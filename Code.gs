// ╔══════════════════════════════════════════════════════════════╗
// ║         GOOGLE APPS SCRIPT - ỨNG DỤNG THI TRỰC TUYẾN       ║
// ║  Copy toàn bộ code này vào Google Apps Script Editor         ║
// ╚══════════════════════════════════════════════════════════════╝

// ============== CẤU HÌNH ==============
const SHEET_ID = 'THAY_BANG_SHEET_ID_CUA_BAN'; // <-- Thay bằng ID Google Sheet của bạn
const ADMIN_PASSWORD = 'admin@2025';
const ROOT_FOLDER_NAME = 'KET_QUA_THI';

// ============== FIREBASE CONFIG ==============
const FIREBASE_URL = 'https://thi-truc-tuyen-967d7-default-rtdb.asia-southeast1.firebasedatabase.app';
const FIREBASE_SECRET = '0qpsUKXIK8egQWxoW5bvwo7z1V1jZTGjycJjOjCD';

// ============== HELPER: LẤY THƯ MỤC GỐC THEO ĐỢT THI ==============
// Cấu trúc: KET_QUA_THI / [Đợt Thi] / [Lớp] / [HS] / files
function getExamRootFolder() {
  var rootFolder = getOrCreateFolder(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
  // Đọc tên đợt thi từ Firebase settings
  try {
    var settings = firebaseGet('settings');
    if (settings && settings.examPeriod && settings.examPeriod.trim()) {
      return getOrCreateFolder(rootFolder, settings.examPeriod.trim());
    }
  } catch (e) {
    Logger.log('getExamRootFolder: cannot read examPeriod, using root. Error: ' + e);
  }
  // Fallback: dùng thư mục gốc nếu chưa cài đợt thi
  return rootFolder;
}

// ============== HELPER: LẤY TÊN ĐỢT THI HIỆN TẠI ==============
function getCurrentExamPeriod() {
  try {
    var settings = firebaseGet('settings');
    if (settings && settings.examPeriod && settings.examPeriod.trim()) {
      return settings.examPeriod.trim();
    }
  } catch (e) {
    Logger.log('getCurrentExamPeriod error: ' + e);
  }
  return '';
}

// ============== FIREBASE REST HELPERS ==============
function firebaseGet(path) {
  var url = FIREBASE_URL + '/' + path + '.json?auth=' + FIREBASE_SECRET;
  var response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  return JSON.parse(response.getContentText());
}

function firebaseSet(path, data) {
  var url = FIREBASE_URL + '/' + path + '.json?auth=' + FIREBASE_SECRET;
  UrlFetchApp.fetch(url, {
    method: 'put',
    contentType: 'application/json',
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  });
}

function firebasePush(path, data) {
  var url = FIREBASE_URL + '/' + path + '.json?auth=' + FIREBASE_SECRET;
  var resp = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  });
  return JSON.parse(resp.getContentText()).name;
}

function firebaseUpdate(path, data) {
  var url = FIREBASE_URL + '/' + path + '.json?auth=' + FIREBASE_SECRET;
  UrlFetchApp.fetch(url, {
    method: 'patch',
    contentType: 'application/json',
    payload: JSON.stringify(data),
    muteHttpExceptions: true
  });
}

function firebaseDelete(path) {
  var url = FIREBASE_URL + '/' + path + '.json?auth=' + FIREBASE_SECRET;
  UrlFetchApp.fetch(url, { method: 'delete', muteHttpExceptions: true });
}

// ============== SESSION SERVER-SIDE (P0-1) ==============
var SESSION_TTL_SECONDS = 2 * 60 * 60; // 2 giờ

function createSession(role, userId, extra) {
  var token = Utilities.getUuid() + '-' + Utilities.getUuid();
  var payload = {
    role: role,
    userId: userId || '',
    createdAt: Date.now(),
    extra: extra || {}
  };
  CacheService.getScriptCache().put('session:' + token, JSON.stringify(payload), SESSION_TTL_SECONDS);
  return token;
}

function getSession(sessionToken) {
  if (!sessionToken) return null;
  try {
    var raw = CacheService.getScriptCache().get('session:' + sessionToken);
    if (!raw) return null;
    return JSON.parse(raw);
  } catch (e) {
    return null;
  }
}

function requireSession(sessionToken, allowedRoles) {
  var session = getSession(sessionToken);
  if (!session || !session.role) return null;
  for (var i = 0; i < allowedRoles.length; i++) {
    if (session.role === allowedRoles[i]) return session;
  }
  return null;
}

function requireAdminSession(sessionToken) {
  return requireSession(sessionToken, ['admin']);
}

function requireTeacherOrAdminSession(sessionToken) {
  return requireSession(sessionToken, ['teacher', 'admin']);
}

// ============== UTILITY HELPERS ==============
function makeSafeFirebaseKey(value) {
  var s = (value || 'default').toString().trim();
  s = removeVietnameseTones(s);
  s = s.replace(/[.#$\[\]\/]/g, '_');
  s = s.replace(/\s+/g, '_');
  return s || 'default';
}

function validateClassName(lop) {
  if (!lop || typeof lop !== 'string') return false;
  var s = lop.trim();
  if (s.length === 0 || s.length > 30) return false;
  return /^[0-9]{1,2}[A-Za-z][0-9A-Za-z_\-]*$/.test(s);
}

function escapeRegex(s) {
  return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function jsonOutput(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============== ACCESS CONTROL HELPERS (P0-4) ==============
function actionRequiresAdmin(action) {
  var adminActions = {
    clearAllData: true, resetForNewPeriod: true, setup: true,
    syncSheetToFirebase: true, syncFirebaseToSheet: true,
    saveTeacher: true, deleteTeacher: true, setSettings: true,
    saveExamConfig: true, migrateQuestionCodes: true
  };
  return !!adminActions[action];
}

function actionRequiresTeacherOrAdmin(action) {
  var actions = {
    updateStudent: true, saveQuestionWithSync: true,
    deleteQuestionWithSync: true, deleteQuestionsByGrade: true,
    bulkSaveQuestions: true, updateQuestion: true,
    uploadStudyFile: true, downloadClassFiles: true,
    downloadStudentFiles: true, cleanupDownloadZip: true,
    addStudentWithSync: true, deleteStudentWithSync: true,
    importStudentsWithSync: true, deleteClassWithSync: true,
    exportResults: true, getExamFolderUrl: true,
    getDashboard: true, getResults: true, getTeachers: true,
    getAllQuestions: true, getExamConfig: true, getPracticeLeaderboard: true,
    sendNotification: true,
    syncStudentStatus: true,
    manualMarkSubmitted: true,
    verifyDriveFiles: true,
    verifyStudentFile: true,
    verifyByFileId: true
  };
  return !!actions[action];
}

// ╔══════════════════════════════════════════════════════════════╗
// ║  HÀM TỰ ĐỘNG TẠO SHEET — Chạy hàm này TRƯỚC TIÊN!         ║
// ║  Vào Apps Script > chọn hàm setupSheets > nhấn ▶ Run        ║
// ╚══════════════════════════════════════════════════════════════╝
function setupSheets() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // ===== Sheet 1: Hoc_Sinh =====
  var hsSheet = ss.getSheetByName('Hoc_Sinh');
  if (!hsSheet) {
    hsSheet = ss.insertSheet('Hoc_Sinh');
  } else {
    hsSheet.clear();
  }
  var hsHeaders = ['STT', 'Ma_HS', 'Ho_Ten', 'Lop', 'Mat_Khau', 'Trang_Thai', 'Thoi_Gian_DN', 'Thoi_Gian_NB'];
  hsSheet.getRange(1, 1, 1, hsHeaders.length).setValues([hsHeaders]);
  hsSheet.getRange(1, 1, 1, hsHeaders.length)
    .setFontWeight('bold')
    .setBackground('#4a86e8')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  hsSheet.setFrozenRows(1);
  hsSheet.setColumnWidth(1, 50);   // STT
  hsSheet.setColumnWidth(2, 100);  // Ma_HS
  hsSheet.setColumnWidth(3, 200);  // Ho_Ten
  hsSheet.setColumnWidth(4, 80);   // Lop
  hsSheet.setColumnWidth(5, 100);  // Mat_Khau
  hsSheet.setColumnWidth(6, 100);  // Trang_Thai
  hsSheet.setColumnWidth(7, 160);  // Thoi_Gian_DN
  hsSheet.setColumnWidth(8, 160);  // Thoi_Gian_NB

  // Thêm 5 dòng dữ liệu mẫu (mật khẩu mặc định: 123456)
  var sampleStudents = [
    [1, 'HS001', 'Nguyễn Văn A', '7A1', '123456', '', '', ''],
    [2, 'HS002', 'Trần Thị B', '7A1', '123456', '', '', ''],
    [3, 'HS003', 'Lê Văn C', '7A2', '123456', '', '', ''],
    [4, 'HS004', 'Phạm Thị D', '8A1', '123456', '', '', ''],
    [5, 'HS005', 'Hoàng Văn E', '6A1', '123456', '', '', '']
  ];
  hsSheet.getRange(2, 1, sampleStudents.length, 8).setValues(sampleStudents);

  // ===== Sheet 2: Cau_Hoi =====
  var chSheet = ss.getSheetByName('Cau_Hoi');
  if (!chSheet) {
    chSheet = ss.insertSheet('Cau_Hoi');
  } else {
    chSheet.clear();
  }
  var chHeaders = ['STT', 'Noi_Dung', 'Dap_An_1', 'Dap_An_2', 'Dap_An_3', 'Dap_An_4', 'Dap_An_Dung', 'Lop'];
  chSheet.getRange(1, 1, 1, chHeaders.length).setValues([chHeaders]);
  chSheet.getRange(1, 1, 1, chHeaders.length)
    .setFontWeight('bold')
    .setBackground('#6aa84f')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  chSheet.setFrozenRows(1);
  chSheet.setColumnWidth(1, 50);
  chSheet.setColumnWidth(2, 400);
  chSheet.setColumnWidth(3, 200);
  chSheet.setColumnWidth(4, 200);
  chSheet.setColumnWidth(5, 200);
  chSheet.setColumnWidth(6, 200);
  chSheet.setColumnWidth(7, 100);
  chSheet.setColumnWidth(8, 60);

  // Thêm 16 câu hỏi mẫu (Tin học) — cột Lop = khối lớp (6, 7, 8, 9 hoặc để trống = chung)
  var sampleQuestions = [
    [1, 'Đơn vị nhỏ nhất của thông tin trong máy tính là gì?', 'Byte', 'Bit', 'KB', 'MB', 'b', '6'],
    [2, '1 Byte bằng bao nhiêu Bit?', '4 Bit', '8 Bit', '16 Bit', '32 Bit', 'b', '6'],
    [3, 'CPU là viết tắt của từ gì?', 'Central Processing Unit', 'Computer Personal Unit', 'Central Program Utility', 'Computer Processing Unit', 'a', '7'],
    [4, 'Phần mềm nào dùng để soạn thảo văn bản?', 'Excel', 'PowerPoint', 'Word', 'Access', 'c', '7'],
    [5, 'RAM là bộ nhớ gì?', 'Bộ nhớ ngoài', 'Bộ nhớ trong (tạm thời)', 'Bộ nhớ chỉ đọc', 'Bộ nhớ vĩnh viễn', 'b', '7'],
    [6, 'Phím tắt Ctrl+C dùng để làm gì?', 'Cắt', 'Sao chép', 'Dán', 'In', 'b', '8'],
    [7, 'Phím tắt Ctrl+V dùng để làm gì?', 'Sao chép', 'Cắt', 'Dán', 'Lưu', 'c', '8'],
    [8, 'Hệ điều hành nào phổ biến nhất trên máy tính cá nhân?', 'Linux', 'MacOS', 'Windows', 'Android', 'c', '8'],
    [9, 'Thiết bị nào là thiết bị đầu vào?', 'Máy in', 'Loa', 'Bàn phím', 'Màn hình', 'c', '9'],
    [10, 'Thiết bị nào là thiết bị đầu ra?', 'Chuột', 'Bàn phím', 'Máy quét', 'Máy in', 'd', '9'],
    [11, 'Đuôi file .docx thuộc phần mềm nào?', 'Excel', 'Word', 'PowerPoint', 'Access', 'b', ''],
    [12, 'Đuôi file .xlsx thuộc phần mềm nào?', 'Word', 'Excel', 'PowerPoint', 'Notepad', 'b', ''],
    [13, 'Phím tắt Ctrl+Z dùng để làm gì?', 'Lưu file', 'Hoàn tác', 'Làm lại', 'Đóng file', 'b', ''],
    [14, 'Internet là gì?', 'Phần mềm máy tính', 'Mạng máy tính toàn cầu', 'Hệ điều hành', 'Thiết bị phần cứng', 'b', ''],
    [15, 'Virus máy tính là gì?', 'Phần cứng hỏng', 'Chương trình gây hại', 'Lỗi hệ điều hành', 'File bị xóa', 'b', '6'],
    [16, 'Phím tắt Ctrl+S dùng để làm gì?', 'Sao chép', 'Lưu file', 'Tìm kiếm', 'Chọn tất cả', 'b', '9']
  ];
  chSheet.getRange(2, 1, sampleQuestions.length, 8).setValues(sampleQuestions);

  // ===== Sheet 3: Ket_Qua =====
  var kqSheet = ss.getSheetByName('Ket_Qua');
  if (!kqSheet) {
    kqSheet = ss.insertSheet('Ket_Qua');
  } else {
    kqSheet.clear();
  }
  var kqHeaders = ['STT', 'Ma_HS', 'Ho_Ten', 'Lop'];
  for (var _h = 1; _h <= 30; _h++) kqHeaders.push('Cau_' + _h);
  kqHeaders.push('Diem', 'Thoi_Gian');
  kqSheet.getRange(1, 1, 1, kqHeaders.length).setValues([kqHeaders]);
  kqSheet.getRange(1, 1, 1, kqHeaders.length)
    .setFontWeight('bold')
    .setBackground('#e69138')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  kqSheet.setFrozenRows(1);
  kqSheet.setColumnWidth(1, 50);
  kqSheet.setColumnWidth(2, 100);
  kqSheet.setColumnWidth(3, 200);
  kqSheet.setColumnWidth(4, 80);
  kqSheet.setColumnWidth(13, 60);
  kqSheet.setColumnWidth(14, 160);

  // ===== Sheet 4: Giao_Vien (MỚI) =====
  var gvSheet = ss.getSheetByName('Giao_Vien');
  if (!gvSheet) {
    gvSheet = ss.insertSheet('Giao_Vien');
  } else {
    gvSheet.clear();
  }
  var gvHeaders = ['MaGV', 'Ho_Ten', 'Mat_Khau', 'Lop_Phu_Trach'];
  gvSheet.getRange(1, 1, 1, gvHeaders.length).setValues([gvHeaders]);
  gvSheet.getRange(1, 1, 1, gvHeaders.length)
    .setFontWeight('bold')
    .setBackground('#8e24aa')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  gvSheet.setFrozenRows(1);
  gvSheet.setColumnWidth(1, 100);
  gvSheet.setColumnWidth(2, 200);
  gvSheet.setColumnWidth(3, 120);
  gvSheet.setColumnWidth(4, 350);

  // Dữ liệu mẫu GV
  var sampleTeachers = [
    ['GV001', 'Nguyễn Văn A', 'abc123', '7A1,7A2,7A3,7A4,6A1,6A2,6A3'],
    ['GV002', 'Trần Thị B', 'xyz789', '8A1,8A2,8A3,8A4,6A4,6A5,6A6'],
    ['GV003', 'Lê Văn C', 'def456', '9A1,9A2,9A3,9A4,6A7,6A8,6A9']
  ];
  gvSheet.getRange(2, 1, sampleTeachers.length, 4).setValues(sampleTeachers);

  // ===== Sheet 5: De_Thi (MỚI) =====
  var dtSheet = ss.getSheetByName('De_Thi');
  if (!dtSheet) {
    dtSheet = ss.insertSheet('De_Thi');
  } else {
    dtSheet.clear();
  }
  var dtHeaders = ['Exam_ID', 'Ten_De', 'Ma_GV', 'Lop_Ap_Dung', 'So_Cau_Thi', 'So_Cau_Thu', 'Trang_Thai', 'Thi_Thu', 'Max_Attempts'];
  dtSheet.getRange(1, 1, 1, dtHeaders.length).setValues([dtHeaders]);
  dtSheet.getRange(1, 1, 1, dtHeaders.length)
    .setFontWeight('bold')
    .setBackground('#c62828')
    .setFontColor('#ffffff')
    .setHorizontalAlignment('center');
  dtSheet.setFrozenRows(1);
  dtSheet.setColumnWidth(1, 100);
  dtSheet.setColumnWidth(2, 200);
  dtSheet.setColumnWidth(3, 80);
  dtSheet.setColumnWidth(4, 250);
  dtSheet.setColumnWidth(5, 90);
  dtSheet.setColumnWidth(6, 90);
  dtSheet.setColumnWidth(7, 90);
  dtSheet.setColumnWidth(8, 80);
  dtSheet.setColumnWidth(9, 100);

  // Xóa sheet mặc định "Sheet1" nếu có
  try {
    var defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Trang tính1');
    if (defaultSheet) ss.deleteSheet(defaultSheet);
  } catch(e) {}

  Logger.log('✅ Đã tạo thành công 5 sheet: Hoc_Sinh, Cau_Hoi, Ket_Qua, Giao_Vien, De_Thi');
  Logger.log('📝 Đã thêm 5 học sinh mẫu, 16 câu hỏi mẫu, 3 giáo viên mẫu');
  Logger.log('👉 Bạn có thể thay đổi dữ liệu mẫu theo ý muốn.');

  return '✅ Setup thành công! Đã tạo 5 sheet với dữ liệu mẫu.';
}

// ============== XỬ LÝ REQUEST ==============
function doGet(e) {
  return handleRequest(e);
}

function doPost(e) {
  return handleRequest(e);
}

function handleRequest(e) {
  // GAS REDIRECT FIX: Đọc action từ URL params TRƯỚC (luôn tồn tại sau redirect)
  // Sau đó mới parse body để lấy thêm data nhạy cảm (password, fileData, etc.)
  var urlAction = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  var params = {};

  if (e && e.postData && e.postData.contents) {
    try {
      params = JSON.parse(e.postData.contents);
    } catch (jsonErr) {
      params = e.parameter || {};
    }
  } else {
    params = e.parameter || {};
  }

  // URL action có độ ưu tiên cao nhất — không bị mất khi redirect
  var action = urlAction || params.action || '';
  var result = {};

  try {
    // P0-4: Access control gate
    if (actionRequiresAdmin(action)) {
      if (!requireAdminSession(params.sessionToken)) {
        return jsonOutput({ success: false, message: '⛔ Không có quyền Admin! Vui lòng đăng nhập lại.' });
      }
    }
    if (actionRequiresTeacherOrAdmin(action)) {
      if (!requireTeacherOrAdminSession(params.sessionToken)) {
        return jsonOutput({ success: false, message: '⛔ Không có quyền Giáo viên/Admin! Vui lòng đăng nhập lại.' });
      }
    }

    switch (action) {
      case 'login':
        result = studentLogin(params.maHS, params.password);
        break;
      case 'loginTeacher':
        result = teacherLogin(params.maGV, params.password);
        break;
      case 'loginAdmin':
        result = adminLogin(params.password);
        break;
      case 'getQuestions':
        result = getRandomQuestions(params.mode || '', params.lop || '', params.sessionToken);
        break;
      case 'getClasses':
        result = getClasses();
        break;
      case 'getDashboard':
        result = getDashboard(params.lop || '');
        break;
      case 'getResults':
        result = getResults(params.lop || '');
        break;
      case 'submitExam':
        result = submitExam(params);
        break;
      case 'saveFile':
        result = saveFileOnly(params);
        break;
      case 'verifyFile':
        result = verifyFileExists(params.maHS);
        break;
      case 'updateStudent':
        var val = params.value === '__EMPTY__' ? '' : (params.value || '');
        result = updateStudent(params.maHS, params.field, val);
        break;
      case 'getSettings':
        result = getSettings();
        break;
      case 'setSettings':
        result = setSettings(params.key, params.value);
        break;
      case 'getStudentResults':
        result = getStudentResults(params.maHS);
        break;
      case 'getPracticeLeaderboard':
        result = getPracticeLeaderboard();
        break;
      case 'exportResults':
        result = exportResultsCSV(params.lop || '');
        break;
      case 'clearAllData':
        result = clearAllData(params.sessionToken, params.confirmPeriod);
        break;
      case 'resetForNewPeriod':
        result = resetForNewPeriod(params.sessionToken);
        break;
      case 'getExamFolderUrl':
        result = getExamFolderUrl();
        break;
      case 'setup':
        result = { success: true, message: setupSheets() };
        break;
      case 'getAllQuestions':
        result = getAllQuestions();
        break;
      case 'saveQuestionWithSync':
        result = saveQuestionWithSync(params);
        break;
      case 'deleteQuestionWithSync':
        result = deleteQuestionWithSync(params.maCH);
        break;
      case 'deleteQuestionsByGrade':
        result = deleteQuestionsByGrade(params.khoi);
        break;
      case 'bulkSaveQuestions':
        result = bulkSaveQuestions(params);
        break;
      case 'getExamConfig':
        result = getExamConfig(params.khoi || '');
        break;
      case 'saveExamConfig':
        result = saveExamConfig(params);
        break;
      case 'migrateQuestionCodes':
        result = migrateQuestionCodes();
        break;
      case 'updateQuestion':
        result = updateQuestion(params);
        break;
      case 'syncSheetToFirebase':
        result = syncSheetToFirebase();
        break;
      case 'syncFirebaseToSheet':
        result = syncFirebaseToSheet();
        break;
      case 'getTeachers':
        result = getTeachers();
        break;
      case 'saveTeacher':
        result = saveTeacher(params);
        break;
      case 'deleteTeacher':
        result = deleteTeacher(params.maGV);
        break;
      case 'uploadStudyFile':
        result = uploadStudyFile(params);
        break;
      case 'downloadClassFiles':
        result = downloadClassFiles(params.lop, params.sessionToken);
        break;
      case 'downloadStudentFiles':
        result = downloadStudentFiles(params.lop, params.maHS, params.sessionToken);
        break;
      case 'cleanupDownloadZip':
        result = cleanupDownloadZip(params.fileId, params.sessionToken);
        break;
      case 'addStudentWithSync':
        result = addStudentWithSync(params);
        break;
      case 'deleteStudentWithSync':
        result = deleteStudentWithSync(params.maHS, params.sessionToken);
        break;
      case 'importStudentsWithSync':
        result = importStudentsWithSync(params);
        break;
      case 'deleteClassWithSync':
        result = deleteClassWithSync(params.lop, params.sessionToken);
        break;
      // P0-5 addition: save student answers server-side
      case 'saveStudentAnswers':
        result = saveStudentAnswers(params);
        break;
      // D1b: release student session khi logout
      case 'releaseStudentSession':
        result = releaseStudentSession(params);
        break;
      case 'sendNotification':
        result = sendNotificationAction(params);
        break;
      // HS nộp lại file khi Drive save thất bại
      case 'resubmitFile':
        result = resubmitFile(params);
        break;
      // Kiểm tra trạng thái file của HS (cho retry UI)
      case 'getFileStatus':
        result = getStudentFileStatus(params.maHS, params.sessionToken);
        break;
      // Fix HS bị kẹt đang thi
      case 'syncStudentStatus':
        result = syncStudentStatus();
        break;
      case 'manualMarkSubmitted':
        result = manualMarkSubmitted(params);
        break;
      // Ki?m tra file th?c t? trên Drive — fix hasFile sai
      case 'verifyDriveFiles':
        result = verifyDriveFiles();
        break;
      case 'verifyStudentFile':
        result = verifyStudentFile(params);
        break;
      case 'verifyByFileId':
        result = verifyByFileId(params);
        break;
      default:
        result = { success: false, message: 'Action không hợp lệ' };
    }
  } catch (err) {
    result = { success: false, message: 'Lỗi hệ thống: ' + err.toString() };
  }

  return jsonOutput(result);
}

// ============== ĐĂNG NHẬP HỌC SINH ==============
function studentLogin(maHS, password) {
  if (!maHS || !password) {
    return { success: false, message: 'Vui lòng nhập đầy đủ Mã HS và Mật khẩu!' };
  }

  // Validate mã HS — chống Firebase path injection
  if (!validateUserCode(maHS)) {
    return { success: false, message: '❌ Mã học sinh không hợp lệ! Chỉ chấp nhận chữ, số, gạch dưới.' };
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = getSheetSafe(ss, 'Hoc_Sinh');
  if (!sheet) return { success: false, message: '❌ Lỗi cấu hình: Không tìm thấy sheet Hoc_Sinh!' };
  var data = sheet.getDataRange().getValues();

  // Kiểm tra chế độ thi thử (null-safe)
  var settingsResult = getSettings();
  var settings = settingsResult && settingsResult.settings ? settingsResult.settings : {};
  var isPractice = settings.practiceMode === true || settings.practiceMode === 'true';

  for (var i = 1; i < data.length; i++) {
    var sheetMaHS = data[i][1].toString().trim();

    if (sheetMaHS === maHS.trim()) {
      // Kiểm tra mật khẩu (cột E = index 4, mặc định 123456 nếu để trống)
      var sheetPass = data[i][4].toString().trim();
      if (!sheetPass) sheetPass = '123456';
      if (sheetPass !== password.trim()) {
        return { success: false, message: '❌ Mật khẩu không đúng!' };
      }

      // Kiểm tra trạng thái thi (cột F = index 5)
      var trangThai = data[i][5].toString().trim().toUpperCase();

      // Kiểm tra giáo viên đã mở kỳ thi chưa (hỗ trợ moThiThat/moThiThu mới + moThi cũ)
      var studentFb = firebaseGet('students/' + maHS.trim());
      if (!studentFb) {
        return { success: false, message: '⏳ Giáo viên chưa mở kỳ thi cho bạn. Vui lòng chờ hướng dẫn từ giáo viên!' };
      }
      var canReal = studentFb.moThiThat !== undefined ? !!studentFb.moThiThat : !!studentFb.moThi;
      var canPractice = studentFb.moThiThu !== undefined ? !!studentFb.moThiThu : !!studentFb.moThi;
      // Auto-switch: nếu chế độ chính bị khóa nhưng chế độ còn lại đang mở → tự chuyển
      if (!isPractice && !canReal && canPractice) {
        isPractice = true; // Thi thật bị khóa, thi thử đang mở → chuyển sang thi thử
      }
      if (isPractice && !canPractice && canReal) {
        isPractice = false; // Thi thử bị khóa, thi thật đang mở → chuyển sang thi thật
      }
      if (isPractice && !canPractice) {
        return { success: false, message: '⏳ Giáo viên chưa mở THI THỬ cho bạn. Vui lòng chờ hướng dẫn từ giáo viên!' };
      }
      if (!isPractice && !canReal) {
        return { success: false, message: '⏳ Giáo viên chưa mở kỳ thi cho bạn. Vui lòng chờ hướng dẫn từ giáo viên!' };
      }

      if (isPractice) {
        // Chế độ thi thử: đếm số lần đã thi
        var attempts = countPracticeAttempts(ss, maHS.trim());
        if (attempts >= 5) {
          return { success: false, message: '⚠️ Bạn đã hết 5 lượt thi thử! Liên hệ giáo viên.', alreadyTested: true };
        }
        // Ghi thời gian đăng nhập
        sheet.getRange(i + 1, 7).setValue(new Date());
        return {
          success: true,
          practiceMode: true,
          practiceAttempts: attempts,
          practiceMaxAttempts: 5,
          sessionToken: createSession('student', sheetMaHS, { lop: data[i][3].toString().trim() }),
          student: {
            stt: data[i][0],
            maHS: sheetMaHS,
            hoTen: data[i][2].toString().trim(),
            lop: data[i][3].toString().trim()
          }
        };
      } else {
        // Chế độ thi thật
        if (trangThai === 'X') {
          return {
            success: false,
            message: '⚠️ Bạn đã hoàn thành bài thi rồi! Không thể thi lại.',
            alreadyTested: true
          };
        }
        // Ghi thời gian đăng nhập (cột G = index 7, row i+1)
        sheet.getRange(i + 1, 7).setValue(new Date());
        return {
          success: true,
          practiceMode: false,
          sessionToken: createSession('student', sheetMaHS, { lop: data[i][3].toString().trim() }),
          student: {
            stt: data[i][0],
            maHS: sheetMaHS,
            hoTen: data[i][2].toString().trim(),
            lop: data[i][3].toString().trim()
          }
        };
      }
    }
  }

  return { success: false, message: '❌ Mã học sinh không tồn tại! Vui lòng kiểm tra lại.' };
}

// Đếm số lần thi thử của HS (đọc từ Firebase — đồng bộ với reset)
function countPracticeAttempts(ss, maHS) {
  // Đếm từ Firebase (nguồn chính) — CHỈ đếm lượt của đợt thi hiện tại
  try {
    var currentPeriod = getCurrentExamPeriod();
    var practiceData = firebaseGet('practiceResults') || {};
    var count = 0;
    for (var key in practiceData) {
      if (practiceData[key] && practiceData[key].maHS === maHS) {
        var rDotThi = (practiceData[key].dotThi || '').toString().trim();
        // Chỉ đếm nếu cùng đợt thi hiện tại
        if (currentPeriod === '' && rDotThi === '') {
          count++;
        } else if (currentPeriod !== '' && rDotThi === currentPeriod) {
          count++;
        }
      }
    }
    return count;
  } catch (e) {
    // Fallback: đếm từ GGSheet nếu Firebase lỗi
    Logger.log('Firebase read error, fallback to GGSheet: ' + e);
    var sheet = getOrCreatePracticeSheet(ss);
    var data = sheet.getDataRange().getValues();
    var currentPeriod2 = getCurrentExamPeriod();
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === maHS) {
        var rowCols = data[i].length;
        var rowDotThi = (data[i][rowCols - 1] || '').toString().trim();
        if (currentPeriod2 === '' && rowDotThi === '') count++;
        else if (currentPeriod2 !== '' && rowDotThi === currentPeriod2) count++;
      }
    }
    return count;
  }
}

// Tạo sheet Ket_Qua_Thu nếu chưa có
function getOrCreatePracticeSheet(ss) {
  var sheet = ss.getSheetByName('Ket_Qua_Thu');
  if (!sheet) {
    sheet = ss.insertSheet('Ket_Qua_Thu');
    var header = ['STT', 'Ma_HS', 'Ho_Ten', 'Lop', 'Lan_Thu'];
    for (var i = 1; i <= 16; i++) header.push('C' + i);
    header.push('Diem', 'Thoi_Gian');
    sheet.appendRow(header);
    sheet.getRange(1, 1, 1, header.length)
      .setFontWeight('bold')
      .setBackground('#9900ff')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============== ĐĂNG NHẬP GIÁO VIÊN ==============
// C2 FIX: Hỗ trợ tìm GV theo password-only khi maGV rỗng (client không cần gửi maGV)
function teacherLogin(maGV, password) {
  if (!password) return { success: false, message: '❌ Vui lòng nhập mật khẩu!' };

  // Nếu client không truyền maGV → tìm GV theo mật khẩu trong toàn bộ teachers
  if (!maGV || maGV.toString().trim() === '') {
    var allTeachers = firebaseGet('teachers') || {};
    var foundTeacher = null;
    var foundMaGV = '';
    for (var gvKey in allTeachers) {
      if (allTeachers[gvKey].matKhau && allTeachers[gvKey].matKhau === password.trim()) {
        foundTeacher = allTeachers[gvKey];
        foundMaGV = gvKey;
        break;
      }
    }
    if (!foundTeacher) return { success: false, message: '❌ Mật khẩu không đúng!' };
    return {
      success: true,
      message: 'Đăng nhập thành công!',
      sessionToken: createSession('teacher', foundMaGV, { lopPhuTrach: foundTeacher.lopPhuTrach || '' }),
      teacher: {
        maGV: foundMaGV,
        hoTen: foundTeacher.hoTen || '',
        lopPhuTrach: foundTeacher.lopPhuTrach || ''
      }
    };
  }

  // Nếu client truyền maGV → validate và tìm trực tiếp
  if (!validateUserCode(maGV)) return { success: false, message: '❌ Mã giáo viên không hợp lệ!' };
  var teacher = firebaseGet('teachers/' + maGV.trim());
  if (!teacher) return { success: false, message: '❌ Mã giáo viên không tồn tại!' };
  if (teacher.matKhau !== password.trim()) return { success: false, message: '❌ Mật khẩu không đúng!' };
  return {
    success: true,
    message: 'Đăng nhập thành công!',
    sessionToken: createSession('teacher', maGV.trim(), { lopPhuTrach: teacher.lopPhuTrach || '' }),
    teacher: {
      maGV: maGV.trim(),
      hoTen: teacher.hoTen || '',
      lopPhuTrach: teacher.lopPhuTrach || ''
    }
  };
}

// ============== ĐĂNG NHẬP ADMIN ==============
function adminLogin(password) {
  if (!password) return { success: false, message: '❌ Vui lòng nhập mật khẩu!' };
  // Check Firebase first, fallback to hardcoded
  var adminData = firebaseGet('admin');
  var adminPw = (adminData && adminData.password) ? adminData.password : ADMIN_PASSWORD;
  if (password === adminPw) {
    return {
      success: true,
      message: 'Đăng nhập Admin thành công!',
      isAdmin: true,
      sessionToken: createSession('admin', 'admin', {})
    };
  }
  return { success: false, message: '❌ Mật khẩu Admin không đúng!' };
}

// ============== QUẢN LÝ GIÁO VIÊN ==============
function getTeachers() {
  var teachers = firebaseGet('teachers') || {};
  var list = [];
  for (var maGV in teachers) {
    list.push({
      maGV: maGV,
      hoTen: teachers[maGV].hoTen || '',
      matKhau: teachers[maGV].matKhau || '',
      lopPhuTrach: teachers[maGV].lopPhuTrach || ''
    });
  }
  return { success: true, teachers: list };
}

function saveTeacher(data) {
  if (!data.maGV) return { success: false, message: 'Thiếu mã giáo viên!' };
  var maGV = data.maGV.toString().trim();
  firebaseUpdate('teachers/' + maGV, {
    hoTen: data.hoTen || '',
    matKhau: data.matKhau || '',
    lopPhuTrach: data.lopPhuTrach || ''
  });
  // Sync to GGSheet
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var gvSheet = ss.getSheetByName('Giao_Vien');
    if (gvSheet) {
      var gvData = gvSheet.getDataRange().getValues();
      var found = false;
      for (var i = 1; i < gvData.length; i++) {
        if (gvData[i][0].toString().trim() === maGV) {
          gvSheet.getRange(i + 1, 2).setValue(data.hoTen || '');
          gvSheet.getRange(i + 1, 3).setValue(data.matKhau || '');
          gvSheet.getRange(i + 1, 4).setValue(data.lopPhuTrach || '');
          found = true;
          break;
        }
      }
      if (!found) {
        gvSheet.appendRow([maGV, data.hoTen || '', data.matKhau || '', data.lopPhuTrach || '']);
      }
    }
  } catch(se) { Logger.log('Sync GV to sheet error: ' + se); }
  return { success: true, message: 'Đã lưu giáo viên ' + maGV + '!' };
}

function deleteTeacher(maGV) {
  if (!maGV) return { success: false, message: 'Thiếu mã giáo viên!' };
  maGV = maGV.toString().trim();
  firebaseDelete('teachers/' + maGV);
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var gvSheet = ss.getSheetByName('Giao_Vien');
    if (gvSheet) {
      var gvData = gvSheet.getDataRange().getValues();
      for (var i = 1; i < gvData.length; i++) {
        if (gvData[i][0].toString().trim() === maGV) {
          gvSheet.deleteRow(i + 1);
          break;
        }
      }
    }
  } catch(se) { Logger.log('Delete GV error: ' + se); }
  return { success: true, message: 'Đã xóa giáo viên ' + maGV + '!' };
}

// ============== LẤY CÂU HỎI ==============
function getRandomQuestions(mode, lop) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Cau_Hoi');
  var data = sheet.getDataRange().getValues();

  // Extract khối from lớp (VD: '8A1' → '8')
  var studentGrade = '';
  if (lop) {
    var gradeMatch = lop.toString().match(/^(\d+)/);
    if (gradeMatch) studentGrade = gradeMatch[1];
  }

  var allQuestions = [];
  for (var i = 1; i < data.length; i++) {
    var maCH = data[i][0].toString().trim();
    var qKhoi = data[i][7] ? data[i][7].toString().trim() : '';
    
    // Lọc theo khối lớp HS (nếu có) — câu chung (qKhoi rỗng) luôn được lấy
    if (studentGrade && qKhoi && qKhoi !== studentGrade) continue;
    
    allQuestions.push({
      stt: maCH,
      maCH: maCH,
      noiDung: data[i][1].toString(),
      dapAn: [
        data[i][2].toString(),
        data[i][3].toString(),
        data[i][4].toString(),
        data[i][5].toString()
      ],
      khoi: qKhoi
      // BẢO MẬT: KHÔNG gửi dapAnDung cho client
    });
  }

  if (mode === 'practice') {
    // Thi thử: đọc examConfig để biết số câu
    var soCauThiThu = allQuestions.length; // mặc định: tất cả
    if (studentGrade) {
      try {
        var examConfig = firebaseGet('examConfig/' + studentGrade);
        if (examConfig && examConfig.soCauThiThu && parseInt(examConfig.soCauThiThu) > 0) {
          soCauThiThu = parseInt(examConfig.soCauThiThu);
        }
      } catch(e) { /* dùng tất cả */ }
    }
    shuffleArray(allQuestions);
    return { success: true, questions: allQuestions.slice(0, soCauThiThu) };
  } else {
    // Thi thật: đọc examConfig/selectedQuestions
    if (studentGrade) {
      try {
        var examConfig = firebaseGet('examConfig/' + studentGrade);
        if (examConfig && examConfig.selectedQuestions && examConfig.selectedQuestions.length > 0) {
          var selectedSet = {};
          for (var s = 0; s < examConfig.selectedQuestions.length; s++) {
            selectedSet[examConfig.selectedQuestions[s]] = true;
          }
          var selected = allQuestions.filter(function(q) { return selectedSet[q.maCH]; });
          if (selected.length > 0) {
            shuffleArray(selected);
            return { success: true, questions: selected };
          }
        }
      } catch(e) { /* fallback random */ }
    }
    // Fallback: random 8 câu nếu chưa cấu hình
    shuffleArray(allQuestions);
    return { success: true, questions: allQuestions.slice(0, 8) };
  }
}

// ============== LẤY DANH SÁCH LỚP ==============
function getClasses() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Hoc_Sinh');
  var data = sheet.getDataRange().getValues();

  var classSet = {};
  for (var i = 1; i < data.length; i++) {
    var lop = data[i][3].toString().trim();
    if (lop) classSet[lop] = true;
  }

  var classes = Object.keys(classSet).sort();
  return { success: true, classes: classes };
}

// ============== NỘP BÀI THI ==============
function submitExam(data) {
  // ===== VALIDATE DỮ LIỆU ĐẦU VÀO (trước khi lấy lock) =====
  if (!data || typeof data !== 'object') {
    return { success: false, message: '❌ Dữ liệu nộp bài không hợp lệ!' };
  }
  if (!data.maHS) {
    return { success: false, message: '❌ Thiếu mã học sinh!' };
  }
  if (!validateUserCode(data.maHS)) {
    return { success: false, message: '❌ Mã học sinh không hợp lệ!' };
  }
  if (!data.answers || !Array.isArray(data.answers) || data.answers.length === 0) {
    return { success: false, message: '❌ Danh sách câu trả lời không hợp lệ hoặc trống!' };
  }

  // Chuẩn hóa maHS 1 lần duy nhất — dùng xuyên suốt hàm
  var maHS = data.maHS.toString().trim();

  // ===== B1: KIỂM TRA SESSION HỌC SINH =====
  var session = requireSession(data.sessionToken, ['student']);
  if (!session || session.userId !== maHS) {
    return { success: false, message: '⛔ Phiên đăng nhập không hợp lệ. Vui lòng đăng nhập lại!' };
  }

  // Chỉ khai báo tạm — sẽ được ghi đè bằng giá trị từ Sheet
  var hoTen = '';
  var lop = '';

  // Race condition protection — chỉ lock phần ghi GGSheet
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // tăng lên 30 giây để hỗ trợ nhiều HS nộp cùng lúc
  } catch (e) {
    // GAS bờờ lock âm thanh: nếu HS retry → server check alreadySubmitted trước lock
    // BUG FIX #7: thêm retryable:true để client không phải parse text message để quyết định retry
    return { success: false, retryable: true, message: '⏳ Hệ thống đang bận (nhiều HS nộp cùng lúc), vui lòng thử lại sau vài giây!' };
  }

  var result; // kết quả trả về
  var needDriveSave = false; // flag để lưu Drive sau khi release lock
  var driveData = null;
  var isAutoSubmitted = false;

  try {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var isPractice = data.isPractice === true || data.isPractice === 'true';
  isAutoSubmitted = data.autoSubmitted === true || data.autoSubmitted === 'true';

  // ===== VALIDATE: Thi thật PHẢI có file thực hành =====
  if (!isPractice) {
    // Kiểm tra file bắt buộc (trừ auto-submit hết giờ)
    if (!data.fileData || !data.fileName) {
      if (!isAutoSubmitted) {
        // KHÔNG gọi lock.releaseLock() thủ công — finally sẽ xử lý
        result = { success: false, message: '⚠️ Thiếu file thực hành! Bạn phải tải file bài thực hành lên trước khi nộp bài.' };
        return result;
      }
      // Auto-submit cho phép thiếu file nhưng ghi cờ cảnh báo
    }
    // Validate file extension (server-side) nếu có file
    if (data.fileName) {
      var allowedExts = ['.doc', '.docx', '.xls', '.xlsx'];
      var fileName = data.fileName.toString().toLowerCase();
      var extValid = false;
      for (var ae = 0; ae < allowedExts.length; ae++) {
        if (fileName.indexOf(allowedExts[ae], fileName.length - allowedExts[ae].length) !== -1) {
          extValid = true;
          break;
        }
      }
      if (!extValid) {
        result = { success: false, message: '⚠️ Sai định dạng file! Chỉ chấp nhận file Word (.doc, .docx) hoặc Excel (.xls, .xlsx).' };
        return result;
      }
    }
  }

  // ===== B2/B3: ĐỌC hoTen VÀ lop TỪ SHEET Hoc_Sinh (không tin client) =====
  var hSheetEarly = getSheetSafe(SpreadsheetApp.openById(SHEET_ID), 'Hoc_Sinh');
  if (!hSheetEarly) {
    result = { success: false, message: '❌ Lỗi cấu hình: Không tìm thấy sheet Hoc_Sinh!' };
    return result;
  }
  var hDataEarly = hSheetEarly.getDataRange().getValues();
  for (var ei = 1; ei < hDataEarly.length; ei++) {
    var rowMaHSEarly = hDataEarly[ei][1] ? hDataEarly[ei][1].toString().trim() : '';
    if (rowMaHSEarly === maHS) {
      hoTen = hDataEarly[ei][2] ? hDataEarly[ei][2].toString().trim() : '';
      lop   = hDataEarly[ei][3] ? hDataEarly[ei][3].toString().trim() : '';
      break;
    }
  }
  if (!hoTen && !lop) {
    // Không tìm thấy maHS trong danh sách
    result = { success: false, message: '❌ Không tìm thấy học sinh trong danh sách!' };
    return result;
  }

  // 1. Lấy đáp án đúng từ sheet Cau_Hoi — LỌC THEO KHỐI LỚP HS
  var qSheet = getSheetSafe(SpreadsheetApp.openById(SHEET_ID), 'Cau_Hoi');
  if (!qSheet) {
    result = { success: false, message: '❌ Lỗi cấu hình: Không tìm thấy sheet Cau_Hoi!' };
    return result;
  }
  var qData = qSheet.getDataRange().getValues();
  var studentGrade = '';
  if (lop) {
    var gradeMatch = lop.match(/^(\d+)/);
    if (gradeMatch) studentGrade = gradeMatch[1];
  }
  var correctAnswers = {};
  var questionContents = {};
  var questionOptions = {};
  for (var i = 1; i < qData.length; i++) {
    var maCH = qData[i][0].toString().trim();
    var qKhoi = qData[i][7] ? qData[i][7].toString().trim() : '';
    if (studentGrade && qKhoi && qKhoi !== studentGrade) continue;
    correctAnswers[maCH] = qData[i][6].toString().trim().toLowerCase();
    questionContents[maCH] = qData[i][1].toString();
    questionOptions[maCH] = {
      a: qData[i][2].toString(),
      b: qData[i][3].toString(),
      c: qData[i][4].toString(),
      d: qData[i][5].toString()
    };
  }

  // ===== B4/B5: officialQuestionList — không fallback sang toàn bộ correctAnswers =====
  var officialQuestionList = []; // danh sách STT câu chính thức theo thứ tự
  var assignedQuestionIds = {};  // map để check nhanh

  try {
    var studentFbData = firebaseGet('students/' + maHS);
    if (studentFbData && studentFbData.assignedQuestions && studentFbData.assignedQuestions.length > 0) {
      for (var aq = 0; aq < studentFbData.assignedQuestions.length; aq++) {
        var aqId = studentFbData.assignedQuestions[aq].toString().trim();
        if (aqId) { officialQuestionList.push(aqId); assignedQuestionIds[aqId] = true; }
      }
    }
  } catch (aqe) {
    Logger.log('Warning: Could not read assignedQuestions: ' + aqe);
  }

  // Nếu chưa có assignedQuestions → đọc từ examConfig (thi thật)
  if (officialQuestionList.length === 0 && !isPractice && studentGrade) {
    try {
      var examCfg = firebaseGet('examConfig/' + studentGrade);
      if (examCfg && examCfg.selectedQuestions && examCfg.selectedQuestions.length > 0) {
        for (var sq = 0; sq < examCfg.selectedQuestions.length; sq++) {
          var sqId = examCfg.selectedQuestions[sq].toString().trim();
          if (sqId) { officialQuestionList.push(sqId); assignedQuestionIds[sqId] = true; }
        }
      }
    } catch (cfgE) { Logger.log('examConfig read error: ' + cfgE); }
  }

  // Nếu thi thử và chưa có danh sách → dùng tất cả câu đúng khối
  if (officialQuestionList.length === 0 && isPractice) {
    var soCauThiThu = 16;
    try {
      var examCfgP = firebaseGet('examConfig/' + studentGrade);
      if (examCfgP && examCfgP.soCauThiThu && parseInt(examCfgP.soCauThiThu) > 0) {
        soCauThiThu = parseInt(examCfgP.soCauThiThu);
      }
    } catch(e) {}
    var allQKeys = Object.keys(correctAnswers);
    for (var pk = 0; pk < Math.min(soCauThiThu, allQKeys.length); pk++) {
      officialQuestionList.push(allQKeys[pk]);
      assignedQuestionIds[allQKeys[pk]] = true;
    }
  }

  // B4: Thi thật PHẢI có officialQuestionList — không fallback sang toàn bộ correctAnswers
  if (!isPractice && officialQuestionList.length === 0) {
    result = { success: false, message: '❌ Lỗi hệ thống: Chưa có danh sách câu hỏi chính thức cho học sinh!' };
    return result;
  }

  var totalQ = officialQuestionList.length > 0 ? officialQuestionList.length : data.answers.length;
  var pointPerQ = totalQ > 0 ? (isPractice ? 10 / totalQ : 4 / totalQ) : 0;
  var score = 0;
  var details = [];
  var answerResults = [];

  // ===== B6: Tạo answerMap từ payload client =====
  var answerMap = {};
  for (var j = 0; j < data.answers.length; j++) {
    var ans = data.answers[j];
    if (!ans || !ans.questionSTT) continue;
    var qSTTmap = ans.questionSTT.toString().trim();
    if (answerMap[qSTTmap] !== undefined) continue; // chống trùng
    var selMap = ans.selected ? ans.selected.toString().trim().toLowerCase() : '';
    if (selMap !== 'a' && selMap !== 'b' && selMap !== 'c' && selMap !== 'd') selMap = '';
    answerMap[qSTTmap] = selMap;
  }

  // ===== B6: Chấm theo officialQuestionList (không loop theo payload) =====
  for (var qi = 0; qi < officialQuestionList.length; qi++) {
    var qSTT = officialQuestionList[qi].toString().trim();
    var selected = answerMap[qSTT] !== undefined ? answerMap[qSTT] : '';
    var correct = correctAnswers[qSTT] !== undefined ? correctAnswers[qSTT] : '';
    var isCorrect = selected !== '' && correct !== '' && selected === correct;

    if (isCorrect) score += pointPerQ;

    var opts = questionOptions[qSTT] || {};
    details.push({
      stt: qSTT,
      noiDung: questionContents[qSTT] || '',
      dapAn: opts,
      dapAnDung: correct,
      dapAnHS: selected,
      ketQua: isCorrect ? 'Đúng' : 'Sai'
    });
    answerResults.push((selected || '').toUpperCase() + (isCorrect ? '✓' : '✗'));
  }

  score = Math.round(score * 100) / 100;

  var ss = SpreadsheetApp.openById(SHEET_ID); // dùng lại cho các sheet khác

  if (isPractice) {
    // ===== THI THỬ: ghi vào Ket_Qua_Thu =====
    var pSheet = getOrCreatePracticeSheet(ss);
    var pLastRow = pSheet.getLastRow();
    var attempts = countPracticeAttempts(ss, maHS);
    // Server-side validation: prevent bypassing max attempts
    if (attempts >= 5) {
      result = { success: false, message: '⚠️ Bạn đã hết 5 lượt thi thử! Không thể nộp thêm.' };
    } else {
      var currentPeriod = getCurrentExamPeriod();
      var rowData = [pLastRow, maHS, hoTen, lop, attempts + 1];
      for (var k = 0; k < 16; k++) {
        rowData.push(k < answerResults.length ? answerResults[k] : '');
      }
      rowData.push(score, new Date(), currentPeriod);
      pSheet.appendRow(rowData);

      // Tô màu đáp án đúng/sai (cột 6-21 = C1-C16)
      var newPRow = pSheet.getLastRow();
      for (var pc = 0; pc < 16; pc++) {
        if (pc < answerResults.length && answerResults[pc]) {
          var pCell = pSheet.getRange(newPRow, 6 + pc);
          if (answerResults[pc].includes('✓')) {
            pCell.setBackground('#d4edda');
          } else {
            pCell.setBackground('#f8d7da');
          }
        }
      }

      // Write to Firebase
      try {
        firebasePush('practiceResults', {
          maHS: maHS,
          hoTen: hoTen,
          lop: lop,
          luot: attempts + 1,
          answers: details,
          correctCount: details.filter(function(d) { return d.ketQua === 'Đúng'; }).length,
          wrongCount: details.filter(function(d) { return d.ketQua === 'Sai'; }).length,
          score: score,
          thoiGian: Date.now(),
          dotThi: currentPeriod,
          synced: true
        });
      } catch (fe) { Logger.log('Firebase practice write error: ' + fe); }

      result = {
        success: true,
        message: '🎯 Nộp bài thi thử thành công!',
        score: score,
        totalQuestions: totalQ,
        details: details,
        attempt: attempts + 1,
        maxAttempts: 5,
        isPractice: true
      };
    }
  } else {
    // ===== THI THẬT =====
    var currentPeriod = getCurrentExamPeriod();

    // ===== B8: CHỐNG NỘP TRÙNG tuyệt đối theo maHS + dotThi =====
    var periodKey = makeSafeFirebaseKey(currentPeriod || 'default');
    var resultKey = periodKey + '_' + maHS;
    try {
      var existingResult = firebaseGet('resultsByStudent/' + resultKey);
      if (existingResult) {
        result = { success: false, alreadySubmitted: true, message: '⚠️ Bạn đã nộp bài rồi! Không thể nộp lại.', driveFileId: existingResult.driveFileId || null };
        return result;
      }
    } catch (dupE) { Logger.log('resultsByStudent check error: ' + dupE); }

    // ===== CHỐNG NỘP TRÙNG: Kiểm tra trạng thái trên Sheet bên trong lock =====
    var hSheet = hSheetEarly; // dùng lại sheet đã đọc sớm
    var studentRowIndex = -1;
    for (var m = 1; m < hDataEarly.length; m++) {
      if (hDataEarly[m][1] && hDataEarly[m][1].toString().trim() === maHS) {
        // Kiểm tra trạng thái: nếu đã nộp (X) → từ chối
        if (hDataEarly[m][5].toString().trim().toUpperCase() === 'X') {
          result = { success: false, alreadySubmitted: true, message: '⚠️ Bạn đã nộp bài rồi! Không thể nộp lại.' };
          return result;
        }
        studentRowIndex = m;
        break;
      }
    }

    // B3: nếu không tìm thấy → fail (không append Ket_Qua)
    if (studentRowIndex < 1) {
      result = { success: false, message: '❌ Không tìm thấy học sinh trong danh sách!' };
      return result;
    }

    // 3. Ghi vào sheet Ket_Qua
    var rSheet = getSheetSafe(ss, 'Ket_Qua');
    if (!rSheet) {
      result = { success: false, message: '❌ Lỗi cấu hình: Không tìm thấy sheet Ket_Qua!' };
      return result;
    }
    var lastRow = rSheet.getLastRow();
    var newSTT = lastRow >= 1 ? lastRow : 1;
    var rowData = [newSTT, maHS, hoTen, lop];
    var maxCols = Math.max(8, answerResults.length);
    for (var k = 0; k < maxCols; k++) {
      rowData.push(k < answerResults.length ? answerResults[k] : '');
    }
    rowData.push(score);
    rowData.push(new Date());
    rowData.push(currentPeriod);
    rSheet.appendRow(rowData);

    // Color correct/wrong answer cells (columns 5-12 = index E-L)
    var newRow = rSheet.getLastRow();
    for (var c = 0; c < answerResults.length; c++) {
      if (answerResults[c]) {
        var cellRange = rSheet.getRange(newRow, 5 + c);
        if (answerResults[c].includes('✓')) {
          cellRange.setBackground('#d4edda');
        } else {
          cellRange.setBackground('#f8d7da');
        }
      }
    }

    // 4. Cập nhật trạng thái thi = X trên GGSheet
    hSheet.getRange(studentRowIndex + 1, 6).setValue('X');
    hSheet.getRange(studentRowIndex + 1, 8).setValue(new Date());

    // 5. Write to Firebase — CHIA THÀNH 2 PHẦN: CRITICAL & NON-CRITICAL

    // === 5A. CRITICAL: ghi trangThai:X ngay lập tức với retry ===
    // Đây là dữ liệu QUAN TRỌNG NHẤT — dashboard đọc trangThai để xác định trạng thái
    var fbCriticalWritten = false;
    for (var fbRetry = 0; fbRetry < 3; fbRetry++) {
      try {
        firebaseUpdate('students/' + maHS, {
          trangThai: 'X',
          thoiGianNB: Date.now()
        });
        fbCriticalWritten = true;
        break; // thành công → thoát vòng lặp
      } catch (fe) {
        Logger.log('Firebase critical write attempt ' + (fbRetry + 1) + ' failed: ' + fe);
        if (fbRetry < 2) Utilities.sleep(500); // chờ 500ms trước retry
      }
    }
    if (!fbCriticalWritten) {
      Logger.log('WARNING: trangThai:X NOT written to Firebase for maHS=' + maHS);
    }

    // === 5B. NON-CRITICAL: ghi kết quả chi tiết & xóa câu hỏi đã giao ===
    try {
      var correctCount = 0, wrongCount = 0;
      answerResults.forEach(function(ar) {
        if (ar && ar.includes('✓')) correctCount++;
        else if (ar && ar.length > 0) wrongCount++;
      });
      var resultPayload = {
        maHS: maHS,
        hoTen: hoTen,
        lop: lop,
        answers: details,
        correctCount: correctCount,
        wrongCount: wrongCount,
        score: score,
        thoiGian: Date.now(),
        dotThi: currentPeriod,
        synced: true
      };
      // B8: Ghi vào resultsByStudent (key cố định) để chống nộp trùng
      firebaseSet('resultsByStudent/' + resultKey, resultPayload);
      // Vẫn push vào results để dashboard cũ
      firebasePush('results', resultPayload);
      // Xóa câu hỏi đã giao + cập nhật metadata file
      firebaseUpdate('students/' + maHS, {
        assignedQuestions: null,
        assignedAnswerOrder: null,
        hasFile: false,
        filePending: !!(data.fileData && data.fileName),
        uploadFailed: false,
        uploadMissing: !data.fileData && !data.fileName
      });
    } catch (fe) { Logger.log('Firebase result write error: ' + fe); }

    // B1 FIX: Đánh dấu cần lưu Drive — truyền hoTen/lop đã validate từ Sheet, không dùng client data
    needDriveSave = true;
    // Tạo bản sao data với hoTen/lop chính xác từ server
    var safeData = {
      maHS: maHS,
      hoTen: hoTen,   // <- từ Sheet Hoc_Sinh, không phải client
      lop: lop,       // <- từ Sheet Hoc_Sinh, không phải client
      fileData: data.fileData || '',
      fileName: data.fileName || '',
      fileMimeType: data.fileMimeType || 'application/octet-stream',
      isPractice: isPractice
    };
    driveData = { data: safeData, details: details, score: score, maHS: maHS };

    // Ghi cờ uploadMissing nếu autoSubmit mà không có file
    if (isAutoSubmitted && (!data.fileData || !data.fileName)) {
      try {
        firebaseUpdate('students/' + maHS, {
          uploadMissing: true,
          uploadError: 'Hết giờ - HS chưa tải file thực hành'
        });
      } catch (fe3) { Logger.log('Flag uploadMissing error: ' + fe3); }
    }

    result = {
      success: true,
      message: '🎉 Nộp bài thành công!',
      score: score,
      isPractice: false,
      fileVerified: false  // B2 FIX: sẽ set true sau khi Drive save thành công
    };
  }

  } finally {
    lock.releaseLock(); // ← Release lock SỚM, trước khi lưu Drive
  }

  // 6. Lưu file lên Google Drive — NGOÀI lock
  if (needDriveSave && driveData) {
    try {
      var driveResult = saveFilesToDrive(driveData.data, driveData.details, driveData.score);
      // driveResult = { folderUrl, fileId } — fileId là ID file thực hành trên Drive
      var driveFolderUrl = driveResult ? driveResult.folderUrl : null;
      var driveFileId   = driveResult ? (driveResult.fileId || null) : null;
      if (driveFolderUrl) {
        // có fileId = chắc chắn đã upload file thực hành — chính xác 100%
        var hasFileUploaded = !!driveFileId;
        // Retry Firebase update để tránh mất trạng thái hasFile khi network chậm
        var driveUpdateOk = false;
        for (var dru = 0; dru < 3; dru++) {
          try {
            firebaseUpdate('students/' + driveData.maHS, {
              driveFolder: driveFolderUrl,
              driveFileId: driveFileId,       // ← Mới: lưu ID file để verify chính xác
              hasFile: hasFileUploaded,
              filePending: false,
              uploadFailed: false,
              uploadMissing: !hasFileUploaded && isAutoSubmitted,
              canResubmit: false
            });
            driveUpdateOk = true;
            break;
          } catch (fe2) {
            Logger.log('Save driveFolder attempt ' + (dru+1) + ' error: ' + fe2);
            if (dru < 2) Utilities.sleep(400);
          }
        }
        if (!driveUpdateOk) Logger.log('WARNING: driveFolder/hasFile NOT saved to Firebase for maHS=' + driveData.maHS);
        // Cập nhật resultsByStudent với driveFileId để check trùng sau cũng trả về đúng ID
        if (driveFileId) {
          try {
            firebaseUpdate('resultsByStudent/' + periodKey + '_' + driveData.maHS, { driveFileId: driveFileId });
          } catch (e2) { Logger.log('Update resultsByStudent driveFileId error: ' + e2); }
        }
        // B2 FIX: cập nhật result để client biết file đã lưu OK
        result.fileVerified = true;
        result.driveFolder = driveFolderUrl;
        result.driveFileId = driveFileId; // ← Trả về client
      }
    } catch (err) {
      Logger.log('Loi luu Drive: ' + err.toString());
      try {
        firebaseUpdate('students/' + driveData.maHS, {
          uploadFailed: true,
          hasFile: false,
          filePending: false,
          canResubmit: !!(driveData.data.fileData && driveData.data.fileName),
          uploadError: 'Drive save failed: ' + err.toString().substring(0, 200)
        });
      } catch (fe3) { /* silent */ }
      // B2 FIX: result.fileVerified vẫn false → client hiển thị cảnh báo retry
    }
  }

  return result;
}

// ============== CẬP NHẬT THÔNG TIN HỌC SINH ==============
function updateStudent(maHS, field, value) {
  if (!maHS || !field) return { success: false, message: 'Thiếu tham số!' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Hoc_Sinh');
  var data = sheet.getDataRange().getValues();
  // Map field to column index: STT=1, Ma_HS=2, Ho_Ten=3, Lop=4, Mat_Khau=5, Trang_Thai=6
  var colMap = { 'Ho_Ten': 3, 'Lop': 4, 'Mat_Khau': 5, 'Trang_Thai': 6 };
  var col = colMap[field];
  if (!col) return { success: false, message: 'Trường không hợp lệ!' };
  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() === maHS.trim()) {
      sheet.getRange(i + 1, col).setValue(value || '');
      // AUTO-SYNC: cập nhật Firebase tương ứng
      var fbFieldMap = { 'Ho_Ten': 'hoTen', 'Lop': 'lop', 'Mat_Khau': 'matKhau', 'Trang_Thai': 'trangThai' };
      var fbField = fbFieldMap[field];
      if (fbField) {
        try {
          var fbUpdate = {};
          fbUpdate[fbField] = value || '';
          firebaseUpdate('students/' + maHS.trim(), fbUpdate);
        } catch(fe) { Logger.log('Auto-sync FB updateStudent error: ' + fe); }
      }
      return { success: true, message: 'Cập nhật thành công!' };
    }
  }
  return { success: false, message: 'Không tìm thấy học sinh!' };
}

// ============== CÀI ĐẶT ==============
function getSettings() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Cai_Dat');
  if (!sheet) {
    sheet = ss.insertSheet('Cai_Dat');
    sheet.appendRow(['Key', 'Value']);
    sheet.appendRow(['practiceMode', 'false']);
  }
  var data = sheet.getDataRange().getValues();
  var settings = {};
  for (var i = 1; i < data.length; i++) {
    settings[data[i][0].toString().trim()] = data[i][1].toString().trim();
  }
  return { success: true, settings: settings };
}

function setSettings(key, value) {
  if (!key) return { success: false, message: 'Thiếu key!' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Cai_Dat');
  if (!sheet) {
    sheet = ss.insertSheet('Cai_Dat');
    sheet.appendRow(['Key', 'Value']);
  }
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value || '');
      return { success: true, message: 'Đã cập nhật!' };
    }
  }
  sheet.appendRow([key, value || '']);
  return { success: true, message: 'Đã thêm!' };
}

// ============== KẾT QUẢ HỌC SINH ==============
function getStudentResults(maHS) {
  if (!maHS) return { success: false, message: 'Thiếu mã HS!' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var rSheet = ss.getSheetByName('Ket_Qua');
  var rData = rSheet.getDataRange().getValues();
  var qSheet = ss.getSheetByName('Cau_Hoi');
  var qData = qSheet.getDataRange().getValues();
  var questionBank = {};
  for (var q = 1; q < qData.length; q++) {
    questionBank[qData[q][0]] = {
      noiDung: qData[q][1].toString(),
      dapAn: [qData[q][2].toString(), qData[q][3].toString(), qData[q][4].toString(), qData[q][5].toString()],
      dapAnDung: qData[q][6].toString().trim().toLowerCase()
    };
  }
  for (var i = 1; i < rData.length; i++) {
    if (rData[i][1].toString().trim() === maHS.trim()) {
      var answers = [];
      for (var j = 4; j < 12; j++) {
        answers.push(rData[i][j] ? rData[i][j].toString() : '');
      }
      var score = rData[i][12] ? parseFloat(rData[i][12]) : 0;
      return { success: true, score: score, answers: answers, questions: questionBank };
    }
  }
  return { success: false, message: 'Không tìm thấy kết quả.' };
}

// ============== BẢNG XẾP HẠNG THI THỬ ==============
function getPracticeLeaderboard() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = getOrCreatePracticeSheet(ss);
  var data = sheet.getDataRange().getValues();

  // Group by Ma_HS: lấy điểm cao nhất, tổng lần thi, điểm TB
  var map = {};
  for (var i = 1; i < data.length; i++) {
    var maHS = data[i][1] ? data[i][1].toString().trim() : '';
    if (!maHS) continue;
    var score = data[i][21] ? parseFloat(data[i][21]) : 0;
    if (!map[maHS]) {
      map[maHS] = {
        maHS: maHS,
        hoTen: data[i][2] ? data[i][2].toString() : '',
        lop: data[i][3] ? data[i][3].toString() : '',
        bestScore: score,
        totalScore: score,
        attempts: 1,
        lastTime: data[i][22] ? formatTime(data[i][22]) : ''
      };
    } else {
      map[maHS].attempts++;
      map[maHS].totalScore += score;
      if (score > map[maHS].bestScore) {
        map[maHS].bestScore = score;
        map[maHS].lastTime = data[i][22] ? formatTime(data[i][22]) : '';
      }
    }
  }

  // Convert to array, calculate avg, sort by bestScore desc
  var leaderboard = [];
  for (var key in map) {
    var entry = map[key];
    entry.avgScore = Math.round(entry.totalScore / entry.attempts * 100) / 100;
    entry.bestScore = Math.round(entry.bestScore * 100) / 100;
    leaderboard.push(entry);
  }
  leaderboard.sort(function(a, b) {
    if (b.bestScore !== a.bestScore) return b.bestScore - a.bestScore;
    return a.attempts - b.attempts; // Ít lần thi hơn xếp trước nếu cùng điểm
  });

  return { success: true, leaderboard: leaderboard };
}

// ============== LƯU FILE RIÊNG ==========================
// ============== HS NỘP LẠI FILE KHI DRIVE SAVE THẤT BẠI ==============
// Cho phép HS nộp lại file mà không cần nộp lại toàn bộ bài thi
function resubmitFile(data) {
  if (!data || !data.sessionToken) {
    return { success: false, message: '❌ Thiếu sessionToken!' };
  }
  // Validate session — chỉ student mới được resubmit
  var session = requireSession(data.sessionToken, ['student']);
  if (!session) return { success: false, message: '⛔ Phiên đăng nhập không hợp lệ!' };

  var maHS = session.userId;
  if (!validateUserCode(maHS)) return { success: false, message: '❌ Mã HS không hợp lệ!' };

  if (!data.fileData || !data.fileName) {
    return { success: false, message: '❌ Thiếu dữ liệu file!' };
  }

  // Validate file extension
  var allowedExts = ['.doc', '.docx', '.xls', '.xlsx'];
  var fn = data.fileName.toString().toLowerCase();
  var extValid = allowedExts.some(function(ext) {
    return fn.indexOf(ext, fn.length - ext.length) !== -1;
  });
  if (!extValid) {
    return { success: false, message: '⚠️ Sai định dạng! Chỉ chấp nhận .doc, .docx, .xls, .xlsx' };
  }

  // B3: kiểm tra size
  if (data.fileData.length > MAX_FILE_BASE64_LEN) {
    return { success: false, message: '⚠️ File quá lớn (> 6MB). Hãy nén file trước khi nộp!' };
  }

  // Kiểm tra HS đã nộp bài chưa (trangThai === X)
  var studentFb = firebaseGet('students/' + maHS);
  if (!studentFb || studentFb.trangThai !== 'X') {
    return { success: false, message: '❌ Học sinh chưa nộp bài hoặc không hợp lệ!' };
  }
  // Chỉ cho resubmit khi uploadFailed = true hoặc canResubmit = true
  if (!studentFb.uploadFailed && !studentFb.canResubmit) {
    return { success: false, message: '✅ File thực hành đã được lưu thành công, không cần nộp lại!' };
  }

  // Lấy hoTen/lop từ Firebase (đã validated trước khi ghi vào Firebase)
  var hoTen = studentFb.hoTen || '';
  var lop = studentFb.lop || '';

  // Thử lưu Drive
  try {
    var safeData = {
      maHS: maHS,
      hoTen: hoTen,
      lop: lop,
      fileData: data.fileData,
      fileName: data.fileName,
      fileMimeType: data.fileMimeType || 'application/octet-stream'
    };
    // Lấy details & score từ Firebase result
    var currentPeriod = getCurrentExamPeriod();
    var periodKey = makeSafeFirebaseKey(currentPeriod || 'default');
    var resultKey = periodKey + '_' + maHS;
    var existingResult = firebaseGet('resultsByStudent/' + resultKey) || {};
    var details = existingResult.answers || [];
    var score = existingResult.score || 0;

    var driveResult = saveFilesToDrive(safeData, details, score);
    var driveFolderUrl = driveResult ? driveResult.folderUrl : null;
    var resubmitFileId = driveResult ? (driveResult.fileId || null) : null;
    if (driveFolderUrl) {
      for (var rr = 0; rr < 3; rr++) {
        try {
          firebaseUpdate('students/' + maHS, {
            driveFolder: driveFolderUrl,
            driveFileId: resubmitFileId,  // ← Cập nhật ID file mới
            hasFile: !!resubmitFileId,
            uploadFailed: false,
            canResubmit: false,
            filePending: false,
            uploadError: null,
            resubmittedAt: Date.now()
          });
          break;
        } catch (fe) { if (rr < 2) Utilities.sleep(400); }
      }
      return { success: true, message: '✅ Nộp lại file thành công!', driveFolder: driveFolderUrl, driveFileId: resubmitFileId };
    }
    return { success: false, message: '❌ Không lấy được URL thư mục Drive!' };
  } catch (e) {
    Logger.log('resubmitFile error: ' + e);
    return { success: false, message: '❌ Lỗi lưu file: ' + e.toString().substring(0, 200) };
  }
}

// ============== KIỂM TRA TRẠNG THÁI FILE HS (cho retry UI) ==============
// ============== ĐỒNG BỘ TRẠNG THÁI: FIX HỌC SINH BỊ KẸT "ĐANG THI" ==============
// Đọc Sheet Hoc_Sinh → cập nhật Firebase trangThai cho tất cả HS đã nộp bài (Trang_Thai=X)
// ============== KIỂM TRA FILE DRIVE THỰC TẾ: FIX hasFile MÃI = FALSE ==============
// Giải thích: Drive đã có file nhưng Firebase hasFile=false do Firebase update thất bại
// Hàm này quét Drive thực tế thay vì tin Firebase

// ============ VERIFY BẰNG FILE ID — CHÍNH XÁC 100%, O(1) ============
// Đây là giải pháp chính xác nhất: lấy file.getId() khi createFile() thành công
// → lưu vào Firebase.driveFileId → verify bằng DriveApp.getFileById() là xong
function verifyByFileId(params) {
  if (!params || !params.maHS) return { success: false, message: 'Thiếu mã HS!' };
  var maHS = params.maHS.toString().trim();
  if (!validateUserCode(maHS)) return { success: false, message: 'Mã HS không hợp lệ!' };

  var s = firebaseGet('students/' + maHS);
  if (!s) return { success: false, message: 'Không tìm thấy HS ' + maHS + ' trong Firebase!' };

  // === Cách 1: Dùng driveFileId (nhanh nhất, chính xác 100%) ===
  if (s.driveFileId) {
    try {
      var file = DriveApp.getFileById(s.driveFileId);
      var exists = !file.isTrashed();
      var fileName = file.getName();
      if (exists && s.hasFile !== true) {
        // Firebase sai → sửa lại
        firebaseUpdate('students/' + maHS, { hasFile: true, uploadFailed: false, filePending: false, canResubmit: false });
      }
      return {
        success: true,
        hasFile: exists,
        driveFileId: s.driveFileId,
        fileName: fileName,
        message: exists
          ? '✅ File tồn tại: <strong>' + fileName + '</strong>'
          : '⚠️ File đã bị xóa khỏi Drive!'
      };
    } catch (e) {
      Logger.log('verifyByFileId getFileById error: ' + e);
      // ID không hợp lệ → fall through sang cách 2
    }
  }

  // === Cách 2: Fall back sang verifyStudentFile (scan folder theo tên) ===
  Logger.log('verifyByFileId: no driveFileId for ' + maHS + ', falling back to folder scan');
  var fallback = _verifySingleStudentDrive(maHS, s, getExamRootFolder());
  return {
    success: true,
    hasFile: fallback.hasFile,
    driveFileId: null,
    message: fallback.hasFile
      ? '✅ Tìm thấy file trong folder Drive (không có File ID).'
      : '⚠️ Không tìm thấy file thực hành — HS chưa nộp hoặc file bị xóa.'
  };
}

// Helper: lấy folderId từ URL Drive
function extractDriveFolderId(url) {

  if (!url) return null;
  var match = url.match(/\/folders\/([a-zA-Z0-9_-]+)/);
  return match ? match[1] : null;
}

// Helper: kiểm tra folder Drive có file thực hành (không phải .txt kết quả) không
function folderHasPracticeFile(folder) {
  try {
    var files = folder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      var name = f.getName().toLowerCase();
      if (!name.endsWith('.txt')) return true; // .doc .docx .xls .xlsx → file thực hành
    }
  } catch (e) { Logger.log('folderHasPracticeFile error: ' + e); }
  return false;
}

// Batch verify: kiểm tra TẤT CẢ học sinh đã nộp bài (trangThai=X)
function verifyDriveFiles() {
  try {
    var students = firebaseGet('students') || {};
    var examRoot = getExamRootFolder();
    var fixedCount = 0;
    var checkedCount = 0;

    for (var maHS in students) {
      var s = students[maHS];
      if ((s.trangThai || '').toString().trim().toUpperCase() !== 'X') continue; // chỉ check HS đã nộp
      checkedCount++;

      var result = _verifySingleStudentDrive(maHS, s, examRoot);
      if (result.updated) fixedCount++;
    }
    return { success: true, message: 'Đã kiểm tra ' + checkedCount + ' HS, cập nhật Firebase cho ' + fixedCount + ' HS!', fixedCount: fixedCount, checkedCount: checkedCount };
  } catch (e) {
    return { success: false, message: 'Lỗi: ' + e.toString() };
  }
}

// Verify 1 học sinh cụ thể
function verifyStudentFile(params) {
  if (!params || !params.maHS) return { success: false, message: 'Thiếu mã học sinh!' };
  var maHS = params.maHS.toString().trim();
  if (!validateUserCode(maHS)) return { success: false, message: 'Mã HS không hợp lệ!' };
  try {
    var s = firebaseGet('students/' + maHS);
    if (!s) return { success: false, message: 'Không tìm thấy HS ' + maHS + ' trong Firebase!' };
    var examRoot = getExamRootFolder();
    var result = _verifySingleStudentDrive(maHS, s, examRoot);
    if (result.updated) {
      return { success: true, message: '✅ Đã cập nhật: ' + maHS + ' có file thực hành trên Drive!', hasFile: result.hasFile, driveFolder: result.driveFolder };
    } else if (result.checked) {
      return { success: true, message: result.hasFile ? '✅ Đã có file (' + maHS + ')' : '⚠️ Không tìm thấy file thực hành trong Drive cho HS ' + maHS, hasFile: result.hasFile };
    }
    return { success: false, message: 'Không kiểm tra được HS ' + maHS };
  } catch (e) {
    return { success: false, message: 'Lỗi: ' + e.toString() };
  }
}

// Internal: kiểm tra 1 HS, cập nhật Firebase nếu cần
function _verifySingleStudentDrive(maHS, s, examRoot) {
  var result = { updated: false, checked: false, hasFile: false, driveFolder: s.driveFolder || '' };

  // === Cách 1: Dùng driveFolder URL sẵn có ===
  if (s.driveFolder) {
    var folderId = extractDriveFolderId(s.driveFolder);
    if (folderId) {
      try {
        var folder = DriveApp.getFolderById(folderId);
        var hasFile = folderHasPracticeFile(folder);
        result.checked = true;
        result.hasFile = hasFile;
        // Nếu Firebase sai → fix
        if (hasFile && s.hasFile !== true) {
          firebaseUpdate('students/' + maHS, {
            hasFile: true,
            uploadFailed: false,
            filePending: false,
            uploadMissing: false,
            canResubmit: false
          });
          result.updated = true;
          Logger.log('verifyDrive: fixed hasFile=true for ' + maHS);
        } else if (!hasFile && s.hasFile === true) {
          // Folder tồn tại nhưng trống (file bị xóa?)
          Logger.log('verifyDrive: folder exists but empty for ' + maHS);
        }
        return result;
      } catch (e) {
        Logger.log('verifyDrive getFolderById fail for ' + maHS + ': ' + e);
        // folderId không hợp lệ → thử tìm theo tên
      }
    }
  }

  // === Cách 2: Tìm folder theo tên HS trong Drive ===
  try {
    var safeLop = ((s.lop || 'Unknown')).toString().replace(/[\/\\:.]/g, '_');
    var classFolders = examRoot.getFoldersByName(safeLop);
    if (!classFolders.hasNext()) return result;
    var classFolder = classFolders.next();
    var studentFolderName = removeVietnameseTones(s.hoTen || '').replace(/\s+/g, '') + '_' + maHS;
    var studentFolders = classFolder.getFoldersByName(studentFolderName);
    if (!studentFolders.hasNext()) return result;
    var studentFolder = studentFolders.next();
    var hasFile = folderHasPracticeFile(studentFolder);
    var folderUrl = studentFolder.getUrl();
    result.checked = true;
    result.hasFile = hasFile;
    result.driveFolder = folderUrl;
    // Cập nhật Firebase
    firebaseUpdate('students/' + maHS, {
      driveFolder: folderUrl,
      hasFile: hasFile,
      uploadFailed: false,
      filePending: false,
      canResubmit: false
    });
    result.updated = true;
    Logger.log('verifyDrive: found by name, fixed ' + maHS + ' hasFile=' + hasFile);
  } catch (e) {
    Logger.log('verifyDrive search by name error for ' + maHS + ': ' + e);
  }
  return result;
}

function syncStudentStatus() {

  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hSheet = getSheetSafe(ss, 'Hoc_Sinh');
    if (!hSheet) return { success: false, message: 'Không tìm thấy sheet Hoc_Sinh!' };
    var hData = hSheet.getDataRange().getValues();
    var fixedCount = 0;
    var fbUpdates = {};
    for (var i = 1; i < hData.length; i++) {
      var maHSRow = hData[i][1] ? hData[i][1].toString().trim() : '';
      if (!maHSRow) continue;
      var trangThaiSheet = hData[i][5] ? hData[i][5].toString().trim().toUpperCase() : '';
      var thoiGianNBSheet = hData[i][7] || null;
      if (trangThaiSheet === 'X') {
        // Kiểm tra Firebase xem đã có trangThai:X chưa
        try {
          var fbStudent = firebaseGet('students/' + maHSRow);
          if (fbStudent) {
            var fbTrangThai = (fbStudent.trangThai || '').toString().trim().toUpperCase();
            if (fbTrangThai !== 'X') {
              // HS nộp bài trong Sheet nhưng Firebase chưa cập nhật → fix ngay
              fbUpdates['students/' + maHSRow + '/trangThai'] = 'X';
              if (thoiGianNBSheet) {
                fbUpdates['students/' + maHSRow + '/thoiGianNB'] = new Date(thoiGianNBSheet).getTime();
              }
              fixedCount++;
            }
          }
        } catch (fe) { Logger.log('syncStudentStatus check error for ' + maHSRow + ': ' + fe); }
      }
    }
    // Ghi tất cả updates một lần
    if (fixedCount > 0) {
      try {
        var db = FirebaseApp.getDatabaseByUrl(FIREBASE_URL, FIREBASE_SECRET);
        db.updateData('/', fbUpdates);
        Logger.log('syncStudentStatus: fixed ' + fixedCount + ' stuck students');
      } catch (batchErr) {
        // Fallback: ghi từng cái
        for (var path in fbUpdates) {
          try { firebaseSet(path.replace('/', ''), fbUpdates[path]); } catch (e2) { }
        }
      }
    }
    return { success: true, message: 'Đã đồng bộ! Sửa ' + fixedCount + ' học sinh bị kẹt.', fixedCount: fixedCount };
  } catch (e) {
    return { success: false, message: 'Lỗi đồng bộ: ' + e.toString() };
  }
}

// ============== GV ĐÁNH DẤU NỘP BÀI THỦ CÔNG (khi HS thực tế đã nộp nhưng hệ thống không ghi) ==============
function manualMarkSubmitted(params) {
  if (!params || !params.maHS) return { success: false, message: 'Thiếu mã học sinh!' };
  var maHS = params.maHS.toString().trim();
  if (!validateUserCode(maHS)) return { success: false, message: 'Mã học sinh không hợp lệ!' };
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var hSheet = getSheetSafe(ss, 'Hoc_Sinh');
    if (!hSheet) return { success: false, message: 'Không tìm thấy sheet Hoc_Sinh!' };
    var hData = hSheet.getDataRange().getValues();
    var rowIdx = -1;
    for (var i = 1; i < hData.length; i++) {
      if (hData[i][1] && hData[i][1].toString().trim() === maHS) { rowIdx = i; break; }
    }
    if (rowIdx < 1) return { success: false, message: 'Không tìm thấy học sinh ' + maHS + ' trong danh sách!' };
    // Ghi vào Sheet
    hSheet.getRange(rowIdx + 1, 6).setValue('X');
    if (!hData[rowIdx][7]) hSheet.getRange(rowIdx + 1, 8).setValue(new Date());
    // Ghi vào Firebase (critical)
    for (var retry = 0; retry < 3; retry++) {
      try {
        firebaseUpdate('students/' + maHS, {
          trangThai: 'X',
          thoiGianNB: Date.now(),
          manuallyMarked: true,
          markedAt: Date.now()
        });
        break;
      } catch (fe) {
        if (retry < 2) Utilities.sleep(500);
      }
    }
    return { success: true, message: '✅ Đã đánh dấu HS ' + maHS + ' là đã nộp bài!' };
  } catch (e) {
    return { success: false, message: 'Lỗi: ' + e.toString() };
  }
}

function getStudentFileStatus(maHS, sessionToken) {

  if (!maHS || !sessionToken) return { success: false, message: 'Thiếu tham số!' };
  var session = requireSession(sessionToken, ['student']);
  if (!session || session.userId !== (maHS || '').trim()) {
    return { success: false, message: '⛔ Không có quyền!' };
  }
  if (!validateUserCode(maHS)) return { success: false, message: '❌ Mã HS không hợp lệ!' };
  var student = firebaseGet('students/' + maHS.trim());
  if (!student) return { success: false, message: 'Không tìm thấy học sinh!' };
  return {
    success: true,
    trangThai: student.trangThai || '',
    hasFile: !!student.hasFile,
    uploadFailed: !!student.uploadFailed,
    canResubmit: !!student.canResubmit,
    uploadError: student.uploadError || '',
    driveFolder: student.driveFolder || '',
    resubmittedAt: student.resubmittedAt || null
  };
}

// ============== SEND NOTIFICATION (stub ACL) ==============
function sendNotificationAction(params) {
  // Đây là action yêu cầu teacher/admin — đã được ACL guard ở handleRequest
  if (!params || !params.message) return { success: false, message: 'Thiếu nội dung thông báo!' };
  try {
    firebasePush('notifications', {
      message: params.message,
      title: params.title || 'Thông báo',
      type: params.type || 'info',
      thoiGian: Date.now(),
      sender: params.sender || 'Hệ thống'
    });
    return { success: true, message: 'Đã gửi thông báo!' };
  } catch (e) {
    return { success: false, message: 'Lỗi gửi thông báo: ' + e.toString() };
  }
}

function saveFileOnly(data) {

  try {
    if (!data.fileData || !data.fileName) {
      return { success: false, message: 'Không có file để lưu!' };
    }
    // Validate file extension
    var allowedExts = ['.doc', '.docx', '.xls', '.xlsx'];
    var fileName = data.fileName.toString().toLowerCase();
    var extValid = false;
    for (var ae = 0; ae < allowedExts.length; ae++) {
      if (fileName.indexOf(allowedExts[ae], fileName.length - allowedExts[ae].length) !== -1) {
        extValid = true;
        break;
      }
    }
    if (!extValid) {
      return { success: false, message: '⚠️ Sai định dạng file! Chỉ chấp nhận Word (.doc, .docx) hoặc Excel (.xls, .xlsx).' };
    }
    var rootFolder = getExamRootFolder();
    var classFolder = getOrCreateFolder(rootFolder, data.lop);
    var studentFolderName = removeVietnameseTones(data.hoTen).replace(/\s+/g, '') + '_' + data.maHS;
    var studentFolder = getOrCreateFolder(classFolder, studentFolderName);
    if (data.fileData && data.fileName) {
      var fileBlob = Utilities.newBlob(
        Utilities.base64Decode(data.fileData),
        data.fileMimeType || 'application/octet-stream',
        data.fileName
      );
      studentFolder.createFile(fileBlob);
    }
    // Share folder and save URL to Firebase
    studentFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var folderUrl = studentFolder.getUrl();
    try {
      firebaseUpdate('students/' + data.maHS, { driveFolder: folderUrl });
    } catch (fe) { Logger.log('Save driveFolder (saveFileOnly) error: ' + fe); }
    return { success: true, message: 'Đã lưu file!', driveFolder: folderUrl };
  } catch(err) {
    return { success: false, message: 'Lỗi lưu file: ' + err.toString() };
  }
}

// ============== XÁC NHẬN FILE ĐÃ UPLOAD ==============
function verifyFileExists(maHS) {
  if (!maHS) return { success: false, message: 'Thiếu mã HS!' };
  try {
    var rootFolder = getExamRootFolder();
    // Search for student folder containing maHS
    var student = firebaseGet('students/' + maHS.trim());
    if (!student) return { success: false, hasFile: false, message: 'Không tìm thấy HS!' };
    var classFolder = getOrCreateFolder(rootFolder, student.lop || 'Unknown');
    var studentFolderName = removeVietnameseTones(student.hoTen || '').replace(/\s+/g, '') + '_' + maHS.trim();
    var folders = classFolder.getFoldersByName(studentFolderName);
    if (folders.hasNext()) {
      var studentFolder = folders.next();
      var files = studentFolder.getFiles();
      var hasFile = false;
      while (files.hasNext()) {
        var f = files.next();
        // Check for uploaded practice files (not the result text file)
        if (f.getName() !== 'KetQuaTracNghiem.txt') {
          hasFile = true;
          break;
        }
      }
      return { success: true, hasFile: hasFile };
    }
    return { success: true, hasFile: false };
  } catch(err) {
    return { success: false, hasFile: false, message: 'Lỗi kiểm tra: ' + err.toString() };
  }
}

// ============== UPLOAD TÀI LIỆU ÔN TẬP ==============
function uploadStudyFile(data) {
  try {
    if (!data.fileData || !data.fileName) {
      return { success: false, message: 'Thiếu dữ liệu file!' };
    }
    // Tạo/tìm thư mục TAI_LIEU_ON_TAP
    var rootFolder = getOrCreateFolder(DriveApp.getRootFolder(), 'TAI_LIEU_ON_TAP');
    
    // Tạo file từ base64
    var fileBlob = Utilities.newBlob(
      Utilities.base64Decode(data.fileData),
      data.fileMimeType || 'application/octet-stream',
      data.fileName
    );
    var file = rootFolder.createFile(fileBlob);
    
    // Chia sẻ công khai để HS có thể tải
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    var fileUrl = file.getUrl();
    var fileId = file.getId();
    
    return {
      success: true,
      fileUrl: fileUrl,
      fileId: fileId,
      fileName: data.fileName,
      message: 'Đã tải file lên thành công!'
    };
  } catch (err) {
    return { success: false, message: 'Lỗi tải file: ' + err.toString() };
  }
}

// ============== LƯU FILE LÊN GOOGLE DRIVE ==============
// B3 FIX: Giới hạn 8MB base64 (≈ 6MB file thực)
var MAX_FILE_BASE64_LEN = 8 * 1024 * 1024;

function saveFilesToDrive(data, details, score) {
  // B1 FIX: data.hoTen/data.lop đã được validate từ Sheet trước khi gọi hàm này
  var rootFolder = getExamRootFolder();

  // Tên lớp và tên thư mục HS được lấy từ server-validated data
  var safeLop = (data.lop || 'Unknown').toString().replace(/[/\\:.]/g, '_');
  var classFolder = getOrCreateFolder(rootFolder, safeLop);

  var studentFolderName = removeVietnameseTones(data.hoTen || '').replace(/\s+/g, '') + '_' + data.maHS;
  var studentFolder = getOrCreateFolder(classFolder, studentFolderName);

  // B3 FIX: Kiểm tra size trước khi decode
  var driveFileId = null;  // ← ID file thực hành (dùng để verify sau này)
  if (data.fileData && data.fileName) {
    if (data.fileData.length > MAX_FILE_BASE64_LEN) {
      Logger.log('File qua lon: ' + data.fileData.length + ' chars base64 (HS: ' + data.maHS + ')');
      throw new Error('File thực hành vượt quá 6MB. Hãy nén file trước khi nộp.');
    }
    var fileBlob = Utilities.newBlob(
      Utilities.base64Decode(data.fileData),
      data.fileMimeType || 'application/octet-stream',
      data.fileName
    );
    var uploadedFile = studentFolder.createFile(fileBlob);
    driveFileId = uploadedFile.getId(); // ← Lưu ID để verify chính xác 100%
    Logger.log('Drive file uploaded: ' + data.fileName + ' ID=' + driveFileId + ' HS=' + data.maHS);
  }

  // Tạo file kết quả text - ĐẦY ĐỦ CHI TIẾT
  var now = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
  
  // Lấy thêm các đáp án từ Cau_Hoi — LỌC THEO KHỐI LỚP HS
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var qSheet = ss.getSheetByName('Cau_Hoi');
  var qData = qSheet.getDataRange().getValues();
  var studentGrade = '';
  if (data.lop) {
    var gradeMatch = data.lop.toString().match(/^(\d+)/);
    if (gradeMatch) studentGrade = gradeMatch[1];
  }
  var questionBank = {};
  for (var q = 1; q < qData.length; q++) {
    var maCH = qData[q][0].toString().trim();
    var qKhoi = qData[q][7] ? qData[q][7].toString().trim() : '';
    // Lọc: chỉ lấy câu hỏi cùng khối hoặc câu chung
    if (studentGrade && qKhoi && qKhoi !== studentGrade) continue;
    questionBank[maCH] = {
      maCH: maCH,
      noiDung: qData[q][1].toString(),
      dapAn1: qData[q][2].toString(),
      dapAn2: qData[q][3].toString(),
      dapAn3: qData[q][4].toString(),
      dapAn4: qData[q][5].toString(),
      dapAnDung: qData[q][6].toString().trim().toLowerCase(),
      khoi: qKhoi
    };
  }

  var resultText = '';
  resultText += '══════════════════════════════════════════════════\n';
  resultText += '           KẾT QUẢ THI TRỰC TUYẾN\n';
  resultText += '══════════════════════════════════════════════════\n\n';
  resultText += '📅 Thời gian nộp bài : ' + now + '\n';
  resultText += '🆔 Mã học sinh       : ' + data.maHS + '\n';
  resultText += '👤 Họ và tên         : ' + data.hoTen + '\n';
  resultText += '🏫 Lớp               : ' + data.lop + '\n';
  var correctCount = 0;
  for (var dc = 0; dc < details.length; dc++) { if (details[dc].ketQua === 'Đúng') correctCount++; }
  var totalQ = details.length;
  resultText += '📊 Điểm trắc nghiệm : ' + score + ' / ' + (data.isPractice ? '10.0' : '4.0') + '\n';
  resultText += '   Số câu đúng      : ' + correctCount + ' / ' + totalQ + '\n\n';
  resultText += '──────────────────────────────────────────────────\n';
  resultText += '           CHI TIẾT CÂU TRẢ LỜI\n';
  resultText += '──────────────────────────────────────────────────\n\n';

  for (var i = 0; i < details.length; i++) {
    var d = details[i];
    var qInfo = questionBank[d.stt] || {};
    var labels = ['A', 'B', 'C', 'D'];
    var dapAnArr = [qInfo.dapAn1 || '', qInfo.dapAn2 || '', qInfo.dapAn3 || '', qInfo.dapAn4 || ''];
    var correctLabel = (d.dapAnDung || '').toUpperCase();
    var studentLabel = (d.dapAnHS || '').toUpperCase();
    var correctIdx = labels.indexOf(correctLabel);
    var studentIdx = labels.indexOf(studentLabel);
    
    resultText += '┌────────────────────────────────────────────────\n';
    resultText += '│ 📝 CÂU ' + (i + 1) + ' (STT gốc: ' + d.stt + ')\n';
    resultText += '│ ' + (d.noiDung || qInfo.noiDung || '') + '\n';
    resultText += '├────────────────────────────────────────────────\n';
    
    for (var a = 0; a < 4; a++) {
      var marker = '';
      if (a === correctIdx && a === studentIdx) marker = ' ← Đáp án đúng ✔ | HS chọn ✔';
      else if (a === correctIdx) marker = ' ← Đáp án đúng ✔';
      else if (a === studentIdx) marker = ' ← HS chọn ✘';
      resultText += '│    ' + labels[a] + '. ' + dapAnArr[a] + marker + '\n';
    }
    
    resultText += '├────────────────────────────────────────────────\n';
    resultText += '│ Đáp án đúng: ' + correctLabel;
    resultText += '  |  HS chọn: ' + studentLabel;
    resultText += '  |  Kết quả: ' + (d.ketQua === 'Đúng' ? '✅ ĐÚNG' : '❌ SAI') + '\n';
    resultText += '└────────────────────────────────────────────────\n\n';
  }

  resultText += '══════════════════════════════════════════════════\n';

  var resultBlob = Utilities.newBlob(resultText, 'text/plain; charset=utf-8', 'KetQuaTracNghiem.txt');
  studentFolder.createFile(resultBlob);

  // Share folder and return object {folderUrl, fileId}
  studentFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { folderUrl: studentFolder.getUrl(), fileId: driveFileId };
}

// ============== DASHBOARD - GIÁO VIÊN ==============
function getDashboard(lopFilter) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var hSheet = ss.getSheetByName('Hoc_Sinh');
  var hData = hSheet.getDataRange().getValues();

  var students = [];
  var classSet = {};
  var totalStudents = 0;
  var loggedIn = 0;
  var submitted = 0;

  for (var i = 1; i < hData.length; i++) {
    var lop = hData[i][3].toString().trim();
    if (lop) classSet[lop] = true;

    if (lopFilter && lop !== lopFilter) continue;

    totalStudents++;
    var trangThai = hData[i][5].toString().trim().toUpperCase();
    var thoiGianDN = hData[i][6] || '';
    var thoiGianNB = hData[i][7] || '';

    var status = 'Chưa đăng nhập';
    var statusCode = 0;
    if (trangThai === 'X') {
      status = 'Đã nộp bài';
      statusCode = 2;
      submitted++;
      loggedIn++;
    } else if (thoiGianDN) {
      status = 'Đang làm bài';
      statusCode = 1;
      loggedIn++;
    }

    students.push({
      stt: hData[i][0],
      maHS: hData[i][1].toString(),
      hoTen: hData[i][2].toString(),
      lop: lop,
      trangThai: status,
      statusCode: statusCode,
      thoiGianDN: thoiGianDN ? formatTime(thoiGianDN) : '—',
      thoiGianNB: thoiGianNB ? formatTime(thoiGianNB) : '—'
    });
  }

  return {
    success: true,
    stats: {
      total: totalStudents,
      loggedIn: loggedIn,
      submitted: submitted,
      pending: totalStudents - submitted,
      notLoggedIn: totalStudents - loggedIn
    },
    students: students,
    classes: Object.keys(classSet).sort()
  };
}

// ============== KẾT QUẢ - GIÁO VIÊN ==============
function getResults(lopFilter) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var rSheet = ss.getSheetByName('Ket_Qua');
  var rData = rSheet.getDataRange().getValues();

  var results = [];
  var totalScore = 0;
  var count = 0;
  var scores = [];

  for (var i = 1; i < rData.length; i++) {
    var lop = rData[i][3] ? rData[i][3].toString().trim() : '';
    if (lopFilter && lop !== lopFilter) continue;

    var answers = [];
    for (var j = 4; j < 12; j++) {
      answers.push(rData[i][j] ? rData[i][j].toString() : '');
    }

    var score = rData[i][12] ? parseFloat(rData[i][12]) : 0;
    totalScore += score;
    count++;
    scores.push(score);

    results.push({
      stt: rData[i][0],
      maHS: rData[i][1] ? rData[i][1].toString() : '',
      hoTen: rData[i][2] ? rData[i][2].toString() : '',
      lop: lop,
      answers: answers,
      score: score,
      thoiGian: rData[i][13] ? formatTime(rData[i][13]) : ''
    });
  }

  // Thống kê điểm
  var scoreDistribution = { '0-1': 0, '1-2': 0, '2-3': 0, '3-4': 0 };
  for (var s = 0; s < scores.length; s++) {
    if (scores[s] <= 1) scoreDistribution['0-1']++;
    else if (scores[s] <= 2) scoreDistribution['1-2']++;
    else if (scores[s] <= 3) scoreDistribution['2-3']++;
    else scoreDistribution['3-4']++;
  }

  return {
    success: true,
    results: results,
    stats: {
      count: count,
      avgScore: count > 0 ? (totalScore / count).toFixed(2) : 0,
      maxScore: count > 0 ? Math.max.apply(null, scores) : 0,
      minScore: count > 0 ? Math.min.apply(null, scores) : 0,
      distribution: scoreDistribution
    }
  };
}

// ============== XUẤT KẾT QUẢ EXCEL ==============
function exportResultsCSV(lopFilter) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var rSheet = ss.getSheetByName('Ket_Qua');
  var rData = rSheet.getDataRange().getValues();
  var currentPeriod = getCurrentExamPeriod();

  // Headers
  var csvRows = [];
  csvRows.push(['STT', 'Mã HS', 'Họ Tên', 'Lớp', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'Số câu đúng', 'Số câu sai', 'Điểm', 'Thời gian nộp', 'Đợt Thi'].join(','));

  var count = 0;
  for (var i = 1; i < rData.length; i++) {
    var lop = rData[i][3] ? rData[i][3].toString().trim() : '';
    if (lopFilter && lop !== lopFilter) continue;
    // Lọc theo đợt thi hiện tại: cột Dot_Thi là cột cuối cùng
    var rowCols = rData[i].length;
    var rowDotThi = (rData[i][rowCols - 1] || '').toString().trim();
    if (currentPeriod !== '') {
      if (rowDotThi !== currentPeriod) continue;
    }
    count++;
    var row = [];
    row.push(count);
    row.push('"' + (rData[i][1] || '').toString().replace(/"/g, '""') + '"');
    row.push('"' + (rData[i][2] || '').toString().replace(/"/g, '""') + '"');
    row.push('"' + lop + '"');
    // Answers C1-C8 + count correct/wrong
    var correctCount = 0;
    var wrongCount = 0;
    for (var j = 4; j < 12; j++) {
      var ans = rData[i][j] ? rData[i][j].toString() : '';
      row.push('"' + ans.replace(/"/g, '""') + '"');
      if (ans.includes('✓')) correctCount++;
      else if (ans.length > 0) wrongCount++;
    }
    // Correct / Wrong counts
    row.push(correctCount);
    row.push(wrongCount);
    // Score
    row.push(rData[i][12] ? parseFloat(rData[i][12]) : 0);
    // Time
    var time = rData[i][13] ? formatTime(rData[i][13]) : '';
    row.push('"' + time + '"');
    // Đợt Thi
    row.push('"' + rowDotThi.replace(/"/g, '""') + '"');
    csvRows.push(row.join(','));
  }

  if (count === 0) {
    return { success: false, message: 'Không có dữ liệu kết quả để xuất!' };
  }

  // BOM + CSV content for Excel UTF-8
  var csvContent = '\uFEFF' + csvRows.join('\n');
  var periodSuffix = currentPeriod ? '_' + currentPeriod.replace(/[^a-zA-Z0-9_\-]/g, '') : '';
  var fileName = 'KetQuaThi' + (lopFilter ? '_' + lopFilter : '') + periodSuffix + '_' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyyMMdd_HHmmss') + '.csv';

  // Save to Drive
  var rootFolder = getExamRootFolder();
  var blob = Utilities.newBlob(csvContent, 'text/csv; charset=utf-8', fileName);
  var file = rootFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    success: true,
    message: 'Đã xuất ' + count + ' kết quả' + (currentPeriod ? ' (đợt: ' + currentPeriod + ')' : '') + '!',
    downloadUrl: file.getDownloadUrl(),
    fileName: fileName
  };
}

// ============== XÓA DỮ LIỆU ĐỢT THI HIỆN TẠI (P0-6) ==============
function clearAllData(sessionToken, confirmPeriod) {
  // E2/F1: Tự kiểm tra quyền Admin bên trong hàm
  if (!requireAdminSession(sessionToken)) {
    return { success: false, message: '⛔ Không có quyền Admin! Vui lòng đăng nhập lại.' };
  }
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {
    return { success: false, message: 'Hệ thống đang bận!' };
  }
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var currentPeriod = getCurrentExamPeriod();
    // Chặn xóa khi chưa đặt tên đợt thi
    if (!currentPeriod || currentPeriod === '') {
      return { success: false, message: '⚠️ Chưa đặt tên đợt thi! Hãy cấu hình đợt thi trước khi xóa dữ liệu.' };
    }
    // Xác nhận đợt thi khớp — chống xóa nhầm
    if (!confirmPeriod || confirmPeriod !== currentPeriod) {
      return { success: false, message: '⚠️ Xác nhận đợt thi không khớp! Vui lòng nhập đúng tên đợt thi hiện tại.' };
    }

    var failedSteps = [];
    var completedSteps = [];
    var deletedReal = 0, deletedPractice = 0;

    // 0. Backup Ket_Qua và Ket_Qua_Thu trước khi xóa (F1)
    var bkTime = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    var bkName = 'Backup_' + currentPeriod.substring(0, 20) + '_' + bkTime;
    try {
      _backupBeforeClear(ss, currentPeriod, bkName);
      completedSteps.push('backup');
    } catch(be) {
      failedSteps.push({ step: 'backup', error: be.toString() });
    }

    // 1. Xóa kết quả thi thật đợt hiện tại trên Sheet
    try {
      var kqSheet = ss.getSheetByName('Ket_Qua');
      if (kqSheet && kqSheet.getLastRow() > 1) {
        var kqData = kqSheet.getDataRange().getValues();
        for (var i = kqData.length - 1; i >= 1; i--) {
          var rowCols = kqData[i].length;
          var rowDotThi = (kqData[i][rowCols - 1] || '').toString().trim();
          if (rowDotThi === currentPeriod) {
            kqSheet.deleteRow(i + 1);
            deletedReal++;
          }
        }
      }
      completedSteps.push('delete_sheet_results');
    } catch(e) {
      failedSteps.push({ step: 'delete_sheet_results', error: e.toString() });
    }

    // 2. Xóa kết quả thi thử đợt hiện tại trên Sheet
    try {
      var kqtSheet = ss.getSheetByName('Ket_Qua_Thu');
      if (kqtSheet && kqtSheet.getLastRow() > 1) {
        var kqtData = kqtSheet.getDataRange().getValues();
        for (var i = kqtData.length - 1; i >= 1; i--) {
          var rowCols2 = kqtData[i].length;
          var rowDotThi2 = (kqtData[i][rowCols2 - 1] || '').toString().trim();
          if (rowDotThi2 === currentPeriod) {
            kqtSheet.deleteRow(i + 1);
            deletedPractice++;
          }
        }
      }
      completedSteps.push('delete_sheet_practice');
    } catch(e) {
      failedSteps.push({ step: 'delete_sheet_practice', error: e.toString() });
    }

    // 3. Reset trạng thái thi + thời gian HS
    try {
      var hSheet = ss.getSheetByName('Hoc_Sinh');
      var hData = hSheet.getDataRange().getValues();
      for (var i = 1; i < hData.length; i++) {
        hSheet.getRange(i + 1, 6).setValue(''); // Trang_Thai
        hSheet.getRange(i + 1, 7).setValue(''); // Thoi_Gian_DN
        hSheet.getRange(i + 1, 8).setValue(''); // Thoi_Gian_NB
      }
      completedSteps.push('reset_student_status');
    } catch(e) {
      failedSteps.push({ step: 'reset_student_status', error: e.toString() });
    }

    // 4. Xóa folder Drive CHỈ của đợt thi hiện tại
    try {
      var rootFolders = DriveApp.getRootFolder().getFoldersByName(ROOT_FOLDER_NAME);
      if (rootFolders.hasNext()) {
        var rootFolder = rootFolders.next();
        var periodFolders = rootFolder.getFoldersByName(currentPeriod);
        while (periodFolders.hasNext()) {
          periodFolders.next().setTrashed(true);
        }
      }
      completedSteps.push('delete_drive');
    } catch(de) {
      failedSteps.push({ step: 'delete_drive', error: de.toString() });
    }

    // 5. Xóa kết quả Firebase CHỈ đợt hiện tại
    try {
      var fbResults = firebaseGet('results') || {};
      for (var rKey in fbResults) {
        var rDotThi = (fbResults[rKey].dotThi || '').toString().trim();
        if (rDotThi === currentPeriod) {
          firebaseDelete('results/' + rKey);
        }
      }
      var fbPractice = firebaseGet('practiceResults') || {};
      for (var pKey in fbPractice) {
        var pDotThi = (fbPractice[pKey].dotThi || '').toString().trim();
        if (pDotThi === currentPeriod) {
          firebaseDelete('practiceResults/' + pKey);
        }
      }
      // Xóa resultsByStudent đợt này
      var safePeriodKey = makeSafeFirebaseKey(currentPeriod);
      var fbRBS = firebaseGet('resultsByStudent') || {};
      for (var rbsKey in fbRBS) {
        if (rbsKey.indexOf(safePeriodKey + '_') === 0) {
          firebaseDelete('resultsByStudent/' + rbsKey);
        }
      }
      // Reset student status on Firebase
      var fbStudents = firebaseGet('students') || {};
      var resetObj = {};
      for (var maHS in fbStudents) {
        resetObj[maHS + '/trangThai'] = '';
        resetObj[maHS + '/thoiGianDN'] = null;
        resetObj[maHS + '/thoiGianNB'] = null;
      }
      if (Object.keys(resetObj).length > 0) {
        firebaseUpdate('students', resetObj);
      }
      completedSteps.push('delete_firebase');
    } catch(fe) {
      failedSteps.push({ step: 'delete_firebase', error: fe.toString() });
    }

    // P0-6: Trả partialSuccess nếu có bước lỗi
    if (failedSteps.length > 0) {
      return {
        success: false,
        partialSuccess: true,
        message: '⚠️ Có bước xóa thất bại. Cần kiểm tra và chạy lại.',
        completedSteps: completedSteps,
        failedSteps: failedSteps
      };
    }

    return { success: true, message: '✅ Đã xóa dữ liệu đợt "' + currentPeriod + '": ' + deletedReal + ' kết quả thi thật, ' + deletedPractice + ' kết quả thi thử + file Drive + Firebase!' };
  } finally {
    lock.releaseLock();
  }
}

// F1 FIX: Backup cả Ket_Qua và Ket_Qua_Thu trước khi xóa
function _backupBeforeClear(ss, currentPeriod) {
  var bkName = 'Backup_' + currentPeriod.substring(0, 20) + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
  var bkSheet = ss.getSheetByName(bkName);
  if (!bkSheet) bkSheet = ss.insertSheet(bkName);

  // Header dòng đầu
  bkSheet.appendRow(['Backup_Time', 'Source_Sheet', 'Dot_Thi', 'Ghi_Chu']);
  bkSheet.appendRow([new Date(), 'KET_QUA + KET_QUA_THU', currentPeriod, 'Auto backup before clearAllData']);

  // Backup Ket_Qua
  var kqSheet = ss.getSheetByName('Ket_Qua');
  if (kqSheet && kqSheet.getLastRow() > 1) {
    bkSheet.appendRow(['--- BACKUP KET_QUA ---', 'Đợt: ' + currentPeriod]);
    var kqData = kqSheet.getDataRange().getValues();
    for (var i = 0; i < kqData.length; i++) {
      var rowCols = kqData[i].length;
      var rowDotThi = i === 0 ? '' : (kqData[i][rowCols - 1] || '').toString().trim();
      if (i === 0 || rowDotThi === currentPeriod) {
        bkSheet.appendRow(kqData[i]);
      }
    }
  }

  // Backup Ket_Qua_Thu
  var kqtSheet = ss.getSheetByName('Ket_Qua_Thu');
  if (kqtSheet && kqtSheet.getLastRow() > 1) {
    bkSheet.appendRow(['--- BACKUP KET_QUA_THU ---', 'Đợt: ' + currentPeriod]);
    var kqtData = kqtSheet.getDataRange().getValues();
    for (var j = 0; j < kqtData.length; j++) {
      var rowCols2 = kqtData[j].length;
      var rowDotThi2 = j === 0 ? '' : (kqtData[j][rowCols2 - 1] || '').toString().trim();
      if (j === 0 || rowDotThi2 === currentPeriod) {
        bkSheet.appendRow(kqtData[j]);
      }
    }
  }
}

// ============== RESET CHO ĐỢT THI MỚI (AN TOÀN) ==============
// Chỉ reset trạng thái HS, KHÔNG xóa kết quả, KHÔNG xóa file Drive
function resetForNewPeriod(sessionToken) {
  // E2/F2: Tự kiểm tra quyền Admin bên trong hàm
  if (!requireAdminSession(sessionToken)) {
    return { success: false, message: '⛔ Không có quyền Admin! Vui lòng đăng nhập lại.' };
  }
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {
    return { success: false, message: 'Hệ thống đang bận!' };
  }
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);
    var failedSteps = [];
    var completedSteps = [];

    // 1. Reset trạng thái HS trên Sheet
    try {
      var hSheet = ss.getSheetByName('Hoc_Sinh');
      var hData = hSheet.getDataRange().getValues();
      var studentCount = 0;
      for (var i = 1; i < hData.length; i++) {
        hSheet.getRange(i + 1, 6).setValue('');
        hSheet.getRange(i + 1, 7).setValue('');
        hSheet.getRange(i + 1, 8).setValue('');
        studentCount++;
      }
      completedSteps.push('reset_sheet');
    } catch(e) {
      failedSteps.push({ step: 'reset_sheet', error: e.toString() });
      var studentCount = 0;
    }

    // 2. Reset trạng thái HS + cờ mở thi trên Firebase
    try {
      var fbStudents = firebaseGet('students') || {};
      var resetObj = {};
      for (var maHS in fbStudents) {
        resetObj[maHS + '/trangThai'] = '';
        resetObj[maHS + '/thoiGianDN'] = null;
        resetObj[maHS + '/thoiGianNB'] = null;
        resetObj[maHS + '/moThiThat'] = false;
        resetObj[maHS + '/moThiThu'] = false;
        resetObj[maHS + '/lastSeen'] = null;
        resetObj[maHS + '/cheatingCount'] = 0;
        resetObj[maHS + '/assignedQuestions'] = null;
        resetObj[maHS + '/assignedAnswerOrder'] = null;
        resetObj[maHS + '/savedAnswers'] = null;
        resetObj[maHS + '/driveFolder'] = null;
        resetObj[maHS + '/uploadFailed'] = false;
        resetObj[maHS + '/uploadMissing'] = false;
        resetObj[maHS + '/uploadError'] = null;
        resetObj[maHS + '/hasFile'] = false;
        resetObj[maHS + '/isLocked'] = false;
        resetObj[maHS + '/lockedAt'] = null;
        resetObj[maHS + '/filePending'] = false;
        resetObj[maHS + '/tags'] = null;
        resetObj[maHS + '/ghiChu'] = null;
        resetObj[maHS + '/loginRequest'] = null;
      }
      if (Object.keys(resetObj).length > 0) {
        firebaseUpdate('students', resetObj);
      }
      completedSteps.push('reset_firebase');
    } catch(fe) {
      failedSteps.push({ step: 'reset_firebase', error: fe.toString() });
    }

    if (failedSteps.length > 0) {
      return {
        success: false, partialSuccess: true,
        message: '⚠️ Reset có bước thất bại. Cần kiểm tra lại.',
        completedSteps: completedSteps, failedSteps: failedSteps
      };
    }

    return {
      success: true,
      message: '✅ Đã reset trạng thái ' + studentCount + ' học sinh. Kết quả thi + file Drive các đợt trước được GIỮ NGUYÊN. Hãy đặt tên đợt thi mới và cấu hình đề thi trước khi mở thi!'
    };
  } finally {
    lock.releaseLock();
  }
}

// ============== LẤY URL THƯ MỤC BÀI THI TRÊN DRIVE ==============
function getExamFolderUrl() {
  try {
    var folder = getExamRootFolder();
    var url = folder.getUrl();
    var name = folder.getName();
    // Đếm số subfolder (lớp) và file
    var subFolderCount = 0;
    var subFolders = folder.getFolders();
    while (subFolders.hasNext()) { subFolders.next(); subFolderCount++; }
    var fileCount = 0;
    var files = folder.getFiles();
    while (files.hasNext()) { files.next(); fileCount++; }
    return {
      success: true,
      url: url,
      name: name,
      subFolders: subFolderCount,
      files: fileCount
    };
  } catch (e) {
    return { success: false, message: 'Lỗi đọc thư mục: ' + e.toString() };
  }
}

// ============== TẢI FILE HỌC SINH (GIÁO VIÊN) ==============

/**
 * Kiểm tra quyền GV với lớp — trả true nếu có quyền
 */
function _checkTeacherPermission(lop, sessionToken) {
  var session = requireTeacherOrAdminSession(sessionToken);
  if (!session) return false;
  if (session.role === 'admin') return true;
  var maGV = session.userId;
  if (!validateUserCode(maGV)) return false;
  var teacher = firebaseGet('teachers/' + maGV);
  if (!teacher || !teacher.lopPhuTrach) return false;
  var allowed = teacher.lopPhuTrach.split(',').map(function(c) { return c.trim(); });
  return allowed.indexOf(lop.toString().trim()) >= 0;
}

/**
 * Thu thập tất cả file Blob từ 1 folder (đệ quy 1 cấp — student folder)
 * Trả về mảng Blob, mỗi blob đặt tên = subFolderName/fileName
 */
function _collectFilesFromFolder(folder, prefix) {
  var blobs = [];
  var files = folder.getFiles();
  while (files.hasNext()) {
    var file = files.next();
    var blob = file.getBlob();
    blob.setName(prefix ? (prefix + '/' + file.getName()) : file.getName());
    blobs.push(blob);
  }
  return blobs;
}

/**
 * Tải tất cả file của cả lớp — nén ZIP
 * @param {string} lop - Tên lớp (vd: '7A1')
 * @param {string} maGV - Mã giáo viên (để kiểm tra quyền)
 * @param {string|boolean} isAdmin - Có phải admin không
 */
function downloadClassFiles(lop, sessionToken) {
  if (!lop) return { success: false, message: 'Vui lòng chọn lớp!' };
  if (!validateClassName(lop)) return { success: false, message: 'Tên lớp không hợp lệ!' };
  if (!_checkTeacherPermission(lop, sessionToken)) {
    return { success: false, message: 'Bạn không có quyền tải file lớp ' + lop + '!' };
  }
  
  var MAX_ZIP_FILES = 80;
  var MAX_ZIP_BYTES = 25 * 1024 * 1024;
  
  try {
    var rootFolder = getExamRootFolder();
    var classFolders = rootFolder.getFoldersByName(lop);
    if (!classFolders.hasNext()) {
      return { success: false, message: 'Chưa có thư mục lớp ' + lop + ' trên Drive!' };
    }
    var classFolder = classFolders.next();
    
    var allBlobs = [];
    var studentCount = 0;
    var fileCount = 0;
    var totalSize = 0;
    var studentFolders = classFolder.getFolders();
    
    while (studentFolders.hasNext()) {
      var sFolder = studentFolders.next();
      var sName = sFolder.getName();
      var sFiles = sFolder.getFiles();
      var hasFiles = false;
      
      while (sFiles.hasNext()) {
        var f = sFiles.next();
        totalSize += f.getSize();
        fileCount++;
        // Kiểm tra giới hạn ZIP
        if (fileCount > MAX_ZIP_FILES || totalSize > MAX_ZIP_BYTES) {
          return {
            success: false, fallback: true,
            message: 'File quá nhiều hoặc quá lớn (' + fileCount + ' file, ' + Math.round(totalSize/1024/1024) + 'MB). Vui lòng tải thủ công từ Drive.',
            folderUrl: classFolder.getUrl()
          };
        }
        var blob = f.getBlob();
        blob.setName(sName + '/' + f.getName());
        allBlobs.push(blob);
        hasFiles = true;
      }
      if (hasFiles) studentCount++;
    }
    
    if (allBlobs.length === 0) {
      return { success: false, message: 'Lớp ' + lop + ' chưa có file bài làm nào!' };
    }
    
    var zipName = 'BaiLam_' + lop + '_' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyyMMdd_HHmm') + '.zip';
    var zipBlob = Utilities.zip(allBlobs, zipName);
    var zipFile = rootFolder.createFile(zipBlob);
    zipFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Lưu ZIP ID vào cache để cleanupDownloadZip verify
    var session = getSession(sessionToken);
    CacheService.getScriptCache().put('tmpzip:' + zipFile.getId(), JSON.stringify({
      createdAt: Date.now(),
      createdBy: session ? session.userId : '',
      createdByRole: session ? session.role : '',
      type: 'class'
    }), 60 * 60);
    
    return {
      success: true,
      message: 'Đã nén ' + fileCount + ' file của ' + studentCount + ' học sinh lớp ' + lop,
      downloadUrl: zipFile.getDownloadUrl(),
      zipFileId: zipFile.getId(),
      fileName: zipName,
      studentCount: studentCount,
      fileCount: fileCount
    };
  } catch (err) {
    try {
      var rf = getExamRootFolder();
      var cf = rf.getFoldersByName(lop);
      if (cf.hasNext()) {
        return {
          success: false, fallback: true,
          message: 'Không thể nén file (quá lớn hoặc timeout). Mở Google Drive để tải thủ công.',
          folderUrl: cf.next().getUrl()
        };
      }
    } catch (e2) { /* silent */ }
    return { success: false, message: 'Lỗi tải file: ' + err.toString() };
  }
}

/**
 * Tải file của 1 học sinh cụ thể
 * @param {string} lop - Tên lớp
 * @param {string} maHS - Mã học sinh
 * @param {string} maGV - Mã giáo viên
 * @param {string|boolean} isAdmin - Có phải admin không
 */
function downloadStudentFiles(lop, maHS, sessionToken) {
  if (!lop || !maHS) return { success: false, message: 'Thiếu thông tin lớp hoặc mã HS!' };
  if (!validateClassName(lop)) return { success: false, message: 'Tên lớp không hợp lệ!' };
  if (!validateUserCode(maHS)) return { success: false, message: 'Mã HS không hợp lệ!' };
  if (!_checkTeacherPermission(lop, sessionToken)) {
    return { success: false, message: 'Bạn không có quyền tải file lớp ' + lop + '!' };
  }
  
  try {
    var rootFolder = getExamRootFolder();
    var classFolders = rootFolder.getFoldersByName(lop);
    if (!classFolders.hasNext()) return { success: false, message: 'Chưa có thư mục lớp ' + lop + '!' };
    var classFolder = classFolders.next();
    
    // P0-8: Regex match cuối chuỗi thay vì indexOf (chống match nhầm HS001 vs HS0011)
    var pattern = new RegExp('_' + escapeRegex(maHS.trim()) + '$');
    var studentFolder = null;
    var subFolders = classFolder.getFolders();
    while (subFolders.hasNext()) {
      var sf = subFolders.next();
      if (pattern.test(sf.getName())) {
        studentFolder = sf;
        break;
      }
    }
    
    if (!studentFolder) {
      return { success: false, message: 'Chưa có thư mục bài làm của HS ' + maHS + '!' };
    }
    
    var blobs = [];
    var firstFile = null;
    var totalBytes = 0;
    var fileCount = 0;
    var MAX_ZIP_FILES = 80;
    var MAX_ZIP_BYTES = 25 * 1024 * 1024;
    var files = studentFolder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      fileCount++;
      totalBytes += f.getSize();
      if (fileCount > MAX_ZIP_FILES || totalBytes > MAX_ZIP_BYTES) {
        return {
          success: false, fallback: true,
          message: 'File quá nhiều hoặc quá lớn, vui lòng tải thủ công từ Drive.',
          folderUrl: studentFolder.getUrl()
        };
      }
      if (!firstFile) firstFile = f;
      blobs.push(f.getBlob());
    }

    if (blobs.length === 0) {
      return { success: false, message: 'HS ' + maHS + ' chưa có file bài làm!' };
    }
    
    // Nếu chỉ 1 file → trả link trực tiếp (dùng firstFile đã lưu, không đọc lại)
    if (blobs.length === 1 && firstFile) {
      firstFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return {
        success: true,
        message: 'Tải file: ' + firstFile.getName(),
        downloadUrl: firstFile.getDownloadUrl(),
        fileName: firstFile.getName(),
        fileCount: 1,
        isDirect: true
      };
    }
    
    var zipName = 'BaiLam_' + maHS.trim() + '_' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyyMMdd_HHmm') + '.zip';
    var zipBlob = Utilities.zip(blobs, zipName);
    var zipFile = rootFolder.createFile(zipBlob);
    zipFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    // Cache ZIP ID
    var session = getSession(sessionToken);
    CacheService.getScriptCache().put('tmpzip:' + zipFile.getId(), JSON.stringify({
      createdAt: Date.now(),
      createdBy: session ? session.userId : '',
      type: 'student'
    }), 60 * 60);
    
    return {
      success: true,
      message: 'Đã nén ' + blobs.length + ' file của HS ' + maHS,
      downloadUrl: zipFile.getDownloadUrl(),
      zipFileId: zipFile.getId(),
      fileName: zipName,
      fileCount: blobs.length
    };
  } catch (err) {
    return { success: false, message: 'Lỗi tải file: ' + err.toString() };
  }
}

/**
 * G4: Xóa file ZIP tạm — kiểm tra session + người tạo ZIP
 */
function cleanupDownloadZip(fileId, sessionToken) {
  if (!fileId) return { success: false, message: 'Thiếu fileId!' };
  // Kiểm tra session GV/Admin
  var session = requireTeacherOrAdminSession(sessionToken);
  if (!session) return { success: false, message: '⛔ Không có quyền!' };
  // Verify ZIP thuộc phiên tải hợp lệ
  var metaRaw = CacheService.getScriptCache().get('tmpzip:' + fileId);
  if (!metaRaw) {
    return { success: false, message: 'File ZIP không thuộc phiên tải hợp lệ hoặc đã hết hạn!' };
  }
  // G4: Kiểm tra chính xác người tạo ZIP (trừ admin có thể xóa bất kỳ)
  try {
    var meta = JSON.parse(metaRaw);
    if (session.role !== 'admin' && meta.createdBy !== session.userId) {
      return { success: false, message: 'Bạn không có quyền xóa ZIP này!' };
    }
  } catch (parseE) {
    return { success: false, message: 'Lỗi đọc metadata ZIP!' };
  }
  try {
    var file = DriveApp.getFileById(fileId);
    if (file.getMimeType() !== MimeType.ZIP && file.getMimeType() !== 'application/zip') {
      return { success: false, message: 'File không phải ZIP tạm!' };
    }
    file.setTrashed(true);
    CacheService.getScriptCache().remove('tmpzip:' + fileId);
    return { success: true, message: 'Đã xóa file ZIP tạm.' };
  } catch (err) {
    return { success: false, message: 'Lỗi xóa file: ' + err.toString() };
  }
}

// ============== HÀM HỖ TRỢ ==============

// Validate mã người dùng (maHS, maGV) — chống Firebase path injection
// Chỉ cho phép chữ, số, gạch dưới, gạch ngang
function validateUserCode(code) {
  if (!code || typeof code !== 'string') return false;
  var trimmed = code.trim();
  if (trimmed.length === 0 || trimmed.length > 50) return false;
  return /^[A-Za-z0-9_\-]+$/.test(trimmed);
}

// Lấy sheet an toàn — trả null nếu không tìm thấy (thay vì crash)
function getSheetSafe(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log('⚠️ Sheet "' + sheetName + '" không tồn tại!');
  }
  return sheet;
}

function shuffleArray(arr) {
  for (var i = arr.length - 1; i > 0; i--) {
    var j = Math.floor(Math.random() * (i + 1));
    var temp = arr[i];
    arr[i] = arr[j];
    arr[j] = temp;
  }
}

function getOrCreateFolder(parent, name) {
  var folders = parent.getFoldersByName(name);
  if (folders.hasNext()) {
    return folders.next();
  }
  return parent.createFolder(name);
}

// P0-10: Server-side saved answers — thay vì client ghi Firebase trực tiếp
function saveStudentAnswers(data) {
  if (!data || !data.sessionToken || !data.maHS) {
    return { success: false, message: 'Thiếu thông tin!' };
  }
  var session = requireSession(data.sessionToken, ['student']);
  if (!session || session.userId !== data.maHS.toString().trim()) {
    return { success: false, message: 'Phiên không hợp lệ!' };
  }
  if (!validateUserCode(data.maHS)) {
    return { success: false, message: 'Mã HS không hợp lệ!' };
  }
  var maHS = data.maHS.toString().trim();
  try {
    firebaseUpdate('students/' + maHS, { savedAnswers: data.answers || {} });
    return { success: true };
  } catch (e) {
    return { success: false, message: 'Lỗi lưu đáp án: ' + e.toString() };
  }
}

// D1b: Release student session khi logout — thay thế việc client ghi Firebase trực tiếp
function releaseStudentSession(data) {
  if (!data || !data.sessionToken) {
    return { success: false, message: 'Thiếu sessionToken!' };
  }
  var session = requireSession(data.sessionToken, ['student']);
  if (!session) {
    return { success: true }; // đã hết hạn hoặc không tồn tại — OK để logout
  }
  var maHS = session.userId;
  if (!validateUserCode(maHS)) return { success: false, message: 'Mã HS không hợp lệ!' };
  try {
    firebaseUpdate('students/' + maHS, { isLocked: false, lockedAt: null });
    // Xóa session khỏi cache
    CacheService.getScriptCache().remove('session:' + data.sessionToken);
    return { success: true };
  } catch (e) {
    return { success: false, message: 'Lỗi release session: ' + e.toString() };
  }
}

function removeVietnameseTones(str) {
  str = str.normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  str = str.replace(/đ/g, 'd').replace(/Đ/g, 'D');
  return str;
}

function formatTime(dateValue) {
  try {
    return Utilities.formatDate(new Date(dateValue), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
  } catch (e) {
    return dateValue.toString();
  }
}

// ============== ĐỒNG BỘ GGSHEET → FIREBASE (MERGE — BẢO TOÀN RUNTIME STATE) ==============
function syncSheetToFirebase() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // Danh sách field runtime cần bảo toàn khi merge (KHÔNG ghi đè)
  var RUNTIME_FIELDS = ['moThiThat', 'moThiThu', 'moThi', 'isLocked', 'lockedAt',
    'assignedQuestions', 'assignedAnswerOrder', 'savedAnswers', 'savedAnswersTime',
    'driveFolder', 'uploadFailed', 'uploadError', 'tags', 'ghiChu'];

  // 1. Sync Hoc_Sinh — MERGE (giữ nguyên runtime state)
  var hSheet = ss.getSheetByName('Hoc_Sinh');
  var hData = hSheet.getDataRange().getValues();
  var existingStudents = firebaseGet('students') || {};
  var mergedStudents = {};
  var sheetMaHSSet = {};
  for (var i = 1; i < hData.length; i++) {
    var maHS = hData[i][1].toString().trim();
    if (!maHS) continue;
    sheetMaHSSet[maHS] = true;
    var existing = existingStudents[maHS] || {};
    // Bắt đầu từ dữ liệu Sheet
    mergedStudents[maHS] = {
      stt: hData[i][0],
      hoTen: hData[i][2].toString(),
      lop: hData[i][3].toString(),
      matKhau: hData[i][4].toString(),
      trangThai: hData[i][5] ? hData[i][5].toString().trim() : '',
      thoiGianDN: hData[i][6] ? new Date(hData[i][6]).getTime() : null,
      thoiGianNB: hData[i][7] ? new Date(hData[i][7]).getTime() : null
    };
    // Giữ nguyên runtime fields từ Firebase
    for (var r = 0; r < RUNTIME_FIELDS.length; r++) {
      var rf = RUNTIME_FIELDS[r];
      if (existing[rf] !== undefined && existing[rf] !== null) {
        mergedStudents[maHS][rf] = existing[rf];
      }
    }
  }
  // Giữ lại HS chỉ có trên Firebase (đã import qua web, chưa có trên Sheet)
  for (var fbMaHS in existingStudents) {
    if (!sheetMaHSSet[fbMaHS]) {
      mergedStudents[fbMaHS] = existingStudents[fbMaHS];
    }
  }
  firebaseSet('students', mergedStudents);

  // 2. Sync Cau_Hoi — dùng maCH làm key (thay vì STT số)
  var qSheet = ss.getSheetByName('Cau_Hoi');
  var qData = qSheet.getDataRange().getValues();
  var questions = {};
  var duplicateMaCH = [];
  for (var j = 1; j < qData.length; j++) {
    var maCH = qData[j][0].toString().trim();
    if (!maCH) continue;
    if (questions[maCH]) {
      duplicateMaCH.push(maCH);
    }
    var qKhoi = qData[j][7] ? qData[j][7].toString().trim() : '';
    // Auto-extract khối từ maCH nếu cột Khoi trống (VD: K6_001 → khối 6)
    if (!qKhoi && maCH.match(/^K(\d+)_/)) {
      qKhoi = maCH.match(/^K(\d+)_/)[1];
    }
    questions[maCH] = {
      maCH: maCH,
      stt: maCH,
      noiDung: qData[j][1].toString(),
      dapAn1: qData[j][2].toString(),
      dapAn2: qData[j][3].toString(),
      dapAn3: qData[j][4].toString(),
      dapAn4: qData[j][5].toString(),
      dapAnDung: qData[j][6].toString().trim().toLowerCase(),
      lop: qKhoi
    };
  }
  if (duplicateMaCH.length > 0) {
    // Cảnh báo nhưng vẫn sync (câu sau ghi đè câu trước)
    Logger.log('WARNING: Mã câu hỏi trùng: ' + duplicateMaCH.join(', '));
  }
  firebaseSet('questions', questions);

  // 3. Sync Giao_Vien — dùng PATCH (giữ field bổ sung nếu có)
  var gvSheet = ss.getSheetByName('Giao_Vien');
  var teacherCount = 0;
  if (gvSheet) {
    var gvData = gvSheet.getDataRange().getValues();
    var teachers = {};
    for (var g = 1; g < gvData.length; g++) {
      var maGV = gvData[g][0].toString().trim();
      if (!maGV) continue;
      teachers[maGV] = {
        hoTen: gvData[g][1].toString(),
        matKhau: gvData[g][2].toString(),
        lopPhuTrach: gvData[g][3].toString()
      };
    }
    teacherCount = Object.keys(teachers).length;
    firebaseSet('teachers', teachers);
  }

  // 4. Sync De_Thi
  var dtSheet = ss.getSheetByName('De_Thi');
  if (dtSheet) {
    var dtData = dtSheet.getDataRange().getValues();
    for (var d = 1; d < dtData.length; d++) {
      var examId = dtData[d][0].toString().trim();
      if (!examId) continue;
      firebaseUpdate('exams/' + examId, {
        tenDe: dtData[d][1].toString(),
        maGV: dtData[d][2].toString().trim(),
        lops: dtData[d][3].toString(),
        soCauThi: parseInt(dtData[d][4]) || 8,
        soCauThu: parseInt(dtData[d][5]) || 16,
        trangThai: dtData[d][6].toString().trim().toLowerCase() === 'mở' ? 'mo' : 'dong',
        thiThu: dtData[d][7] ? dtData[d][7].toString().trim().toLowerCase() === 'có' : true,
        maxAttempts: parseInt(dtData[d][8]) || 5
      });
    }
  }

  // 5. Sync Cai_Dat
  var cSheet = ss.getSheetByName('Cai_Dat');
  if (cSheet) {
    var cData = cSheet.getDataRange().getValues();
    var settings = {};
    for (var k = 1; k < cData.length; k++) {
      var key = cData[k][0].toString().trim();
      var val = cData[k][1].toString().trim();
      if (key === 'practiceMode') settings.practiceMode = (val === 'true');
    }
    firebaseSet('settings', settings);
  }

  // 6. Set admin password
  firebaseUpdate('admin', { password: ADMIN_PASSWORD });

  // 7. Update meta
  firebaseUpdate('meta', {
    lastSyncFromSheet: Date.now(),
    totalStudents: Object.keys(mergedStudents).length,
    totalQuestions: Object.keys(questions).length,
    totalTeachers: teacherCount
  });

  return {
    success: true,
    message: '✅ Đồng bộ thành công (MERGE)! ' + Object.keys(mergedStudents).length + ' HS, ' + Object.keys(questions).length + ' câu hỏi, ' + teacherCount + ' GV. Runtime state đã được bảo toàn.'
  };
}

// ============== ĐỒNG BỘ FIREBASE → GGSHEET ==============
function syncFirebaseToSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // 1. Sync results from Firebase to Ket_Qua sheet
  var fbResults = firebaseGet('results') || {};
  var rSheet = ss.getSheetByName('Ket_Qua');

  // Clear existing data (keep header)
  if (rSheet.getLastRow() > 1) {
    rSheet.deleteRows(2, rSheet.getLastRow() - 1);
  }

  var resultKeys = Object.keys(fbResults);
  var count = 0;
  for (var i = 0; i < resultKeys.length; i++) {
    var r = fbResults[resultKeys[i]];
    count++;
    var rowData = [count, r.maHS || '', r.hoTen || '', r.lop || ''];
    var answers = r.answers || [];
    // answers có thể là array of objects {dapAnHS, ketQua, ...} hoặc array of strings
    var answerStrings = [];
    for (var a = 0; a < 8; a++) {
      if (a < answers.length && answers[a]) {
        if (typeof answers[a] === 'object') {
          // Format: object {dapAnHS, ketQua, ...}
          var ans = answers[a];
          var label = (ans.dapAnHS || ans.selected || '').toUpperCase();
          var mark = (ans.ketQua === 'Đúng' || ans.isCorrect) ? '✓' : '✗';
          answerStrings.push(label + mark);
        } else {
          // Format: string (VD: "A✓", "B✗")
          answerStrings.push(answers[a].toString());
        }
      } else {
        answerStrings.push('');
      }
      rowData.push(answerStrings[a]);
    }
    rowData.push(r.score || 0);
    rowData.push(r.thoiGian ? new Date(r.thoiGian) : '');
    rowData.push(r.dotThi || '');
    rSheet.appendRow(rowData);

    // Color cells
    var newRow = rSheet.getLastRow();
    for (var c = 0; c < 8; c++) {
      if (answerStrings[c]) {
        var cell = rSheet.getRange(newRow, 5 + c);
        if (answerStrings[c].indexOf('✓') >= 0) cell.setBackground('#d4edda');
        else cell.setBackground('#f8d7da');
      }
    }
  }

  // 2. Sync practice results
  var fbPractice = firebaseGet('practiceResults') || {};
  var pSheet = ss.getSheetByName('Ket_Qua_Thu');
  if (pSheet) {
    if (pSheet.getLastRow() > 1) pSheet.deleteRows(2, pSheet.getLastRow() - 1);
    var pKeys = Object.keys(fbPractice);
    for (var p = 0; p < pKeys.length; p++) {
      var pr = fbPractice[pKeys[p]];
      pSheet.appendRow([p + 1, pr.maHS || '', pr.hoTen || '', pr.lop || '', pr.lanThu || 1, pr.score || 0, pr.thoiGian ? new Date(pr.thoiGian) : '', pr.dotThi || '']);
    }
  }

  // 3. Sync student status
  var fbStudents = firebaseGet('students') || {};
  var hSheet = ss.getSheetByName('Hoc_Sinh');
  var hData = hSheet.getDataRange().getValues();
  for (var s = 1; s < hData.length; s++) {
    var maHS = hData[s][1].toString().trim();
    var fbStudent = fbStudents[maHS];
    if (fbStudent) {
      hSheet.getRange(s + 1, 6).setValue(fbStudent.trangThai || '');
      if (fbStudent.thoiGianDN) hSheet.getRange(s + 1, 7).setValue(new Date(fbStudent.thoiGianDN));
      if (fbStudent.thoiGianNB) hSheet.getRange(s + 1, 8).setValue(new Date(fbStudent.thoiGianNB));
    }
  }

  firebaseUpdate('meta', { lastSyncToSheet: Date.now() });

  return {
    success: true,
    message: '✅ Đã đồng bộ ' + count + ' kết quả thi + ' + Object.keys(fbPractice).length + ' kết quả thi thử về Google Sheet thành công!'
  };
}

function getAllQuestions() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Cau_Hoi');
  var data = sheet.getDataRange().getValues();

  var questions = [];
  for (var i = 1; i < data.length; i++) {
    questions.push({
      stt: data[i][0],
      noiDung: data[i][1].toString(),
      dapAn1: data[i][2].toString(),
      dapAn2: data[i][3].toString(),
      dapAn3: data[i][4].toString(),
      dapAn4: data[i][5].toString(),
      dapAnDung: data[i][6].toString().trim().toLowerCase(),
      lop: data[i][7] ? data[i][7].toString().trim() : ''
    });
  }

  return { success: true, questions: questions };
}

function updateQuestion(params) {
  if (!params.stt) return { success: false, message: 'Thiếu STT câu hỏi!' };
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Cau_Hoi');
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString() === params.stt.toString()) {
      if (params.noiDung !== undefined) sheet.getRange(i + 1, 2).setValue(params.noiDung);
      if (params.dapAn1 !== undefined) sheet.getRange(i + 1, 3).setValue(params.dapAn1);
      if (params.dapAn2 !== undefined) sheet.getRange(i + 1, 4).setValue(params.dapAn2);
      if (params.dapAn3 !== undefined) sheet.getRange(i + 1, 5).setValue(params.dapAn3);
      if (params.dapAn4 !== undefined) sheet.getRange(i + 1, 6).setValue(params.dapAn4);
      if (params.dapAnDung !== undefined) sheet.getRange(i + 1, 7).setValue(params.dapAnDung.toLowerCase());
      if (params.lop !== undefined) sheet.getRange(i + 1, 8).setValue(params.lop);
      // AUTO-SYNC: cập nhật Firebase
      try {
        var fbQ = {};
        if (params.noiDung !== undefined) fbQ.noiDung = params.noiDung;
        if (params.dapAn1 !== undefined) fbQ.dapAn1 = params.dapAn1;
        if (params.dapAn2 !== undefined) fbQ.dapAn2 = params.dapAn2;
        if (params.dapAn3 !== undefined) fbQ.dapAn3 = params.dapAn3;
        if (params.dapAn4 !== undefined) fbQ.dapAn4 = params.dapAn4;
        if (params.dapAnDung !== undefined) fbQ.dapAnDung = params.dapAnDung.toLowerCase();
        if (params.lop !== undefined) fbQ.lop = params.lop;
        if (Object.keys(fbQ).length > 0) {
          firebaseUpdate('questions/' + params.stt, fbQ);
        }
      } catch(fe) { Logger.log('Auto-sync FB updateQuestion error: ' + fe); }
      return { success: true, message: 'Cập nhật câu ' + params.stt + ' thành công!' };
    }
  }

  return { success: false, message: 'Không tìm thấy câu hỏi STT ' + params.stt };
}

// ============== AUTO-SYNC CRUD: HỌC SINH ==============

/**
 * Thêm 1 HS mới vào Sheet + Firebase (auto-sync)
 * @param {Object} data - { maHS, hoTen, lop, matKhau, maGV, isAdmin }
 */
function addStudentWithSync(data) {
  if (!data.maHS || !data.hoTen || !data.lop) {
    return { success: false, message: 'Thiếu thông tin: Mã HS, Họ tên, Lớp!' };
  }
  var maHS = data.maHS.toString().trim();
  var hoTen = data.hoTen.toString().trim();
  var lop = data.lop.toString().trim();
  var matKhau = data.matKhau ? data.matKhau.toString().trim() : '123456';

  // 1. Kiểm tra quyền GV (session-based)
  if (!_checkTeacherPermission(lop, data.sessionToken)) {
    return { success: false, message: '⛔ Bạn không có quyền thêm HS vào lớp ' + lop + '!' };
  }

  // 2. Kiểm tra trùng mã HS trên Sheet
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Hoc_Sinh');
  var sheetData = sheet.getDataRange().getValues();
  for (var i = 1; i < sheetData.length; i++) {
    if (sheetData[i][1].toString().trim() === maHS) {
      return { success: false, message: '⚠️ Mã HS "' + maHS + '" đã tồn tại trên Google Sheet (HS: ' + sheetData[i][2] + ', Lớp: ' + sheetData[i][3] + ')!' };
    }
  }

  // 3. Kiểm tra trùng mã HS trên Firebase
  var existingFb = firebaseGet('students/' + maHS);
  if (existingFb) {
    return { success: false, message: '⚠️ Mã HS "' + maHS + '" đã tồn tại trên Firebase (HS: ' + (existingFb.hoTen || '') + ', Lớp: ' + (existingFb.lop || '') + ')!' };
  }

  // 4. Ghi vào Sheet
  var newSTT = sheetData.length; // STT mới = số dòng hiện tại (đã trừ header)
  sheet.appendRow([newSTT, maHS, hoTen, lop, matKhau, '', '', '']);

  // 5. Ghi vào Firebase (PATCH — không ảnh hưởng HS khác)
  try {
    firebaseUpdate('students/' + maHS, {
      stt: newSTT,
      hoTen: hoTen,
      lop: lop,
      matKhau: matKhau,
      trangThai: '',
      moThiThat: false,
      moThiThu: false
    });
  } catch(fe) { Logger.log('addStudentWithSync Firebase error: ' + fe); }

  return { success: true, message: '✅ Đã thêm HS ' + hoTen + ' (' + maHS + ') vào lớp ' + lop + ' + đồng bộ Firebase!' };
}

/**
 * Xóa 1 HS khỏi Sheet + Firebase (auto-sync)
 * @param {string} maHS - Mã HS cần xóa
 * @param {string} maGV - Mã GV (kiểm tra quyền)
 * @param {string|boolean} isAdmin - Có phải admin không
 */
function deleteStudentWithSync(maHS, sessionToken) {
  if (!maHS) return { success: false, message: 'Thiếu mã HS!' };
  maHS = maHS.toString().trim();

  // 1. Tìm HS trên Sheet để lấy lớp (kiểm tra quyền)
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Hoc_Sinh');
  var data = sheet.getDataRange().getValues();
  var foundRow = -1;
  var studentLop = '';
  for (var i = 1; i < data.length; i++) {
    if (data[i][1].toString().trim() === maHS) {
      foundRow = i + 1;
      studentLop = data[i][3].toString().trim();
      break;
    }
  }

  // Nếu không có trên Sheet, thử lấy từ Firebase
  if (foundRow === -1) {
    var fbStudent = firebaseGet('students/' + maHS);
    if (fbStudent) {
      studentLop = fbStudent.lop || '';
    }
  }

  // 2. Kiểm tra quyền GV (session-based)
  if (studentLop && !_checkTeacherPermission(studentLop, sessionToken)) {
    return { success: false, message: '⛔ Bạn không có quyền xóa HS lớp ' + studentLop + '!' };
  }

  // 3. Xóa khỏi Sheet
  if (foundRow > 0) {
    sheet.deleteRow(foundRow);
  }

  // 4. Xóa khỏi Firebase
  try {
    firebaseDelete('students/' + maHS);
  } catch(fe) { Logger.log('deleteStudentWithSync Firebase error: ' + fe); }

  return { success: true, message: '✅ Đã xóa HS ' + maHS + ' khỏi Sheet + Firebase!' };
}

/**
 * Import nhiều HS vào Sheet + Firebase (auto-sync)
 * @param {Object} data - { students: [{maHS, hoTen, lop, matKhau}], mode: 'add'|'overwrite', maGV, isAdmin }
 */
function importStudentsWithSync(data) {
  if (!data.students || !data.students.length) {
    return { success: false, message: 'Không có dữ liệu HS để import!' };
  }
  var mode = data.mode || 'add';
  var importList = data.students;

  // 1. Kiểm tra quyền GV cho tất cả lớp
  var classesInImport = {};
  for (var c = 0; c < importList.length; c++) {
    classesInImport[importList[c].lop.toString().trim()] = true;
  }
  var blockedClasses = [];
  for (var cls in classesInImport) {
    if (!_checkTeacherPermission(cls, data.sessionToken)) {
      blockedClasses.push(cls);
    }
  }
  if (blockedClasses.length > 0) {
    return { success: false, message: '⛔ Bạn không có quyền import HS vào lớp: ' + blockedClasses.join(', ') + '!' };
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Hoc_Sinh');
  var sheetData = sheet.getDataRange().getValues();

  // 2. Kiểm tra trùng mã HS (chỉ ở mode 'add')
  if (mode === 'add') {
    var duplicates = [];
    for (var d = 0; d < importList.length; d++) {
      var checkMa = importList[d].maHS.toString().trim();
      for (var s = 1; s < sheetData.length; s++) {
        if (sheetData[s][1].toString().trim() === checkMa) {
          duplicates.push(checkMa + ' (' + sheetData[s][2] + ')');
          break;
        }
      }
    }
    if (duplicates.length > 0) {
      return { success: false, message: '⚠️ Các mã HS đã tồn tại trên Sheet: ' + duplicates.slice(0, 10).join(', ') + (duplicates.length > 10 ? '... và ' + (duplicates.length - 10) + ' mã khác' : ''), duplicates: duplicates };
    }
  }

  // 3. Nếu mode overwrite: xóa HS cũ cùng lớp khỏi Sheet
  if (mode === 'overwrite') {
    var classesToOverwrite = Object.keys(classesInImport);
    // Xóa từ dưới lên để không bị shift index
    for (var del = sheetData.length - 1; del >= 1; del--) {
      var delLop = sheetData[del][3].toString().trim();
      if (classesToOverwrite.indexOf(delLop) >= 0) {
        sheet.deleteRow(del + 1);
      }
    }
  }

  // 4. Ghi vào Sheet + Firebase
  var fbUpdates = {};
  var currentLastRow = sheet.getLastRow();
  for (var imp = 0; imp < importList.length; imp++) {
    var student = importList[imp];
    var maHS = student.maHS.toString().trim();
    var hoTen = student.hoTen.toString().trim();
    var lop = student.lop.toString().trim();
    var matKhau = student.matKhau ? student.matKhau.toString().trim() : '123456';
    if (!matKhau) matKhau = '123456';

    // Ghi Sheet
    currentLastRow++;
    sheet.appendRow([currentLastRow - 1, maHS, hoTen, lop, matKhau, '', '', '']);

    // Chuẩn bị Firebase batch update
    fbUpdates['students/' + maHS + '/stt'] = currentLastRow - 1;
    fbUpdates['students/' + maHS + '/hoTen'] = hoTen;
    fbUpdates['students/' + maHS + '/lop'] = lop;
    fbUpdates['students/' + maHS + '/matKhau'] = matKhau;
    fbUpdates['students/' + maHS + '/trangThai'] = '';
    // Chỉ set moThiThat/moThiThu nếu chưa có (bảo toàn nếu đã tồn tại)
    var existingFb = firebaseGet('students/' + maHS);
    if (!existingFb) {
      fbUpdates['students/' + maHS + '/moThiThat'] = false;
      fbUpdates['students/' + maHS + '/moThiThu'] = false;
    }
  }

  // 5. Batch ghi Firebase (1 request duy nhất — nhanh + atomic)
  try {
    firebaseUpdate('', fbUpdates);
  } catch(fe) { Logger.log('importStudentsWithSync Firebase error: ' + fe); }

  return {
    success: true,
    message: '✅ Đã import ' + importList.length + ' HS vào Sheet + Firebase!' + (mode === 'overwrite' ? ' (Chế độ ghi đè)' : ''),
    count: importList.length
  };
}

/**
 * Xóa toàn bộ HS 1 lớp khỏi Sheet + Firebase (auto-sync)
 * @param {string} lop - Tên lớp cần xóa
 * @param {string} maGV - Mã GV (kiểm tra quyền)
 * @param {string|boolean} isAdmin - Có phải admin không
 */
function deleteClassWithSync(lop, sessionToken) {
  if (!lop) return { success: false, message: 'Thiếu tên lớp!' };
  lop = lop.toString().trim();

  // 1. Kiểm tra quyền GV (session-based)
  if (!_checkTeacherPermission(lop, sessionToken)) {
    return { success: false, message: '⛔ Bạn không có quyền xóa lớp ' + lop + '!' };
  }

  // 2. Tìm HS của lớp trên Sheet
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Hoc_Sinh');
  var data = sheet.getDataRange().getValues();
  var maHSList = [];
  var rowsToDelete = [];
  for (var i = 1; i < data.length; i++) {
    if (data[i][3].toString().trim() === lop) {
      maHSList.push(data[i][1].toString().trim());
      rowsToDelete.push(i + 1);
    }
  }

  // 3. Xóa khỏi Sheet (từ dưới lên)
  for (var r = rowsToDelete.length - 1; r >= 0; r--) {
    sheet.deleteRow(rowsToDelete[r]);
  }

  // 4. Xóa khỏi Firebase
  try {
    var fbDeletes = {};
    for (var f = 0; f < maHSList.length; f++) {
      fbDeletes['students/' + maHSList[f]] = null;
    }
    if (Object.keys(fbDeletes).length > 0) {
      firebaseUpdate('', fbDeletes);
    }
  } catch(fe) { Logger.log('deleteClassWithSync Firebase error: ' + fe); }

  return { success: true, message: '✅ Đã xóa ' + maHSList.length + ' HS lớp ' + lop + ' khỏi Sheet + Firebase!', count: maHSList.length };
}

// ============== QUẢN LÝ CÂU HỎI: THÊM + ĐỒNG BỘ ==============

/**
 * Thêm 1 câu hỏi mới — auto-generate mã (K6_001, K9_002...) + sync GGSheet + Firebase
 * @param {Object} data - { khoi, noiDung, dapAn1, dapAn2, dapAn3, dapAn4, dapAnDung, [maCH] }
 *   - khoi: '6','7','8','9' hoặc '' (chung)
 *   - Nếu maCH được truyền thì dùng luôn, nếu không thì auto-generate
 */
function saveQuestionWithSync(data) {
  if (!data.noiDung) return { success: false, message: 'Thiếu nội dung câu hỏi!' };
  if (!data.dapAnDung) return { success: false, message: 'Thiếu đáp án đúng!' };

  var khoi = data.khoi ? data.khoi.toString().trim() : '';
  var maCH = '';

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Cau_Hoi');
  var sheetData = sheet.getDataRange().getValues();

  if (data.maCH && data.maCH.toString().trim()) {
    // Dùng mã câu hỏi do người dùng truyền vào
    maCH = data.maCH.toString().trim();
  } else {
    // Auto-generate mã: K{khoi}_{số thứ tự 3 chữ số}
    var prefix = khoi ? ('K' + khoi + '_') : 'KC_';
    var maxNum = 0;
    for (var i = 1; i < sheetData.length; i++) {
      var existingMa = sheetData[i][0].toString().trim();
      if (existingMa.indexOf(prefix) === 0) {
        var numPart = parseInt(existingMa.substring(prefix.length));
        if (!isNaN(numPart) && numPart > maxNum) maxNum = numPart;
      }
    }
    var nextNum = maxNum + 1;
    maCH = prefix + ('00' + nextNum).slice(-3); // K6_001, K6_002, ...
  }

  // Kiểm tra trùng mã trên GGSheet
  for (var j = 1; j < sheetData.length; j++) {
    if (sheetData[j][0].toString().trim() === maCH) {
      return { success: false, message: '⚠️ Mã câu hỏi "' + maCH + '" đã tồn tại trên Google Sheet!' };
    }
  }

  // Kiểm tra trùng mã trên Firebase
  var existingFb = firebaseGet('questions/' + maCH);
  if (existingFb) {
    return { success: false, message: '⚠️ Mã câu hỏi "' + maCH + '" đã tồn tại trên Firebase!' };
  }

  // Ghi vào GGSheet (cột: STT/MaCH, NoiDung, DA1, DA2, DA3, DA4, DADung, Lop)
  sheet.appendRow([
    maCH,
    data.noiDung.toString(),
    data.dapAn1 ? data.dapAn1.toString() : '',
    data.dapAn2 ? data.dapAn2.toString() : '',
    data.dapAn3 ? data.dapAn3.toString() : '',
    data.dapAn4 ? data.dapAn4.toString() : '',
    data.dapAnDung.toString().trim().toLowerCase(),
    khoi
  ]);

  // Ghi vào Firebase
  try {
    firebaseUpdate('questions/' + maCH, {
      maCH: maCH,
      stt: maCH,
      noiDung: data.noiDung.toString(),
      dapAn1: data.dapAn1 ? data.dapAn1.toString() : '',
      dapAn2: data.dapAn2 ? data.dapAn2.toString() : '',
      dapAn3: data.dapAn3 ? data.dapAn3.toString() : '',
      dapAn4: data.dapAn4 ? data.dapAn4.toString() : '',
      dapAnDung: data.dapAnDung.toString().trim().toLowerCase(),
      lop: khoi
    });
  } catch (fe) { Logger.log('saveQuestionWithSync Firebase error: ' + fe); }

  return {
    success: true,
    message: '✅ Đã thêm câu hỏi ' + maCH + ' thành công!',
    maCH: maCH
  };
}

/**
 * Thêm NHIỀU câu hỏi cùng lúc — auto-generate mã + sync GGSheet + Firebase
 * @param {Object} data - { questions: [{khoi, noiDung, dapAn1, dapAn2, dapAn3, dapAn4, dapAnDung, maCH?}] }
 */
function bulkSaveQuestions(data) {
  if (!data.questions || !Array.isArray(data.questions) || data.questions.length === 0) {
    return { success: false, message: 'Không có câu hỏi để lưu!' };
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Cau_Hoi');
  var sheetData = sheet.getDataRange().getValues();

  // Build existing code map to auto-generate
  var existingCodes = {};
  for (var e = 1; e < sheetData.length; e++) {
    existingCodes[sheetData[e][0].toString().trim()] = true;
  }

  var counters = {};
  // Scan existing codes for max numbers per prefix
  for (var key in existingCodes) {
    var match = key.match(/^(K\d+_|KC_)(\d+)$/);
    if (match) {
      var pf = match[1];
      var num = parseInt(match[2]);
      if (!counters[pf] || num > counters[pf]) counters[pf] = num;
    }
  }

  var savedQuestions = [];
  var fbUpdates = {};

  for (var i = 0; i < data.questions.length; i++) {
    var q = data.questions[i];
    if (!q.noiDung) continue;

    var khoi = q.khoi ? q.khoi.toString().trim() : (q.lop ? q.lop.toString().trim() : '');
    var maCH = q.maCH ? q.maCH.toString().trim() : '';

    // Auto-generate nếu maCH rỗng hoặc là số thuần
    if (!maCH || /^\d+$/.test(maCH)) {
      var prefix = khoi ? ('K' + khoi + '_') : 'KC_';
      if (!counters[prefix]) counters[prefix] = 0;
      counters[prefix]++;
      maCH = prefix + ('00' + counters[prefix]).slice(-3);
      // Tránh trùng
      while (existingCodes[maCH]) {
        counters[prefix]++;
        maCH = prefix + ('00' + counters[prefix]).slice(-3);
      }
    }

    // Skip trùng
    if (existingCodes[maCH]) continue;
    existingCodes[maCH] = true;

    var dapAnDung = (q.dapAnDung || 'a').toString().trim().toLowerCase();

    // Ghi GGSheet
    sheet.appendRow([
      maCH,
      (q.noiDung || '').toString(),
      (q.dapAn1 || '').toString(),
      (q.dapAn2 || '').toString(),
      (q.dapAn3 || '').toString(),
      (q.dapAn4 || '').toString(),
      dapAnDung,
      khoi
    ]);

    // Chuẩn bị Firebase update
    fbUpdates[maCH] = {
      maCH: maCH, stt: maCH,
      noiDung: (q.noiDung || '').toString(),
      dapAn1: (q.dapAn1 || '').toString(),
      dapAn2: (q.dapAn2 || '').toString(),
      dapAn3: (q.dapAn3 || '').toString(),
      dapAn4: (q.dapAn4 || '').toString(),
      dapAnDung: dapAnDung,
      lop: khoi
    };

    savedQuestions.push(maCH);
  }

  // Batch Firebase update
  if (Object.keys(fbUpdates).length > 0) {
    try { firebaseUpdate('questions', fbUpdates); } catch (fe) { Logger.log('bulkSaveQuestions Firebase error: ' + fe); }
  }

  return {
    success: true,
    message: '✅ Đã lưu ' + savedQuestions.length + ' câu hỏi vào GGSheet + Firebase!',
    count: savedQuestions.length,
    codes: savedQuestions
  };
}


// ============== QUẢN LÝ CÂU HỎI: XÓA + ĐỒNG BỘ ==============

/**
 * Xóa 1 câu hỏi theo mã — đồng bộ GGSheet + Firebase
 * @param {string} maCH - Mã câu hỏi cần xóa (VD: K6_003)
 */
function deleteQuestionWithSync(maCH) {
  if (!maCH) return { success: false, message: 'Thiếu mã câu hỏi!' };
  maCH = maCH.toString().trim();

  // Xóa khỏi GGSheet
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Cau_Hoi');
  var data = sheet.getDataRange().getValues();
  var found = false;
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim() === maCH) {
      sheet.deleteRow(i + 1);
      found = true;
      break;
    }
  }

  // Xóa khỏi Firebase
  try {
    firebaseDelete('questions/' + maCH);
  } catch (fe) { Logger.log('deleteQuestionWithSync Firebase error: ' + fe); }

  // Xóa khỏi examConfig/selectedQuestions nếu có
  try {
    var allConfig = firebaseGet('examConfig') || {};
    for (var grade in allConfig) {
      var cfg = allConfig[grade];
      if (cfg && cfg.selectedQuestions && cfg.selectedQuestions.length > 0) {
        var idx = cfg.selectedQuestions.indexOf(maCH);
        if (idx >= 0) {
          cfg.selectedQuestions.splice(idx, 1);
          firebaseUpdate('examConfig/' + grade, { selectedQuestions: cfg.selectedQuestions });
        }
      }
    }
  } catch (fe2) { Logger.log('Cleanup examConfig error: ' + fe2); }

  if (found) {
    return { success: true, message: '✅ Đã xóa câu hỏi ' + maCH + ' khỏi Sheet + Firebase!' };
  } else {
    return { success: true, message: '⚠️ Không tìm thấy ' + maCH + ' trên Sheet, đã xóa trên Firebase.' };
  }
}

/**
 * Xóa TẤT CẢ câu hỏi của 1 khối — đồng bộ GGSheet + Firebase
 * @param {string} khoi - Khối cần xóa (VD: '8', '9')
 */
function deleteQuestionsByGrade(khoi) {
  if (!khoi) return { success: false, message: 'Thiếu khối lớp!' };
  khoi = khoi.toString().trim();

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Cau_Hoi');
  var data = sheet.getDataRange().getValues();

  // Tìm các dòng cần xóa (từ dưới lên để không bị lệch index)
  var rowsToDelete = [];
  var deletedMaCHs = [];
  for (var i = 1; i < data.length; i++) {
    var maCH = data[i][0].toString().trim();
    var qKhoi = data[i][7] ? data[i][7].toString().trim() : '';
    // Kiểm tra cả cột Khoi và mã câu (K8_001 → khối 8)
    var matchByCol = (qKhoi === khoi);
    var matchByCode = false;
    if (!matchByCol && maCH) {
      var m = maCH.match(/^K(\d+)_/);
      if (m && m[1] === khoi) matchByCode = true;
    }
    if (matchByCol || matchByCode) {
      rowsToDelete.push(i + 1); // Sheet row (1-indexed + header)
      deletedMaCHs.push(maCH);
    }
  }

  if (deletedMaCHs.length === 0) {
    return { success: false, message: '⚠️ Không tìm thấy câu hỏi nào cho khối ' + khoi + '!' };
  }

  // Xóa GGSheet (từ dưới lên)
  for (var r = rowsToDelete.length - 1; r >= 0; r--) {
    sheet.deleteRow(rowsToDelete[r]);
  }

  // Xóa Firebase
  try {
    var fbDeletes = {};
    for (var f = 0; f < deletedMaCHs.length; f++) {
      fbDeletes[deletedMaCHs[f]] = null;
    }
    firebaseUpdate('questions', fbDeletes);
  } catch (fe) { Logger.log('deleteQuestionsByGrade Firebase error: ' + fe); }

  // Cleanup examConfig
  try {
    var cfg = firebaseGet('examConfig/' + khoi);
    if (cfg && cfg.selectedQuestions) {
      firebaseUpdate('examConfig/' + khoi, { selectedQuestions: [] });
    }
  } catch (fe2) { Logger.log('Cleanup examConfig error: ' + fe2); }

  return { success: true, message: '✅ Đã xóa ' + deletedMaCHs.length + ' câu hỏi khối ' + khoi + ' khỏi Sheet + Firebase!', count: deletedMaCHs.length };
}

// ============== CẤU HÌNH ĐỀ THI: ĐỌC ==============

/**
 * Đọc cấu hình đề thi theo khối (hoặc tất cả khối)
 * @param {string} khoi - Khối lớp ('6','7','8','9') hoặc '' = tất cả
 */
function getExamConfig(khoi) {
  try {
    if (khoi && khoi.toString().trim()) {
      var config = firebaseGet('examConfig/' + khoi.toString().trim());
      return {
        success: true,
        config: config || { soCauThiThu: 0, selectedQuestions: [] },
        khoi: khoi.toString().trim()
      };
    } else {
      // Trả tất cả cấu hình
      var allConfig = firebaseGet('examConfig') || {};
      return { success: true, configs: allConfig };
    }
  } catch (e) {
    return { success: false, message: 'Lỗi đọc cấu hình: ' + e.toString() };
  }
}

// ============== CẤU HÌNH ĐỀ THI: LƯU ==============

/**
 * Lưu cấu hình đề thi cho 1 khối
 * @param {Object} data - { khoi, soCauThiThu, selectedQuestions: ['K6_001','K6_003',...] }
 */
function saveExamConfig(data) {
  if (!data.khoi) return { success: false, message: 'Thiếu khối lớp!' };
  var khoi = data.khoi.toString().trim();
  var soCauThiThu = parseInt(data.soCauThiThu) || 0;
  var selectedQuestions = data.selectedQuestions || [];

  // Validate: kiểm tra các mã câu hỏi trong selectedQuestions có tồn tại không
  if (selectedQuestions.length > 0) {
    var existingQuestions = firebaseGet('questions') || {};
    var invalid = [];
    for (var i = 0; i < selectedQuestions.length; i++) {
      if (!existingQuestions[selectedQuestions[i]]) {
        invalid.push(selectedQuestions[i]);
      }
    }
    if (invalid.length > 0) {
      return { success: false, message: '⚠️ Các mã câu hỏi không tồn tại: ' + invalid.join(', ') };
    }
  }

  // Lưu vào Firebase
  try {
    firebaseSet('examConfig/' + khoi, {
      soCauThiThu: soCauThiThu,
      selectedQuestions: selectedQuestions,
      lastUpdated: Date.now()
    });
  } catch (e) {
    return { success: false, message: 'Lỗi lưu cấu hình: ' + e.toString() };
  }

  return {
    success: true,
    message: '✅ Đã lưu cấu hình đề thi khối ' + khoi + ': ' +
      soCauThiThu + ' câu thi thử, ' +
      selectedQuestions.length + ' câu thi thật!'
  };
}

// ============== MIGRATE MÃ CÂU HỎI: STT → K{khối}_{số} ==============

/**
 * Chuyển đổi mã câu hỏi từ STT số (1,2,3...) sang mã khối K6_001, K9_001...
 * Đọc cột Lop (cột H) để xác định khối.
 * Câu không có khối → mã KC_001 (chung)
 */
function migrateQuestionCodes() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Cau_Hoi');
  var data = sheet.getDataRange().getValues();

  // Bước 1: Đếm số câu hiện có theo khối
  var counters = {}; // { '6': 0, '9': 0, 'C': 0, ... }
  var migrations = []; // [{ row, oldMa, newMa }]

  for (var i = 1; i < data.length; i++) {
    var oldMa = data[i][0].toString().trim();
    // Bỏ qua câu đã có mã dạng K{x}_ rồi
    if (oldMa.match(/^K\w+_\d+$/)) continue;

    var khoi = data[i][7] ? data[i][7].toString().trim() : '';
    var prefix = khoi ? khoi : 'C';

    if (!counters[prefix]) counters[prefix] = 0;
    counters[prefix]++;
    var newMa = 'K' + prefix + '_' + ('00' + counters[prefix]).slice(-3);

    migrations.push({ row: i + 1, oldMa: oldMa, newMa: newMa, khoi: khoi });
  }

  if (migrations.length === 0) {
    return { success: true, message: '✅ Tất cả câu hỏi đã có mã chuẩn. Không cần migrate!' };
  }

  // Bước 2: Cập nhật GGSheet (cột STT = cột 1)
  for (var m = 0; m < migrations.length; m++) {
    sheet.getRange(migrations[m].row, 1).setValue(migrations[m].newMa);
  }

  // Bước 3: Xóa questions cũ trên Firebase và ghi lại với key mới
  try {
    // Đọc tất cả questions hiện có trên Firebase
    var oldQuestions = firebaseGet('questions') || {};
    var newQuestions = {};

    // Giữ lại các câu đã có mã chuẩn
    for (var key in oldQuestions) {
      if (key.match(/^K\w+_\d+$/)) {
        newQuestions[key] = oldQuestions[key];
      }
    }

    // Thêm các câu sau khi migrate (đọc lại từ Sheet để chính xác)
    var updatedData = sheet.getDataRange().getValues();
    for (var u = 1; u < updatedData.length; u++) {
      var maCH = updatedData[u][0].toString().trim();
      var qKhoi = updatedData[u][7] ? updatedData[u][7].toString().trim() : '';
      // Auto-extract khối từ maCH nếu cột Khoi trống
      if (!qKhoi && maCH.match(/^K(\d+)_/)) {
        qKhoi = maCH.match(/^K(\d+)_/)[1];
      }
      newQuestions[maCH] = {
        maCH: maCH,
        stt: maCH,
        noiDung: updatedData[u][1].toString(),
        dapAn1: updatedData[u][2].toString(),
        dapAn2: updatedData[u][3].toString(),
        dapAn3: updatedData[u][4].toString(),
        dapAn4: updatedData[u][5].toString(),
        dapAnDung: updatedData[u][6].toString().trim().toLowerCase(),
        lop: qKhoi
      };
    }

    firebaseSet('questions', newQuestions);
  } catch (fe) {
    Logger.log('migrateQuestionCodes Firebase error: ' + fe);
    return {
      success: true,
      message: '⚠️ Đã migrate ' + migrations.length + ' câu trên Sheet, nhưng Firebase bị lỗi: ' + fe.toString(),
      migrations: migrations
    };
  }

  return {
    success: true,
    message: '✅ Đã migrate ' + migrations.length + ' câu hỏi sang mã mới! ' +
      Object.keys(counters).map(function(k) { return 'K' + k + ': ' + counters[k] + ' câu'; }).join(', '),
    migrations: migrations,
    totalMigrated: migrations.length
  };
}
