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
  var kqHeaders = ['STT', 'Ma_HS', 'Ho_Ten', 'Lop', 'Cau_1', 'Cau_2', 'Cau_3', 'Cau_4', 'Cau_5', 'Cau_6', 'Cau_7', 'Cau_8', 'Diem', 'Thoi_Gian'];
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
  var params = e.parameter || {};
  var action = params.action || '';
  var result = {};

  try {
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
        result = getRandomQuestions(params.mode || '');
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
        // Fallback: also handle via GET param
        if (e.postData && e.postData.contents) {
          var postData = JSON.parse(e.postData.contents);
          result = submitExam(postData);
        } else if (params.examData) {
          var getData = JSON.parse(decodeURIComponent(params.examData));
          result = submitExam(getData);
        } else {
          result = { success: false, message: 'Không nhận được dữ liệu bài thi!' };
        }
        break;
      case 'saveFile':
        if (e.postData && e.postData.contents) {
          var filePost = JSON.parse(e.postData.contents);
          result = saveFileOnly(filePost);
        } else {
          result = { success: false, message: 'Không nhận được file!' };
        }
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
        result = clearAllData();
        break;
      case 'setup':
        result = { success: true, message: setupSheets() };
        break;
      case 'getAllQuestions':
        result = getAllQuestions();
        break;
      case 'updateQuestion':
        if (e.postData && e.postData.contents) {
          var qPost = JSON.parse(e.postData.contents);
          result = updateQuestion(qPost);
        } else {
          result = updateQuestion(params);
        }
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
        if (e.postData && e.postData.contents) {
          var tPost = JSON.parse(e.postData.contents);
          result = saveTeacher(tPost);
        } else {
          result = saveTeacher(params);
        }
        break;
      case 'deleteTeacher':
        result = deleteTeacher(params.maGV);
        break;
      case 'uploadStudyFile':
        if (e.postData && e.postData.contents) {
          var sfPost = JSON.parse(e.postData.contents);
          result = uploadStudyFile(sfPost);
        } else {
          result = { success: false, message: 'Missing POST data' };
        }
        break;
      case 'downloadClassFiles':
        result = downloadClassFiles(params.lop, params.maGV, params.isAdmin);
        break;
      case 'downloadStudentFiles':
        result = downloadStudentFiles(params.lop, params.maHS, params.maGV, params.isAdmin);
        break;
      case 'cleanupDownloadZip':
        result = cleanupDownloadZip(params.fileId);
        break;
      case 'addStudentWithSync':
        if (e.postData && e.postData.contents) {
          var asPost = JSON.parse(e.postData.contents);
          result = addStudentWithSync(asPost);
        } else {
          result = addStudentWithSync(params);
        }
        break;
      case 'deleteStudentWithSync':
        result = deleteStudentWithSync(params.maHS, params.maGV, params.isAdmin);
        break;
      case 'importStudentsWithSync':
        if (e.postData && e.postData.contents) {
          var isPost = JSON.parse(e.postData.contents);
          result = importStudentsWithSync(isPost);
        } else {
          result = { success: false, message: 'Cần POST data!' };
        }
        break;
      case 'deleteClassWithSync':
        result = deleteClassWithSync(params.lop, params.maGV, params.isAdmin);
        break;
      default:
        result = { success: false, message: 'Action không hợp lệ' };
    }
  } catch (err) {
    result = { success: false, message: 'Lỗi hệ thống: ' + err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============== ĐĂNG NHẬP HỌC SINH ==============
function studentLogin(maHS, password) {
  if (!maHS || !password) {
    return { success: false, message: 'Vui lòng nhập đầy đủ Mã HS và Mật khẩu!' };
  }

  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Hoc_Sinh');
  var data = sheet.getDataRange().getValues();

  // Kiểm tra chế độ thi thử
  var settingsResult = getSettings();
  var isPractice = settingsResult.settings.practiceMode === true || settingsResult.settings.practiceMode === 'true';

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
  // Đếm từ Firebase (nguồn chính, đồng bộ với reset trên client)
  try {
    var practiceData = firebaseGet('practiceResults') || {};
    var count = 0;
    for (var key in practiceData) {
      if (practiceData[key] && practiceData[key].maHS === maHS) count++;
    }
    return count;
  } catch (e) {
    // Fallback: đếm từ GGSheet nếu Firebase lỗi
    Logger.log('Firebase read error, fallback to GGSheet: ' + e);
    var sheet = getOrCreatePracticeSheet(ss);
    var data = sheet.getDataRange().getValues();
    var count = 0;
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][1].toString().trim() === maHS) count++;
    }
    return count;
  }
}

// Tạo sheet Ket_Qua_Thu nếu chưa có
function getOrCreatePracticeSheet(ss) {
  var sheet = ss.getSheetByName('Ket_Qua_Thu');
  if (!sheet) {
    sheet = ss.insertSheet('Ket_Qua_Thu');
    sheet.appendRow(['STT', 'Ma_HS', 'Ho_Ten', 'Lop', 'Lan_Thu', 'Diem', 'Thoi_Gian']);
    sheet.getRange(1, 1, 1, 7)
      .setFontWeight('bold')
      .setBackground('#9900ff')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ============== ĐĂNG NHẬP GIÁO VIÊN ==============
function teacherLogin(maGV, password) {
  if (!maGV || !password) return { success: false, message: '❌ Vui lòng nhập mã GV và mật khẩu!' };
  var teacher = firebaseGet('teachers/' + maGV.trim());
  if (!teacher) return { success: false, message: '❌ Mã giáo viên không tồn tại!' };
  if (teacher.matKhau !== password) return { success: false, message: '❌ Mật khẩu không đúng!' };
  return {
    success: true,
    message: 'Đăng nhập thành công!',
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
    return { success: true, message: 'Đăng nhập Admin thành công!', isAdmin: true };
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
function getRandomQuestions(mode) {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sheet = ss.getSheetByName('Cau_Hoi');
  var data = sheet.getDataRange().getValues();

  var allQuestions = [];
  for (var i = 1; i < data.length; i++) {
    allQuestions.push({
      stt: data[i][0],
      noiDung: data[i][1].toString(),
      dapAn: [
        data[i][2].toString(),
        data[i][3].toString(),
        data[i][4].toString(),
        data[i][5].toString()
      ]
      // BẢO MẬT: KHÔNG gửi dapAnDung cho client
    });
  }

  // Trộn ngẫu nhiên
  shuffleArray(allQuestions);

  if (mode === 'practice') {
    return { success: true, questions: allQuestions };
  } else {
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
  // Race condition protection — chỉ lock phần ghi GGSheet
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000); // tăng lên 30 giây để hỗ trợ nhiều HS nộp cùng lúc
  } catch (e) {
    return { success: false, message: '⏳ Hệ thống đang bận, vui lòng thử lại sau vài giây!' };
  }

  var result; // kết quả trả về
  var needDriveSave = false; // flag để lưu Drive sau khi release lock
  var driveData = null;

  try {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var isPractice = data.isPractice === true || data.isPractice === 'true';

  // 1. Lấy đáp án đúng từ sheet Cau_Hoi
  var qSheet = ss.getSheetByName('Cau_Hoi');
  var qData = qSheet.getDataRange().getValues();
  var correctAnswers = {};
  var questionContents = {};
  var questionOptions = {};
  for (var i = 1; i < qData.length; i++) {
    correctAnswers[qData[i][0]] = qData[i][6].toString().trim().toLowerCase();
    questionContents[qData[i][0]] = qData[i][1].toString();
    questionOptions[qData[i][0]] = {
      a: qData[i][2].toString(),
      b: qData[i][3].toString(),
      c: qData[i][4].toString(),
      d: qData[i][5].toString()
    };
  }

  // 2. Chấm điểm
  var totalQ = data.answers.length;
  var pointPerQ = totalQ > 0 ? (isPractice ? 10 / totalQ : 4 / totalQ) : 0;
  var score = 0;
  var details = [];
  var answerResults = [];

  for (var j = 0; j < data.answers.length; j++) {
    var ans = data.answers[j];
    var correct = correctAnswers[ans.questionSTT];
    var isCorrect = ans.selected.toLowerCase() === correct;
    if (isCorrect) score += pointPerQ;

    var opts = questionOptions[ans.questionSTT] || {};
    details.push({
      stt: ans.questionSTT,
      noiDung: ans.noiDung || questionContents[ans.questionSTT] || '',
      dapAn: opts,
      dapAnDung: correct,
      dapAnHS: ans.selected.toLowerCase(),
      ketQua: isCorrect ? 'Đúng' : 'Sai'
    });

    answerResults.push(ans.selected.toUpperCase() + (isCorrect ? '✓' : '✗'));
  }

  score = Math.round(score * 100) / 100;

  if (isPractice) {
    // ===== THI THỬ: ghi vào Ket_Qua_Thu =====
    var pSheet = getOrCreatePracticeSheet(ss);
    var pLastRow = pSheet.getLastRow();
    var attempts = countPracticeAttempts(ss, data.maHS.trim());
    // Server-side validation: prevent bypassing max attempts
    if (attempts >= 5) {
      result = { success: false, message: '⚠️ Bạn đã hết 5 lượt thi thử! Không thể nộp thêm.' };
    } else {
      pSheet.appendRow([pLastRow, data.maHS, data.hoTen, data.lop, attempts + 1, score, new Date()]);

      // Write to Firebase
      try {
        firebasePush('practiceResults', {
          maHS: data.maHS,
          hoTen: data.hoTen,
          lop: data.lop,
          lanThu: attempts + 1,
          score: score,
          thoiGian: Date.now(),
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
    // 3. Ghi vào sheet Ket_Qua
    var rSheet = ss.getSheetByName('Ket_Qua');
    var lastRow = rSheet.getLastRow();
    var newSTT = lastRow >= 1 ? lastRow : 1;
    var rowData = [newSTT, data.maHS, data.hoTen, data.lop];
    for (var k = 0; k < 8; k++) {
      rowData.push(k < answerResults.length ? answerResults[k] : '');
    }
    rowData.push(score);
    rowData.push(new Date());
    rSheet.appendRow(rowData);

    // Color correct/wrong answer cells (columns 5-12 = index E-L)
    var newRow = rSheet.getLastRow();
    for (var c = 0; c < 8; c++) {
      if (c < answerResults.length && answerResults[c]) {
        var cellRange = rSheet.getRange(newRow, 5 + c);
        if (answerResults[c].includes('✓')) {
          cellRange.setBackground('#d4edda'); // light green
        } else {
          cellRange.setBackground('#f8d7da'); // light red
        }
      }
    }

    // 4. Cập nhật trạng thái thi = X trên GGSheet
    var hSheet = ss.getSheetByName('Hoc_Sinh');
    var hData = hSheet.getDataRange().getValues();
    for (var m = 1; m < hData.length; m++) {
      if (hData[m][1].toString().trim() === data.maHS.trim()) {
        hSheet.getRange(m + 1, 6).setValue('X');
        hSheet.getRange(m + 1, 8).setValue(new Date());
        break;
      }
    }

    // 5. Write to Firebase
    try {
      var correctCount = 0, wrongCount = 0;
      answerResults.forEach(function(ar) {
        if (ar && ar.includes('✓')) correctCount++;
        else if (ar && ar.length > 0) wrongCount++;
      });
      firebasePush('results', {
        maHS: data.maHS,
        hoTen: data.hoTen,
        lop: data.lop,
        answers: details,
        correctCount: correctCount,
        wrongCount: wrongCount,
        score: score,
        thoiGian: Date.now(),
        synced: true
      });
      // Update student status on Firebase + clear assigned questions
      firebaseUpdate('students/' + data.maHS, {
        trangThai: 'X',
        thoiGianNB: Date.now(),
        assignedQuestions: null,
        assignedAnswerOrder: null
      });
    } catch (fe) { Logger.log('Firebase result write error: ' + fe); }

    // Đánh dấu cần lưu Drive — sẽ thực hiện SAU KHI release lock
    needDriveSave = true;
    driveData = { data: data, details: details, score: score };

    result = {
      success: true,
      message: '🎉 Nộp bài thành công!',
      score: score,
      isPractice: false
    };
  }

  } finally {
    lock.releaseLock(); // ← Release lock SỚM, trước khi lưu Drive
  }

  // 6. Lưu file lên Google Drive — NGOÀI lock (tốn 3-5s, không cần chiếm lock)
  if (needDriveSave && driveData) {
    try {
      var driveFolderUrl = saveFilesToDrive(driveData.data, driveData.details, driveData.score);
      if (driveFolderUrl) {
        try {
          firebaseUpdate('students/' + driveData.data.maHS, { driveFolder: driveFolderUrl });
        } catch (fe2) { Logger.log('Save driveFolder error: ' + fe2); }
      }
    } catch (err) {
      Logger.log('Lỗi lưu Drive: ' + err.toString());
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
    var score = data[i][5] ? parseFloat(data[i][5]) : 0;
    if (!map[maHS]) {
      map[maHS] = {
        maHS: maHS,
        hoTen: data[i][2] ? data[i][2].toString() : '',
        lop: data[i][3] ? data[i][3].toString() : '',
        bestScore: score,
        totalScore: score,
        attempts: 1,
        lastTime: data[i][6] ? formatTime(data[i][6]) : ''
      };
    } else {
      map[maHS].attempts++;
      map[maHS].totalScore += score;
      if (score > map[maHS].bestScore) {
        map[maHS].bestScore = score;
        map[maHS].lastTime = data[i][6] ? formatTime(data[i][6]) : '';
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
function saveFileOnly(data) {
  try {
    var rootFolder = getOrCreateFolder(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
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
    var rootFolder = getOrCreateFolder(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
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
function saveFilesToDrive(data, details, score) {
  // Tạo/tìm thư mục gốc KET_QUA_THI
  var rootFolder = getOrCreateFolder(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);

  // Tạo/tìm thư mục Lớp
  var classFolder = getOrCreateFolder(rootFolder, data.lop);

  // Tạo thư mục học sinh: HoTen_MaHS
  var studentFolderName = removeVietnameseTones(data.hoTen).replace(/\s+/g, '') + '_' + data.maHS;
  var studentFolder = getOrCreateFolder(classFolder, studentFolderName);

  // Lưu file thực hành
  if (data.fileData && data.fileName) {
    var fileBlob = Utilities.newBlob(
      Utilities.base64Decode(data.fileData),
      data.fileMimeType || 'application/octet-stream',
      data.fileName
    );
    studentFolder.createFile(fileBlob);
  }

  // Tạo file kết quả text - ĐẦY ĐỦ CHI TIẾT
  var now = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'dd/MM/yyyy HH:mm:ss');
  
  // Lấy thêm các đáp án từ Cau_Hoi
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var qSheet = ss.getSheetByName('Cau_Hoi');
  var qData = qSheet.getDataRange().getValues();
  var questionBank = {};
  for (var q = 1; q < qData.length; q++) {
    questionBank[qData[q][0]] = {
      noiDung: qData[q][1].toString(),
      dapAn1: qData[q][2].toString(),
      dapAn2: qData[q][3].toString(),
      dapAn3: qData[q][4].toString(),
      dapAn4: qData[q][5].toString(),
      dapAnDung: qData[q][6].toString().trim().toLowerCase()
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

  // Share folder and return URL for teacher access
  studentFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return studentFolder.getUrl();
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

  // Headers
  var csvRows = [];
  csvRows.push(['STT', 'Mã HS', 'Họ Tên', 'Lớp', 'C1', 'C2', 'C3', 'C4', 'C5', 'C6', 'C7', 'C8', 'Số câu đúng', 'Số câu sai', 'Điểm', 'Thời gian nộp'].join(','));

  var count = 0;
  for (var i = 1; i < rData.length; i++) {
    var lop = rData[i][3] ? rData[i][3].toString().trim() : '';
    if (lopFilter && lop !== lopFilter) continue;
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
    csvRows.push(row.join(','));
  }

  if (count === 0) {
    return { success: false, message: 'Không có dữ liệu kết quả để xuất!' };
  }

  // BOM + CSV content for Excel UTF-8
  var csvContent = '\uFEFF' + csvRows.join('\n');
  var fileName = 'KetQuaThi' + (lopFilter ? '_' + lopFilter : '') + '_' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyyMMdd_HHmmss') + '.csv';

  // Save to Drive
  var rootFolder = getOrCreateFolder(DriveApp.getRootFolder(), ROOT_FOLDER_NAME);
  var blob = Utilities.newBlob(csvContent, 'text/csv; charset=utf-8', fileName);
  var file = rootFolder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    success: true,
    message: 'Đã xuất ' + count + ' kết quả!',
    downloadUrl: file.getDownloadUrl(),
    fileName: fileName
  };
}

// ============== XÓA TOÀN BỘ DỮ LIỆU ==============
function clearAllData() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch(e) {
    return { success: false, message: 'Hệ thống đang bận!' };
  }
  try {
    var ss = SpreadsheetApp.openById(SHEET_ID);

    // 1. Xóa dữ liệu Ket_Qua (giữ header)
    var kqSheet = ss.getSheetByName('Ket_Qua');
    if (kqSheet && kqSheet.getLastRow() > 1) {
      kqSheet.deleteRows(2, kqSheet.getLastRow() - 1);
    }

    // 2. Xóa dữ liệu Ket_Qua_Thu (giữ header)
    var kqtSheet = ss.getSheetByName('Ket_Qua_Thu');
    if (kqtSheet && kqtSheet.getLastRow() > 1) {
      kqtSheet.deleteRows(2, kqtSheet.getLastRow() - 1);
    }

    // 3. Reset trạng thái thi + thời gian HS
    var hSheet = ss.getSheetByName('Hoc_Sinh');
    var hData = hSheet.getDataRange().getValues();
    for (var i = 1; i < hData.length; i++) {
      hSheet.getRange(i + 1, 6).setValue(''); // Trang_Thai
      hSheet.getRange(i + 1, 7).setValue(''); // Thoi_Gian_DN
      hSheet.getRange(i + 1, 8).setValue(''); // Thoi_Gian_NB
    }

    // 4. Xóa folder KET_QUA_THI trên Drive
    try {
      var folders = DriveApp.getRootFolder().getFoldersByName(ROOT_FOLDER_NAME);
      while (folders.hasNext()) {
        var folder = folders.next();
        folder.setTrashed(true);
      }
    } catch(de) { Logger.log('Drive delete error: ' + de); }

    // 5. Xóa dữ liệu Firebase
    try {
      firebaseDelete('results');
      firebaseDelete('practiceResults');
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
    } catch(fe) { Logger.log('Firebase clear error: ' + fe); }

    return { success: true, message: '✅ Đã xóa toàn bộ: kết quả thi, kết quả thi thử, trạng thái HS, file trên Drive và Firebase!' };
  } finally {
    lock.releaseLock();
  }
}

// ============== TẢI FILE HỌC SINH (GIÁO VIÊN) ==============

/**
 * Kiểm tra quyền GV với lớp — trả true nếu có quyền
 */
function _checkTeacherPermission(lop, maGV, isAdmin) {
  if (isAdmin === 'true' || isAdmin === true) return true;
  if (!maGV) return false;
  var teacher = firebaseGet('teachers/' + maGV.toString().trim());
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
function downloadClassFiles(lop, maGV, isAdmin) {
  if (!lop) return { success: false, message: 'Vui lòng chọn lớp!' };
  
  // Kiểm tra quyền
  if (!_checkTeacherPermission(lop, maGV, isAdmin)) {
    return { success: false, message: 'Bạn không có quyền tải file lớp ' + lop + '!' };
  }
  
  try {
    // Tìm folder lớp: KET_QUA_THI/{lop}/
    var rootFolders = DriveApp.getRootFolder().getFoldersByName(ROOT_FOLDER_NAME);
    if (!rootFolders.hasNext()) {
      return { success: false, message: 'Chưa có thư mục kết quả thi trên Drive!' };
    }
    var rootFolder = rootFolders.next();
    var classFolders = rootFolder.getFoldersByName(lop);
    if (!classFolders.hasNext()) {
      return { success: false, message: 'Chưa có thư mục lớp ' + lop + ' trên Drive!' };
    }
    var classFolder = classFolders.next();
    
    // Thu thập file từ tất cả sub-folder (mỗi HS 1 folder)
    var allBlobs = [];
    var studentCount = 0;
    var fileCount = 0;
    var studentFolders = classFolder.getFolders();
    
    while (studentFolders.hasNext()) {
      var sFolder = studentFolders.next();
      var sName = sFolder.getName(); // VD: NguyenVanA_HS001
      var sFiles = sFolder.getFiles();
      var hasFiles = false;
      
      while (sFiles.hasNext()) {
        var f = sFiles.next();
        var blob = f.getBlob();
        // Đặt tên: TênHSFolder/TênFile để giữ cấu trúc thư mục trong ZIP
        blob.setName(sName + '/' + f.getName());
        allBlobs.push(blob);
        fileCount++;
        hasFiles = true;
      }
      if (hasFiles) studentCount++;
    }
    
    if (allBlobs.length === 0) {
      // Không có file — trả folder URL cho GV xem
      classFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return { success: false, message: 'Lớp ' + lop + ' chưa có file bài làm nào!', folderUrl: classFolder.getUrl() };
    }
    
    // Nén ZIP
    var zipName = 'BaiLam_' + lop + '_' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyyMMdd_HHmm') + '.zip';
    var zipBlob = Utilities.zip(allBlobs, zipName);
    
    // Lưu ZIP lên folder gốc KET_QUA_THI
    var zipFile = rootFolder.createFile(zipBlob);
    zipFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
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
    // Fallback: trả URL folder nếu ZIP thất bại (quá lớn / timeout)
    try {
      var rootF = DriveApp.getRootFolder().getFoldersByName(ROOT_FOLDER_NAME);
      if (rootF.hasNext()) {
        var rf = rootF.next();
        var cf = rf.getFoldersByName(lop);
        if (cf.hasNext()) {
          var cFolder = cf.next();
          cFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          return {
            success: false,
            message: 'Không thể nén file (quá lớn hoặc timeout). Mở Google Drive để tải thủ công.',
            folderUrl: cFolder.getUrl(),
            fallback: true
          };
        }
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
function downloadStudentFiles(lop, maHS, maGV, isAdmin) {
  if (!lop || !maHS) return { success: false, message: 'Thiếu thông tin lớp hoặc mã HS!' };
  
  // Kiểm tra quyền
  if (!_checkTeacherPermission(lop, maGV, isAdmin)) {
    return { success: false, message: 'Bạn không có quyền tải file lớp ' + lop + '!' };
  }
  
  try {
    // Tìm folder học sinh
    var rootFolders = DriveApp.getRootFolder().getFoldersByName(ROOT_FOLDER_NAME);
    if (!rootFolders.hasNext()) return { success: false, message: 'Chưa có thư mục kết quả trên Drive!' };
    var rootFolder = rootFolders.next();
    
    var classFolders = rootFolder.getFoldersByName(lop);
    if (!classFolders.hasNext()) return { success: false, message: 'Chưa có thư mục lớp ' + lop + '!' };
    var classFolder = classFolders.next();
    
    // Tìm folder HS: tìm folder có chứa _MãHS trong tên
    var studentFolder = null;
    var subFolders = classFolder.getFolders();
    while (subFolders.hasNext()) {
      var sf = subFolders.next();
      if (sf.getName().indexOf('_' + maHS.trim()) >= 0) {
        studentFolder = sf;
        break;
      }
    }
    
    if (!studentFolder) {
      return { success: false, message: 'Chưa có thư mục bài làm của HS ' + maHS + '!' };
    }
    
    // Thu thập file
    var blobs = [];
    var files = studentFolder.getFiles();
    while (files.hasNext()) {
      var f = files.next();
      blobs.push(f.getBlob());
    }
    
    if (blobs.length === 0) {
      return { success: false, message: 'HS ' + maHS + ' chưa có file bài làm!' };
    }
    
    // Nếu chỉ 1 file → trả link tải trực tiếp
    if (blobs.length === 1) {
      var singleFile = studentFolder.getFiles().next();
      singleFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      return {
        success: true,
        message: 'Tải file: ' + singleFile.getName(),
        downloadUrl: singleFile.getDownloadUrl(),
        fileName: singleFile.getName(),
        fileCount: 1,
        isDirect: true
      };
    }
    
    // Nhiều file → nén ZIP
    var zipName = maHS.trim() + '_' + Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyyMMdd_HHmm') + '.zip';
    var zipBlob = Utilities.zip(blobs, zipName);
    var zipFile = rootFolder.createFile(zipBlob);
    zipFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return {
      success: true,
      message: 'Đã nén ' + blobs.length + ' file của HS ' + maHS,
      downloadUrl: zipFile.getDownloadUrl(),
      zipFileId: zipFile.getId(),
      fileName: zipName,
      fileCount: blobs.length
    };
  } catch (err) {
    // Fallback: mở folder HS trên Drive
    try {
      var student = firebaseGet('students/' + maHS.trim());
      if (student && student.driveFolder) {
        return {
          success: false,
          message: 'Không thể nén, mở Google Drive để tải thủ công.',
          folderUrl: student.driveFolder,
          fallback: true
        };
      }
    } catch (e2) { /* silent */ }
    return { success: false, message: 'Lỗi tải file: ' + err.toString() };
  }
}

/**
 * Xóa file ZIP tạm sau khi đã tải
 * @param {string} fileId - ID file ZIP cần xóa
 */
function cleanupDownloadZip(fileId) {
  if (!fileId) return { success: false, message: 'Thiếu fileId!' };
  try {
    var file = DriveApp.getFileById(fileId);
    file.setTrashed(true);
    return { success: true, message: 'Đã xóa file ZIP tạm.' };
  } catch (err) {
    return { success: false, message: 'Lỗi xóa file: ' + err.toString() };
  }
}

// ============== HÀM HỖ TRỢ ==============
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
    'driveFolder', 'uploadFailed', 'uploadError'];

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

  // 2. Sync Cau_Hoi — dùng multi-path update
  var qSheet = ss.getSheetByName('Cau_Hoi');
  var qData = qSheet.getDataRange().getValues();
  var questions = {};
  for (var j = 1; j < qData.length; j++) {
    var stt = qData[j][0];
    questions[stt] = {
      stt: stt,
      noiDung: qData[j][1].toString(),
      dapAn1: qData[j][2].toString(),
      dapAn2: qData[j][3].toString(),
      dapAn3: qData[j][4].toString(),
      dapAn4: qData[j][5].toString(),
      dapAnDung: qData[j][6].toString().trim().toLowerCase(),
      lop: qData[j][7] ? qData[j][7].toString().trim() : ''
    };
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
    for (var a = 0; a < 8; a++) {
      rowData.push(a < answers.length ? answers[a] : '');
    }
    rowData.push(r.score || 0);
    rowData.push(r.thoiGian ? new Date(r.thoiGian) : '');
    rSheet.appendRow(rowData);

    // Color cells
    var newRow = rSheet.getLastRow();
    for (var c = 0; c < 8; c++) {
      if (c < answers.length && answers[c]) {
        var cell = rSheet.getRange(newRow, 5 + c);
        if (answers[c].includes('✓')) cell.setBackground('#d4edda');
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
      pSheet.appendRow([p + 1, pr.maHS || '', pr.hoTen || '', pr.lop || '', pr.lanThu || 1, pr.score || 0, pr.thoiGian ? new Date(pr.thoiGian) : '']);
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

  // 1. Kiểm tra quyền GV
  if (!_checkTeacherPermission(lop, data.maGV, data.isAdmin)) {
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
function deleteStudentWithSync(maHS, maGV, isAdmin) {
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

  // 2. Kiểm tra quyền GV
  if (studentLop && !_checkTeacherPermission(studentLop, maGV, isAdmin)) {
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
    if (!_checkTeacherPermission(cls, data.maGV, data.isAdmin)) {
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
function deleteClassWithSync(lop, maGV, isAdmin) {
  if (!lop) return { success: false, message: 'Thiếu tên lớp!' };
  lop = lop.toString().trim();

  // 1. Kiểm tra quyền GV
  if (!_checkTeacherPermission(lop, maGV, isAdmin)) {
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
