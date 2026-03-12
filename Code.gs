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
  var chHeaders = ['STT', 'Noi_Dung', 'Dap_An_1', 'Dap_An_2', 'Dap_An_3', 'Dap_An_4', 'Dap_An_Dung'];
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

  // Thêm 16 câu hỏi mẫu (Tin học)
  var sampleQuestions = [
    [1, 'Đơn vị nhỏ nhất của thông tin trong máy tính là gì?', 'Byte', 'Bit', 'KB', 'MB', 'b'],
    [2, '1 Byte bằng bao nhiêu Bit?', '4 Bit', '8 Bit', '16 Bit', '32 Bit', 'b'],
    [3, 'CPU là viết tắt của từ gì?', 'Central Processing Unit', 'Computer Personal Unit', 'Central Program Utility', 'Computer Processing Unit', 'a'],
    [4, 'Phần mềm nào dùng để soạn thảo văn bản?', 'Excel', 'PowerPoint', 'Word', 'Access', 'c'],
    [5, 'RAM là bộ nhớ gì?', 'Bộ nhớ ngoài', 'Bộ nhớ trong (tạm thời)', 'Bộ nhớ chỉ đọc', 'Bộ nhớ vĩnh viễn', 'b'],
    [6, 'Phím tắt Ctrl+C dùng để làm gì?', 'Cắt', 'Sao chép', 'Dán', 'In', 'b'],
    [7, 'Phím tắt Ctrl+V dùng để làm gì?', 'Sao chép', 'Cắt', 'Dán', 'Lưu', 'c'],
    [8, 'Hệ điều hành nào phổ biến nhất trên máy tính cá nhân?', 'Linux', 'MacOS', 'Windows', 'Android', 'c'],
    [9, 'Thiết bị nào là thiết bị đầu vào?', 'Máy in', 'Loa', 'Bàn phím', 'Màn hình', 'c'],
    [10, 'Thiết bị nào là thiết bị đầu ra?', 'Chuột', 'Bàn phím', 'Máy quét', 'Máy in', 'd'],
    [11, 'Đuôi file .docx thuộc phần mềm nào?', 'Excel', 'Word', 'PowerPoint', 'Access', 'b'],
    [12, 'Đuôi file .xlsx thuộc phần mềm nào?', 'Word', 'Excel', 'PowerPoint', 'Notepad', 'b'],
    [13, 'Phím tắt Ctrl+Z dùng để làm gì?', 'Lưu file', 'Hoàn tác', 'Làm lại', 'Đóng file', 'b'],
    [14, 'Internet là gì?', 'Phần mềm máy tính', 'Mạng máy tính toàn cầu', 'Hệ điều hành', 'Thiết bị phần cứng', 'b'],
    [15, 'Virus máy tính là gì?', 'Phần cứng hỏng', 'Chương trình gây hại', 'Lỗi hệ điều hành', 'File bị xóa', 'b'],
    [16, 'Phím tắt Ctrl+S dùng để làm gì?', 'Sao chép', 'Lưu file', 'Tìm kiếm', 'Chọn tất cả', 'b']
  ];
  chSheet.getRange(2, 1, sampleQuestions.length, 7).setValues(sampleQuestions);

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
  var isPractice = settingsResult.settings.practiceMode === 'true';

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

// Đếm số lần thi thử của HS
function countPracticeAttempts(ss, maHS) {
  var sheet = getOrCreatePracticeSheet(ss);
  var data = sheet.getDataRange().getValues();
  var count = 0;
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] && data[i][1].toString().trim() === maHS) count++;
  }
  return count;
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
  // Race condition protection: chỉ 1 HS nộp bài tại 1 thời điểm
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // chờ tối đa 15 giây
  } catch (e) {
    return { success: false, message: '⏳ Hệ thống đang bận, vui lòng thử lại sau vài giây!' };
  }

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
      return { success: false, message: '⚠️ Bạn đã hết 5 lượt thi thử! Không thể nộp thêm.' };
    }
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

    return {
      success: true,
      message: '🎯 Nộp bài thi thử thành công!',
      score: score,
      totalQuestions: totalQ,
      details: details,
      attempt: attempts + 1,
      maxAttempts: 5,
      isPractice: true
    };
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
        answers: answerResults,
        correctCount: correctCount,
        wrongCount: wrongCount,
        score: score,
        thoiGian: Date.now(),
        synced: true
      });
      // Update student status on Firebase
      firebaseUpdate('students/' + data.maHS, {
        trangThai: 'X',
        thoiGianNB: Date.now()
      });
    } catch (fe) { Logger.log('Firebase result write error: ' + fe); }

    // 6. Lưu file lên Google Drive
    try {
      saveFilesToDrive(data, details, score);
    } catch (err) {
      Logger.log('Lỗi lưu Drive: ' + err.toString());
    }

    return {
      success: true,
      message: '🎉 Nộp bài thành công!',
      score: score,
      isPractice: false
    };
  }

  } finally {
    lock.releaseLock();
  }
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
    return { success: true, message: 'Đã lưu file!' };
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
  resultText += '📊 Điểm trắc nghiệm : ' + score + ' / 4.0\n';
  resultText += '   Số câu đúng      : ' + (score / 0.5) + ' / 8\n\n';
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

// ============== ĐỒNG BỘ GGSHEET → FIREBASE ==============
function syncSheetToFirebase() {
  var ss = SpreadsheetApp.openById(SHEET_ID);

  // 1. Sync Hoc_Sinh
  var hSheet = ss.getSheetByName('Hoc_Sinh');
  var hData = hSheet.getDataRange().getValues();
  var students = {};
  for (var i = 1; i < hData.length; i++) {
    var maHS = hData[i][1].toString().trim();
    if (!maHS) continue;
    students[maHS] = {
      stt: hData[i][0],
      hoTen: hData[i][2].toString(),
      lop: hData[i][3].toString(),
      matKhau: hData[i][4].toString(),
      trangThai: hData[i][5] ? hData[i][5].toString().trim() : '',
      thoiGianDN: hData[i][6] ? new Date(hData[i][6]).getTime() : null,
      thoiGianNB: hData[i][7] ? new Date(hData[i][7]).getTime() : null
    };
  }
  firebaseSet('students', students);

  // 2. Sync Cau_Hoi
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
      dapAnDung: qData[j][6].toString().trim().toLowerCase()
    };
  }
  firebaseSet('questions', questions);

  // 3. Sync Giao_Vien
  var gvSheet = ss.getSheetByName('Giao_Vien');
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
    firebaseSet('teachers', teachers);
  }

  // 4. Sync De_Thi
  var dtSheet = ss.getSheetByName('De_Thi');
  if (dtSheet) {
    var dtData = dtSheet.getDataRange().getValues();
    for (var d = 1; d < dtData.length; d++) {
      var examId = dtData[d][0].toString().trim();
      if (!examId) continue;
      // Only sync exam metadata, not questions (questions are managed separately)
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
  var teacherCount = gvSheet ? Object.keys(teachers).length : 0;
  firebaseUpdate('meta', {
    lastSyncFromSheet: Date.now(),
    totalStudents: Object.keys(students).length,
    totalQuestions: Object.keys(questions).length,
    totalTeachers: teacherCount
  });

  return {
    success: true,
    message: '✅ Đồng bộ thành công! ' + Object.keys(students).length + ' HS, ' + Object.keys(questions).length + ' câu hỏi, ' + teacherCount + ' GV đã được push lên Firebase.'
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
      dapAnDung: data[i][6].toString().trim().toLowerCase()
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
      return { success: true, message: 'Cập nhật câu ' + params.stt + ' thành công!' };
    }
  }

  return { success: false, message: 'Không tìm thấy câu hỏi STT ' + params.stt };
}
