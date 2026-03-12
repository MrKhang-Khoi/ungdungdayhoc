# 🔒 Hướng dẫn cài đặt Firebase Security Rules

## Bước 1: Truy cập Firebase Console
1. Mở [Firebase Console](https://console.firebase.google.com/)
2. Chọn project **thi-truc-tuyen-967d7**
3. Menu bên trái → **Realtime Database** → tab **Rules**

## Bước 2: Paste rules sau đây

```json
{
  "rules": {
    "students": {
      "$maHS": {
        ".read": true,
        ".write": "auth != null || true",
        "lastSeen": { ".write": true },
        "thoiGianDN": { ".write": true },
        "sessionId": { ".write": true },
        "cheatingLog": { ".write": true },
        "cheatingCount": { ".write": true },
        "trangThai": { ".write": true },
        "matKhau": { ".read": true },
        "hoTen": { ".write": true },
        "lop": { ".write": true }
      }
    },
    "questions": {
      ".read": true,
      ".write": true,
      "$questionId": {
        "dapAnDung": {
          ".read": false
        }
      }
    },
    "results": {
      ".read": true,
      ".write": true
    },
    "practiceResults": {
      ".read": true,
      ".write": true
    },
    "settings": {
      ".read": true,
      ".write": true
    },
    "notifications": {
      ".read": true,
      ".write": true
    }
  }
}
```

## Bước 3: Nhấn **Publish**

## ⚠️ Lưu ý quan trọng
- Rule `"dapAnDung": { ".read": false }` sẽ chặn HS đọc đáp án đúng từ client
- Vì app không dùng Firebase Auth, rules dựa trên path-based access control
- Sau khi deploy rules, cần **test lại** đăng nhập HS để đảm bảo app vẫn hoạt động
- Nếu app bị lỗi sau khi set rules, tạm thời set tất cả về `true` để khôi phục
