# 🔒 Hướng dẫn cài đặt Firebase Security Rules

## ⚠️ SỬA LỖI KHẨN CẤP

Nếu dashboard GV đang bị lỗi "permission_denied", hãy paste rules dưới đây **ngay lập tức**.

## Bước 1: Truy cập Firebase Console
1. Mở [Firebase Console](https://console.firebase.google.com/)
2. Chọn project **thi-truc-tuyen-967d7**
3. Menu bên trái → **Realtime Database** → tab **Rules**

## Bước 2: Paste rules sau đây (THAY THẾ TOÀN BỘ)

```json
{
  "rules": {
    ".read": true,
    ".write": true
  }
}
```

## Bước 3: Nhấn **Publish**

## ⚠️ Giải thích

Firebase Realtime Database **không hỗ trợ chặn đọc field con** nếu cha đã cho phép đọc. Ví dụ: không thể cho đọc `/questions` nhưng chặn `/questions/1/dapAnDung` — đây là giới hạn của Firebase RTDB.

**Vì ứng dụng không dùng Firebase Auth**, không phân biệt được ai là HS / ai là GV ở cấp rules. Do đó:
- Bảo mật **mật khẩu GV** → đã giải quyết bằng **SHA-256 hash** trong code
- Bảo mật **đáp án đúng** → đã giải quyết bằng cách **không gửi `dapAnDung` về client** trong code JS
- Rules giữ đơn giản: cho phép đọc/ghi tất cả

> 💡 Nếu muốn bảo mật mạnh hơn trong tương lai, cần bật **Firebase Anonymous Auth** và rewrite rules dựa trên `auth.uid`.
