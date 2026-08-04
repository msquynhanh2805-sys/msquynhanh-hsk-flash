# 📚 HSK 1 Flashcard Learner

Ứng dụng web học từ vựng HSK 1 bằng Flashcard theo phong cách **Quizlet** hiện đại, mượt mà. Đọc dữ liệu trực tiếp từ file Excel `HSK1_flashcards.xlsx` (143 từ vựng HSK 1 chuẩn).

Được xây dựng hoàn toàn bằng **HTML5, CSS3, JavaScript (ES6) thuần**, không phụ thuộc framework hay backend, sẵn sàng deploy trực tiếp trên **GitHub Pages**.

---

## 🔥 Tính năng nổi bật

- 🗂 **Đọc file Excel linh hoạt**: Tự động parse dữ liệu từ `HSK1_flashcards.xlsx` thông qua thư viện `SheetJS`.
- 🎴 **Hiệu ứng lật thẻ 3D**: Lật mặt mượt mà, hiển thị rõ ràng chữ Hán, Pinyin, Nghĩa tiếng Việt và Ví dụ minh họa.
- 🧠 **Thuật toán ôn tập thông minh**: 
  - Chọn **"Đã nhớ"**: Lưu tiến độ hoàn thành.
  - Chọn **"Chưa nhớ"**: Đưa từ vựng xuống cuối danh sách để lặp lại cho tới khi nhớ.
- 🔍 **Tìm kiếm & Bộ lọc linh hoạt**:
  - Tìm theo **Hanzi**, **Pinyin**, hoặc **Nghĩa tiếng Việt**.
  - Lọc danh sách theo: *Tất cả*, *Chưa nhớ*, *Đã nhớ*, *Từ khó (⭐)*.
- ⭐ **Đánh dấu từ khó (Favorite)**: Dễ dàng lưu các từ hay quên để tập trung ôn tập riêng.
- 🔀 **Học ngẫu nhiên**: Xáo trộn từ vựng (Shuffle) tránh học vẹt theo thứ tự.
- 📊 **Thống kê tiến độ trực quan**:
  - Đếm tổng số từ, số từ đã nhớ, chưa nhớ.
  - Thanh tiến độ (Progress bar) và % hoàn thành thời gian thực.
- 🌙 **Giao diện sáng / tối (Dark Mode)**: Chuyển đổi giao diện bảo vệ mắt khi học ban đêm.
- 💾 **Tự động lưu tiến độ (LocalStorage)**: Giữ nguyên trạng thái học tập kể cả khi tắt trình duyệt hay reload trang.
- 🎉 **Thông báo hoàn thành**: Popup chúc mừng sinh động khi học xong toàn bộ bộ từ.
- ⌨️ **Hỗ trợ Phím tắt tiện lợi**:
  - `Space`: Lật thẻ
  - `Mũi tên Trái (←)` / `Mũi tên Phải (→)`: Từ trước / Từ sau
  - `Phím 1`: Đánh dấu "Chưa nhớ"
  - `Phím 2`: Đánh dấu "Đã nhớ"

---

## 📁 Cấu trúc thư mục Project

```text
/
├── index.html            # Giao diện ứng dụng chính
├── style.css             # CSS Styling, 3D flip effect, Responsive, Dark Mode
├── script.js             # Logic xử lý chính, SheetJS integration, LocalStorage
├── README.md             # Tài liệu hướng dẫn sử dụng & deploy
└── HSK1_flashcards.xlsx  # File dữ liệu 143 từ vựng HSK 1
