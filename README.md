# 📚 HSK Flashcard Learner - Modern Web App

Ứng dụng học từ vựng HSK bằng Flashcard giao diện Quizlet hiện đại, được viết hoàn toàn bằng **HTML5, CSS3, JavaScript (ES6) thuần**. Không dùng framework, không backend, chạy mượt mà 100% trên **GitHub Pages**.

![App Preview](https://img.shields.io/badge/Status-Active-brightgreen) ![License](https://img.shields.io/badge/License-MIT-blue)

## 🌟 Tính năng nổi bật

- 🗂 **Đọc trực tiếp File Excel**: Sử dụng thư viện `SheetJS` đọc file `HSK1_flashcards.xlsx` linh hoạt.
- 🔄 **Hiệu ứng 3D Flashcard**: Flip hiệu ứng mượt như Quizlet.
- 🧠 **Thuật toán học thông minh**: Chọn **Chưa nhớ** lập tức đẩy từ vựng về cuối danh sách để ôn lại.
- 📊 **Thống kê tiến độ trực quan**: Dashboard đo đếm số từ đã nhớ, chưa nhớ và tỷ lệ % hoàn thành.
- 🔍 **Tìm kiếm & Bộ lọc**: Tìm nhanh từ theo Hanzi, Pinyin, Nghĩa; lọc từ khó (Favorite ⭐).
- 🌙 **Dark Mode**: Tối ưu học đêm, giảm mỏi mắt.
- 💾 **Lưu tiến độ**: Sử dụng `LocalStorage` tự động lưu lại toàn bộ quá trình học.
- ⌨ **Hỗ trợ phím tắt**:
  - `Space`: Lật thẻ
  - `Mũi tên Trái/Phải`: Chuyển thẻ
  - `Phím 1`: Chưa nhớ
  - `Phím 2`: Đã nhớ

---

## 🚀 Hướng dẫn Deploy lên GitHub Pages

1. **Tạo Repository mới trên GitHub**:
   - Đặt tên repository (ví dụ: `hsk1-flashcard`).

2. **Upload các file vào thư mục gốc (`/`)**:
   - `index.html`
   - `style.css`
   - `script.js`
   - `README.md`
   - `HSK1_flashcards.xlsx`

3. **Bật GitHub Pages**:
   - Vào **Settings** của Repository trên GitHub.
   - Chọn mục **Pages** ở thanh menu bên trái.
   - Tại mục **Build and deployment** -> **Source**, chọn `Deploy from a branch`.
   - Chọn Branch `main` (hoặc `master`) và thư mục `/root`, sau đó nhấn **Save**.
   - Chờ 1-2 phút, truy cập đường dẫn do GitHub cung cấp (ví dụ: `https://your-username.github.io/hsk1-flashcard`).

---

## 📊 Chuẩn hóa cấu trúc File Excel (`HSK1_flashcards.xlsx`)

Đảm bảo file Excel có đúng 5 cột sau ở dòng đầu tiên:

| Hanzi | Pinyin | Nghĩa | Ví dụ VN | Ví dụ pinyin |
| :--- | :--- | :--- | :--- | :--- |
| 爱 | ài | Yêu, thích | Tôi yêu gia đình tôi. | Wǒ ài wǒ de jiātíng. |
| 八 | bā | Số 8 | Anh ấy có 8 cái bánh. | Tā yǒu bā gè dànɡāo. |
