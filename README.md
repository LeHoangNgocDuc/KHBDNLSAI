# Smart KHBD AI 1.2.0

Web bổ sung năng lực số (NLS), năng lực AI (NLAI) và hỗ trợ học sinh khuyết tật (HSHN/GDHN) vào **DOCX gốc**, dành cho giáo viên THCS Toán và Tin học 6–9.

## Các nâng cấp của bản 1.2.0

- **Mục tiêu HSHN đầy đủ**: phần mới nằm sau `3. Phẩm chất` và trước `II. Thiết bị dạy học`, với ba mục `Kiến thức`, `Năng lực` (chung/riêng), `Phẩm chất`, cộng `[GDHN.KT]` và `[GDHN.KN]`. Các nội dung Kiến thức và hai Phẩm chất lấy từ chính mục I của bài học đang chọn, lược bớt yêu cầu phức tạp; giáo viên được duyệt trước khi xuất.
- **Nhận dạng mẫu Tin học**: chấp nhận `2.1. Năng lực chung`, `2.2. Năng lực Tin học`, bảng `Tổ chức thực hiện (Hoạt động của GV và HS)`; tên hoạt động dạng `1. Hoạt động 1`, `2.1. Hoạt động 1`, `2.1.2. Hoạt động 2`, v.v. NLS và NLAI ưu tiên Bước 1; hỗ trợ HSHN ưu tiên Bước 2, bên cột hướng dẫn giáo viên/học sinh.
- **DOCX nhiều bài**: tự phát hiện các bài, cho chọn đúng bài đang xử lý, không gán đoạn sang bài kế tiếp. Với tài liệu không đủ thông tin phân đoạn, cần kiểm tra thủ công vị trí đề xuất.
- **Tùy chọn nâng cao**: giao diện thẻ màu, phần HSHN mở rộng khi tích; tùy chọn Có/Không tích hợp NLS cho HSKT (mặc định Không), sáu dạng khuyết tật có thể chọn nhiều. Không yêu cầu thông tin định danh hay hồ sơ y tế học sinh.
- **Không trùng HSHN ở bước dạy**: một đoạn hỗ trợ đã có ở Bước 1 không làm mất khả năng thêm hỗ trợ tại Bước 2 khi thích hợp.
- **Bảo toàn DOCX**: thay đổi tối thiểu `word/document.xml`; mọi tệp khác trong gói DOCX được giữ nguyên bytes. Không tái tạo bảng/hình/Equation từ HTML. Các đoạn mới dùng Times New Roman, kế thừa cỡ chữ đoạn gốc; NLS `#DC2626`, NLAI `#2563EB`, HSHN `#7C3AED`.
- **LaTeX mới** trong đề xuất được chuyển thành Equation Word (OMML). Tùy chọn chuyển LaTeX đã tồn tại trong Word là **opt-in** vì sửa trực tiếp nội dung cũ.
- **Dự phòng an toàn**: không xác định đúng vị trí thì chuyển đề xuất xuống phụ lục cuối, ghi lý do. Không tự chèn vào ô Sản phẩm.

**Các tính năng Ngoại ngữ, Song ngữ Việt–Anh và Toán KTD** có các thẻ tương ứng theo giao diện tham khảo nhưng **chưa kích hoạt**, không có checkbox giả hoạt động.

## Triển khai GitHub + Vercel

- Đẩy các tệp dự án vào GitHub (giữ nguyên thư mục `api/`, `public/`, `package.json`, `vercel.json`).
- Vercel → Add New Project → Import Git Repository. Framework `Other`, Output Directory `public`, Build Command để trống.
- Với web đã triển khai, giải nén ZIP **Chỉ cập nhật** và sao chép đè đúng đường dẫn; commit để Vercel tự redeploy. Không cần đổi URL.
- Sau khi triển khai, tải lại bằng `Ctrl + Shift + R` và kiểm tra footer `1.2.0`.

## Sử dụng

1. Tải một file **`.docx` gốc**. PDF chỉ dùng làm tài liệu tham khảo, không phải nguồn xuất Word bảo toàn định dạng.
2. Nếu file có nhiều bài (như mẫu Tin 7), chọn **bài trong file Word**. Kiểm tra môn/khối/tiết KHDH mà hệ thống ghép.
3. Chọn NLS, NLAI và/hoặc HSHN. Với HSHN, chọn Có/Không tích hợp NLS và dạng hỗ trợ phù hợp.
4. Phân tích bằng Gemini hoặc chọn chế độ không API. Duyệt từng đoạn, thay đổi/loại bỏ theo chuyên môn.
5. Xuất Word. Kiểm tra lại bằng Microsoft Word, nhất là trang chứa bảng và công thức toán học.

### Cấu hình API

Gemini API Key được người sử dụng nhập trong Cài đặt; cơ chế chọn model và key dự phòng vẫn giữ từ 1.1.0. Mỗi project Google Cloud dùng chung quota cho các key thuộc project đó; ứng dụng không tăng quota bằng cách đổi key.

### Kiểm thử

Chạy `npm test` (Node.js 20+) hoặc `node --test tests/*.test.mjs`. Trong môi trường kiểm thử đã chạy **26/26** bài kiểm thử, bao gồm mẫu Tin 7 gồm ba bài trong một Word và mẫu Toán Bài 36 có bảng, hình, Equation. Đã dựng ảnh Word và kiểm tra phần mục tiêu/ô Bước 2. **Chưa xác nhận** gọi Gemini thật với key cá nhân và thử nghiệm trình duyệt sau deploy Vercel.

### Lưu ý chuyên môn

HSHN điều chỉnh dựa trên nhu cầu cá nhân và kế hoạch giáo dục cá nhân nếu có, không chỉ dựa trên nhãn dạng khuyết tật. Chế độ `Không tích hợp NLS cho HSKT` không xóa hoạt động Tin học cơ bản của bài; chỉ không giao thêm nhiệm vụ phát triển chỉ số NLS riêng cho HSKT.
