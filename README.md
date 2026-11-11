# SMART KHBD AI 1.1.0 - Tích hợp NLS / NLAI / HSHN vào giáo án Word gốc

Trường THCS Nam Nha Trang - Năm học 2026-2027. Nâng cấp từ Smart KHBD AI 1.0.1.

## 1. Triển khai trên GitHub và Vercel

1. Giải nén `Smart_KHBD_AI_1.1.0_GitHub_Vercel.zip`.
2. Mở thư mục `smart-khbd-ai`. Upload **nội dung bên trong** thư mục đó lên **gốc** repository GitHub đang liên kết với web Vercel (không đặt thêm một lớp thư mục ở gốc). Thay thế các file trùng tên. Giữ lại các tệp từ bản cũ nếu có điều chỉnh cá nhân chưa được gộp.
3. Vercel sẽ tự build/deploy khi nhận commit nếu đang kết nối GitHub. Nếu tạo dự án mới: Framework `Other`, Output Directory `public`, Build Command để trống. Các cấu hình Function và header nằm trong `vercel.json`.
4. Sau khi Vercel triển khai, mở trang, Ctrl+Shift+R. Footer hiển thị `Smart KHBD AI 1.1.0`.
5. Mở Cài đặt (⚙), nhập Gemini API Key; chọn hoặc tự lấy danh sách model. Có thể chọn chế độ không API nếu cần kiểm tra luồng chỉnh sửa Word mà chưa có key.

## 2. Cách sử dụng

1. Chọn môn, khối và tải lên **giáo án gốc .docx**. Chọn đúng bài/tiết KHDH. Ứng dụng đã có danh mục Toán/Tin 6-9, không cần tải KHDH mỗi lần.
2. Chọn NLS, NLAI và/hoặc HSHN. Khi đánh dấu HSHN, một khung viền tím xuất hiện:
   - Radio: `Có tích hợp NLS` / `Không tích hợp NLS`; mặc định **không tích hợp NLS cho HSKT**.
   - Chọn nhiều dạng khuyết tật: nhìn, nghe, vận động, trí tuệ/phát triển, ngôn ngữ/giao tiếp, khác. Không cần lưu tên hay tình trạng y tế học sinh.
3. Tùy chọn **chuyển LaTeX đã có trong Word thành Equation** mặc định TẮT. Chỉ bật nếu Word gốc chứa công thức ở dạng `$...$` cần chuyển; lưu ý việc này sẽ thay đổi chính các đoạn có LaTeX thô (không thể vừa chuyển chúng vừa giữ nguyên tuyệt đối từng ký tự của đoạn đó).
4. Nhấn `PHÂN TÍCH VÀ ĐỀ XUẤT`, duyệt/sửa từng nội dung, kiểm tra vị trí, nhấn `XUẤT WORD GIỮ NGUYÊN BẢN GỐC`.
5. Luôn mở file DOCX xuất ra bằng Microsoft Word và xem lại vị trí chèn, phân trang, bảng, Equation trước khi sử dụng.

## 3. Quy tắc sư phạm bản 1.1.0

- **NLS** màu đỏ `#DC2626`, **NLAI** màu xanh lam `#2563EB`, **HSHN** màu tím `#7C3AED`. Nội dung thêm dùng Times New Roman.
- Mục tiêu NLS/NLAI ở cuối phần `2. Năng lực`, trước `3. Phẩm chất`.
- Mục tiêu HSKT ở cuối phần I, sau phần `3. Phẩm chất`, tạo khối riêng `* MỤC TIÊU DÀNH CHO HỌC SINH KHUYẾT TẬT` gồm kiến thức đơn giản hóa, năng lực chung/riêng, phẩm chất, `GDHN.KT`, `GDHN.KN`.
- NLS, NLAI gắn bài tập/ví dụ SGK ở `Bước 1: Chuyển giao nhiệm vụ`. NLS còn có thể tổng kết cuối Bước 1 và tạo nhiệm vụ trong Hướng dẫn về nhà nếu xác định được anchor.
- HSHN gắn `Bước 2: Thực hiện nhiệm vụ`, đúng cột hoạt động giáo viên và học sinh. HSKT không chọn NLS sẽ được hỗ trợ với phiếu gợi ý, nhiệm vụ nhỏ, máy tính cầm tay (khi cần), không giao thao tác phần mềm/Internet như nhiệm vụ NLS.
- AI chỉ trả lại **đề xuất có vị trí**. Chương trình tự kiểm tra anchor, cột và bước dạy trước khi chèn. Không được chèn vào cột `SẢN PHẨM DỰ KIẾN` chỉ vì đó là cột phải.
- Khi không tìm chắc vị trí chèn, chương trình đưa nội dung vào **phụ lục cuối giáo án**, ghi mã, vị trí dự kiến và nguyên nhân. Nội dung đó có trạng thái `APPENDED_UNPLACED`, không tính là đã chèn thành công.
- Chống lặp theo mã chỉ số và hoạt động: tài liệu đã có nội dung tương ứng sẽ không thêm lần nữa.

## 4. Equation và tài liệu nguồn

- **Công thức OMML/Equation gốc** trong DOCX được giữ nguyên. File DOCX được chỉnh sửa bằng OOXML, không tái tạo từ HTML.
- **Công thức LaTeX mới** do AI đề xuất trong `$...$` sẽ được chương trình chuyển thành **Word Equation (OMML)** khi xuất. Có hỗ trợ đại số thường dùng: số mũ, chỉ số, phân số, căn, bất đẳng thức, tập số, một số ký hiệu hình học, hệ phương trình `cases` dạng cơ bản.
- Khi bật tùy chọn, các công thức LaTeX cũ nằm **trọn trong một `w:t`** của Word được chuyển sang Equation. Công thức nằm rải qua nhiều run hoặc cú pháp không hỗ trợ không được chuyển tự động; báo thống kê nội dung không xử lý được. Có loại cú pháp không hỗ trợ sẽ ngăn việc chèn mới, tránh xuất công thức hỏng.
- File `.pdf` được cung cấp **chỉ là mẫu đối chiếu**. Để bảo toàn tài liệu tối đa và giữ các Equation gốc, hãy tải file **Word `.docx`**, không dùng PDF để thay Word. Việc tái tạo PDF thành Word sửa được không thể cam kết giữ nguyên 100% bảng, hình, công thức.
- Có thể thay đổi số trang hoặc vị trí xuống dòng do nội dung bổ sung. Các thuộc tính gốc của bảng, hình ảnh, công thức không được viết lại.

## 5. Bảo mật và API

- Nhiều Gemini API Key do người dùng nhập trong tab. Không lưu khóa vào `localStorage` hay GitHub. Khóa đi qua HTTPS tới Vercel Function để gọi Gemini. Cấu hình nhóm project/quota và model dự phòng giữ từ bản 1.0.1.
- Nhiều key trong cùng một Google Cloud project **không tăng quota**. Tên model hợp lệ phụ thuộc key; nút `Lấy danh sách model` lấy từ Google API thay vì tin hoàn toàn tên model có sẵn.
- Không tải các hồ sơ y tế, thông tin cá nhân học sinh vào giáo án cần gửi AI. Phần thiết lập chỉ cần loại hỗ trợ chung, không có thông tin định danh.
- Khi Gemini báo lỗi, hệ thống giữ nguyên tài liệu; có chế độ gợi ý từ KHDH không dùng API để giáo viên duyệt.

## 6. Kiểm thử

Yêu cầu Node.js >= 20. Trong thư mục chứa `package.json`, chạy:

```bash
npm install
npm test
```

Bản 1.1.0 có **21 kiểm thử tự động** (API proxy, model fallback, nhận diện bảng, chống chèn trùng, mục tiêu HSKT, chế độ NLS HSKT, LaTeX -> OMML, chuyển LaTeX đã có, bảo toàn các ZIP part ngoài `word/document.xml`). Đã dựng được file Word kiểm thử thành 20 trang bằng LibreOffice và kiểm tra hình Equation hiển thị đúng. **Chưa kiểm thử cuộc gọi Gemini thật bằng API Key của người dùng hoặc xác nhận UI trên Vercel sau deployment**.

## 7. Cấu trúc nguồn chính

```text
smart-khbd-ai/
  api/gemini.js               - API proxy Gemini / model list
  public/index.html           - UI responsive
  public/src/main.js          - Điều phối giao diện/AI/duyệt
  public/src/catalog.js       - Tra cứu KHDH và gợi ý offline
  public/src/inclusion-policy.js - Nhiệm vụ HSKT, 6 dạng hỗ trợ
  public/src/docx-engine.js   - Chỉnh sửa DOCX OOXML
  public/src/word-math.js     - LaTeX sang Word Equation (OMML)
  public/src/model-policy.js  - Quản lý ưu tiên model
  public/src/styles.css       - Giao diện
  public/data/catalog.json   - Dữ liệu KHDH Toán/Tin 6-9
  public/vendor/jszip.min.js  - JSZip
  tests/                     - Kiểm thử
  scripts/build_catalog.py    - Script lập lại danh mục KHDH
  package.json / vercel.json
```

**Tiêu chuẩn nghiệm thu:** kết quả AI là bản đề xuất, không phải nội dung bắt buộc; giáo viên duyệt trước khi xuất. Quy trình bảo toàn Word được ưu tiên hơn việc cố gắng chèn mọi nội dung ở những vị trí thiếu thông tin.
