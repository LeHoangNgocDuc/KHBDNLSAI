# Smart KHBD AI 1.0 – Tích hợp NLS, NLAI và HSHN vào giáo án Word gốc

**Dành cho giáo viên Trường THCS Nam Nha Trang – Năm học 2026–2027.**

Ứng dụng sử dụng giao diện HTML/CSS/JavaScript thuần, chạy trên Vercel với một Vercel Function nhỏ làm cầu nối Gemini API. **Không cần tải PPCT mỗi lần làm việc.** File DOCX được đọc và sửa trực tiếp trong trình duyệt; chỉ gửi ngữ cảnh ngắn gồm đề mục, hoạt động và chỉ số năng lực tới Gemini khi giáo viên chọn sử dụng API.

## 1. Tính năng đã lập trình

- Tải một file **`.docx`** (mặc định giới hạn 20 MB), tự đọc cấu trúc tài liệu, số bảng và các đối tượng Equation.
- Tự gợi ý môn, khối, tên bài, đối chiếu với danh mục KHDH Toán 6–9 và Tin học 6–9 (411 mục, gồm cả bài ôn tập/kiểm tra).
- Lấy mã và mô tả NLS/NLAI từ các nguồn KHDH, đồng thời ghi rõ trường hợp nguồn chỉ là **giáo án mẫu tham khảo**.
- Gemini đề xuất nội dung và vị trí chèn dưới dạng JSON; chỉ giữ lại những mã có thật trong dữ liệu được chọn.
- Có chế độ **không API** tạo đề xuất ban đầu dựa trên KHDH để giáo viên tự duyệt.
- Duyệt/sửa/xóa chọn từng đề xuất, đổi vị trí chèn, hoặc chọn "chưa xác định → ghi cuối giáo án".
- Chức năng nhập đề xuất thủ công cho trường hợp tài liệu KHDH thiếu mã. Nội dung thủ công được đánh dấu nguồn.
- Chèn vào **cấu trúc DOCX gốc** (không tạo lại file từ HTML); giữ nguyên mọi tệp ZIP khác ngoài `word/document.xml`.
- Có chống thêm lặp nội dung theo mã ở cùng mục/hoạt động và theo nội dung tương đương; nội dung HSHN đã có trong hoạt động được ưu tiên giữ nguyên.
- Tích hợp NLS màu đỏ, NLAI xanh lam, HSHN tím; font Times New Roman cho đoạn **mới**.
- Nếu không tìm được vị trí, ghi đủ: loại, mã, vị trí mong muốn, lý do, nội dung vào **phụ lục cuối giáo án**.
- Nhiều API Key do giáo viên nhập, chia nhóm project/quota, chọn model chính và dự phòng, chế độ luân phiên model và kiểm tra key.

**Điều chưa tự động làm:** Chuyển hoặc thay thế công thức Equation cũ. Đây là chủ ý nhằm giữ nguyên các công thức/đối tượng nhúng đang có. Nếu đề xuất văn bản của AI chứa LaTeX, hệ thống **chưa tạo Equation mới** từ LaTeX, vì vậy không nên yêu cầu AI bổ sung công thức toán dài; kiểm tra bản Word sau xuất khi có công thức.

## 2. TRIỂN KHAI LÊN GITHUB → VERCEL (dễ nhất)

1. Giải nén tệp ZIP này. Tạo kho GitHub mới (ví dụ `smart-khbd-ai`). **Tải toàn bộ nội dung bên trong thư mục `smart-khbd-ai`** lên gốc kho, không chỉ tải `index.html`. Quan trọng có: `vercel.json`, `package.json`, `api/gemini.js`, `public/index.html`, `public/src/`, `public/data/` và `public/vendor/`.
2. Vào [vercel.com](https://vercel.com), chọn **Add New → Project**, Import kho vừa tạo.
3. Chọn Framework Preset **Other**. Không cần Build Command (để trống). Output Directory **public**; cấu hình đã ghi trong `vercel.json`.
4. Nhấn Deploy. Đường dẫn có dạng `https://ten-du-an.vercel.app`.
5. Vào trang, nhấn biểu tượng **⚙**, nhập một hoặc nhiều Gemini API Key tạo tại [Google AI Studio](https://aistudio.google.com/app/apikey); bấm **Kiểm tra** rồi **Lấy danh sách model** để chọn model đang khả dụng.
6. Tải giáo án `.docx`, kiểm tra môn/lớp/bài/tiết, chọn NLS–NLAI–HSHN, nhấn **Phân tích và đề xuất**, duyệt nội dung và **Xuất Word**.

**Không cần điền API Key vào file mã nguồn hoặc Vercel Environment Variables.** Khóa được giáo viên nhập trong tab làm việc; chỉ gửi qua kết nối HTTPS tới `api/gemini.js` (cầu nối server) khi gọi Google. Trang không lưu khóa lên server, GitHub hoặc `localStorage`. Khi tải lại/đóng tab cần nhập lại khóa.

Nếu triển khai web cho nhiều người, nên bật **Deployment Protection / xác thực người dùng** tại Vercel để hạn chế truy cập không mong muốn. Không nên tải lên văn bản chứa thông tin cá nhân nhạy cảm của học sinh.

## 3. Quy trình sử dụng

1. **Chọn file Word gốc:** Không cần tạo lại giáo án. Ứng dụng không ghi đè file trên máy người dùng.
2. **Kiểm tra bài KHDH:** Trường `Bài học khớp ...` có thể chọn lại bằng tay. Với Toán, lưu ý trùng tên giữa Đại số/Hình học hoặc khác nhau số tiết.
3. **Tích hợp:** Tick những loại mong muốn, chọn mức độ và điều kiện thiết bị.
4. **AI hoặc không API:** Chọn `Gemini AI` để tạo đề xuất; chọn `Không dùng API` khi chưa có key hoặc chỉ muốn tham khảo các mô tả đã có trong KHDH.
5. **Duyệt:** Giữ hoặc bỏ từng đề xuất, chỉnh lại câu chữ và chọn đúng ô bảng/bước tổ chức. Khi nghi ngờ chỗ chèn, chọn `Chưa xác định → ghi vào cuối`.
6. **Xuất Word:** Hệ thống chạy kiểm tra XML, bảng, Equation và các đối tượng trước khi tạo file tải về. Giáo viên cần mở bằng Microsoft Word để rà lại ngắt trang và độ dài bảng.

### Nguồn dữ liệu tích hợp

- Toán 6: `KHDH Toan 6_26_27_Dieu chinh_8_9_2026.pdf`.
- Toán 7–9: 3 file `KHDH-TOÁN ...docx`.
- Tin học 6–9: 4 file `KHDH ...-cá nhân.docx`, ưu tiên vì có cột NLS và AI, trong khi các bản chia tuần chủ yếu cung cấp tiến độ.
- Riêng Toán 9 Bài 6, bộ dữ liệu bổ sung các mã từ giáo án mẫu `DS9_T9_T23_24_B6.docx` và **ghi rõ là mã tham khảo từ mẫu, không coi là mã trích từ KHDH Toán 9**.

Một số bài KHDH thực sự không có mã NLS hoặc NLAI. Chương trình không tự bịa mã. Giáo viên có thể bổ sung mã thủ công khi có nguồn chuyên môn thích hợp.

### Màu quy ước

- **NLS:** Times New Roman, `#FF0000`.
- **NLAI:** Times New Roman, `#0066CC`.
- **HSHN:** Times New Roman, `#7C3AED`.

## 4. Điều phối API/model

- **Tự động:** Ưu tiên model chính và key đang bật, chuyển model/key khi có lỗi phù hợp.
- **Luân phiên:** Luân phiên giữa model chính và dự phòng theo lần phân tích.
- **Thủ công:** Chỉ sử dụng model chính; vẫn có thể chuyển sang key khác đang bật khi cần.
- Khi HTTP **429** xuất hiện, đánh dấu nhóm project/quota đó không tiếp tục thử trong cùng lượt phân tích. Các key thuộc **cùng project dùng chung quota**; việc tạo thêm key không tăng hạn mức của project.
- Khi key bị thu hồi hoặc không đủ quyền, web báo lỗi. Khi tất cả key/model lỗi, file Word gốc được giữ nguyên, không có file lỗi âm thầm.

## 5. Nguyên tắc kỹ thuật bảo toàn tài liệu

`public/src/docx-engine.js` sử dụng JSZip và DOMParser trên trình duyệt để thao tác **thêm paragraph XML**, không xóa/mở rộng bảng, không viết lại các paragraph đã tồn tại. Sau khi chèn, kiểm tra:

1. Từng paragraph gốc vẫn giữ XML tương đương trước xử lý.
2. Số lượng bảng giữ nguyên.
3. Số đối tượng OMML và object Word giữ nguyên.
4. Tệp DOCX parse lại được, xuất dưới dạng DOCX hợp lệ.
5. Tất cả mục ZIP khác ngoài `word/document.xml` giữ nguyên (kiểm thử binary).

**Giới hạn kỹ thuật:** `XMLSerializer` có thể chuẩn hóa cách biểu diễn XML của tài liệu mà không làm thay đổi ý nghĩa OOXML. Do đó không thể cam kết toàn bộ file DOCX có byte-level equality với nguồn; mục tiêu là bảo toàn các thành phần, nội dung và định dạng. Nội dung bổ sung có thể làm bảng giãn trang. Tài liệu quá đặc biệt (liên kết OLE phức tạp, macro, nội dung được bảo vệ, bố cục VBA) cần kiểm tra thủ công. Chỉ nhận `.docx`, không nhận `.doc` và `.docm`.

## 6. Chạy thử không dùng Vercel

**Chỉ kiểm thử chức năng tải/đọc/soạn offline/xuất Word:**

```bash
python -m http.server 8080 -d public
```

Mở `http://localhost:8080`, chọn **Không dùng API**. Nút gọi Gemini cần Vercel Function, vì Python HTTP server chỉ phục vụ file tĩnh. Để thử nghiệm đầy đủ trên máy có cài Vercel CLI:

```bash
npx vercel dev
```

Hệ thống không có cơ sở dữ liệu; tất cả kế hoạch dạy học đã được đưa vào `public/data/catalog.json`.

### Cập nhật KHDH

Để cập nhật toàn bộ danh mục từ 8 file kế hoạch đúng tên được cung cấp, chỉnh thư mục nguồn trong `scripts/build_catalog.py` (biến `SRC`), cài `python-docx`, `PyMuPDF`, sau đó chạy:

```bash
python scripts/build_catalog.py
```

Kiểm tra đầu ra `public/data/catalog.json` và đối chiếu mã chỉ số trước khi đưa lên GitHub. Script chứa phần bổ sung tham khảo từ giáo án mẫu; có thể bỏ hoặc thay bằng dữ liệu KHDH được phê duyệt.

## 7. Kiểm thử

Với Node.js 20+ và Internet:

```bash
npm install
npm test
```

Bộ test chạy kiểm tra API proxy giả lập, vị trí chèn, chống chèn trùng, và so sánh binary ZIP (trừ `word/document.xml`). Trong môi trường phát triển đã chạy thử với file Toán 9 `DS9_T9_T23_24_B6.docx`: **7/7 test đạt**; file xuất có thể chuyển sang PDF bằng LibreOffice với 20 trang, kiểm tra hình ảnh đầu/cuối trang và bảng hoạt động. Đây chưa phải kiểm thử Gemini API thật với khóa người dùng hay nghiệm thu giao diện sau triển khai Vercel.

## 8. Cấu trúc mã nguồn

```text
smart-khbd-ai/
├── api/gemini.js                 # Vercel Function, không lưu API key
├── public/index.html            # Web mobile/desktop
├── public/src/main.js           # Giao diện, kiểm duyệt, quản trị API/model
├── public/src/catalog.js        # Đối chiếu bài/KHDH
├── public/src/docx-engine.js    # Chèn nội dung đúng vị trí DOCX gốc
├── public/src/styles.css        # UI EdTech
├── public/data/catalog.json     # 8 bộ kế hoạch được chuẩn hóa
├── public/vendor/jszip.min.js   # Thư viện ZIP độc lập, bản phân phối
├── scripts/build_catalog.py     # Tạo lại dữ liệu nguồn
├── tests/                        # Bộ test + file Word thử nghiệm tự tạo
├── package.json
└── vercel.json
```

Thư viện JSZip có giấy phép riêng trong `public/vendor/jszip-LICENSE.md`.

---

**Điều kiện nghiệm thu chuyên môn:** Luôn kiểm tra đúng môn, khối, bài, tiết, mã năng lực; trường hợp không tìm thấy vị trí, chuyển đề xuất xuống cuối và ghi rõ lý do. Giáo viên là người duyệt và quyết định việc sử dụng.
