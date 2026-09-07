# Ký Hiệu Toán Học — Word Add-in

Task pane add-in cho Microsoft Word (Windows) giúp chèn nhanh ký hiệu toán học Unicode và công thức Word Equation (OMML).

## Tính năng

- 6 nhóm ký hiệu Unicode: Hy Lạp, Toán tử, Tập hợp, Logic, Mũi tên, Giải tích — click để chèn vào vị trí con trỏ.
- 51 mẫu công thức OMML chia 4 nhóm — chèn thành Equation Object thật của Word với **ô trống điền được**, Tab để nhảy giữa các ô, lồng được vào nhau:
  - *Phân số & mũ*: 4 kiểu phân số, căn bậc hai/bậc n, lũy thừa, chỉ số dưới/trên/trước, mũ của ngoặc.
  - *Tổng & tích phân*: Σ, Π, ∫ (có cận/không cận/hai lớp/đường), ⋃, ⋂.
  - *Ngoặc & ma trận*: 6 loại ngoặc, tổ hợp, hệ 2–3 phương trình, ma trận 2x2/3x3, định thức, vector cột.
  - *Hàm & dấu*: lim, max, sin/cos/tan/ln, log cơ số, đạo hàm thường và riêng, vector, hat, chấm trên, gạch trên/dưới, ngoặc nhọn trên/dưới.
- **Ba tab công cụ soạn đề thi**:
  - *Đề thi*: chọn **Times New Roman cỡ 12** (hoặc phông/cỡ khác) áp cho mọi thứ add-in chèn ra, và chọn **kiểu đề**:
    - *Tự luận / học sinh giỏi*: đầu đề Phòng GD&ĐT + ĐỀ CHÍNH THỨC / kỳ thi / lớp + năm học / MÔN / thời gian "(không kể thời gian giao đề)" / số trang; câu dạng `Câu 1 (4,0 điểm).` kèm các ý a) b) c); dòng kết `…Hết…` với lời dặn và chỗ ghi họ tên; bảng hướng dẫn chấm `Câu | Ý | Nội dung | Điểm` và bảng cấu trúc đề `Câu | Nội dung | Điểm`.
    - *Trắc nghiệm*: đầu đề Sở/Trường có mã đề, tiêu đề Phần I/II/III theo cấu trúc đề THPT, chèn hàng loạt câu trắc nghiệm A–D (1/2/4 cột), câu đúng/sai a–d, câu trả lời ngắn.

    Cả hai kiểu đều nhớ lại nội dung đã nhập cho lần mở sau, và ô "Bắt đầu từ câu" tự tăng sau mỗi lần chèn.
  - *Bảng BT*: dựng bảng biến thiên (3 dòng) hoặc bảng xét dấu (2 dòng) — nhập nghiệm, chọn dấu từng khoảng, xem trước ngay trong task pane rồi chèn ra thành bảng Word thật; mũi tên ↗ ↘ tự suy ra từ dấu đạo hàm.
  - *Đáp án*: chèn bảng đáp án từ chuỗi `1A 2B 3C` hoặc `ABCD...` (cả đáp án đúng/sai kiểu `1 ĐSSĐ`), bảng đáp án trống, đánh số lại toàn bộ câu hỏi, trộn thứ tự câu hỏi và chuyển bảng đáp án cũ sang thứ tự mới.
- Tìm kiếm theo tên tiếng Việt hoặc tiếng Anh (`alpha`, `integral`, `phân số`...).
- Đánh dấu yêu thích (★) — lưu trong localStorage của task pane.

## Cấu trúc

```
manifest.xml          # Manifest add-in
set-host.ps1          # Script điền URL GitHub Pages vào manifest
installer/
  MathSymbols.iss     # Script Inno Setup tạo MathSymbolsSetup.exe
  build.ps1           # Biên dịch bộ cài
  Register-Addin.ps1  # Đăng ký/gỡ add-in bằng PowerShell (triển khai hàng loạt)
assets/               # Icon add-in (16/32/64/80/128 px)
taskpane/
  taskpane.html       # Giao diện task pane
  taskpane.css        # Style (hỗ trợ dark mode)
  taskpane.js         # Logic render + chèn vào Word qua Office.js
  symbols.js          # Dữ liệu ký hiệu + mẫu OMML + hàm đóng gói Flat OPC
  exam.js             # Dựng OOXML cho đầu đề, câu hỏi, bảng biến thiên, bảng đáp án
  tools.js            # Form của 3 tab công cụ soạn đề thi
```

Không có build step — toàn bộ là static file, host thẳng lên GitHub Pages.

Chạy thử tại chỗ (không cần Word, các nút sẽ copy OOXML vào clipboard thay vì chèn):

```bash
python -m http.server 8144
```

rồi mở `http://localhost:8144/taskpane/taskpane.html`.

## Triển khai lên GitHub Pages

1. Tạo repo GitHub tên `math_symbols` và push code lên nhánh `main`.
2. Trong repo: **Settings → Pages → Source: Deploy from a branch → Branch: `main` / `(root)`**.
3. Điền URL thật vào manifest:

```bash
powershell -File set-host.ps1 -User duchop0974 -Repo math_symbols
```

4. Commit và push lại `manifest.xml` đã cập nhật.
5. Kiểm tra `https://duchop0974.github.io/math_symbols/taskpane/taskpane.html` mở được trên trình duyệt.

## Cài vào Word bằng bộ cài

Tải bộ cài mới nhất tại [Releases](https://github.com/duchop0974/math_symbols/releases/latest).

Người dùng cuối chỉ cần chạy `MathSymbolsSetup.exe`:

1. Đóng Microsoft Word.
2. Chạy `MathSymbolsSetup.exe` (cài theo từng user, **không cần quyền admin**).
3. Mở Word → tab **Home** → nút **Add-ins** → mục **Developer Add-ins** → bấm **Ký Hiệu Toán Học**.
4. Từ lần này trở đi, nút **Bảng Ký Hiệu** nằm sẵn ở tab Home.

Bước 3 chỉ phải làm một lần trên mỗi máy: Word nạp manifest ngay khi khởi động, nhưng chỉ gắn nút lên ribbon sau lần mở đầu tiên.

## Lưu ý khi trộn đề

Chức năng **Trộn thứ tự câu hỏi** đọc tài liệu qua Office.js, nhận diện mọi đoạn bắt đầu bằng `Câu 1:` / `Câu 1.` / `Câu 1)`, rồi đảo thứ tự các khối đó **từ câu đầu tiên đến hết tài liệu** và đánh số lại từ 1.

- Lưu tài liệu trước khi trộn — thao tác này ghi đè phần thân đề.
- Để bảng đáp án ở file riêng, đừng để cuối đề, vì nó cũng nằm trong vùng bị trộn.
- Sau khi trộn, task pane hiện bảng *câu mới ← câu gốc*; dán đáp án đề gốc vào ô "Chuyển đáp án theo thứ tự mới" để lấy đáp án tương ứng.

Add-in nạp trực tiếp từ GitHub Pages nên khi cập nhật tính năng chỉ cần push code — người dùng **không phải cài lại** bộ cài.

Gỡ cài đặt: **Settings → Apps → Ký Hiệu Toán Học → Uninstall** (bộ cài tự xoá khoá registry đã ghi).

Add-in tải giao diện từ GitHub Pages nên **máy cần có Internet** khi dùng.

### Cách bộ cài hoạt động

- Copy `manifest.xml` vào `%LOCALAPPDATA%\MathSymbolsAddin\`.
- Ghi một giá trị `REG_SZ` vào `HKCU\SOFTWARE\Microsoft\Office\16.0\WEF\Developer`: tên = `<Id>` trong manifest, dữ liệu = đường dẫn manifest. Đây chính là cơ chế mà công cụ sideload chính thức của Microsoft (`office-addin-dev-settings`) dùng.

Muốn triển khai hàng loạt không qua giao diện, dùng `installer\Register-Addin.ps1` (có tham số `-Uninstall`) hoặc chạy `MathSymbolsSetup.exe /SILENT`.

### Build lại bộ cài

Cần [Inno Setup 6](https://jrsoftware.org/isinfo.php) (`winget install --id JRSoftware.InnoSetup`):

```powershell
powershell -File installer\build.ps1
```

Kết quả: `dist\MathSymbolsSetup.exe`.

### Lựa chọn thay thế

Nếu tổ chức dùng Microsoft 365 Admin Center, có thể dùng **Centralized Deployment** (Settings → Integrated apps → Upload custom apps) để đẩy add-in tới toàn bộ người dùng mà không cần chạy bộ cài trên từng máy.

## Phát triển cục bộ

```bash
python -m http.server 8123
```

Mở `http://127.0.0.1:8123/taskpane/taskpane.html` để xem UI ở chế độ xem thử (click ký hiệu sẽ copy vào clipboard thay vì chèn vào Word).

Để test thật trong Word, sửa `SourceLocation` và các `bt:Url` trong manifest trỏ về `https://localhost:3000/...` kèm chứng chỉ dev HTTPS (`npx office-addin-dev-certs install`) — Word yêu cầu HTTPS.

## Thêm ký hiệu mới

Sửa `taskpane/symbols.js`:

- Ký hiệu Unicode: thêm cặp `['⊗', 'tensor product']` vào mảng của nhóm tương ứng.
- Mẫu công thức: thêm object `{ name, preview, type: 'omml', ooxml: wrapOMath('<m:...>') }`. Cách lấy XML mẫu nhanh: tạo công thức trong Word, lưu file .docx, giải nén và xem `word/document.xml`.
