# Trợ Lý Soạn Đề — Word Add-in

Task pane add-in cho Microsoft Word (Windows) giúp giáo viên **soạn đề thi nhanh hơn**: dựng sẵn khung đề, bảng biến thiên, bảng đáp án và các thao tác đánh số/trộn đề. Bảng ký hiệu Unicode và công thức Word Equation (OMML) là phần phụ đi kèm — không nhằm thay thế trình soạn công thức của Word.

## Tính năng

Task pane mở ra ở tab **Đề thi**. Bốn tab trên một hàng: *Đề thi · Bảng BT · Đáp án · Ký hiệu* (10 nhóm ký hiệu gom vào ô chọn nhóm trong tab Ký hiệu).

Trên cùng là **thanh nút nhanh** dính theo màn hình, luôn thấy dù đang ở mục nào:

- ô số câu + kiểu câu + nút **+ Chèn** — chèn câu tiếp theo rồi tự tăng số, đồng bộ với ô "Bắt đầu từ câu" của tab Đề thi;
- hàng ký hiệu **vừa dùng và đánh dấu ★** — chèn lại công thức mà không phải rời tab đang làm.

Tab **Đề thi** xếp theo việc chính trước, việc lẻ sau:

1. **Dựng cả khung đề** (mở sẵn) — mọi thứ để ra một đề hoàn chỉnh trong một lần bấm.
2. **Chèn từng khối** — đầu đề, dòng kết, thêm câu lẻ khi đề đã dựng xong.
3. **Bảng hướng dẫn chấm**.
4. **Định dạng chuẩn** — phông, cỡ chữ, lề trang, giãn dòng; đặt một lần rồi thôi.
5. **Nếu đề phải nộp bằng MathType**.

Mục **Dựng cả khung đề** chứa cả kiểu đề, thông tin đầu đề và một danh sách phần tự thêm/bớt — không cố định ba phần:

| | |
|---|---|
| Số phần | thêm/bớt tuỳ ý, đánh số La Mã tự cập nhật |
| Mỗi phần | tự chọn dạng câu (trắc nghiệm A–D / đúng-sai / trả lời ngắn / tự luận), số câu, số cột phương án hoặc điểm và số ý |
| Tiêu đề phần | để trống thì tự sinh theo câu chữ đề Bộ GD&ĐT từ 2025, gõ vào để thay bằng chữ của mình; tắt hẳn bằng ô *Có tiêu đề phần* |
| Đánh số | lại từ 1 mỗi phần (chuẩn đề Bộ) hoặc liên tục cả đề |
| Đầu đề, dòng kết | bật/tắt riêng |
| Nút ⤓ | chèn riêng một phần |

Dòng tóm tắt cập nhật ngay theo từng ô — *"Sẽ dựng: đầu đề + Phần I 12 câu trắc nghiệm + Phần II 4 câu đúng/sai + Phần III 6 câu trả lời ngắn + dòng kết đề — tổng 22 câu, đánh số lại từ 1 mỗi phần."* — nên biết trước sẽ ra gì rồi mới bấm **Dựng vào Word**. Cấu hình được nhớ riêng cho từng kiểu đề.

- 6 nhóm ký hiệu Unicode: Hy Lạp, Toán tử, Tập hợp, Logic, Mũi tên, Giải tích — click để chèn vào vị trí con trỏ.
- 65 mẫu công thức OMML chia 5 nhóm — chèn thành Equation Object thật của Word với **ô trống điền được**, Tab để nhảy giữa các ô, lồng được vào nhau:
  - *Phân số & mũ*: 4 kiểu phân số, căn bậc hai/bậc n, lũy thừa, chỉ số dưới/trên/trước, mũ của ngoặc.
  - *Tổng & tích phân*: Σ, Π, ∫ (có cận/không cận/hai lớp/đường), ⋃, ⋂.
  - *Ngoặc & ma trận*: 6 loại ngoặc, tổ hợp, hệ 2–3 phương trình, ma trận 2x2/3x3, định thức, vector cột.
  - *Hàm & dấu*: lim, max, sin/cos/tan/ln, log cơ số, đạo hàm thường và riêng, vector, hat, chấm trên, gạch trên/dưới, ngoặc nhọn trên/dưới.
  - *Hàm số*: f(▫), g(▫), u(▫), f′(▫), f″(▫), y = f(▫), hàm hợp f(g(▫)), hàm ngược f⁻¹(▫), nguyên hàm ∫f(x)dx, tích phân có cận, F(x) + C, vi phân dx/dt (chữ d đứng đúng chuẩn), tập xác định D = ℝ∖{▫}.
- **Ba tab công cụ soạn đề thi**:
  - *Đề thi*: mục **Định dạng chuẩn** đặt phông (mặc định Times New Roman 12), cỡ chữ, **lề trang** và **giãn dòng** theo Nghị định 30/2020/NĐ-CP (A4, lề trái 30–35mm, phải 15–20mm, trên/dưới 20–25mm). Lề trang dùng để tính bề rộng bảng cho khớp cột chữ — add-in không đổi được lề của tài liệu. Sau đó chọn **kiểu đề**:
    - *Tự luận / học sinh giỏi*: đầu đề Phòng GD&ĐT + ĐỀ CHÍNH THỨC / kỳ thi / lớp + năm học / MÔN / thời gian "(không kể thời gian giao đề)" / số trang; câu dạng `Câu 1 (4,0 điểm).` kèm các ý a) b) c); dòng kết `…Hết…` với lời dặn và chỗ ghi họ tên; bảng hướng dẫn chấm `Câu | Ý | Nội dung | Điểm` và bảng cấu trúc đề `Câu | Nội dung | Điểm`.
    - *Trắc nghiệm*: đầu đề theo mẫu đề Bộ GD&ĐT (`Môn thi: TOÁN`, `Thời gian làm bài: 90 phút, không kể thời gian phát đề`, `Mã đề thi 101`, `(Đề thi có 04 trang)`, họ tên và số báo danh trên hai dòng); tiêu đề Phần I/II/III lấy đúng câu chữ đề tham khảo từ 2025, chọn được bản đầy đủ hoặc bản ngắn, số câu đặt được (mặc định Toán 12 – 4 – 6) — mỗi phần đánh số lại từ câu 1 nên bấm tiêu đề phần sẽ đưa ô "Bắt đầu từ câu" về 1; chèn hàng loạt câu trắc nghiệm A–D (1/2/4 cột), câu đúng/sai a–d, câu trả lời ngắn.

    Cả hai kiểu đều nhớ lại nội dung đã nhập cho lần mở sau, và ô "Bắt đầu từ câu" tự tăng sau mỗi lần chèn.
  - *Bảng BT*: dựng bảng biến thiên (3 dòng) hoặc bảng xét dấu (2 dòng) — nhập nghiệm, chọn dấu từng khoảng, xem trước ngay trong task pane rồi chèn ra thành bảng Word thật; mũi tên ↗ ↘ tự suy ra từ dấu đạo hàm.
  - *Đáp án*: chèn bảng đáp án từ chuỗi `1A 2B 3C` hoặc `ABCD...` (cả đáp án đúng/sai kiểu `1 ĐSSĐ`), bảng đáp án trống, đánh số lại toàn bộ câu hỏi, và **trộn đề thành nhiều mã** (xem mục dưới).
- Nếu đơn vị bắt buộc nộp đề gõ bằng **MathType**: soạn cả đề bằng add-in rồi chuyển một lần ở bước cuối — Word → tab MathType → **Convert Equations** → nguồn *Word 2007 and later (OMML) equations*, phạm vi *Whole document*, đích *MathType equations (OLE objects)*. Hướng dẫn này có sẵn trong tab Đề thi.
- Tìm kiếm theo tên tiếng Việt hoặc tiếng Anh (`alpha`, `integral`, `phân số`...).
- Đánh dấu yêu thích (★) và mẫu vừa dùng — hiện trên thanh nút nhanh, lưu trong localStorage của task pane.

## Cấu trúc

```
manifest.xml          # Manifest add-in
set-host.ps1          # Script điền URL GitHub Pages vào manifest
installer/
  TroLySoanDe.iss     # Script Inno Setup tạo TroLySoanDeSetup.exe
  build.ps1           # Biên dịch bộ cài
  Register-Addin.ps1  # Đăng ký/gỡ add-in bằng PowerShell (triển khai hàng loạt)
assets/               # Icon add-in (16/32/64/80/128 px)
taskpane/
  taskpane.html       # Giao diện task pane
  taskpane.css        # Style (hỗ trợ dark mode)
  taskpane.js         # Logic render + chèn vào Word qua Office.js
  symbols.js          # Dữ liệu ký hiệu + mẫu OMML + hàm đóng gói Flat OPC
  exam.js             # Dựng OOXML cho đầu đề, câu hỏi, bảng biến thiên, bảng đáp án
  shuffle.js          # Trộn đề nhiều mã: tách/ghép gói OPC, đảo câu và phương án
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

Người dùng cuối chỉ cần chạy `TroLySoanDeSetup.exe`:

1. Đóng Microsoft Word.
2. Chạy `TroLySoanDeSetup.exe` (cài theo từng user, **không cần quyền admin**).
3. Mở Word → tab **Home** → nút **Add-ins** → mục **Developer Add-ins** → bấm **Trợ Lý Soạn Đề**.
4. Từ lần này trở đi, nút **Soạn đề** nằm sẵn ở tab Home.

Bước 3 chỉ phải làm một lần trên mỗi máy: Word nạp manifest ngay khi khởi động, nhưng chỉ gắn nút lên ribbon sau lần mở đầu tiên.

## Cập nhật

Giao diện task pane nạp trực tiếp từ GitHub Pages, nên **hầu hết thay đổi chỉ cần push code** — máy giáo viên có ngay, không phải cài lại. Hai ngoại lệ:

- **Sửa `taskpane/taskpane.html`** (thêm file JS mới, đổi tiêu đề…): Word giữ bản HTML cũ trong cache. Đóng Word rồi xoá sạch bên trong `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef`, mở lại Word. Các file `.js`/`.css` không bị, vì đã có tham số `?v=` tăng theo mỗi bản.
- **Sửa `manifest.xml`** (tên add-in, nhãn nút ribbon…): phải phát hành bộ cài mới và chạy lại trên từng máy. `AppId` của bộ cài và `Id` của add-in giữ nguyên nên bản mới **nâng cấp đè** lên bản cũ, không tạo mục thứ hai trong Apps.

## Trộn đề nhiều mã

Tab **Đáp án → Trộn đề nhiều mã**: nhập danh sách mã đề (`101, 102, 103, 104`) và đáp án đề gốc, add-in đọc đề đang mở rồi **nối thêm từng mã đề vào cuối tài liệu**, mỗi mã bắt đầu ở một trang mới. Đề gốc được giữ nguyên.

Với mỗi mã đề:

- thứ tự câu bị đảo **trong từng phần** — câu trắc nghiệm không bị trộn lẫn sang phần đúng/sai;
- thứ tự **A/B/C/D trong từng câu** cũng bị đảo, nhãn được viết lại đúng vị trí mới, công thức trong phương án giữ nguyên;
- mỗi phần đánh số lại từ câu 1, đúng cấu trúc đề Bộ;
- dòng `Mã đề thi ...` ở đầu đề được thay bằng mã tương ứng.

Nhập đáp án đề gốc thì add-in chèn thêm **bảng đáp án tổng hợp** — mỗi dòng một mã đề, mỗi cột một câu — đã tính lại theo cả thứ tự câu lẫn thứ tự phương án. Task pane cũng hiện bảng *câu mới ← câu gốc* của từng mã.

Nhận diện dựa trên chữ đầu đoạn: `PHẦN I/II/III` cho ranh giới phần, `Câu 1.` / `Câu 1:` / `Câu 1)` cho đầu mỗi câu. Câu không có đúng 4 phương án (đúng/sai, trả lời ngắn, tự luận) chỉ bị đảo thứ tự, không đụng vào ruột.

Hãy **lưu tài liệu trước khi trộn**.

Về mặt kỹ thuật, `getOoxml()` trả về nguyên một gói Flat OPC nên không nối thẳng được; `shuffle.js` bóc ruột `<w:body>` của từng khối rồi ghép lại vào chính gói Word trả về, nhờ vậy `styles.xml` và các part khác được giữ nguyên.

Gỡ cài đặt: **Settings → Apps → Trợ Lý Soạn Đề → Uninstall** (bộ cài tự xoá khoá registry đã ghi).

Add-in tải giao diện từ GitHub Pages nên **máy cần có Internet** khi dùng.

### Cách bộ cài hoạt động

- Copy `manifest.xml` vào `%LOCALAPPDATA%\TroLySoanDe\`.
- Ghi một giá trị `REG_SZ` vào `HKCU\SOFTWARE\Microsoft\Office\16.0\WEF\Developer`: tên = `<Id>` trong manifest, dữ liệu = đường dẫn manifest. Đây chính là cơ chế mà công cụ sideload chính thức của Microsoft (`office-addin-dev-settings`) dùng.

Muốn triển khai hàng loạt không qua giao diện, dùng `installer\Register-Addin.ps1` (có tham số `-Uninstall`) hoặc chạy `TroLySoanDeSetup.exe /SILENT`.

### Build lại bộ cài

Cần [Inno Setup 6](https://jrsoftware.org/isinfo.php) (`winget install --id JRSoftware.InnoSetup`):

```powershell
powershell -ExecutionPolicy Bypass -File installer\build.ps1
```

Cần `-ExecutionPolicy Bypass` vì Windows mặc định chặn chạy file `.ps1`; cờ này chỉ áp cho đúng tiến trình đó, không đổi thiết lập máy. Thiếu nó sẽ báo *"running scripts is disabled on this system"* trước khi script kịp chạy. Cách khác là gọi thẳng trình biên dịch:

```powershell
& "$env:LOCALAPPDATA\Programs\Inno Setup 6\ISCC.exe" installer\TroLySoanDe.iss
```

Kết quả: `dist\TroLySoanDeSetup.exe`. Lưu ý tương tự khi chạy `set-host.ps1` và `Register-Addin.ps1`.

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
