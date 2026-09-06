# Ký Hiệu Toán Học — Word Add-in

Task pane add-in cho Microsoft Word (Windows) giúp chèn nhanh ký hiệu toán học Unicode và công thức Word Equation (OMML).

## Tính năng

- 6 nhóm ký hiệu Unicode: Hy Lạp, Toán tử, Tập hợp, Logic, Mũi tên, Giải tích — click để chèn vào vị trí con trỏ.
- 9 mẫu công thức OMML (phân số, căn bậc hai/bậc n, lũy thừa, chỉ số dưới, tích phân có cận, tổng Σ, giới hạn, ma trận 2x2) — chèn thành Equation Object thật của Word, edit tiếp được bằng Equation Editor.
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
  symbols.js          # Dữ liệu ký hiệu + mẫu OMML
```

Không có build step — toàn bộ là static file, host thẳng lên GitHub Pages.

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

Người dùng cuối chỉ cần chạy `MathSymbolsSetup.exe`:

1. Đóng Microsoft Word.
2. Chạy `MathSymbolsSetup.exe` (cài theo từng user, **không cần quyền admin**).
3. Mở Word → tab **Home** → nút **Add-ins** → mục **Developer Add-ins** → bấm **Ký Hiệu Toán Học**.
4. Từ lần này trở đi, nút **Bảng Ký Hiệu** nằm sẵn ở tab Home.

Bước 3 chỉ phải làm một lần trên mỗi máy: Word nạp manifest ngay khi khởi động, nhưng chỉ gắn nút lên ribbon sau lần mở đầu tiên.

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
