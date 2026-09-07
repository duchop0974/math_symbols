; Bộ cài add-in "Trợ Lý Soạn Đề" cho Microsoft Word.
; Cài theo từng user (không cần quyền admin): copy manifest vào LocalAppData
; và đăng ký vào khoá WEF\Developer để Word tự nạp add-in khi khởi động.

#define AppName "Tro Ly Soan De"
#define AppNameVi "Trợ Lý Soạn Đề"
#define AppVersion "1.1.0"
#define AppPublisher "Trợ Lý Soạn Đề"
#define AppUrl "https://github.com/duchop0974/math_symbols"
#define AddinId "0f864aba-8b36-40d2-8d98-c7cd44356d53"

[Setup]
AppId={{9E4B1C7A-3F52-4D18-9A6E-5C0D2B8F71A3}
AppName={#AppNameVi}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppSupportURL={#AppUrl}
DefaultDirName={localappdata}\MathSymbolsAddin
DisableDirPage=yes
DisableProgramGroupPage=yes
UninstallDisplayName={#AppNameVi} (Word Add-in)
UninstallDisplayIcon={app}\assets\icon-128.png
OutputDir=..\dist
OutputBaseFilename=MathSymbolsSetup
Compression=lzma2
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "vi"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "..\manifest.xml"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\assets\icon-128.png"; DestDir: "{app}\assets"; Flags: ignoreversion

[Registry]
; Word đọc khoá này để nạp add-in sideload: tên giá trị = Id trong manifest.
Root: HKCU; Subkey: "SOFTWARE\Microsoft\Office\16.0\WEF\Developer"; \
    ValueType: string; ValueName: "{#AddinId}"; ValueData: "{app}\manifest.xml"; \
    Flags: uninsdeletevalue createvalueifdoesntexist

[Code]
function WordIsRunning(): Boolean;
var
  ResultCode: Integer;
begin
  Result := False;
  if Exec('cmd.exe', '/c tasklist /FI "IMAGENAME eq WINWORD.EXE" | find /I "WINWORD.EXE"',
          '', SW_HIDE, ewWaitUntilTerminated, ResultCode) then
    Result := (ResultCode = 0);
end;

function InitializeSetup(): Boolean;
begin
  Result := True;
  if WordIsRunning() then
  begin
    if MsgBox('Microsoft Word đang mở.' + #13#10#13#10 +
              'Hãy đóng hẳn Word trước khi cài để add-in được nạp đúng.' + #13#10 +
              'Bạn vẫn muốn tiếp tục cài?', mbConfirmation, MB_YESNO) = IDNO then
      Result := False;
  end;
end;

function InitializeUninstall(): Boolean;
begin
  Result := True;
  if WordIsRunning() then
  begin
    if MsgBox('Microsoft Word đang mở.' + #13#10#13#10 +
              'Hãy đóng Word trước khi gỡ cài đặt.' + #13#10 +
              'Bạn vẫn muốn tiếp tục gỡ?', mbConfirmation, MB_YESNO) = IDNO then
      Result := False;
  end;
end;

[Messages]
vi.FinishedLabel=Đã cài xong [name].%n%nMở Microsoft Word, vào tab Home và bấm nút "Soạn đề" để mở bảng công cụ.%n%nAdd-in tải giao diện từ Internet nên máy cần có kết nối mạng khi sử dụng.
