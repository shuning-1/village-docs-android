; 智慧文档管理系统 - Inno Setup 安装脚本
; 全中文界面，默认安装到 D:\智慧文档管理系统
; 自动检测并安装 WebView2 运行时

[Setup]
; 应用信息
AppName=智慧文档管理系统
AppVersion=1.0.0
AppPublisher=智慧文档管理系统
AppPublisherURL=https://www.example.com
AppSupportURL=https://www.example.com
AppUpdatesURL=https://www.example.com

; 默认安装目录
DefaultDirName=D:\智慧文档管理系统
DefaultGroupName=智慧文档管理系统

; 升级安装时记住上次的安装目录
UsePreviousAppDir=yes
; 安装前强制关闭正在运行的程序，确保 exe 能被替换
CloseApplications=force

; 输出配置
OutputDir=installer_output
OutputBaseFilename=智慧文档管理系统安装程序
Compression=lzma2/ultra64
SolidCompression=yes
LZMAUseSeparateProcess=yes

; 界面设置
DisableProgramGroupPage=yes
DisableDirPage=no
PrivilegesRequired=admin
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

; 图标
SetupIconFile=app_icon.ico
UninstallDisplayIcon={app}\智慧文档管理系统.exe

; 卸载
UninstallDisplayName=智慧文档管理系统

; 版本信息
VersionInfoVersion=1.0.0.0
VersionInfoCompany=智慧文档管理系统
VersionInfoProductName=智慧文档管理系统
VersionInfoProductVersion=1.0.0.0

[Languages]
Name: "chinesesimp"; MessagesFile: "ChineseSimplified.isl"

[Tasks]
Name: "desktopicon"; Description: "在桌面上创建快捷方式"; GroupDescription: "附加任务:"; Flags: checkedonce

[Files]
; 主程序
Source: "dist\智慧文档管理系统.exe"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
; 开始菜单快捷方式
Name: "{group}\智慧文档管理系统"; Filename: "{app}\智慧文档管理系统.exe"; IconFilename: "{app}\智慧文档管理系统.exe"; Comment: "智慧文档管理系统"
Name: "{group}\卸载智慧文档管理系统"; Filename: "{uninstallexe}"; Comment: "卸载智慧文档管理系统"

; 桌面快捷方式
Name: "{commondesktop}\智慧文档管理系统"; Filename: "{app}\智慧文档管理系统.exe"; IconFilename: "{app}\智慧文档管理系统.exe"; Comment: "智慧文档管理系统"; Tasks: desktopicon

[Run]
; 安装完成后启动程序
Filename: "{app}\智慧文档管理系统.exe"; Description: "立即启动智慧文档管理系统"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: files; Name: "{app}\launcher.log"
Type: files; Name: "{app}\console.log"
Type: files; Name: "{app}\crash.log"
Type: files; Name: "{app}\fatal_error.log"

[Code]
function InitializeSetup(): Boolean;
begin
  Result := True;
end;

// 检查 WebView2 运行时是否已安装
function WebView2Installed: Boolean;
var
  Version: String;
begin
  Result := False;
  if RegQueryStringValue(HKLM, 'SOFTWARE\WOW6432Node\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'pv', Version) then
  begin
    Result := (Version <> '');
    if Result then Exit;
  end;
  if RegQueryStringValue(HKLM, 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'pv', Version) then
  begin
    Result := (Version <> '');
    if Result then Exit;
  end;
  if RegQueryStringValue(HKCU, 'SOFTWARE\Microsoft\EdgeUpdate\Clients\{F3017226-FE2A-4295-8BDF-00C3A9A7E4C5}', 'pv', Version) then
  begin
    Result := (Version <> '');
    if Result then Exit;
  end;
end;

// 安装完成后自动检测并安装 WebView2
procedure CurStepChanged(CurStep: TSetupStep);
var
  ResultCode: Integer;
  TempDir: String;
  ScriptPath: String;
  ScriptContent: String;
begin
  if CurStep = ssPostInstall then
  begin
    // 如果 WebView2 已安装，跳过
    if WebView2Installed then
      Exit;

    // 提示用户需要安装 WebView2
    if MsgBox('系统检测到未安装 WebView2 运行时组件，程序需要此组件才能正常运行。' + #13#10 + #13#10 +
      '是否现在自动下载并安装？（需要网络连接，约 100-200 MB）',
      mbConfirmation, MB_YESNO) = IDNO then
    begin
      MsgBox('未安装 WebView2 运行时组件，程序可能无法正常启动。' + #13#10 + #13#10 +
        '您可稍后手动下载安装：' + #13#10 +
        'https://developer.microsoft.com/en-us/microsoft-edge/webview2/',
        mbInformation, MB_OK);
      Exit;
    end;

    // 生成 PowerShell 脚本用于下载安装 WebView2
    TempDir := ExpandConstant('{tmp}');
    ScriptPath := TempDir + '\install_webview2.ps1';
    ScriptContent :=
      '$ProgressPreference = ''Continue''' + #13#10 +
      'Write-Host "正在下载 WebView2 运行时组件..."' + #13#10 +
      '$url = ''https://go.microsoft.com/fwlink/p/?LinkId=2124701''' + #13#10 +
      '$out = Join-Path $env:TEMP ''MicrosoftEdgeWebview2Setup.exe''' + #13#10 +
      'try {' + #13#10 +
      '  $wc = New-Object System.Net.WebClient' + #13#10 +
      '  $wc.DownloadFile($url, $out)' + #13#10 +
      '  Write-Host "下载完成，正在安装..."' + #13#10 +
      '  $p = Start-Process -FilePath $out -ArgumentList ''/silent /install'' -Wait -PassThru' + #13#10 +
      '  Write-Host "安装完成 (代码: $($p.ExitCode))"' + #13#10 +
      '  Remove-Item $out -Force -ErrorAction SilentlyContinue' + #13#10 +
      '  if ($p.ExitCode -eq 0) { exit 0 } else { exit $p.ExitCode }' + #13#10 +
      '} catch {' + #13#10 +
      '  Write-Host "下载失败: $_"' + #13#10 +
      '  exit 1' + #13#10 +
      '}';

    // 写入脚本文件
    SaveStringToFile(ScriptPath, ScriptContent, False);

    // 执行 PowerShell 脚本
    if Exec('powershell.exe',
         '-ExecutionPolicy Bypass -NoProfile -File "' + ScriptPath + '"',
         '', SW_SHOW, ewWaitUntilTerminated, ResultCode) then
    begin
      if ResultCode = 0 then
        Log('WebView2 安装成功')
      else
        Log('WebView2 安装返回代码: ' + IntToStr(ResultCode));
    end
    else
      Log('WebView2 安装脚本执行失败');

    // 清理脚本文件
    DeleteFile(ScriptPath);
  end;
end;
