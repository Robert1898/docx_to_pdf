; Installer Inno Setup 6. Se construieste prin build.ps1, care trimite numele si versiunea
; din convertor/__init__.py (/DAppVersion, /DAppName, /DAppExe, /DCliExe, /DAppDataId).
; Implicit se instaleaza doar pentru utilizatorul curent, fara drepturi de administrator.

#ifndef AppName
  #define AppName "Document Converter"
#endif
#ifndef AppExe
  #define AppExe "Document Converter.exe"
#endif
#ifndef CliExe
  #define CliExe "docconv.exe"
#endif
#ifndef AppDataId
  #define AppDataId "DocumentConverter"
#endif
#ifndef AppVersion
  #define AppVersion "0.0.0"
#endif

[Setup]
; acelasi AppId ca „Convertor DOCX-PDF” (pana la 2.0): Windows o vede ca actualizare
AppId={{8F1C2E3A-5B7D-4C9E-A1F0-2D6B9E4C7A13}
AppName={#AppName}
AppVersion={#AppVersion}
AppVerName={#AppName} {#AppVersion}
AppPublisher=Robert-Raul Mîrza
VersionInfoVersion={#AppVersion}
SourceDir=..
OutputDir=dist\installer
OutputBaseFilename={#StringChange(AppName, " ", "-")}-Setup-{#AppVersion}
DefaultDirName={autopf}\{#AppName}
; dupa redenumire nu refolosim folderul vechi („Convertor DOCX-PDF”)
UsePreviousAppDir=no
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
SetupIconFile=icon.ico
UninstallDisplayIcon={app}\{#AppExe}
UninstallDisplayName={#AppName}
LicenseFile=LICENSE
Compression=lzma2/max
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible
CloseApplications=yes
MinVersion=10.0

[Languages]
Name: "ro"; MessagesFile: "installer\Romanian.isl"
Name: "en"; MessagesFile: "compiler:Default.isl"

[CustomMessages]
ro.Integration=Integrare cu Windows Explorer:
en.Integration=Windows Explorer integration:
ro.ContextMenu=„Convertește în PDF / DOCX” în meniul click-dreapta
en.ContextMenu="Convert to PDF / DOCX" in the right-click menu
ro.SendTo=Aplicația în „Trimite către” (pentru mai multe fișiere deodată)
en.SendTo=The app in "Send to" (for several files at once)

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked
Name: "contextmenu"; Description: "{cm:ContextMenu}"; GroupDescription: "{cm:Integration}"
Name: "sendto"; Description: "{cm:SendTo}"; GroupDescription: "{cm:Integration}"

[Files]
Source: "dist\{#AppName}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{autoprograms}\{#AppName}"; Filename: "{app}\{#AppExe}"
Name: "{autodesktop}\{#AppName}"; Filename: "{app}\{#AppExe}"; Tasks: desktopicon

[Run]
; integrarea cu Explorer e facuta de aplicatie, ca sa existe o singura implementare
Filename: "{app}\{#CliExe}"; Parameters: "--register-shell"; Flags: runhidden; Tasks: contextmenu
Filename: "{app}\{#CliExe}"; Parameters: "--register-sendto"; Flags: runhidden; Tasks: sendto
Filename: "{app}\{#AppExe}"; Description: "{cm:LaunchProgram,{#AppName}}"; Flags: nowait postinstall skipifsilent

[UninstallRun]
Filename: "{app}\{#CliExe}"; Parameters: "--unregister-shell"; Flags: runhidden; RunOnceId: "UnregisterShell"

[UninstallDelete]
; jurnalele locale; setarile din %APPDATA% raman pentru o eventuala reinstalare
Type: filesandordirs; Name: "{localappdata}\{#AppDataId}\logs"

[Code]
{ O versiune instalata anterior (inclusiv sub numele vechi) e dezinstalata silentios inainte,
  ca sa nu ramana exe-uri si scurtaturi cu numele vechi. }
function PreviousUninstaller(): String;
var
  Key, Value: String;
begin
  Result := '';
  Key := ExpandConstant('Software\Microsoft\Windows\CurrentVersion\Uninstall\{#SetupSetting("AppId")}_is1');
  if RegQueryStringValue(HKCU, Key, 'UninstallString', Value) or
     RegQueryStringValue(HKLM, Key, 'UninstallString', Value) then
    Result := RemoveQuotes(Value);
end;

function PrepareToInstall(var NeedsRestart: Boolean): String;
var
  Uninstaller: String;
  Code, Waited: Integer;
begin
  Result := '';
  Uninstaller := PreviousUninstaller();
  if (Uninstaller <> '') and FileExists(Uninstaller) then
  begin
    Exec(Uninstaller, '/VERYSILENT /SUPPRESSMSGBOXES /NORESTART', '', SW_HIDE,
         ewWaitUntilTerminated, Code);
    { dezinstalatorul se copiaza in Temp si continua acolo; asteptam sa termine }
    Waited := 0;
    while FileExists(Uninstaller) and (Waited < 60) do
    begin
      Sleep(500);
      Waited := Waited + 1;
    end;
  end;
end;
