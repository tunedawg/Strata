; ============================================================
;  Universal Search — Inno Setup 6 installer script
;  Produces: dist\UniversalSearch_Setup.exe
;
;  Run after build.bat, or compile manually:
;    "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer.iss
; ============================================================

#define AppName      "Universal Search"
#define AppVersion   "1.0"
#define AppPublisher "Your Firm Name"
#define AppURL       ""
#define AppExeName   "UniversalSearch.exe"
#define SourceDir    "dist\UniversalSearch"

[Setup]
AppId={{A7F2C3D1-4E8B-4F9A-B2D3-6C7E8F9A0B1C}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher={#AppPublisher}
AppPublisherURL={#AppURL}
AppSupportURL={#AppURL}
AppUpdatesURL={#AppURL}
DefaultDirName={autopf}\{#AppName}
DefaultGroupName={#AppName}
AllowNoIcons=yes
OutputDir=dist
OutputBaseFilename=UniversalSearch_Setup
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=commandline
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Tasks]
Name: "desktopicon"; Description: "Create a &desktop shortcut"; GroupDescription: "Additional icons:"
Name: "startmenuicon"; Description: "Create a Start &Menu shortcut"; GroupDescription: "Additional icons:"

[Files]
Source: "{#SourceDir}\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{userdesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon; Comment: "Open Universal Search"
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: startmenuicon
Name: "{group}\Uninstall {#AppName}"; Filename: "{uninstallexe}"

[Run]
Filename: "{app}\{#AppExeName}"; Description: "Launch {#AppName} now"; Flags: nowait postinstall skipifsilent

[UninstallDelete]
Type: filesandordirs; Name: "{app}\__pycache__"

[Code]
procedure CurStepChanged(CurStep: TSetupStep);
var
  DataPath: String;
begin
  if CurStep = ssDone then
  begin
    DataPath := ExpandConstant('{userdocs}') + '\UniversalSearch';
    MsgBox(
      'Installation complete!' + #13#10 + #13#10 +
      'Your datasets and exports will be saved to:' + #13#10 +
      '  ' + DataPath + #13#10 + #13#10 +
      'This folder is preserved if you update or reinstall the app.',
      mbInformation, MB_OK
    );
  end;
end;
