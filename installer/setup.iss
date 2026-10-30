; Inno Setup script for Tacitus Lite for Outlook.
; Build on Windows after PyInstaller:  ISCC.exe installer\setup.iss
; Installs per user (no administrator), under %LOCALAPPDATA%\Programs.

#define AppName "Tacitus Lite for Outlook"
#define AppVersion GetEnv("TACITUS_LITE_VERSION")
#if AppVersion == ""
  #define AppVersion "0.0.0"
#endif
#define TaskName "Tacitus Lite mail dump"

[Setup]
AppId={{7B1F3C2E-5A2D-4E1B-9C77-2D6B3F0A1E55}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisher=ARPG, University of Colorado Boulder
AppPublisherURL=https://github.com/arpg/tacitus-lite-outlook
DefaultDirName={localappdata}\Programs\TacitusLite
DisableDirPage=yes
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
OutputDir=Output
OutputBaseFilename=TacitusLiteOutlook-Setup
Compression=lzma2
SolidCompression=yes
ChangesEnvironment=yes
UninstallDisplayName={#AppName}
WizardStyle=modern

[Files]
Source: "..\dist\TacitusLite\*"; DestDir: "{app}"; Flags: recursesubdirs ignoreversion
; The mail folder gets the agent's instructions and a notes README once;
; a reinstall never overwrites what the person changed.
Source: "..\AGENTS.md"; DestDir: "{userdocs}\OutlookMail"; Flags: onlyifdoesntexist uninsneveruninstall
Source: "..\notes\README.md"; DestDir: "{userdocs}\OutlookMail\notes"; Flags: onlyifdoesntexist uninsneveruninstall

[Dirs]
Name: "{userdocs}\OutlookMail\mail"; Flags: uninsneveruninstall
Name: "{userdocs}\OutlookMail\drafts"; Flags: uninsneveruninstall

[Icons]
Name: "{userprograms}\Tacitus Lite\Self-test (Outlook must be open)"; Filename: "{app}\tacitus-lite.exe"; Parameters: "selftest --draft --pause"; WorkingDir: "{userdocs}\OutlookMail"
Name: "{userprograms}\Tacitus Lite\Copy all mail now"; Filename: "{app}\tacitus-lite.exe"; Parameters: "dump --full --pause"; WorkingDir: "{userdocs}\OutlookMail"
Name: "{userprograms}\Tacitus Lite\Copy new mail now"; Filename: "{app}\tacitus-lite.exe"; Parameters: "dump --pause"; WorkingDir: "{userdocs}\OutlookMail"
Name: "{userprograms}\Tacitus Lite\Status"; Filename: "{app}\tacitus-lite.exe"; Parameters: "status --pause"; WorkingDir: "{userdocs}\OutlookMail"
Name: "{userprograms}\Tacitus Lite\Mail folder"; Filename: "{userdocs}\OutlookMail"
Name: "{userprograms}\Tacitus Lite\Ask Codex"; Filename: "{cmd}"; Parameters: "/k cd /d ""{userdocs}\OutlookMail"" && codex || echo Codex is not installed: see the README."; WorkingDir: "{userdocs}\OutlookMail"
Name: "{userprograms}\Tacitus Lite\Uninstall"; Filename: "{uninstallexe}"

[Registry]
; tacitus-lite on the user's PATH, so `tacitus-lite draft ...` works in any terminal and for the agent.
Root: HKCU; Subkey: "Environment"; ValueType: expandsz; ValueName: "Path"; ValueData: "{olddata};{app}"; Check: NeedsAddPath(ExpandConstant('{app}'))

[Run]
; The scheduled task: every 30 minutes while logged in, windowless, logging to mail\dump.log.
Filename: "schtasks.exe"; Parameters: "/Create /F /SC MINUTE /MO 30 /TN ""{#TaskName}"" /TR ""\""{app}\tacitus-lite-bg.exe\"" dump"""; Flags: runhidden
Filename: "{app}\tacitus-lite.exe"; Parameters: "selftest --draft --pause"; WorkingDir: "{userdocs}\OutlookMail"; Description: "Run the self-test now (Outlook must be open)"; Flags: postinstall nowait skipifsilent

[UninstallRun]
Filename: "schtasks.exe"; Parameters: "/Delete /F /TN ""{#TaskName}"""; Flags: runhidden; RunOnceId: "deltask"

[Code]
function NeedsAddPath(Param: string): boolean;
var
  OrigPath: string;
begin
  if not RegQueryStringValue(HKCU, 'Environment', 'Path', OrigPath) then
  begin
    Result := True;
    exit;
  end;
  Result := Pos(';' + Uppercase(Param) + ';', ';' + Uppercase(OrigPath) + ';') = 0;
end;

procedure CurUninstallStepChanged(CurUninstallStep: TUninstallStep);
var
  OrigPath, App: string;
  P: Integer;
begin
  if CurUninstallStep = usPostUninstall then
  begin
    App := ExpandConstant('{app}');
    if RegQueryStringValue(HKCU, 'Environment', 'Path', OrigPath) then
    begin
      P := Pos(';' + App, OrigPath);
      if P > 0 then
      begin
        Delete(OrigPath, P, Length(App) + 1);
        RegWriteExpandStringValue(HKCU, 'Environment', 'Path', OrigPath);
      end;
    end;
  end;
end;
