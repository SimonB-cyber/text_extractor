; ============================================================
; Text-Extraktor Pro - Inno Setup Installer Script
; ============================================================
#define AppName "Text-Extraktor Pro"
#define AppVersion "1.2"
#define AppPublisher "Simon B."
#define AppExeName "TextExtraktPro.exe"
#define TessSourceDir "C:\Program Files\Tesseract-OCR"

[Setup]
AppId={{A1B2C3D4-E5F6-7890-ABCD-EF1234567890}
AppName={#AppName}
AppVersion={#AppVersion}
AppPublisherURL=
AppPublisher={#AppPublisher}
DefaultDirName={autopf}\TextExtraktPro
DefaultGroupName={#AppName}
AllowNoIcons=no
Compression=lzma2/ultra64
SolidCompression=yes
WizardStyle=modern
OutputDir=C:\Users\Simon B\Downloads\installer_output
OutputBaseFilename=TextExtraktPro_Setup_v1.2
; Show nice welcome page
DisableWelcomePage=no
; Run as admin so we can install to Program Files
PrivilegesRequired=admin

[Languages]
Name: "german"; MessagesFile: "compiler:Languages\German.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Dirs]
; Create Tesseract directory in Program Files
Name: "{autopf}\Tesseract-OCR"

[Files]
; ------ Main Application ------
Source: "C:\Users\Simon B\Downloads\dist\TextExtraktPro\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

; ------ Tesseract OCR (deine eigene, fertige Kopie!) ------
Source: "C:\Program Files\Tesseract-OCR\*"; DestDir: "{autopf}\Tesseract-OCR"; Flags: ignoreversion recursesubdirs createallsubdirs

; ------ tesseract_path.txt damit das Programm es direkt findet ------
Source: "C:\Users\Simon B\Downloads\tesseract_path.txt"; DestDir: "{userappdata}\TextExtraktPro"; Flags: ignoreversion skipifsourcedoesntexist

[Icons]
Name: "{group}\{#AppName}"; Filename: "{app}\{#AppExeName}"
Name: "{group}\{cm:UninstallProgram,{#AppName}}"; Filename: "{uninstallexe}"
Name: "{commondesktop}\{#AppName}"; Filename: "{app}\{#AppExeName}"; Tasks: desktopicon

[Registry]
; Speichere den Tesseract-Pfad in der Windows Registry als Backup
Root: HKLM; Subkey: "SOFTWARE\Tesseract-OCR"; ValueType: string; ValueName: "Path"; ValueData: "{autopf}\Tesseract-OCR\tesseract.exe"; Flags: createvalueifdoesntexist

[Run]
Filename: "{app}\{#AppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(AppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent

[Code]
// Schreibe die tesseract_path.txt mit dem echten Pfad nach der Installation in APPDATA
procedure WriteConfigFile();
var
  TessPath: String;
  ConfigLines: TArrayOfString;
  AppDataDir: String;
begin
  TessPath := ExpandConstant('{autopf}\Tesseract-OCR\tesseract.exe');
  AppDataDir := ExpandConstant('{userappdata}\TextExtraktPro');
  
  // Ordner erstellen falls er nicht existiert
  ForceDirectories(AppDataDir);
  
  SetArrayLength(ConfigLines, 1);
  ConfigLines[0] := TessPath;
  SaveStringsToFile(AppDataDir + '\tesseract_path.txt', ConfigLines, False);
end;

procedure CurStepChanged(CurStep: TSetupStep);
begin
  if CurStep = ssPostInstall then
    WriteConfigFile();
end;
