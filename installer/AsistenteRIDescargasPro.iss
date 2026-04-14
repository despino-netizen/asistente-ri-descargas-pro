#define MyAppName "Asistente RI Descargas Pro"
#define MyAppVersion "3.4.2"
#define MyAppPublisher "Dariel Espino"
#define MyAppExeName "AsistenteRIDescargasPro.exe"
#define MyAppFolder "AsistenteRIDescargasPro"

[Setup]
AppId={{9A9BF494-2E33-4E29-B89E-B4AECE062A42}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={localappdata}\Programs\{#MyAppFolder}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
PrivilegesRequired=lowest
PrivilegesRequiredOverridesAllowed=dialog
OutputDir=..\dist
OutputBaseFilename=Instalar_AsistenteRIDescargasPro_v{#MyAppVersion}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
UninstallDisplayIcon={app}\{#MyAppExeName}
ArchitecturesAllowed=x64compatible
ArchitecturesInstallIn64BitMode=x64compatible

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el escritorio"; Flags: unchecked

[Files]
Source: "..\dist\AsistenteRIDescargasPro\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall skipifsilent
