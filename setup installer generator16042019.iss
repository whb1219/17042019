; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

#define MyAppName "SHIPLYS LCT"
#define MyAppVersion "1.2"
#define MyAppPublisher "Strathclyde University"
#define MyAppURL1 "http://www.strath.ac.uk/" 
#define MyAppURL2 "http://www.shiplys.com/"
#define MyAppExeName "SHIPLYS LCT"

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{FC211E0A-2076-43E3-BFB8-18B0882A9706}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
;AppVerName={#MyAppName} {#MyAppVersion}
AppPublisher={#MyAppPublisher}
AppPublisherURL={#MyAppURL1}
AppSupportURL={#MyAppURL2}
AppUpdatesURL={#MyAppURL2}
DefaultDirName={pf}\{#MyAppName}
DisableProgramGroupPage=yes
OutputBaseFilename=LCT_setup v1.2
Compression=lzma
SolidCompression=yes
PrivilegesRequired=admin

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"
Name: "corsican"; MessagesFile: "compiler:Languages\Corsican.isl"
Name: "french"; MessagesFile: "compiler:Languages\French.isl"
Name: "german"; MessagesFile: "compiler:Languages\German.isl"
Name: "greek"; MessagesFile: "compiler:Languages\Greek.isl"
Name: "portuguese"; MessagesFile: "compiler:Languages\Portuguese.isl"
Name: "scottishgaelic"; MessagesFile: "compiler:Languages\ScottishGaelic.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
Source: "C:\Users\tjb12178\git\17042019\e_deploy\db\*"; DestDir: "{app}\db\"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\tjb12178\git\17042019\e_deploy\03042019_lib\*"; DestDir: "{app}\03042019_lib\"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\tjb12178\git\17042019\e_deploy\pic\*"; DestDir: "{app}\pic\"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\tjb12178\git\17042019\e_deploy\reports\*"; DestDir: "{app}\reports\"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Users\tjb12178\git\17042019\e_deploy\03042019.jar"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\tjb12178\git\17042019\e_deploy\SHIPLYS LCT.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\tjb12178\git\17042019\e_deploy\Help.pdf"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\tjb12178\git\17042019\e_deploy\Training Material.pdf"; DestDir: "{app}"; Flags: ignoreversion
; NOTE: Don't use "Flags: ignoreversion" on any shared system files

[Icons]
Name: "{commonprograms}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{commondesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

;[Run]
;Filename: "{app}\{#MyAppExeName}"; Description: "{cm:LaunchProgram,{#StringChange(MyAppName, '&', '&&')}}"; Flags: nowait postinstall skipifsilent
