; Script generated by the Inno Setup Script Wizard.
; SEE THE DOCUMENTATION FOR DETAILS ON CREATING INNO SETUP SCRIPT FILES!

[Setup]
; NOTE: The value of AppId uniquely identifies this application.
; Do not use the same AppId value in installers for other applications.
; (To generate a new GUID, click Tools | Generate GUID inside the IDE.)
AppId={{A7111B9A-F2BE-4497-82D5-F58FBA8D5CCF}
AppName=QS:Doc
AppVersion=1.5
;AppVerName=QS:Doc
AppPublisher=GEAR
DefaultDirName=c:\qlik\QSDoc
DefaultGroupName=QSDoc
DisableProgramGroupPage=yes
LicenseFile=C:\Projects\QSDoc\versions\QlikSense3.2\license.txt
OutputBaseFilename=QSDoc
Compression=lzma
SolidCompression=yes

[Languages]
Name: "english"; MessagesFile: "compiler:Default.isl"

[Files]
Source: "C:\Projects\QSDoc\versions\QlikSense3.2\create_QSDoc.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Projects\QSDoc\versions\QlikSense3.2\WindowsUCRT\*"; DestDir: "{app}\.\pre_reqs"; Flags: ignoreversion recursesubdirs createallsubdirs
Source: "C:\Projects\QSDoc\versions\QlikSense3.2\word_template\*"; DestDir: "{app}\.\word_template"; Flags: ignoreversion recursesubdirs createallsubdirs
; NOTE: Don't use "Flags: ignoreversion" on any shared system files


