; Package installer for windows clients

[Setup]
AppName=Sage To PDF
AppVersion=1.0
DefaultDirName={pf}\SageToPDF
DefaultGroupName=SageToPDF
UninstallDisplayIcon={app}\SageToPDF.exe
Compression=lzma2
SolidCompression=yes
OutputDir=userdocs:SageToPDFSetup
PrivilegesRequired=admin

[Files]
Source: "dist\main.exe"; DestDir: "{app}"; DestName: "SageToPDF.exe"
Source: "templates\TribeList.xlsx"; DestDir: "{app}\templates"; DestName: "TribeList.xlsx"
Source: "templates\Tribals.docx"; DestDir: "{app}\templates"; DestName: "Tribals.docx"
Source: "templates\Mapping.docx"; DestDir: "{app}\templates"; DestName: "Mapping.docx"
Source: "templates\Ethno.docx"; DestDir: "{app}\templates"; DestName: "Ethno.docx"
Source: "templates\Field Survey.docx"; DestDir: "{app}\templates"; DestName: "Field Survey.docx"
Source: "templates\TribeList.xlsx"; DestDir: "{app}\templates"; DestName: "TribeList.xlsx"
Source: "help.html"; DestDir: "{app}"
Source: "Readme.md"; DestDir: "{app}"; DestName: "Readme.txt"; Flags: isreadme
Source: "default.prop"; DestDir: "{app}"; DestName: ".prop";
; DLLs
Source: dist\main\mfc100u.dll; DestDir: {app}\bin; Flags: restartreplace ignoreversion
Source: dist\main\msvcr100.dll; DestDir: {app}\bin; Flags: restartreplace ignoreversion
Source: dist\main\python34.dll; DestDir: {app}\bin; Flags: restartreplace ignoreversion
Source: dist\main\pythoncom34.dll; DestDir: {app}\bin; Flags: restartreplace ignoreversion
Source: dist\main\pywintypes34.dll; DestDir: {app}\bin; Flags: restartreplace ignoreversion
Source: dist\main\tcl86t.dll; DestDir: {app}\bin; Flags: restartreplace ignoreversion
Source: dist\main\tk86t.dll; DestDir: {app}\bin; Flags: restartreplace ignoreversion

[Icons]
Name: "{group}\Sage To PDF"; Filename: "{app}\SageToPDF.exe"
