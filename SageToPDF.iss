; Package installer for windows clients

[Setup]
AppName=Sage To PDF
AppVersion=1.0
DefaultDirName={pf}\Sage To PDF
DefaultGroupName=Sage To PDF
UninstallDisplayIcon={app}\SageToPDF.exe
Compression=lzma2
SolidCompression=yes
OutputDir=userdocs:SageToPDFSetup

[Files]
Source: "SageToPDF.exe"; DestDir: "{app}"; DestName: "SageToPDF.exe"
Source: "templates\TribeList.xlsx"; DestDir: "{app}\templates"; DestName: "TribeList.xlsx"
Source: "templates\Tribals.docx"; DestDir: "{app}\templates"; DestName: "Tribals.docx"
Source: "templates\TribeList.xlsx"; DestDir: "{app}\templates"; DestName: "TribeList.xlsx"
Source: "help.html"; DestDir: "{app}"
Source: "Readme.md"; DestDir: "{app}"; DestName: "Readme.txt"; Flags: isreadme

[Icons]
Name: "{group}\Sage To PDF"; Filename: "{app}\SageToPDF.exe"
