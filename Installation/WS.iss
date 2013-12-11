[Setup]
AppName=Web Stratego
AppVerName=Classic Stratego IP2IP
AppCopyright=Copyright © Uri Goren
DefaultDirName={pf}\Web Stratego
DefaultGroupName=\Uri Goren\Web Stratego
BackColor=clgray
BackColor2=clblack
AppPublisher=Uri Goren
AppPublisherURL=http://www.goren4u.com/
AppUpdatesURL=http://www.goren4u.com/
AppVersion=1.00
ChangesAssociations=true
LicenseFile=License.txt
DisableStartupPrompt=true
WizardStyle=classic
WindowVisible=true
AlwaysCreateUninstallIcon=true
DirExistsWarning=false
CompressLevel=9
OutputDir=..

[Files]
Source: "..\wstratego.exe"; DestDir: "{app}"
Source: "..\sbuild.exe"; DestDir: "{app}"
Source: "..\skins.exe"; DestDir: "{app}"
Source: "..\*.ssz"; DestDir: "{app}"
Source: "VBRUN60.exe"; DestDir: "{tmp}"; CopyMode: alwaysoverwrite; Flags: deleteafterinstall
Source: "mswinsck.ocx"; DestDir: "{sys}"; CopyMode: onlyifdoesntexist; Flags: regserver sharedfile
Source: "SWFlash.ocx"; DestDir: "{sys}"; CopyMode: onlyifdoesntexist; Flags: regserver
Source: "..\images\*.*"; DestDir: "{app}\images"
Source: "..\effects\*.*"; DestDir: "{app}\effects"
Source: "..\skins\unzip32.dll"; DestDir: "{app}"

[Icons]
Name: "{group}\Web Stratego"; Filename: "{app}\Wstratego.exe"; WorkingDir: "{app}"; IconIndex: 1
Name: "{group}\Strategy Builder"; Filename: "{app}\Sbuild.exe"; WorkingDir: "{app}"; IconIndex: 2
Name: "{group}\Skin Selector"; Filename: "{app}\SkinS.exe"; WorkingDir: "{app}"; IconIndex: 3

[Run]
filename: "{tmp}\vbrun60.exe"

[Registry]
Root: HKCR; subkey: ".sss"; valuetype: string; valuename: "default"; valuedata: "WStratego.sss"; flags: uninsdeletekey
Root: HKCR; subkey: "WStratego.sss"; valuetype: string; valuename: "default"; valuedata: "Stratego Strategy sheet"; flags: uninsdeletekey
Root: HKCR; subkey: "WStratego.sss\Shell"; flags: uninsdeletekey
Root: HKCR; subkey: "WStratego.sss\shell\open"; valuetype: string; valuename: "default"; valuedata: "Edit it !"; flags: uninsdeletekey
Root: HKCR; subkey: "WStratego.sss\shell\open\command"; valuetype: string; valuename: "default"; valuedata: "{app}\sbuild.exe %1"; flags: uninsdeletekey
Root: HKCR; subkey: "WStratego.sss\defaulticon"; valuetype: string; valuename: "default"; valuedata: "{app}\sbuild.exe,1"; flags: uninsdeletekey
Root: HKCR; subkey: ".ssz"; valuetype: string; valuename: "default"; valuedata: "WStratego.sss"; flags: uninsdeletekey
Root: HKCR; subkey: "WStratego.ssz"; valuetype: string; valuename: "default"; valuedata: "Stratego Skin Zip"; flags: uninsdeletekey
Root: HKCR; subkey: "WStratego.ssz\Shell"; flags: uninsdeletekey
Root: HKCR; subkey: "WStratego.ssz\shell\open"; valuetype: string; valuename: "default"; valuedata: "Deploy it !"; flags: uninsdeletekey
Root: HKCR; subkey: "WStratego.ssz\shell\open\command"; valuetype: string; valuename: "default"; valuedata: "{app}\skins.exe %1"; flags: uninsdeletekey
Root: HKCR; subkey: "WStratego.ssz\defaulticon"; valuetype: string; valuename: "default"; valuedata: "{app}\skins.exe,1"; flags: uninsdeletekey
