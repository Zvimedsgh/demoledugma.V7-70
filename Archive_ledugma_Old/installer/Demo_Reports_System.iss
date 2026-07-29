[Setup]
AppName=Demo Reports System
AppVersion=7.70
DefaultDirName={localappdata}\Demo Reports System
DefaultGroupName=Demo Reports System
OutputDir=C:\\Temp
OutputBaseFilename=Demo_Reports_System_Setup_V7.70
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest
DisableProgramGroupPage=yes

[Files]
Source: "C:\ledugma\installer\Demo_Reports_System_V7.70.xlsm"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{userdesktop}\Demo Reports System"; Filename: "{app}\Demo_Reports_System_V7.70.xlsm"
Name: "{group}\Demo Reports System"; Filename: "{app}\Demo_Reports_System_V7.70.xlsm"

[Registry]
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\DemoReportsSystem"; ValueType: string; ValueName: "Path"; ValueData: "{app}\"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\DemoReportsSystem"; ValueType: dword; ValueName: "AllowSubfolders"; ValueData: "1"; Flags: uninsdeletekey
Root: HKCU; Subkey: "Software\Microsoft\Office\16.0\Excel\Security\Trusted Locations\DemoReportsSystem"; ValueType: string; ValueName: "Description"; ValueData: "Demo Reports System trusted location"; Flags: uninsdeletekey

[Run]
Filename: "{app}\Demo_Reports_System_V7.70.xlsm"; Description: "Open Demo Reports System"; Flags: postinstall shellexec skipifsilent