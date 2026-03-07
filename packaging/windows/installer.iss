[Setup]
AppName=Gestion Pacientes
AppVersion=1.0.0
DefaultDirName={autopf}\Gestion Pacientes
DefaultGroupName=Gestion Pacientes
OutputDir=..\..\dist\installer
OutputBaseFilename=GestionPacientesSetup
Compression=lzma
SolidCompression=yes
WizardStyle=modern

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear acceso directo en el escritorio"; GroupDescription: "Accesos directos:"

[Files]
Source: "..\..\dist\GestionPacientes\*"; DestDir: "{app}"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{group}\Gestion Pacientes"; Filename: "{app}\GestionPacientes.exe"
Name: "{autodesktop}\Gestion Pacientes"; Filename: "{app}\GestionPacientes.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\GestionPacientes.exe"; Description: "Abrir Gestion Pacientes"; Flags: nowait postinstall skipifsilent
