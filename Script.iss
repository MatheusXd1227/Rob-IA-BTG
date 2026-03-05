[Setup]
AppName=Extrator de Boletos BTG IA
AppVersion=1.0
AppPublisher=BE RH
DefaultDirName={autopf}\Extrator Boletos BTG IA
DisableProgramGroupPage=yes
OutputBaseFilename=Instalador_Extrator_Boletos_IA
Compression=lzma
SolidCompression=yes
PrivilegesRequired=lowest

[Languages]
Name: "portuguese"; MessagesFile: "compiler:Languages\Portuguese.isl"

[Tasks]
Name: "desktopicon"; Description: "{cm:CreateDesktopIcon}"; GroupDescription: "{cm:AdditionalIcons}"; Flags: unchecked

[Files]
; ATENÇÃO: Substitua os caminhos abaixo pelos caminhos reais do seu computador!
Source: "C:\Users\TI_028\OneDrive - be.rh benefícios\Área de Trabalho\meu aplicativo\dist\Extrator de Boletos BTG IA.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "C:\Users\TI_028\Downloads\boletos\modelo.xlsx"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{autoprograms}\Extrator de Boletos BTG IA"; Filename: "{app}\Extrator de Boletos BTG IA.exe"
Name: "{autodesktop}\Extrator de Boletos BTG IA"; Filename: "{app}\Extrator de Boletos BTG IA.exe"; Tasks: desktopicon

[Run]
Filename: "{app}\Extrator de Boletos BTG IA.exe"; Description: "{cm:LaunchProgram,Extrator de Boletos BTG IA}"; Flags: nowait postinstall skipifsilent