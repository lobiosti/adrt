[Setup]
AppName=ADRT
AppVersion=1.0
DefaultDirName={userpf}\ADRT
DisableProgramGroupPage=yes
OutputDir=.\Output
OutputBaseFilename=ADRT-Setup
Compression=lzma
SolidCompression=yes

[Files]
; Nenhum arquivo local. O script PowerShell baixa os arquivos do GitHub.
; Colocaremos apenas o script de instalação.

Source: "Install-ADRT.ps1"; DestDir: "{app}"; Flags: ignoreversion

[Icons]
Name: "{userdesktop}\ADRT"; Filename: "pwsh"; Parameters: "-ExecutionPolicy Bypass -File `"{app}\Start-ADRT.ps1`""
WorkingDir: "{app}"; IconFilename: "pwsh"; IconIndex: 0

[Run]
; Executar script de instalação em PowerShell 5 (pré-requisito do sistema)
Filename: "powershell.exe"; \
  Parameters: "-ExecutionPolicy Bypass -WindowStyle Hidden -File `"{app}\Install-ADRT.ps1`""; \
  StatusMsg: "Instalando ADRT..."; \
  Flags: runhidden

; Perguntar se deseja rodar ao final
Filename: "pwsh"; \
  Parameters: "-ExecutionPolicy Bypass -File `"{app}\Start-ADRT.ps1`""; \
  Description: "Iniciar ADRT agora"; \
  Flags: postinstall runascurrentuser unchecked

[Code]
// Garante que o PowerShell 7 esteja disponível antes de executar Start-ADRT.ps1
// Isso será tratado dentro do script PowerShell mesmo
