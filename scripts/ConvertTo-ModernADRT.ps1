#Requires -Version 5.1
<#
.SYNOPSIS
    Converte scripts ADRT existentes para a versão moderna com interface Lobios
.DESCRIPTION
    Este script converte todos os scripts ADRT (.ps1) em um diretório para
    usar o módulo ADRT-Modern, gerando relatórios com interface gráfica moderna.
.PARAMETER Path
    Caminho onde os scripts ADRT originais estão localizados.
.PARAMETER BackupOriginals
    Cria backup dos arquivos originais antes de substituí-los.
.PARAMETER CreateNew
    Cria novos arquivos com o sufixo "-modern" em vez de substituir os originais.
.EXAMPLE
    .\ConvertTo-ModernADRT.ps1 -Path "C:\ADRT" -CreateNew
.NOTES
    Author: Lobios Segurança • Tecnologia • Inovação
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory=$false)]
    [string]$Path = (Get-Location).Path,
    
    [Parameter(Mandatory=$false)]
    [switch]$BackupOriginals,
    
    [Parameter(Mandatory=$false)]
    [switch]$CreateNew
)

# Verificar se o módulo está disponível
$modulePath = Join-Path -Path $Path -ChildPath "modules\ADRT-Modern.psm1"
if (-not (Test-Path -Path $modulePath)) {
    Write-Error "Módulo ADRT-Modern não encontrado em: $modulePath"
    Write-Error "Por favor, execute primeiro o script Install-ADRT-Modern.ps1"
    return
}

# Importar o módulo
try {
    Import-Module $modulePath -ErrorAction Stop
    Write-Host "Módulo ADRT-Modern importado com sucesso." -ForegroundColor Green
}
catch {
    Write-Error "Erro ao importar o módulo ADRT-Modern: $_"
    return
}

# Banner
Write-Host @"

 █████╗ ██████╗ ██████╗ ████████╗    ███╗   ███╗ ██████╗ ██████╗ ███████╗██████╗ ███╗   ██╗
██╔══██╗██╔══██╗██╔══██╗╚══██╔══╝    ████╗ ████║██╔═══██╗██╔══██╗██╔════╝██╔══██╗████╗  ██║
███████║██║  ██║██████╔╝   ██║       ██╔████╔██║██║   ██║██║  ██║█████╗  ██████╔╝██╔██╗ ██║
██╔══██║██║  ██║██╔══██╗   ██║       ██║╚██╔╝██║██║   ██║██║  ██║██╔══╝  ██╔══██╗██║╚██╗██║
██║  ██║██████╔╝██║  ██║   ██║       ██║ ╚═╝ ██║╚██████╔╝██████╔╝███████╗██║  ██║██║ ╚████║
╚═╝  ╚═╝╚═════╝ ╚═╝  ╚═╝   ╚═╝       ╚═╝     ╚═╝ ╚═════╝ ╚═════╝ ╚══════╝╚═╝  ╚═╝╚═╝  ╚═══╝
                                                                                                                
                         Conversor de Scripts para Interface Moderna
                         Lobios Segurança • Tecnologia • Inovação

"@ -ForegroundColor Magenta

# Localizar todos os scripts ADRT
Write-Host "Procurando scripts ADRT em: $Path" -ForegroundColor Cyan
$adrtScripts = Get-ChildItem -Path $Path -Filter "ad-*.ps1" -File | Where-Object { $_.Name -notlike "*-modern.ps1" }

if ($adrtScripts.Count -eq 0) {
    Write-Warning "Nenhum script ADRT encontrado em: $Path"
    return
}

Write-Host "Encontrados $($adrtScripts.Count) scripts ADRT." -ForegroundColor Green

# Processar cada script
$convertedCount = 0
$errorCount = 0

foreach ($script in $adrtScripts) {
    Write-Host "Processando: $($script.Name)" -ForegroundColor Cyan
    
    try {
        # Definir caminho de saída
        $outputPath = ""
        if ($CreateNew) {
            $baseName = [System.IO.Path]::GetFileNameWithoutExtension($script.Name)
            $extension = [System.IO.Path]::GetExtension($script.Name)
            $outputPath = Join-Path -Path $script.DirectoryName -ChildPath "$baseName-modern$extension"
        }
        else {
            # Fazer backup se solicitado
            if ($BackupOriginals) {
                $backupPath = Join-Path -Path $script.DirectoryName -ChildPath "$($script.Name).bak"
                Copy-Item -Path $script.FullName -Destination $backupPath -Force
                Write-Host "  Backup criado: $backupPath" -ForegroundColor Gray
            }
            
            $outputPath = $script.FullName
        }
        
        # Converter o script
        $result = Convert-ADRTScript -ScriptPath $script.FullName -OutputPath $outputPath
        
        if ($result) {
            $convertedCount++
            Write-Host "  Convertido com sucesso: $outputPath" -ForegroundColor Green
        }
        else {
            $errorCount++
            Write-Host "  Erro ao converter o script." -ForegroundColor Red
        }
    }
    catch {
        $errorCount++
        Write-Host "  Erro: $_" -ForegroundColor Red
    }
    
    Write-Host ""
}

# Resumo
Write-Host "Conversão concluída!" -ForegroundColor Green
Write-Host "Scripts convertidos: $convertedCount" -ForegroundColor Green
if ($errorCount -gt 0) {
    Write-Host "Erros encontrados: $errorCount" -ForegroundColor Red
}

# Mostrar próximos passos
Write-Host @"

Próximos Passos:
1. Execute um dos scripts convertidos para testar a nova interface.
2. Verifique se os logotipos da Lobios estão no diretório: $Path\web\img
3. Personalize as cores no arquivo CSS, se necessário.

Exemplo de script para executar: $Path\ad-users-modern.ps1

"@ -ForegroundColor Cyan

# Oferecer para executar um script convertido
$runExample = Read-Host "Deseja executar um script de exemplo agora? (S/N)"
if ($runExample -eq "S" -or $runExample -eq "s") {
    $examplePath = Join-Path -Path $Path -ChildPath "ad-users-modern.ps1"
    if (Test-Path -Path $examplePath) {
        Write-Host "Executando script de exemplo: $examplePath" -ForegroundColor Cyan
        & $examplePath
    }
    else {
        Write-Host "Script de exemplo não encontrado: $examplePath" -ForegroundColor Yellow
        $anyScript = $adrtScripts | Where-Object { $_.Name -like "ad-users*" } | Select-Object -First 1
        if ($anyScript) {
            $modernPath = Join-Path -Path $anyScript.DirectoryName -ChildPath "$([System.IO.Path]::GetFileNameWithoutExtension($anyScript.Name))-modern.ps1"
            if (Test-Path -Path $modernPath) {
                Write-Host "Executando script alternativo: $modernPath" -ForegroundColor Cyan
                & $modernPath
            }
        }
    }
}