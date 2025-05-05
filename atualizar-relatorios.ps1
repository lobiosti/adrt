# atualizar-relatorios.ps1
# Script melhorado para atualizar todos os relatórios ADRT Modernos
# Inclui correção de codificação para acentuação correta

# Definir codificação para garantir acentuação correta
$OutputEncoding = [System.Text.UTF8Encoding]::new()
$PSDefaultParameterValues['Out-File:Encoding'] = 'UTF8'

# Banner
Write-Host @"

╔═══════════════════════════════════════════════════════════════╗
║                                                               ║
║      ██╗      ██████╗ ██████╗ ██╗ ██████╗ ███████╗           ║
║      ██║     ██╔═══██╗██╔══██╗██║██╔═══██╗██╔════╝           ║
║      ██║     ██║   ██║██████╔╝██║██║   ██║███████╗           ║
║      ██║     ██║   ██║██╔══██╗██║██║   ██║╚════██║           ║
║      ███████╗╚██████╔╝██████╔╝██║╚██████╔╝███████║           ║
║      ╚══════╝ ╚═════╝ ╚═════╝ ╚═╝ ╚═════╝ ╚══════╝           ║
║                                                               ║
║      ADRT - Active Directory Report Tool (Versão Moderna)     ║
║                                                               ║
╚═══════════════════════════════════════════════════════════════╝
"@ -ForegroundColor Magenta

Write-Host "Iniciando atualização de todos os relatórios ADRT..." -ForegroundColor Cyan
Write-Host "Este processo pode levar alguns minutos, dependendo do tamanho do seu Active Directory." -ForegroundColor Yellow
Write-Host ""

# Lista de scripts a serem executados em ordem
$scripts = @(
    "ad-users-modern.ps1",
    "ad-admins-modern.ps1",
    "ad-enterprise-admins-modern.ps1",
    "ad-disabled-modern.ps1",
    "ad-lastlogon-modern.ps1",
    "ad-neverexpires-modern.ps1",
    "ad-groups-modern.ps1",
    "ad-membergroups-modern.ps1",
    "ad-ous-modern.ps1",
    "ad-computers-modern.ps1",
    "ad-servers-modern.ps1",
    "ad-dcs-modern.ps1",
    "ad-gpos-modern.ps1",
    "ad-inventory-modern.ps1"
)

# Inicializar contadores
$totalScripts = $scripts.Count
$sucessos = 0
$falhas = 0
$currentScript = 0

Write-Host "Total de relatórios a processar: $totalScripts" -ForegroundColor White
Write-Host ""

# Executar cada script
foreach ($script in $scripts) {
    $currentScript++
    $porcentagem = [math]::Round(($currentScript / $totalScripts) * 100)
    
    # Barra de progresso
    $barSize = 30
    $progressBar = "["
    $completeSize = [math]::Floor($barSize * ($porcentagem / 100))
    $remainingSize = $barSize - $completeSize
    
    $progressBar += "".PadLeft($completeSize, "█")
    $progressBar += "".PadLeft($remainingSize, "░")
    $progressBar += "] $porcentagem%"
    
    Write-Host "$progressBar Executando: $script" -ForegroundColor Cyan
    
    $scriptPath = Join-Path -Path (Get-Location).Path -ChildPath $script
    
    if (Test-Path -Path $scriptPath) {
        try {
            # Executar o script em um novo escopo para evitar conflitos
            & $scriptPath
            
            if ($LASTEXITCODE -eq 0 -or $null -eq $LASTEXITCODE) {
                $sucessos++
                Write-Host "  ✓ Concluído com sucesso!" -ForegroundColor Green
            } else {
                $falhas++
                Write-Host "  ✗ Falha na execução. Código de saída: $LASTEXITCODE" -ForegroundColor Red
            }
        }
        catch {
            $falhas++
            Write-Host "  ✗ Erro ao executar $script : $_" -ForegroundColor Red
        }
    }
    else {
        Write-Host "  ✗ Arquivo não encontrado: $scriptPath" -ForegroundColor Red
        $falhas++
    }
    
    Write-Host ""
}

# Executar a análise completa
Write-Host "Gerando análise completa do Active Directory..." -ForegroundColor Cyan
$analiseScript = Join-Path -Path (Get-Location).Path -ChildPath "analise-completa.ps1"

if (Test-Path -Path $analiseScript) {
    try {
        & $analiseScript
        $sucessos++
        Write-Host "  ✓ Análise completa gerada com sucesso!" -ForegroundColor Green
    }
    catch {
        $falhas++
        Write-Host "  ✗ Erro ao gerar análise completa: $_" -ForegroundColor Red
    }
}
else {
    Write-Host "  ✗ Script de análise não encontrado: $analiseScript" -ForegroundColor Red
    $falhas++
}

Write-Host ""

# Corrigir a codificação de todos os arquivos HTML gerados
Write-Host "Corrigindo codificação UTF-8 em todos os relatórios HTML..." -ForegroundColor Cyan

$diretorioReports = Join-Path -Path (Get-Location).Path -ChildPath "ad-reports"
if (Test-Path -Path $diretorioReports) {
    try {
        $htmlFiles = Get-ChildItem -Path $diretorioReports -Filter "*.html" -Recurse -ErrorAction Continue
        $countFiles = 0
        
        foreach ($arquivo in $htmlFiles) {
            try {
                # Ler o conteúdo do arquivo
                $conteudo = Get-Content -Path $arquivo.FullName -Raw -ErrorAction Continue
                
                # Reescrever com codificação UTF-8 sem BOM
                if ($conteudo) {
                    [System.IO.File]::WriteAllText($arquivo.FullName, $conteudo, [System.Text.UTF8Encoding]::new($false))
                    $countFiles++
                }
            }
            catch {
                Write-Host "  ✗ Erro ao corrigir $($arquivo.Name): $_" -ForegroundColor Red
            }
        }
        
        Write-Host "  ✓ $countFiles arquivos HTML processados e corrigidos." -ForegroundColor Green
    }
    catch {
        Write-Host "  ✗ Erro ao listar arquivos HTML: $_" -ForegroundColor Red
    }
}
else {
    Write-Host "  ! Diretório de relatórios não encontrado: $diretorioReports" -ForegroundColor Yellow
}

Write-Host ""

# Abrir o painel de controle
Write-Host "Abrindo o painel de controle..." -ForegroundColor Cyan

try {
    $indexPath = Join-Path -Path (Get-Location).Path -ChildPath "index-modern.html"
    if (Test-Path -Path $indexPath) {
        Start-Process $indexPath
        Write-Host "  ✓ Painel de controle aberto com sucesso!" -ForegroundColor Green
    }
    else {
        Write-Host "  ✗ Arquivo index-modern.html não encontrado." -ForegroundColor Red
    }
}
catch {
    Write-Host "  ✗ Erro ao abrir o painel de controle: $_" -ForegroundColor Red
}

# Resumo da execução
Write-Host ""
Write-Host "╔═══════════════════════════════════════════════════════════════╗" -ForegroundColor Cyan
Write-Host "║                      RESUMO DA EXECUÇÃO                       ║" -ForegroundColor Cyan
Write-Host "╚═══════════════════════════════════════════════════════════════╝" -ForegroundColor Cyan
Write-Host ""
Write-Host "  • Scripts executados com sucesso: $sucessos" -ForegroundColor Green
if ($falhas -gt 0) {
    Write-Host "  • Scripts com falhas: $falhas" -ForegroundColor Red
}
else {
    Write-Host "  • Scripts com falhas: 0" -ForegroundColor Green
}
Write-Host ""
Write-Host "Os relatórios estão disponíveis na pasta 'ad-reports' e podem"
Write-Host "ser acessados pelo painel de controle que foi aberto automaticamente."
Write-Host ""
Write-Host "Para visualizar novamente o painel de controle, abra o arquivo 'index-modern.html'."
Write-Host ""
Write-Host "Atualização concluída!" -ForegroundColor Green