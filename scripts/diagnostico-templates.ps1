# Script de diagnóstico para o módulo ADRT-Templates (Versão corrigida)

# Verificar se o módulo pode ser encontrado
$modulePath = "C:\Files\ad-assessment\ad-assessment\modules\ADRT-Templates.psm1"
if (Test-Path -Path $modulePath) {
    Write-Host "✓ Módulo encontrado: $modulePath" -ForegroundColor Green
} else {
    Write-Host "✗ Módulo não encontrado: $modulePath" -ForegroundColor Red
}

# Tentar importar o módulo
try {
    Import-Module $modulePath -Force -ErrorAction Stop
    Write-Host "✓ Módulo importado com sucesso!" -ForegroundColor Green
} catch {
    Write-Host "✗ Erro ao importar o módulo: $_" -ForegroundColor Red
    exit
}

# Verificar se a função está disponível
if (Get-Command -Name "New-ADRTHtmlFromTemplate" -ErrorAction SilentlyContinue) {
    Write-Host "✓ Função New-ADRTHtmlFromTemplate está disponível" -ForegroundColor Green
    
    # Olhar os parâmetros da função
    $funcParams = (Get-Command -Name "New-ADRTHtmlFromTemplate").Parameters
    Write-Host "Parâmetros da função:" -ForegroundColor Cyan
    $funcParams.Keys | ForEach-Object {
        $param = $funcParams[$_]
        $mandatory = if ($param.Attributes.Mandatory) { "Obrigatório" } else { "Opcional" }
        Write-Host "  - $_ ($mandatory)" -ForegroundColor Gray
    }
} else {
    Write-Host "✗ Função New-ADRTHtmlFromTemplate NÃO está disponível" -ForegroundColor Red
}

# Examinar o código da função para entender o que ela faz
Write-Host "`nVamos examinar a função para ver como ela localiza os templates:" -ForegroundColor Yellow
$functionDefinition = (Get-Command -Name "New-ADRTHtmlFromTemplate" -ErrorAction SilentlyContinue).ScriptBlock
Write-Host $functionDefinition

# Verificar os caminhos que a função usa para os templates
Write-Host "`nVamos verificar os caminhos usados pela função:" -ForegroundColor Yellow
$baseDir = Split-Path -Parent (Split-Path -Parent $modulePath)
Write-Host "Diretório base (baseDir): $baseDir" -ForegroundColor Gray

$headerPath = Join-Path -Path $baseDir -ChildPath "templates\header.html"
$sidebarPath = Join-Path -Path $baseDir -ChildPath "templates\sidebar.html"
$footerPath = Join-Path -Path $baseDir -ChildPath "templates\footer.html"

$headerExists = Test-Path -Path $headerPath
$sidebarExists = Test-Path -Path $sidebarPath
$footerExists = Test-Path -Path $footerPath

if ($headerExists) {
    Write-Host "✓ Arquivo header.html encontrado: $headerPath" -ForegroundColor Green
} else {
    Write-Host "✗ Arquivo header.html NÃO encontrado: $headerPath" -ForegroundColor Red
}

if ($sidebarExists) {
    Write-Host "✓ Arquivo sidebar.html encontrado: $sidebarPath" -ForegroundColor Green
} else {
    Write-Host "✗ Arquivo sidebar.html NÃO encontrado: $sidebarPath" -ForegroundColor Red
}

if ($footerExists) {
    Write-Host "✓ Arquivo footer.html encontrado: $footerPath" -ForegroundColor Green
} else {
    Write-Host "✗ Arquivo footer.html NÃO encontrado: $footerPath" -ForegroundColor Red
}

# Sugerir uma solução baseada no que foi encontrado
Write-Host "`nBaseado na análise, aqui está uma sugestão:" -ForegroundColor Yellow
if ($headerExists -and $sidebarExists -and $footerExists) {
    Write-Host "Todos os arquivos de template foram encontrados. O problema pode estar na forma como o módulo está calculando os caminhos." -ForegroundColor Gray
    Write-Host "Vamos ver o código do módulo para identificar o problema." -ForegroundColor Gray
    
    # Examinar o conteúdo do módulo
    $moduleContent = Get-Content -Path $modulePath -Raw
    
    # Verificar se existe um problema semelhante no módulo
    if ($moduleContent -match "Test-Path.*-and.*Test-Path") {
        Write-Host "`nPROBLEMA IDENTIFICADO:" -ForegroundColor Red
        Write-Host "O módulo está usando 'Test-Path -and Test-Path', o que causa o erro 'Cannot bind parameter because parameter 'Path' is specified more than once'" -ForegroundColor Yellow
        Write-Host "Este é o mesmo erro que encontramos no script de diagnóstico." -ForegroundColor Yellow
        
        Write-Host "`nSOLUÇÃO:" -ForegroundColor Green
        Write-Host "O módulo precisa ser modificado para verificar cada caminho separadamente." -ForegroundColor Gray
        Write-Host "Em vez de:" -ForegroundColor Gray
        Write-Host "  if (-not (Test-Path -Path \$headerPath) -or -not (Test-Path -Path \$sidebarPath) -or -not (Test-Path -Path \$footerPath))" -ForegroundColor Gray
        Write-Host "Use:" -ForegroundColor Gray
        Write-Host "  \$headerExists = Test-Path -Path \$headerPath" -ForegroundColor Gray
        Write-Host "  \$sidebarExists = Test-Path -Path \$sidebarPath" -ForegroundColor Gray
        Write-Host "  \$footerExists = Test-Path -Path \$footerPath" -ForegroundColor Gray
        Write-Host "  if (-not \$headerExists -or -not \$sidebarExists -or -not \$footerExists)" -ForegroundColor Gray
    } else {
        Write-Host "Não foi possível identificar o problema específico no módulo." -ForegroundColor Yellow
    }
} else {
    Write-Host "Alguns arquivos de template não foram encontrados. Você precisa verificar se eles existem nos locais corretos." -ForegroundColor Gray
    Write-Host "Caminhos esperados:" -ForegroundColor Gray
    Write-Host "  - Header: $headerPath" -ForegroundColor Gray
    Write-Host "  - Sidebar: $sidebarPath" -ForegroundColor Gray
    Write-Host "  - Footer: $footerPath" -ForegroundColor Gray
}

# Manter aberto para visualização
Write-Host "`nPressione qualquer tecla para fechar..." -ForegroundColor Cyan
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")