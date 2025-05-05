# Setup simplificado para ADRT Moderno
# Script reduzido para evitar problemas de sintaxe

Write-Host "Iniciando configuração do ADRT Moderno..." -ForegroundColor Cyan

# Diretório atual
$currentDir = (Get-Item -Path ".").FullName
Write-Host "Diretório atual: $currentDir"

# 1. Criar estrutura de diretórios
$diretorios = @(
    "ad-reports\ad-users",
    "ad-reports\ad-admins",
    "ad-reports\ad-enterprise-admins",
    "ad-reports\ad-disabled",
    "ad-reports\ad-lastlogon",
    "ad-reports\ad-neverexpires",
    "ad-reports\ad-groups",
    "ad-reports\ad-membergroups",
    "ad-reports\ad-ous",
    "ad-reports\ad-computers",
    "ad-reports\ad-servers",
    "ad-reports\ad-dcs",
    "ad-reports\ad-gpos",
    "ad-reports\ad-inventory",
    "ad-reports\ad-all",
    "ad-reports\ad-analysis",
    "web\img",
    "templates"
)

foreach ($dir in $diretorios) {
    $dirPath = Join-Path -Path $currentDir -ChildPath $dir
    if (-not (Test-Path -Path $dirPath)) {
        Write-Host "Criando diretório: $dir" -ForegroundColor Yellow
        New-Item -ItemType Directory -Path $dirPath -Force | Out-Null
    }
}

Write-Host "Estrutura de diretórios verificada e criada" -ForegroundColor Green

# 2. Verificar logos
$logoPath = Join-Path -Path $currentDir -ChildPath "web\img\lobios-logo.png"
$logoSmallPath = Join-Path -Path $currentDir -ChildPath "web\img\lobios-logo-small.png"

if (-not (Test-Path -Path $logoPath)) {
    Write-Host "Aviso: Logo principal não encontrado: web\img\lobios-logo.png" -ForegroundColor Yellow
}

if (-not (Test-Path -Path $logoSmallPath)) {
    Write-Host "Aviso: Logo pequeno não encontrado: web\img\lobios-logo-small.png" -ForegroundColor Yellow
}

# 3. Corrigir codificação UTF-8 nos arquivos HTML existentes
Write-Host "Verificando e corrigindo arquivos HTML..." -ForegroundColor Cyan

try {
    $htmlFiles = Get-ChildItem -Path (Join-Path -Path $currentDir -ChildPath "ad-reports") -Filter "*.html" -Recurse -ErrorAction SilentlyContinue
    
    if ($htmlFiles) {
        Write-Host "Encontrados $($htmlFiles.Count) arquivos HTML" -ForegroundColor Gray
        
        foreach ($file in $htmlFiles) {
            Write-Host "Corrigindo codificação: $($file.Name)" -ForegroundColor Gray
            
            try {
                $content = Get-Content -Path $file.FullName -Raw -ErrorAction SilentlyContinue
                if ($content) {
                    [System.IO.File]::WriteAllText($file.FullName, $content, [System.Text.UTF8Encoding]::new($false))
                }
            }
            catch {
                Write-Host "Erro ao processar $($file.Name): $_" -ForegroundColor Red
            }
        }
    }
    else {
        Write-Host "Nenhum arquivo HTML encontrado para corrigir" -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Erro ao pesquisar arquivos HTML: $_" -ForegroundColor Red
}

# 4. Corrigir codificação dos scripts
Write-Host "Corrigindo codificação dos scripts PS1..." -ForegroundColor Cyan

try {
    $psScripts = Get-ChildItem -Path $currentDir -Filter "*.ps1" -ErrorAction SilentlyContinue
    
    foreach ($script in $psScripts) {
        Write-Host "Corrigindo codificação: $($script.Name)" -ForegroundColor Gray
        
        try {
            $content = Get-Content -Path $script.FullName -Raw -ErrorAction SilentlyContinue
            if ($content) {
                [System.IO.File]::WriteAllText($script.FullName, $content, [System.Text.UTF8Encoding]::new($false))
            }
        }
        catch {
            Write-Host "Erro ao processar $($script.Name): $_" -ForegroundColor Red
        }
    }
}
catch {
    Write-Host "Erro ao pesquisar scripts PS1: $_" -ForegroundColor Red
}

# 5. Corrigir codificação do template HTML
$templatePath = Join-Path -Path $currentDir -ChildPath "templates\modern-template.html"
if (Test-Path -Path $templatePath) {
    Write-Host "Corrigindo codificação do template HTML..." -ForegroundColor Cyan
    
    try {
        $content = Get-Content -Path $templatePath -Raw -ErrorAction SilentlyContinue
        if ($content) {
            [System.IO.File]::WriteAllText($templatePath, $content, [System.Text.UTF8Encoding]::new($false))
        }
    }
    catch {
        Write-Host "Erro ao processar o template HTML: $_" -ForegroundColor Red
    }
}

# 6. Corrigir codificação do painel de controle
$indexPath = Join-Path -Path $currentDir -ChildPath "index-modern.html"
if (Test-Path -Path $indexPath) {
    Write-Host "Corrigindo codificação do painel de controle..." -ForegroundColor Cyan
    
    try {
        $content = Get-Content -Path $indexPath -Raw -ErrorAction SilentlyContinue
        if ($content) {
            [System.IO.File]::WriteAllText($indexPath, $content, [System.Text.UTF8Encoding]::new($false))
        }
    }
    catch {
        Write-Host "Erro ao processar o painel de controle: $_" -ForegroundColor Red
    }
}

Write-Host "Configuração concluída!" -ForegroundColor Green
Write-Host "Você pode executar update-all-reports.ps1 para gerar todos os relatórios" -ForegroundColor Cyan

# Perguntar se deseja abrir o painel de controle
$openIndex = Read-Host "Deseja abrir o painel de controle agora? (S/N)"
if ($openIndex -eq "S" -or $openIndex -eq "s") {
    try {
        Start-Process $indexPath
    }
    catch {
        Write-Host "Erro ao abrir o painel de controle: $_" -ForegroundColor Red
    }
}