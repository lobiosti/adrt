# ADRT-Notification.psm1
# Fun��es para enviar notifica��es por email e Telegram

function Send-EmailNotification {
    param (
        [string]$To = "suporte@lobios.com.br",
        [string]$From = "bot@enviodelog.com.br",
        [string]$SmtpServer = "mail.enviodelog.com.br",
        [string]$Port = "587",
        [string]$Subject,
        [string]$HtmlBody,
        [array]$Attachments = @(),
        [bool]$UseSSL = $false,
        [PSCredential]$Credential = $null
    )
    
    try {
        # Configurar a codifica��o para UTF-8
        $encoding = [System.Text.Encoding]::UTF8
        
        $emailParams = @{
            From = $From
            To = $To
            Subject = $Subject
            BodyAsHtml = $true
            Body = $HtmlBody
            SmtpServer = $SmtpServer
            Port = $Port
            Encoding = $encoding
        }
        
        # Adicionar anexos se houver
        if ($Attachments.Count -gt 0) {
            $validAttachments = @()
            foreach ($attachment in $Attachments) {
                if (Test-Path -Path $attachment) {
                    $validAttachments += $attachment
                }
            }
            
            if ($validAttachments.Count -gt 0) {
                $emailParams.Add("Attachments", $validAttachments)
            }
        }
        
        # Adicionar credenciais se fornecidas
        if ($Credential) {
            $emailParams.Add("Credential", $Credential)
        }
        
        if ($UseSSL) {
            $emailParams.Add("UseSsl", $true)
        }
        
        Send-MailMessage @emailParams
        Write-Host "? Email com relat�rio enviado com sucesso para $To" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "? Erro ao enviar email: $_" -ForegroundColor Red
        return $false
    }
}

function Send-ADRTUpdateNotification {
    [CmdletBinding()]
    param (
        [string]$ScriptName,
        [int]$TotalScripts = 0,
        [int]$Sucessos = 0,
        [int]$Falhas = 0,
        [array]$ScriptsList = @(),
        [string]$OutputPath = ""
    )
    
    # Valores fixos para email e Telegram
    $emailEnabled = $true
    $emailFrom = "bot@enviodelog.com.br"
    $emailTo = "suporte@lobios.com.br"
    $smtpServer = "mail.enviodelog.com.br"
    $smtpPort = 587
    $emailUseSSL = $false
    $emailUsername = "bot@enviodelog.com.br"
    $emailPassword = "n4!ve9e70s2="
    
    $telegramEnabled = $true
    $telegramBotToken = "6023316555:AAEj6mmY0gYiPVJt67c10Cj7aobE5HnLi58"
    $telegramChatID = "-4708211611"
    
    $date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $hostname = [System.Net.Dns]::GetHostName()
    $username = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    $domain = (Get-ADDomain).Forest
    
    # Criar mensagem de notificação para Telegram
    $message = @"
<b>Relatório ADRT - Atualização de Relatórios</b>
📊 <b>Script:</b> $ScriptName
📅 <b>Data:</b> $date
🖥︝ <b>Servidor:</b> $hostname
👤 <b>Usuário:</b> $username
🝢 <b>Domínio:</b> $domain

<b>Estatísticas:</b>
• Total de relatórios: $TotalScripts
• Sucessos: $Sucessos
• Falhas: $Falhas
"@
    
    # Enviar email se habilitado
    $emailSuccess = $false
    if ($emailEnabled) {
        $emailSubject = "ADRT - Atualização de Relatórios [$date]"
        $emailBody = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: Calibri; font-size: 14px; }
        table { margin: auto; width: 80%; border-collapse: collapse; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: center; }
        th { background-color: #f2f2f2; }
        .header { background-color: #6a3094; color: white; padding: 10px; text-align: center; }
        .footer { background-color: #f2f2f2; color: #666; padding: 5px; text-align: center; font-size: 12px; }
        .info { background-color: #f8f9fa; border-left: 4px solid #6a3094; padding: 15px; margin: 10px 0; }
        .success { color: green; }
        .error { color: red; }
    </style>
</head>
<body>
    <div class="header">
        <h2>Atualização de Relatórios ADRT - $date</h2>
    </div>
    
    <div class="info">
        <p><strong>Script:</strong> $ScriptName</p>
        <p><strong>Servidor:</strong> $hostname</p>
        <p><strong>Usuário:</strong> $username</p>
        <p><strong>Domínio:</strong> $domain</p>
    </div>
    
    <table>
        <tr>
            <th colspan="2">Resumo da Execução</th>
        </tr>
        <tr>
            <td>Total de Relatórios</td>
            <td>$TotalScripts</td>
        </tr>
        <tr>
            <td>Sucessos</td>
            <td class="success">$Sucessos</td>
        </tr>
        <tr>
            <td>Falhas</td>
            <td class="error">$Falhas</td>
        </tr>
    </table>
"@

        # Adicionar a lista de scripts executados se disponível
        if ($ScriptsList.Count -gt 0) {
            $emailBody += @"
    
    <h3>Scripts Executados:</h3>
    <ul>
"@
            foreach ($script in $ScriptsList) {
                $emailBody += @"
        <li>$script</li>
"@
            }
            
            $emailBody += @"
    </ul>
"@
        }

        # Fechar o HTML
        $emailBody += @"
    
    <p>Este é um email automático enviado pelo sistema ADRT.</p>
    
    <div class="footer">
        <p>ADRT - Active Directory Report Tool v2.0 | Desenvolvido por Lobios Segurança • Tecnologia • Inovação</p>
    </div>
</body>
</html>
"@

        # Anexos - apenas se o caminho existir
        $attachments = @()
        if ($OutputPath -and (Test-Path -Path $OutputPath)) {
            $attachments += $OutputPath
        }
        
        # Configurar credenciais se fornecidas
        $credential = $null
        if ($emailUsername -and $emailPassword) {
            $securePassword = ConvertTo-SecureString $emailPassword -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($emailUsername, $securePassword)
        }
        
        $emailSuccess = Send-EmailNotification -To $emailTo `
                              -From $emailFrom `
                              -SmtpServer $smtpServer `
                              -Port $smtpPort `
                              -Subject $emailSubject `
                              -HtmlBody $emailBody `
                              -Attachments $attachments `
                              -UseSSL $emailUseSSL `
                              -Credential $credential
    }
    
    # Enviar notificação Telegram se habilitado
    $telegramSuccess = $false
    if ($telegramEnabled) {
        $telegramSuccess = Send-TelegramNotification -BotToken $telegramBotToken `
                                 -ChatID $telegramChatID `
                                 -Message $message
    }
    
    # Retornar sucesso se pelo menos um método funcionou
    return ($emailSuccess -or $telegramSuccess)
}
function Send-TelegramNotification {
    param (
        [string]$BotToken = "6023316555:AAEj6mmY0gYiPVJt67c10Cj7aobE5HnLi58",
        [string]$ChatID = "-4708211611",
        [string]$Message
    )
    
    try {
        # Remove "bot" prefix if present
        if ($BotToken.StartsWith("bot")) {
            $BotToken = $BotToken.Substring(3)
        }
        
        # Construir URL correto garantindo que n�o haja espa�os ou caracteres estranhos
        $BotToken = $BotToken.Trim()
        $ChatID = $ChatID.Trim()
        $telegramURL = "https://api.telegram.org/bot$BotToken/sendMessage"
        
        $params = @{
            chat_id = $ChatID
            text = $Message
            parse_mode = "HTML"
        }
        
        # Converter para JSON
        $body = $params | ConvertTo-Json
        
        # Enviar a solicita��o HTTP POST
        $result = Invoke-RestMethod -Method Post -Uri $telegramURL -ContentType "application/json" -Body $body
        
        if ($result.ok) {
            Write-Host "? Notifica��o Telegram enviada com sucesso" -ForegroundColor Green
            return $true
        } else {
            Write-Host "? Erro ao enviar notifica��o Telegram: $($result | ConvertTo-Json)" -ForegroundColor Red
            return $false
        }
    }
    catch {
        Write-Host "? Erro ao enviar notifica��o Telegram: $_" -ForegroundColor Red
        return $false
    }
}

function Send-ADRTNotification {
    [CmdletBinding()]
    param (
        [string]$ScriptName,
        [string]$Type,
        [hashtable]$Stats,
        [string]$Domain,
        [string]$ReportPath,
        [array]$Attachments = @()
    )
    
    # Valores fixos para email e Telegram
    $emailEnabled = $true
    $emailFrom = "bot@enviodelog.com.br"
    $emailTo = "suporte@lobios.com.br"
    $smtpServer = "mail.enviodelog.com.br"
    $smtpPort = 587
    $emailUseSSL = $false
    $emailUsername = "bot@enviodelog.com.br"
    $emailPassword = "n4!ve9e70s2="
    
    $telegramEnabled = $true
    $telegramBotToken = "6023316555:AAEj6mmY0gYiPVJt67c10Cj7aobE5HnLi58"
    $telegramChatID = "-4708211611"
    
    $date = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $hostname = [System.Net.Dns]::GetHostName()
    $username = [System.Security.Principal.WindowsIdentity]::GetCurrent().Name
    
    # Criar mensagem de notifica��o para Telegram
    $message = @"
<b>Relat�rio ADRT - $Type</b>
?? <b>Script:</b> $ScriptName
?? <b>Data:</b> $date
??? <b>Servidor:</b> $hostname
?? <b>Usu�rio:</b> $username
?? <b>Dom�nio:</b> $Domain

<b>Estat�sticas:</b>
� Usu�rios: $($Stats.TotalUsers)
� Computadores: $($Stats.TotalComputers)
� Servidores: $($Stats.TotalServers)
� Grupos: $($Stats.TotalGroups)
� OUs: $($Stats.TotalOUs)
� GPOs: $($Stats.TotalGPOs)
"@
    
    # Enviar email se habilitado
    $emailSuccess = $false
    if ($emailEnabled) {
        $emailSubject = "ADRT - Relat�rio $Type Gerado [$date]"
        $emailBody = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="UTF-8">
    <style>
        body { font-family: Calibri; font-size: 14px; }
        table { margin: auto; width: 80%; border-collapse: collapse; }
        th, td { border: 1px solid #ddd; padding: 8px; text-align: center; }
        th { background-color: #f2f2f2; }
        .header { background-color: #6a3094; color: white; padding: 10px; text-align: center; }
        .footer { background-color: #f2f2f2; color: #666; padding: 5px; text-align: center; font-size: 12px; }
        .info { background-color: #f8f9fa; border-left: 4px solid #6a3094; padding: 15px; margin: 10px 0; }
    </style>
</head>
<body>
    <div class="header">
        <h2>Relatorio do Active Directory - $date</h2>
    </div>
    
    <div class="info">
        <p><strong>Script:</strong> $ScriptName</p>
        <p><strong>Servidor:</strong> $hostname</p>
        <p><strong>Usuario:</strong> $username</p>
        <p><strong>Dominio:</strong> $Domain</p>
        <p><strong>Caminho do Relatorio:</strong> $ReportPath</p>
    </div>
    
    <table>
        <tr>
            <th colspan="2">Resumo do Active Directory</th>
        </tr>
        <tr>
            <td>Usuarios</td>
            <td>$($Stats.TotalUsers)</td>
        </tr>
        <tr>
            <td>Computadores</td>
            <td>$($Stats.TotalComputers)</td>
        </tr>
        <tr>
            <td>Servidores</td>
            <td>$($Stats.TotalServers)</td>
        </tr>
        <tr>
            <td>Grupos</td>
            <td>$($Stats.TotalGroups)</td>
        </tr>
        <tr>
            <td>OUs</td>
            <td>$($Stats.TotalOUs)</td>
        </tr>
        <tr>
            <td>GPOs</td>
            <td>$($Stats.TotalGPOs)</td>
        </tr>
    </table>
    
    <p>Este e um email automatico enviado pelo sistema ADRT.</p>
    
    <div class="footer">
        <p>ADRT - Active Directory Report Tool v2.0 | Desenvolvido por Lobios Seguranca � Tecnologia � Inovacao</p>
    </div>
</body>
</html>
"@

        # Adicionar o relat�rio aos anexos se existir
        if (Test-Path -Path $ReportPath) {
            $Attachments += $ReportPath
        }
        
        # Configurar credenciais se fornecidas
        $credential = $null
        if ($emailUsername -and $emailPassword) {
            $securePassword = ConvertTo-SecureString $emailPassword -AsPlainText -Force
            $credential = New-Object System.Management.Automation.PSCredential($emailUsername, $securePassword)
        }
        
        $emailSuccess = Send-EmailNotification -To $emailTo `
                              -From $emailFrom `
                              -SmtpServer $smtpServer `
                              -Port $smtpPort `
                              -Subject $emailSubject `
                              -HtmlBody $emailBody `
                              -Attachments $Attachments `
                              -UseSSL $emailUseSSL `
                              -Credential $credential
    }
    
    # Enviar notifica��o Telegram se habilitado
    $telegramSuccess = $false
    if ($telegramEnabled) {
        $telegramSuccess = Send-TelegramNotification -BotToken $telegramBotToken `
                                 -ChatID $telegramChatID `
                                 -Message $message
    }
    
    # Retornar sucesso se pelo menos um m�todo funcionou
    return ($emailSuccess -or $telegramSuccess)
}

# Exportar as funcoes
Export-ModuleMember -Function Send-EmailNotification, Send-TelegramNotification, Send-ADRTNotification, Send-ADRTUpdateNotification