# Importar modulos do ActiveDirectory e ExchangeOnlineManagement
Import-Module ActiveDirectory
Import-Module ExchangeOnlineManagement

# Pesquisa os usuários do AD habilitados
$users = Get-ADUser -Filter "Enabled -eq 'true'" -Properties passwordlastset,mail

# Define variáveis
$dataAtual = Get-Date

# Conecta no 365 para envio de emails
$pathsenha365 = "C:\ExpiraSenha\o365clear.xml"
[xml]$cred = Get-Content $pathsenha365
$smtpUsername = $cred.credentials.username
$smtpPassword = ConvertTo-SecureString $cred.credentials.password -AsPlainText -Force
$smtpCred = New-Object System.Management.Automation.PSCredential ($smtpUsername, $smtpPassword)

Connect-ExchangeOnline -UserPrincipalName 'SuaCaixaDeEmail' -ShowProgress $true -Credential $smtpCred
#$cred = Import-Clixml -Path "C:\ExpiraSenha\o365clear.xml"

$assunto = "Expiração de senha"
$enc = [System.Text.Encoding]::UTF8

# Para cada usuário habilitado do Active Directory
foreach ($user in $users) {
    $nomeUser = $user.Name
    $login = $user.SamAccountName
    $destinatario = $user.mail
    $dataExpiracaoSenha = $user.PasswordLastSet + (Get-ADDefaultDomainPasswordPolicy).maxpasswordage
    $diasRestantes = ($dataExpiracaoSenha - $dataAtual).Days

    # Verifica se faltam 14, 7, 6, 5, 4, 3, 2 ou 1 dia para a expiração
    if ($diasRestantes -eq 14 -or $diasRestantes -eq 7 -or $diasRestantes -eq 6 -or $diasRestantes -eq 5 -or $diasRestantes -eq 4 -or $diasRestantes -eq 3 -or $diasRestantes -eq 2 -or $diasRestantes -eq 1) {
        $corpoHtml = @"
        <html>
        <head>
            <meta charset="UTF-8">
            <style>
                body {
                    font-family: Arial, Helvetica, sans-serif;
                    color: black; /* Cor do texto principal */
                }
                h1 {
                    color: #EB1946; /* Cor vermelha para o título  */
                }
                .red-text {
                    color: #EB1946; /* Cor vermelha para palavras-chave */
                }
                .blue-text {
                    color: #0073e6; /* Cor azul para palavras-chave */
                }
                .container {
                    max-width: 600px;
                    margin: 0 auto;
                    padding: 20px;
                }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>Nome da empresa</h1>
                <h1>Olá <span class="blue-text"></span></h1>
                <p>A data de expiração de sua senha <span class="red-text"> NomeDaEmpresa</span> está programada para ocorrer em um prazo de <span class="red-text">$diasRestantes dias</span>. Recomendamos que esteja atento a esta data, a fim de evitar a perda de acesso aos ambientes dos sistemas <span class="red-text">NomeDaEmpresa</span>.</p>
                <p>Login: <span class="red-text">$login</span></p>
                <p>Lembre-se ao redefinir sua senha:</p>
                <ul>
                    <li>A senha diferencia maiúsculas de minúsculas;</li>
                    <li>Deve ter no mínimo 12 caracteres;</li>
                    <li>Não deve incluir os seguintes valores: password test;</li>
                    <li>Não deve incluir parte do seu nome ou nome de usuário;</li>
                    <li>Não deve incluir uma palavra comum ou sequência de caracteres comumente usada;</li>
                    <li>Não deve incluir nenhuma das últimas 2 senhas utilizadas;</li>
                    <li>Requer no mínimo 1 caractere maiúsculo;</li>
                    <li>Requer no mínimo 1 número;</li>
                    <li>Requer no mínimo 1 caractere especial (por exemplo: !, $, #, @).</li>
                </ul>
                <p>A senha expira a cada 90 dias. Enviaremos uma nova mensagem faltando 14 dias da expiração.</p>
                <p>Acesse esse site para redefinir sua senha, copie essa URL e cole no seu navegador: </span>
                <p></p>
                <span class="red-text"> </span></p>
                <p><span class="blue-text">Lembrando, essa senha se refere ao acesso WTS, login utilizado para acessar aos ambientes, não sistemas. </span></p>
            </div>
        </body>
        </html>
"@
        
        # Envie o e-mail
        Send-MailMessage -From "SuaCaixaDeEmail" -To $destinatario -Subject $assunto -Body $corpoHtml -BodyAsHtml -SmtpServer "smtp.office365.com" -UseSsl -Port 587 -Credential $smtpCred -Encoding UTF8
    }
}
