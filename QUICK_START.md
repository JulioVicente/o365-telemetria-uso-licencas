# Guia rapido

Abra o PowerShell 7 como administrador. Baixe, inspecione, simule e execute o instalador:

```powershell
Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/JulioVicente/o365-telemetria-uso-licencas/main/Install.ps1' -OutFile "$env:TEMP\Install-M365LicenseAssessment.ps1"
Get-Content "$env:TEMP\Install-M365LicenseAssessment.ps1"
& "$env:TEMP\Install-M365LicenseAssessment.ps1" -WhatIf
& "$env:TEMP\Install-M365LicenseAssessment.ps1"
```

O destino padrao e `%ProgramData%\M365LicenseAssessment`. O instalador instala o modulo Graph e abre o login. A conta autenticada envia e recebe o proprio relatorio automaticamente; a BestSoft recebe a copia oculta. Nenhum endereco precisa ser digitado no instalador. A instalacao somente termina depois que a coleta e o envio forem aceitos pelo Microsoft Graph.

Depois, execute quando desejar:

```powershell
$root = "$env:ProgramData\M365LicenseAssessment"
& "$root\Run-Assessment.ps1"
```

Na primeira conexao, autentique-se com a conta remetente informada e consinta as permissoes solicitadas, inclusive `Mail.Send`. Toda execucao instalada envia o relatorio ao destinatario e uma copia oculta para a BestSoft. Nenhuma licenca ou caixa postal e alterada.
