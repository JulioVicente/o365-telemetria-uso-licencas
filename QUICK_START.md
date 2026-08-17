# Guia rapido

Abra o PowerShell 7 como administrador. Baixe, inspecione, simule e execute o instalador:

```powershell
Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/JulioVicente/o365-telemetria-uso-licencas/main/Install.ps1' -OutFile "$env:TEMP\Install-M365LicenseAssessment.ps1"
Get-Content "$env:TEMP\Install-M365LicenseAssessment.ps1"
& "$env:TEMP\Install-M365LicenseAssessment.ps1" -WhatIf
& "$env:TEMP\Install-M365LicenseAssessment.ps1"
```

O destino padrao e `%ProgramData%\M365LicenseAssessment`. O instalador solicita somente o e-mail do usuario que concedera as permissoes e recebera o relatorio; nao solicita um remetente separado. No login, a conta autenticada deve corresponder ao e-mail informado. A BestSoft recebe a copia oculta. A instalacao somente termina depois que a coleta e o envio forem aceitos pelo Microsoft Graph.

Depois, execute quando desejar:

```powershell
$root = "$env:ProgramData\M365LicenseAssessment"
& "$root\Run-Assessment.ps1"
```

Na primeira conexao, autentique-se com a conta remetente informada e consinta as permissoes solicitadas, inclusive `Mail.Send`. Toda execucao instalada envia o relatorio ao destinatario e uma copia oculta para a BestSoft. Nenhuma licenca ou caixa postal e alterada.
