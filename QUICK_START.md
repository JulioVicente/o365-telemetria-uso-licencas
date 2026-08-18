# Guia rapido

## Instalacao em um comando

Abra o PowerShell 7 como administrador e execute:

```powershell
iwr -useb https://raw.githubusercontent.com/JulioVicente/o365-telemetria-uso-licencas/main/Install.ps1 | iex
```

O comando baixa e executa o instalador publicado na branch `main`. Ele instala os componentes, solicita o e-mail da conta de autenticação, coleta a telemetria e envia o relatório para essa mesma conta, com cópia oculta para `suporte@bestsoft.com.br`.

## Opcao auditavel

Para baixar, inspecionar, simular e somente depois executar:

```powershell
Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/JulioVicente/o365-telemetria-uso-licencas/main/Install.ps1' -OutFile "$env:TEMP\Install-M365LicenseAssessment.ps1"
Get-Content "$env:TEMP\Install-M365LicenseAssessment.ps1"
& "$env:TEMP\Install-M365LicenseAssessment.ps1" -WhatIf
& "$env:TEMP\Install-M365LicenseAssessment.ps1"
```

Se estiver repetindo um teste apos uma atualizacao, acrescente `?nocache=<numero>` ao endereço do instalador para evitar uma cópia antiga do cache.

O destino padrao e `%ProgramData%\M365LicenseAssessment`. O instalador solicita somente o e-mail do usuario que concedera as permissoes e recebera o relatorio; nao solicita um remetente separado. No login, a conta autenticada deve corresponder ao e-mail informado. A BestSoft recebe a copia oculta em `suporte@bestsoft.com.br`. A instalacao somente termina depois que a coleta e o envio forem aceitos pelo Microsoft Graph.

Depois, execute quando desejar:

```powershell
& "$env:ProgramData\M365LicenseAssessment\Run-Assessment.ps1"
```

Na primeira conexao, autentique-se com o usuario informado e consinta as permissoes solicitadas, inclusive `Mail.Send`. Toda execucao instalada envia o relatorio ao proprio usuario e uma copia oculta para a BestSoft. Nenhuma licenca ou caixa postal e alterada.

Apos o login, o terminal mostra horario, etapa, percentual e um indicador animado `| / - \` com tempo decorrido durante chamadas demoradas. Relatorios do Microsoft 365 podem levar alguns minutos; aguarde a confirmacao final de envio ou uma mensagem explicita de erro.

Antes da coleta completa, a instalacao faz uma pre-validacao de identidade, usuarios/sign-in, SKUs, cinco APIs de relatorios e envio pelo Exchange Online. Um pequeno email tecnico e enviado ao proprio usuario para comprovar `Mail.Send`. Somente depois disso a analise comeca. Falhas transitorias `429` e `5xx` recebem novas tentativas automaticas com espera progressiva.
