# Avaliacao de uso e licencas Microsoft 365

PowerShell para autenticar interativamente em um tenant, coletar usuarios, licencas e relatorios de uso do Microsoft Graph e gerar uma analise de reducao. O script **nao altera licencas**.

## Instalacao guiada

O processo segue o mesmo padrao do projeto `sharepoint-version-cleanup`: instalador auditavel, suporte a `-WhatIf`, instalacao em `%ProgramData%`, configuracao interativa e rollback local em caso de falha.

```powershell
Invoke-WebRequest -Uri 'https://raw.githubusercontent.com/JulioVicente/o365-telemetria-uso-licencas/main/Install.ps1' -OutFile "$env:TEMP\Install-M365LicenseAssessment.ps1"
Get-Content "$env:TEMP\Install-M365LicenseAssessment.ps1"
& "$env:TEMP\Install-M365LicenseAssessment.ps1" -WhatIf
& "$env:TEMP\Install-M365LicenseAssessment.ps1"
```

Veja [QUICK_START.md](QUICK_START.md) para a operacao depois da instalacao.

## Requisitos

- PowerShell 7.2 ou superior.
- Uma conta corporativa com acesso aos relatorios e leitura do diretorio.
- Modulo `Microsoft.Graph.Authentication`.

```powershell
Install-Module Microsoft.Graph.Authentication -Scope CurrentUser
```

Na primeira execucao, o administrador deve consentir com `User.Read.All`, `AuditLog.Read.All`, `LicenseAssignment.Read.All` e `Reports.Read.All`. Para envio, tambem e solicitado `Mail.Send`. A disponibilidade de `signInActivity` depende da licenca/role do tenant. O Graph pode anonimizar nomes nos relatorios; desative essa opcao no Centro de Administracao do Microsoft 365 para correlacao por UPN.

## Executar

```powershell
pwsh ./Invoke-M365LicenseAssessment.ps1
```

Na instalacao guiada, e solicitado somente o e-mail do usuario que concedera as permissoes e recebera o relatorio. Nao existe pergunta separada sobre remetente: o envio ocorre pela propria conta autenticada no Microsoft Graph. O instalador valida que o login corresponde ao usuario informado e somente conclui depois que o Graph aceitar o envio e a copia oculta para `suprote@bestsoft.com.br`.

Para uma execucao avulsa do script principal:

```powershell
pwsh ./Invoke-M365LicenseAssessment.ps1 -SendEmail -EmailTo gestor@contoso.com
```

Caso o endereco pretendido seja `suporte@bestsoft.com.br`, use `-BccAddress suporte@bestsoft.com.br`.

Cada execucao cria `output/<data-hora>/relatorio.html`, `analise-usuarios.csv` e `resumo.json`.

## Criterios e limites

- Login e cada carga de trabalho sao classificados em ativo ate 30 dias, inativo entre 31 e 90 dias, inativo acima de 90 dias ou sem atividade observada.
- Uma conta com licenca que inclui Exchange e sem login ha mais de 90 dias e marcada como candidata a caixa compartilhada. Isso preserva um endereco funcional e pode permitir a retirada da licenca, mas requer validacao antes da mudanca.
- A recomendacao usa atividade observada em ate 180 dias. Sem atividade em todas as cargas, o usuario vira candidato a remocao, sujeito a validacao humana.
- O menor plano sugerido cobre apenas email, OneDrive, SharePoint, Office Web e Office Desktop observados. Seguranca, Intune, Defender, Teams/telefonia, Power BI, compliance, arquivo e retencao exigem revisao humana; por isso planos Premium/E5 nao sao sugeridos automaticamente.
- O limite de 300 usuarios dos planos Business deve ser considerado pelo administrador. Para tenants maiores, revise sugestoes Business.
- SKUs adicionais ou sem preco publico aparecem em `LicencasSemPrecoPublico` e nao entram na economia estimada.
- Caixa compartilhada sem licenca e limitada a 50 GB. Mais de 50 GB, arquivo, litigation hold e determinados recursos avancados de retencao/seguranca podem exigir licenca. O script apenas recomenda: nao converte a caixa, nao bloqueia a conta e nao remove licencas.
- Precos mudam e podem variar por impostos, canal, contrato e promocao. Atualize `config/license-catalog.pt-BR.json` antes de uma decisao comercial.

Fontes: [relatorios de uso do Microsoft Graph](https://learn.microsoft.com/graph/api/resources/report), [SKUs inscritos](https://learn.microsoft.com/graph/api/subscribedsku-list), [precos Business Brasil](https://www.microsoft.com/pt-br/microsoft-365/business/microsoft-365-plans-and-pricing) e [precos Office 365 Enterprise Brasil](https://www.microsoft.com/pt-br/microsoft-365/enterprise/office-365-plans-and-pricing).
