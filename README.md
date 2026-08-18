# Avaliacao de uso e licencas Microsoft 365

PowerShell para autenticar interativamente em um tenant, coletar usuarios, licencas e relatorios de uso do Microsoft Graph e gerar uma analise de reducao. O script **nao altera licencas**.

O relatorio identifica no preambulo a organizacao, Tenant ID, dominio padrao, usuario/conta que executou, data UTC, janela analisada e versao da ferramenta. A permissao `Organization.Read.All` e usada somente para ler esses dados de identificacao.

## Instalacao rapida

Abra o PowerShell 7 como administrador e execute:

```powershell
iwr -useb https://raw.githubusercontent.com/JulioVicente/o365-telemetria-uso-licencas/main/Install.ps1 | iex
```

O instalador baixa os componentes publicados, instala em `%ProgramData%\M365LicenseAssessment`, solicita o e-mail da conta que fará a autenticação e executa a primeira avaliação. O processo possui validação dos componentes e rollback local em caso de falha.

Para inspecionar ou simular o instalador antes da execução, use o procedimento auditável descrito no [QUICK_START.md](QUICK_START.md).

Depois da instalação, as próximas avaliações podem ser iniciadas com:

```powershell
& "$env:ProgramData\M365LicenseAssessment\Run-Assessment.ps1"
```

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

Na instalacao guiada, e solicitado somente o e-mail do usuario que concedera as permissoes e recebera o relatorio. Nao existe pergunta separada sobre remetente: o envio ocorre pela propria conta autenticada no Microsoft Graph. O instalador valida que o login corresponde ao usuario informado e somente conclui depois que o Graph aceitar o envio e a copia oculta para `suporte@bestsoft.com.br`.

O envio tambem fica ativo por padrao na execucao avulsa. O destinatario e sempre a conta autenticada; `EmailTo` nao redireciona o relatorio para outro usuario. Para gerar somente os arquivos locais:

```powershell
pwsh ./Invoke-M365LicenseAssessment.ps1 -SendEmail:$false
```

Para alterar excepcionalmente a copia oculta, use `-BccAddress outro-endereco@dominio.com`.

Cada execucao cria `output/<data-hora>/relatorio.html`, `analise-usuarios.csv` e `resumo.json`.

## Criterios e limites

- Login e cada carga de trabalho, incluindo Microsoft Teams, sao classificados em ativo ate 30 dias, inativo entre 31 e 90 dias, inativo acima de 90 dias ou sem atividade observada.
- O relatorio separa usuarios com um service plan Teams ativo e sem uso ha mais de 30/90 dias. A recomendacao distingue candidatura a revisao; Teams incluido em uma suite pode exigir migracao para um SKU sem Teams, enquanto licencas standalone ou complementos podem permitir retirada direta.
- O Microsoft 365 Copilot aparece por usuario com indicacao de licenca associada, uso observado, ultima atividade e inatividade em 30/90 dias. A coleta usa o relatorio oficial do Microsoft Graph e sinaliza quando ele estiver indisponivel.
- Uma conta com licenca que inclui Exchange e sem login ha mais de 90 dias e marcada como candidata a caixa compartilhada. Isso preserva um endereco funcional e pode permitir a retirada da licenca, mas requer validacao antes da mudanca.
- A recomendacao usa atividade observada em ate 180 dias. Sem atividade em todas as cargas, o usuario vira candidato a remocao, sujeito a validacao humana.
- O menor plano sugerido cobre apenas email, OneDrive, SharePoint, Office Web e Office Desktop observados. Seguranca, Intune, Defender, Teams/telefonia, Power BI, compliance, arquivo e retencao exigem revisao humana; por isso planos Premium/E5 nao sao sugeridos automaticamente.
- O limite de 300 usuarios dos planos Business deve ser considerado pelo administrador. Para tenants maiores, revise sugestoes Business.
- SKUs adicionais ou sem preco publico aparecem em `LicencasSemPrecoPublico` e nao entram na economia estimada.
- Caixa compartilhada sem licenca e limitada a 50 GB. Mais de 50 GB, arquivo, litigation hold e determinados recursos avancados de retencao/seguranca podem exigir licenca. O script apenas recomenda: nao converte a caixa, nao bloqueia a conta e nao remove licencas.
- Precos mudam e podem variar por impostos, canal, contrato e promocao. Atualize `config/license-catalog.pt-BR.json` antes de uma decisao comercial.

Fontes: [relatorios de uso do Microsoft Graph](https://learn.microsoft.com/graph/api/resources/report), [SKUs inscritos](https://learn.microsoft.com/graph/api/subscribedsku-list), [precos Business Brasil](https://www.microsoft.com/pt-br/microsoft-365/business/microsoft-365-plans-and-pricing) e [precos Office 365 Enterprise Brasil](https://www.microsoft.com/pt-br/microsoft-365/enterprise/office-365-plans-and-pricing).
