# coletor_cadastros_salesforce

Puxa os clientes (Account + CA_Endereco__c) do Salesforce e **adiciona** os CardCodes novos na tabela do
`PowerApp_Clientes.xlsx` (SharePoint sincronizado via OneDrive), que alimenta o Power Apps.
Nunca altera linhas existentes; faz backup `.bak-*.xlsx` antes de gravar.

## Servidor
1. `git clone` e copie `.env.example` para `.env` (preencha SF_* e XLSX_PATH).
2. Teste sem gravar: `python sf_clientes_para_excel.py --xlsx "<caminho>" --dry-run`
3. Agendar: `schtasks /Create /TN CadastrosSalesforce /SC DAILY /ST 07:30 /TR "cmd.exe /c \"<pasta>\run_coletor.bat\"" /RL HIGHEST`

Colunas: CardCode, CardName, Cod_Nome, MailZipCod, E_Mail, MailStrNo, Telefone, Tipo_de_Conta (CPF/CNPJ),
Password (o CPF/CNPJ formatado), __PowerAppsId__ (uuid4). Contas sem CPF/CNPJ ficam de fora.
