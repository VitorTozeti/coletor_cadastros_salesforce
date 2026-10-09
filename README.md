# coletor_cadastros_salesforce

Puxa os clientes (Account + CA_Endereco__c) do Salesforce e **adiciona** os CardCodes novos na tabela `Consulta1`
do `PowerApp_Clientes.xlsx` (SharePoint "Site da Equipe > Documentos", sincronizado pelo OneDrive no servidor),
que alimenta o app **Contratos** (Power Apps). Nunca altera linhas existentes.

## Servidor
1. `git clone` e copie `.env.example` para `.env` (preencha SF_* e XLSX_PATH = caminho COMPLETO do arquivo `.xlsx`, não a pasta).
2. Teste sem gravar: `python sf_clientes_para_excel.py --xlsx "<caminho>" --dry-run`
3. Agendar: `schtasks /Create /TN CadastrosSalesforce /SC DAILY /ST 06:30 /TR "cmd.exe /c \"<pasta>\run_coletor.bat\"" /RL HIGHEST`
   (horário sem ninguém com o Excel aberto).

Colunas: CardCode, CardName, Cod_Nome, MailZipCod, E_Mail, MailStrNo, Telefone, Tipo_de_Conta (CPF/CNPJ),
Password (o CPF/CNPJ formatado), __PowerAppsId__ (uuid4).

## Quem entra no Excel (e por que um cliente "não aparece")
1. precisa de **CardCode** (`CA_CodigoSAP__c`) — só existe depois que a integração Salesforce → SAP conclui;
2. precisa de **CPF ou CNPJ** válido;
3. só entra na **próxima rodada** do script.

Diagnóstico de um cliente: `python sf_clientes_para_excel.py --xlsx "<caminho>" --diagnostico C0031886` (ou parte do nome).
Toda rodada imprime quantas contas ficaram de fora e por quê.

## Como a gravação é segura (incidente de 06/10/2026)
- **Não usa openpyxl.** O openpyxl reconstruía o arquivo e apagava a consulta do Power Query
  ("Intervalo de dados externos") → o Excel pedia reparo e o app perdeu a tabela. Agora `xlsx_append.py` muda só
  o XML da aba (linhas novas) e da tabela (faixa); todo o resto é copiado byte a byte.
- Recusa planilha com consulta/conexão externa (as linhas sumiriam na atualização da consulta).
- Recusa se não houver tabela, se faltar coluna ou se houver conteúdo abaixo da tabela.
- Aborta se vierem mais de 2000 novos (tabela errada/vazia); `--forcar` só se for esperado.
- Backup `.bak-*.xlsx` antes; grava em `.tmp`, reabre e valida, e só então troca (atômico). Arquivo aberto no
  Excel → sai sem gravar e sem deixar lixo.

## Se o app quebrar de novo (tabela sumiu / erro 404 "Nenhuma tabela foi encontrada")
1. Abra o `PowerApp_Clientes.xlsx` no Excel Online. Sem cabeçalho com filtros = sem tabela:
   selecione `A1:J<última linha>` → Inserir → Tabela → marque **"Minha tabela tem cabeçalhos"** → OK →
   Design da Tabela → nome **`Consulta1`**.
2. No Studio do Contratos: Dados → remova `Consulta1` (se estiver com erro) → Adicionar dados → Excel Online
   (Business) → Site da Equipe → Documentos → `PowerApp_Clientes.xlsx` → tabela `Consulta1` → identificador
   "ID gerada automaticamente" → Conectar. **Salvar e Publicar.**
3. O Power Apps guarda a tabela pelo ID interno: tabela recriada = fonte tem de ser re-adicionada.

## Limite conhecido do app (não é do script)
O combo "Código/Nome SAP" usa `Filter(Consulta1, Tipo_de_Conta <> "CNPJ")`. **Excel não é delegável**: o app só
carrega as primeiras linhas da tabela (limite de linhas de dados do app, máx. 2000). Clientes além dessa linha —
inclusive os novos, que o script adiciona no fim — podem não aparecer no combo. Solução definitiva: trocar a fonte
para uma lista do SharePoint/Dataverse (delegável).

## Arquitetura atual (07/10/2026) — o app NÃO lê mais o Excel
O combo de clientes do app **Contratos** lê a **lista SharePoint `Clientes_PowerApp`** (busca por nome, delegável; o
Excel não era delegável e o app só enxergava as 500 primeiras linhas). O fluxo é:

1. **Este coletor** (servidor, 06:30) → acrescenta CardCodes novos em `PowerApp_Clientes.xlsx` (histórico/backup) **e**
   recria `PowerApp_Clientes_Novos.xlsx` (tabela `Novos`) com os clientes dos últimos 7 dias.
2. **Power Automate** (07:00) lê a tabela `Novos` e cria os itens na lista (`CardCode` é único: repetido = ignorado).
3. O app lê a lista.

Proteções do coletor: nunca usa openpyxl no arquivo principal; recusa Power Query/sem tabela/coluna faltando; backup
+ arquivo temporário + validação + troca atômica; trava `coletor.lock` (sem rodada dupla); 4 tentativas (2 min) se o
OneDrive segurar o arquivo; estado `novos_estado.json` gravado ANTES da escrita (nada se perde se falhar no meio);
sai com código ≠ 0 em qualquer falha (aparece no agendador e no `coletor.log`).
Limitação conhecida: cliente já existente que mudar nome/e-mail no Salesforce não é atualizado (só entram CardCodes novos).

## Endereço completo e vínculo PF↔PJ (branch `enderecos-completos`, 09/10/2026)
O Excel principal (`Consulta1`) e o fluxo atual **não mudam**. Cada rodada também recria, do zero e de forma atômica,
`PowerApp_Clientes_Enderecos.xlsx` (tabela `Enderecos`) ao lado do principal, com **todas as contas aptas**:
`CardCode, Logradouro (tipo+nome), Numero, Complemento, Bairro, Cidade, UF, CEP, CardCode_Vinculado, Nome_Vinculado`
(endereço = ativo mais recente de `CA_Endereco__c`, entrega antes de cobrança). Se esse arquivo falhar, a rodada principal
não é afetada (só loga AVISO).
- **Vínculo PF→PJ:** descubra o campo com `python sf_clientes_para_excel.py --xlsx "<caminho>" --descrever` (lista os
  lookups Account→Account e quantos estão preenchidos) e ponha no `.env`: `SF_CAMPO_VINCULO=<campo>` (ex. `ParentId`).
  Sem a variável, as colunas `*_Vinculado` ficam em branco.
- **Para o app usar:** (1) criar na lista `Clientes_PowerApp` as colunas novas (só adicionar; as atuais ficam);
  (2) um fluxo à parte lê `PowerApp_Clientes_Enderecos.xlsx`/`Enderecos` e atualiza os itens por `CardCode`;
  (3) só então ler as colunas no app. Nada disso é necessário para o app continuar funcionando como hoje.
- `--dry-run` calcula e mostra as contagens sem gravar o arquivo.
