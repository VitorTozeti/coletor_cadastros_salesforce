"""Salesforce (Account) -> PowerApp_Clientes.xlsx na pasta do SharePoint sincronizada pelo OneDrive.

Roda no servidor do Caetano, 1x/dia. Só ADICIONA CardCodes que ainda não estão na tabela; linhas
existentes nunca são alteradas. Faz backup antes de gravar e reabre o resultado para validar antes de trocar o arquivo.

Uso:  python sf_clientes_para_excel.py --xlsx "<caminho local do PowerApp_Clientes.xlsx>" [--dry-run]

.env ao lado do script (nunca versionar): SF_CLIENT_ID, SF_CLIENT_SECRET, SF_LOGIN_URL.
"""
import argparse
import os
import re
import sys
import uuid
from pathlib import Path

import requests
import xlsx_append

AQUI = Path(__file__).parent
COLUNAS = ["CardCode", "CardName", "Cod_Nome", "MailZipCod", "E_Mail", "MailStrNo", "Telefone", "Tipo_de_Conta",
           "Password", "__PowerAppsId__"]
# Tipo_de_Conta = CPF|CNPJ; Password = o próprio CPF/CNPJ formatado (padrão da planilha); __PowerAppsId__ = uuid4.

SOQL_CONTAS = (
    "SELECT Id, Name, CA_CodigoSAP__c, CA_Email__c, CA_EmailEmpresa__c, CA_Telefone1__c, CA_Celular__c, Phone, CA_CPF__c, CA_CNPJ__c "
    "FROM Account WHERE CA_CodigoSAP__c != null"
)
# Endereço: mais recente ATIVO, preferindo entrega (S). CA_TipoEndereco__c guarda 'S'/'B', não os rótulos.
SOQL_END = (
    "SELECT CA_Conta__c, CA_CEP__c, CA_Numero__c, CA_TipoEndereco__c, CreatedDate "
    "FROM CA_Endereco__c WHERE CA_StatusEndereco__c = 'Ativo' AND CA_Conta__c != null"
)


def carrega_env(caminho=AQUI / ".env"):
    for linha in caminho.read_text(encoding="utf-8").splitlines():
        if "=" in linha and not linha.lstrip().startswith("#"):
            k, v = linha.split("=", 1)
            os.environ.setdefault(k.strip(), v.split("#")[0].strip())


def sf_query(soql):
    r = requests.post(
        os.environ["SF_LOGIN_URL"] + "/services/oauth2/token",
        data={"grant_type": "client_credentials", "client_id": os.environ["SF_CLIENT_ID"],
              "client_secret": os.environ["SF_CLIENT_SECRET"]}, timeout=60)
    r.raise_for_status()
    t = r.json()
    h = {"Authorization": "Bearer " + t["access_token"]}
    url, params = t["instance_url"] + "/services/data/v60.0/query", {"q": soql}
    while url:
        resp = requests.get(url, params=params, headers=h, timeout=120)
        if not resp.ok:
            sys.exit(f"Salesforce recusou a consulta ({resp.status_code}): {resp.text[:200]}")
        d = resp.json()
        yield from d["records"]
        url, params = (t["instance_url"] + d["nextRecordsUrl"], None) if not d["done"] else (None, None)


def melhor_endereco():
    """conta_id -> (cep, numero): ativo mais recente, entrega (S) antes de cobrança."""
    mapa = {}
    for e in sf_query(SOQL_END):
        chave = (e["CA_TipoEndereco__c"] == "S", e["CreatedDate"])
        atual = mapa.get(e["CA_Conta__c"])
        if atual is None or chave > atual[0]:
            mapa[e["CA_Conta__c"]] = (chave, (e["CA_CEP__c"] or "", e["CA_Numero__c"] or ""))
    return {k: v[1] for k, v in mapa.items()}


def documento(a):
    """(tipo, documento formatado) a partir de CA_CPF__c / CA_CNPJ__c; (None, None) se não tiver."""
    cpf, cnpj = re.sub(r"\D", "", a["CA_CPF__c"] or ""), re.sub(r"\D", "", a["CA_CNPJ__c"] or "")
    if len(cpf) == 11:
        return "CPF", f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
    if len(cnpj) == 14:
        return "CNPJ", f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
    return None, None


def linhas_salesforce():
    ends = melhor_endereco()
    for a in sf_query(SOQL_CONTAS):
        cep, num = ends.get(a["Id"], ("", ""))
        tipo, doc = documento(a)
        if not doc:  # sem CPF/CNPJ não há Password -> o cliente não conseguiria entrar no app; fica de fora
            continue
        cod, nome = a["CA_CodigoSAP__c"].strip(), (a["Name"] or "").strip()
        yield {
            "CardCode": cod.upper(), "CardName": nome, "Cod_Nome": f"{cod} - {nome}",
            "MailZipCod": cep, "E_Mail": a["CA_Email__c"] or a["CA_EmailEmpresa__c"] or "",
            "MailStrNo": num, "Telefone": a["CA_Telefone1__c"] or a["CA_Celular__c"] or a["Phone"] or "",
            "Tipo_de_Conta": tipo, "Password": doc, "__PowerAppsId__": str(uuid.uuid4()),
        }


LIMITE_NOVAS = 2000  # uso normal: dezenas por dia; um número alto indica tabela errada/vazia


LIMITE_NOVAS = 2000  # uso normal: dezenas por dia; um número alto indica tabela errada/vazia


def anexa_no_excel(xlsx, novas_fn, dry_run, forcar=False):
    """Grava via xlsx_append (só XML da aba/tabela; o resto do arquivo fica intacto — openpyxl NÃO é usado)."""
    return xlsx_append.anexa(xlsx, COLUNAS, novas_fn, dry_run, forcar, LIMITE_NOVAS)


if __name__ == "__main__":
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--forcar", action="store_true", help="ignora o limite de segurança de novos CardCodes")
    a = ap.parse_args()
    carrega_env()
    try:
        n = anexa_no_excel(a.xlsx, lambda: list(linhas_salesforce()), a.dry_run, a.forcar)
    except PermissionError:
        sys.exit("Excel em uso/bloqueado pelo OneDrive — tente de novo mais tarde. Nada foi gravado.")
    print("dry-run: nada gravado" if a.dry_run else f"{n} linhas adicionadas")
