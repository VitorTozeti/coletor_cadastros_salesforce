"""Salesforce (Account) -> PowerApp_Clientes.xlsx na pasta do SharePoint sincronizada pelo OneDrive.

Roda no servidor do Caetano, 1x/dia. Só ADICIONA CardCodes que ainda não estão na tabela; linhas
existentes nunca são alteradas. Faz backup antes de gravar e reabre o resultado para validar antes de trocar o arquivo.

Uso:
  python sf_clientes_para_excel.py --xlsx "<caminho do PowerApp_Clientes.xlsx>" [--dry-run]
  python sf_clientes_para_excel.py --xlsx "<caminho>" --diagnostico C0031886      (ou parte do nome)

Regra de entrada (por que um cliente NÃO aparece no app):
  1. precisa de CardCode (CA_CodigoSAP__c) -> só existe depois que a integração SF -> SAP conclui;
  2. precisa de CPF ou CNPJ válido (vira Tipo_de_Conta e Password);
  3. só entra na próxima rodada do script (agendado).

.env ao lado do script (nunca versionar): SF_CLIENT_ID, SF_CLIENT_SECRET, SF_LOGIN_URL.
"""
import argparse
import json
import os
import time
import re
import sys
import uuid
import zipfile
from collections import Counter
from datetime import datetime, timedelta
from pathlib import Path

import requests
import novos_xlsx
import xlsx_append

AQUI = Path(__file__).parent
COLUNAS = ["CardCode", "CardName", "Cod_Nome", "MailZipCod", "E_Mail", "MailStrNo", "Telefone", "Tipo_de_Conta",
           "Password", "__PowerAppsId__"]
# Tipo_de_Conta = CPF|CNPJ; Password = o próprio CPF/CNPJ formatado (padrão da planilha); __PowerAppsId__ = uuid4.
DIAS_NOVOS = 7  # quantos dias um cliente novo fica no arquivo _Novos (o fluxo ignora CardCode repetido)
LIMITE_NOVAS = 2000  # uso normal: dezenas por dia; um número alto indica tabela errada/vazia

SOQL_CONTAS = (
    "SELECT Id, Name, CA_CodigoSAP__c, CA_Email__c, CA_EmailEmpresa__c, CA_Telefone1__c, CA_Celular__c, Phone, "
    "CA_CPF__c, CA_CNPJ__c, CA_StatusIntegracao__c FROM Account"
)
# Endereço: mais recente ATIVO, preferindo entrega (S). CA_TipoEndereco__c guarda 'S'/'B', não os rótulos.
SOQL_END = (
    "SELECT CA_Conta__c, CA_CEP__c, CA_Numero__c, CA_TipoEndereco__c, CreatedDate, CA_TipoLogradouro__c, "
    "CA_Logradouro__c, CA_Complemento__c, CA_Bairro__c, CA_NomeMunicipio__c, CA_SiglaEstado__c "
    "FROM CA_Endereco__c WHERE CA_StatusEndereco__c = 'Ativo' AND CA_Conta__c != null"
)


def log(msg):
    print(f"[{datetime.now():%d/%m/%Y %H:%M:%S}] {msg}", flush=True)


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


def melhor_endereco_completo():
    """conta_id -> registro CA_Endereco__c: ativo mais recente, entrega (S) antes de cobrança."""
    mapa = {}
    for e in sf_query(SOQL_END):
        chave = (e["CA_TipoEndereco__c"] == "S", e["CreatedDate"])
        atual = mapa.get(e["CA_Conta__c"])
        if atual is None or chave > atual[0]:
            mapa[e["CA_Conta__c"]] = (chave, e)
    return {k: v[1] for k, v in mapa.items()}


def melhor_endereco():
    """conta_id -> (cep, numero) — o que vai para o Excel principal (colunas inalteradas)."""
    return {k: (e["CA_CEP__c"] or "", e["CA_Numero__c"] or "") for k, e in melhor_endereco_completo().items()}


def documento(a):
    """(tipo, documento formatado) a partir de CA_CPF__c / CA_CNPJ__c; (None, None) se não tiver."""
    cpf, cnpj = re.sub(r"\D", "", a["CA_CPF__c"] or ""), re.sub(r"\D", "", a["CA_CNPJ__c"] or "")
    if len(cpf) == 11:
        return "CPF", f"{cpf[:3]}.{cpf[3:6]}.{cpf[6:9]}-{cpf[9:]}"
    if len(cnpj) == 14:
        return "CNPJ", f"{cnpj[:2]}.{cnpj[2:5]}.{cnpj[5:8]}/{cnpj[8:12]}-{cnpj[12:]}"
    return None, None


def motivo_fora(a):
    """None se a conta entra no Excel; senão o motivo legível."""
    if not (a["CA_CodigoSAP__c"] or "").strip():
        return f"sem CardCode (integração SAP: {a.get('CA_StatusIntegracao__c') or 'não iniciada'})"
    if not documento(a)[1]:
        return "sem CPF/CNPJ válido"
    return None


def linhas_salesforce(contas=None, resumo=None):
    ends = melhor_endereco()
    for a in contas if contas is not None else sf_query(SOQL_CONTAS):
        motivo = motivo_fora(a)
        if resumo is not None:
            resumo[motivo or "entra"] += 1
        if motivo:
            continue
        cep, num = ends.get(a["Id"], ("", ""))
        tipo, doc = documento(a)
        cod, nome = a["CA_CodigoSAP__c"].strip(), (a["Name"] or "").strip()
        yield {
            "CardCode": cod.upper(), "CardName": nome, "Cod_Nome": f"{cod} - {nome}",
            "MailZipCod": cep, "E_Mail": a["CA_Email__c"] or a["CA_EmailEmpresa__c"] or "",
            "MailStrNo": num, "Telefone": a["CA_Telefone1__c"] or a["CA_Celular__c"] or a["Phone"] or "",
            "Tipo_de_Conta": tipo, "Password": doc, "__PowerAppsId__": str(uuid.uuid4()),
        }


# Arquivo de endereços/vínculo: recriado do zero a cada rodada (todas as contas aptas), SEM tocar no Excel principal
# nem na tabela Consulta1 (o app e o fluxo atual continuam iguais). Quem lê é um fluxo/consulta à parte.
COLUNAS_END = ["CardCode", "Logradouro", "Numero", "Complemento", "Bairro", "Cidade", "UF", "CEP",
               "CardCode_Vinculado", "Nome_Vinculado"]
# SF_CAMPO_VINCULO (.env) = campo do Account que liga uma conta PF a uma PJ (lookup). Vazio = colunas *_Vinculado
# ficam em branco. Descobrir o campo com --descrever; ex.: ParentId ou um CA_*__c customizado.


def campo_vinculo():
    c = os.environ.get("SF_CAMPO_VINCULO", "").strip()
    if c and not re.fullmatch(r"[A-Za-z0-9_]+", c):
        sys.exit(f"SF_CAMPO_VINCULO inválido: {c!r}")
    return c


def linhas_enderecos():
    campo = campo_vinculo()
    soql = SOQL_CONTAS.replace(" FROM Account", f", {campo} FROM Account") if campo else SOQL_CONTAS
    contas = list(sf_query(soql))
    por_id = {a["Id"]: a for a in contas}
    ends = melhor_endereco_completo()
    for a in contas:
        if motivo_fora(a):
            continue
        e = ends.get(a["Id"], {})
        vinc = por_id.get(a.get(campo)) if campo and a.get(campo) else None
        logr = " ".join(x for x in ((e.get("CA_TipoLogradouro__c") or "").strip(),
                                    (e.get("CA_Logradouro__c") or "").strip()) if x)
        yield {
            "CardCode": a["CA_CodigoSAP__c"].strip().upper(), "Logradouro": logr, "Numero": e.get("CA_Numero__c") or "",
            "Complemento": e.get("CA_Complemento__c") or "", "Bairro": e.get("CA_Bairro__c") or "",
            "Cidade": e.get("CA_NomeMunicipio__c") or "", "UF": e.get("CA_SiglaEstado__c") or "",
            "CEP": e.get("CA_CEP__c") or "",
            "CardCode_Vinculado": ((vinc or {}).get("CA_CodigoSAP__c") or "").strip().upper(),
            "Nome_Vinculado": (vinc or {}).get("Name") or "",
        }


def grava_enderecos(xlsx, dry_run=False):
    """Recria PowerApp_Clientes_Enderecos.xlsx (tabela `Enderecos`) ao lado do principal."""
    linhas = list(linhas_enderecos())
    log(f"endereços: {len(linhas)} contas aptas; com logradouro={sum(1 for l in linhas if l['Logradouro'])}; "
        f"com vínculo={sum(1 for l in linhas if l['CardCode_Vinculado'])}")
    if dry_run:
        return
    destino = Path(xlsx).with_name(Path(xlsx).stem + "_Enderecos.xlsx")
    novos_xlsx.grava(destino, linhas, COLUNAS_END, "Enderecos")
    log(f"{destino.name} gravado")


def descrever():
    """Lista os campos de Account que são lookup para Account (candidatos a ligar PF -> PJ) e quantos usam cada um."""
    r = requests.post(os.environ["SF_LOGIN_URL"] + "/services/oauth2/token", data={
        "grant_type": "client_credentials", "client_id": os.environ["SF_CLIENT_ID"],
        "client_secret": os.environ["SF_CLIENT_SECRET"]}, timeout=60)
    r.raise_for_status()
    t = r.json()
    d = requests.get(t["instance_url"] + "/services/data/v60.0/sobjects/Account/describe",
                     headers={"Authorization": "Bearer " + t["access_token"]}, timeout=60)
    d.raise_for_status()
    for f in d.json()["fields"]:
        if f["type"] == "reference" and f["referenceTo"] == ["Account"]:
            n = next(sf_query(f"SELECT COUNT(Id) n FROM Account WHERE {f['name']} != null"))["n"]
            print(f"{f['name']:<35} {f['label']:<40} preenchidos={n}")


ESTADO = AQUI / "novos_estado.json"  # nunca versionar (tem dados de clientes)
TRAVA = AQUI / "coletor.lock"


def _agora():
    return datetime.now().strftime("%Y-%m-%d")


def registra_novos(novas):
    """Guarda os clientes desta rodada (por data) para o arquivo _Novos; descarta os mais velhos que DIAS_NOVOS."""
    estado = json.loads(ESTADO.read_text(encoding="utf-8")) if ESTADO.exists() else {}
    for l in novas:
        estado[l["CardCode"]] = {"em": _agora(), "linha": {c: l.get(c, "") for c in novos_xlsx.COLUNAS}}
    ESTADO.write_text(json.dumps(estado, ensure_ascii=False), encoding="utf-8")


def grava_novos(xlsx):
    """Recria PowerApp_Clientes_Novos.xlsx (ao lado do principal) a partir do estado. Roda SEMPRE, mesmo sem novos."""
    estado = json.loads(ESTADO.read_text(encoding="utf-8")) if ESTADO.exists() else {}
    corte = (datetime.now() - timedelta(days=DIAS_NOVOS)).strftime("%Y-%m-%d")
    estado = {k: v for k, v in estado.items() if v["em"] >= corte}
    ESTADO.write_text(json.dumps(estado, ensure_ascii=False), encoding="utf-8")
    destino = Path(xlsx).with_name(Path(xlsx).stem + "_Novos.xlsx")
    novos_xlsx.grava(destino, [v["linha"] for v in estado.values()])
    log(f"{destino.name}: {len(estado)} clientes dos últimos {DIAS_NOVOS} dias (o fluxo lê a tabela Novos)")


def com_tentativas(fn, vezes=4, espera=120):
    """Arquivo em uso/bloqueado pelo OneDrive costuma liberar em minutos: tenta de novo antes de desistir."""
    for i in range(1, vezes + 1):
        try:
            return fn()
        except PermissionError:
            if i == vezes:
                raise
            log(f"arquivo em uso/bloqueado (tentativa {i}/{vezes}); nova tentativa em {espera}s")
            time.sleep(espera)


def pega_trava():
    """Impede duas rodadas ao mesmo tempo (agendador + execução manual). Trava velha (>3h) é ignorada."""
    if TRAVA.exists() and time.time() - TRAVA.stat().st_mtime < 3 * 3600:
        sys.exit("Já existe uma rodada em andamento (coletor.lock). Nada feito.")
    TRAVA.write_text(str(os.getpid()), encoding="utf-8")


def anexa_no_excel(xlsx, novas_fn, dry_run, forcar=False):
    """Grava via xlsx_append (só XML da aba/tabela; o resto do arquivo fica intacto — openpyxl NÃO é usado)."""
    return xlsx_append.anexa(xlsx, COLUNAS, novas_fn, dry_run, forcar, LIMITE_NOVAS, antes_de_gravar=registra_novos)


def cardcodes_no_excel(xlsx):
    with zipfile.ZipFile(xlsx) as z:
        aba, tabela = xlsx_append._localiza_tabela(z)
        if not aba:
            return set()
        tab = z.read(tabela).decode("utf-8")
        m = re.search(r'<table\b[^>]*?\sref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', tab)
        cab = re.findall(r'<tableColumn\b[^>]*\sname="([^"]*)"', tab)
        letra = xlsx_append._letra(xlsx_append._num(m.group(1)) + cab.index("CardCode"))
        return xlsx_append._valores_da_coluna(z, aba, letra, int(m.group(2)) + 1, int(m.group(4)))


def diagnostico(xlsx, termo):
    """Explica por que um cliente aparece ou não no app: Salesforce -> regra de entrada -> Excel."""
    termo_n = termo.strip().upper()
    contas = [a for a in sf_query(SOQL_CONTAS)
              if termo_n in (a["CA_CodigoSAP__c"] or "").upper() or termo_n in (a["Name"] or "").upper()]
    if not contas:
        print(f"'{termo}': não encontrado no Salesforce (Account). Cadastro ainda não existe ou nome/código diferente.")
        return
    no_excel = cardcodes_no_excel(xlsx)
    for a in contas[:20]:
        cod = (a["CA_CodigoSAP__c"] or "").strip().upper()
        motivo = motivo_fora(a)
        if motivo:
            estado = f"NÃO entra no Excel: {motivo}"
        elif cod in no_excel:
            estado = "JÁ está no Excel (se não aparece no app: publicar o app / atualizar a fonte Consulta1)"
        else:
            estado = "entra na PRÓXIMA rodada do script (ainda não está no Excel)"
        print(f"- {a['Name']} | CardCode={cod or '-'} -> {estado}")
    if len(contas) > 20:
        print(f"... e mais {len(contas) - 20} contas com '{termo}'")


def main(argv=None):
    ap = argparse.ArgumentParser()
    ap.add_argument("--xlsx", required=True)
    ap.add_argument("--dry-run", action="store_true")
    ap.add_argument("--forcar", action="store_true", help="ignora o limite de segurança de novos CardCodes")
    ap.add_argument("--descrever", action="store_true",
                    help="lista lookups Account->Account (candidatos ao vínculo PF/PJ) e quantos estão preenchidos")
    ap.add_argument("--diagnostico", metavar="CARDCODE_OU_NOME", help="explica por que um cliente aparece ou não")
    a = ap.parse_args(argv)
    carrega_env()
    if a.descrever:
        descrever()
        sys.exit(0)
    if a.diagnostico:
        diagnostico(a.xlsx, a.diagnostico)
        sys.exit(0)
    resumo = Counter()
    pega_trava()
    try:
        n = com_tentativas(lambda: anexa_no_excel(a.xlsx, lambda: list(linhas_salesforce(resumo=resumo)), a.dry_run,
                                                  a.forcar))
        if not a.dry_run:
            com_tentativas(lambda: grava_novos(a.xlsx))
    except PermissionError:
        sys.exit("Excel em uso/bloqueado pelo OneDrive mesmo após várias tentativas. Nada foi gravado.")
    finally:
        TRAVA.unlink(missing_ok=True)
    try:  # arquivo à parte: se falhar não derruba a rodada principal (o Excel principal já foi tratado acima)
        com_tentativas(lambda: grava_enderecos(a.xlsx, a.dry_run))
    except Exception as e:  # noqa: BLE001
        log(f"AVISO: arquivo de endereços não gravado: {e}")
    fora = {k: v for k, v in resumo.items() if k != "entra"}
    log(f"Salesforce: {sum(resumo.values())} contas; aptas={resumo['entra']}; fora do Excel={sum(fora.values())}")
    for motivo, qtd in sorted(fora.items(), key=lambda x: -x[1]):
        log(f"   fora: {qtd} {motivo}")
    log("dry-run: nada gravado" if a.dry_run else f"{n} linhas adicionadas")


if __name__ == "__main__":
    main()
