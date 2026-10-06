"""Adiciona linhas ao fim de uma TABELA de um .xlsx mexendo só no XML necessário.

Por que não openpyxl: ele reconstrói o arquivo inteiro e descarta o que não conhece (Power Query, dados
externos, conexões...), e o Excel passa a pedir reparo. Aqui só 3 partes mudam — a aba (linhas novas + dimension)
e a tabela (ref/autoFilter); todo o resto do pacote é copiado byte a byte. Só biblioteca padrão.
"""
import os
import posixpath
import re
import shutil
import sys
import zipfile
from datetime import datetime
from pathlib import Path
from xml.etree import ElementTree as ET
from xml.sax.saxutils import escape

NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
# Com consulta/conexão externa as linhas seriam apagadas na próxima atualização e a ref da consulta não bateria.
BLOQUEIA = ("xl/queryTables/", "xl/connections.xml")


def _letra(n):
    s = ""
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


def _num(letras):
    n = 0
    for ch in letras:
        n = n * 26 + ord(ch) - 64
    return n


def _caminho(base_dir, alvo):
    return alvo.lstrip("/") if alvo.startswith("/") else posixpath.normpath(posixpath.join(base_dir, alvo))


def _localiza_tabela(z):
    """(caminho_da_aba, caminho_da_tabela) da primeira aba que tem tabela."""
    wb = ET.fromstring(z.read("xl/workbook.xml"))
    alvo = {r.get("Id"): r.get("Target") for r in ET.fromstring(z.read("xl/_rels/workbook.xml.rels"))}
    for sh in wb.find(f"{{{NS}}}sheets"):
        aba = _caminho("xl", alvo[sh.get(f"{{{NS_REL}}}id")])
        rels = f"{posixpath.dirname(aba)}/_rels/{posixpath.basename(aba)}.rels"
        if rels in z.namelist():
            for r in ET.fromstring(z.read(rels)):
                if r.get("Type", "").endswith("/table"):
                    return aba, _caminho(posixpath.dirname(aba), r.get("Target"))
    return None, None


def _strings_compartilhadas(z):
    if "xl/sharedStrings.xml" not in z.namelist():
        return []
    raiz = ET.fromstring(z.read("xl/sharedStrings.xml"))
    return ["".join(t.text or "" for t in si.iter(f"{{{NS}}}t")) for si in raiz.findall(f"{{{NS}}}si")]


def _valores_da_coluna(z, aba, letra, linha_ini, linha_fim):
    compart, vistos = _strings_compartilhadas(z), set()
    with z.open(aba) as fh:  # fechar é obrigatório: handle aberto impede a troca do arquivo no Windows
        for _, el in ET.iterparse(fh, events=("end",)):
            if el.tag != f"{{{NS}}}c":
                continue
            ref = el.get("r", "")
            m = re.fullmatch(r"([A-Z]+)(\d+)", ref)
            if not m or m.group(1) != letra or not (linha_ini <= int(m.group(2)) <= linha_fim):
                el.clear()
                continue
            t, v = el.get("t"), el.find(f"{{{NS}}}v")
            if t == "s" and v is not None:
                valor = compart[int(v.text)]
            elif t == "inlineStr":
                valor = "".join(x.text or "" for x in el.iter(f"{{{NS}}}t"))
            else:
                valor = v.text if v is not None else ""
            vistos.add((valor or "").strip().upper())
            el.clear()
    return vistos


def _linha_xml(n, c1, c2, colunas, linha):
    celulas = []
    for i, nome in enumerate(colunas):
        valor = linha.get(nome)
        if valor in (None, ""):
            continue  # célula realmente vazia
        celulas.append(f'<c r="{_letra(c1 + i)}{n}" t="inlineStr"><is><t xml:space="preserve">'
                       f'{escape(str(valor))}</t></is></c>')
    return f'<row r="{n}" spans="{c1}:{c2}">{"".join(celulas)}</row>'


def anexa(xlsx, colunas, novas_fn, dry_run=False, forcar=False, limite=2000):
    xlsx = str(xlsx)
    with zipfile.ZipFile(xlsx) as z:
        achados = sorted(n for n in z.namelist() if n.startswith(BLOQUEIA))
        if achados:
            sys.exit("A planilha tem consulta/conexão externa (" + ", ".join(achados[:3]) + "). Linhas gravadas aqui "
                     "seriam apagadas na atualização da consulta. Nada gravado.")
        aba, tabela = _localiza_tabela(z)
        if not aba:
            sys.exit("A planilha não tem nenhuma TABELA do Excel (Inserir > Tabela). Nada gravado.")
        tab_xml = z.read(tabela).decode("utf-8")
        m = re.search(r'<table\b[^>]*?\sref="([A-Z]+)(\d+):([A-Z]+)(\d+)"', tab_xml)
        l1, r1, l2, r2 = m.group(1), int(m.group(2)), m.group(3), int(m.group(4))
        c1, c2 = _num(l1), _num(l2)
        cab = [c.get("name") for c in ET.fromstring(tab_xml).iter(f"{{{NS}}}tableColumn")]
        nome_tab = ET.fromstring(tab_xml).get("displayName")
        faltam = [c for c in colunas if c not in cab]
        if faltam:
            sys.exit(f"Colunas ausentes na tabela {nome_tab}: {faltam}. Cabeçalho: {cab}. Nada gravado.")
        sheet = z.read(aba).decode("utf-8")
        ultima = max(int(x) for x in re.findall(r'<row r="(\d+)"', sheet))
        if ultima > r2:
            sys.exit(f"Há conteúdo abaixo da tabela (linha {ultima} > {r2}). Nada gravado.")
        letra_cod = _letra(c1 + cab.index("CardCode"))
        existentes = _valores_da_coluna(z, aba, letra_cod, r1 + 1, r2)
        novas = [l for l in novas_fn() if l["CardCode"] not in existentes]
        print(f"tabela {nome_tab}: {r2 - r1} linhas existentes; {len(novas)} CardCodes novos")
        if dry_run or not novas:
            return len(novas)
        if len(novas) > limite and not forcar:
            sys.exit(f"{len(novas)} novos > limite de {limite} (arquivo/tabela errado ou vazio?). "
                     "Nada gravado. Se for esperado, rode com --forcar.")

        fim = r2 + len(novas)
        antigo, novo = f"{l1}{r1}:{l2}{r2}", f"{l1}{r1}:{l2}{fim}"
        corpo = "".join(_linha_xml(r2 + i, c1, c2, cab, l) for i, l in enumerate(novas, start=1))
        if "</sheetData>" not in sheet or antigo not in tab_xml:
            sys.exit("Estrutura inesperada na aba/tabela. Nada gravado.")
        sheet = sheet.replace("</sheetData>", corpo + "</sheetData>", 1)
        sheet = re.sub(r'<dimension ref="[^"]*"/>', f'<dimension ref="{l1}1:{_letra(max(c2, 1))}{fim}"/>', sheet, count=1)
        tab_xml = tab_xml.replace(antigo, novo)

        backup = Path(xlsx).with_name(f"{Path(xlsx).stem}.bak-{datetime.now():%Y%m%d-%H%M%S}.xlsx")
        shutil.copy2(xlsx, backup)
        tmp = Path(xlsx).with_suffix(".tmp.xlsx")
        try:
            with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zo:
                for info in z.infolist():
                    dados = z.read(info.filename)
                    if info.filename == aba:
                        dados = sheet.encode("utf-8")
                    elif info.filename == tabela:
                        dados = tab_xml.encode("utf-8")
                    zo.writestr(info, dados)
            _verifica(tmp, aba, tabela, fim, [i.filename for i in z.infolist()])
        except BaseException:
            tmp.unlink(missing_ok=True)
            raise
    try:
        os.replace(tmp, xlsx)  # troca atômica: o arquivo nunca fica pela metade
    except OSError:  # arquivo aberto no Excel/OneDrive: limpa o temporário e deixa o original intacto
        tmp.unlink(missing_ok=True)
        raise
    return len(novas)


def _verifica(tmp, aba, tabela, fim, nomes):
    """Reabre o que acabou de escrever; se algo não fechar, aborta antes de trocar o arquivo original."""
    with zipfile.ZipFile(tmp) as z:
        if z.testzip() is not None or [i.filename for i in z.infolist()] != nomes:
            raise RuntimeError("arquivo gerado inválido (zip/partes)")
        ET.fromstring(z.read(tabela))
        linhas = [int(x) for x in re.findall(r'<row r="(\d+)"', z.read(aba).decode("utf-8"))]
        ET.fromstring(z.read(aba))
        if max(linhas) != fim or len(linhas) != len(set(linhas)):
            raise RuntimeError("arquivo gerado inconsistente (linhas)")
