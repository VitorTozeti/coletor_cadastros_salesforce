"""Gera PowerApp_Clientes_Novos.xlsx: arquivo PEQUENO só com os clientes que entraram nos últimos dias.

Quem lê é o fluxo do Power Automate (tabela `Novos` -> lista SharePoint `Clientes_PowerApp`). Ler só dezenas de
linhas em vez das 17 mil do arquivo principal deixa o fluxo diário rápido. O arquivo é SEMPRE recriado do zero
(nada a preservar: sem Power Query, sem conexões) e trocado de forma atômica — se falhar, o anterior continua valendo.
Escrito à mão (zipfile), sem openpyxl.
"""
import os
import zipfile
from pathlib import Path
from xml.sax.saxutils import escape

COLUNAS = ["CardCode", "CardName", "Cod_Nome", "MailZipCod", "E_Mail", "MailStrNo", "Telefone", "Tipo_de_Conta",
           "Password"]
NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def _letra(i):
    return chr(ord("A") + i)


def _cel(ref, v):
    return f'<c r="{ref}" t="inlineStr"><is><t xml:space="preserve">{escape(str(v or ""))}</t></is></c>'


def grava(destino, linhas):
    """linhas: lista de dicts com as COLUNAS. Mantém ao menos 1 linha (uma tabela do Excel não pode ficar vazia)."""
    linhas = list(linhas)
    if not linhas:  # linha-sentinela: o fluxo ignora CardCode vazio
        linhas = [{c: "" for c in COLUNAS}]
    ultima = len(linhas) + 1
    fim = f"{_letra(len(COLUNAS) - 1)}{ultima}"
    rows = ['<row r="1">' + "".join(_cel(f"{_letra(i)}1", c) for i, c in enumerate(COLUNAS)) + "</row>"]
    for n, l in enumerate(linhas, start=2):
        rows.append(f'<row r="{n}">' + "".join(_cel(f"{_letra(i)}{n}", l.get(c, "")) for i, c in enumerate(COLUNAS)) + "</row>")
    sheet = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="{NS}" '
             'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
             f'<dimension ref="A1:{fim}"/><sheetData>{"".join(rows)}</sheetData>'
             '<tableParts count="1"><tablePart r:id="rId1"/></tableParts></worksheet>')
    cols = "".join(f'<tableColumn id="{i + 1}" name="{c}"/>' for i, c in enumerate(COLUNAS))
    tabela = (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><table xmlns="{NS}" id="1" name="Novos" '
              f'displayName="Novos" ref="A1:{fim}" totalsRowShown="0"><autoFilter ref="A1:{fim}"/>'
              f'<tableColumns count="{len(COLUNAS)}">{cols}</tableColumns>'
              '<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" showRowStripes="1" '
              'showColumnStripes="0"/></table>')
    partes = {
        "[Content_Types].xml": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/></Types>',
        "_rels/.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
        "xl/workbook.xml": f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="{NS}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="Novos" sheetId="1" r:id="rId1"/></sheets></workbook>',
        "xl/_rels/workbook.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/></Relationships>',
        "xl/worksheets/sheet1.xml": sheet,
        "xl/worksheets/_rels/sheet1.xml.rels": '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/></Relationships>',
        "xl/tables/table1.xml": tabela,
    }
    destino = Path(destino)
    tmp = destino.with_suffix(".tmp.xlsx")
    try:
        with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as z:
            for nome, conteudo in partes.items():
                z.writestr(nome, conteudo)
        with zipfile.ZipFile(tmp) as z:  # reabre antes de trocar
            if z.testzip() is not None or len(z.namelist()) != len(partes):
                raise RuntimeError("Novos.xlsx gerado inválido")
        os.replace(tmp, destino)
    except BaseException:
        tmp.unlink(missing_ok=True)
        raise
