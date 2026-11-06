"""Extrator de faturas: lê PDFs de uma pasta e junta os dados num Excel.

Uso:
    python extrator.py                       # lê ./faturas, gera faturas.xlsx
    python extrator.py minha_pasta out.xlsx
"""
import re
import sys
from datetime import datetime
from pathlib import Path

import pdfplumber
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

PADROES = {
    "numero": r"Fatura\s*N\.?\s*[ºo°]?\s*[:\-]?\s*([A-Z]{1,3}\s*\d{4}/\d+)",
    "data": r"Data\s*[:\-]?\s*(\d{2}[/-]\d{2}[/-]\d{4})",
    "nif": r"NIF\s*[:\-]?\s*(\d{9})",
    "base": r"Base\s+tribut[áa]vel\s*[:\-]?\s*([\d.\s]+,\d{2})",
    "iva": r"IVA[^:\n]*[:\-]?\s*([\d.\s]+,\d{2})",
    "total": r"Total\s*[:\-]?\s*([\d.\s]+,\d{2})",
}
COLUNAS = ["Ficheiro", "Fornecedor", "NIF", "N.º Fatura", "Data",
           "Base (€)", "IVA (€)", "Total (€)"]


def para_numero(texto: str) -> float:
    """'1.234,56' -> 1234.56"""
    return float(texto.replace(" ", "").replace(".", "").replace(",", "."))


def extrair(pdf: Path) -> dict:
    with pdfplumber.open(pdf) as doc:
        texto = "\n".join(p.extract_text() or "" for p in doc.pages)
    if not texto.strip():
        raise ValueError("PDF sem texto (pode ser digitalizado)")

    achados = {}
    for campo, padrao in PADROES.items():
        m = re.search(padrao, texto, re.IGNORECASE)
        achados[campo] = m.group(1).strip() if m else None
    if not achados["total"]:
        raise ValueError("não encontrei o total")

    return {
        "Ficheiro": pdf.name,
        "Fornecedor": texto.strip().splitlines()[0].strip(),
        "NIF": achados["nif"],
        "N.º Fatura": achados["numero"],
        "Data": datetime.strptime(achados["data"].replace("-", "/"), "%d/%m/%Y")
                if achados["data"] else None,
        "Base (€)": para_numero(achados["base"]) if achados["base"] else None,
        "IVA (€)": para_numero(achados["iva"]) if achados["iva"] else None,
        "Total (€)": para_numero(achados["total"]),
    }


def escrever_excel(linhas: list[dict], erros: list[tuple], destino: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Faturas"
    ws.append(COLUNAS)
    for linha in sorted(linhas, key=lambda l: l["Data"] or datetime.min):
        ws.append([linha[c] for c in COLUNAS])

    ultima = ws.max_row
    ws.append(["TOTAL", "", "", "", "",
               f"=SUM(F2:F{ultima})", f"=SUM(G2:G{ultima})", f"=SUM(H2:H{ultima})"])

    cabecalho = PatternFill("solid", fgColor="1F4E78")
    for cel in ws[1]:
        cel.font = Font(bold=True, color="FFFFFF")
        cel.fill = cabecalho
        cel.alignment = Alignment(horizontal="center")
    for cel in ws[ws.max_row]:
        cel.font = Font(bold=True)
    for row in ws.iter_rows(min_row=2):
        row[4].number_format = "DD/MM/YYYY"
        for cel in row[5:8]:
            cel.number_format = '#,##0.00 "€"'
    for i, largura in enumerate([16, 28, 12, 16, 12, 12, 12, 12], start=1):
        ws.column_dimensions[get_column_letter(i)].width = largura
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:H{ultima}"

    if erros:
        we = wb.create_sheet("Por rever")
        we.append(["Ficheiro", "Problema"])
        for cel in we[1]:
            cel.font = Font(bold=True)
        for erro in erros:
            we.append(list(erro))
        we.column_dimensions["A"].width = 20
        we.column_dimensions["B"].width = 50

    wb.save(destino)


def main() -> None:
    pasta = Path(sys.argv[1] if len(sys.argv) > 1 else "faturas")
    destino = Path(sys.argv[2] if len(sys.argv) > 2 else "faturas.xlsx")
    pdfs = sorted(pasta.glob("*.pdf"))
    if not pdfs:
        sys.exit(f"Nenhum PDF encontrado em {pasta}/")

    linhas, erros = [], []
    for pdf in pdfs:
        try:
            linhas.append(extrair(pdf))
            print(f"  ✓ {pdf.name}")
        except Exception as e:  # um PDF mau não deve parar os outros
            erros.append((pdf.name, str(e)))
            print(f"  ✗ {pdf.name}: {e}")

    escrever_excel(linhas, erros, destino)
    print(f"\n{len(linhas)} faturas lidas, {len(erros)} por rever → {destino}")


if __name__ == "__main__":
    main()
