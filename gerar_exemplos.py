"""Gera faturas PDF fictícias em ./faturas para testar o extrator."""
import random
from datetime import date, timedelta
from pathlib import Path

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas

FORNECEDORES = [
    ("Padaria Central, Lda.", "501234567"),
    ("TecnoLuz Eletricidade", "509876543"),
    ("Papelaria Avenida", "503333444"),
    ("CleanPro Limpezas", "514567890"),
    ("Café Estrela Unipessoal", "516111222"),
]


def gerar(pasta: Path, quantidade: int = 8) -> None:
    pasta.mkdir(exist_ok=True)
    random.seed(42)
    for i in range(1, quantidade + 1):
        nome, nif = random.choice(FORNECEDORES)
        data = date(2026, 10, 1) + timedelta(days=random.randint(0, 30))
        base = round(random.uniform(20, 900), 2)
        iva = round(base * 0.23, 2)
        total = round(base + iva, 2)

        c = canvas.Canvas(str(pasta / f"fatura_{i:03d}.pdf"), pagesize=A4)
        y = 800
        for linha in [
            nome,
            f"NIF: {nif}",
            "",
            f"Fatura N.º FT 2026/{100 + i}",
            f"Data: {data.strftime('%d/%m/%Y')}",
            "",
            f"Base tributável: {base:.2f} €".replace(".", ","),
            f"IVA (23%): {iva:.2f} €".replace(".", ","),
            f"Total: {total:.2f} €".replace(".", ","),
        ]:
            c.drawString(60, y, linha)
            y -= 22
        c.save()
    print(f"{quantidade} faturas criadas em {pasta}/")


if __name__ == "__main__":
    gerar(Path("faturas"))
