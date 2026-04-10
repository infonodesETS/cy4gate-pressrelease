import requests
from bs4 import BeautifulSoup
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import re
from datetime import datetime

BASE_URL = "https://www.cy4gate.com"

PAGES = {
    2024: "/company/investor-relations/financial-press-releases/year-2024/",
    2025: "/company/investor-relations/financial-press-releases/year-2025/",
    2026: "/company/investor-relations/financial-press-releases/year-3/",
}

HEADERS = {
    "User-Agent": "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
}


def parse_date_time(raw: str):
    """Estrae data e ora da stringhe tipo '27/12/2024 7.42' o '14.05.2025 18.42'."""
    raw = raw.strip()
    # Pulisce doppi slash tipo "30/12//2025"
    raw = re.sub(r'/+', '/', raw)
    # Formato con punto come separatore di data tipo "14.05.2025"
    raw = re.sub(r'^(\d{2})\.(\d{2})\.(\d{4})', r'\1/\2/\3', raw)

    parts = raw.split()
    date_str = parts[0] if parts else ""
    time_str = parts[1] if len(parts) > 1 else ""

    # Prova a parsare la data
    try:
        date_obj = datetime.strptime(date_str, "%d/%m/%Y").date()
    except ValueError:
        date_obj = None

    return date_obj, time_str


def scrape_year(year: int, path: str):
    url = BASE_URL + path
    print(f"  Scaricando {year}: {url}")
    resp = requests.get(url, headers=HEADERS, timeout=30)
    resp.raise_for_status()
    soup = BeautifulSoup(resp.text, "html.parser")

    # Il contenuto dei comunicati è nell'area principale (fuori nav/footer)
    # La struttura è: <ul><li>DATA ORA</li></ul><p><a href="...">TITOLO</a></p>
    # Cerchiamo il blocco principale
    main = soup.find("div", class_=re.compile(r"main|content|body", re.I))
    if not main:
        main = soup.body

    records = []
    current_date = None
    current_time = None
    seen_pdfs = set()  # per deduplicare sub-voci dello stesso comunicato

    elements = main.find_all(["ul", "p", "h1", "h2", "h3"])

    for el in elements:
        if el.name == "ul":
            li_texts = [li.get_text(strip=True) for li in el.find_all("li")]
            for text in li_texts:
                # Verifica se il testo sembra una data (contiene cifre e /)
                if re.search(r'\d{1,2}[/\.]\d{1,2}[/\.]\d{4}', text):
                    date_obj, time_str = parse_date_time(text)
                    current_date = date_obj
                    current_time = time_str
                    seen_pdfs = set()  # reset per nuova data

        elif el.name == "p" and current_date:
            a = el.find("a", href=True)
            if not a:
                continue
            href = a.get("href", "").strip()
            if not href:
                continue
            # Costruisce URL completo
            if href.startswith("/"):
                full_url = BASE_URL + href
            elif href.startswith("http"):
                full_url = href
            else:
                continue

            # Salta link non-PDF o non-asset
            if not ("/assets/" in href or href.endswith(".pdf")):
                continue

            # Deduplicazione: prende solo il primo paragrafo per ogni PDF
            if full_url in seen_pdfs:
                continue
            seen_pdfs.add(full_url)

            title = a.get_text(" ", strip=True)
            # Pulisce titolo da spazi/newline multipli
            title = re.sub(r'\s+', ' ', title).strip()

            records.append({
                "anno": year,
                "data": current_date,
                "ora": current_time,
                "titolo": title,
                "url": full_url,
            })

    print(f"    → {len(records)} comunicati trovati")
    return records


def build_excel(all_records, output_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Comunicati Stampa"

    # Intestazioni
    headers = ["Anno", "Data", "Ora", "Titolo", "Link"]
    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)

    for col, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col, value=h)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[1].height = 22

    # Dati
    link_font = Font(color="0563C1", underline="single")
    alt_fill = PatternFill(start_color="EBF3FB", end_color="EBF3FB", fill_type="solid")

    for row_idx, rec in enumerate(all_records, 2):
        fill = alt_fill if row_idx % 2 == 0 else None

        ws.cell(row=row_idx, column=1, value=rec["anno"])

        data_cell = ws.cell(row=row_idx, column=2)
        data_cell.value = rec["data"]
        data_cell.number_format = "DD/MM/YYYY"

        ws.cell(row=row_idx, column=3, value=rec["ora"])
        ws.cell(row=row_idx, column=4, value=rec["titolo"])

        # Cella link cliccabile
        link_cell = ws.cell(row=row_idx, column=5, value="Apri PDF")
        link_cell.hyperlink = rec["url"]
        link_cell.font = link_font

        if fill:
            for col in range(1, 6):
                ws.cell(row=row_idx, column=col).fill = fill

        # Wrap text per il titolo
        ws.cell(row=row_idx, column=4).alignment = Alignment(wrap_text=True, vertical="top")

    # Larghezze colonne
    ws.column_dimensions["A"].width = 8    # Anno
    ws.column_dimensions["B"].width = 14   # Data
    ws.column_dimensions["C"].width = 8    # Ora
    ws.column_dimensions["D"].width = 80   # Titolo
    ws.column_dimensions["E"].width = 12   # Link

    # Freeze header
    ws.freeze_panes = "A2"

    # Filtro automatico
    ws.auto_filter.ref = f"A1:E{len(all_records) + 1}"

    wb.save(output_path)
    print(f"\nFile salvato: {output_path}")
    print(f"Totale comunicati: {len(all_records)}")


if __name__ == "__main__":
    all_records = []
    for year, path in PAGES.items():
        records = scrape_year(year, path)
        all_records.extend(records)

    # Ordina per anno desc, data desc
    all_records.sort(key=lambda r: (r["anno"], r["data"] or datetime.min.date()), reverse=True)

    output = "/home/nelsonmau/Coding/infonodes/cy4gate/cy4gate_comunicati_stampa.xlsx"
    build_excel(all_records, output)
