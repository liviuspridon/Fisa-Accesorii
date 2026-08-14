"""
Genereaza o Fisa Accesorii in formatul exact al 617461120_Delta_Ap_Madalina.xlsx,
pornind de la un template si de la datele extrase din proforma PDF.

Structura template (Sheet1):
  A3:B3  = titlu (spatiu liber, ex. "DELTA Studio_ 617461120")
  C3:E3  = data
  A4:B4  = denumire proiect (spatiu liber, ex. "AP MADALINA")
  C4:E4  = numar document
  Rand 5 = header tabel (Crt / Denumire Produs / Cant / Conf / Dif)
  Rand 6...  = randuri articole (se multiplica stilul dupa nevoie)
  Ultimele 2 randuri = footer "Fisa Accesorii : ..." / "Depozit : ..."
"""
import copy
import datetime
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

TEMPLATE_PATH = "/mnt/user-data/uploads/617461120_Delta_Ap_Madalina.xlsx"
FIRST_DATA_ROW = 6
HEADER_ROW = 5


def _copy_row_style(ws, src_row, dst_row, ncols=5):
    for col in range(1, ncols + 1):
        src = ws.cell(row=src_row, column=col)
        dst = ws.cell(row=dst_row, column=col)
        dst.font = copy.copy(src.font)
        dst.border = copy.copy(src.border)
        dst.fill = copy.copy(src.fill)
        dst.alignment = copy.copy(src.alignment)
        dst.number_format = src.number_format
    ws.row_dimensions[dst_row].height = ws.row_dimensions[src_row].height


def build_fisa(items, titlu, proiect, data, nr_document,
                fisa_de=None, depozit=None, output_path="fisa_accesorii.xlsx"):
    wb = load_workbook(TEMPLATE_PATH)
    ws = wb["Sheet1"]

    # gaseste in template ultimul rand de date si primul rand de footer
    # (in template: date pe randurile 6..27, footer pe 28..29)
    template_last_data_row = None
    for r in range(FIRST_DATA_ROW, ws.max_row + 1):
        if ws.cell(row=r, column=1).value not in (None, "") and \
           isinstance(ws.cell(row=r, column=1).value, (int, float)):
            template_last_data_row = r
        else:
            break
    footer_row_1 = template_last_data_row + 1
    footer_row_2 = footer_row_1 + 1

    # NU pastram numele din template (Teo Pavel / Stelian Grecu) - sunt specifice
    # exemplului original. Folosim doar etichetele, cu spatiu de completat manual.
    footer_vals = {
        "a1": "Fisa Accesorii :   ",
        "c1": None,
        "a2": "Depozit       :        ",
    }

    # openpyxl nu realiniaza intotdeauna corect merged cells la insert/delete_rows,
    # asa ca eliminam manual merge-ul din footer inainte de a muta randurile
    for rng in list(ws.merged_cells.ranges):
        if rng.min_row >= footer_row_1:
            ws.unmerge_cells(str(rng))

    n_items = len(items)
    n_template_rows = template_last_data_row - FIRST_DATA_ROW + 1

    # ajusteaza numarul de randuri din tabel la numarul de articole
    if n_items > n_template_rows:
        ws.insert_rows(template_last_data_row + 1, n_items - n_template_rows)
        for i in range(n_template_rows, n_items):
            _copy_row_style(ws, FIRST_DATA_ROW, FIRST_DATA_ROW + i)
    elif n_items < n_template_rows:
        ws.delete_rows(FIRST_DATA_ROW + n_items, n_template_rows - n_items)

    footer_row_1 = FIRST_DATA_ROW + n_items
    footer_row_2 = footer_row_1 + 1
    ws.merge_cells(f"C{footer_row_1}:E{footer_row_2}")

    # curata orice randuri goale ramase dupa footer (relevant la delete_rows)
    if ws.max_row > footer_row_2:
        ws.delete_rows(footer_row_2 + 1, ws.max_row - footer_row_2)

    # --- Header ---
    ws["A3"] = titlu
    ws["C3"] = data if isinstance(data, datetime.date) else \
        datetime.datetime.strptime(data, "%d/%m/%Y")
    ws["A4"] = proiect
    ws["C4"] = nr_document

    # --- Randuri articole ---
    for i, item in enumerate(items):
        r = FIRST_DATA_ROW + i
        ws.cell(row=r, column=1, value=i + 1)
        ws.cell(row=r, column=2, value=item["denumire"])
        ws.cell(row=r, column=3, value=item["cantitate"])
        # coloanele D (Conf) si E (Dif) raman libere, de completat manual

    # --- Footer ---
    ws.cell(row=footer_row_1, column=1, value=footer_vals["a1"] if fisa_de is None
             else f"Fisa Accesorii :   {fisa_de}")
    ws.cell(row=footer_row_1, column=3, value=footer_vals["c1"] or datetime.date.today())
    ws.cell(row=footer_row_2, column=1, value=footer_vals["a2"] if depozit is None
             else f"Depozit       :        {depozit}")

    wb.save(output_path)
    return output_path


if __name__ == "__main__":
    from extract_pdf import extract_proforma

    data = extract_proforma("/mnt/user-data/uploads/FP_260800280.pdf")

    # linia libera de pe proforma: "EGGER PROIECT 635243588 - SIDEBOARD"
    # se imparte in titlu (partea dinainte de " - ") si proiect (partea de dupa)
    linia = data["proiect"] or ""
    if " - " in linia:
        titlu, proiect = [p.strip() for p in linia.split(" - ", 1)]
    else:
        titlu, proiect = linia, ""

    build_fisa(
        items=data["items"],
        titlu=titlu,
        proiect=proiect,
        data=data["data_livrare"],
        nr_document=data["nr_intern"],
        fisa_de="TEODOR PAVEL",
        depozit="STELIAN GRECU",
        output_path="/home/claude/fisa_app/fisa_accesorii_demo.xlsx",
    )
    print("done")
