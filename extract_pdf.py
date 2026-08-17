import re
import pdfplumber


def extract_proforma(pdf_path):
    """
    Extrage datele din proforma PDF.

    Procesează fiecare pagină separat și ignoră anteturile repetate.
    Nu presupune la ce poziție se termină o pagină.
    """

    with pdfplumber.open(pdf_path) as pdf:
        page_texts = [page.extract_text() or "" for page in pdf.pages]

    full_text = "\n".join(page_texts)

    # ---------------------------------------------------------
    # DATE DOCUMENT
    # ---------------------------------------------------------

    nr_intern = None
    data_livrare = None

    m = re.search(r"Nr\.\s*intern\s+(\d+)", full_text)
    if m:
        nr_intern = m.group(1)

    m = re.search(
        r"Data livrare\s+(\d{2}/\d{2}/\d{4})",
        full_text,
    )
    if m:
        data_livrare = m.group(1)

    # ---------------------------------------------------------
    # PROIECT
    # ---------------------------------------------------------

    proiect = None

    lines = full_text.splitlines()

    for i, line in enumerate(lines):
        if line.strip().startswith("Acest document contine"):
            if i + 1 < len(lines):
                candidate = lines[i + 1].strip()
                if candidate:
                    proiect = candidate
                    break

    # ---------------------------------------------------------
    # REGEX ARTICOLE
    # ---------------------------------------------------------

    item_start_re = re.compile(r"^(\d+)\s+(.*)$")

    trailing_re = re.compile(
        r"^(?P<desc>.*?)\s+"
        r"(?P<um>Buc|buc\.|SET|Set|Kg|ML|ml)\s+"
        r"(?P<cant>[\d.,]+)\s+"
        r"[\d.,]+\s+"
        r"[\d.,]+\s+"
        r"\d+%\s+"
        r"[\d.,]+\s*$"
    )

    items = []

    # ---------------------------------------------------------
    # FUNCȚII PENTRU ANTET / FOOTER
    # ---------------------------------------------------------

    header_token_re = re.compile(
        r"F\s*A\s*C\s*T\s*U\s*R\s*A",
        re.IGNORECASE,
    )

    proforma_token_re = re.compile(
        r"P\s*R\s*O\s*F\s*O\s*R\s*M\s*A",
        re.IGNORECASE,
    )

    def is_header_or_footer(line):
        s = line.strip()

        patterns = [
            r"^Data\s*\(",
            r"^Nr\.\s*intern\b",
            r"^Furnizor\b",
            r"^Cumparator\b",
            r"^CIF/CUI\b",
            r"^Nr\.\s*Reg\.\s*Com\.",
            r"^Sediul\b",
            r"^Loc\./Jud\.",
            r"^Contul\b",
            r"^Banca\b",
            r"^Punct lucru\b",
            r"^Agent\b",
            r"^Pag\.\s*\d+\s*/\s*\d+",
            r"^Nr\.$",
            r"^crt\.$",
            r"^Denumirea produselor\b",
            r"^sau a serviciilor$",
            r"^U\.?M\.?$",
            r"^Cant$",
            r"^Pretul unitar$",
            r"^\(fara TVA\)$",
            r"^Valoarea$",
            r"^TVA$",
            r"^%$",
            r"^Semnatura\b",
            r"^stampila\b",
            r"^furnizorului\b",
            r"^Date privind expeditia\b",
            r"^Semnatura de\b",
            r"^primire\b",
            r"^Numele delegatului\b",
            r"^Total de plata\b",
            r"^C\.I\.\s*Seria\b",
            r"^Mijlocul de transport\b",
            r"^Expedierea s-a efectuat\b",
            r"^Exemplar:\b",
            r"^Emis cu Neomanager\b",
            r"^Taxabila cf\.",
            r"^Acest document contine:",
        ]

        return (
            header_token_re.search(s)
            or proforma_token_re.search(s)
            or any(
                re.search(pattern, s, re.IGNORECASE)
                for pattern in patterns
            )
        )

    # ---------------------------------------------------------
    # PROCESARE PAGINĂ CU PAGINĂ
    # ---------------------------------------------------------

    for page_text in page_texts:

        page_lines = page_text.splitlines()

        current = None
        in_table = False

        for raw_line in page_lines:

            line = raw_line.strip()

            if not line:
                continue

            # -----------------------------------------------------
            # FOARTE IMPORTANT:
            # Uneori pdfplumber unește ultima linie de produs cu
            # antetul paginii următoare.
            #
            # Exemplu real:
            # "262.50.359 Panel comp.Keku AS black F A C T U R A
            #  P R O F O R M A Data..."
            #
            # În acest caz păstrăm doar partea produsului.
            # -----------------------------------------------------

            header_match = header_token_re.search(line)

            if header_match:
                before_header = line[:header_match.start()].strip()

                # Dacă există conținut înainte de antet și avem un
                # produs în lucru, partea dinaintea antetului aparține
                # produsului.
                if current and before_header:
                    current["denumire"] += " " + before_header

                # Antetul marchează sfârșitul tabelului de pe pagina
                # curentă. Salvăm produsul și ignorăm tot ce urmează.
                if current:
                    items.append(current)
                    current = None

                in_table = False
                continue

            # -----------------------------------------------------
            # Dacă apare PROFORMA fără FACTURA, tratăm la fel.
            # -----------------------------------------------------

            if proforma_token_re.search(line):
                if current:
                    items.append(current)
                    current = None

                in_table = False
                continue

            # -----------------------------------------------------
            # Detectare început tabel.
            # -----------------------------------------------------

            if re.search(
                r"Denumirea produselor",
                line,
                re.IGNORECASE,
            ):
                in_table = True
                continue

            if in_table and line.lower() in {
                "sau a serviciilor",
                "u.m.",
                "um",
                "cant",
                "valoarea",
                "tva",
                "%",
            }:
                continue

            if in_table and re.search(
                r"pretul unitar",
                line,
                re.IGNORECASE,
            ):
                continue

            if in_table and re.search(
                r"fara TVA",
                line,
                re.IGNORECASE,
            ):
                continue

            # -----------------------------------------------------
            # Footer.
            # -----------------------------------------------------

            if line.startswith("Taxabila cf."):
                if current:
                    items.append(current)
                    current = None

                in_table = False
                continue

            if line.startswith("Acest document contine"):
                if current:
                    items.append(current)
                    current = None

                in_table = False
                continue

            # Alte elemente de antet/footer.
            if is_header_or_footer(line):
                continue

            # -----------------------------------------------------
            # Articol nou.
            # -----------------------------------------------------

            m_item = item_start_re.match(line)

            if m_item:

                crt = int(m_item.group(1))
                rest = m_item.group(2)

                m_full = trailing_re.match(rest)

                if m_full:

                    if current:
                        items.append(current)

                    current = {
                        "crt": crt,
                        "denumire": m_full.group("desc").strip(),
                        "um": m_full.group("um"),
                        "cantitate": float(
                            m_full.group("cant").replace(",", ".")
                        ),
                    }

                    in_table = True
                    continue

                # Început de descriere care continuă pe liniile următoare.
                if current:
                    items.append(current)

                current = {
                    "crt": crt,
                    "denumire": rest,
                    "um": None,
                    "cantitate": None,
                }

                in_table = True
                continue

            # -----------------------------------------------------
            # Continuare articol.
            # -----------------------------------------------------

            if current:

                m_end = re.search(
                    r"\s+(Buc|buc\.|SET|Set|Kg|ML|ml)\s+"
                    r"([\d.,]+)\s+"
                    r"[\d.,]+\s+"
                    r"[\d.,]+\s+"
                    r"\d+%\s+"
                    r"[\d.,]+\s*$",
                    line,
                )

                if m_end:

                    desc = line[:m_end.start()].strip()

                    if desc:
                        current["denumire"] += " " + desc

                    current["um"] = m_end.group(1)
                    current["cantitate"] = float(
                        m_end.group(2).replace(",", ".")
                    )

                    continue

                current["denumire"] += " " + line

        # Produsul de la finalul paginii.
        if current:
            items.append(current)

    # ---------------------------------------------------------
    # CURĂȚARE FINALĂ
    # ---------------------------------------------------------

    clean_items = []

    for item in items:
        if item["cantitate"] is None:
            continue

        clean_items.append(item)

    return {
        "nr_intern": nr_intern,
        "data_livrare": data_livrare,
        "proiect": proiect,
        "items": clean_items,
    }


if __name__ == "__main__":
    import json
    import sys

    pdf_path = (
        sys.argv[1]
        if len(sys.argv) > 1
        else "FP_260800280.pdf"
    )

    print(
        json.dumps(
            extract_proforma(pdf_path),
            indent=2,
            ensure_ascii=False,
        )
    )
