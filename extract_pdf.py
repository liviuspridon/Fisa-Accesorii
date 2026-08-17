import re
import pdfplumber


def extract_proforma(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        page_texts = [page.extract_text() or "" for page in pdf.pages]

    full_text = "\n".join(page_texts)

    nr_intern = None
    data_livrare = None
    m = re.search(r"Nr\.\s*intern\s+(\d+)", full_text)
    if m:
        nr_intern = m.group(1)
    m = re.search(r"Data livrare\s+(\d{2}/\d{2}/\d{4})", full_text)
    if m:
        data_livrare = m.group(1)

    proiect = None
    lines = full_text.splitlines()
    for i, line in enumerate(lines):
        if line.strip().startswith("Acest document contine") and i + 1 < len(lines):
            candidate = lines[i + 1].strip()
            if candidate:
                proiect = candidate
                break

    item_start_re = re.compile(r"^(\d+)\s+(.*)$")
    trailing_re = re.compile(
        r"^(?P<desc>.*?)\s+(?P<um>Buc|buc\.|SET|Set|Kg|ML|ml)\s+"
        r"(?P<cant>[\d.,]+)\s+[\d.,]+\s+[\d.,]+\s+\d+%\s+[\d.,]+\s*$"
    )

    def is_header(line):
        s = line.strip()
        patterns = [
            r"^F\s*A\s*C\s*T\s*U\s*R\s*A$",
            r"^P\s*R\s*O\s*F\s*O\s*R\s*M\s*A\b",
            r"^Data\s*\(", r"^Nr\.\s*intern\b", r"^Furnizor\b",
            r"^Cumparator\b", r"^CIF/CUI\b", r"^Nr\.\s*Reg\.\s*Com\.",
            r"^Sediul\b", r"^Loc\./Jud\.", r"^Contul\b", r"^Banca\b",
            r"^Punct lucru\b", r"^Agent\b", r"^Pag\.\s*\d+\s*/\s*\d+",
            r"^Nr\.$", r"^crt\.$", r"^Denumirea produselor\b",
            r"^sau a serviciilor$", r"^U\.?M\.?$", r"^Cant$",
            r"^Pretul unitar$", r"^\(fara TVA\)$", r"^Valoarea$",
            r"^TVA$", r"^%$", r"^Semnatura\b", r"^stampila\b",
            r"^furnizorului\b", r"^Date privind expeditia\b",
            r"^Semnatura de\b", r"^primire\b", r"^Numele delegatului\b",
            r"^Total de plata\b", r"^C\.I\.\s*Seria\b",
            r"^Mijlocul de transport\b", r"^Expedierea s-a efectuat\b",
            r"^Exemplar:\b", r"^Emis cu Neomanager\b", r"^Taxabila cf\."
        ]
        return any(re.search(p, s, re.I) for p in patterns)

    items = []

    # Each page is parsed independently. No fixed product number is used.
    for page_text in page_texts:
        page_lines = page_text.splitlines()
        current = None
        in_table = False

        for raw in page_lines:
            line = raw.strip()
            if not line:
                continue

            if re.search(r"Denumirea produselor", line, re.I):
                in_table = True
                continue
            if in_table and (
                line.lower() in {"sau a serviciilor", "u.m.", "um", "cant", "valoarea", "tva", "%"}
                or re.search(r"Pretul unitar|fara TVA", line, re.I)
            ):
                continue

            if line.startswith("Taxabila cf.") or line.startswith("Acest document contine"):
                if current:
                    items.append(current)
                    current = None
                in_table = False
                continue

            if is_header(line):
                continue

            m = item_start_re.match(line)
            if m:
                crt = int(m.group(1))
                rest = m.group(2)
                full = trailing_re.match(rest)
                if current:
                    items.append(current)
                    current = None
                if full:
                    current = {
                        "crt": crt,
                        "denumire": full.group("desc").strip(),
                        "um": full.group("um"),
                        "cantitate": float(full.group("cant").replace(",", ".")),
                    }
                else:
                    current = {"crt": crt, "denumire": rest, "um": None, "cantitate": None}
                in_table = True
                continue

            if current:
                end = re.search(
                    r"\s+(Buc|buc\.|SET|Set|Kg|ML|ml)\s+([\d.,]+)\s+[\d.,]+\s+[\d.,]+\s+\d+%\s+[\d.,]+\s*$",
                    line,
                )
                if end:
                    desc = line[:end.start()].strip()
                    if desc:
                        current["denumire"] += " " + desc
                    current["um"] = end.group(1)
                    current["cantitate"] = float(end.group(2).replace(",", "."))
                elif not is_header(line):
                    current["denumire"] += " " + line

        if current:
            items.append(current)

    # Only completed product rows are returned.
    items = [x for x in items if x["cantitate"] is not None]

    return {
        "nr_intern": nr_intern,
        "data_livrare": data_livrare,
        "proiect": proiect,
        "items": items,
    }
