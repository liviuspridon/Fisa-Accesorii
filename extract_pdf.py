import re
import pdfplumber


def extract_proforma(pdf_path):
    """
    Extrage din proforma PDF:
    - Nr. intern
    - Data livrare
    - linia liber-text de proiect
    - lista de accesorii: denumire + cantitate
    """

    with pdfplumber.open(pdf_path) as pdf:
        text = "\n".join(
            page.extract_text()
            for page in pdf.pages
            if page.extract_text()
        )

    lines = text.split("\n")

    # --- Header fields ---
    nr_intern = None
    data_livrare = None

    m = re.search(r"Nr\.\s*intern\s+(\d+)", text)
    if m:
        nr_intern = m.group(1)

    m = re.search(r"Data livrare\s+(\d{2}/\d{2}/\d{4})", text)
    if m:
        data_livrare = m.group(1)

    # --- Linia libera de proiect ---
    # Apare intre "Acest document contine..." si "Taxabila"
    proiect = None

    for i, line in enumerate(lines):
        if line.strip().startswith("Acest document contine"):
            if i + 1 < len(lines):
                candidate = lines[i + 1].strip()
                if candidate and not candidate.lower().startswith("taxabila"):
                    proiect = candidate
            break

    # --- Linii de produs ---
    item_start_re = re.compile(r"^(\d+)\s+(.*)$")

    trailing_re = re.compile(
        r"^(?P<desc>.*?)\s+"
        r"(?P<um>Buc|SET|Set|Kg|buc\.|ml)\s+"
        r"(?P<cant>[\d.]+)\s+"
        r"[\d.]+\s+"
        r"[\d.]+\s+"
        r"\d+%\s+"
        r"[\d.]+\s*$"
    )

    items = []
    current = None
    in_table = False

    for line in lines:
        if re.search(r"crt\.\s*sau a serviciilor", line):
            in_table = True
            continue

        if line.strip().startswith("Acest document contine"):
            in_table = False
            if current:
                items.append(current)
                current = None
            continue

        if not in_table:
            continue

        m_item = item_start_re.match(line)

        if m_item and trailing_re.match(m_item.group(2)):
            if current:
                items.append(current)

            crt = m_item.group(1)
            rest = m_item.group(2)
            mt = trailing_re.match(rest)

            current = {
                "crt": int(crt),
                "denumire": mt.group("desc").strip(),
                "um": mt.group("um"),
                "cantitate": float(mt.group("cant")),
            }

        else:
            # Linie de continuare a descrierii
            if current and line.strip():
                current["denumire"] += " " + line.strip()

    if current:
        items.append(current)

    return {
        "nr_intern": nr_intern,
        "data_livrare": data_livrare,
        "proiect": proiect,
        "items": items,
    }


if __name__ == "__main__":
    import json
    import sys

    pdf_path = (
        sys.argv[1]
        if len(sys.argv) > 1
        else "FP_260800280.pdf"
    )

    data = extract_proforma(pdf_path)
    print(json.dumps(data, indent=2, ensure_ascii=False))
