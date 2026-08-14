import re
import tempfile
from pathlib import Path

import streamlit as st

from extract_pdf import extract_proforma
from generate_fisa import build_fisa


st.set_page_config(
    page_title="Fisa Accesorii",
    page_icon="📋",
    layout="centered",
)

st.title("📋 Generare Fișă Accesorii")
st.write(
    "Încarcă proforma PDF și generează automat fișa de accesorii în format Excel."
)

pdf_file = st.file_uploader(
    "Încarcă proforma PDF",
    type=["pdf"],
)


if pdf_file is not None:

    if st.button("Extrage și generează fișa", type="primary"):

        try:

            with tempfile.TemporaryDirectory() as tmp_dir:

                # ---------------------------------------------------------
                # SALVARE PDF TEMPORAR
                # ---------------------------------------------------------

                pdf_path = Path(tmp_dir) / pdf_file.name
                pdf_path.write_bytes(pdf_file.getvalue())

                # ---------------------------------------------------------
                # EXTRAGERE PDF
                # ---------------------------------------------------------

                with st.spinner("Se extrag datele din PDF..."):
                    data = extract_proforma(str(pdf_path))

                if not data["items"]:
                    st.error(
                        "Nu au fost găsite articole în proformă. "
                        "Verifică formatul PDF-ului."
                    )
                    st.stop()

                # ---------------------------------------------------------
                # EXTRAGERE TITLU + NUMĂR PROIECT
                # ---------------------------------------------------------

                linia = (data.get("proiect") or "").strip()

                st.write("### Date extrase din PDF")
                st.write("**Linia proiectului:**", linia)

                # Exemplu:
                #
                # EGGER - PROIECT 635243586 - EGG STORAGE
                #
                # Rezultat:
                #
                # titlu   = EGGER - PROIECT - 635243586
                # proiect = EGG STORAGE

                match = re.search(
                    r"^(.*?)\s*-\s*PROIECT\s+(\d+)(.*)$",
                    linia,
                    flags=re.IGNORECASE,
                )

                if match:

                    prefix = match.group(1).strip()
                    nr_proiect = match.group(2).strip()
                    rest = match.group(3).strip()

                    rest = rest.lstrip("-").strip()

                    if prefix:
                        titlu = (
                            f"{prefix} - PROIECT - {nr_proiect}"
                        )
                    else:
                        titlu = (
                            f"PROIECT - {nr_proiect}"
                        )

                    proiect = rest

                else:

                    # -----------------------------------------------------
                    # FALLBACK
                    # -----------------------------------------------------

                    titlu = linia
                    proiect = ""

                    # Dacă nu găsim PROIECT + număr, încercăm să
                    # găsim orice număr de minimum 5 cifre.
                    nr_match = re.search(
                        r"\b(\d{5,})\b",
                        linia,
                    )

                    if nr_match:

                        nr_proiect = nr_match.group(1)

                        partea_dinainte = linia[
                            :nr_match.start()
                        ].strip()

                        partea_dinainte = (
                            partea_dinainte
                            .rstrip("-")
                            .strip()
                        )

                        titlu = (
                            f"{partea_dinainte} - {nr_proiect}"
                        )

                        partea_de_dupa = linia[
                            nr_match.end():
                        ].strip()

                        proiect = (
                            partea_de_dupa
                            .lstrip("-")
                            .strip()
                        )

                # ---------------------------------------------------------
                # CURĂȚARE TITLU
                # ---------------------------------------------------------

                titlu = re.sub(
                    r"\s+",
                    " ",
                    titlu,
                ).strip()

                proiect = re.sub(
                    r"\s+",
                    " ",
                    proiect,
                ).strip()

                # ---------------------------------------------------------
                # NUME FIȘIER
                # ---------------------------------------------------------

                nume_fisa = titlu

                # Caractere interzise în Windows
                nume_fisa = re.sub(
                    r'[<>:"/\\|?*]',
                    "",
                    nume_fisa,
                )

                nume_fisa = re.sub(
                    r"\s+",
                    " ",
                    nume_fisa,
                ).strip()

                # Limităm lungimea numelui
                nume_fisa = nume_fisa[:150].strip()

                if not nume_fisa:
                    nume_fisa = "fisa_accesorii"

                file_name = f"{nume_fisa}.xlsx"

                output_path = (
                    Path(tmp_dir) / file_name
                )

                # ---------------------------------------------------------
                # AFIȘARE DEBUG
                # ---------------------------------------------------------

                st.write("**Titlu extras:**", titlu)
                st.write("**Număr proiect:**", nr_proiect if "nr_proiect" in locals() else "-")
                st.write("**Denumire proiect:**", proiect)

                st.warning(
                    f"Fișierul va fi: {nume_fisa}.xlsx"
                )

                # ---------------------------------------------------------
                # ARTICOLE
                # ---------------------------------------------------------

                st.success(
                    f"Au fost găsite {len(data['items'])} articole."
                )

                st.dataframe(
                    [
                        {
                            "Crt.": item["crt"],
                            "Denumire produs": item["denumire"],
                            "Cantitate": item["cantitate"],
                            "UM": item["um"],
                        }
                        for item in data["items"]
                    ],
                    use_container_width=True,
                    hide_index=True,
                )

                # ---------------------------------------------------------
                # GENERARE EXCEL
                # ---------------------------------------------------------

                with st.spinner(
                    "Se generează fișierul Excel..."
                ):

                    build_fisa(
                        items=data["items"],
                        titlu=titlu,
                        proiect=proiect,
                        data=data["data_livrare"],
                        nr_document=data["nr_intern"],
                        output_path=str(output_path),
                    )

                st.success(
                    "Fișa Accesorii a fost generată."
                )

                # ---------------------------------------------------------
                # DOWNLOAD
                # ---------------------------------------------------------

                st.download_button(
                    label=f"⬇️ Descarcă {file_name}",
                    data=output_path.read_bytes(),
                    file_name=file_name,
                    mime=(
                        "application/"
                        "vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    ),
                    type="primary",
                )

        except Exception as e:

            st.error(
                f"A apărut o eroare: {e}"
            )

            st.exception(e)
