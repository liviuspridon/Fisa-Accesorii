import streamlit as st
import tempfile
from pathlib import Path

from extract_pdf import extract_proforma
from generate_fisa import build_fisa


st.set_page_config(
    page_title="Fisa Accesorii",
    page_icon="📋",
    layout="centered",
)

st.title("📋 Generare Fișă Accesorii")
st.error("TEST VERSIUNE NOUA APP.PY")
st.write("Încarcă proforma PDF și generează automat fișa de accesorii în format Excel.")

pdf_file = st.file_uploader(
    "Încarcă proforma PDF",
    type=["pdf"],
)

if pdf_file is not None:
    if st.button("Extrage și generează fișa", type="primary"):
        try:
            with tempfile.TemporaryDirectory() as tmp_dir:
                pdf_path = Path(tmp_dir) / pdf_file.name
                pdf_path.write_bytes(pdf_file.getvalue())

                with st.spinner("Se extrag datele din PDF..."):
                    data = extract_proforma(str(pdf_path))

                if not data["items"]:
                    st.error(
                        "Nu au fost găsite articole în proformă. "
                        "Verifică formatul PDF-ului."
                    )
                    st.stop()

                # Afișăm datele extrase
                st.success(f"Au fost găsite {len(data['items'])} articole.")

                col1, col2 = st.columns(2)
                with col1:
                    st.write("**Nr. intern:**", data["nr_intern"] or "-")
                    st.write("**Data livrare:**", data["data_livrare"] or "-")
                with col2:
                    st.write("**Proiect:**", data["proiect"] or "-")

                st.subheader("Articole extrase")
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

                # Separă titlul și proiectul din linia:
                # EGGER PROIECT 635243588 - SIDEBOARD
                linia = data["proiect"] or ""
                if " - " in linia:
                    titlu, proiect = [p.strip() for p in linia.split(" - ", 1)]
                else:
                    titlu = linia.strip()
                    proiect = ""

                output_path = Path(tmp_dir) / "fisa_accesorii.xlsx"

                with st.spinner("Se generează fișierul Excel..."):
                    build_fisa(
                        items=data["items"],
                        titlu=titlu,
                        proiect=proiect,
                        data=data["data_livrare"],
                        nr_document=data["nr_intern"],
                        output_path=str(output_path),
                    )

                st.success("Fișa Accesorii a fost generată.")

                st.download_button(
                    label="⬇️ Descarcă Fișa Accesorii Excel",
                    data=output_path.read_bytes(),
                    file_name="fisa_accesorii.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    type="primary",
                )

        except Exception as e:
            st.error(f"A apărut o eroare: {e}")
            st.exception(e)
