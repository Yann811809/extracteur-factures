import streamlit as st
import pandas as pd
import requests
import os
import urllib.parse
import re
import zipfile
import io
from concurrent.futures import ThreadPoolExecutor

# === CONFIG ===
COL_URL = "URL"
COL_INVOICE = "Num Facture"
COL_FIRSTNAME = "Prénom"
COL_LASTNAME = "Nom"
MAX_WORKERS = 5

# === PAGE ===
st.set_page_config(
    page_title="📄 Extracteur de Factures PDF",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Extracteur de Factures PDF")
st.markdown("Téléchargez votre fichier Excel, puis récupérez tous les PDF en un clic.")

# === SESSION ===
session = requests.Session()
session.headers.update({
    "User-Agent": "Mozilla/5.0",
    "Accept": "application/pdf"
})

# === FONCTIONS ===
def build_pdf_url(url):
    base = url.split("/invoices/")[1]
    invoice_id = base.split("?")[0]
    key = base.split("key=")[1]
    inner_url = f"https://www.free2move.com/invoice/print/invoices/{invoice_id}?key={key}"
    encoded = urllib.parse.quote(inner_url, safe="")
    return f"https://www.free2move.com/api/media/{encoded}"

def clean(text):
    text = str(text)
    text = re.sub(r"[^\w\-]", "_", text)
    return text.strip("_")

def download_row(row):
    try:
        url = row[COL_URL]
        invoice = clean(row[COL_INVOICE])
        firstname = clean(row[COL_FIRSTNAME])
        lastname = clean(row[COL_LASTNAME])
        filename = f"{invoice}_{firstname}_{lastname}.pdf"
        pdf_url = build_pdf_url(url)
        response = session.get(pdf_url, timeout=15)

        if "application/pdf" in response.headers.get("Content-Type", ""):
            return filename, response.content, "✅ Succès"
        else:
            return filename, None, "⚠️ Réponse non-PDF"
    except Exception as e:
        return f"erreur", None, f"❌ Erreur : {e}"

# === UPLOAD ===
uploaded_file = st.file_uploader(
    "📂 Chargez votre fichier Excel (.xlsx)",
    type=["xlsx"]
)

if uploaded_file:
    df = pd.read_excel(uploaded_file)
    st.success(f"✅ Fichier chargé : **{len(df)} lignes** détectées")

    # Vérification des colonnes
    required_cols = [COL_URL, COL_INVOICE, COL_FIRSTNAME, COL_LASTNAME]
    missing = [c for c in required_cols if c not in df.columns]

    if missing:
        st.error(f"❌ Colonnes manquantes dans votre Excel : `{'`, `'.join(missing)}`")
        st.info(f"Colonnes trouvées : `{'`, `'.join(df.columns.tolist())}`")
        st.stop()

    st.dataframe(df[[COL_INVOICE, COL_FIRSTNAME, COL_LASTNAME]].head(10), use_container_width=True)

    if st.button("🚀 Lancer l'extraction", type="primary"):

        results_log = []
        pdf_files = {}

        progress_bar = st.progress(0, text="Démarrage...")
        status_box = st.empty()

        rows = [row for _, row in df.iterrows()]
        total = len(rows)

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {executor.submit(download_row, row): i for i, row in enumerate(rows)}

            completed = 0
            for future in futures:
                filename, content, status = future.result()
                completed += 1
                progress_bar.progress(completed / total, text=f"Traitement {completed}/{total}...")
                results_log.append({"Fichier": filename, "Statut": status})
                if content:
                    pdf_files[filename] = content

        progress_bar.progress(1.0, text="✅ Terminé !")

        # Résumé
        nb_ok = sum(1 for r in results_log if "✅" in r["Statut"])
        nb_err = len(results_log) - nb_ok

        col1, col2, col3 = st.columns(3)
        col1.metric("Total", total)
        col2.metric("✅ Succès", nb_ok)
        col3.metric("❌ Erreurs", nb_err)

        # Log détaillé
        with st.expander("📋 Voir le détail des résultats"):
            st.dataframe(pd.DataFrame(results_log), use_container_width=True)

        # Téléchargement ZIP
        if pdf_files:
            zip_buffer = io.BytesIO()
            with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zf:
                for fname, content in pdf_files.items():
                    zf.writestr(fname, content)
            zip_buffer.seek(0)

            st.download_button(
                label=f"⬇️ Télécharger les {len(pdf_files)} PDF (ZIP)",
                data=zip_buffer,
                file_name="factures.zip",
                mime="application/zip",
                type="primary"
            )
        else:
            st.warning("Aucun PDF n'a pu être téléchargé.")
