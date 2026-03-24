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
MAX_WORKERS = 5

COL_CANDIDATES = {
    "url":       ["URL", "Url", "url", "Link", "Lien", "Invoice link"],
    "invoice":   ["Num Facture", "Invoice Number", "Invoice", "Facture", "Num Invoice", "Invoice number"],
    "firstname": ["Prénom", "Firstname", "First Name", "First_Name", "Prenom", "Fisrt name"],
    "lastname":  ["Nom", "Lastname", "Last Name", "Last_Name", "Name", "Last name"],
}

def detect_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

# === PAGE ===
st.set_page_config(
    page_title="📄 Extracteur de Factures PDF",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Extracteur de Factures PDF")
st.markdown("Téléchargez votre fichier Excel, puis récupérez tous les PDF en un clic.")

st.info("""
**📋 Comment préparer votre fichier Excel ?**

➡️ Pour récupérer les URL : connectez-vous sur [free2move.com](https://www.free2move.com), \
Afin de savoir comment accéder à “l'Export pour la gestion comptable des factures", il faut suivre les étapes suivantes : \
•	Aller dans le menu “Location de voiture”, onglet “Exports” et choisir l’export “Invoices PSA” \
•	Remplir les informations nécessaires et exporter \
 \
•	L'export vous est transmis par notification, où vous pouvez le télécharger. 

Votre fichier `.xlsx` doit contenir exactement ces 4 colonnes :

| Colonne | Description |
|---|---|
| `URL` | Lien vers la facture sur Free2Move |
| `Num Facture` | Numéro de la facture |
| `Prénom` | Prénom du client |
| `Nom` | Nom du client |


""")

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

    # Détection automatique des colonnes (FR ou EN)
    COL_URL       = detect_col(df, COL_CANDIDATES["url"])
    COL_INVOICE   = detect_col(df, COL_CANDIDATES["invoice"])
    COL_FIRSTNAME = detect_col(df, COL_CANDIDATES["firstname"])
    COL_LASTNAME  = detect_col(df, COL_CANDIDATES["lastname"])

    missing = [label for label, val in {
        "URL / Link": COL_URL,
        "Num Facture / Invoice": COL_INVOICE,
        "Prénom / Firstname": COL_FIRSTNAME,
        "Nom / Lastname": COL_LASTNAME
    }.items() if val is None]

    if missing:
        st.error(f"❌ Colonnes non détectées : `{'`, `'.join(missing)}`")
        st.info(f"Colonnes trouvées dans le fichier : `{'`, `'.join(df.columns.tolist())}`")
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