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
    "url":       ["URL", "Url", "url", "Link", "Lien","Invoice link"],
    "invoice":   ["Num Facture", "Invoice Number", "Invoice", "Facture", "Num Invoice", "Invoice number"],
    "firstname": ["Prénom", "Firstname", "First Name", "First_Name", "Prenom", "Fisrt name"],
    "lastname":  ["Nom", "Lastname", "Last Name", "Last_Name", "Name", "Last name"],
    "company":   ["Principal", "Societe", "Société", "Company", "Entreprise", "Agency", "Agence Résa"],
}

# === TABLE DE CORRESPONDANCE SOCIÉTÉS ===
COMPANY_MAP = {
    "GARAGE MODERNE SAS - Citroën Rent & Smile - GARAGE MODERNE SAS - CHALON SUR SAONE": "Chalon_AC",
    "GARAGE MODERNE SAS - Citroën Rent & Smile - GARAGE MODERNE SAS - MACON":            "Macon_AC",
    "GARAGE MODERNE SAS - DS Rent - GARAGE MODERNE SAS - MACON":                          "Macon_DS",
    "NOMBLOT SAS - Peugeot Rent - NOMBLOT VILLEFRANCHE":                                  "Villefranche_AP",
    "NOMBLOT VILLEFRANCHE - Free2move (C) VILLEFRANCHE-SUR-SAONE":                        "Villefranche_AC",
    "NOMBLOT VILLEFRANCHE - Free2move (F) VILLEFRANCHE-SUR-SAONE":                        "Villefranche_Fiat",
    "NOMBLOT VILLEFRANCHE - Free2move (J) VILLEFRANCHE-SUR-SAONE":                        "Villefranche_Jeep",
    "NOMBLOT VILLEFRANCHE - Free2move (O) VILLEFRANCHE-SUR-SAONE":                        "Villefranche_Opel",
}

def detect_col(df, candidates):
    for c in candidates:
        if c in df.columns:
            return c
    return None

def clean(text):
    text = str(text)
    text = re.sub(r"[^\w\-]", "_", text)
    return text.strip("_")

def get_company_short(raw):
    """Retourne le raccourci société, ou une version nettoyée si non trouvé dans la table."""
    raw = str(raw).strip()
    if raw in COMPANY_MAP:
        return COMPANY_MAP[raw]
    # Fallback : nettoyage brut (40 car. max)
    return clean(raw)[:40]

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

Votre fichier `.xlsx` doit contenir exactement ces 4 colonnes :

| Colonne | Description |
|---|---|
| `URL` | Lien vers la facture sur Free2Move |
| `Num Facture` | Numéro de la facture |
| `Prénom` | Prénom du client |
| `Nom` | Nom du client |
| `Principal` ou `Societe` | Nom de la société |

➡️ Pour récupérer les URL : connectez-vous sur [free2move.com](https://www.free2move.com), \
allez dans **Mes Factures**, faites un clic droit sur chaque facture → **Copier le lien**, \
et collez-le dans la colonne `URL` de votre Excel.
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

def download_row(row, col_url, col_invoice, col_firstname, col_lastname, col_company):
    try:
        url = row[col_url]
        invoice   = clean(row[col_invoice])
        firstname = clean(row[col_firstname])
        lastname  = clean(row[col_lastname])
        company   = get_company_short(row[col_company]) if col_company else "Inconnu"

        filename = f"{company}_{invoice}_{firstname}_{lastname}.pdf"
        pdf_url  = build_pdf_url(url)
        response = session.get(pdf_url, timeout=15)

        if "application/pdf" in response.headers.get("Content-Type", ""):
            return filename, response.content, "✅ Succès"
        else:
            return filename, None, "⚠️ Réponse non-PDF"
    except Exception as e:
        return "erreur", None, f"❌ Erreur : {e}"

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
    COL_COMPANY   = detect_col(df, COL_CANDIDATES["company"])

    missing = [label for label, val in {
        "URL / Link": COL_URL,
        "Num Facture / Invoice": COL_INVOICE,
        "Prénom / Firstname": COL_FIRSTNAME,
        "Nom / Lastname": COL_LASTNAME,
    }.items() if val is None]

    if missing:
        st.error(f"❌ Colonnes non détectées : `{'`, `'.join(missing)}`")
        st.info(f"Colonnes trouvées dans le fichier : `{'`, `'.join(df.columns.tolist())}`")
        st.stop()

    if COL_COMPANY is None:
        st.warning("⚠️ Colonne société non détectée — le préfixe sera remplacé par 'Inconnu'.")

    # Aperçu
    preview_cols = [c for c in [COL_COMPANY, COL_INVOICE, COL_FIRSTNAME, COL_LASTNAME] if c]
    st.dataframe(df[preview_cols].head(10), use_container_width=True)

    # Sociétés non mappées
    if COL_COMPANY:
        unmapped = df[COL_COMPANY].dropna().unique()
        unmapped = [s for s in unmapped if str(s).strip() not in COMPANY_MAP]
        if unmapped:
            with st.expander(f"⚠️ {len(unmapped)} société(s) absente(s) de la table de correspondance (nom brut utilisé)"):
                for s in unmapped:
                    st.markdown(f"- `{s}`")

    if st.button("🚀 Lancer l'extraction", type="primary"):

        results_log = []
        pdf_files   = {}

        progress_bar = st.progress(0, text="Démarrage...")
        rows  = [row for _, row in df.iterrows()]
        total = len(rows)

        with ThreadPoolExecutor(max_workers=MAX_WORKERS) as executor:
            futures = {
                executor.submit(download_row, row, COL_URL, COL_INVOICE, COL_FIRSTNAME, COL_LASTNAME, COL_COMPANY): i
                for i, row in enumerate(rows)
            }
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
        nb_ok  = sum(1 for r in results_log if "✅" in r["Statut"])
        nb_err = len(results_log) - nb_ok

        col1, col2, col3 = st.columns(3)
        col1.metric("Total", total)
        col2.metric("✅ Succès", nb_ok)
        col3.metric("❌ Erreurs", nb_err)

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