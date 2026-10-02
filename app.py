import streamlit as st
import pandas as pd
import requests
import urllib.parse
import re
import zipfile
import io

from concurrent.futures import ThreadPoolExecutor, as_completed


# ============================================================
# CONFIGURATION
# ============================================================

MAX_WORKERS = 5
REQUEST_TIMEOUT = 30


COL_CANDIDATES = {
    "url": [
        "URL",
        "Url",
        "url",
        "Link",
        "Lien",
        "Invoice link"
    ],
    "invoice": [
        "Num Facture",
        "Invoice Number",
        "Invoice",
        "Facture",
        "Num Invoice",
        "Invoice number"
    ],
    "firstname": [
        "Prénom",
        "Firstname",
        "First Name",
        "First_Name",
        "Prenom",
        "Fisrt name"
    ],
    "lastname": [
        "Nom",
        "Lastname",
        "Last Name",
        "Last_Name",
        "Name",
        "Last name"
    ],
    "company": [
        "Principal",
        "Societe",
        "Société",
        "Company",
        "Entreprise",
        "Agency",
        "Agence Résa"
    ],
}


# ============================================================
# TABLE DE CORRESPONDANCE DES SOCIÉTÉS
# ============================================================

COMPANY_MAP = {
    "GARAGE MODERNE SAS - Citroën Rent & Smile - GARAGE MODERNE SAS - CHALON SUR SAONE":
        "Chalon_AC",

    "GARAGE MODERNE SAS - Citroën Rent & Smile - GARAGE MODERNE SAS - MACON":
        "Macon_AC",

    "GARAGE MODERNE SAS - DS Rent - GARAGE MODERNE SAS - MACON":
        "Macon_DS",

    "NOMBLOT SAS - Peugeot Rent - NOMBLOT VILLEFRANCHE":
        "Villefranche_AP",

    "NOMBLOT VILLEFRANCHE - Free2move (C) VILLEFRANCHE-SUR-SAONE":
        "Villefranche_AC",

    "NOMBLOT VILLEFRANCHE - Free2move (F) VILLEFRANCHE-SUR-SAONE":
        "Villefranche_Fiat",

    "NOMBLOT VILLEFRANCHE - Free2move (J) VILLEFRANCHE-SUR-SAONE":
        "Villefranche_Jeep",

    "NOMBLOT VILLEFRANCHE - Free2move (O) VILLEFRANCHE-SUR-SAONE":
        "Villefranche_Opel",

    "FREE2MOVE RENT - NOMBLOT AUTOMOBILES SAS (C) VILLEFRANCHE S/SAONE CEDEX":
        "Villefranche_AC",

    "NOMBLOT CHALON - Free2move (C) CHALON SUR SAONE":
        "Chalon_AC",

    "NOMBLOT CHALON - Free2move (D) Chalon sur Saone":
        "Chalon_DS",
}


# ============================================================
# CONFIGURATION DE LA PAGE STREAMLIT
# ============================================================

st.set_page_config(
    page_title="Extracteur de factures PDF",
    page_icon="📄",
    layout="centered"
)

st.title("📄 Extracteur de factures PDF")

st.markdown(
    """
    Chargez votre export Excel Free2Move, puis téléchargez toutes les
    factures disponibles dans une seule archive ZIP.
    """
)


st.info(
    """
**📋 Comment préparer votre fichier Excel ?**

Connectez-vous sur [Free2Move Nimda](https://free2move.rent/nimda/).

Pour récupérer l'export comptable :

1. Aller dans le menu **Location de voiture**
2. Ouvrir l'onglet **Exports**
3. Choisir l'export **Invoices PSA**
4. Remplir les informations nécessaires
5. Lancer l'export
6. Télécharger le fichier depuis la notification reçue

Le fichier `.xlsx` doit contenir les colonnes suivantes :

| Colonne | Description |
|---|---|
| `URL` | Lien vers la facture Free2Move |
| `Num Facture` | Numéro de la facture |
| `Prénom` | Prénom du client |
| `Nom` | Nom du client |
| `Principal` ou `Societe` | Nom de la société |
"""
)


# ============================================================
# FONCTIONS UTILITAIRES
# ============================================================

def detect_col(df, candidates):
    """
    Recherche la première colonne existante parmi les noms proposés.
    """

    for candidate in candidates:
        if candidate in df.columns:
            return candidate

    return None


def clean(text):
    """
    Nettoie une valeur pour pouvoir l'utiliser dans un nom de fichier.
    """

    if pd.isna(text):
        return ""

    text = str(text).strip()

    # Remplace les caractères interdits ou gênants par un underscore.
    text = re.sub(r"[^\w\-]", "_", text, flags=re.UNICODE)

    # Supprime les underscores multiples.
    text = re.sub(r"_+", "_", text)

    return text.strip("_")


def get_safe_value(value, default_value):
    """
    Retourne une valeur nettoyée ou une valeur par défaut.
    """

    cleaned_value = clean(value)

    if not cleaned_value:
        return default_value

    return cleaned_value


def get_company_short(raw_company):
    """
    Retourne le raccourci société défini dans COMPANY_MAP.

    Si la société n'existe pas dans la table, son nom est nettoyé et
    limité à 40 caractères.
    """

    if pd.isna(raw_company):
        return "Inconnu"

    raw_company = str(raw_company).strip()

    if not raw_company:
        return "Inconnu"

    if raw_company in COMPANY_MAP:
        return COMPANY_MAP[raw_company]

    cleaned_company = clean(raw_company)[:40]

    return cleaned_company or "Inconnu"


def make_unique_filename(filename, existing_files):
    """
    Évite l'écrasement si deux factures produisent le même nom de fichier.
    """

    if filename not in existing_files:
        return filename

    base_name = filename.rsplit(".", 1)[0]
    extension = filename.rsplit(".", 1)[1]

    counter = 2

    while True:
        new_filename = f"{base_name}_{counter}.{extension}"

        if new_filename not in existing_files:
            return new_filename

        counter += 1


# ============================================================
# CONSTRUCTION DE L'URL PDF
# ============================================================

def build_pdf_url(source_url):
    """
    Construit l'URL de téléchargement du PDF Free2Move.

    Exemple de résultat :

    https://free2move.rent/api/media/
    https://free2move.rent/invoice/print/invoices/
    IDENTIFIANT%3Fkey=CLE

    Le paramètre modified n'est volontairement pas ajouté.
    """

    if pd.isna(source_url):
        raise ValueError("URL vide")

    source_url = str(source_url).strip()

    if not source_url:
        raise ValueError("URL vide")

    # Si le fichier Excel contient déjà une URL /api/media/,
    # on retire éventuellement le paramètre modified.
    if "/api/media/" in source_url:
        source_url = re.sub(
            r"([?&])modified=[^&]*",
            "",
            source_url,
            flags=re.IGNORECASE
        )

        source_url = source_url.rstrip("?&")

        return source_url

    # Décodage limité pour prendre en charge les URL contenant %3Fkey.
    decoded_url = urllib.parse.unquote(source_url)

    parsed_url = urllib.parse.urlparse(decoded_url)

    # Recherche de l'identifiant situé après /invoices/.
    invoice_match = re.search(
        r"/invoices/([^/?&]+)",
        parsed_url.path
    )

    if not invoice_match:
        # Solution de secours si la structure de l'URL est inhabituelle.
        invoice_match = re.search(
            r"/invoices/([^/?&]+)",
            decoded_url
        )

    if not invoice_match:
        raise ValueError(
            "Identifiant de facture introuvable dans l'URL"
        )

    invoice_id = invoice_match.group(1).strip()

    # Recherche normale du paramètre key.
    query_params = urllib.parse.parse_qs(
        parsed_url.query,
        keep_blank_values=True
    )

    key = query_params.get("key", [None])[0]

    # Solution de secours pour les URL contenant ?key= ou %3Fkey=.
    if not key:
        key_match = re.search(
            r"(?:\?|%3F)key=([^&]+)",
            source_url,
            flags=re.IGNORECASE
        )

        if key_match:
            key = urllib.parse.unquote(key_match.group(1))

    if not key:
        raise ValueError(
            "Clé de téléchargement 'key' introuvable dans l'URL"
        )

    invoice_id = urllib.parse.quote(
        invoice_id,
        safe=""
    )

    key = urllib.parse.quote(
        str(key).strip(),
        safe=""
    )

    return (
        "https://free2move.rent/api/media/"
        "https://free2move.rent/invoice/print/invoices/"
        f"{invoice_id}%3Fkey={key}"
    )


# ============================================================
# TÉLÉCHARGEMENT D'UNE FACTURE
# ============================================================

def download_row(
    row,
    row_number,
    col_url,
    col_invoice,
    col_firstname,
    col_lastname,
    col_company
):
    """
    Télécharge une facture et retourne son contenu ainsi qu'un journal
    détaillé du traitement.
    """

    source_url = ""

    try:
        source_url = row[col_url]

        if pd.isna(source_url) or not str(source_url).strip():
            raise ValueError("URL de facture vide")

        invoice = get_safe_value(
            row[col_invoice],
            "SANS_NUMERO"
        )

        firstname = get_safe_value(
            row[col_firstname],
            "SANS_PRENOM"
        )

        lastname = get_safe_value(
            row[col_lastname],
            "SANS_NOM"
        )

        if col_company:
            company = get_company_short(row[col_company])
        else:
            company = "Inconnu"

        filename = (
            f"{company}_{invoice}_{firstname}_{lastname}.pdf"
        )

        pdf_url = build_pdf_url(source_url)

        headers = {
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) "
                "Chrome/140.0.0.0 Safari/537.36"
            ),
            "Accept": (
                "application/pdf,"
                "application/octet-stream;q=0.9,"
                "*/*;q=0.8"
            ),
            "Referer": "https://free2move.rent/",
        }

        response = requests.get(
            pdf_url,
            headers=headers,
            timeout=REQUEST_TIMEOUT,
            allow_redirects=True
        )

        status_code = response.status_code

        content_type = response.headers.get(
            "Content-Type",
            ""
        ).lower()

        is_pdf = (
            "application/pdf" in content_type
            or response.content.startswith(b"%PDF")
        )

        if response.ok and is_pdf:
            return {
                "row_number": row_number,
                "filename": filename,
                "content": response.content,
                "status": "✅ Succès",
                "http_code": status_code,
                "content_type": content_type,
                "details": ""
            }

        try:
            response_preview = response.text[:250]

            response_preview = re.sub(
                r"\s+",
                " ",
                response_preview
            ).strip()

        except Exception:
            response_preview = "Contenu de la réponse illisible"

        details = (
            f"Réponse reçue : {response_preview}"
            if response_preview
            else "Le serveur n'a retourné aucun contenu exploitable"
        )

        return {
            "row_number": row_number,
            "filename": filename,
            "content": None,
            "status": "⚠️ Réponse non-PDF",
            "http_code": status_code,
            "content_type": content_type or "Type inconnu",
            "details": details
        }

    except requests.Timeout:
        return {
            "row_number": row_number,
            "filename": "Non téléchargé",
            "content": None,
            "status": "❌ Délai dépassé",
            "http_code": "",
            "content_type": "",
            "details": (
                f"Le serveur n'a pas répondu sous "
                f"{REQUEST_TIMEOUT} secondes"
            )
        }

    except requests.HTTPError as error:
        status_code = ""

        if error.response is not None:
            status_code = error.response.status_code

        return {
            "row_number": row_number,
            "filename": "Non téléchargé",
            "content": None,
            "status": "❌ Erreur HTTP",
            "http_code": status_code,
            "content_type": "",
            "details": str(error)
        }

    except requests.RequestException as error:
        return {
            "row_number": row_number,
            "filename": "Non téléchargé",
            "content": None,
            "status": "❌ Erreur réseau",
            "http_code": "",
            "content_type": "",
            "details": str(error)
        }

    except Exception as error:
        return {
            "row_number": row_number,
            "filename": "Non téléchargé",
            "content": None,
            "status": "❌ Erreur",
            "http_code": "",
            "content_type": "",
            "details": str(error)
        }


# ============================================================
# IMPORT DU FICHIER EXCEL
# ============================================================

uploaded_file = st.file_uploader(
    "📂 Chargez votre fichier Excel",
    type=["xlsx"],
    help="Sélectionnez l'export Excel généré par Free2Move."
)


if uploaded_file is not None:

    try:
        df = pd.read_excel(
            uploaded_file,
            engine="openpyxl"
        )

    except Exception as error:
        st.error(
            f"Impossible de lire le fichier Excel : {error}"
        )
        st.stop()

    if df.empty:
        st.warning("Le fichier Excel ne contient aucune ligne.")
        st.stop()

    # Nettoyage des noms de colonnes.
    df.columns = [
        str(column).strip()
        for column in df.columns
    ]

    st.success(
        f"✅ Fichier chargé : **{len(df)} ligne(s)** détectée(s)"
    )

    # Détection automatique des colonnes.
    COL_URL = detect_col(
        df,
        COL_CANDIDATES["url"]
    )

    COL_INVOICE = detect_col(
        df,
        COL_CANDIDATES["invoice"]
    )

    COL_FIRSTNAME = detect_col(
        df,
        COL_CANDIDATES["firstname"]
    )

    COL_LASTNAME = detect_col(
        df,
        COL_CANDIDATES["lastname"]
    )

    COL_COMPANY = detect_col(
        df,
        COL_CANDIDATES["company"]
    )

    detected_columns = {
        "URL / Link": COL_URL,
        "Num Facture / Invoice": COL_INVOICE,
        "Prénom / Firstname": COL_FIRSTNAME,
        "Nom / Lastname": COL_LASTNAME,
    }

    missing_columns = [
        label
        for label, column_name in detected_columns.items()
        if column_name is None
    ]

    if missing_columns:
        st.error(
            "❌ Colonnes obligatoires non détectées : "
            f"`{'`, `'.join(missing_columns)}`"
        )

        st.info(
            "Colonnes trouvées dans le fichier : "
            f"`{'`, `'.join(df.columns.tolist())}`"
        )

        st.stop()

    if COL_COMPANY is None:
        st.warning(
            "⚠️ Colonne société non détectée. "
            "Le préfixe `Inconnu` sera utilisé."
        )

    # Résumé des colonnes détectées.
    with st.expander("🔎 Voir les colonnes détectées"):
        st.write(f"URL : `{COL_URL}`")
        st.write(f"Numéro de facture : `{COL_INVOICE}`")
        st.write(f"Prénom : `{COL_FIRSTNAME}`")
        st.write(f"Nom : `{COL_LASTNAME}`")
        st.write(
            f"Société : `{COL_COMPANY or 'Non détectée'}`"
        )

    # Aperçu du fichier.
    st.subheader("Aperçu des données")

    preview_columns = [
        column
        for column in [
            COL_COMPANY,
            COL_INVOICE,
            COL_FIRSTNAME,
            COL_LASTNAME
        ]
        if column
    ]

    st.dataframe(
        df[preview_columns].head(10),
        use_container_width=True
    )

    # Détection des sociétés absentes de la table.
    if COL_COMPANY:
        companies = (
            df[COL_COMPANY]
            .dropna()
            .astype(str)
            .str.strip()
            .unique()
        )

        unmapped_companies = [
            company
            for company in companies
            if company and company not in COMPANY_MAP
        ]

        if unmapped_companies:
            with st.expander(
                "⚠️ "
                f"{len(unmapped_companies)} société(s) absente(s) "
                "de la table de correspondance"
            ):
                st.caption(
                    "Le nom brut nettoyé sera utilisé dans le nom du PDF."
                )

                for company in unmapped_companies:
                    st.markdown(f"- `{company}`")

    # Détection des factures sans numéro.
    missing_invoice_count = int(
        df[COL_INVOICE].isna().sum()
    )

    if missing_invoice_count:
        st.warning(
            f"⚠️ {missing_invoice_count} ligne(s) sans numéro de facture. "
            "Le texte `SANS_NUMERO` sera utilisé."
        )

    # Test de la première URL.
    with st.expander("🧪 Vérifier la première URL PDF"):
        try:
            first_source_url = df.iloc[0][COL_URL]
            first_pdf_url = build_pdf_url(first_source_url)

            st.write("URL PDF générée :")
            st.code(
                first_pdf_url,
                language=None
            )

            st.link_button(
                "🔗 Ouvrir le PDF de test",
                first_pdf_url
            )

            st.caption(
                "La clé présente dans l'URL peut donner accès à une facture. "
                "Ne partagez pas ce lien."
            )

        except Exception as error:
            st.error(
                f"Impossible de construire la première URL : {error}"
            )

    # ========================================================
    # LANCEMENT DE L'EXTRACTION
    # ========================================================

    if st.button(
        "🚀 Lancer l'extraction",
        type="primary",
        use_container_width=True
    ):

        results_log = []
        pdf_files = {}

        rows = [
            (index, row)
            for index, row in df.iterrows()
        ]

        total = len(rows)

        progress_bar = st.progress(
            0,
            text="Préparation de l'extraction..."
        )

        status_placeholder = st.empty()

        with ThreadPoolExecutor(
            max_workers=MAX_WORKERS
        ) as executor:

            futures = {
                executor.submit(
                    download_row,
                    row,
                    index + 2,
                    COL_URL,
                    COL_INVOICE,
                    COL_FIRSTNAME,
                    COL_LASTNAME,
                    COL_COMPANY
                ): index
                for index, row in rows
            }

            completed = 0

            for future in as_completed(futures):
                result = future.result()

                completed += 1

                progress_bar.progress(
                    completed / total,
                    text=(
                        f"Traitement des factures : "
                        f"{completed}/{total}"
                    )
                )

                if result["content"]:
                    unique_filename = make_unique_filename(
                        result["filename"],
                        pdf_files
                    )

                    pdf_files[unique_filename] = result["content"]
                    result["filename"] = unique_filename

                results_log.append({
                    "Ligne Excel": result["row_number"],
                    "Fichier": result["filename"],
                    "Statut": result["status"],
                    "HTTP": result["http_code"],
                    "Type de contenu": result["content_type"],
                    "Détails": result["details"]
                })

                status_placeholder.caption(
                    f"Dernier traitement : {result['status']} "
                    f"• {result['filename']}"
                )

        progress_bar.progress(
            1.0,
            text="✅ Extraction terminée"
        )

        status_placeholder.empty()

        # Trie les résultats selon la ligne Excel d'origine.
        results_log = sorted(
            results_log,
            key=lambda result: result["Ligne Excel"]
        )

        # ====================================================
        # RÉSUMÉ
        # ====================================================

        success_count = sum(
            1
            for result in results_log
            if result["Statut"] == "✅ Succès"
        )

        error_count = total - success_count

        success_rate = (
            round((success_count / total) * 100, 1)
            if total
            else 0
        )

        st.subheader("Résultat de l'extraction")

        metric_total, metric_success, metric_error = st.columns(3)

        metric_total.metric(
            "Total",
            total
        )

        metric_success.metric(
            "✅ Succès",
            success_count
        )

        metric_error.metric(
            "❌ Erreurs",
            error_count
        )

        st.caption(
            f"Taux de réussite : {success_rate} %"
        )

        with st.expander(
            "📋 Voir le détail des résultats",
            expanded=error_count > 0
        ):
            st.dataframe(
                pd.DataFrame(results_log),
                use_container_width=True,
                hide_index=True
            )

        # ====================================================
        # CRÉATION ET TÉLÉCHARGEMENT DU ZIP
        # ====================================================

        if pdf_files:
            zip_buffer = io.BytesIO()

            with zipfile.ZipFile(
                zip_buffer,
                mode="w",
                compression=zipfile.ZIP_DEFLATED
            ) as zip_file:

                for filename, content in pdf_files.items():
                    zip_file.writestr(
                        filename,
                        content
                    )

            zip_buffer.seek(0)

            st.download_button(
                label=(
                    f"⬇️ Télécharger les "
                    f"{len(pdf_files)} PDF dans un ZIP"
                ),
                data=zip_buffer.getvalue(),
                file_name="factures_free2move.zip",
                mime="application/zip",
                type="primary",
                use_container_width=True
            )

        else:
            st.warning(
                "Aucun PDF n'a pu être téléchargé. "
                "Consultez le détail des résultats."
            )
