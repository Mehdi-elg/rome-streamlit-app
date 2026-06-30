import streamlit as st
import requests
import pandas as pd
import io
from io import BytesIO
from openpyxl.utils import get_column_letter

# --- CONFIGURATION & SECRETS ---
CLIENT_ID     = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]
TOKEN_URL     = "https://entreprise.francetravail.fr/connexion/oauth2/access_token?realm=/partenaire"
API_BASE      = "https://api.francetravail.io/partenaire/rome-metiers"
# Ajout du scope nomenclature pour lister les métiers
SCOPES        = "nomenclatureRome api_rome-metiersv1"

# --- LOGIQUE API ---

def get_token():
    data = {
        "grant_type": "client_credentials",
        "client_id": CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope": SCOPES,
    }
    r = requests.post(TOKEN_URL, data=data)
    r.raise_for_status()
    return r.json()["access_token"]

def get_all_rome_referentiel(headers):
    """Récupère la liste de tous les codes ROME existants"""
    url = f"{API_BASE}/v1/metiers/metier" # Endpoint liste
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    return r.json() # Retourne une liste de dictionnaires {'code': '...', 'libelle': '...'}

def get_metier_details(headers, code_rome):
    """Récupère les détails (contextes) d'un métier spécifique"""
    champs = "code,libelle,contextestravail(categorie,libelle),"
    url = f"{API_BASE}/v1/metiers/metier/{code_rome}"
    params = {"champs": champs}
    r = requests.get(url, headers=headers, params=params)
    r.raise_for_status()
    return r.json()

# --- TRAITEMENT DES DONNÉES ---

def get_contextes_by_categorie(metier, categorie):
    contextes = []
    if 'contextesTravail' in metier:
        for ctx in metier['contextesTravail']:
            if ctx.get('categorie') == categorie:
                libelle = ctx.get('libelle', '').strip()
                if libelle:
                    contextes.append(libelle)
    return contextes

def is_fipu(conditions_str: str, horaires_str: str) -> bool:
    if not conditions_str and not horaires_str:
        return False
    criteres = [
        "En altitude", "En milieu nucléaire", "En milieu hyperbare", "En milieu exigu ou confiné",
        "En grande hauteur", "En zone frigorifique", "Exposition à de hautes températures",
        "En environnement climatique difficile", "Manipulation d'un engin, équipement ou outil dangereux",
        "Port et manipulation de charges lourdes ou encombrantes", "Position pénible",
        "Station debout prolongée", "Travail répétitif ou cadence imposée", "En environnement bruyant",
        "Travail dans des environnements hostiles et dangereux", "Exposition à de basses températures",
        "Exposition possible à gaz, aérosol, fumées …", "Station assise prolongée", "Risques de chutes",
        "Travail dans des milieux difficiles et exigeants pour l'humain", "Travail posté (2x8, 3x8, 5x8, etc.)",
        "Travail de nuit", "Travail en astreinte", "Travail en horaires décalés", "Travail par roulement"
    ]
    texte = (conditions_str + " " + horaires_str).lower()
    return any(critere.lower() in texte for critere in criteres)

def build_enriched_df(metiers_list):
    rows = []
    for metier in metiers_list:
        conditions = get_contextes_by_categorie(metier, "CONDITIONS_TRAVAIL")
        horaires   = get_contextes_by_categorie(metier, "HORAIRE_ET_DUREE_TRAVAIL")
        cond_joined = ', '.join(conditions) if conditions else ''
        hor_joined  = ', '.join(horaires) if horaires else ''
        
        rows.append({
            'Code ROME': metier.get('code'),
            'Libellé': metier.get('libelle'),
            'FIPU': "OUI" if is_fipu(cond_joined, hor_joined) else "NON",
            'Conditions de travail': cond_joined,
            'Horaires et durée': hor_joined
        })
    return pd.DataFrame(rows)

def df_to_excel_bytes(df):
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name="Metiers_ROME")
        ws = writer.sheets["Metiers_ROME"]
        for col_idx, column_cells in enumerate(ws.columns, start=1):
            column_letter = get_column_letter(col_idx)
            ws.column_dimensions[column_letter].width = 30
        ws.freeze_panes = "A2"
    buffer.seek(0)
    return buffer

# --- INTERFACE STREAMLIT ---

def main():
    st.set_page_config(page_title="ROME & FIPU Explorer", layout="wide")
    st.title("🔎 Explorateur de Métiers ROME & FIPU")

    try:
        token = get_token()
        headers = {"Authorization": f"Bearer {token}"}
    except Exception as e:
        st.error(f"Erreur d'authentification : {e}")
        return

    tab_all, tab_list = st.tabs(["Extraction complète (Référentiel)", "Recherche par liste"])

    # --- ONGLET 1 : EXTRACTION COMPLÈTE ---
    with tab_all:
        st.subheader("Extraction de tous les métiers du ROME")
        st.info("Cette option récupère l'intégralité des métiers (approx. 530) pour analyser le FIPU. Cela peut prendre 2-3 minutes.")
        
        if "df_all_fipu" not in st.session_state: st.session_state.df_all_fipu = None
        if "stop_all" not in st.session_state: st.session_state.stop_all = False

        c1, c2 = st.columns(2)
        btn_run = c1.button("🚀 Lancer l'extraction complète")
        btn_stop = c2.button("🛑 STOP", type="primary")

        if btn_stop:
            st.session_state.stop_all = True

        prog_txt = st.empty()
        prog_bar = st.progress(0)

        if btn_run:
            st.session_state.stop_all = False
            try:
                referentiel = get_all_rome_referentiel(headers)
                total = len(referentiel)
                data_accumulated = []
                
                for i, item in enumerate(referentiel):
                    if st.session_state.stop_all:
                        st.warning("Extraction interrompue par l'utilisateur.")
                        break
                    
                    code = item['code']
                    try:
                        details = get_metier_details(headers, code)
                        data_accumulated.append(details)
                    except:
                        pass # On ignore les erreurs individuelles
                    
                    prog_bar.progress((i + 1) / total)
                    prog_txt.text(f"Traitement : {i+1}/{total} ({code})")
                
                st.session_state.df_all_fipu = build_enriched_df(data_accumulated)
            except Exception as e:
                st.error(f"Erreur lors de l'extraction : {e}")

        if st.session_state.df_all_fipu is not None:
            df = st.session_state.df_all_fipu
            st.success(f"Extraction terminée : {len(df)} métiers récupérés.")
            st.dataframe(df, use_container_width=True, height=400)
            
            excel_data = df_to_excel_bytes(df)
            st.download_button(
                "📥 Télécharger l'extraction complète (Excel)",
                data=excel_data,
                file_name="ROME_FIPU_complet.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

    # --- ONGLET 2 : RECHERCHE PAR LISTE ---
    with tab_list:
        st.subheader("Recherche ciblée")
        input_codes = st.text_area("Entrez vos codes ROME (un par ligne)", height=200, placeholder="A1413\nM1805")
        
        if st.button("🔍 Rechercher la sélection"):
            codes = [c.strip().upper() for c in input_codes.splitlines() if c.strip()]
            if not codes:
                st.warning("Veuillez entrer au moins un code.")
            else:
                results = []
                p_bar = st.progress(0)
                for i, c in enumerate(codes):
                    try:
                        res = get_metier_details(headers, c)
                        results.append(res)
                    except:
                        st.error(f"Code {c} non trouvé.")
                    p_bar.progress((i+1)/len(codes))
                
                if results:
                    df_list = build_enriched_df(results)
                    st.session_state.df_list_fipu = df_list
                    st.dataframe(df_list, use_container_width=True)
                    
                    excel_list = df_to_excel_bytes(df_list)
                    st.download_button(
                        "📥 Télécharger la sélection (Excel)",
                        data=excel_list,
                        file_name="ROME_FIPU_selection.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

if __name__ == "__main__":
    main()
