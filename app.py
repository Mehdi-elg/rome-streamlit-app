import streamlit as st
import requests
import pandas as pd
import io
from openpyxl.utils import get_column_letter

CLIENT_ID     = st.secrets["CLIENT_ID"]
CLIENT_SECRET = st.secrets["CLIENT_SECRET"]

TOKEN_URL = "https://entreprise.francetravail.fr/connexion/oauth2/access_token?realm=/partenaire"
API_BASE  = "https://api.francetravail.io/partenaire/rome-metiers"
SCOPES    = "nomenclatureRome api_rome-metiersv1"

STATS_BASE_URL = "https://api.francetravail.io/partenaire/stats-offres-demandes-emploi"

if 'search_done' not in st.session_state:
    st.session_state.search_done = False
    st.session_state.statuts = []
    st.session_state.codes_list = []


# ── Auth ──────────────────────────────────────────────────────────────────────

def get_token():
    data = {
        "grant_type":    "client_credentials",
        "client_id":     CLIENT_ID,
        "client_secret": CLIENT_SECRET,
        "scope":         SCOPES,
    }
    r = requests.post(TOKEN_URL, data=data)
    r.raise_for_status()
    return r.json()["access_token"]


# ── ROME métier ───────────────────────────────────────────────────────────────

def get_metier(code_rome):
    token = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    champs = "code,libelle,contextestravail(categorie,libelle),"
    url    = f"{API_BASE}/v1/metiers/metier/{code_rome}"
    r = requests.get(url, headers=headers, params={"champs": champs})
    r.raise_for_status()
    return r.json()


def get_contextes_by_categorie(metier, categorie):
    contextes = []
    if 'contextesTravail' in metier:
        for ctx in metier['contextesTravail']:
            if ctx.get('categorie') == categorie:
                libelle = ctx.get('libelle', '').strip()
                if libelle:
                    contextes.append(libelle)
    return contextes


def flatten_dict(d, parent_key='', sep='_'):
    items = []
    for k, v in d.items():
        new_key = f"{parent_key}{sep}{k}" if parent_key else k
        if isinstance(v, dict):
            items.extend(flatten_dict(v, new_key, sep=sep).items())
        elif isinstance(v, list):
            for i, item in enumerate(v):
                if isinstance(item, dict):
                    items.extend(flatten_dict(item, f"{new_key}_{i}", sep=sep).items())
                else:
                    items.append((f"{new_key}_{i}", item))
        else:
            items.append((new_key, v))
    return dict(items)


CRITERES_FIPU = [
    "En altitude", "En milieu nucléaire", "En milieu hyperbare",
    "En milieu exigu ou confiné", "En grande hauteur", "En zone frigorifique",
    "Exposition à de hautes températures", "En environnement climatique difficile",
    "Manipulation d'un engin, équipement ou outil dangereux",
    "Port et manipulation de charges lourdes ou encombrantes",
    "Position pénible", "Station debout prolongée",
    "Travail répétitif ou cadence imposée", "En environnement bruyant",
    "Travail dans des environnements hostiles et dangereux",
    "Exposition à de basses températures",
    "Exposition possible à gaz, aérosol, fumées …",
    "Station assise prolongée", "Risques de chutes",
    "Travail dans des milieux difficiles et exigeants pour l'humain",
    "Travail posté (2x8, 3x8, 5x8, etc.)", "Travail de nuit",
    "Travail en astreinte", "Travail en horaires décalés",
    "Travail par roulement",
]


def is_fipu(conditions_str: str, horaires_str: str) -> bool:
    if not conditions_str and not horaires_str:
        return False
    texte = (conditions_str + " " + horaires_str).lower()
    return any(c.lower() in texte for c in CRITERES_FIPU)


def create_enriched_df(metiers_data):
    rows = []
    for metier in metiers_data:
        flat       = flatten_dict(metier)
        conditions = get_contextes_by_categorie(metier, "CONDITIONS_TRAVAIL")
        horaires   = get_contextes_by_categorie(metier, "HORAIRE_ET_DUREE_TRAVAIL")
        cond_str   = ', '.join(conditions) if conditions else ''
        hor_str    = ', '.join(horaires)   if horaires   else ''
        flat['Conditions de travail et risques professionnels'] = cond_str
        flat['Horaires et durée du travail']                    = hor_str
        flat['FIPU'] = "OUI" if is_fipu(cond_str, hor_str) else "NON"
        rows.append(flat)
    df = pd.DataFrame(rows)
    desired = ['code', 'libelle', 'FIPU',
               'Conditions de travail et risques professionnels',
               'Horaires et durée du travail']
    return df[[c for c in desired if c in df.columns]]


def to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Metiers_ROME") -> bytes:
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name=sheet_name, index=False)
        ws = writer.sheets[sheet_name]
        for col_idx, _ in enumerate(ws.columns, start=1):
            col_letter   = get_column_letter(col_idx)
            header_value = ws[f"{col_letter}1"].value
            if header_value:
                ws.column_dimensions[col_letter].width = min(len(str(header_value)) + 5, 80)
        ws.freeze_panes = "A2"
    buf.seek(0)
    return buf.getvalue()


# ── Référentiel ROME (pour l'extraction complète) ─────────────────────────────

def get_all_rome_codes():
    """Récupère la liste de tous les codes ROME existants."""
    token   = get_token()
    headers = {"Authorization": f"Bearer {token}"}
    url     = f"{API_BASE}/v1/metiers/metier"
    r = requests.get(url, headers=headers)
    r.raise_for_status()
    data = r.json()
    # L'API renvoie soit une liste directe, soit un objet avec une clé
    items = data if isinstance(data, list) else data.get("metiers", data.get("results", []))
    codes = []
    for item in items:
        code = item.get("code") if isinstance(item, dict) else item
        if code:
            codes.append(code)
    return codes


# ── UI ────────────────────────────────────────────────────────────────────────

st.title("🔎 Recherche Multi-Métiers ROME")

tab_list, tab_all = st.tabs(
    ["Recherche par liste de codes ROME", "Extraction complète (tous les ROME)"]
)

# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 1 — Recherche ciblée (code d'origine, inchangé)
# ══════════════════════════════════════════════════════════════════════════════
with tab_list:
    st.markdown("**Entrez plusieurs codes ROME (1 par ligne) et consultez les résultats détaillés**")

    codes_input = st.text_area(
        "Codes ROME (un par ligne, ex: A1413\nM1805\nH1203)",
        height=150,
        placeholder="A1413\nM1805\nH1203",
    )

    if st.button("🔍 Rechercher TOUS les métiers", type="primary"):
        if not codes_input.strip():
            st.warning("⚠️ Veuillez entrer au moins un code ROME.")
        else:
            seen, codes_list = set(), []
            for line in codes_input.strip().split('\n'):
                code = line.strip().upper()
                if code and code not in seen:
                    seen.add(code)
                    codes_list.append(code)

            original_len = len([c for c in codes_input.strip().split('\n') if c.strip()])
            if len(codes_list) < original_len:
                st.info(f"ℹ️ {original_len - len(codes_list)} doublon(s) ignoré(s) (ordre conservé)")

            if not codes_list:
                st.warning("⚠️ Aucun code ROME valide détecté.")
            else:
                st.info(f"🔄 Recherche de **{len(codes_list)}** métiers...")
                progress_bar = st.progress(0)
                metiers_data, statuts = [], []

                for i, code_rome in enumerate(codes_list):
                    try:
                        metier  = get_metier(code_rome)
                        libelle = metier.get('libelle', 'Sans libellé')
                        metiers_data.append(metier)
                        statuts.append({'code': code_rome, 'libelle': libelle, 'metier_data': metier, 'success': True})
                    except requests.HTTPError:
                        statuts.append({'code': code_rome, 'libelle': 'Non trouvé', 'success': False})
                    except Exception as e:
                        statuts.append({'code': code_rome, 'libelle': f'Erreur: {str(e)[:30]}…', 'success': False})
                    progress_bar.progress((i + 1) / len(codes_list))

                st.session_state.statuts       = statuts
                st.session_state.reussis_data  = [s['metier_data'] for s in statuts if s.get('success')]
                st.session_state.codes_list    = codes_list
                st.session_state.search_done   = True

    if st.session_state.search_done:
        statuts      = st.session_state.statuts
        reussis_data = st.session_state.reussis_data
        reussis      = sum(1 for s in statuts if s.get('success'))

        st.subheader("📊 Résumé de la recherche")
        col1, col2 = st.columns([3, 1])
        with col1:
            st.metric("Métiers trouvés", f"{reussis} / {len(st.session_state.codes_list)}")

        if reussis_data:
            df           = create_enriched_df(reussis_data)
            excel_bytes  = to_excel_bytes(df)
            st.download_button(
                label=f"📥 Télécharger Excel ({len(reussis_data)} métiers)",
                data=excel_bytes,
                file_name=f"ROME_multi_metiers_{len(reussis_data)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )
        else:
            st.info("Aucun métier trouvé → pas de fichier à télécharger.")

        st.divider()
        st.subheader("📋 Détails par métier")
        for statut in statuts:
            code_rome = statut['code']
            libelle   = statut['libelle']
            if statut.get('success'):
                metier_data    = statut['metier_data']
                cond_str       = ', '.join(get_contextes_by_categorie(metier_data, "CONDITIONS_TRAVAIL"))
                hor_str        = ', '.join(get_contextes_by_categorie(metier_data, "HORAIRE_ET_DUREE_TRAVAIL"))
                fipu_oui       = is_fipu(cond_str, hor_str)
                st.success(f"✅ **{libelle}** ({code_rome})")
                if fipu_oui:
                    st.success("**FIPU : OUI** ✅")
                else:
                    st.error("**FIPU : NON** ❌")
                st.markdown("**🏭 Conditions de travail et risques professionnels :**")
                if cond_str:
                    for item in cond_str.split(', '):
                        st.markdown(f"- {item}")
                else:
                    st.markdown("*Aucune condition trouvée*")
                st.markdown("**⏰ Horaires et durée du travail :**")
                if hor_str:
                    for item in hor_str.split(', '):
                        st.markdown(f"- {item}")
                else:
                    st.markdown("*Aucun horaire spécifique trouvé*")
            else:
                st.error(f"❌ **{code_rome}** - {libelle}")
            st.divider()

        if reussis_data:
            df          = create_enriched_df(reussis_data)
            excel_bytes = to_excel_bytes(df)
            st.download_button(
                label=f"📊 Télécharger Excel ({len(reussis_data)} métiers)",
                data=excel_bytes,
                file_name=f"ROME_multi_metiers_{len(reussis_data)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    with st.expander("💡 Exemple d'utilisation"):
        st.code("A1413\nM1805\nH1203\nK2110", language="text")


# ══════════════════════════════════════════════════════════════════════════════
# ONGLET 2 — Extraction complète (tous les ROME disponibles)
# ══════════════════════════════════════════════════════════════════════════════
with tab_all:
    st.subheader("Extraction complète de tous les métiers ROME disponibles")
    st.markdown(
        "Cet onglet récupère **tous les codes ROME** exposés par l'API, "
        "puis interroge chaque fiche métier pour en extraire les conditions de travail, "
        "les horaires et le statut FIPU. L'opération peut prendre plusieurs minutes."
    )

    # 1. Initialisation du session_state
    for key, default in [
        ("stop_full",        False),
        ("extraction_en_cours", False),
        ("df_all_fipu",      None),
        ("statuts_all",      []),
        ("metiers_data_all", []),
        ("all_codes",        []),
        ("extraction_done",  False),
    ]:
        if key not in st.session_state:
            st.session_state[key] = default

    # 2. DÉFINITION DES CALLBACKS (Ils s'exécuteront avant le rendu de l'interface)
    def on_start_extraction():
        st.session_state.stop_full = False
        st.session_state.extraction_en_cours = True
        st.session_state.extraction_done = False
        st.session_state.df_all_fipu = None
        st.session_state.statuts_all = []
        st.session_state.metiers_data_all = []
        st.session_state.all_codes = []

    def on_stop_extraction():
        st.session_state.stop_full = True

    # 3. AFFICHAGE DES BOUTONS (Liés aux callbacks via on_click)
    col1, col2 = st.columns(2)
    with col1:
        st.button("🚀 Lancer l'extraction complète", key="btn_start_full",
                  disabled=st.session_state.extraction_en_cours,
                  on_click=on_start_extraction)
    with col2:
        st.button("⛔ STOP", type="primary", key="btn_stop_full",
                  disabled=not st.session_state.extraction_en_cours,
                  on_click=on_stop_extraction)

    # 4. LOGIQUE D'EXTRACTION
    
    # Si on vient de lancer l'extraction et qu'on n'a pas encore les codes
    if st.session_state.extraction_en_cours and not st.session_state.all_codes:
        try:
            with st.spinner("Récupération de la liste de tous les codes ROME…"):
                st.session_state.all_codes = get_all_rome_codes()
        except Exception as e:
            st.error(f"Erreur lors de la récupération des codes ROME : {e}")
            st.session_state.extraction_en_cours = False

    # Boucle de traitement (s'exécute tant que l'extraction est en cours)
    if st.session_state.extraction_en_cours and st.session_state.all_codes:
        all_codes = st.session_state.all_codes
        total     = len(all_codes)
        done_so_far = len(st.session_state.statuts_all)

        st.info(f"🔄 {total} codes ROME — traitement en cours…")
        progress_bar_full  = st.progress(done_so_far / total if total > 0 else 0)
        progress_text_full = st.empty()

        for i in range(done_so_far, total):
            # Vérification du bouton STOP
            if st.session_state.stop_full:
                st.session_state.extraction_en_cours = False
                progress_text_full.text(f"⛔ Arrêté à {i} / {total} codes traités.")
                break

            code_rome = all_codes[i]
            try:
                metier  = get_metier(code_rome)
                libelle = metier.get('libelle', 'Sans libellé')
                st.session_state.metiers_data_all.append(metier)
                st.session_state.statuts_all.append(
                    {'code': code_rome, 'libelle': libelle, 'success': True}
                )
            except requests.HTTPError:
                st.session_state.statuts_all.append(
                    {'code': code_rome, 'libelle': 'Non trouvé', 'success': False}
                )
            except Exception as e:
                st.session_state.statuts_all.append(
                    {'code': code_rome, 'libelle': f'Erreur: {str(e)[:30]}…', 'success': False}
                )

            progress_bar_full.progress((i + 1) / total)
            progress_text_full.text(f"{i + 1} / {total} codes traités…")

        else:
            # Boucle terminée normalement (pas de break)
            st.session_state.extraction_en_cours = False
            st.session_state.extraction_done     = True

        # Construction du DataFrame final dès que la boucle est finie ou stoppée
        if not st.session_state.extraction_en_cours:
            metiers_data = st.session_state.metiers_data_all
            df_all_fipu  = create_enriched_df(metiers_data) if metiers_data else pd.DataFrame()
            st.session_state.df_all_fipu = df_all_fipu
            
            nb_ok  = sum(1 for s in st.session_state.statuts_all if s.get('success'))
            nb_err = len(st.session_state.statuts_all) - nb_ok
            progress_text_full.text(
                f"✅ Extraction terminée — {nb_ok} métiers récupérés, {nb_err} erreurs."
            )
            st.rerun()

    # ── Affichage des résultats persistants ───────────────────────────────────
    df_all_fipu = st.session_state.df_all_fipu
    statuts_all = st.session_state.statuts_all

    if df_all_fipu is not None:
        if df_all_fipu.empty:
            st.warning("Aucun métier récupéré.")
        else:
            nb_ok   = sum(1 for s in statuts_all if s.get('success'))
            nb_err  = len(statuts_all) - nb_ok
            nb_fipu = (df_all_fipu['FIPU'] == 'OUI').sum() if 'FIPU' in df_all_fipu.columns else 0

            st.success(f"✅ Extraction terminée — {nb_ok} métiers récupérés, {nb_err} erreurs.")

            col_a, col_b, col_c = st.columns(3)
            col_a.metric("Métiers récupérés", nb_ok)
            col_b.metric("Erreurs / non trouvés", nb_err)
            col_c.metric("Métiers FIPU (OUI)", nb_fipu)

            st.dataframe(df_all_fipu, use_container_width=True, height=500)

            excel_bytes = to_excel_bytes(df_all_fipu, sheet_name="Tous_ROME_FIPU")
            st.download_button(
                label=f"📥 Télécharger les résultats (.xlsx) — {len(df_all_fipu)} métiers",
                data=excel_bytes,
                file_name=f"ROME_tous_metiers_FIPU_{len(df_all_fipu)}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True,
            )

            if nb_err > 0:
                with st.expander(f"⚠️ {nb_err} code(s) en erreur"):
                    for s in statuts_all:
                        if not s.get('success'):
                            st.markdown(f"- **{s['code']}** — {s['libelle']}")
