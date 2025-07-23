import streamlit as st
import pandas as pd
from datetime import date
# La signature de la fonction d'analyse a changé
from analyse_logic import analyser_pointages, exporter_excel

# --- Configuration de la Page ---
st.set_page_config(
    page_title="PerformCheck",
    page_icon="🔍",
    layout="wide"
)

st.title("Système d'Analyse de Pointage et de Congés")

# --- Initialisation du Session State pour la saisie manuelle ---
if 'affectations_manuelles' not in st.session_state:
    st.session_state.affectations_manuelles = []

# --- ÉTAPE 1 : CHARGEMENT DES FICHIERS ---
st.subheader("Étape 1 : Charger les fichiers")
col1, col2, col3 = st.columns(3)
with col1:
    pointage_file = st.file_uploader("Fichier de Pointage (Obligatoire)", type=["xlsx", "xls"])
with col2:
    conges_file = st.file_uploader("Fichier des Congés (Optionnel)", type=["xlsx", "xls"])
with col3:
    affectations_file = st.file_uploader("Fichier des Affectations (Optionnel)", type=["xlsx", "xls"])

# --- SAISIE MANUELLE ---
with st.expander("✍️ Ajouter/Corriger une affectation manuellement (pour les cas urgents)"):
    with st.form("formulaire_affectation", clear_on_submit=True):
        st.write("Les entrées manuelles écrasent les données du fichier Excel pour le même jour.")
        col_form1, col_form2, col_form3 = st.columns(3)
        with col_form1:
            manuel_matricule = st.text_input("Matricule de l'employé")
        with col_form2:
            manuel_date = st.date_input("Date de l'affectation", key="manuel_date")
        with col_form3:
            manuel_affectation = st.selectbox("Type d'affectation", ["Chantier", "Domicile", "Chantier et Bureau"])
        
        manuel_lieu = st.text_input("Lieu du Chantier (si applicable)")
        manuel_projet = st.text_input("Projet (si travail à domicile)")

        submitted = st.form_submit_button("Ajouter l'affectation")
        if submitted and manuel_matricule and manuel_date and manuel_affectation:
            st.session_state.affectations_manuelles.append({
                'Matricule': manuel_matricule, 'Date': manuel_date,
                'Affectation': manuel_affectation, 'Lieu_Chantier': manuel_lieu,
                'Projet_Domicile': manuel_projet
            })
            st.success(f"Affectation pour {manuel_matricule} le {manuel_date.strftime('%d/%m/%Y')} ajoutée !")

# --- AFFICHAGE DES AFFECTATIONS MANUELLES ---
st.divider()
st.write("#### Affectations manuelles en cours :")
if not st.session_state.affectations_manuelles:
    st.info("Aucune affectation manuelle n'a été ajoutée pour le moment.")
else:
    col_spec = [1.5, 2, 2, 4, 2]
    cols = st.columns(col_spec)
    cols[0].markdown("**Matricule**"); cols[1].markdown("**Date**"); cols[2].markdown("**Type**"); cols[3].markdown("**Lieu / Projet**")
    for i, affectation in enumerate(st.session_state.affectations_manuelles):
        cols = st.columns(col_spec)
        cols[0].write(affectation['Matricule'])
        cols[1].write(affectation['Date'].strftime('%d/%m/%Y'))
        cols[2].write(affectation['Affectation'])
        cols[3].write(affectation.get('Lieu_Chantier', '') or affectation.get('Projet_Domicile', ''))
        if cols[4].button("Supprimer", key=f"delete_{i}", use_container_width=True):
            del st.session_state.affectations_manuelles[i]
            st.rerun()

# --- La suite de l'application ne s'affiche que si le fichier de pointage est chargé ---
if pointage_file is not None:
    st.subheader("Étape 2 : Définir les paramètres d'analyse")
    param_col1, param_col2 = st.columns(2)
    with param_col1:
        current_year = date.today().year
        annee = st.selectbox("Année", range(current_year, current_year + 10), index=0)
    with param_col2:
        current_month = date.today().month
        mois = st.selectbox("Mois", range(1, 13), index=current_month - 1)
    dates_du_mois = [date(annee, mois, jour) for jour in range(1, pd.Timestamp(annee, mois, 1).days_in_month + 1)]
    jours_feries_selectionnes = st.multiselect(
        "Confirmez les jours fériés pour ce mois :",
        options=dates_du_mois,
        format_func=lambda d: d.strftime('%A %d %B %Y')
    )

    # --- ÉTAPE 3 : CHOIX DES COLONNES ---
    st.subheader("Étape 3 : Choisir les colonnes pour le rapport final")
    toutes_les_colonnes_possibles = [
        'Matricule', 'Jours Payés (par Employeur)', 'Nb Jours Absence Injustifiée',
        'Nb Jours Chantier', 'Nb Jours Domicile',
        'Total Heures Normales', 'Heures Normales Bureau', 'Heures Normales Chantier', 'Heures Normales Domicile',
        'Total HS 25%', 'Total HS 50%', 'Total HS 100%',
        'Total Majorations', 'Majoration 25% (Val)', 'Majoration 50% (Val)', 'Majoration 100% (Val)',
        'Nb Jours Congé Payé par Employeur', 'Nb Jours Congé Non Payé', 'Nb Jours Payé par CNSS',
        'Détail des Absences', 'Détail Jours de Congé', 'Détail des Types de Congé',
        'Lieu(x) de Chantier', 'Projet(s) (Domicile)'
    ]
    selection_par_defaut = [
        'Matricule', 'Jours Payés (par Employeur)', 'Nb Jours Absence Injustifiée', 'Total Heures Normales', 'Total Majorations',
        'Nb Jours Chantier', 'Nb Jours Domicile'
    ]
    colonnes_choisies = st.multiselect(
        "Cochez les colonnes à inclure dans le rapport :",
        options=toutes_les_colonnes_possibles,
        default=selection_par_defaut
    )

    # --- BOUTON FINAL ---
    st.divider()
    if st.button("Lancer l'analyse et préparer le rapport", type="primary"):
        try:
            df_pointage = pd.read_excel(pointage_file, header=None)
            df_conges = pd.read_excel(conges_file, header=None) if conges_file else pd.DataFrame()
            
            # --- CORRECTION : Préparation des données d'affectation ---
            df_affectations_fichier = pd.read_excel(affectations_file, header=None) if affectations_file else pd.DataFrame()
            df_affectations_manuel = pd.DataFrame(st.session_state.affectations_manuelles)
            if not df_affectations_manuel.empty:
                df_affectations_manuel['Date'] = pd.to_datetime(df_affectations_manuel['Date'])
            # --- FIN CORRECTION ---

            with st.spinner("Analyse en cours... Cette opération peut prendre un moment."):
                result_df = analyser_pointages(
                    df_pointage, df_conges,
                    df_affectations_fichier, # On passe le DF brut du fichier
                    df_affectations_manuel,  # On passe le DF propre de la saisie manuelle
                    mois, annee, jours_feries_selectionnes
                )

            if result_df.empty:
                st.warning("Aucun résultat généré. Vérifiez les fichiers d'entrée.")
            else:
                st.success("Analyse terminée !")
                st.subheader("Aperçu du rapport final")
                df_pour_apercu = exporter_excel(result_df.copy(), "apercu", toutes_les_colonnes_possibles, return_df=True)
                colonnes_a_afficher = [col for col in colonnes_choisies if col in df_pour_apercu.columns]
                st.dataframe(df_pour_apercu[colonnes_a_afficher])
                
                output_file_name = f"rapport_final_{mois}_{annee}.xlsx"
                output_data = exporter_excel(result_df, output_file_name, colonnes_choisies)
                
                st.download_button(
                    label="⬇️ Télécharger le rapport complet",
                    data=output_data,
                    file_name=output_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Une erreur est survenue lors de l'analyse : {e}")
            st.error("Veuillez vérifier le format de vos fichiers et les données saisies.")
