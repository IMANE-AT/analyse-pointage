import streamlit as st
import pandas as pd
from datetime import date
from analyse_logic import analyser_pointages, exporter_excel

st.set_page_config(layout="wide")
st.title("Système d'Analyse de Pointage et de Congés")

# --- ÉTAPE 1 : CHARGEMENT DES FICHIERS ---
st.subheader("Étape 1 : Charger les fichiers")
col1, col2 = st.columns(2)
with col1:
    pointage_file = st.file_uploader("Fichier de Pointage (Obligatoire)", type=["xlsx", "xls"])
with col2:
    conges_file = st.file_uploader("Fichier des Congés (Optionnel)", type=["xlsx", "xls"])

if pointage_file is not None:
    # --- ÉTAPE 2 : PARAMÈTRES DE L'ANALYSE ---
    st.subheader("Étape 2 : Définir les paramètres")
    param_col1, param_col2 = st.columns(2)
    with param_col1:
        # Correction de l'index pour éviter les erreurs si l'année change
        current_year = date.today().year
        year_range = range(current_year, current_year + 10)
        annee = st.selectbox("Année", year_range, index=0)
    with param_col2:
        # Correction de l'index pour éviter les erreurs au changement de mois
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
    
    # AJOUT : Nouvelles colonnes dans la liste de choix
    toutes_les_colonnes_possibles = [
        'Matricule', 
        'Jours Payés (par Employeur)', 
        'Nb Jours Absence Injustifiée', 
        'Nb Jours Congé Payé', 
        'Nb Jours Congé Non Payé',
        'Nb Jours Payé par CNSS',
        'Détail des Types de Congé',
        'Détail Jours de Congé',
        'Payé Par', 
        'Heures Normales', 
        'Heures Sup 25%', 'Heures Sup 50%', 'Heures Sup 100%', 
        'Majoration 25% (Val)', 'Majoration 50% (Val)', 'Majoration 100% (Val)', 
        'Total Majorations', 
        'Heures Pause Déjeuner', 
        'Détail des Absences'
    ]
    
    selection_par_defaut = ['Matricule', 'Jours Payés (par Employeur)', 'Nb Jours Absence Injustifiée', 'Total Majorations']
    
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
            df_conges = pd.DataFrame()
            if conges_file is not None:
                df_conges = pd.read_excel(conges_file, header=None)

            with st.spinner("Analyse en cours..."):
                result_df = analyser_pointages(df_pointage, df_conges, mois, annee, jours_feries_selectionnes)

            if result_df.empty:
                st.warning("Aucun résultat généré.")
            else:
                st.success("Analyse terminée !")
                
                # MODIFICATION : Dictionnaire pour renommer les colonnes pour l'affichage et l'export
                rename_dict = {
                    'Nb Jours Absence Injustifiée': 'Nb Jours Absence Injustifiée',
                    'Nb Jours Congé Payé': 'Nb Jours Congé Payé',
                    'Nb Jours Congé Non Payé': 'Nb Jours Congé Non Payé',
                    'Nb Jours Payé par CNSS': 'Nb Jours Payé par CNSS',
                    'Détail des Congés': 'Détail des Types de Congé',
                    'Détail des Jours de Congé': 'Détail Jours de Congé',
                    'Jours_Payés': 'Jours Payés (par Employeur)',
                    'Maj_25': 'Majoration 25% (Val)',
                    'Maj_50': 'Majoration 50% (Val)',
                    'Maj_100': 'Majoration 100% (Val)',
                    'Total_Majorations': 'Total Majorations',
                    'Heures_Normales': 'Heures Normales',
                    'Heures_Pause_Dej': 'Heures Pause Déjeuner',
                    'Détail des Absences': 'Détail des Absences',
                    'Payé Par': 'Payé Par',
                    'Heures_Sup_Maj25': 'Heures Sup 25%',
                    'Heures_Sup_Maj50': 'Heures Sup 50%',
                    'Heures_Sup_Maj100': 'Heures Sup 100%'
                }
                result_df_renamed = result_df.rename(columns=rename_dict)
                apercu_cols = [col for col in colonnes_choisies if col in result_df_renamed.columns]
                
                st.subheader("Aperçu du rapport final")
                st.dataframe(result_df_renamed[apercu_cols])

                output_file_name = f"rapport_final_{mois}_{annee}.xlsx"
                # L'export utilise directement le result_df et le dictionnaire de renommage dans la fonction
                output_data = exporter_excel(result_df, output_file_name, colonnes_choisies)
                
                st.download_button(
                    label="Télécharger le rapport",
                    data=output_data,
                    file_name=output_file_name,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
        except Exception as e:
            st.error(f"Une erreur est survenue lors de l'analyse : {e}")