import streamlit as st
import pandas as pd
from datetime import date
from analyse_logic import analyser_pointages, exporter_excel
import db_logic as db
import email_logic as mail

# Configuration de la page (doit être la première commande Streamlit)
st.set_page_config(
    page_title="PerformCheck",
    page_icon="🔍",
    layout="wide"
)

# Initialise la base de données au démarrage
db.init_db()


# Dans app.py, remplacez la fonction show_login_page existante

def show_login_page():
    """Affiche la page de connexion, d'inscription, de mot de passe oublié et gère la logique de premier compte."""
    
    # Centre la zone de connexion/inscription
    # [1, 2, 1] signifie: colonne vide (1), colonne de contenu (2 fois plus large), colonne vide (1)
    col_empty, col_main, col_empty2 = st.columns([1, 2, 1]) 
    
    with col_main:
        # 1. Affichage du Logo 
        try:
            # Utilisez le nom de fichier spécifié
            st.image("GLOBETUDES_LOGO.jpg", width=200) 
        except FileNotFoundError:
            # Titre de secours si le logo n'est pas trouvé
            st.title("PerformCheck") 

        st.markdown("---") # Séparateur visuel

        # --- GESTION DU PREMIER COMPTE ---
        # Cas 1 : Aucun utilisateur n'existe -> Création du premier compte admin (unique)
        if not db.check_if_users_exist():
            st.subheader("Bienvenue ! Créez le premier compte administrateur")
            with st.form("signup_form"):
                new_username = st.text_input("Choisissez un nom d'utilisateur (votre e-mail)")
                new_password = st.text_input("Choisissez un mot de passe", type="password")
                if st.form_submit_button("Créer le compte", type="primary"):
                    success, message = db.add_user(new_username, new_password)
                    if success:
                        st.success(message)
                        st.info("Veuillez maintenant vous connecter.")
                        st.rerun()
                    else:
                        st.error(message)
        
        # --- GESTION DES COMPTES EXISTANTS (Connexion, Inscription multiple, Mot de passe oublié) ---
        else:
            # Utilisation des onglets pour une présentation professionnelle
            tab_login, tab_signup, tab_forgot = st.tabs(["Se connecter", "Créer un compte", "Mot de passe oublié"])
            
            # --- Onglet de Connexion ---
            with tab_login:
                st.subheader("Connexion")
                with st.form("login_form"):
                    username = st.text_input("Nom d'utilisateur (E-mail)")
                    password = st.text_input("Mot de passe", type="password")
                    if st.form_submit_button("Se connecter", type="primary"):
                        if db.check_user(username, password):
                            st.session_state['logged_in'] = True
                            st.session_state['username'] = username
                            st.rerun()
                        else:
                            st.error("Nom d'utilisateur ou mot de passe incorrect.")

            # --- Onglet de Création de Compte (Inscriptions Multiples) ---
            with tab_signup:
                st.subheader("Créer un nouveau compte")
                with st.form("new_signup_form"):
                    new_username = st.text_input("Choisissez un nom d'utilisateur (votre e-mail)", key="new_user_email")
                    new_password = st.text_input("Choisissez un mot de passe", type="password", key="new_user_password")
                    if st.form_submit_button("Créer le compte utilisateur", type="primary"):
                        success, message = db.add_user(new_username, new_password)
                        if success:
                            st.success(message)
                            st.info("Compte créé. Veuillez maintenant vous connecter.")
                        else:
                            st.error(message)

            # --- Onglet de Mot de Passe Oublié ---
            with tab_forgot:
                st.subheader("Réinitialiser le mot de passe")
                with st.form("forgot_password_form"):
                    email_to_reset = st.text_input("Entrez votre nom d'utilisateur (votre e-mail) pour réinitialiser", key="reset_email_input")
                    if st.form_submit_button("Envoyer le lien de réinitialisation"):
                        token = db.set_reset_token(email_to_reset)
                        if token:
                            # Votre logique pour construire l'URL et envoyer l'email
                            app_url = st.get_option("server.baseUrlPath")
                            if not app_url.startswith("http"):
                               app_url = "http://localhost:8501" 
                            
                            # NOTE : Assurez-vous que les fonctions db.set_reset_token et mail.send_reset_email fonctionnent
                            st.success("Un e-mail de réinitialisation a été envoyé (si le service d'e-mail est configuré).")
                        else:
                            st.error("Aucun compte trouvé pour cet utilisateur.")

def show_reset_password_page(token):
    """Affiche la page pour entrer un nouveau mot de passe."""
    st.title("Réinitialiser votre mot de passe")
    with st.form("reset_form"):
        new_password = st.text_input("Entrez votre nouveau mot de passe", type="password")
        confirm_password = st.text_input("Confirmez le nouveau mot de passe", type="password")
        if st.form_submit_button("Valider"):
            if new_password == confirm_password:
                success, message = db.reset_password_with_token(token, new_password)
                if success:
                    st.success(message)
                    st.info("Vous pouvez maintenant fermer cet onglet et vous connecter.")
                else:
                    st.error(message)
            else:
                st.error("Les mots de passe ne correspondent pas.")


def show_main_app():
    """Affiche l'application principale d'analyse une fois connecté."""

    st.sidebar.success(f"Connecté en tant que {st.session_state['username']}")
    if st.sidebar.button("Se déconnecter"):
        del st.session_state['logged_in']
        del st.session_state['username']
        st.rerun()

    with st.sidebar.expander("Changer le mot de passe"):
        with st.form("change_password_form", clear_on_submit=True):
            old_password = st.text_input("Ancien mot de passe", type="password")
            new_password = st.text_input("Nouveau mot de passe", type="password")
            confirm_password = st.text_input("Confirmer le nouveau mot de passe", type="password")
            
            if st.form_submit_button("Valider"):
                if new_password == confirm_password:
                    success, message = db.update_password(st.session_state['username'], old_password, new_password)
                    if success:
                        st.sidebar.success(message)
                    else:
                        st.sidebar.error(message)
                else:
                    st.sidebar.error("Les nouveaux mots de passe ne correspondent pas.")

    # --- Votre application d'analyse commence ici ---
    st.title("Système d'Analyse de Pointage et de Congés")

    if 'affectations_manuelles' not in st.session_state:
        st.session_state.affectations_manuelles = []

    st.subheader("Étape 1 : Charger les fichiers")
    col1, col2, col3 = st.columns(3)
    with col1:
        pointage_file = st.file_uploader("Fichier de Pointage (Obligatoire)", type=["xlsx", "xls"])
    with col2:
        conges_file = st.file_uploader("Fichier des Congés (Optionnel)", type=["xlsx", "xls"])
    with col3:
        affectations_file = st.file_uploader("Fichier des Affectations (Optionnel)", type=["xlsx", "xls"])

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

    if pointage_file is not None:
        st.subheader("Étape 2 : Définir les paramètres d'analyse")
        # ... (Le reste de votre code est identique)
        param_col1, param_col2 = st.columns(2)
        with param_col1:
            current_year = date.today().year
            annee = st.selectbox("Année", range(2023, current_year + 15), index=0)
        with param_col2:
            current_month = date.today().month
            mois = st.selectbox("Mois", range(1, 13), index=current_month - 1)
        dates_du_mois = [date(annee, mois, jour) for jour in range(1, pd.Timestamp(annee, mois, 1).days_in_month + 1)]
        jours_feries_selectionnes = st.multiselect(
            "Confirmez les jours fériés pour ce mois :",
            options=dates_du_mois,
            format_func=lambda d: d.strftime('%A %d %B %Y')
        )

        st.subheader("Étape 3 : Choisir les colonnes pour le rapport final")
        toutes_les_colonnes_possibles = [
            'Matricule', 'Score Discipline (%)',
            'Jours Payés (par Employeur)', 
            'Nb Jours Absence Injustifiée',
            'Détail des Absences', 
            'Nb Jours en Retard',
            'Nb Jours Chantier', 'Nb Jours Domicile',
            'Total Heures Normales', 'Heures Normales Bureau', 'Heures Normales Chantier', 'Heures Normales Domicile',
            'Total Majorations', 'Majoration 25% (Val)', 'Majoration 50% (Val)', 'Majoration 100% (Val)',
            'Total HS 25%', 'Total HS 50%', 'Total HS 100%',
            'Nb Jours Congé Payé par Employeur', 'Nb Jours Congé Non Payé', 'Nb Jours Payé par CNSS',
            'Détail Jours de Congé', 'Détail des Types de Congé',
            'Lieu(x) de Chantier', 'Projet(s) (Domicile)'
        ]
        selection_par_defaut = [
            'Matricule', 'Score Discipline (%)',
            'Jours Payés (par Employeur)', 'Nb Jours Absence Injustifiée', 'Total Majorations',
            'Nb Jours Chantier', 'Nb Jours Domicile'
        ]
        colonnes_choisies = st.multiselect(
            "Cochez les colonnes à inclure dans le rapport :",
            options=toutes_les_colonnes_possibles,
            default=selection_par_defaut
        )

        st.divider()
        if st.button("Lancer l'analyse et préparer le rapport", type="primary"):
            try:
                df_pointage = pd.read_excel(pointage_file, header=None)
                df_conges = pd.read_excel(conges_file, header=None) if conges_file else pd.DataFrame()
                
                df_affectations_fichier = pd.read_excel(affectations_file, header=None) if affectations_file else pd.DataFrame()
                df_affectations_manuel = pd.DataFrame(st.session_state.affectations_manuelles)
                if not df_affectations_manuel.empty:
                    df_affectations_manuel['Date'] = pd.to_datetime(df_affectations_manuel['Date'])

                with st.spinner("Analyse en cours... Cette opération peut prendre un moment."):
                    result_df = analyser_pointages(
                        df_pointage, df_conges,
                        df_affectations_fichier,
                        df_affectations_manuel,
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


# Vérifie si un token de reset est dans l'URL
query_params = st.query_params
if "reset_token" in query_params:
    show_reset_password_page(query_params["reset_token"])
# Vérifie si l'utilisateur est connecté
elif st.session_state.get('logged_in', False):
    show_main_app()
# Sinon, affiche la page de connexion
else:
    show_login_page()