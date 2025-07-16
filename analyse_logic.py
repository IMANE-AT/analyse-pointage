import pandas as pd
from datetime import time, timedelta, date
from itertools import product
from io import BytesIO

# --- Dictionnaire de Configuration et Règles ---
CONFIG = {
    'BASE_JOURS_PAYES': 26, 'DUREE_JOURNEE_NORMALE_HEURES': 8.0, 'JOUR_DIMANCHE': 'Sunday',
    'HEURE_DEBUT_JOURNEE_NORMALE': time(8, 30), 'HEURE_DEBUT_NUIT': time(21, 0),
    'HEURE_FIN_NUIT': time(6, 0), 'JOUR_VENDREDI': 'Friday', 'JOUR_SAMEDI': 'Saturday',
    'HEURE_FIN_TRAVAIL_SAMEDI_MATIN': time(12, 30), 'HEURE_DEBUT_PAUSE_DEJ': time(12, 30),
    'HEURE_FIN_PAUSE_DEJ_LUN_JEU': time(14, 30), 'HEURE_FIN_PAUSE_DEJ_VENDREDI': time(15, 0),
    'TOLERANCE_RETARD_MIN': 20, 'HEURE_FIN_JOURNEE_NORMALE_STANDARD_THEORIQUE': time(18, 30),
    'HEURE_FIN_JOURNEE_NORMALE_VENDREDI_THEORIQUE': time(19, 0),
    'JOURS_OUVRABLES_LUN_JEU': ['Monday', 'Tuesday', 'Wednesday', 'Thursday']
}
REGLES_CONGES = {
    "CONGE_PAYE": {"mots_cles": ["annuel", "payé"], "duree_max": 18, "paye_par": "Employeur", "est_paye": True},
    "CONGE_MALADIE": {"mots_cles": ["maladie", "medical"], "duree_max": 180, "paye_par": "CNSS", "est_paye": True},
    # AJOUT : Règle spécifique pour les congés maladie courts
    "CONGE_MALADIE_COURT": {"mots_cles": [], "duree_max": 3, "paye_par": "Personne", "est_paye": False},
    "CONGE_MARIAGE": {"mots_cles": ["mariage"], "duree_max": 4, "paye_par": "Employeur", "est_paye": True},
    "CONGE_PATERNITE": {"mots_cles": ["paternité", "naissance"], "duree_max": 3, "paye_par": "Employeur", "est_paye": True},
    "CONGE_DECES": {"mots_cles": ["décès", "deces"], "duree_max": 3, "paye_par": "Employeur", "est_paye": True},
    "CONGE_SANS_SOLDE": {"mots_cles": ["sans solde", "non payé"], "duree_max": 999, "paye_par": "Personne", "est_paye": False},
    "AUTRE": {"mots_cles": [], "duree_max": 999, "paye_par": "Inconnu", "est_paye": False}
}


# --- Fonctions Utilitaires et de Calcul (inchangées) ---
def find_and_rename_header(df, columns_map):
    df_copy = df.copy()
    for i in range(min(5, len(df_copy))):
        try:
            row_values = [str(v).lower() for v in df_copy.iloc[i].values]
            found_keywords = 0
            for cell_value in row_values:
                for keywords in columns_map.values():
                    if any(keyword in cell_value for keyword in keywords):
                        found_keywords += 1
                        break
            if found_keywords >= 2:
                df_copy.columns = df_copy.iloc[i]
                df_copy = df_copy.iloc[i+1:].reset_index(drop=True)
                new_names = {}
                used_standard_names = set()
                for col_name in df_copy.columns:
                    if pd.isna(col_name): continue
                    col_name_str = str(col_name).lower()
                    for standard_name, keywords in columns_map.items():
                        if standard_name not in used_standard_names and any(keyword in col_name_str for keyword in keywords):
                            new_names[col_name] = standard_name
                            used_standard_names.add(standard_name)
                            break
                df_copy = df_copy.rename(columns=new_names)
                return df_copy
        except:
            continue
    num_cols_to_assign = min(len(df_copy.columns), len(columns_map.keys()))
    df_copy.columns = list(columns_map.keys())[:num_cols_to_assign]
    return df_copy

def prepare_conges_df(df_conges_raw):
    if df_conges_raw.empty:
        return pd.DataFrame(columns=['Matricule', 'Type_Congé', 'Date_Debut', 'Date_Fin', 'Type_Congé_Standard'])
    conges_cols_map = {'Matricule': ['matr', 'id'], 'Type_Congé': ['type', 'motif'], 'Date_Debut': ['debut', 'début', 'start'], 'Date_Fin': ['fin', 'end']}
    df_conges = find_and_rename_header(df_conges_raw.copy(), conges_cols_map)
    df_conges.drop_duplicates(inplace=True)
    if 'Date_Fin' not in df_conges.columns:
        df_conges['Date_Fin'] = pd.NaT
    df_conges['Date_Debut'] = pd.to_datetime(df_conges['Date_Debut'], errors='coerce')
    df_conges['Date_Fin'] = pd.to_datetime(df_conges['Date_Fin'], errors='coerce').fillna(df_conges['Date_Debut'])
    df_conges.dropna(subset=['Matricule', 'Date_Debut', 'Date_Fin'], inplace=True)
    def find_type(text):
        text_lower = str(text).lower()
        for rule_name, rule_details in REGLES_CONGES.items():
            if rule_name == "CONGE_MALADIE_COURT": continue # On ne détecte pas ce type via mot-clé
            if any(keyword in text_lower for keyword in rule_details['mots_cles']):
                return rule_name
        return "AUTRE"
    df_conges['Type_Congé_Standard'] = df_conges['Type_Congé'].apply(find_type)
    return df_conges

def get_full_date_range(mois, annee):
    start_date = f'{annee}-{mois:02d}-01'
    end_date = pd.to_datetime(start_date) + pd.offsets.MonthEnd(1)
    return pd.date_range(start=start_date, end=end_date, freq='D').date

def calculer_heures_supplementaires(debut, fin, jour_semaine, est_jour_ferie, heures_normales_deja_comptees):
    heures_normales, h_sup25, h_sup50, h_sup100 = 0.0, 0.0, 0.0, 0.0
    if est_jour_ferie or jour_semaine == CONFIG['JOUR_DIMANCHE']:
        duree_totale = (fin - debut).total_seconds() / 3600
        nuit_debut_soir = pd.to_datetime(f"{debut.date()} {CONFIG['HEURE_DEBUT_NUIT']}")
        nuit_fin_matin = pd.to_datetime(f"{debut.date()} {CONFIG['HEURE_FIN_NUIT']}")
        heures_nuit = 0
        if max(debut, nuit_debut_soir) < fin: heures_nuit += (fin - max(debut, nuit_debut_soir)).total_seconds() / 3600
        if min(fin, nuit_fin_matin) > debut: heures_nuit += (min(fin, nuit_fin_matin) - debut).total_seconds() / 3600
        heures_jour = duree_totale - heures_nuit
        h_sup100 += heures_nuit
        h_sup50 += heures_jour
    elif jour_semaine == CONFIG['JOUR_SAMEDI']:
        limite_samedi = pd.to_datetime(f"{debut.date()} {CONFIG['HEURE_FIN_TRAVAIL_SAMEDI_MATIN']}")
        if fin <= limite_samedi: heures_normales = (fin - debut).total_seconds() / 3600
        elif debut < limite_samedi:
            heures_normales = (limite_samedi - debut).total_seconds() / 3600
            h_sup50 = (fin - limite_samedi).total_seconds() / 3600
        else: h_sup50 = (fin - debut).total_seconds() / 3600
    else:
        pause_debut = pd.to_datetime(f"{debut.date()} {CONFIG['HEURE_DEBUT_PAUSE_DEJ']}")
        pause_fin_heure = CONFIG['HEURE_FIN_PAUSE_DEJ_VENDREDI'] if jour_semaine == CONFIG['JOUR_VENDREDI'] else CONFIG['HEURE_FIN_PAUSE_DEJ_LUN_JEU']
        pause_fin = pd.to_datetime(f"{debut.date()} {pause_fin_heure}")
        duree_avant_pause = max(0, (min(fin, pause_debut) - debut).total_seconds() / 3600)
        duree_apres_pause = max(0, (fin - max(debut, pause_fin)).total_seconds() / 3600)
        heures_travaillees_reelles = duree_avant_pause + duree_apres_pause
        heures_normales_a_ajouter = min(heures_travaillees_reelles, CONFIG['DUREE_JOURNEE_NORMALE_HEURES'] - heures_normales_deja_comptees)
        heures_normales += heures_normales_a_ajouter
        heures_sup_jour = heures_travaillees_reelles - heures_normales_a_ajouter
        nuit_debut_soir = pd.to_datetime(f"{debut.date()} {CONFIG['HEURE_DEBUT_NUIT']}")
        nuit_fin_matin = pd.to_datetime(f"{debut.date()} {CONFIG['HEURE_FIN_NUIT']}")
        heures_nuit = 0
        if max(debut, nuit_debut_soir) < fin: heures_nuit += (fin - max(debut, nuit_debut_soir)).total_seconds() / 3600
        if min(fin, nuit_fin_matin) > debut: heures_nuit += (min(fin, nuit_fin_matin) - debut).total_seconds() / 3600
        heures_nuit_valide = min(heures_sup_jour, heures_nuit)
        h_sup50 += heures_nuit_valide
        h_sup25 += max(0, heures_sup_jour - heures_nuit_valide)
    return heures_normales, h_sup25, h_sup50, h_sup100

def calculer_indicateurs_jour_travaille(jour_pointages):
    heures_normales, h_sup25, h_sup50, h_sup100, heures_pause_dej = 0.0, 0.0, 0.0, 0.0, 0.0
    date_jour = jour_pointages.iloc[0]['Date']
    jour_semaine = pd.to_datetime(date_jour).day_name()
    est_jour_ferie = jour_pointages.iloc[0]['Est_JourFerie']
    pointages_du_jour = jour_pointages.sort_values('Pointage')
    heures_normales_deja_comptees = 0.0
    num_pointages = len(pointages_du_jour)
    if num_pointages % 2 != 0: num_pointages -= 1
    for i in range(0, num_pointages, 2):
        debut = pointages_du_jour.iloc[i]['Pointage']
        fin = pointages_du_jour.iloc[i+1]['Pointage']
        if pd.isna(fin) or pd.isna(debut): continue
        pause_debut = pd.to_datetime(f"{date_jour} {CONFIG['HEURE_DEBUT_PAUSE_DEJ']}")
        pause_fin_heure = CONFIG['HEURE_FIN_PAUSE_DEJ_VENDREDI'] if jour_semaine == CONFIG['JOUR_VENDREDI'] else CONFIG['HEURE_FIN_PAUSE_DEJ_LUN_JEU']
        pause_fin = pd.to_datetime(f"{date_jour} {pause_fin_heure}")
        overlap_debut_pause = max(debut, pause_debut)
        overlap_fin_pause = min(fin, pause_fin)
        if overlap_fin_pause > overlap_debut_pause:
            heures_pause_dej += (overlap_fin_pause - overlap_debut_pause).total_seconds() / 3600
        hn, h25, h50, h100 = calculer_heures_supplementaires(debut, fin, jour_semaine, est_jour_ferie, heures_normales_deja_comptees)
        heures_normales += hn
        heures_normales_deja_comptees += hn
        h_sup25 += h25
        h_sup50 += h50
        h_sup100 += h100
    return pd.Series({
        'Heures_Normales': round(heures_normales, 2), 'Heures_Sup_Maj25': round(h_sup25, 2),
        'Heures_Sup_Maj50': round(h_sup50, 2), 'Heures_Sup_Maj100': round(h_sup100, 2),
        'Heures_Pause_Dej': round(heures_pause_dej, 2)
    })

# --- FONCTION PRINCIPALE D'ANALYSE ---
def analyser_pointages(df_pointage_raw, df_conges_raw, mois, annee, jours_feries):
    pointage_cols_map = {'Matricule': ['matr', 'id'], 'Pointage': ['pointage', 'date']}
    df_pointage = find_and_rename_header(df_pointage_raw.copy(), pointage_cols_map)
    df_pointage['Pointage'] = pd.to_datetime(df_pointage['Pointage'], errors='coerce')
    df_pointage.dropna(subset=['Pointage'], inplace=True)
    df_pointage.drop_duplicates(subset=['Matricule', 'Pointage'], inplace=True)
    df_pointage.sort_values(['Matricule', 'Pointage'], inplace=True)
    diff_temps = df_pointage.groupby('Matricule')['Pointage'].diff()
    df_pointage = df_pointage[(diff_temps > timedelta(minutes=10)) | (diff_temps.isnull())]
    df_pointage['Date'] = df_pointage['Pointage'].dt.date
    
    df_conges = prepare_conges_df(df_conges_raw)

    all_matricules = df_pointage['Matricule'].astype(str).unique()
    if len(all_matricules) == 0: return pd.DataFrame()
    
    full_date_range = get_full_date_range(mois, annee)
    calendrier_ref = pd.DataFrame(list(product(all_matricules, full_date_range)), columns=['Matricule', 'Date'])
    calendrier_ref['Jour_Semaine'] = pd.to_datetime(calendrier_ref['Date']).dt.day_name()
    calendrier_ref['Est_JourFerie'] = calendrier_ref['Date'].isin(jours_feries)
    calendrier_ref['Est_Jour_Ouvrable'] = ((calendrier_ref['Jour_Semaine'] != CONFIG['JOUR_DIMANCHE']) & (calendrier_ref['Est_JourFerie'] == False))

    df_pointage['Present'] = 1
    calendrier_ref['Matricule'] = calendrier_ref['Matricule'].astype(str)
    df_pointage['Matricule'] = df_pointage['Matricule'].astype(str)
    df_analyse = pd.merge(calendrier_ref, df_pointage, on=['Matricule', 'Date'], how='left')
    
    df_analyse['Type_Congé'] = ""
    if not df_conges.empty:
        df_conges['Matricule'] = df_conges['Matricule'].astype(str)
        for index, conge in df_conges.iterrows():
            regle = REGLES_CONGES.get(conge['Type_Congé_Standard'], REGLES_CONGES["AUTRE"])
            duree_max_autorisee = regle['duree_max']
            tous_les_jours_du_conge = pd.date_range(start=conge['Date_Debut'], end=conge['Date_Fin'])
            jours_ouvrables_conge = [j for j in tous_les_jours_du_conge if j.weekday() < 6 and pd.to_datetime(j.date()) not in jours_feries]
            
            # AJOUT : Logique pour congé maladie court
            type_conge_a_appliquer = conge['Type_Congé_Standard']
            if type_conge_a_appliquer == "CONGE_MALADIE" and len(jours_ouvrables_conge) < 4:
                type_conge_a_appliquer = "CONGE_MALADIE_COURT"
            
            jours_conge_justifies = jours_ouvrables_conge[:duree_max_autorisee]
            mask = (df_analyse['Matricule'] == conge['Matricule']) & (df_analyse['Date'].isin([j.date() for j in jours_conge_justifies]))
            df_analyse.loc[mask, 'Type_Congé'] = type_conge_a_appliquer

    df_analyse['Absence_Injustifiee'] = ((df_analyse['Est_Jour_Ouvrable'] == True) & (df_analyse['Present'].isnull()) & (df_analyse['Type_Congé'] == ""))
    
    jours_travailles = df_analyse[df_analyse['Present'].notnull()].groupby(['Matricule', 'Date'])
    if not jours_travailles.groups: indicateurs_calcules = pd.DataFrame()
    else: indicateurs_calcules = jours_travailles.apply(calculer_indicateurs_jour_travaille).reset_index()
    
    df_analyse = pd.merge(df_analyse.drop(columns=['Pointage', 'Present']), indicateurs_calcules, on=['Matricule', 'Date'], how='left')
    cols_to_fill = ['Heures_Normales', 'Heures_Sup_Maj25', 'Heures_Sup_Maj50', 'Heures_Sup_Maj100', 'Heures_Pause_Dej']
    df_analyse[cols_to_fill] = df_analyse[cols_to_fill].fillna(0)
    
    resume_mensuel = df_analyse.groupby('Matricule').agg(
        Heures_Normales=('Heures_Normales', 'sum'),
        Heures_Sup_Maj25=('Heures_Sup_Maj25', 'sum'), Heures_Sup_Maj50=('Heures_Sup_Maj50', 'sum'),
        Heures_Sup_Maj100=('Heures_Sup_Maj100', 'sum'), Heures_Pause_Dej=('Heures_Pause_Dej', 'sum'),
        Nb_Jours_Absence_Injustifiee=('Absence_Injustifiee', 'sum')
    ).reset_index()

    # --- ### MODIFICATION : Agrégation des congés avec décompte CNSS ### ---
    def aggregate_conges(series):
        types = [REGLES_CONGES.get(t, {}) for t in series if t and t != ""]
        paye = sum(1 for t in types if t.get('est_paye'))
        non_paye = sum(1 for t in types if not t.get('est_paye'))
        paye_cnss = sum(1 for t in types if t.get('paye_par') == 'CNSS') # Nouveau compteur
        details = ", ".join(sorted(list(set(s for s in series if s and s != ""))))
        payers = ", ".join(sorted(list(set(t.get('paye_par', 'Inconnu') for t in types))))
        return paye, non_paye, paye_cnss, details, payers

    if 'Type_Congé' not in df_analyse.columns: df_analyse['Type_Congé'] = ""
    agg_conges = df_analyse.groupby('Matricule')['Type_Congé'].apply(aggregate_conges).apply(pd.Series)
    if not agg_conges.empty:
        agg_conges.columns = ['Nb Jours Congé Payé', 'Nb Jours Congé Non Payé', 'Nb Jours Payé par CNSS', 'Détail des Congés', 'Payé Par']
    else:
        agg_conges = pd.DataFrame(columns=['Nb Jours Congé Payé', 'Nb Jours Congé Non Payé', 'Nb Jours Payé par CNSS', 'Détail des Congés', 'Payé Par'])

    resume_mensuel['Matricule'] = resume_mensuel['Matricule'].astype(str)
    agg_conges.index = agg_conges.index.astype(str)
    resume_mensuel = pd.merge(resume_mensuel, agg_conges, on='Matricule', how='left').fillna(0)
    
    # --- ### MODIFICATION : Calcul des Jours Payés ### ---
    # On soustrait les absences, les congés non payés et les congés payés par la CNSS
    resume_mensuel['Jours_Payés'] = CONFIG['BASE_JOURS_PAYES'] - resume_mensuel['Nb_Jours_Absence_Injustifiee'] - resume_mensuel['Nb Jours Congé Non Payé'] - resume_mensuel['Nb Jours Payé par CNSS']
    resume_mensuel['Jours_Payés'] = resume_mensuel['Jours_Payés'].clip(lower=0)
    
    resume_mensuel['Maj_25'] = (resume_mensuel['Heures_Sup_Maj25'] * 0.25).round(2)
    resume_mensuel['Maj_50'] = (resume_mensuel['Heures_Sup_Maj50'] * 0.50).round(2)
    resume_mensuel['Maj_100'] = (resume_mensuel['Heures_Sup_Maj100'] * 1.00).round(2)
    resume_mensuel['Total_Majorations'] = (resume_mensuel['Maj_25'] + resume_mensuel['Maj_50'] + resume_mensuel['Maj_100']).round(2)
    
    jours_absence_detail = df_analyse[df_analyse['Absence_Injustifiee']].groupby('Matricule')['Date'].apply(
        lambda x: ', '.join(sorted([d.strftime('%d/%m/%Y') for d in x]))
    ).reset_index(name='Détail des Absences')
    
    if not jours_absence_detail.empty:
        resume_mensuel = pd.merge(resume_mensuel, jours_absence_detail, on='Matricule', how='left')
        resume_mensuel['Détail des Absences'].fillna("", inplace=True)
    else:
        resume_mensuel['Détail des Absences'] = ''

    # --- ### NOUVEAU : Création de la colonne de détail des dates de congé ### ---
    jours_conge_detail = df_analyse[df_analyse['Type_Congé'] != ""].drop_duplicates(subset=['Matricule', 'Date']).groupby('Matricule')['Date'].apply(
    lambda x: ', '.join(sorted([d.strftime('%d/%m/%Y') for d in x]))
).reset_index(name='Détail des Jours de Congé')

    if not jours_conge_detail.empty:
        resume_mensuel = pd.merge(resume_mensuel, jours_conge_detail, on='Matricule', how='left')
        resume_mensuel['Détail des Jours de Congé'].fillna("", inplace=True)
    else:
        resume_mensuel['Détail des Jours de Congé'] = ''
        
    resume_mensuel = resume_mensuel.rename(columns={'Nb_Jours_Absence_Injustifiee': 'Nb Jours Absence Injustifiée'})
    
    resume_mensuel['Matricule_numeric'] = pd.to_numeric(resume_mensuel['Matricule'], errors='coerce')
    resume_mensuel.sort_values(by='Matricule_numeric', inplace=True, na_position='first')
    resume_mensuel.drop(columns=['Matricule_numeric'], inplace=True)
    
    return resume_mensuel

# --- FONCTION D'EXPORTATION ---
def exporter_excel(df_resultats, nom_fichier, colonnes_choisies):
    output = BytesIO()
    df_export = df_resultats.copy()
    
    # AJOUT des nouvelles colonnes pour l'export
    rename_dict_export = {
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
        'Payé Par': 'Payé Par'
    }
    df_export = df_export.rename(columns=rename_dict_export)
    
    colonnes_a_exporter = [col for col in colonnes_choisies if col in df_export.columns]
    df_final_export = df_export[colonnes_a_exporter]
    
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        df_final_export.to_excel(writer, index=False, sheet_name='Rapport')
        worksheet = writer.sheets['Rapport']
        for idx, col in enumerate(df_final_export):
            series = df_final_export[col]
            try:
                max_len = max((series.astype(str).map(len).max(), len(str(series.name)))) + 2
            except (ValueError, TypeError):
                max_len = len(str(col)) + 2
            worksheet.set_column(idx, idx, max_len)
            
    processed_data = output.getvalue()
    return processed_data