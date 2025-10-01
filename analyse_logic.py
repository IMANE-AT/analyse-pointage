import pandas as pd
from datetime import time, timedelta, date
from itertools import product
from io import BytesIO

CONFIG = {
    'BASE_JOURS_PAYES': 26, 'DUREE_JOURNEE_NORMALE_HEURES': 8.0, 'JOUR_DIMANCHE': 'Sunday',
    'HEURE_DEBUT_JOURNEE_NORMALE': time(8, 30), 'HEURE_DEBUT_NUIT': time(21, 0),
    'HEURE_FIN_NUIT': time(6, 0), 'JOUR_VENDREDI': 'Friday', 'JOUR_SAMEDI': 'Saturday',
    'HEURE_FIN_TRAVAIL_SAMEDI_MATIN': time(12, 30), 'HEURE_DEBUT_PAUSE_DEJ': time(12, 30),
    'HEURE_FIN_PAUSE_DEJ_LUN_JEU': time(14, 30), 'HEURE_FIN_PAUSE_DEJ_VENDREDI': time(15, 0),
    'TOLERANCE_RETARD_MIN': 40, 'HEURE_FIN_JOURNEE_NORMALE_STANDARD_THEORIQUE': time(18, 30),
    'HEURE_FIN_JOURNEE_NORMALE_VENDREDI_THEORIQUE': time(19, 0),
    'JOURS_OUVRABLES_LUN_JEU': ['Monday', 'Tuesday', 'Wednesday', 'Thursday']
}
REGLES_CONGES = {
    "CONGE_PAYE": {"mots_cles": ["annuel", "payé"], "duree_max": 18, "paye_par": "Employeur", "est_paye": True},
    "CONGE_MATERNITE": {"mots_cles": ["maternité", "maternite"], "duree_max": 98, "paye_par": "CNSS", "est_paye": True},
    "CONGE_MALADIE": {"mots_cles": ["maladie", "medical"], "duree_max": 180, "paye_par": "CNSS", "est_paye": True},
    "CONGE_MALADIE_COURT": {"mots_cles": [], "duree_max": 3, "paye_par": "Personne", "est_paye": False},
    "CONGE_MARIAGE": {"mots_cles": ["mariage"], "duree_max": 4, "paye_par": "Employeur", "est_paye": True},
    "CONGE_PATERNITE": {"mots_cles": ["paternité", "naissance"], "duree_max": 3, "paye_par": "Employeur", "est_paye": True},
    "CONGE_DECES": {"mots_cles": ["décès", "deces"], "duree_max": 3, "paye_par": "Employeur", "est_paye": True},
    "CONGE_SANS_SOLDE": {"mots_cles": ["sans solde", "non payé"], "duree_max": 999, "paye_par": "Personne", "est_paye": False},
    "AUTRE": {"mots_cles": [], "duree_max": 999, "paye_par": "Inconnu", "est_paye": False}
}

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
    df_conges['Date_Debut'] = pd.to_datetime(df_conges['Date_Debut'], errors='coerce', dayfirst=True)
    df_conges['Date_Fin'] = pd.to_datetime(df_conges['Date_Fin'], errors='coerce', dayfirst=True).fillna(df_conges['Date_Debut'])
    df_conges.dropna(subset=['Matricule', 'Date_Debut', 'Date_Fin'], inplace=True)
    def find_type(text):
        text_lower = str(text).lower()
        for rule_name, rule_details in REGLES_CONGES.items():
            if rule_name == "CONGE_MALADIE_COURT": continue
            if any(keyword in text_lower for keyword in rule_details['mots_cles']):
                return rule_name
        return "AUTRE"
    df_conges['Type_Congé_Standard'] = df_conges['Type_Congé'].apply(find_type)
    return df_conges

def prepare_affectations_df(df_affectations_raw):
    if df_affectations_raw.empty:
        return pd.DataFrame()
    affectations_cols_map = {
        'Matricule': ['matr', 'id'], 'Date': ['date'],
        'Affectation': ['affectation', 'tâche', 'tache', 'type'],
        'Lieu_Chantier': ['lieu', 'chantier', 'ville'],
        'Projet_Domicile': ['projet', 'domicile']
    }
    df_affectations = find_and_rename_header(df_affectations_raw.copy(), affectations_cols_map)
    return df_affectations

def get_full_date_range(mois, annee):
    start_date = f'{annee}-{mois:02d}-01'
    end_date = pd.to_datetime(start_date) + pd.offsets.MonthEnd(1)
    return pd.date_range(start=start_date, end=end_date, freq='D')

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


def calculer_presence_par_blocs(pointages_du_jour, jour_dt, jour_semaine):
    if pointages_du_jour.empty:
        type_absence = "Journée complète (Samedi)" if jour_semaine == CONFIG['JOUR_SAMEDI'] else "Journée complète"
        return 0.0, type_absence

    periodes_travail = []
    pointages = sorted(pointages_du_jour['Pointage'].tolist())
    
    pointage_isole = None
    if len(pointages) % 2 != 0:
        pointage_isole = pointages.pop()

    for i in range(0, len(pointages), 2):
        debut_travail = pointages[i]
        fin_travail = pointages[i+1]
        if pd.notna(debut_travail) and pd.notna(fin_travail):
            periodes_travail.append((debut_travail, fin_travail))

    if not periodes_travail and not pointage_isole:
        type_absence = "Journée complète (Samedi)" if jour_semaine == CONFIG['JOUR_SAMEDI'] else "Journée complète"
        return 0.0, type_absence

    date_jour = jour_dt.date()
    bloc_matin_debut = pd.to_datetime(f"{date_jour} {CONFIG['HEURE_DEBUT_JOURNEE_NORMALE']}")
    bloc_matin_fin = pd.to_datetime(f"{date_jour} {CONFIG['HEURE_DEBUT_PAUSE_DEJ']}")
    
    if jour_semaine == CONFIG['JOUR_VENDREDI']:
        bloc_soir_debut = pd.to_datetime(f"{date_jour} {CONFIG['HEURE_FIN_PAUSE_DEJ_VENDREDI']}")
        bloc_soir_fin = pd.to_datetime(f"{date_jour} {CONFIG['HEURE_FIN_JOURNEE_NORMALE_VENDREDI_THEORIQUE']}")
    else:
        bloc_soir_debut = pd.to_datetime(f"{date_jour} {CONFIG['HEURE_FIN_PAUSE_DEJ_LUN_JEU']}")
        bloc_soir_fin = pd.to_datetime(f"{date_jour} {CONFIG['HEURE_FIN_JOURNEE_NORMALE_STANDARD_THEORIQUE']}")

    presence_matin = False
    for debut, fin in periodes_travail:
        if debut < bloc_matin_fin and fin > bloc_matin_debut:
            presence_matin = True
            break

    presence_soir = False
    if jour_semaine != CONFIG['JOUR_SAMEDI']:
        for debut, fin in periodes_travail:
            if debut < bloc_soir_fin and fin > bloc_soir_debut:
                presence_soir = True
                break
    
    if pointage_isole:
        if not presence_matin and (bloc_matin_debut <= pointage_isole < bloc_matin_fin):
            presence_matin = True
            
        # --- DÉBUT DE LA MODIFICATION ---
        # Ancienne condition : if not presence_soir and (bloc_soir_debut <= pointage_isole < bloc_soir_fin):
        # Nouvelle condition : On vérifie juste si le pointage est après le début du bloc soir.
        if not presence_soir and (pointage_isole >= bloc_soir_debut):
            presence_soir = True
        # --- FIN DE LA MODIFICATION ---

    if jour_semaine == CONFIG['JOUR_SAMEDI']:
        return (1.0, "") if presence_matin else (0.0, "Journée complète (Samedi)")
    else:
        score = 0.0
        if presence_matin: score += 0.5
        if presence_soir: score += 0.5
        
        type_absence = ""
        if score == 0.0:
            type_absence = "Journée complète"
        elif score == 0.5:
            type_absence = "Soir" if presence_matin else "Matin"
        
        return score, type_absence

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
    return {
        'Heures_Bureau': round(heures_normales, 2), 'HS_Bureau_25': round(h_sup25, 2),
        'HS_Bureau_50': round(h_sup50, 2), 'HS_Bureau_100': round(h_sup100, 2),
        'Heures_Pause_Dej': round(heures_pause_dej, 2)
    }

def analyser_pointages(df_pointage_raw, df_conges_raw, df_affectations_file_raw, df_affectations_manuel, mois, annee, jours_feries):
    pointage_cols_map = {'Matricule': ['matr', 'id'], 'Pointage': ['pointage', 'date']}
    df_pointage = find_and_rename_header(df_pointage_raw.copy(), pointage_cols_map)
    if not df_pointage.empty:
        df_pointage['Pointage'] = pd.to_datetime(df_pointage['Pointage'], errors='coerce')
        df_pointage.dropna(subset=['Pointage'], inplace=True)
        df_pointage.sort_values(['Matricule', 'Pointage'], inplace=True)
        
        indices_a_garder = []
        last_matricule = None
        last_time_kept = pd.NaT
        intervalle_securite = timedelta(minutes=10)

        for index, row in df_pointage.iterrows():
            if row['Matricule'] != last_matricule:
                last_matricule = row['Matricule']
                last_time_kept = pd.NaT
            
            if pd.isna(last_time_kept) or (row['Pointage'] - last_time_kept > intervalle_securite):
                indices_a_garder.append(index)
                last_time_kept = row['Pointage']
        
        df_pointage = df_pointage.loc[indices_a_garder].reset_index(drop=True)
        df_pointage['Date'] = pd.to_datetime(df_pointage['Pointage'].dt.date)

    df_conges = prepare_conges_df(df_conges_raw)
    df_affectations_file = prepare_affectations_df(df_affectations_file_raw)
    df_affectations = pd.concat([df_affectations_file, df_affectations_manuel], ignore_index=True)
    if not df_affectations.empty:
        if 'Date' in df_affectations.columns:
            df_affectations['Date'] = pd.to_datetime(df_affectations['Date'], errors='coerce', dayfirst=True)
            df_affectations.dropna(subset=['Matricule', 'Date', 'Affectation'], inplace=True)
            df_affectations.drop_duplicates(subset=['Matricule', 'Date'], keep='last', inplace=True)
            df_affectations['Date'] = df_affectations['Date'].dt.normalize()
        else:
            df_affectations = pd.DataFrame()

    matricules_pointage = pd.Series(df_pointage['Matricule'].astype(str).unique()) if not df_pointage.empty else pd.Series([], dtype=str)
    matricules_conges = pd.Series(df_conges['Matricule'].astype(str).unique()) if not df_conges.empty else pd.Series([], dtype=str)
    matricules_affectations = pd.Series(df_affectations['Matricule'].astype(str).unique()) if not df_affectations.empty else pd.Series([], dtype=str)
    all_matricules = pd.concat([matricules_pointage, matricules_conges, matricules_affectations]).unique()

    if len(all_matricules) == 0: return pd.DataFrame()
    
    jours_feries_dates = {d.date() for d in pd.to_datetime(jours_feries)}
    full_date_range = get_full_date_range(mois, annee)
    
    resultats_journaliers = []
    for matricule in all_matricules:
        for jour in full_date_range:
            jour_dt = pd.to_datetime(jour)
            jour_semaine = jour_dt.day_name()
            est_jour_ferie = jour_dt.date() in jours_feries_dates
            est_jour_ouvrable = (jour_semaine != CONFIG['JOUR_DIMANCHE']) and not est_jour_ferie

            jour_actuel = {
                'Matricule': matricule, 'Date': jour_dt.date(), 'Jour_Semaine': jour_semaine, 
                'Est_JourFerie': est_jour_ferie, 'Est_Jour_Ouvrable': est_jour_ouvrable,
                'Type_Congé': '', 'Affectation': '', 'Lieu_Chantier': '', 'Projet_Domicile': '',
                'Jours_Presence_Travail': 0.0, 'Type_Absence_Jour': '', 'Statut_Jour': 'Non Travaillé',
                'Est_En_Retard': False, 'Heures_Bureau': 0, 'HS_Bureau_25': 0, 'HS_Bureau_50': 0,
                'HS_Bureau_100': 0, 'Heures_Chantier': 0, 'HS_Chantier_25': 0, 'HS_Chantier_50': 0,
                'HS_Chantier_100': 0, 'Heures_Domicile': 0, 'HS_Domicile_25': 0, 'HS_Domicile_50': 0,
                'HS_Domicile_100': 0,
            }

            conge_du_jour = df_conges[(df_conges['Matricule'].astype(str) == str(matricule)) & (df_conges['Date_Debut'] <= jour_dt) & (df_conges['Date_Fin'] >= jour_dt)]
            if not conge_du_jour.empty:
                type_conge = conge_du_jour.iloc[0]['Type_Congé_Standard']
                jour_est_decompte = (jour_dt.weekday() < 6 and not est_jour_ferie) or type_conge == "CONGE_MATERNITE"
                if jour_est_decompte:
                    jour_actuel['Type_Congé'] = type_conge
                    jour_actuel['Statut_Jour'] = 'En Congé'
                    resultats_journaliers.append(jour_actuel)
                    continue
            
            if not est_jour_ouvrable:
                resultats_journaliers.append(jour_actuel)
                continue

            affectation_du_jour = df_affectations[(df_affectations['Matricule'].astype(str) == str(matricule)) & (df_affectations['Date'].dt.date == jour_dt.date())] if not df_affectations.empty else pd.DataFrame()
            if not affectation_du_jour.empty:
                affectation_info = affectation_du_jour.iloc[0]
                jour_actuel.update({'Affectation': affectation_info.get('Affectation', ''), 'Lieu_Chantier': affectation_info.get('Lieu_Chantier', ''), 'Projet_Domicile': affectation_info.get('Projet_Domicile', '')})
                jour_actuel['Jours_Presence_Travail'] = 1.0
                jour_actuel['Statut_Jour'] = affectation_info.get('Affectation', 'Affectation')
                
                affectation_type = str(affectation_info.get('Affectation', '')).lower()
                if 'chantier' in affectation_type or 'domicile' in affectation_type:
                    debut_virtuel = pd.to_datetime(f"{jour.isoformat()} {CONFIG['HEURE_DEBUT_JOURNEE_NORMALE']}")
                    heure_fin_theorique = CONFIG['HEURE_FIN_TRAVAIL_SAMEDI_MATIN'] if jour_semaine == CONFIG['JOUR_SAMEDI'] else CONFIG['HEURE_FIN_JOURNEE_NORMALE_STANDARD_THEORIQUE']
                    fin_virtuel = pd.to_datetime(f"{jour.isoformat()} {heure_fin_theorique}")
                    hn, h25, h50, h100 = calculer_heures_supplementaires(debut_virtuel, fin_virtuel, jour_semaine, est_jour_ferie, 0)
                    
                    if 'chantier' in affectation_type: jour_actuel.update({'Heures_Chantier': hn, 'HS_Chantier_25': h25, 'HS_Chantier_50': h50, 'HS_Chantier_100': h100})
                    if 'domicile' in affectation_type: jour_actuel.update({'Heures_Domicile': hn, 'HS_Domicile_25': h25, 'HS_Domicile_50': h50, 'HS_Domicile_100': h100})
                resultats_journaliers.append(jour_actuel)
                continue

            pointages_du_jour = df_pointage[(df_pointage['Matricule'].astype(str) == str(matricule)) & (df_pointage['Date'] == jour_dt)] if not df_pointage.empty else pd.DataFrame()
            
            score_presence, type_absence = calculer_presence_par_blocs(pointages_du_jour, jour_dt, jour_semaine)
            
            jour_actuel['Jours_Presence_Travail'] = score_presence
            jour_actuel['Type_Absence_Jour'] = type_absence

            if score_presence > 0:
                jour_actuel['Statut_Jour'] = 'Bureau'
                pointages_du_jour_copy = pointages_du_jour.copy()
                pointages_du_jour_copy['Est_JourFerie'] = est_jour_ferie
                heures_calculees = calculer_indicateurs_jour_travaille(pointages_du_jour_copy)
                jour_actuel.update(heures_calculees)
                
                premier_pointage = pointages_du_jour.iloc[0]['Pointage']
                heure_debut_theorique = CONFIG['HEURE_DEBUT_JOURNEE_NORMALE']
                heure_debut_dt = pd.to_datetime(f"{jour_dt.date()} {heure_debut_theorique}")
                limite_retard = heure_debut_dt + timedelta(minutes=CONFIG['TOLERANCE_RETARD_MIN'])
                if premier_pointage > limite_retard: jour_actuel['Est_En_Retard'] = True
            
            elif est_jour_ouvrable:
                jour_actuel['Statut_Jour'] = 'Absence Injustifiée'
                
            resultats_journaliers.append(jour_actuel)

    if not resultats_journaliers: return pd.DataFrame()
    df_analyse = pd.DataFrame(resultats_journaliers)

    agg_dict = {
        'Heures_Bureau': ('Heures_Bureau', 'sum'), 'HS_Bureau_25': ('HS_Bureau_25', 'sum'),
        'HS_Bureau_50': ('HS_Bureau_50', 'sum'), 'HS_Bureau_100': ('HS_Bureau_100', 'sum'),
        'Heures_Chantier': ('Heures_Chantier', 'sum'), 'HS_Chantier_25': ('HS_Chantier_25', 'sum'),
        'HS_Chantier_50': ('HS_Chantier_50', 'sum'), 'HS_Chantier_100': ('HS_Chantier_100', 'sum'),
        'Heures_Domicile': ('Heures_Domicile', 'sum'), 'HS_Domicile_25': ('HS_Domicile_25', 'sum'),
        'HS_Domicile_50': ('HS_Domicile_50', 'sum'), 'HS_Domicile_100': ('HS_Domicile_100', 'sum'),
        'Lieu_Chantier': ('Lieu_Chantier', lambda x: ', '.join(x.dropna().astype(str).unique())),
        'Projet_Domicile': ('Projet_Domicile', lambda x: ', '.join(x.dropna().astype(str).unique())),
        'Nb Jours Chantier': ('Statut_Jour', lambda x: x.str.contains('Chantier').sum()),
        'Nb Jours Domicile': ('Statut_Jour', lambda x: x.str.contains('Domicile').sum()),
        'Nb_Retards': ('Est_En_Retard', 'sum')
    }
    resume_mensuel = df_analyse.groupby('Matricule').agg(**agg_dict).reset_index()

    def aggregate_conges(series):
        types = [REGLES_CONGES.get(t, {}) for t in series if t and t != ""]
        paye_employeur = sum(1 for t in types if t.get('paye_par') == 'Employeur')
        non_paye = sum(1 for t in types if not t.get('est_paye'))
        paye_cnss = sum(1 for t in types if t.get('paye_par') == 'CNSS')
        details_types = ", ".join(sorted(list(set(s for s in series if s and s != ""))))
        return paye_employeur, non_paye, paye_cnss, details_types
    agg_conges = df_analyse.groupby('Matricule')['Type_Congé'].apply(aggregate_conges).apply(pd.Series)
    if not agg_conges.empty:
        agg_conges.columns = ['Nb Jours Congé Payé par Employeur', 'Nb Jours Congé Non Payé', 'Nb Jours Payé par CNSS', 'Détail des Types de Congé']
        resume_mensuel = pd.merge(resume_mensuel, agg_conges, on='Matricule', how='left')
    
    resume_mensuel.fillna(0, inplace=True)

    df_analyse['Absence_Fraction'] = df_analyse.apply(
        lambda row: 1.0 - row['Jours_Presence_Travail']
        if row['Est_Jour_Ouvrable'] and not row['Type_Congé'] and not row['Affectation'] else 0.0,
        axis=1
    )
    total_absences = df_analyse.groupby('Matricule')['Absence_Fraction'].sum().reset_index(name='Nb Jours Absence Injustifiée')
    resume_mensuel = pd.merge(resume_mensuel, total_absences, on='Matricule', how='left')
    resume_mensuel['Nb Jours Absence Injustifiée'] = resume_mensuel['Nb Jours Absence Injustifiée'].fillna(0).round(2)
    
    # --- DÉBUT DE LA MODIFICATION ---
    # Retour à l'ancienne méthode de calcul des Jours Payés
    resume_mensuel['Jours Payés (par Employeur)'] = (
        CONFIG['BASE_JOURS_PAYES'] -
        resume_mensuel['Nb Jours Absence Injustifiée'] -
        resume_mensuel.get('Nb Jours Congé Non Payé', 0) -
        resume_mensuel.get('Nb Jours Payé par CNSS', 0)
    )
    resume_mensuel['Jours Payés (par Employeur)'] = resume_mensuel['Jours Payés (par Employeur)'].clip(lower=0).round(2)
    # --- FIN DE LA MODIFICATION ---
    
    def format_absence_detail(df_group):
        details = []
        for _, row in df_group.iterrows():
            if row['Type_Absence_Jour']:
                details.append(f"{row['Date'].strftime('%d/%m/%Y')} ({row['Type_Absence_Jour']})")
        return ', '.join(details)
    absence_details = df_analyse.groupby('Matricule').apply(format_absence_detail).reset_index(name='Détail des Absences')
    resume_mensuel = pd.merge(resume_mensuel, absence_details, on='Matricule', how='left')

    jours_conge_detail = df_analyse[df_analyse['Type_Congé'] != ""].groupby('Matricule')['Date'].apply(lambda x: ', '.join(sorted([d.strftime('%d/%m/%Y') for d in x]))).reset_index(name='Détail Jours de Congé')
    if not jours_conge_detail.empty:
        resume_mensuel = pd.merge(resume_mensuel, jours_conge_detail, on='Matricule', how='left')

    resume_mensuel['Total Heures Normales'] = (resume_mensuel['Heures_Bureau'] + resume_mensuel['Heures_Chantier'] + resume_mensuel['Heures_Domicile']).round(2)
    resume_mensuel['Total HS 25%'] = (resume_mensuel['HS_Bureau_25'] + resume_mensuel['HS_Chantier_25'] + resume_mensuel['HS_Domicile_25']).round(2)
    resume_mensuel['Total HS 50%'] = (resume_mensuel['HS_Bureau_50'] + resume_mensuel['HS_Chantier_50'] + resume_mensuel['HS_Domicile_50']).round(2)
    resume_mensuel['Total HS 100%'] = (resume_mensuel['HS_Bureau_100'] + resume_mensuel['HS_Chantier_100'] + resume_mensuel['HS_Domicile_100']).round(2)
    resume_mensuel['Majoration 25% (Val)'] = (resume_mensuel['Total HS 25%'] * 0.25).round(2)
    resume_mensuel['Majoration 50% (Val)'] = (resume_mensuel['Total HS 50%'] * 0.50).round(2)
    resume_mensuel['Majoration 100% (Val)'] = (resume_mensuel['Total HS 100%'] * 1.00).round(2)
    resume_mensuel['Total Majorations'] = (resume_mensuel['Majoration 25% (Val)'] + resume_mensuel['Majoration 50% (Val)'] + resume_mensuel['Majoration 100% (Val)']).round(2)

    jours_ouvrables = df_analyse[df_analyse['Est_Jour_Ouvrable']].groupby('Matricule').size().reset_index(name='Jours_Ouvrables')
    resume_mensuel = pd.merge(resume_mensuel, jours_ouvrables, on='Matricule', how='left')
    resume_mensuel['Jours_Ouvrables'] = resume_mensuel['Jours_Ouvrables'].fillna(0)
    resume_mensuel['Conges_Autorises'] = resume_mensuel.get('Nb Jours Congé Payé par Employeur', 0) + resume_mensuel.get('Nb Jours Congé Non Payé', 0) + resume_mensuel.get('Nb Jours Payé par CNSS', 0)
    resume_mensuel['Jours_Prevus'] = resume_mensuel['Jours_Ouvrables'] - resume_mensuel['Conges_Autorises']

    penalite_points = (1 * resume_mensuel['Nb_Retards']) + (4 * resume_mensuel['Nb Jours Absence Injustifiée'])
    
    
    bonus_points = resume_mensuel['Total Majorations']
    
    
    points_nets_penalite = (penalite_points - bonus_points).clip(lower=0) 
    
    
    score = 100 - (points_nets_penalite / resume_mensuel['Jours_Prevus'].replace(0, pd.NA)) * 100
    
    
    resume_mensuel['Score Discipline (%)'] = score.clip(0, 100).round(2).astype(str)
    resume_mensuel.loc[resume_mensuel['Jours_Prevus'] <= 0, 'Score Discipline (%)'] = '-'

    resume_mensuel.fillna({'Détail des Absences': '', 'Détail Jours de Congé': '', 'Détail des Types de Congé': ''}, inplace=True)
    resume_mensuel['Matricule_numeric'] = pd.to_numeric(resume_mensuel['Matricule'], errors='coerce')
    resume_mensuel.sort_values(by='Matricule_numeric', inplace=True, na_position='first')
    resume_mensuel.drop(columns=['Matricule_numeric', 'Jours_Ouvrables', 'Conges_Autorises', 'Jours_Prevus'], errors='ignore', inplace=True)
    
    return resume_mensuel

def exporter_excel(df_resultats, nom_fichier, colonnes_choisies, return_df=False):
    output = BytesIO()
    df_export = df_resultats.copy()
    
    if 'Matricule' in colonnes_choisies:
        colonnes_a_exporter = ['Matricule'] + [col for col in colonnes_choisies if col in df_export.columns and col != 'Matricule']
    else:
        colonnes_a_exporter = [col for col in colonnes_choisies if col in df_export.columns]
        
    df_final_export = df_export.reindex(columns=colonnes_a_exporter)
    
    if return_df: 
        return df_final_export

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