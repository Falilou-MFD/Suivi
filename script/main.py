import os
from datetime import datetime
import pandas as pd
from sqlalchemy import create_engine
import duckdb

def get_db_engine(db_name):
    DB_USER = os.getenv("DB_USER", "postgres")
    DB_PASSWORD = os.getenv("DB_PASSWORD", "12345678") 
    DB_HOST = os.getenv("DB_HOST", "postgres_db")      
    DB_PORT = os.getenv("DB_PORT", "5432")
    connection_string = f"postgresql://{DB_USER}:{DB_PASSWORD}@{DB_HOST}:{DB_PORT}/{db_name}"
    return create_engine(connection_string)

def extract_and_transform():
    # --- CONFIGURATION MULTI-SITES ---
    SITES_CONFIG = {
        "Dakar Plateau": "dkp",
        "Ngor-Almadies": "nga",
        "Mbour": "mb"
    }
    
    BUREAUX_CONFIG = {
        "cadastre": "Cadastre",
        "domaines": "Domaines",
        "conservation": "Conservation"
    }
    
    all_dataframes = []

    # Requête de suivi des mouvements
    query_mouvements = """
        SELECT 
            u.first_name || ' ' || u.last_name AS operateur,
            r.record_date AS registration_date,
            r.return_date AS date_retour,
            r.folder_is_update AS dossier_mis_a_jour,
            r.is_for_scan, 
            r.folder_number AS numero_dossier,
            c.code AS csf_code, 
            c.label AS csf_label,
            cl.value AS motif_enregistrement, 
            i.is_unique_item,
            f.label AS folder_label
        FROM public.e_s_record r
        LEFT JOIN public.user_entity u ON r.user_id = u.id
        LEFT JOIN public.csf c ON r.csf_uid = c.uid
        LEFT JOIN public.code_list cl ON r.record_type = cl.uid
        LEFT JOIN public.recorded_item ri ON r.uid = ri.record_ref
        LEFT JOIN public.item i ON ri.item_ref = i.uid
        LEFT JOIN public.folder f ON i.folder_id = f.uid;
    """

    # Requête d'indexation
    query_index = """
        SELECT 
            gi.folder_number AS numero_dossier, 
            gi.cote, 
            gi.count_sheet, 
            gi.count_item, 
            gi.date_indexation, 
            gi.update_date AS date_maj,
            COALESCE(prop.nb_proprietaires_uniques, 0) AS nb_proprietaires_uniques
        FROM public.general_index gi
        LEFT JOIN (
            SELECT 
                general_index_id, 
                COUNT(DISTINCT (nom || ' ' || prenoms)) AS nb_proprietaires_uniques
            FROM public.onomastique_index
            GROUP BY general_index_id
        ) prop ON gi.uid = prop.general_index_id;
    """

    # --- BOUCLES D'EXTRACTION SUR TOUS LES SITES ET BUREAUX ---
    for current_csf_name, site_code in SITES_CONFIG.items():
        print(f"\n🌍 Début de l'extraction pour le site : {current_csf_name} ({site_code})")
        
        for bureau_prefix, service_label in BUREAUX_CONFIG.items():
            # Correspondance parfaite avec les bases créées par init-multiple-dbs.sh
            db_name = f"{bureau_prefix}_{site_code}_db"
            print(f"  📥 Extraction de {service_label} depuis [{db_name}]...")
            
            try:
                engine = get_db_engine(db_name)
                df_mvt = pd.read_sql_query(query_mouvements, con=engine)
                df_idx = pd.read_sql_query(query_index, con=engine)
                
                # Normalisation des clés de jointure
                df_mvt['numero_dossier'] = df_mvt['numero_dossier'].astype(str).str.strip()
                df_idx['numero_dossier'] = df_idx['numero_dossier'].astype(str).str.strip()

                # Fusion des mouvements et de l'indexation
                df_service = pd.merge(df_mvt, df_idx, on="numero_dossier", how="left")
                df_service["service_origine"] = service_label
                df_service["csf_geographique"] = current_csf_name
                
                # --- STRATÉGIE ANTI-CUMUL ET DÉDUPLICATION ---
                df_service['registration_date'] = pd.to_datetime(df_service['registration_date'], errors='coerce')
                df_service['date_maj'] = pd.to_datetime(df_service['date_maj'], errors='coerce')
                df_service['date_ref_fraicheur'] = df_service['registration_date'].combine_first(df_service['date_maj'])
                
                df_service = df_service.sort_values(by='date_ref_fraicheur', ascending=False)
                df_service = df_service.drop_duplicates(subset=['numero_dossier'], keep='first')
                df_service = df_service.drop(columns=['date_ref_fraicheur'])
                
                all_dataframes.append(df_service)
                
            except Exception as e:
                print(f"  ⚠️ Échec sur {db_name} : {e}")

    if not all_dataframes:
        print("❌ Aucune donnée extraite sur l'ensemble des sites.")
        return

    # Regroupement des données
    df_total = pd.concat(all_dataframes, ignore_index=True)

    colonnes_dates = ["registration_date", "date_retour", "date_indexation", "date_maj"]
    for col in colonnes_dates:
        if col in df_total.columns:
            df_total[col] = pd.to_datetime(df_total[col], errors='coerce').dt.tz_localize(None)

    df_total['operateur'] = df_total.get('operateur', pd.Series()).fillna('Operateur Inconnu').str.title()
    df_total['folder_label'] = df_total.get('folder_label', pd.Series()).fillna('')
    df_total['anomalie_chronologie'] = (df_total['date_retour'] < df_total['registration_date']).fillna(False)

    # --- PERSISTANCE DUCKDB ---
    duckdb_path = os.getenv("DUCKDB_PATH", "/app/duckdb_store/analytics.duckdb")
    print(f"\n🚀 Initialisation de la base DuckDB à l'emplacement : {duckdb_path}")
    
    dossier_parent = os.path.dirname(duckdb_path)
    if dossier_parent and not os.path.exists(dossier_parent):
        os.makedirs(dossier_parent, exist_ok=True)

    con = duckdb.connect(duckdb_path)
    
    table_exists = con.execute("SELECT count(*) FROM information_schema.tables WHERE table_name = 'fact_suivi_global'").fetchone()[0] > 0
    
    if not table_exists:
        con.execute("CREATE TABLE fact_suivi_global AS SELECT * FROM df_total WHERE 1=0")
    else:
        res = con.execute("SELECT * FROM fact_suivi_global LIMIT 0")
        existing_columns = [desc[0] for desc in res.description]
        
        for col in df_total.columns:
            if col not in existing_columns:
                if df_total[col].dtype == 'bool': 
                    col_type = "BOOLEAN"
                elif "date" in col or "maj" in col:
                    col_type = "TIMESTAMP"
                elif "nb_" in col or "count_" in col:
                    col_type = "BIGINT"
                elif df_total[col].dtype == 'float64':
                    col_type = "DOUBLE"
                else:
                    col_type = "VARCHAR"
                con.execute(f"ALTER TABLE fact_suivi_global ADD COLUMN {col} {col_type}")

    # Nettoyage TOTAL de la table et réinsertion
    con.execute("DELETE FROM fact_suivi_global")
    
    final_db_columns = [desc[0] for desc in con.execute("SELECT * FROM fact_suivi_global LIMIT 0").description]
    df_to_insert = df_total.reindex(columns=final_db_columns)
    
    con.execute("INSERT INTO fact_suivi_global SELECT * FROM df_to_insert")
    
    # Vérification
    total_lignes = con.execute("SELECT COUNT(*) FROM fact_suivi_global").fetchone()[0]
    resume_sites = con.execute("SELECT csf_geographique, COUNT(*) FROM fact_suivi_global GROUP BY csf_geographique").df()
    
    con.close()

    try:
        os.chmod(duckdb_path, 0o666)
        print(f"🔓 Permissions ajustées (666) sur {duckdb_path} pour Reflex.")
    except Exception as perm_err:
        print(f"⚠️ Impossible d'ajuster les permissions du fichier : {perm_err}")
            
    print(f"✅ Data Warehouse DuckDB mis à jour avec succès !")
    print(f"📊 Total des dossiers consolidés : {total_lignes}")
    print(resume_sites)

    
    # print(f"✅ Data Warehouse DuckDB mis à jour avec succès !")
    # print(f"📊 Total des dossiers consolidés : {total_lignes}")
    # print(resume_sites)

if __name__ == "__main__":
    extract_and_transform()