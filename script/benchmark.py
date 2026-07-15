import os
import time
import statistics
import psycopg2
import duckdb

DB_HOST = os.getenv("DB_HOST", "postgres_db")
DB_USER = os.getenv("DB_USER", "postgres")
DB_PASSWORD = os.getenv("DB_PASSWORD", "12345678")
DUCKDB_PATH = "/app/duckdb_store/analytics.duckdb"
TABLE_NAME = "e_s_record"

DATABASES_TO_CHECK = [
    "cadastre_dkp_db", "cadastre_mb_db", "cadastre_nga_db",
    "conservation_dkp_db", "conservation_mb_db", "conservation_nga_db",
    "domaines_dkp_db", "domaines_mb_db", "domaines_nga_db"
]

def discover_correct_db():
    for db in DATABASES_TO_CHECK:
        try:
            conn = psycopg2.connect(host=DB_HOST, database=db, user=DB_USER, password=DB_PASSWORD, port=5432)
            with conn.cursor() as cur:
                cur.execute(f"SELECT EXISTS (SELECT FROM information_schema.tables WHERE table_schema = 'public' AND table_name = '{TABLE_NAME}');")
                exists = cur.fetchone()[0]
            conn.close()
            if exists:
                return db
        except Exception:
            continue
    raise ValueError(f"Impossible de trouver 'public.{TABLE_NAME}' dans les bases.")

def init_duckdb_data(target_db):
    print(f"[INIT] Synchronisation de DuckDB avec {target_db}...")
    os.makedirs(os.path.dirname(DUCKDB_PATH), exist_ok=True)
    conn = duckdb.connect(DUCKDB_PATH)
    conn.execute("INSTALL postgres;")
    conn.execute("LOAD postgres;")
    conn.execute(f"ATTACH 'host={DB_HOST} user={DB_USER} password={DB_PASSWORD} dbname={target_db} port=5432' AS pg_source (TYPE POSTGRES);")
    conn.execute(f"CREATE OR REPLACE TABLE {TABLE_NAME} AS SELECT * FROM pg_source.public.{TABLE_NAME};")
    count = conn.execute(f"SELECT COUNT(*) FROM {TABLE_NAME};").fetchone()[0]
    print(f"[INIT] Table {TABLE_NAME} importée ({count} lignes).\n")
    conn.close()

def run_test(target_db, query_pg, query_duck, title, iterations=30):
    print(f"Série : {title}")
    print("-" * 60)
    
    # PostgreSQL
    pg_conn = psycopg2.connect(host=DB_HOST, database=target_db, user=DB_USER, password=DB_PASSWORD, port=5432)
    pg_times = []
    for _ in range(iterations):
        start = time.perf_counter()
        with pg_conn.cursor() as cur:
            cur.execute(query_pg)
            cur.fetchall()
        pg_times.append((time.perf_counter() - start) * 1000)
    pg_conn.close()
    
    # DuckDB
    duck_conn = duckdb.connect(DUCKDB_PATH)
    duck_times = []
    for _ in range(iterations):
        start = time.perf_counter()
        duck_conn.execute(query_duck).fetchall()
        duck_times.append((time.perf_counter() - start) * 1000)
    duck_conn.close()
    
    pg_med = statistics.median(pg_times)
    duck_med = statistics.median(duck_times)
    
    print(f"POSTGRESQL -> Médiane : {pg_med:.3f} ms [Min : {min(pg_times):.3f} ms | Max : {max(pg_times):.3f} ms]")
    print(f"DUCKDB     -> Médiane : {duck_med:.3f} ms [Min : {min(duck_times):.3f} ms | Max : {max(duck_times):.3f} ms]")
    
    ratio = pg_med / duck_med if duck_med > 0 else 0
    print(f"Résultat : DuckDB est {ratio:.2f}x plus rapide.\n")

if __name__ == "__main__":
    try:
        active_db = discover_correct_db()
        init_duckdb_data(active_db)
        
        # TEST 1 : Tri complet (Avantage structurel à PostgreSQL sur les faibles volumes)
        q1_pg = f"SELECT * FROM public.{TABLE_NAME} ORDER BY uid ASC;"
        q1_duck = f"SELECT * FROM {TABLE_NAME} ORDER BY uid ASC;"
        run_test(active_db, q1_pg, q1_duck, f"TEST 1 : SELECT * FROM {TABLE_NAME} ORDER BY uid")
        
        # TEST 2 : Agrégation basée sur le premier caractère de l'UUID (Nettoyé du bug de type)
        q2_pg = f"SELECT SUBSTRING(uid::text FROM 1 FOR 1) as prefixe, COUNT(*) FROM public.{TABLE_NAME} GROUP BY 1 ORDER BY 1;"
        q2_duck = f"SELECT SUBSTRING(uid::text FROM 1 FOR 1) as prefixe, COUNT(*) FROM {TABLE_NAME} GROUP BY 1 ORDER BY 1;"
        run_test(active_db, q2_pg, q2_duck, "TEST 2 : Agrégation et Transformation de chaînes (UUID)")
        
    except Exception as e:
        print(f"Erreur d'exécution : {e}")