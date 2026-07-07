from datetime import datetime, timedelta
from airflow import DAG
from airflow.providers.docker.operators.docker import DockerOperator
from docker.types import Mount # 💡 IMPORTATION CRUCIALE ICI

default_args = {
    'owner': 'falilou_mfd',
    'depends_on_past': False,
    'start_date': datetime(2026, 1, 1),
    'retries': 1,
    'retry_delay': timedelta(minutes=1),
}

with DAG(
    'pipeline_ingestion_suivi',
    default_args=default_args,
    description='ETL Postgres vers DuckDB',
    schedule_interval='@hourly', 
    catchup=False,
) as dag:

    run_etl = DockerOperator(
        task_id='execution_pipeline_etl',
        image='suivi_data_processor:v1',
        # container_name='airflow_worker_etl_run',
        api_version='auto',
        auto_remove=True,
        force_pull=False,
        mount_tmp_dir=False,
        # network_mode='bridge', 
        # 💡 REMPLACEMENT DE 'volumes' PAR 'mounts'
        network_mode='suivi_network',
        mounts=[
            Mount(
                # source='duckdb_data', # Le nom du volume Docker
                source='dgid_duckdb_shared', # Le nom du volume Docker
                target='/app/duckdb_store',       # Le chemin de destination dans le conteneur
                type='volume'
            )
        ], 
        environment={
            'PYTHONUNBUFFERED': '1',
            'DB_USER': 'postgres',
            'DB_PASSWORD': '12345678',
            'DB_HOST': 'postgres_db',
            'DUCKDB_PATH': '/app/duckdb_store/analytics.duckdb'
        }
    )

    run_etl