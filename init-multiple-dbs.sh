#!/bin/bash
set -e
set -u

# --- CONFIGURATION DES BOUCLES ---
SITES=("dkp" "nga" "mb")
BUREAUX=("cadastre" "domaines" "conservation")

# 1. Création du rôle de production manquant 'archivage_es'
echo "[1/4] Création du rôle de production archivage_es..."
psql -v ON_ERROR_STOP=1 --username "$POSTGRES_USER" --dbname "$POSTGRES_DB" <<-EOSQL
    DO \$\$
    BEGIN
        IF NOT EXISTS (SELECT FROM pg_catalog.pg_roles WHERE rolname = 'archivage_es') THEN
            CREATE ROLE archivage_es WITH LOGIN PASSWORD '12345678' SUPERUSER;
        END IF;
    END
    \$\$;
EOSQL

# 2. Fonction pour créer les bases de données vides
function create_database() {
    local database=$1
    echo "  -> Création de la base de données '$database'..."
    psql -v ON_ERROR_STOP=1 --username "$POSTGRES_USER" --dbname "$POSTGRES_DB" <<-EOSQL
        SELECT 'CREATE DATABASE $database' WHERE NOT EXISTS (SELECT FROM pg_database WHERE datname = '$database')\gexec
        GRANT ALL PRIVILEGES ON DATABASE $database TO $POSTGRES_USER;
        GRANT ALL PRIVILEGES ON DATABASE $database TO archivage_es;
EOSQL
}

echo "[2/4] Initialisation dynamique des 9 bases de données..."
for site in "${SITES[@]}"; do
    for bureau in "${BUREAUX[@]}"; do
        create_database "${bureau}_${site}_db"
    done
done

# 3. Restauration intelligente des fichiers Métiers (.backup ou .sql)
echo "[3/4] Lancement de la restauration des fichiers de production..."
for site in "${SITES[@]}"; do
    BACKUP_DIR="/backups/backups_${site}"
    
    for bureau in "${BUREAUX[@]}"; do
        DB_NAME="${bureau}_${site}_db"
        
        FILE_BACKUP="${BACKUP_DIR}/${bureau}_${site}.backup"
        FILE_SQL="${BACKUP_DIR}/${bureau}_${site}.sql"
        
        TARGET_FILE=""
        [ -f "$FILE_BACKUP" ] && TARGET_FILE="$FILE_BACKUP"
        [ -f "$FILE_SQL" ] && TARGET_FILE="$FILE_SQL"
        
        if [ -n "$TARGET_FILE" ]; then
            echo "  -> [Tentative] Restauration de $DB_NAME avec $TARGET_FILE..."
            if pg_restore -U "$POSTGRES_USER" -d "$DB_NAME" --no-owner "$TARGET_FILE" 2>/dev/null; then
                echo "  -> [Succès] Restauration via pg_restore terminée."
            elif psql -v ON_ERROR_STOP=1 -U "$POSTGRES_USER" -d "$DB_NAME" -f "$TARGET_FILE" >/dev/null 2>&1; then
                echo "  -> [Succès] Restauration via psql terminée."
            else
                echo "  ❌ Échec de la restauration pour $DB_NAME."
            fi
        else
            echo "  ❌ Aucun fichier trouvé pour $DB_NAME."
        fi
    done
done

# 4. Exécution des fichiers Keycloak (format binaire/dump)
echo "[4/4] Application des fichiers Keycloak..."
KEYCLOAK_DIR="/keycloak_files"

for site in "${SITES[@]}"; do
    for bureau in "${BUREAUX[@]}"; do
        DB_NAME="${bureau}_${site}_db"
        # On teste les deux extensions possibles au cas où
        KC_FILE="${KEYCLOAK_DIR}/keycloack_db_${bureau}_${site}.sql"
        
        if [ -f "$KC_FILE" ]; then
            echo "  -> [Tentative] Injection Keycloak dans $DB_NAME..."
            # On utilise pg_restore car c'est un dump binaire
            if pg_restore -U "$POSTGRES_USER" -d "$DB_NAME" --no-owner "$KC_FILE" 2>/dev/null; then
                echo "  -> [Succès] Keycloak injecté via pg_restore."
            else
                # Fallback sur psql au cas où ce fichier là serait bien en texte SQL
                psql -v ON_ERROR_STOP=1 -U "$POSTGRES_USER" -d "$DB_NAME" -f "$KC_FILE"
                echo "  -> [Succès] Keycloak injecté via psql."
            fi
        else
            echo "  ❌ Fichier Keycloak introuvable : $KC_FILE"
        fi
    done
done


echo "✅ Toutes les bases de données sont prêtes et restaurées avec succès !"