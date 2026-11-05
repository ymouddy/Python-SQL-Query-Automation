import pandas as pd
from sqlalchemy import create_engine
import warnings
import time
import datetime

# 🔕 Ignorer les avertissements inutiles
warnings.filterwarnings("ignore", category=UserWarning, module="pandas")

# --- 🔐 Connexion Oracle ---
username = "GHAMADA"
password = "password1test"
dsn = "192.168.0.166:1521/TPXPROD"

# --- 📘 Création du moteur SQLAlchemy ---
engine = create_engine(f"oracle+oracledb://{username}:{password}@{dsn}")

# --- 🧠 Requête SQL avec filtre sur la date de création ---
sql_query = """
SELECT 
STMSITE CODE_SITE,
ARTCEXR CODE_ARTICLE,
V5PROD.PKSTRUCOBJ.GET_DESC(ARTCINR,'FR') LIBELLE,
ARTSTAT CODE_STAT,
V5PROD.PKARTCOCA.GET_CLOSESTEAN(ARUCINL) CODE_CAISSE, 
V5PROD.PKSTRUCOBJ.GET_CEXT(V5PROD.PKSTRUCREL.GET_NIVEAU(1,ARTCINR,2,SYSDATE)) CODE_GRP,
V5PROD.PKSTRUCOBJ.GET_DESC(V5PROD.PKSTRUCREL.GET_NIVEAU(1,ARTCINR,2,SYSDATE),'FR') GRP,
V5PROD.PKSTRUCOBJ.GET_CEXT(V5PROD.PKSTRUCREL.GET_NIVEAU(1,ARTCINR,3,SYSDATE)) CODE_DEP,
V5PROD.PKSTRUCOBJ.GET_DESC(V5PROD.PKSTRUCREL.GET_NIVEAU(1,ARTCINR,3,SYSDATE),'FR') DEP,
V5PROD.PKSTRUCOBJ.GET_CEXT(V5PROD.PKSTRUCREL.GET_NIVEAU(1,ARTCINR,4,SYSDATE)) CODE_RAYON,
V5PROD.PKSTRUCOBJ.GET_DESC(V5PROD.PKSTRUCREL.GET_NIVEAU(1,ARTCINR,4,SYSDATE),'FR') RAYON,
V5PROD.PKSTRUCOBJ.GET_CEXT(V5PROD.PKSTRUCREL.GET_NIVEAU(1,ARTCINR,5,SYSDATE)) CODE_FAM,
V5PROD.PKSTRUCOBJ.GET_DESC(V5PROD.PKSTRUCREL.GET_NIVEAU(1,ARTCINR,5,SYSDATE),'FR') FAM,
V5PROD.PKSTRUCOBJ.GET_CEXT(V5PROD.PKSTRUCREL.GET_NIVEAU(1,ARTCINR,6,SYSDATE)) CODE_SFAM,
V5PROD.PKSTRUCOBJ.GET_DESC(V5PROD.PKSTRUCREL.GET_NIVEAU(1,ARTCINR,6,SYSDATE),'FR') SFAM,
V5PROD.PKSTRUCOBJ.GET_CEXT(V5PROD.PKSTRUCREL.GET_NIVEAU(1,ARTCINR,7,SYSDATE)) CODE_SSFAM,
V5PROD.PKSTRUCOBJ.GET_DESC(V5PROD.PKSTRUCREL.GET_NIVEAU(1,ARTCINR,7,SYSDATE),'FR') SSFAM,
AATCATT CNUF,
V5PROD.PKFOUDGENE.GET_DESCRIPTIONCNUF(AATCATT) LIB_FOURNISSEUR,
-SUM(STMVAL) QTE_VENDUE,
-SUM(STMVPV) CA_TTC,
-SUM(STMTVA) VALEUR_TVA,
-SUM(STMVPV-STMTVA) CA_HT,
-SUM(STMVPR) VALEUR_PR,
-SUM(NVL(STMVPV,0) - NVL(STMTVA,0) - NVL(STMVPR,0)) MARGE_HT
FROM V5PROD.STOMVT,V5PROD.ARTUL,V5PROD.ARTRAC,V5PROD.ARTATTRI
WHERE STMCINL=ARUCINL
AND ARUCINR=ARTCINR
AND ARTCINR=AATCINR
AND AATCCLA = 'CNUF' AND TRUNC(SYSDATE) BETWEEN TRUNC(AATDDEB) AND TRUNC(AATDFIN)
AND STMTMVT=150

AND V5PROD.PKSTRUCREL.GET_NIVEAU(1,ARTCINR,4,SYSDATE) IN  (258732)

AND STMDMVT >= TO_DATE('16/10/2025','DD/MM/RRRR') AND STMDMVT < TO_DATE('17/10/2025','DD/MM/RRRR') -- Besoin cloture 1
--AND STMDMVT >= TO_DATE('01/02/2017','DD/MM/RRRR') AND STMDMVT < TO_DATE('01/03/2017','DD/MM/RRRR') -- Besoin cloture 2
GROUP BY STMSITE,ARTCEXR,ARTCINR,ARTSTAT,ARUCINL,AATCATT
"""

# --- 📁 Nom de base du fichier Excel ---
excel_base = "Ca_tr_part"

# --- 📅 Générer un timestamp pour éviter les conflits ---
timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

try:
    # --- Temps de connexion ---
    start_conn = time.time()
    connection = engine.connect()
    end_conn = time.time()
    conn_time_min = (end_conn - start_conn) / 60
    print(f"✅ Successfully connected! (Connection time: {conn_time_min:.2f} min)\n")

    # --- Temps d'extraction ---
    start_extract = time.time()
    df = pd.read_sql(sql_query, connection)
    end_extract = time.time()
    extract_time_min = (end_extract - start_extract) / 60
    total_rows = len(df)
    print(f"✅ {total_rows:,} ligne(s) récupérée(s) depuis Oracle.")
    print(f"⏱ Extraction time: {extract_time_min:.2f} min\n")

    # 📊 Découpage si > 1 048 576 lignes
    max_rows_excel = 1_048_000
    if total_rows > max_rows_excel:
        print("⚠️ Données trop volumineuses pour un seul fichier Excel — division automatique...")
        chunks = [df[i:i + max_rows_excel] for i in range(0, total_rows, max_rows_excel)]

        for idx, chunk in enumerate(chunks, start=1):
            part_file = f"{excel_base}_part_{idx}_{timestamp}.xlsx"
            start_write = time.time()
            chunk.to_excel(part_file, index=False)
            end_write = time.time()
            write_time_min = (end_write - start_write) / 60
            print(f"📁 Partie {idx} sauvegardée : {part_file} ({len(chunk):,} lignes, write time: {write_time_min:.2f} min)")

        print("\n✅ Export terminé — plusieurs fichiers générés avec succès.")
    else:
        output_file = f"{excel_base}_{timestamp}.xlsx"
        start_write = time.time()
        df.to_excel(output_file, index=False)
        end_write = time.time()
        write_time_min = (end_write - start_write) / 60
        print(f"📊 Résultat sauvegardé dans : {output_file} (write time: {write_time_min:.2f} min)")

except Exception as e:
    print(f"\n❌ Unexpected error: {e}")

finally:
    # --- Temps de fermeture ---
    if 'connection' in locals():
        start_close = time.time()
        connection.close()
        end_close = time.time()
        close_time_min = (end_close - start_close) / 60
        print(f"\n🔒 Connection closed. (Close time: {close_time_min:.2f} min)")
