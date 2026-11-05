import os
import re
import pandas as pd
import warnings
import time
import datetime
from sqlalchemy import create_engine
from concurrent.futures import ThreadPoolExecutor, as_completed
from dotenv import load_dotenv

warnings.filterwarnings("ignore", category=UserWarning, module="pandas")

# ---------------------------------------------------------
# 📅 ASK USER FOR DATE
# ---------------------------------------------------------
def ask_date():
    while True:
        date_str = input("📅 Veuillez entrer la date (format JJ/MM/AA) : ")
        try:
            datetime.datetime.strptime(date_str, "%d/%m/%y")
            return date_str
        except ValueError:
            print("❌ Format invalide ! Exemple valide : 04/12/25")

USER_DATE = ask_date()
print(f"➡️ Date utilisée pour les extractions : {USER_DATE}\n")

# ---------------------------------------------------------
# 🔐 DB CONNECTION
# ---------------------------------------------------------
load_dotenv()

username = os.getenv("ORACLE_USER")
password = os.getenv("ORACLE_PASSWORD")
host = os.getenv("ORACLE_HOST", "192.168.0.166")
port = os.getenv("ORACLE_PORT", "1521")
service_name = os.getenv("ORACLE_SERVICE", "TPXPROD")
MAX_THREADS = int(os.getenv("MAX_THREADS", 3))

SQL_FOLDER = "sql"
OUTPUT_FOLDER = "roquete"

dsn = f"(DESCRIPTION=(ADDRESS=(PROTOCOL=TCP)(HOST={host})(PORT={port}))(CONNECT_DATA=(SERVICE_NAME={service_name})))"
engine = create_engine(f"oracle+oracledb://{username}:{password}@{dsn}")

# ---------------------------------------------------------
# 🔥 HANDLE &date → :DATE_PARAM
# ---------------------------------------------------------
def handle_date_parameter(sql_query):
    sql_query = re.sub(r"['\"]\s*&date\s*['\"]", "&date", sql_query, flags=re.IGNORECASE)
    sql_query = re.sub(r"&date", ":DATE_PARAM", sql_query, flags=re.IGNORECASE)

    bind_match = re.search(r":([A-Za-z0-9_]+)", sql_query)
    if bind_match:
        bind_name = bind_match.group(1)
        return sql_query, {bind_name: USER_DATE}

    return sql_query, None


# ---------------------------------------------------------
# 📄 READ SQL FILE
# ---------------------------------------------------------
def read_sql_file(file_path):
    try:
        with open(file_path, "r", encoding="utf-8") as f:
            return f.read()
    except UnicodeDecodeError:
        try:
            with open(file_path, "r", encoding="latin-1") as f:
                return f.read()
        except UnicodeDecodeError:
            with open(file_path, "r", encoding="utf-8", errors="replace") as f:
                return f.read()


# ---------------------------------------------------------
# 📊 SAVE DATAFRAME TO EXCEL (AUTO SPLIT)
# ---------------------------------------------------------
def save_dataframe_to_excel(df, output_file, sheet_prefix):
    MAX_ROWS = 1048576  # Excel limit

    if len(df) <= MAX_ROWS:
        df.to_excel(output_file, index=False)
        return

    with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
        for i in range(0, len(df), MAX_ROWS):
            chunk = df.iloc[i:i + MAX_ROWS]
            sheet_name = f"{sheet_prefix}_{i // MAX_ROWS + 1}"
            chunk.to_excel(writer, sheet_name=sheet_name, index=False)


# ---------------------------------------------------------
# 🧠 EXECUTE SQL
# ---------------------------------------------------------
def execute_sql_file(file_path, engine):
    sql_name = os.path.splitext(os.path.basename(file_path))[0]
    output_base = os.path.join(OUTPUT_FOLDER, sql_name)
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")

    print(f"\n--- 🚀 Starting extraction: {sql_name} ---")

    try:
        sql_query = read_sql_file(file_path)
        print(f"✅ File loaded: {sql_name}")
    except Exception as e:
        print(f"❌ Error reading {sql_name}: {e}")
        return

    sql_query, params = handle_date_parameter(sql_query)

    try:
        connection = engine.connect()
        print("✅ Connected to Oracle\n")
        print("➡️ Bind parameters :", params)

        if params:
            df = pd.read_sql(sql_query, connection, params=params)
        else:
            df = pd.read_sql(sql_query, connection)

        print(f"✅ {len(df):,} rows extracted\n")

        output_file = f"{output_base}_{timestamp}.xlsx"
        save_dataframe_to_excel(df, output_file, sql_name)

        print(f"📁 Saved : {output_file}")

    except Exception as e:
        print(f"❌ SQL Error for {sql_name}: {e}")

    finally:
        try:
            connection.close()
            print(f"🔒 Connection closed for {sql_name}\n")
        except:
            pass


# ---------------------------------------------------------
# 🧵 MAIN LOOP
# ---------------------------------------------------------
if not os.path.exists(OUTPUT_FOLDER):
    os.makedirs(OUTPUT_FOLDER)

sql_files = [f for f in os.listdir(SQL_FOLDER) if f.endswith(".sql") or f.endswith(".txt")]

if not sql_files:
    print("⚠️ Aucun fichier SQL trouvé dans /sql")
else:
    print(f"🧵 Starting with {MAX_THREADS} threads...\n")
    start_time = time.time()

    with ThreadPoolExecutor(max_workers=MAX_THREADS) as executor:
        futures = {
            executor.submit(execute_sql_file, os.path.join(SQL_FOLDER, f), engine): f
            for f in sql_files
        }

        for future in as_completed(futures):
            try:
                future.result()
            except Exception as e:
                print(f"❌ Thread error: {e}")

    print(f"\n✅ DONE in {(time.time() - start_time)/60:.2f} min.")
