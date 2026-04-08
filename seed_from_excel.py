from app import init_db, read_excel_file, replace_records_from_dataframe, app
import os

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
FILE_PATH = os.path.join(BASE_DIR, "sample_data", "RENOVACION ABRIL.xlsm")

if __name__ == "__main__":
    init_db()
    with app.app_context():
        df = read_excel_file(FILE_PATH)
        replace_records_from_dataframe(df, os.path.basename(FILE_PATH))
    print(f"Base precargada con {len(df)} registros.")
