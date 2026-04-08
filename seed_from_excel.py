from pathlib import Path

from app import read_excel_file, replace_records_from_rows, init_db

BASE_DIR = Path(__file__).resolve().parent
SAMPLE_FILE = BASE_DIR / "sample_data" / "RENOVACION ABRIL.xlsm"

if __name__ == "__main__":
    init_db()
    if not SAMPLE_FILE.exists():
        raise FileNotFoundError(f"No se encontró el archivo de muestra: {SAMPLE_FILE}")
    rows = read_excel_file(str(SAMPLE_FILE))
    replace_records_from_rows(rows, SAMPLE_FILE.name)
    print(f"Carga completada. Registros importados: {len(rows)}")
