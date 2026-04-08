import os
import sqlite3
from datetime import datetime, date
from functools import wraps
from io import BytesIO

import pandas as pd
from flask import (
    Flask,
    flash,
    g,
    redirect,
    render_template,
    request,
    send_file,
    session,
    url_for,
)
from werkzeug.security import check_password_hash, generate_password_hash
from werkzeug.utils import secure_filename

BASE_DIR = os.path.abspath(os.path.dirname(__file__))
INSTANCE_DIR = os.path.join(BASE_DIR, "instance")
UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
DB_PATH = os.path.join(INSTANCE_DIR, "epps.db")

REQUIRED_COLUMNS = [
    "NRO. DNI",
    "NOMBRE AUXILIAR",
    "DESCRIPCION",
    "CANTIDAD",
    "FECHA DE MOVIMIENTO",
    "FECHA DE RENOVACION",
    "ESTADO",
]
OPTIONAL_COLUMNS = [
    "ESTADO MANUAL",
    "TIPO PLLA",
    "AREA",
    "CARGO",
    "OPERACIÓN",
    "OPERACION",
    "OPERACIÓN ",
]
ALLOWED_EXTENSIONS = {"xlsx", "xlsm", "xls"}

app = Flask(__name__)
app.config["SECRET_KEY"] = os.environ.get("SECRET_KEY", "transportes-libertad-epps-2026")
app.config["UPLOAD_FOLDER"] = UPLOAD_DIR
app.config["MAX_CONTENT_LENGTH"] = 25 * 1024 * 1024

os.makedirs(INSTANCE_DIR, exist_ok=True)
os.makedirs(UPLOAD_DIR, exist_ok=True)


def normalize_col(value: str) -> str:
    normalized = " ".join(str(value).replace("\n", " ").replace("\r", " ").strip().upper().split())
    aliases = {
        "FECHA MOVIMIENTO": "FECHA DE MOVIMIENTO",
        "OPERACION": "OPERACIÓN",
        "OPERACIÓN ": "OPERACIÓN",
    }
    return aliases.get(normalized, normalized)


def allowed_file(filename: str) -> bool:
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS


def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(error=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def init_db():
    db = sqlite3.connect(DB_PATH)
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS admin_user (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            created_at TEXT NOT NULL
        )
        """
    )
    db.execute(
        """
        CREATE TABLE IF NOT EXISTS epp_records (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            dni TEXT NOT NULL,
            nombre_auxiliar TEXT,
            descripcion TEXT,
            cantidad TEXT,
            fecha_movimiento TEXT,
            fecha_renovacion TEXT,
            estado TEXT,
            estado_manual TEXT,
            tipo_plla TEXT,
            area TEXT,
            cargo TEXT,
            operacion TEXT,
            archivo_origen TEXT,
            fecha_importacion TEXT NOT NULL,
            renovacion_sort TEXT,
            movimiento_sort TEXT,
            estado_visual TEXT
        )
        """
    )
    db.commit()

    username = os.environ.get("ADMIN_USERNAME", "admin")
    password = os.environ.get("ADMIN_PASSWORD", "Admin123*")
    exists = db.execute("SELECT id FROM admin_user WHERE username = ?", (username,)).fetchone()
    if not exists:
        db.execute(
            "INSERT INTO admin_user (username, password_hash, created_at) VALUES (?, ?, ?)",
            (username, generate_password_hash(password), datetime.now().strftime("%Y-%m-%d %H:%M:%S")),
        )
        db.commit()
    db.close()


def login_required(view):
    @wraps(view)
    def wrapped_view(*args, **kwargs):
        if not session.get("admin_logged_in"):
            flash("Debe iniciar sesión para acceder al panel administrador.", "warning")
            return redirect(url_for("admin_login"))
        return view(*args, **kwargs)

    return wrapped_view




def parse_date_value(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return pd.NaT
    text_value = str(value).strip()
    if not text_value or text_value.lower() == "nan":
        return pd.NaT
    parsed = pd.to_datetime(text_value, errors="coerce")
    if pd.isna(parsed):
        parsed = pd.to_datetime(text_value, dayfirst=True, errors="coerce")
    return parsed


def to_display_date(value):
    parsed = parse_date_value(value)
    if pd.isna(parsed):
        return "" if value is None else str(value).strip()
    return parsed.strftime("%d/%m/%Y")


def to_sort_date(value):
    parsed = parse_date_value(value)
    if pd.isna(parsed):
        return ""
    return parsed.strftime("%Y-%m-%d")


def detect_header_row(file_path: str, sheet_name: str) -> int:
    preview = pd.read_excel(file_path, sheet_name=sheet_name, header=None, nrows=12, dtype=str)
    required_norm = {normalize_col(col) for col in REQUIRED_COLUMNS}
    for idx, row in preview.iterrows():
        row_values = {normalize_col(v) for v in row.tolist() if str(v).strip() and str(v) != "nan"}
        if required_norm.issubset(row_values):
            return idx
    return 0


def read_excel_file(file_path: str) -> pd.DataFrame:
    xls = pd.ExcelFile(file_path)
    if not xls.sheet_names:
        raise ValueError("El archivo no contiene hojas.")

    sheet_name = xls.sheet_names[0]
    header_row = detect_header_row(file_path, sheet_name)
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row, dtype=str)
    df.columns = [normalize_col(c) for c in df.columns]
    df = df.loc[:, ~pd.Index(df.columns).duplicated()].copy()
    df = df.dropna(how="all")

    required_norm = [normalize_col(c) for c in REQUIRED_COLUMNS]
    missing = [col for col in required_norm if col not in df.columns]
    if missing:
        raise ValueError("Faltan columnas obligatorias: " + ", ".join(missing))

    keep_cols = list(dict.fromkeys(required_norm + [normalize_col(c) for c in OPTIONAL_COLUMNS if normalize_col(c) in df.columns]))
    df = df[keep_cols].copy().dropna(how="all")

    for col in df.columns:
        df[col] = df[col].fillna("").astype(str).str.strip()

    df = df[df["NRO. DNI"].str.replace(r"\D", "", regex=True) != ""]
    if df.empty:
        raise ValueError("No se encontraron registros válidos con DNI.")

    df["NRO. DNI"] = df["NRO. DNI"].str.replace(r"\D", "", regex=True).str[:8]

    mov_dt = pd.to_datetime(df["FECHA DE MOVIMIENTO"], errors="coerce")
    ren_dt = pd.to_datetime(df["FECHA DE RENOVACION"], errors="coerce")
    invalid_mov = mov_dt.isna()
    invalid_ren = ren_dt.isna()
    if invalid_mov.any():
        mov_dt.loc[invalid_mov] = pd.to_datetime(df.loc[invalid_mov, "FECHA DE MOVIMIENTO"], dayfirst=True, errors="coerce")
    if invalid_ren.any():
        ren_dt.loc[invalid_ren] = pd.to_datetime(df.loc[invalid_ren, "FECHA DE RENOVACION"], dayfirst=True, errors="coerce")

    df["FECHA DE MOVIMIENTO_DISPLAY"] = mov_dt.dt.strftime("%d/%m/%Y").fillna(df["FECHA DE MOVIMIENTO"])
    df["FECHA DE RENOVACION_DISPLAY"] = ren_dt.dt.strftime("%d/%m/%Y").fillna(df["FECHA DE RENOVACION"])
    df["FECHA DE MOVIMIENTO_SORT"] = mov_dt.dt.strftime("%Y-%m-%d").fillna("")
    df["FECHA DE RENOVACION_SORT"] = ren_dt.dt.strftime("%Y-%m-%d").fillna("")

    estado_manual = df["ESTADO MANUAL"].fillna("").astype(str).str.strip() if "ESTADO MANUAL" in df.columns else pd.Series("", index=df.index)
    estado = df["ESTADO"].fillna("").astype(str).str.strip()
    estado_visual = estado.copy()
    today = pd.Timestamp(date.today())
    delta_days = (ren_dt.dt.normalize() - today).dt.days
    estado_visual = estado_visual.mask(delta_days < 0, "VENCIDO")
    estado_visual = estado_visual.mask((delta_days >= 0) & (delta_days <= 30), "POR VENCER")
    estado_visual = estado_visual.mask(estado_visual.eq(""), "SIN ESTADO")
    estado_visual = estado_visual.mask(estado_manual.ne(""), estado_manual)
    df["ESTADO_VISUAL"] = estado_visual
    return df


def replace_records_from_dataframe(df: pd.DataFrame, original_filename: str):
    db = get_db()
    now = datetime.now().strftime("%d/%m/%Y %H:%M:%S")
    db.execute("DELETE FROM epp_records")

    for _, row in df.iterrows():
        operacion = row.get("OPERACIÓN", "") or row.get("OPERACION", "") or row.get("OPERACIÓN ", "") or ""
        db.execute(
            """
            INSERT INTO epp_records (
                dni, nombre_auxiliar, descripcion, cantidad, fecha_movimiento,
                fecha_renovacion, estado, estado_manual, tipo_plla, area,
                cargo, operacion, archivo_origen, fecha_importacion,
                renovacion_sort, movimiento_sort, estado_visual
            ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            """,
            (
                str(row.get("NRO. DNI", "")).strip(),
                str(row.get("NOMBRE AUXILIAR", "")).strip(),
                str(row.get("DESCRIPCION", "")).strip(),
                str(row.get("CANTIDAD", "")).strip(),
                str(row.get("FECHA DE MOVIMIENTO_DISPLAY", "")).strip(),
                str(row.get("FECHA DE RENOVACION_DISPLAY", "")).strip(),
                str(row.get("ESTADO", "")).strip(),
                str(row.get("ESTADO MANUAL", "")).strip(),
                str(row.get("TIPO PLLA", "")).strip(),
                str(row.get("AREA", "")).strip(),
                str(row.get("CARGO", "")).strip(),
                str(operacion).strip(),
                original_filename,
                now,
                str(row.get("FECHA DE RENOVACION_SORT", "")).strip(),
                str(row.get("FECHA DE MOVIMIENTO_SORT", "")).strip(),
                str(row.get("ESTADO_VISUAL", "")).strip(),
            ),
        )
    db.commit()


def get_dashboard_stats():
    db = get_db()
    base = db.execute(
        """
        SELECT COUNT(*) AS total_registros,
               COUNT(DISTINCT dni) AS total_trabajadores,
               COALESCE(MAX(fecha_importacion), 'Sin carga') AS ultima_importacion,
               COALESCE(MAX(archivo_origen), 'Sin archivo') AS archivo_origen
        FROM epp_records
        """
    ).fetchone()

    estados = db.execute(
        """
        SELECT
            SUM(CASE WHEN UPPER(estado_visual) = 'VENCIDO' THEN 1 ELSE 0 END) AS vencidos,
            SUM(CASE WHEN UPPER(estado_visual) = 'POR VENCER' THEN 1 ELSE 0 END) AS por_vencer,
            SUM(CASE WHEN UPPER(estado_visual) IN ('ENTREGADO', 'VIGENTE') THEN 1 ELSE 0 END) AS vigentes
        FROM epp_records
        """
    ).fetchone()

    result = dict(base)
    result.update({
        "vencidos": estados["vencidos"] or 0,
        "por_vencer": estados["por_vencer"] or 0,
        "vigentes": estados["vigentes"] or 0,
    })
    return result


@app.route("/", methods=["GET", "POST"])
def home():
    records = []
    dni = ""
    nombre = ""
    resumen = None

    if request.method == "POST":
        dni = "".join(ch for ch in request.form.get("dni", "") if ch.isdigit())
        if len(dni) != 8:
            flash("Ingrese un DNI válido de 8 dígitos.", "error")
        else:
            db = get_db()
            records = db.execute(
                """
                SELECT dni, nombre_auxiliar, descripcion, cantidad,
                       fecha_movimiento, fecha_renovacion, estado, estado_visual,
                       area, cargo, operacion
                FROM epp_records
                WHERE dni = ?
                ORDER BY renovacion_sort ASC, descripcion ASC
                """,
                (dni,),
            ).fetchall()
            if records:
                nombre = records[0]["nombre_auxiliar"]
                resumen = {
                    "total_items": len(records),
                    "vencidos": sum(1 for r in records if (r["estado_visual"] or "").upper() == "VENCIDO"),
                    "por_vencer": sum(1 for r in records if (r["estado_visual"] or "").upper() == "POR VENCER"),
                    "vigentes": sum(1 for r in records if (r["estado_visual"] or "").upper() in {"ENTREGADO", "VIGENTE"}),
                }
            else:
                flash("No se encontraron registros para ese DNI.", "warning")

    return render_template("home.html", records=records, dni=dni, nombre=nombre, resumen=resumen)


@app.route("/admin/login", methods=["GET", "POST"])
def admin_login():
    if request.method == "POST":
        username = request.form.get("username", "").strip()
        password = request.form.get("password", "")
        db = get_db()
        user = db.execute("SELECT * FROM admin_user WHERE username = ?", (username,)).fetchone()
        if user and check_password_hash(user["password_hash"], password):
            session.clear()
            session["admin_logged_in"] = True
            session["admin_username"] = username
            return redirect(url_for("admin_dashboard"))
        flash("Usuario o contraseña incorrectos.", "error")
    return render_template("admin_login.html")


@app.route("/admin/logout")
def admin_logout():
    session.clear()
    return redirect(url_for("admin_login"))


@app.route("/admin", methods=["GET", "POST"])
@login_required
def admin_dashboard():
    db = get_db()

    if request.method == "POST":
        excel = request.files.get("excel_file")
        if not excel or excel.filename == "":
            flash("Seleccione un archivo Excel.", "error")
            return redirect(url_for("admin_dashboard"))
        if not allowed_file(excel.filename):
            flash("Formato no permitido. Use .xlsx, .xlsm o .xls.", "error")
            return redirect(url_for("admin_dashboard"))

        filename = secure_filename(excel.filename)
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        final_name = f"{timestamp}_{filename}"
        save_path = os.path.join(app.config["UPLOAD_FOLDER"], final_name)
        excel.save(save_path)

        try:
            df = read_excel_file(save_path)
            replace_records_from_dataframe(df, final_name)
            flash(f"Archivo cargado correctamente. Registros importados: {len(df)}", "success")
        except Exception as exc:
            flash(f"No se pudo importar el archivo: {exc}", "error")
        return redirect(url_for("admin_dashboard"))

    filtro_dni = "".join(ch for ch in request.args.get("dni", "") if ch.isdigit())
    filtro_estado = request.args.get("estado", "").strip().upper()

    query = """
        SELECT dni, nombre_auxiliar, descripcion, fecha_renovacion, estado_visual, archivo_origen
        FROM epp_records
        WHERE 1=1
    """
    params = []
    if filtro_dni:
        query += " AND dni = ?"
        params.append(filtro_dni)
    if filtro_estado:
        query += " AND UPPER(estado_visual) = ?"
        params.append(filtro_estado)
    query += " ORDER BY renovacion_sort ASC, id DESC LIMIT 25"

    recent_records = db.execute(query, params).fetchall()
    stats = get_dashboard_stats()
    return render_template(
        "admin_dashboard.html",
        stats=stats,
        recent_records=recent_records,
        filtro_dni=filtro_dni,
        filtro_estado=filtro_estado,
    )


@app.route("/admin/change-password", methods=["POST"])
@login_required
def change_password():
    current_password = request.form.get("current_password", "")
    new_password = request.form.get("new_password", "")
    confirm_password = request.form.get("confirm_password", "")

    if len(new_password) < 8:
        flash("La nueva contraseña debe tener al menos 8 caracteres.", "error")
        return redirect(url_for("admin_dashboard"))
    if new_password != confirm_password:
        flash("La confirmación de contraseña no coincide.", "error")
        return redirect(url_for("admin_dashboard"))

    db = get_db()
    user = db.execute("SELECT * FROM admin_user WHERE username = ?", (session["admin_username"],)).fetchone()
    if not user or not check_password_hash(user["password_hash"], current_password):
        flash("La contraseña actual es incorrecta.", "error")
        return redirect(url_for("admin_dashboard"))

    db.execute("UPDATE admin_user SET password_hash = ? WHERE id = ?", (generate_password_hash(new_password), user["id"]))
    db.commit()
    flash("Contraseña actualizada correctamente.", "success")
    return redirect(url_for("admin_dashboard"))


@app.route("/admin/exportar")
@login_required
def admin_exportar():
    db = get_db()
    df = pd.read_sql_query(
        """
        SELECT dni AS 'NRO. DNI', nombre_auxiliar AS 'NOMBRE AUXILIAR', descripcion AS 'DESCRIPCION',
               cantidad AS 'CANTIDAD', fecha_movimiento AS 'FECHA DE MOVIMIENTO',
               fecha_renovacion AS 'FECHA DE RENOVACION', estado AS 'ESTADO',
               estado_manual AS 'ESTADO MANUAL', tipo_plla AS 'TIPO PLLA', area AS 'AREA',
               cargo AS 'CARGO', operacion AS 'OPERACIÓN', estado_visual AS 'ESTADO VISUAL',
               archivo_origen AS 'ARCHIVO ORIGEN', fecha_importacion AS 'FECHA IMPORTACION'
        FROM epp_records
        ORDER BY dni, renovacion_sort, descripcion
        """,
        db,
    )

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="EPPS")
    output.seek(0)
    filename = f"reporte_epps_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    return send_file(output, as_attachment=True, download_name=filename, mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


@app.template_filter("estado_badge")
def estado_badge(value):
    value = (value or "").strip().upper()
    if value == "VENCIDO":
        return "badge-red"
    if value == "POR VENCER":
        return "badge-yellow"
    if value in {"ENTREGADO", "VIGENTE"}:
        return "badge-green"
    return "badge-gray"


@app.context_processor
def inject_now():
    return {"current_year": datetime.now().year}


if __name__ == "__main__":
    init_db()
    app.run(debug=True)
