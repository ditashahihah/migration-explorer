"""
Logika data murni (framework-agnostic) buat Migration Progress Explorer -
dipakai bareng oleh app.py (Streamlit) dan dash_app.py (Dash), supaya kedua
UI itu selalu konsisten hasilnya (satu sumber kebenaran buat aturan bisnis).

Modul ini SENGAJA tidak import streamlit atau dash - cuma pandas/requests/
stdlib, jadi aman dipakai dari environment manapun.
"""

from __future__ import annotations

import os
import shutil

import pandas as pd
import requests

# ---------------------------------------------------------------------
# Konfigurasi & konstanta
# ---------------------------------------------------------------------
SHEET_NAME = "Coretan Checking Migration"
DEFAULT_PATH = "Checking_Progress_Migration.xlsx"

# Pemetaan nama kolom asli (di Excel) -> nama kolom yang dipakai di app ini.
# Kalau nanti header di Excel berubah, cukup update dictionary ini.
COLUMN_MAP = {
    "Project": "Project",
    "Data Source (Tables)": "Data Source",
    "Phase 1": "Phase 1",
    "Column DWH": "Column DWH",
    "Source Table in DataLake": "Bronze Table",
    "Status Column Source": "Status",
    "Source Column In datalake": "Bronze Column",
    "Stage to silver 2": "Stage to Silver",
    "Table Silver Tier 1 In Datalake": "Silver1 Table",
    "Column Silver Tier 1 In Datalake": "Silver1 Column",
    "Table Silver Tier 2 in datalake": "Silver2 Table",
    "Silver Column Datalake Tier 2": "Silver2 Column",
    "Domain": "Domain",
    "Kategori Project": "Kategori Project",
    "Formula Hardcode": "Formula Hardcode",
}

# Urutan & kolom yang ditampilkan di tabel detail (biar konsisten DWH -> Bronze
# -> Silver1 -> Silver2 dari kiri ke kanan).
DISPLAY_COLUMNS = [
    "Project",
    "Short Table",
    "Column DWH",
    "Status",
    "Bronze Table",
    "Bronze Column",
    "Silver1 Table",
    "Silver1 Column",
    "Silver2 Table",
    "Silver2 Column",
    "Domain",
    "Kategori Project",
    "Phase 1",
    "Formula Hardcode",
]

STATUS_COLORS = {
    "Gap": "#f8696b",
    "Hardcode": "#ffeb84",
    "Table": "#63be7b",
}

SHEET_SELECTED = "Selected Columns"
SHEET_UNIQUE = "Unique Columns per Table"
SHEET_COMPILE = "Compile per Table"
BACKUP_PATH = DEFAULT_PATH + ".bak"

# kolom yang disimpan di sheet "Selected Columns" (sama seperti DISPLAY_COLUMNS
# + recipe Dataiku mana aja yang confirmed makai kolom ini + jejak waktu submit)
SELECTED_SHEET_COLUMNS = DISPLAY_COLUMNS + ["Dipakai di Recipe", "Selected At"]


def get_secret(name: str, default=None):
    """Baca secret dari environment variable - portable ke semua host
    (Docker/Hugging Face Spaces/Render/dst pakai env var, bukan file secrets)."""
    if name in os.environ:
        return os.environ[name]
    return os.environ.get(name.upper(), default)


def has_secret(name: str) -> bool:
    return get_secret(name) is not None


def gsheet_enabled() -> bool:
    """False kalau env var GSHEET_WEBAPP_URL/GSHEET_WEBAPP_TOKEN belum diisi."""
    return has_secret("gsheet_webapp_url") and has_secret("gsheet_webapp_token")


# ---------------------------------------------------------------------
# Helper: ekstrak nama table pendek dari "Data Source (Tables)"
# Contoh: "[DB_ANALYTICS].[AMFS_DWH].[dbo].[IB_BRANCH]" -> "IB_BRANCH"
# ---------------------------------------------------------------------
def extract_short_table(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    last = s.split(".")[-1].strip().strip("[]").strip()
    return last.upper() if last else None


def load_data(file_source) -> pd.DataFrame:
    """Baca sheet 'Coretan Checking Migration' dan siapkan kolom turunan."""
    df = pd.read_excel(file_source, sheet_name=SHEET_NAME, engine="openpyxl")
    df = df.rename(columns=COLUMN_MAP)

    keep_cols = [c for c in COLUMN_MAP.values() if c in df.columns]
    df = df[keep_cols].copy()

    df["Short Table"] = df["Data Source"].apply(extract_short_table)

    # buang baris yang benar-benar kosong (tidak ada project & table sama sekali)
    df = df.dropna(subset=["Project", "Short Table"], how="all").reset_index(drop=True)
    return df


def pipeline_stage_counts(df_subset: pd.DataFrame) -> dict:
    """Hitung berapa kolom DWH unik di df_subset yang sampai ke tiap stage
    pipeline (Bronze/Silver) dan Hardcode/Gap. Di-dedup dulu by (Short Table,
    Column DWH) supaya kolom yang dipakai banyak project tidak dobel-hitung,
    dan kolom senama di table berbeda tetap dihitung terpisah. "Silver"
    dihitung dari Silver Tier 2 saja (stage akhir)."""
    unique = df_subset.dropna(subset=["Column DWH"]).drop_duplicates(subset=["Short Table", "Column DWH"])
    return {
        "unik": len(unique),
        "bronze": int((unique["Status"] == "Table").sum()) if "Status" in unique else 0,
        "silver": int(unique["Silver2 Table"].notna().sum()) if "Silver2 Table" in unique else 0,
        "hardcode": int((unique["Status"] == "Hardcode").sum()) if "Status" in unique else 0,
        "gap": int((unique["Status"] == "Gap").sum()) if "Status" in unique else 0,
    }


def norm_col(s) -> str:
    """Normalisasi nama kolom buat dibandingkan - Coretan sering nyimpen
    Column DWH dengan bracket SQL Server (mis. '[OWN_MOBILE_PH]'), sedangkan
    schema di dokumen Dataiku polos tanpa bracket. Tanpa normalisasi ini,
    kolom yang sebenarnya sama bakal keliatan cuma-di-DWH DAN cuma-di-
    Dokumen secara terpisah."""
    return str(s).strip().strip("[]").strip().upper()


def keterangan_label(in_dwh: bool, in_used: bool) -> str:
    if in_dwh and in_used:
        return "✅ Ada di DWH & dipakai di Dataiku"
    if in_dwh:
        return "📗 Ada di DWH, TIDAK dipakai di Dataiku"
    return "📄 Dipakai di Dataiku, TIDAK ada di DWH"


def build_table_confirmed_map(doc_recipes: list) -> dict:
    """Ubah list Recipe (dataclass, hasil parse_dataiku_doc/json) jadi dict
    polos {table: {column: [nama_recipe, ...]}} - JSON-serializable (bisa
    disimpan di dcc.Store buat Dash), dan jadi input buat build_table_union()
    di bawah supaya kedua UI (Streamlit & Dash) pakai fungsi union yang sama
    tanpa perlu passing objek Recipe langsung."""
    result: dict[str, dict[str, list]] = {}
    for r in doc_recipes:
        for table, cols in r.confirmed_columns.items():
            bucket = result.setdefault(table, {})
            for c in cols:
                bucket.setdefault(c, []).append(r.name)
    return result


def build_table_union(proj_df: pd.DataFrame, table_confirmed_map: dict, table: str) -> pd.DataFrame:
    """Bangun tabel union DWH + kolom confirmed dipakai di Dataiku buat 1
    table, dipakai bareng oleh mode 'Info Table by Project' di Streamlit
    maupun Dash. Cuma nampilin kolom yang confirmed dipakai di Dataiku
    (union dwh_norm ada di situ HANYA kalau juga confirmed dipakai).

    `table_confirmed_map`: hasil build_table_confirmed_map() di atas."""
    dwh_rows = (
        proj_df[proj_df["Short Table"] == table][["Column DWH", "Status"]]
        .dropna(subset=["Column DWH"])
        .drop_duplicates()
    )
    dwh_status = dict(zip(dwh_rows["Column DWH"], dwh_rows["Status"]))
    dwh_norm = {norm_col(c): c for c in dwh_status}

    used_norm: dict[str, str] = {}
    used_recipes: dict[str, list] = {}
    for c, recipe_names in table_confirmed_map.get(table, {}).items():
        n = norm_col(c)
        used_norm.setdefault(n, c)
        used_recipes.setdefault(n, []).extend(recipe_names)

    all_norm = sorted(set(used_norm))
    rows = []
    for n in all_norm:
        in_dwh, in_used = n in dwh_norm, n in used_norm
        label = dwh_norm[n] if in_dwh else used_norm[n]
        rows.append(
            {
                "Pilih": n in dwh_norm and n in used_norm,
                "Column DWH": label,
                "Status": dwh_status.get(label, "-") if in_dwh else "-",
                "Keterangan": keterangan_label(in_dwh, in_used),
                "Dipakai di Recipe": ", ".join(sorted(set(used_recipes.get(n, [])))) or "-",
            }
        )
    return pd.DataFrame(rows, columns=["Pilih", "Column DWH", "Status", "Keterangan", "Dipakai di Recipe"])


# ---------------------------------------------------------------------
# Seleksi Kolom: backend Google Sheets lewat Apps Script Web App
# ---------------------------------------------------------------------
def _appscript_call(action: str, **payload) -> dict:
    """POST ke Apps Script Web App. Apps Script selalu balas HTTP 200 (tidak
    bisa set status code custom), jadi sukses/gagal dicek dari field `ok`
    di body JSON-nya, bukan dari status code."""
    url = get_secret("gsheet_webapp_url")
    token = get_secret("gsheet_webapp_token")
    resp = requests.post(url, json={"token": token, "action": action, **payload}, timeout=20)
    resp.raise_for_status()
    data = resp.json()
    if not data.get("ok"):
        raise RuntimeError(data.get("error", "Apps Script mengembalikan error tanpa detail"))
    return data


def _df_to_appscript_payload(df: pd.DataFrame) -> dict:
    header = list(df.columns)
    body = df.astype(object).where(pd.notna(df), "").values.tolist() if not df.empty else []
    return {"headers": header, "rows": body}


def load_selected_sheet_gsheet() -> pd.DataFrame:
    data = _appscript_call("read", sheet=SHEET_SELECTED)
    rows = data.get("rows", [])
    return pd.DataFrame(rows) if rows else pd.DataFrame(columns=SELECTED_SHEET_COLUMNS)


def save_project_selection_gsheet(project: str, new_rows: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Sama seperti save_project_selection_local, tapi 3 sheet-nya ditulis ke
    tab Google Sheets lewat Apps Script Web App, bukan ke file Excel lokal —
    supaya persisten di hosting gratis yang storage lokalnya sementara."""
    existing = load_selected_sheet_gsheet()
    existing = existing[existing["Project"] != project] if not existing.empty else existing

    new_rows = new_rows.copy()
    new_rows["Selected At"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    new_rows = new_rows[[c for c in SELECTED_SHEET_COLUMNS if c in new_rows.columns]]

    combined = pd.concat([existing, new_rows], ignore_index=True)
    unique_df = build_unique_sheet(combined)
    compile_df = build_compile_sheet(unique_df)

    _appscript_call("write", sheet=SHEET_SELECTED, **_df_to_appscript_payload(combined))
    _appscript_call("write", sheet=SHEET_UNIQUE, **_df_to_appscript_payload(unique_df))
    _appscript_call("write", sheet=SHEET_COMPILE, **_df_to_appscript_payload(compile_df))

    return combined, unique_df, compile_df


# ---------------------------------------------------------------------
# Seleksi Kolom: backend file Excel lokal (fallback waktu tidak ada
# kredensial Google Sheets, misal jalan lokal di laptop)
# ---------------------------------------------------------------------
def load_selected_sheet_local(path: str) -> pd.DataFrame:
    """Baca sheet 'Selected Columns' dari file output kalau sudah ada."""
    try:
        return pd.read_excel(path, sheet_name=SHEET_SELECTED, engine="openpyxl")
    except (FileNotFoundError, ValueError):
        return pd.DataFrame(columns=SELECTED_SHEET_COLUMNS)


def build_unique_sheet(selected_df: pd.DataFrame) -> pd.DataFrame:
    """Union unik (Short Table, Column DWH) dari seluruh project yang pernah diseleksi."""
    key_cols = ["Short Table", "Column DWH"]
    detail_cols = [
        c
        for c in [
            "Status",
            "Bronze Table",
            "Bronze Column",
            "Silver1 Table",
            "Silver1 Column",
            "Silver2 Table",
            "Silver2 Column",
            "Domain",
        ]
        if c in selected_df.columns
    ]
    unique_df = selected_df.drop_duplicates(subset=key_cols)[key_cols + detail_cols].copy()

    projects_by_key = (
        selected_df.groupby(key_cols)["Project"]
        .apply(lambda s: ", ".join(sorted(set(s.dropna()))))
        .reset_index()
        .rename(columns={"Project": "Projects"})
    )
    unique_df = unique_df.merge(projects_by_key, on=key_cols, how="left")
    return unique_df.sort_values(key_cols).reset_index(drop=True)


def build_compile_sheet(unique_df: pd.DataFrame) -> pd.DataFrame:
    """Compile hitungan per table: dari DWH berapa yang sampai Bronze/Silver/Hardcode/Gap."""
    if unique_df.empty:
        return pd.DataFrame(
            columns=[
                "Short Table",
                "Jumlah Kolom DWH",
                "Jumlah ke Bronze",
                "Jumlah ke Silver 1",
                "Jumlah ke Silver 2",
                "Jumlah Hardcode",
                "Jumlah Gap",
            ]
        )

    grouped = unique_df.groupby("Short Table").agg(
        **{
            "Jumlah Kolom DWH": ("Column DWH", "count"),
            "Jumlah ke Bronze": ("Status", lambda s: (s == "Table").sum()),
            "Jumlah Hardcode": ("Status", lambda s: (s == "Hardcode").sum()),
            "Jumlah Gap": ("Status", lambda s: (s == "Gap").sum()),
        }
    )
    if "Silver1 Table" in unique_df.columns:
        grouped["Jumlah ke Silver 1"] = unique_df.groupby("Short Table")["Silver1 Table"].apply(
            lambda s: s.notna().sum()
        )
    if "Silver2 Table" in unique_df.columns:
        grouped["Jumlah ke Silver 2"] = unique_df.groupby("Short Table")["Silver2 Table"].apply(
            lambda s: s.notna().sum()
        )

    grouped = grouped.reset_index()
    col_order = [
        "Short Table",
        "Jumlah Kolom DWH",
        "Jumlah ke Bronze",
        "Jumlah ke Silver 1",
        "Jumlah ke Silver 2",
        "Jumlah Hardcode",
        "Jumlah Gap",
    ]
    return grouped[[c for c in col_order if c in grouped.columns]]


def save_project_selection_local(path: str, project: str, new_rows: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    """Ganti seleksi lama project ini dengan yang baru, lalu tulis ulang 3 sheet
    seleksi ke dalam `path` (file master) tanpa mengubah sheet-sheet lain."""
    existing = load_selected_sheet_local(path)
    existing = existing[existing["Project"] != project] if not existing.empty else existing

    new_rows = new_rows.copy()
    new_rows["Selected At"] = pd.Timestamp.now().strftime("%Y-%m-%d %H:%M:%S")
    new_rows = new_rows[[c for c in SELECTED_SHEET_COLUMNS if c in new_rows.columns]]

    combined = pd.concat([existing, new_rows], ignore_index=True)
    unique_df = build_unique_sheet(combined)
    compile_df = build_compile_sheet(unique_df)

    # backup 1-langkah-mundur sebelum nulis, jaga-jaga kalau proses tulis gagal
    shutil.copy2(path, path + ".bak")

    with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        combined.to_excel(writer, sheet_name=SHEET_SELECTED, index=False)
        unique_df.to_excel(writer, sheet_name=SHEET_UNIQUE, index=False)
        compile_df.to_excel(writer, sheet_name=SHEET_COMPILE, index=False)

    return combined, unique_df, compile_df


# ---------------------------------------------------------------------
# Seleksi Kolom: dispatcher — pilih backend Google Sheets kalau ada
# kredensialnya (deploy di hosting gratis), fallback ke file Excel lokal
# kalau tidak (jalan lokal di laptop).
# ---------------------------------------------------------------------
def save_project_selection(project: str, new_rows: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame, pd.DataFrame]:
    if gsheet_enabled():
        return save_project_selection_gsheet(project, new_rows)
    return save_project_selection_local(DEFAULT_PATH, project, new_rows)
