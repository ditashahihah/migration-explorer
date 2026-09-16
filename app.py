"""
Migration Progress Explorer
============================
Streamlit app untuk cari progress migrasi data (DWH -> Bronze -> Silver Tier 1
-> Silver Tier 2) dengan dua mode:
  1. Cari by Table  -> lihat project mana saja yang pakai table itu, plus
     detail tiap kolomnya di semua stage.
  2. Cari by Project -> lihat table & kolom apa saja yang dipakai project itu,
     plus detail tiap kolomnya di semua stage.

Sumber data: sheet 'Coretan Checking Migration' di file
Checking_Progress_Migration.xlsx. Baca README.md untuk penjelasan lengkap
struktur data & kolomnya.
"""

import io
import json
import os
import shutil

import pandas as pd
import requests
import streamlit as st

from dataiku_doc import build_project_column_report, parse_dataiku_doc


def get_secret(name: str, default=None):
    """Baca secret dari st.secrets (Streamlit Cloud/lokal, .streamlit/secrets.toml)
    kalau ada, fallback ke environment variable (Hugging Face Spaces & host lain
    yang nyimpen secrets sebagai env var, bukan secrets.toml). Ini bikin app-nya
    portable ke berbagai platform hosting tanpa ubah kode."""
    try:
        if name in st.secrets:
            return st.secrets[name]
    except Exception:
        pass
    # env var lookup: coba nama persis dulu, lalu versi UPPERCASE (konvensi
    # umum di host lain kayak Hugging Face Spaces) - env var case-sensitive
    # di Linux, beda dari Windows yang case-insensitive.
    if name in os.environ:
        return os.environ[name]
    return os.environ.get(name.upper(), default)


def has_secret(name: str) -> bool:
    return get_secret(name) is not None


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
    "Gap": "background-color: #f8696b; color: white;",
    "Hardcode": "background-color: #ffeb84;",
    "Table": "background-color: #63be7b;",
}

# ---------------------------------------------------------------------
# Seleksi Kolom: 2 backend penyimpanan.
#
# 1. Google Sheets lewat Google Apps Script Web App (kalau secrets
#    `gsheet_webapp_url` + `gsheet_webapp_token` ada) — dipakai kalau app
#    di-deploy ke hosting gratis (storage-nya sementara, jadi hasil seleksi
#    PIC harus disimpan di luar container biar tidak hilang tiap restart).
#    Dipilih lewat Apps Script (bukan service account Google Cloud) karena
#    Apps Script cuma butuh akun Google biasa — tidak perlu bikin project
#    Google Cloud / kartu kredit sama sekali. Lihat README bagian 7.
# 2. File lokal `Checking_Progress_Migration.xlsx` (fallback kalau secrets
#    di atas tidak ada) — dipakai waktu jalan lokal di laptop. Ditulis balik
#    sebagai 3 sheet tambahan (sheet Coretan & sheet lain tidak disentuh),
#    dengan backup 1-langkah-mundur tiap sebelum nulis.
# ---------------------------------------------------------------------
SHEET_SELECTED = "Selected Columns"
SHEET_UNIQUE = "Unique Columns per Table"
SHEET_COMPILE = "Compile per Table"
BACKUP_PATH = DEFAULT_PATH + ".bak"

# kolom yang disimpan di sheet "Selected Columns" (sama seperti DISPLAY_COLUMNS
# + jejak waktu submit-nya)
SELECTED_SHEET_COLUMNS = DISPLAY_COLUMNS + ["Selected At"]


# ---------------------------------------------------------------------
# Helper: ekstrak nama table pendek dari "Data Source (Tables)"
# Contoh: "[DB_ANALYTICS].[AMFS_DWH].[dbo].[IB_BRANCH]" -> "IB_BRANCH"
# Ini rule yang sama persis dipakai waktu regenerasi sheet
# "List Table Compile - DWH", supaya hasilnya konsisten.
# ---------------------------------------------------------------------
def extract_short_table(value):
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return None
    s = str(value).strip()
    last = s.split(".")[-1].strip().strip("[]").strip()
    return last.upper() if last else None


@st.cache_data(show_spinner="Membaca dokumen Dataiku Flow (dokumen besar bisa makan waktu ~10-20 detik)...")
def parse_uploaded_doc(file_bytes: bytes):
    """Wrapper cached di atas parse_dataiku_doc supaya dokumen besar tidak
    di-parse ulang tiap kali ada interaksi UI lain (checkbox, dsb)."""
    return parse_dataiku_doc(io.BytesIO(file_bytes))


@st.cache_data(show_spinner="Memuat data dari Excel...")
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


def style_status(val):
    return STATUS_COLORS.get(val, "")


def render_detail_table(df_subset: pd.DataFrame):
    cols = [c for c in DISPLAY_COLUMNS if c in df_subset.columns]
    view = df_subset[cols]
    if "Status" in view.columns:
        try:
            styled = view.style.map(style_status, subset=["Status"])
        except AttributeError:
            # pandas lama belum punya Styler.map, fallback ke applymap
            styled = view.style.applymap(style_status, subset=["Status"])
        st.dataframe(styled, use_container_width=True, hide_index=True)
    else:
        st.dataframe(view, use_container_width=True, hide_index=True)


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


# ---------------------------------------------------------------------
# Seleksi Kolom: saran AI (Groq) buat kolom yang tidak bisa dipastikan dari
# dokumen (kena recipe "Prepare" tapi nama kolomnya tidak tercatat).
# Groq CUMA BOLEH MILIH dari daftar kolom yang dikasih (closed-set) - tidak
# boleh mengarang nama kolom baru, supaya tidak ada hasil yang halusinasi.
# ---------------------------------------------------------------------
GROQ_API_URL = "https://api.groq.com/openai/v1/chat/completions"
GROQ_MODEL = "openai/gpt-oss-120b"


def groq_enabled() -> bool:
    return has_secret("groq_api_key")


def ask_groq_column_suggestions(table_name: str, candidate_columns: list, context_note: str) -> tuple[set, str]:
    """Minta Groq milih kolom mana (dari candidate_columns SAJA) yang
    kemungkinan masih dipakai. Return (set_kolom_disarankan, alasan_singkat)."""
    if not candidate_columns:
        return set(), ""

    prompt = (
        f"Kamu membantu menganalisis pipeline data Dataiku untuk table '{table_name}'.\n"
        f"Konteks: {context_note}\n\n"
        "Daftar kolom yang TERSEDIA di table ini (belum ada bukti eksplisit "
        "dipakai atau tidak di dokumen flow-nya):\n"
        + "\n".join(f"- {c}" for c in candidate_columns)
        + "\n\nDari daftar itu SAJA (jangan sebutkan nama kolom lain yang "
        "tidak ada di daftar), kolom mana yang menurutmu paling mungkin "
        "masih dipakai/relevan di pipeline ini berdasarkan nama & "
        'konteksnya? Balas HANYA dalam format JSON: '
        '{"likely_used": ["nama_kolom", ...], "reasoning": "alasan singkat 1-2 kalimat"}'
    )
    resp = requests.post(
        GROQ_API_URL,
        headers={"Authorization": f"Bearer {get_secret('groq_api_key')}"},
        json={
            "model": get_secret("groq_model", GROQ_MODEL),
            "messages": [{"role": "user", "content": prompt}],
            "response_format": {"type": "json_object"},
            "temperature": 0,
        },
        timeout=30,
    )
    resp.raise_for_status()
    content = resp.json()["choices"][0]["message"]["content"]
    parsed = json.loads(content)

    # jaga2: cuma terima nama yang beneran ada di daftar candidate, buang
    # kalau Groq ngarang nama yang tidak diminta
    candidate_set = set(candidate_columns)
    likely = {c for c in parsed.get("likely_used", []) if c in candidate_set}
    reasoning = str(parsed.get("reasoning", "")).strip()
    return likely, reasoning


# ---------------------------------------------------------------------
# Seleksi Kolom: backend Google Sheets lewat Apps Script Web App
# ---------------------------------------------------------------------
def gsheet_enabled() -> bool:
    """False kalau secret gsheet_webapp_url/gsheet_webapp_token belum diisi,
    baik lewat secrets.toml (Streamlit Cloud/lokal) maupun environment
    variable (host lain kayak Hugging Face Spaces)."""
    return has_secret("gsheet_webapp_url") and has_secret("gsheet_webapp_token")


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


# ---------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------
st.set_page_config(page_title="Migration Progress Explorer", layout="wide")
st.title("🔍 Migration Progress Explorer")
st.caption(
    "Sumber data: sheet **'Coretan Checking Migration'** di "
    "`Checking_Progress_Migration.xlsx` — DWH → Bronze → Silver Tier 1 → Silver Tier 2"
)

# ---- Sidebar: sumber data ----
with st.sidebar:
    st.header("📂 Sumber Data")
    st.caption(f"Terkunci ke file default: `{DEFAULT_PATH}`")
    if st.button("🔄 Reload data (clear cache)"):
        st.cache_data.clear()

source = DEFAULT_PATH

try:
    df = load_data(source)
except FileNotFoundError:
    st.error(
        f"File '{DEFAULT_PATH}' tidak ditemukan di folder yang sama dengan app.py. "
        "Upload manual lewat sidebar, atau sesuaikan DEFAULT_PATH di app.py."
    )
    st.stop()
except ValueError as e:
    st.error(
        f"Gagal baca sheet '{SHEET_NAME}'. Pastikan nama sheet di Excel belum berubah. "
        f"Detail error: {e}"
    )
    st.stop()
except Exception as e:  # noqa: BLE001 - tampilkan apapun errornya ke user
    st.error(f"Gagal membaca file: {e}")
    st.stop()

st.sidebar.success(f"Data termuat: {len(df):,} baris")

# ---- Sidebar: filter global ----
st.sidebar.header("🔧 Filter")
kategori_opts = sorted(df["Kategori Project"].dropna().unique().tolist()) if "Kategori Project" in df else []
status_opts = sorted(df["Status"].dropna().unique().tolist()) if "Status" in df else []

kategori_sel = st.sidebar.multiselect("Kategori Project (DA/BI/DS/DQ)", kategori_opts, default=kategori_opts)
status_sel = st.sidebar.multiselect("Status Column Source", status_opts, default=status_opts)

df_f = df.copy()
if kategori_opts:
    df_f = df_f[df_f["Kategori Project"].isin(kategori_sel) | df_f["Kategori Project"].isna()]
if status_opts:
    df_f = df_f[df_f["Status"].isin(status_sel) | df_f["Status"].isna()]

# ---- Mode pencarian ----
mode = st.radio(
    "Mode pencarian",
    ["🔎 Cari by Table", "🔎 Cari by Project", "🧩 Seleksi Kolom"],
    horizontal=True,
)
st.divider()

if mode == "🔎 Cari by Table":
    tables = sorted(t for t in df_f["Short Table"].dropna().unique())
    keyword = st.text_input("Ketik untuk filter nama table (contoh: IB_POLICY)")
    if keyword:
        tables = [t for t in tables if keyword.strip().upper() in t]

    if not tables:
        st.warning("Tidak ada table yang cocok dengan filter saat ini.")
        st.stop()

    selected_table = st.selectbox(f"Pilih table ({len(tables)} ditemukan)", tables)
    subset = df_f[df_f["Short Table"] == selected_table]
    projects_using = sorted(subset["Project"].dropna().unique())
    stats = pipeline_stage_counts(subset)

    c1, c2, c3 = st.columns(3)
    c1.metric("Jumlah Project Pemakai", len(projects_using))
    c2.metric("Jumlah Baris Kolom", len(subset))
    c3.metric("Jumlah Kolom DWH Unik", stats["unik"])

    d1, d2, d3, d4 = st.columns(4)
    d1.metric("🟩 ke Bronze", stats["bronze"])
    d2.metric("ke Silver", stats["silver"])
    d3.metric("🟨 Hardcode", stats["hardcode"])
    d4.metric("🟥 Gap", stats["gap"])

    st.subheader("📋 Project yang memakai table ini")
    st.write(", ".join(projects_using) if projects_using else "-")

    st.subheader("📊 Detail Kolom: DWH → Bronze → Silver Tier 1 → Silver Tier 2")
    render_detail_table(subset.sort_values(["Project", "Column DWH"]))

elif mode == "🔎 Cari by Project":
    projects = sorted(p for p in df_f["Project"].dropna().unique())
    keyword = st.text_input("Ketik untuk filter nama project")
    if keyword:
        projects = [p for p in projects if keyword.strip().lower() in p.lower()]

    if not projects:
        st.warning("Tidak ada project yang cocok dengan filter saat ini.")
        st.stop()

    selected_project = st.selectbox(f"Pilih project ({len(projects)} ditemukan)", projects)
    subset = df_f[df_f["Project"] == selected_project]
    tables_used = sorted(subset["Short Table"].dropna().unique())
    stats = pipeline_stage_counts(subset)

    c1, c2, c3 = st.columns(3)
    c1.metric("Jumlah Table Dipakai", len(tables_used))
    c2.metric("Jumlah Baris Kolom", len(subset))
    c3.metric("Jumlah Kolom DWH Unik", stats["unik"])

    d1, d2, d3, d4 = st.columns(4)
    d1.metric("🟩 ke Bronze", stats["bronze"])
    d2.metric("ke Silver", stats["silver"])
    d3.metric("🟨 Hardcode", stats["hardcode"])
    d4.metric("🟥 Gap", stats["gap"])

    st.subheader("📋 Table yang dipakai project ini")
    st.write(", ".join(tables_used) if tables_used else "-")

    st.subheader("📊 Detail Kolom: DWH → Bronze → Silver Tier 1 → Silver Tier 2")
    render_detail_table(subset.sort_values(["Short Table", "Column DWH"]))

else:  # 🧩 Seleksi Kolom
    if gsheet_enabled():
        st.caption(
            "Pilih project → pilih table yang relevan → centang kolom DWH yang "
            "bener-bener dipakai project ini. Hasilnya disimpan ke **Google "
            "Sheets** (3 tab: Selected Columns, Unique Columns per Table, "
            "Compile per Table) — jadi tetap persisten meskipun hosting "
            "di-restart. Submit ulang untuk project yang sama akan "
            "menggantikan seleksi lama project itu, bukan menumpuk duplikat."
        )
    else:
        st.caption(
            "Pilih project → pilih table yang relevan → centang kolom DWH yang "
            "bener-bener dipakai project ini. Hasilnya ditulis balik sebagai 3 "
            f"sheet tambahan di `{DEFAULT_PATH}` sendiri (Selected Columns, Unique "
            "Columns per Table, Compile per Table) — sheet Coretan & sheet lain "
            "tidak disentuh. Submit ulang untuk project yang sama akan "
            "menggantikan seleksi lama project itu, bukan menumpuk duplikat. "
            f"Sebelum tiap kali nulis, file di-backup dulu ke `{BACKUP_PATH}`. "
            "(Belum ada kredensial Google Sheets — kalau di-deploy ke hosting "
            "gratis, hasil seleksi ini bisa hilang tiap container restart; "
            "lihat README bagian setup Google Sheets.)"
        )

    projects = sorted(p for p in df_f["Project"].dropna().unique())
    sel_project = st.selectbox("Pilih Project", projects, key="selkol_project")

    proj_df = df_f[df_f["Project"] == sel_project]
    tables_for_project = sorted(proj_df["Short Table"].dropna().unique())

    with st.expander("📄 Bantu pre-fill dari Dokumen Dataiku Flow (opsional)"):
        st.caption(
            "Upload dokumen 'Dataiku Flow Documentation' (.docx) buat project ini. "
            "App baca table & kolom apa saja yang benar-benar dipakai di flow-nya "
            "(dari join key / group key / distinct key yang tertulis eksplisit di "
            "dokumen), lalu otomatis pre-fill centang di bawah — kolom yang kena "
            "recipe 'Prepare' (transformasi tanpa nama kolom tercatat) dibiarkan "
            "TIDAK tercentang karena dokumennya sendiri tidak menyebut nama "
            "kolomnya, jadi tetap perlu direview manual."
        )
        doc_file = st.file_uploader(
            "Upload Dataiku Flow Documentation (.docx)", type=["docx"], key=f"selkol_doc_{sel_project}"
        )

    doc_report = {}
    doc_datasets = {}
    groq_suggestions = st.session_state.get(f"groq_suggestions_{sel_project}", {})
    if doc_file is not None:
        doc_datasets, doc_recipes = parse_uploaded_doc(doc_file.getvalue())
        coretan_short_tables = set(df["Short Table"].dropna().unique())
        doc_report = build_project_column_report(doc_datasets, doc_recipes, coretan_short_tables)

        matched_in_project = [t for t in tables_for_project if t in doc_report]
        st.success(
            f"Dokumen terbaca: {len(doc_datasets)} dataset, {len(doc_recipes)} recipe. "
            f"{len(matched_in_project)} dari {len(tables_for_project)} table project ini "
            "ketemu namanya persis di dokumen (sisanya kemungkinan bukan tabel DWH "
            "yang di-track Coretan, atau beda nama)."
        )
        with st.expander("🔍 Hasil ekstraksi mentah per table (nama table → daftar kolom di schema-nya)"):
            for t in matched_in_project:
                st.markdown(f"**{t}** ({len(doc_datasets[t].columns)} kolom di schema)")
                st.write(", ".join(doc_datasets[t].columns) if doc_datasets[t].columns else "-")

        uncertain_tables = [t for t in matched_in_project if doc_report[t]["uncertain"]]
        groq_key = f"groq_suggestions_{sel_project}"
        if uncertain_tables:
            if not groq_enabled():
                st.info(
                    f"{len(uncertain_tables)} table ({', '.join(uncertain_tables)}) punya kolom yang "
                    "belum pasti (kena recipe Prepare). Tambahkan `groq_api_key` di secrets kalau mau "
                    "dapat saran AI buat kolom-kolom itu — lihat README."
                )
            elif st.button("🤖 Minta saran AI (Groq) untuk kolom yang belum pasti", key=f"groq_btn_{sel_project}"):
                suggestions = {}
                try:
                    with st.spinner(f"Tanya Groq buat {len(uncertain_tables)} table..."):
                        for t in uncertain_tables:
                            info = doc_report[t]
                            candidates = [c for c in doc_datasets[t].columns if c not in info["confirmed_columns"]]
                            prep_types = sorted(
                                {
                                    s
                                    for r in doc_recipes
                                    if t in r.inputs
                                    for s in r.prepare_step_types
                                }
                            )
                            context = (
                                f"Step transformasi yang menyentuh table ini: {', '.join(prep_types) or '-'}. "
                                f"Kolom yang sudah pasti dipakai (dari join/group/distinct key): "
                                f"{', '.join(sorted(info['confirmed_columns'])) or '-'}."
                            )
                            likely, reasoning = ask_groq_column_suggestions(t, candidates, context)
                            suggestions[t] = {"columns": likely, "reasoning": reasoning}
                    st.session_state[groq_key] = suggestions
                    st.rerun()
                except Exception as e:  # noqa: BLE001 - tampilkan apapun errornya ke user
                    st.error(f"Gagal minta saran Groq: {e}")

    sel_tables = st.multiselect(
        "Pilih Table yang relevan untuk project ini",
        tables_for_project,
        default=tables_for_project,
        key=f"selkol_tables_{sel_project}",
    )

    picked_frames = []
    for t in sel_tables:
        st.markdown(f"**📄 {t}**")
        cols_for_table = (
            proj_df[proj_df["Short Table"] == t][["Column DWH", "Status"]]
            .dropna(subset=["Column DWH"])
            .drop_duplicates()
            .sort_values("Column DWH")
            .reset_index(drop=True)
        )

        info = doc_report.get(t)
        if info is not None:
            confirmed = info["confirmed_columns"]
            uncertain = info["uncertain"]
            ai_suggested = groq_suggestions.get(t, {}).get("columns", set())

            def _sumber(c, confirmed=confirmed, uncertain=uncertain, ai_suggested=ai_suggested):
                if c in confirmed:
                    return "✅ confirmed"
                if c in ai_suggested:
                    return "🤖 saran AI (belum pasti)"
                return "⚠️ cek manual" if uncertain else "📄 ikut alur"

            cols_for_table["Sumber Dokumen"] = cols_for_table["Column DWH"].apply(_sumber)
            default_pilih = (
                cols_for_table["Column DWH"].isin(confirmed)
                | cols_for_table["Column DWH"].isin(ai_suggested)
                | (not uncertain)
            )
            if t in groq_suggestions and groq_suggestions[t].get("reasoning"):
                st.caption(f"🤖 Alasan saran AI: {groq_suggestions[t]['reasoning']}")
        elif doc_file is not None:
            cols_for_table["Sumber Dokumen"] = "❓ table tidak ada di dokumen"
            default_pilih = True
        else:
            cols_for_table["Sumber Dokumen"] = "-"
            default_pilih = True
        cols_for_table.insert(0, "Pilih", default_pilih)

        edited = st.data_editor(
            cols_for_table,
            hide_index=True,
            use_container_width=True,
            disabled=["Column DWH", "Status", "Sumber Dokumen"],
            key=f"selkol_editor_{sel_project}_{t}",
        )
        picked_cols = edited.loc[edited["Pilih"], "Column DWH"].tolist()
        if picked_cols:
            picked_frames.append(
                proj_df[(proj_df["Short Table"] == t) & (proj_df["Column DWH"].isin(picked_cols))]
            )

    st.divider()
    if st.button("✅ Go Selection", type="primary", disabled=not picked_frames):
        new_rows = pd.concat(picked_frames, ignore_index=True) if picked_frames else pd.DataFrame(columns=DISPLAY_COLUMNS)

        try:
            combined, unique_df, compile_df = save_project_selection(sel_project, new_rows)
        except Exception as e:  # noqa: BLE001 - tampilkan apapun errornya ke user
            st.error(f"Gagal menyimpan seleksi: {e}")
            st.stop()

        target = "Google Sheets" if gsheet_enabled() else f"`{DEFAULT_PATH}`"
        st.success(
            f"Tersimpan ke {target}: {len(new_rows)} kolom untuk "
            f"project **{sel_project}**. Total keseluruhan: {len(combined)} baris seleksi, "
            f"{len(unique_df)} kolom unik di {compile_df['Short Table'].nunique() if not compile_df.empty else 0} table."
        )

        if not gsheet_enabled():
            try:
                with open(DEFAULT_PATH, "rb") as f:
                    st.download_button(
                        f"⬇️ Download {DEFAULT_PATH}",
                        f.read(),
                        file_name=DEFAULT_PATH,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
            except FileNotFoundError:
                pass

st.divider()
with st.expander("ℹ️ Keterangan warna Status"):
    st.markdown(
        "- 🟩 **Table** — kolom DWH berhasil dipetakan ke table & kolom Bronze\n"
        "- 🟨 **Hardcode** — tidak ada source column, pakai formula/nilai tetap (lihat kolom Formula Hardcode)\n"
        "- 🟥 **Gap** — source table ada tapi source column kosong, dan tidak ada formula hardcode"
    )
