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
import os

import pandas as pd
import streamlit as st

from dataiku_json import parse_dataiku_json
from stage_mapping import build_full_lineage_table
from data_core import (
    BACKUP_PATH,
    DEFAULT_PATH,
    DISPLAY_COLUMNS,
    SELECTED_SHEET_COLUMNS,
    SHEET_NAME,
    STATUS_COLORS,
    build_table_confirmed_map,
    build_table_union,
    gsheet_enabled,
    pipeline_stage_counts,
    rows_from_union,
)
from data_core import load_data as _load_data_core
from data_core import save_project_selection as _save_project_selection_core

# Jembatan ke st.secrets (Streamlit Cloud/lokal, .streamlit/secrets.toml) -
# data_core.get_secret cuma baca env var (portable ke semua host), jadi di
# sini secrets.toml (kalau ada) disalin ke os.environ SEKALI di awal supaya
# tetap kepakai tanpa data_core perlu tahu soal Streamlit sama sekali.
try:
    for _key in ("gsheet_webapp_url", "gsheet_webapp_token"):
        if _key not in os.environ and _key in st.secrets:
            os.environ[_key] = st.secrets[_key]
except Exception:
    pass


@st.cache_data(show_spinner="Membaca dump Dataiku Flow (dump besar bisa makan waktu ~10-20 detik)...")
def parse_uploaded_doc(file_bytes: bytes, file_name: str):
    """Wrapper cached di atas parser dump JSON supaya dump besar tidak
    di-parse ulang tiap kali ada interaksi UI lain (checkbox, dsb).

    Cuma JSON (dari dump_flow_via_notebook.py/export_code.txt) - dukungan
    .docx sudah dihapus (JSON lebih portable, jalan di environment manapun
    tanpa butuh python-docx)."""
    return parse_dataiku_json(io.BytesIO(file_bytes))


@st.cache_data(show_spinner="Membangun tabel lineage DWH → Bronze → Silver 1 → Silver 2...")
def load_full_lineage_table() -> pd.DataFrame:
    """Wrapper cached di atas build_full_lineage_table() - baca ulang 3
    sumber file mapping teknis (dwh_to_flat.csv, Source ke Silver 1.csv,
    7 file domain Silver1->Silver2) makan waktu ~1-2 detik, tidak perlu
    diulang tiap interaksi UI."""
    return build_full_lineage_table()


@st.cache_data(show_spinner="Memuat data dari Excel...")
def load_data(file_source) -> pd.DataFrame:
    """Wrapper cached di atas data_core.load_data()."""
    return _load_data_core(file_source)


def style_status(val):
    color = STATUS_COLORS.get(val)
    if not color:
        return ""
    text_color = "white" if val == "Gap" else "black"
    return f"background-color: {color}; color: {text_color};"


def render_detail_table(df_subset: pd.DataFrame):
    cols = [c for c in DISPLAY_COLUMNS if c in df_subset.columns]
    # tampilkan "-" buat stage yang belum termapping (bukan "None"/kosong) -
    # tiap stage (Bronze/Silver1/Silver2) independen, tidak diasumsikan
    # harus berurutan (bisa aja cuma ada Silver 2 tanpa Silver 1, dst).
    view = df_subset[cols].fillna("-")
    if "Status" in view.columns:
        try:
            styled = view.style.map(style_status, subset=["Status"])
        except AttributeError:
            # pandas lama belum punya Styler.map, fallback ke applymap
            styled = view.style.applymap(style_status, subset=["Status"])
        st.dataframe(styled, use_container_width=True, hide_index=True)
    else:
        st.dataframe(view, use_container_width=True, hide_index=True)


def render_dataset_table(datasets: dict, names: list, key_prefix: str):
    """Render daftar dataset (hasil parsing dokumen Dataiku) sebagai tabel
    yang bisa difilter by Type & Connection, plus pilih 1 buat lihat daftar
    kolomnya."""
    if not names:
        st.write("-")
        return

    table_df = pd.DataFrame(
        [
            {
                "Dataset": n,
                "Type": datasets[n].type or "-",
                "Connection": datasets[n].connection or "-",
                "Jumlah Kolom": len(datasets[n].columns),
            }
            for n in names
        ]
    )
    # Connection yang namanya nyebut "DWH" ditampilin duluan (mis. MPI_AMFS_DWH
    # sebelum WRCSQL32_EDM), baru sisanya alfabetis - biar sumber utama (DWH)
    # selalu di atas sebelum source tambahan lain (EDM, dst).
    conn_priority = table_df["Connection"].str.upper().str.contains("DWH").map({True: 0, False: 1})
    table_df = (
        table_df.assign(_conn_priority=conn_priority)
        .sort_values(["_conn_priority", "Connection", "Dataset"])
        .drop(columns="_conn_priority")
        .reset_index(drop=True)
    )

    type_opts = sorted(table_df["Type"].unique())
    conn_opts = sorted(table_df["Connection"].unique())
    c1, c2 = st.columns(2)
    type_sel = c1.multiselect("Filter Type", type_opts, default=type_opts, key=f"{key_prefix}_type")
    conn_sel = c2.multiselect("Filter Connection", conn_opts, default=conn_opts, key=f"{key_prefix}_conn")
    keyword = st.text_input("Filter nama dataset", key=f"{key_prefix}_kw")

    filtered = table_df[table_df["Type"].isin(type_sel) & table_df["Connection"].isin(conn_sel)]
    if keyword:
        filtered = filtered[filtered["Dataset"].str.upper().str.contains(keyword.strip().upper())]

    st.caption(f"Menampilkan {len(filtered)} dari {len(table_df)} dataset.")
    st.dataframe(filtered, hide_index=True, use_container_width=True)

    if not filtered.empty:
        detail_pick = st.selectbox(
            "Lihat daftar kolom dataset",
            ["-"] + filtered["Dataset"].tolist(),
            key=f"{key_prefix}_detail",
        )
        if detail_pick != "-":
            st.write(", ".join(datasets[detail_pick].columns) or "-")


save_project_selection = _save_project_selection_core


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
    ["🔎 Cari by Table", "🔎 Cari by Project", "🧩 Info Table by Project", "📈 Kelengkapan Stage"],
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
    project_filter = st.multiselect(
        "Filter by Project (kosongkan buat tampilkan semua)",
        projects_using,
        default=[],
        key=f"detail_project_filter_{selected_table}",
    )
    detail_subset = subset[subset["Project"].isin(project_filter)] if project_filter else subset
    render_detail_table(detail_subset.sort_values(["Project", "Column DWH"]))

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

elif mode == "🧩 Info Table by Project":
    if gsheet_enabled():
        st.caption(
            "Pilih project → upload dump Dataiku Flow (JSON) → pilih table → "
            "centang kolom yang bener-bener dipakai. Hasilnya disimpan ke **Google "
            "Sheets** (2 tab: Selected Columns, Unique Columns) — jadi tetap "
            "persisten meskipun hosting di-restart. Submit ulang untuk project "
            "yang sama akan menggantikan seleksi lama project itu, bukan menumpuk duplikat."
        )
    else:
        st.caption(
            "Pilih project → upload dump Dataiku Flow (JSON) → pilih table → "
            "centang kolom yang bener-bener dipakai. Hasilnya ditulis balik sebagai 2 "
            f"sheet tambahan di `{DEFAULT_PATH}` sendiri (Selected Columns, Unique "
            "Columns) — sheet Coretan & sheet lain tidak disentuh. Submit ulang "
            "untuk project yang sama akan menggantikan seleksi lama project itu, "
            "bukan menumpuk duplikat. "
            f"Sebelum tiap kali nulis, file di-backup dulu ke `{BACKUP_PATH}`. "
            "(Belum ada kredensial Google Sheets — kalau di-deploy ke hosting "
            "gratis, hasil seleksi ini bisa hilang tiap container restart; "
            "lihat README bagian setup Google Sheets.)"
        )

    # sengaja pakai df (bukan df_f) - filter sidebar (Kategori Project/Status)
    # itu buat mode pencarian, tidak relevan buat Seleksi Kolom. Kalau ikut
    # df_f, table yang semua kolomnya Gap bisa hilang dari pilihan cuma
    # gara-gara filter Status di sidebar tidak nyentang "Gap".
    projects = sorted(p for p in df["Project"].dropna().unique())
    sel_project = st.selectbox("Pilih Project", projects, key="selkol_project")

    proj_df = df[df["Project"] == sel_project]
    tables_for_project = sorted(proj_df["Short Table"].dropna().unique())

    with st.expander("📄 Upload dump Dataiku Flow (JSON, opsional)"):
        st.caption(
            "Upload dump JSON dari `dump_flow_via_notebook.py`/`export_code.txt`. "
            "Buat tiap table, kolomnya digabung dari 2 sumber: yang tercatat di "
            "Coretan (DWH) dan yang confirmed dipakai di recipe Dataiku — lalu "
            "ditandai ada di DWH & dipakai di Dataiku, cuma di DWH, atau cuma di "
            "Dataiku. Default kecentang kalau ADA DI DWH DAN dipakai di Dataiku; "
            "sisanya perlu direview manual dulu."
        )
        doc_file = st.file_uploader(
            "Upload dump (.json)",
            type=["json"],
            key=f"selkol_doc_{sel_project}",
        )

    doc_datasets = {}
    doc_recipes = []
    table_confirmed_map = {}
    if doc_file is not None:
        try:
            doc_datasets, doc_recipes = parse_uploaded_doc(doc_file.getvalue(), doc_file.name)
        except Exception as e:  # noqa: BLE001 - tampilkan apapun errornya ke user
            st.error(f"Gagal parse file: {e}")
            st.stop()

        table_confirmed_map = build_table_confirmed_map(doc_recipes)
        input_datasets = {n for r in doc_recipes for n in r.inputs}
        output_datasets = {n for r in doc_recipes for n in r.outputs}

        total_table = len(doc_datasets) or len(table_confirmed_map)
        total_columns = sum(len(ds.columns) for ds in doc_datasets.values()) or len(
            {(t, c) for t, cols in table_confirmed_map.items() for c in cols}
        )
        st.success(f"Dump terbaca: {total_table} table, {total_columns} kolom, {len(doc_recipes)} recipe.")

        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Total Table", total_table)
        m2.metric("Total Kolom", total_columns)
        m3.metric("Jadi Input Recipe", len(input_datasets))
        m4.metric("Jadi Output Recipe", len(output_datasets))
        st.caption(
            "Dihitung di level dataset (bukan kolom formula) — dataset yang jadi INPUT "
            "sekaligus OUTPUT (dataset perantara di tengah flow) dihitung di dua-duanya, "
            "jadi totalnya bisa lebih besar dari jumlah dataset unik."
        )

        matched_in_project = [t for t in tables_for_project if t in doc_datasets or t in table_confirmed_map]
        if doc_datasets:
            edm_datasets = sorted(
                n
                for n in doc_datasets
                if n not in matched_in_project and "EDM" in (doc_datasets[n].connection or "").upper()
            )
            with st.expander(f"📘 Data DWH ({len(matched_in_project)})"):
                render_dataset_table(doc_datasets, matched_in_project, key_prefix=f"matched_ds_{sel_project}")
            with st.expander(f"📦 Data EDM ({len(edm_datasets)})"):
                st.caption(
                    "Dataset di flow Dataiku ini yang connection-nya EDM (bukan DWH) dan "
                    "TIDAK match nama-nya ke Short Table Coretan. Table di sini bisa dipilih "
                    "juga di bawah lewat 'Filter Connection', buat lihat detail kolomnya "
                    "yang dipakai di Dataiku."
                )
                render_dataset_table(doc_datasets, edm_datasets, key_prefix=f"other_ds_{sel_project}")

    if not table_confirmed_map:
        st.info("Upload dump JSON dulu buat lihat kolom yang dipakai di Dataiku.")
        st.stop()

    # Pool table: default cuma table yang di-track project ini di Coretan
    # (DWH). Table lain (mis. EDM, yang nggak ada di Coretan sama sekali)
    # cuma muncul kalau user explicit milih connection-nya di Filter
    # Connection - biar default-nya tetap ringkas & fokus ke project ini.
    all_tables_pool = sorted(
        set(tables_for_project) | {t for t in table_confirmed_map if t not in tables_for_project}
    )
    table_connections = {
        t: (doc_datasets[t].connection or "-") if t in doc_datasets else "-" for t in all_tables_pool
    }
    conn_opts = sorted(set(table_connections.values()))
    sel_conn = st.multiselect(
        "Filter Connection (kosongkan buat cuma table project ini)",
        conn_opts,
        default=[],
        key=f"selkol_connfilter_{sel_project}",
    )
    candidate_tables = (
        [t for t in all_tables_pool if table_connections[t] in sel_conn] if sel_conn else tables_for_project
    )

    # Key ikut sel_conn - kalau filter Connection ganti, widget-nya "fresh"
    # lagi (default kosong) daripada nyimpen seleksi lama yang bisa jadi
    # tidak valid lagi buat opsi (candidate_tables) yang baru.
    conn_key_part = ",".join(sorted(sel_conn))
    sel_tables_input = st.multiselect(
        "Pilih Table (kosongkan buat tampilkan semua dari filter di atas)",
        candidate_tables,
        default=[],
        key=f"selkol_tables_{sel_project}_{conn_key_part}",
    )
    sel_tables = sel_tables_input if sel_tables_input else candidate_tables

    picked_frames = []
    for t in sel_tables:
        st.markdown(f"**📄 {t}**")
        cols_for_table = build_table_union(proj_df, table_confirmed_map, t)
        if cols_for_table.empty:
            st.caption("Nggak ada kolom yang confirmed dipakai di Dataiku buat table ini.")
            continue

        used_only_count = int(
            (cols_for_table["Keterangan"] == "📄 Dipakai di Dataiku, TIDAK ada di DWH").sum()
        )
        if used_only_count:
            st.caption(
                f"📄 {used_only_count} kolom dipakai di Dataiku tapi belum tercatat "
                "di Coretan — centang di sini TIDAK otomatis nambah ke Coretan, cuma "
                "penanda buat direview & ditambah manual kalau memang relevan."
            )

        edited = st.data_editor(
            cols_for_table,
            hide_index=True,
            use_container_width=True,
            disabled=["Column DWH", "Status", "Keterangan", "Dipakai di Recipe"],
            key=f"selkol_editor_{sel_project}_{t}",
        )
        picked = edited.loc[edited["Pilih"]]
        if not picked.empty:
            picked_frames.append(rows_from_union(sel_project, t, picked))

    st.divider()
    if st.button("✅ Go Compare", type="primary", disabled=not picked_frames):
        new_rows = (
            pd.concat(picked_frames, ignore_index=True)
            if picked_frames
            else pd.DataFrame(columns=SELECTED_SHEET_COLUMNS[:-1])
        )

        try:
            combined, unique_df = save_project_selection(sel_project, new_rows)
        except Exception as e:  # noqa: BLE001 - tampilkan apapun errornya ke user
            st.error(f"Gagal menyimpan seleksi: {e}")
            st.stop()

        target = "Google Sheets" if gsheet_enabled() else f"`{DEFAULT_PATH}`"
        st.success(
            f"Tersimpan ke {target}: {len(new_rows)} kolom untuk "
            f"project **{sel_project}**. Total keseluruhan: {len(combined)} baris seleksi, "
            f"{len(unique_df)} kolom unik."
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

else:  # 📈 Kelengkapan Stage
    st.caption(
        "Cross-check independen kelengkapan pipeline DWH → Bronze → Silver "
        "Tier 1 → Silver Tier 2 — dihitung ulang dari 3 file mapping teknis "
        "terpisah (`dwh_to_flat.csv`, `Source ke Silver 1.csv`, 7 file domain "
        "Silver1→Silver2 di folder `Silver 1 ke Silver 2/`), **BUKAN** dari "
        "kolom Bronze/Silver yang sudah tercatat di Coretan — jadi bisa "
        "kepakai buat cross-check apakah Coretan-nya masih sinkron. Cari "
        "mulai dari stage manapun (tidak harus dari DWH). Satu kolom DWH "
        "bisa punya lebih dari satu jalur (misal sumber Health & Life "
        "sekaligus), jadi wajar kalau muncul >1 baris buat 1 kolom yang "
        "sama. Ini murni tampilan, tidak disimpan ke mana pun."
    )

    try:
        lineage_df = load_full_lineage_table()
    except FileNotFoundError as e:
        st.error(f"File sumber mapping tidak ditemukan di folder app: {e}")
        st.stop()

    stage_col_map = {
        "DWH": ("DWH Table", "DWH Column"),
        "Bronze": ("Bronze Table", "Bronze Column"),
        "Silver 1": ("Silver1 Table", "Silver1 Column"),
        "Silver 2": ("Silver2 Table", "Silver2 Column"),
    }
    stage_pick = st.radio(
        "Cari mulai dari stage", list(stage_col_map.keys()), horizontal=True, key="stage_search_from"
    )
    table_col, col_col = stage_col_map[stage_pick]

    c1, c2 = st.columns(2)
    table_kw = c1.text_input(f"Nama table di stage {stage_pick}", key="stage_table_kw")
    col_kw = c2.text_input(f"Nama kolom di stage {stage_pick} (opsional)", key="stage_col_kw")

    domain_opts = sorted(d for d in lineage_df["Domain"].unique() if d != "-")
    domain_sel = st.multiselect("Filter Domain (kosongkan buat semua)", domain_opts, default=[], key="stage_domain")

    filtered = lineage_df
    if table_kw:
        filtered = filtered[filtered[table_col].str.upper().str.contains(table_kw.strip().upper(), na=False)]
    if col_kw:
        filtered = filtered[filtered[col_col].str.upper().str.contains(col_kw.strip().upper(), na=False)]
    if domain_sel:
        filtered = filtered[filtered["Domain"].isin(domain_sel)]

    with_dwh = filtered[filtered["DWH Table"] != "-"]

    c1, c2, c3 = st.columns(3)
    c1.metric("Jumlah Baris Jalur", len(filtered))
    c2.metric(
        "Kombinasi DWH Table.Column Unik",
        with_dwh[["DWH Table", "DWH Column"]].drop_duplicates().shape[0] if not with_dwh.empty else 0,
    )
    c3.metric("Sampai Silver 2", int((filtered["Silver2 Table"] != "-").sum()))

    st.dataframe(filtered, hide_index=True, use_container_width=True)

st.divider()
with st.expander("ℹ️ Keterangan warna Status"):
    st.markdown(
        "- 🟩 **Table** — kolom DWH berhasil dipetakan ke table & kolom Bronze\n"
        "- 🟨 **Hardcode** — tidak ada source column, pakai formula/nilai tetap (lihat kolom Formula Hardcode)\n"
        "- 🟥 **Gap** — source table ada tapi source column kosong, dan tidak ada formula hardcode"
    )
