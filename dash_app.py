"""
Migration Progress Explorer - versi Dash.

Alternatif dari app.py (Streamlit) buat kondisi jaringan/device yang blokir
protokol WebSocket (Streamlit WAJIB pakai WebSocket buat reaktivitasnya,
Dash defaultnya pakai HTTP request/AJAX biasa buat callback-nya, jadi lebih
tahan terhadap proxy/firewall korporat yang ketat).

Logika data (load_data, union DWH+recipe, simpan seleksi, dst) 100% reuse
dari data_core.py - SAMA PERSIS dengan yang dipakai app.py, jadi hasil dari
kedua UI ini selalu konsisten.

Run lokal:  python dash_app.py
Run Docker: lihat Dockerfile.dash
"""

from __future__ import annotations

import base64
import io

import dash
import dash_bootstrap_components as dbc
import pandas as pd
from dash import Input, Output, State, callback_context, dash_table, dcc, html
from dash.exceptions import PreventUpdate

from dataiku_json import parse_dataiku_json
from stage_mapping import build_full_lineage_table

from data_core import (
    DEFAULT_PATH,
    DISPLAY_COLUMNS,
    SELECTED_SHEET_COLUMNS,
    STATUS_COLORS,
    build_table_confirmed_map,
    build_table_union,
    gsheet_enabled,
    load_data,
    pipeline_stage_counts,
    rows_from_union,
    save_project_selection,
)

# ---------------------------------------------------------------------
# Data dimuat sekali di awal proses (module level) - beda dari Streamlit
# yang bisa cache-per-fungsi dgn gampang, Dash lebih simpel load sekali di
# awal karena app ini single-file-source (Checking_Progress_Migration.xlsx
# tidak berubah selama proses jalan, restart proses = data ke-refresh).
# ---------------------------------------------------------------------
df = load_data(DEFAULT_PATH)
lineage_df = build_full_lineage_table()

STAGE_COL_MAP = {
    "DWH": ("DWH Table", "DWH Column"),
    "Bronze": ("Bronze Table", "Bronze Column"),
    "Silver 1": ("Silver1 Table", "Silver1 Column"),
    "Silver 2": ("Silver2 Table", "Silver2 Column"),
}

MODE_OPTIONS = [
    {"label": "🔎 Cari by Table", "value": "table"},
    {"label": "🔎 Cari by Project", "value": "project"},
    {"label": "🧩 Info Table by Project", "value": "info"},
    {"label": "📈 Kelengkapan Stage", "value": "stage"},
]

app = dash.Dash(
    __name__,
    external_stylesheets=[
        dbc.themes.DARKLY,
        "https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap",
    ],
    suppress_callback_exceptions=True,
    title="Migration Progress Explorer",
)
server = app.server  # dipakai gunicorn/Docker

app.index_string = """<!DOCTYPE html>
<html>
    <head>
        {%metas%}
        <title>{%title%}</title>
        {%favicon%}
        {%css%}
    </head>
    <body>
        {%app_entry%}
        <footer>
            {%config%}
            {%scripts%}
            {%renderer%}
        </footer>
    </body>
</html>"""


def status_style_conditional(status_col: str = "Status") -> list:
    return [
        {
            "if": {"filter_query": f'{{{status_col}}} = "{status}"', "column_id": status_col},
            "backgroundColor": color,
            "color": "white" if status == "Gap" else "black",
        }
        for status, color in STATUS_COLORS.items()
    ]


def metric_row(items: list[tuple[str, object]]) -> dbc.Row:
    return dbc.Row(
        [
            dbc.Col(
                dbc.Card(
                    dbc.CardBody(
                        [
                            html.Div(label, className="text-muted", style={"fontSize": "12.5px"}),
                            html.Div(str(value), style={"fontSize": "26px", "fontWeight": 600}),
                        ]
                    ),
                    className="metric-card mb-3",
                ),
                width="auto",
            )
            for label, value in items
        ],
        className="g-3 mb-2",
    )


def data_table(df_in: pd.DataFrame, id_: str, **kwargs) -> dash_table.DataTable:
    return dash_table.DataTable(
        id=id_,
        data=df_in.to_dict("records"),
        columns=[{"name": c, "id": c} for c in df_in.columns],
        page_size=25,
        sort_action="native",
        filter_action="native",
        style_table={"overflowX": "auto"},
        style_cell={"textAlign": "left", "padding": "6px", "fontFamily": "sans-serif", "fontSize": 13},
        style_header={"fontWeight": "bold"},
        **kwargs,
    )


# ---------------------------------------------------------------------
# Layout
# ---------------------------------------------------------------------
app.layout = dbc.Container(
    [
        html.H2("🔍 Migration Progress Explorer", className="mt-3"),
        html.P(
            "Sumber data: sheet 'Coretan Checking Migration' di Checking_Progress_Migration.xlsx "
            "— DWH → Bronze → Silver Tier 1 → Silver Tier 2",
            className="text-muted",
        ),
        dbc.RadioItems(
            id="mode",
            options=MODE_OPTIONS,
            value="table",
            inline=True,
            className="btn-group flex-wrap mb-2",
            inputClassName="btn-check",
            labelClassName="btn btn-outline-primary",
            labelCheckedClassName="active",
        ),
        html.Hr(),
        html.Div(id="mode-content"),
        # Store buat share parsed docx/json antar callback mode "info" (dataset/
        # recipe objek Python tidak JSON-serializable, jadi disimpan dalam bentuk
        # ringkas: nama dataset->type/connection/columns, dan table_confirmed_map
        # dari build_table_confirmed_map() yang sudah plain dict).
        dcc.Store(id="info-doc-store"),
    ],
    fluid=False,
    className="pb-5",
)


# ---------------------------------------------------------------------
# Router: render konten sesuai mode
# ---------------------------------------------------------------------
@app.callback(Output("mode-content", "children"), Input("mode", "value"))
def render_mode(mode):
    if mode == "table":
        return render_table_mode()
    if mode == "project":
        return render_project_mode()
    if mode == "info":
        return render_info_mode()
    return render_stage_mode()


# ============================== Mode: Cari by Table ==============================
def render_table_mode():
    tables = sorted(t for t in df["Short Table"].dropna().unique())
    return html.Div(
        [
            dbc.Select(
                id="table-dropdown",
                options=[{"label": t, "value": t} for t in tables],
                placeholder="Pilih table...",
                className="mb-3",
                style={"maxWidth": "400px"},
            ),
            html.Div(id="table-mode-body"),
        ]
    )


@app.callback(Output("table-mode-body", "children"), Input("table-dropdown", "value"))
def update_table_mode(selected_table):
    if not selected_table:
        return html.P("Pilih table dulu.", className="text-muted")

    subset = df[df["Short Table"] == selected_table]
    projects_using = sorted(subset["Project"].dropna().unique())
    stats = pipeline_stage_counts(subset)

    return html.Div(
        [
            metric_row(
                [
                    ("Jumlah Project Pemakai", len(projects_using)),
                    ("Jumlah Baris Kolom", len(subset)),
                    ("Jumlah Kolom DWH Unik", stats["unik"]),
                ]
            ),
            metric_row(
                [
                    ("🟩 ke Bronze", stats["bronze"]),
                    ("ke Silver", stats["silver"]),
                    ("🟨 Hardcode", stats["hardcode"]),
                    ("🟥 Gap", stats["gap"]),
                ]
            ),
            html.H4("📋 Project yang memakai table ini"),
            html.P(", ".join(projects_using) if projects_using else "-"),
            html.H4("📊 Detail Kolom: DWH → Bronze → Silver Tier 1 → Silver Tier 2"),
            data_table(
                subset[[c for c in DISPLAY_COLUMNS if c in subset.columns]]
                .fillna("-")
                .sort_values(["Project", "Column DWH"]),
                "table-detail-table",
                style_data_conditional=status_style_conditional(),
            ),
        ]
    )


# ============================== Mode: Cari by Project ==============================
def render_project_mode():
    projects = sorted(p for p in df["Project"].dropna().unique())
    return html.Div(
        [
            dbc.Select(
                id="project-dropdown",
                options=[{"label": p, "value": p} for p in projects],
                placeholder="Pilih project...",
                className="mb-3",
                style={"maxWidth": "400px"},
            ),
            html.Div(id="project-mode-body"),
        ]
    )


@app.callback(Output("project-mode-body", "children"), Input("project-dropdown", "value"))
def update_project_mode(selected_project):
    if not selected_project:
        return html.P("Pilih project dulu.", className="text-muted")

    subset = df[df["Project"] == selected_project]
    tables_used = sorted(subset["Short Table"].dropna().unique())
    stats = pipeline_stage_counts(subset)

    return html.Div(
        [
            metric_row(
                [
                    ("Jumlah Table Dipakai", len(tables_used)),
                    ("Jumlah Baris Kolom", len(subset)),
                    ("Jumlah Kolom DWH Unik", stats["unik"]),
                ]
            ),
            metric_row(
                [
                    ("🟩 ke Bronze", stats["bronze"]),
                    ("ke Silver", stats["silver"]),
                    ("🟨 Hardcode", stats["hardcode"]),
                    ("🟥 Gap", stats["gap"]),
                ]
            ),
            html.H4("📋 Table yang dipakai project ini"),
            html.P(", ".join(tables_used) if tables_used else "-"),
            html.H4("📊 Detail Kolom: DWH → Bronze → Silver Tier 1 → Silver Tier 2"),
            data_table(
                subset[[c for c in DISPLAY_COLUMNS if c in subset.columns]]
                .fillna("-")
                .sort_values(["Short Table", "Column DWH"]),
                "project-detail-table",
                style_data_conditional=status_style_conditional(),
            ),
        ]
    )


# ============================== Mode: Kelengkapan Stage ==============================
def render_stage_mode():
    domain_opts = sorted(d for d in lineage_df["Domain"].unique() if d != "-")
    return html.Div(
        [
            html.P(
                "Cross-check independen kelengkapan pipeline DWH → Bronze → Silver Tier 1 → "
                "Silver Tier 2 — dihitung ulang dari 3 file mapping teknis terpisah, BUKAN dari "
                "kolom Bronze/Silver yang sudah tercatat di Coretan.",
                className="text-muted",
            ),
            dbc.RadioItems(
                id="stage-pick",
                options=[{"label": k, "value": k} for k in STAGE_COL_MAP],
                value="DWH",
                inline=True,
                className="btn-group mb-2",
                inputClassName="btn-check",
                labelClassName="btn btn-outline-secondary",
                labelCheckedClassName="active",
            ),
            dbc.Row(
                [
                    dbc.Col(dbc.Input(id="stage-table-kw", placeholder="Nama table..."), width="auto"),
                    dbc.Col(dbc.Input(id="stage-col-kw", placeholder="Nama kolom (opsional)..."), width="auto"),
                ],
                className="g-2 mb-2 mt-1",
            ),
            dcc.Dropdown(
                id="stage-domain",
                options=[{"label": d, "value": d} for d in domain_opts],
                multi=True,
                placeholder="Filter Domain (kosongkan buat semua)...",
                className="mb-3",
            ),
            html.Div(id="stage-mode-body"),
        ]
    )


@app.callback(
    Output("stage-mode-body", "children"),
    Input("stage-pick", "value"),
    Input("stage-table-kw", "value"),
    Input("stage-col-kw", "value"),
    Input("stage-domain", "value"),
)
def update_stage_mode(stage_pick, table_kw, col_kw, domain_sel):
    table_col, col_col = STAGE_COL_MAP[stage_pick or "DWH"]
    filtered = lineage_df
    if table_kw:
        filtered = filtered[filtered[table_col].str.upper().str.contains(table_kw.strip().upper(), na=False)]
    if col_kw:
        filtered = filtered[filtered[col_col].str.upper().str.contains(col_kw.strip().upper(), na=False)]
    if domain_sel:
        filtered = filtered[filtered["Domain"].isin(domain_sel)]

    with_dwh = filtered[filtered["DWH Table"] != "-"]
    return html.Div(
        [
            metric_row(
                [
                    ("Jumlah Baris Jalur", len(filtered)),
                    (
                        "Kombinasi DWH Table.Column Unik",
                        with_dwh[["DWH Table", "DWH Column"]].drop_duplicates().shape[0] if not with_dwh.empty else 0,
                    ),
                    ("Sampai Silver 2", int((filtered["Silver2 Table"] != "-").sum())),
                ]
            ),
            data_table(filtered, "stage-table"),
        ]
    )


# ============================== Mode: Info Table by Project ==============================
def render_info_mode():
    projects = sorted(p for p in df["Project"].dropna().unique())
    return html.Div(
        [
            html.P(
                "Pilih project → upload dump Dataiku Flow (JSON) → pilih table → "
                "centang kolom yang bener-bener dipakai. Hasilnya disimpan ke "
                + ("Google Sheets" if gsheet_enabled() else f"`{DEFAULT_PATH}`") + ".",
                className="text-muted",
            ),
            dbc.Select(
                id="info-project-dropdown",
                options=[{"label": p, "value": p} for p in projects],
                placeholder="Pilih project...",
                className="mb-2",
                style={"maxWidth": "400px"},
            ),
            dcc.Upload(
                id="info-upload",
                children=html.Div(["Drag & drop atau ", html.A("klik buat pilih file (.json)")]),
                style={
                    "width": "100%", "height": "50px", "lineHeight": "50px", "borderWidth": "1px",
                    "borderStyle": "dashed", "borderRadius": "5px", "textAlign": "center", "marginTop": "5px",
                },
            ),
            html.Div(id="info-upload-status", className="mt-2"),
            html.Div(id="info-mode-body"),
        ]
    )


def _parse_doc_contents(contents: str, filename: str):
    _, content_string = contents.split(",", 1)
    decoded = base64.b64decode(content_string)
    return parse_dataiku_json(io.BytesIO(decoded))


@app.callback(
    Output("info-doc-store", "data"),
    Output("info-upload-status", "children"),
    Input("info-upload", "contents"),
    State("info-upload", "filename"),
)
def parse_uploaded(contents, filename):
    if contents is None:
        raise PreventUpdate
    try:
        doc_datasets, doc_recipes = _parse_doc_contents(contents, filename)
    except ImportError as e:
        return None, html.P(str(e), style={"color": "red"})
    except Exception as e:  # noqa: BLE001
        return None, html.P(f"Gagal parse file: {e}", style={"color": "red"})

    table_confirmed_map = build_table_confirmed_map(doc_recipes)
    dataset_meta = {
        name: {"type": ds.type, "connection": ds.connection, "columns": ds.columns}
        for name, ds in doc_datasets.items()
    }
    store = {
        "filename": filename,
        "dataset_meta": dataset_meta,
        "table_confirmed_map": table_confirmed_map,
        "n_recipes": len(doc_recipes),
        "input_datasets": sorted({n for r in doc_recipes for n in r.inputs}),
        "output_datasets": sorted({n for r in doc_recipes for n in r.outputs}),
    }
    total_cols_from_map = len({(t, c) for t, cols in table_confirmed_map.items() for c in cols})
    total_table = len(dataset_meta) or len(table_confirmed_map)
    total_kolom = sum(len(v["columns"]) for v in dataset_meta.values()) or total_cols_from_map
    status = html.P(
        f"✅ {filename} terbaca: {total_table} table, {total_kolom} kolom, {len(doc_recipes)} recipe.",
        style={"color": "#3cb371"},
    )
    return store, status


# Connection yang dipakai di seluruh mode "Info Table by Project" - table
# dari connection lain (mis. AMFS_dataset, MPI_DATAMART) di-exclude TOTAL
# (bukan cuma disembunyiin di default), nggak muncul sama sekali di filter,
# metric, atau tabel union manapun.
ALLOWED_CONN_KEYWORDS = ("DWH", "EDM", "BICC")


def _conn_of(t: str, dataset_meta: dict) -> str:
    return (dataset_meta.get(t, {}).get("connection") or "-").upper()


def _allowed_tables_pool(all_tables_pool: list, dataset_meta: dict) -> list:
    return [
        t for t in all_tables_pool if any(k in _conn_of(t, dataset_meta) for k in ALLOWED_CONN_KEYWORDS)
    ]


def _dataset_mini_table(dataset_meta: dict, names: list, id_: str):
    if not names:
        return html.P("-", className="text-muted")
    rows = [
        {
            "Dataset": n,
            "Type": dataset_meta.get(n, {}).get("type") or "-",
            "Connection": dataset_meta.get(n, {}).get("connection") or "-",
            "Jumlah Kolom": len(dataset_meta.get(n, {}).get("columns") or []),
        }
        for n in names
    ]
    rows.sort(key=lambda r: (0 if "DWH" in r["Connection"].upper() else 1, r["Connection"], r["Dataset"]))
    return data_table(pd.DataFrame(rows), id_)


@app.callback(
    Output("info-mode-body", "children"),
    Input("info-project-dropdown", "value"),
    Input("info-doc-store", "data"),
)
def update_info_mode(sel_project, doc_store):
    if not sel_project:
        return html.P("Pilih project dulu.", className="text-muted")

    doc_store = doc_store or {}
    table_confirmed_map = doc_store.get("table_confirmed_map", {})
    if not table_confirmed_map:
        return html.P(
            "Belum ada dokumen/dump di-upload (atau belum ada recipe yang confirmed "
            "makai table project ini) — upload dulu buat lihat kolom yang dipakai di Dataiku.",
            className="text-muted",
        )

    proj_df = df[df["Project"] == sel_project]
    tables_for_project = sorted(proj_df["Short Table"].dropna().unique())
    dataset_meta = doc_store.get("dataset_meta", {})

    matched_in_project = [t for t in tables_for_project if t in dataset_meta or t in table_confirmed_map]
    edm_datasets = sorted(
        n for n in dataset_meta if n not in matched_in_project and "EDM" in (dataset_meta[n].get("connection") or "").upper()
    )

    # Pool connection: table project ini (DWH) DAN semua table lain yang
    # confirmed dipakai di Dataiku (termasuk EDM), TAPI cuma yang connection-
    # nya DWH/EDM/BICC - AMFS_dataset/MPI_DATAMART/dll dibuang total.
    raw_pool = sorted(set(tables_for_project) | set(table_confirmed_map))
    all_tables_pool = _allowed_tables_pool(raw_pool, dataset_meta) if dataset_meta else list(table_confirmed_map)
    conn_opts = sorted({dataset_meta.get(t, {}).get("connection") or "-" for t in all_tables_pool})

    total_table = len(all_tables_pool)
    total_kolom = sum(len(dataset_meta.get(t, {}).get("columns") or []) for t in all_tables_pool) or len(
        {(t, c) for t in all_tables_pool for c in table_confirmed_map.get(t, {})}
    )

    children = [
        metric_row(
            [
                ("Total Table", total_table),
                ("Total Kolom", total_kolom),
                ("Jadi Input Recipe", len(doc_store.get("input_datasets", []))),
                ("Jadi Output Recipe", len(doc_store.get("output_datasets", []))),
            ]
        ),
    ]
    if dataset_meta:
        children += [
            html.H5(f"📘 Data DWH ({len(matched_in_project)})", className="mt-3"),
            _dataset_mini_table(dataset_meta, matched_in_project, "info-dwh-table"),
            html.H5(f"📦 Data EDM ({len(edm_datasets)})", className="mt-3"),
            _dataset_mini_table(dataset_meta, edm_datasets, "info-edm-table"),
        ]

    children += [
        html.H5("🧩 Kolom yang dipakai di Dataiku per table", className="mt-3"),
        dbc.Select(
            id="info-conn-filter",
            options=[{"label": "Semua Connection", "value": "__all__"}]
            + [{"label": c, "value": c} for c in conn_opts],
            value="__all__",
            className="mb-2",
            style={"maxWidth": "420px"},
        ),
        html.Div(id="info-table-area"),
    ]
    return html.Div(children)


@app.callback(
    Output("info-table-area", "children"),
    Input("info-project-dropdown", "value"),
    Input("info-doc-store", "data"),
    Input("info-conn-filter", "value"),
)
def update_info_table(sel_project, doc_store, conn_filter):
    if not sel_project:
        raise PreventUpdate

    doc_store = doc_store or {}
    table_confirmed_map = doc_store.get("table_confirmed_map", {})
    dataset_meta = doc_store.get("dataset_meta", {})
    proj_df = df[df["Project"] == sel_project]
    tables_for_project = sorted(proj_df["Short Table"].dropna().unique())
    raw_pool = sorted(set(tables_for_project) | set(table_confirmed_map))
    all_tables_pool = _allowed_tables_pool(raw_pool, dataset_meta) if dataset_meta else list(table_confirmed_map)

    if conn_filter and conn_filter != "__all__":
        tables_to_show = [t for t in all_tables_pool if _conn_of(t, dataset_meta) == conn_filter.upper()]
    else:
        tables_to_show = all_tables_pool

    frames = []
    for t in tables_to_show:
        union_df = build_table_union(proj_df, table_confirmed_map, t)
        if union_df.empty:
            continue
        union_df.insert(0, "Connection", dataset_meta.get(t, {}).get("connection") or "-")
        union_df.insert(0, "Short Table", t)
        frames.append(union_df)

    if not frames:
        return html.P(
            "Nggak ada kolom project ini yang confirmed dipakai di dokumen/dump yang di-upload "
            "(buat filter connection ini).",
            className="text-muted",
        )

    combined = pd.concat(frames, ignore_index=True)
    # Checkbox seleksi (row_selectable) - default kecentang buat baris yang
    # "Pilih"=True (ada di DWH DAN confirmed dipakai di Dataiku, lihat
    # build_table_union() di data_core.py). Kolom "Pilih" mentahnya
    # disembunyikan dari tampilan, diganti checkbox di kolom paling kiri.
    selected_rows = combined.index[combined["Pilih"]].tolist()
    display_cols = [c for c in combined.columns if c != "Pilih"]

    table_out = dash_table.DataTable(
        id="info-union-table",
        data=combined[display_cols].to_dict("records"),
        columns=[{"name": c, "id": c} for c in display_cols],
        row_selectable="multi",
        selected_rows=selected_rows,
        page_size=50,
        sort_action="native",
        filter_action="native",
        style_table={"overflowX": "auto"},
        style_cell={"textAlign": "left", "padding": "6px", "fontFamily": "sans-serif", "fontSize": 13},
        style_header={"fontWeight": "bold"},
        style_data_conditional=status_style_conditional(),
    )
    return html.Div(
        [
            html.P(
                "Kolom yang ada di DWH & dipakai di Dataiku sudah otomatis tercentang. "
                "Centang/hilangkan centang di kiri buat sesuaikan, lalu klik Go Compare.",
                className="text-muted",
            ),
            table_out,
            dbc.Button("✅ Go Compare", id="info-go-btn", n_clicks=0, color="primary", className="mt-3"),
            html.Div(id="info-save-status", className="mt-2"),
        ]
    )


@app.callback(
    Output("info-save-status", "children"),
    Input("info-go-btn", "n_clicks"),
    State("info-union-table", "data"),
    State("info-union-table", "selected_rows"),
    State("info-project-dropdown", "value"),
    prevent_initial_call=True,
)
def go_compare(n_clicks, table_data, selected_rows, sel_project):
    if not n_clicks or not table_data:
        raise PreventUpdate

    picked = pd.DataFrame(table_data).iloc[selected_rows or []]
    if picked.empty:
        return html.P("Nggak ada baris yang dicentang.", style={"color": "orange"})

    frames = [
        rows_from_union(sel_project, t, group) for t, group in picked.groupby("Short Table")
    ]
    new_rows = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=SELECTED_SHEET_COLUMNS[:-1])
    try:
        combined, unique_df = save_project_selection(sel_project, new_rows)
    except Exception as e:  # noqa: BLE001
        return html.P(f"Gagal menyimpan seleksi: {e}", style={"color": "red"})

    target = "Google Sheets" if gsheet_enabled() else f"`{DEFAULT_PATH}`"
    return html.P(
        f"✅ Tersimpan ke {target}: {len(new_rows)} kolom untuk project {sel_project}. "
        f"Total keseluruhan: {len(combined)} baris seleksi, {len(unique_df)} kolom unik.",
        style={"color": "green"},
    )


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8501, debug=False)
