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
import pandas as pd
from dash import Input, Output, State, callback_context, dash_table, dcc, html
from dash.exceptions import PreventUpdate

from dataiku_doc import build_project_column_report, parse_dataiku_doc
from dataiku_json import parse_dataiku_json
from stage_mapping import build_full_lineage_table

from data_core import (
    DEFAULT_PATH,
    DISPLAY_COLUMNS,
    STATUS_COLORS,
    build_table_confirmed_map,
    build_table_union,
    gsheet_enabled,
    load_data,
    pipeline_stage_counts,
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

app = dash.Dash(__name__, suppress_callback_exceptions=True, title="Migration Progress Explorer")
server = app.server  # dipakai gunicorn/Docker


def status_style_conditional(status_col: str = "Status") -> list:
    return [
        {
            "if": {"filter_query": f'{{{status_col}}} = "{status}"', "column_id": status_col},
            "backgroundColor": color,
            "color": "white" if status == "Gap" else "black",
        }
        for status, color in STATUS_COLORS.items()
    ]


def metric_row(items: list[tuple[str, object]]) -> html.Div:
    return html.Div(
        [
            html.Div(
                [html.Div(label, className="metric-label"), html.Div(str(value), className="metric-value")],
                className="metric-box",
            )
            for label, value in items
        ],
        className="metric-row",
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
app.layout = html.Div(
    [
        html.H2("🔍 Migration Progress Explorer"),
        html.P(
            "Sumber data: sheet 'Coretan Checking Migration' di Checking_Progress_Migration.xlsx "
            "— DWH → Bronze → Silver Tier 1 → Silver Tier 2",
            style={"color": "#666"},
        ),
        dcc.RadioItems(
            id="mode",
            options=MODE_OPTIONS,
            value="table",
            labelStyle={"display": "inline-block", "marginRight": "20px"},
        ),
        html.Hr(),
        html.Div(id="mode-content"),
        # Store buat share parsed docx/json antar callback mode "info" (dataset/
        # recipe objek Python tidak JSON-serializable, jadi disimpan dalam bentuk
        # ringkas: nama dataset->type/connection/columns, dan table_confirmed_map
        # dari build_table_confirmed_map() yang sudah plain dict).
        dcc.Store(id="info-doc-store"),
    ],
    style={"maxWidth": "1200px", "margin": "0 auto", "padding": "20px", "fontFamily": "sans-serif"},
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
            dcc.Dropdown(
                id="table-dropdown",
                options=[{"label": t, "value": t} for t in tables],
                placeholder="Pilih table...",
            ),
            html.Div(id="table-mode-body"),
        ]
    )


@app.callback(Output("table-mode-body", "children"), Input("table-dropdown", "value"))
def update_table_mode(selected_table):
    if not selected_table:
        return html.P("Pilih table dulu.", style={"color": "#888"})

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
            dcc.Dropdown(
                id="project-dropdown",
                options=[{"label": p, "value": p} for p in projects],
                placeholder="Pilih project...",
            ),
            html.Div(id="project-mode-body"),
        ]
    )


@app.callback(Output("project-mode-body", "children"), Input("project-dropdown", "value"))
def update_project_mode(selected_project):
    if not selected_project:
        return html.P("Pilih project dulu.", style={"color": "#888"})

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
                style={"color": "#666"},
            ),
            dcc.RadioItems(
                id="stage-pick",
                options=[{"label": k, "value": k} for k in STAGE_COL_MAP],
                value="DWH",
                labelStyle={"display": "inline-block", "marginRight": "15px"},
            ),
            html.Div(
                [
                    dcc.Input(id="stage-table-kw", placeholder="Nama table...", style={"marginRight": "10px"}),
                    dcc.Input(id="stage-col-kw", placeholder="Nama kolom (opsional)..."),
                ],
                style={"marginTop": "10px", "marginBottom": "10px"},
            ),
            dcc.Dropdown(
                id="stage-domain",
                options=[{"label": d, "value": d} for d in domain_opts],
                multi=True,
                placeholder="Filter Domain (kosongkan buat semua)...",
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
                "Pilih project → upload dokumen/dump Dataiku Flow (opsional) → pilih table → "
                "centang kolom yang mau dibawa. Hasilnya disimpan ke "
                + ("Google Sheets" if gsheet_enabled() else f"`{DEFAULT_PATH}`") + ".",
                style={"color": "#666"},
            ),
            dcc.Dropdown(
                id="info-project-dropdown",
                options=[{"label": p, "value": p} for p in projects],
                placeholder="Pilih project...",
            ),
            dcc.RadioItems(
                id="info-doc-format",
                options=[{"label": "JSON", "value": "json"}, {"label": "DOCX", "value": "docx"}],
                value="json",
                labelStyle={"display": "inline-block", "marginRight": "15px"},
                style={"marginTop": "10px"},
            ),
            dcc.Upload(
                id="info-upload",
                children=html.Div(["Drag & drop atau ", html.A("klik buat pilih file")]),
                style={
                    "width": "100%", "height": "50px", "lineHeight": "50px", "borderWidth": "1px",
                    "borderStyle": "dashed", "borderRadius": "5px", "textAlign": "center", "marginTop": "5px",
                },
            ),
            html.Div(id="info-upload-status"),
            html.Div(id="info-mode-body"),
        ]
    )


def _parse_doc_contents(contents: str, filename: str):
    _, content_string = contents.split(",", 1)
    decoded = base64.b64decode(content_string)
    if filename.lower().endswith(".json"):
        return parse_dataiku_json(io.BytesIO(decoded))
    return parse_dataiku_doc(io.BytesIO(decoded))


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
    }
    total_cols_from_map = len({(t, c) for t, cols in table_confirmed_map.items() for c in cols})
    total_table = len(dataset_meta) or len(table_confirmed_map)
    total_kolom = sum(len(v["columns"]) for v in dataset_meta.values()) or total_cols_from_map
    status = html.Div(
        [
            html.P(f"✅ {filename} terbaca: {total_table} table, {total_kolom} kolom, {len(doc_recipes)} recipe."),
        ],
        style={"color": "green"},
    )
    return store, status


@app.callback(
    Output("info-mode-body", "children"),
    Input("info-project-dropdown", "value"),
    Input("info-doc-store", "data"),
)
def update_info_mode(sel_project, doc_store):
    if not sel_project:
        return html.P("Pilih project dulu.", style={"color": "#888"})

    proj_df = df[df["Project"] == sel_project]
    tables_for_project = sorted(proj_df["Short Table"].dropna().unique())
    table_confirmed_map = (doc_store or {}).get("table_confirmed_map", {})

    if not table_confirmed_map:
        return html.Div(
            [
                html.P(
                    "Belum ada dokumen/dump di-upload (atau belum ada recipe yang confirmed "
                    "makai table project ini) — upload dulu buat lihat kolom yang dipakai di Dataiku.",
                    style={"color": "#888"},
                ),
            ]
        )

    frames = []
    for t in tables_for_project:
        union_df = build_table_union(proj_df, table_confirmed_map, t)
        if union_df.empty:
            continue
        union_df.insert(0, "Short Table", t)
        frames.append(union_df)

    if not frames:
        return html.P(
            "Nggak ada kolom project ini yang confirmed dipakai di dokumen/dump yang di-upload.",
            style={"color": "#888"},
        )

    combined = pd.concat(frames, ignore_index=True)
    table_out = dash_table.DataTable(
        id="info-union-table",
        data=combined.to_dict("records"),
        columns=[{"name": c, "id": c, "editable": c == "Pilih"} for c in combined.columns],
        row_selectable=False,
        editable=True,
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
                "Centang (edit kolom 'Pilih' jadi True/False) baris yang mau dibawa, lalu klik Go Compare.",
                style={"color": "#666"},
            ),
            table_out,
            html.Button("✅ Go Compare", id="info-go-btn", n_clicks=0, style={"marginTop": "15px"}),
            html.Div(id="info-save-status"),
        ]
    )


@app.callback(
    Output("info-save-status", "children"),
    Input("info-go-btn", "n_clicks"),
    State("info-union-table", "data"),
    State("info-project-dropdown", "value"),
    prevent_initial_call=True,
)
def go_compare(n_clicks, table_data, sel_project):
    if not n_clicks or not table_data:
        raise PreventUpdate

    picked = pd.DataFrame(table_data)
    picked = picked[picked["Pilih"].isin([True, "True", "true"])]
    if picked.empty:
        return html.P("Nggak ada baris yang dicentang (Pilih=True).", style={"color": "orange"})

    proj_df = df[df["Project"] == sel_project]
    frames = []
    for t, group in picked.groupby("Short Table"):
        recipe_map = dict(zip(group["Column DWH"], group["Dipakai di Recipe"]))
        subset = proj_df[
            (proj_df["Short Table"] == t) & (proj_df["Column DWH"].isin(group["Column DWH"]))
        ].copy()
        subset["Dipakai di Recipe"] = subset["Column DWH"].map(recipe_map)
        frames.append(subset)

    new_rows = pd.concat(frames, ignore_index=True) if frames else pd.DataFrame(columns=DISPLAY_COLUMNS)
    try:
        combined, unique_df, compile_df = save_project_selection(sel_project, new_rows)
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
