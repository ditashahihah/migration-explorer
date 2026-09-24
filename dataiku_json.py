"""
Parser buat dump JSON dari `export_code.py` (dijalanin di notebook Dataiku -
lihat docstring file itu buat cara pakainya).

Semua struktur di bawah ini divalidasi langsung lawan dump asli (export_ecm.json,
211 recipe/13 tipe), bukan tebakan dari dokumentasi API - raw JSON recipe
(payload) Dataiku nyimpen kolom SPESIFIK yang dipakai tiap recipe per tipe:

- Join/VStack   -> payload["selectedColumns"] (Join: per-item ada "table"
  index ke recipe.inputs; VStack: flat list, berlaku sama ke SEMUA input
  soalnya vstack nyatuin skema yang sama). Join juga ambil kolom kondisi
  dari payload["joins"][]["on"][] (column1/column2, masing2 {"name","table"}).
- Grouping       -> payload["keys"] + payload["values"] (list of {"column":..})
- Distinct       -> payload["keys"] (list of {"column":..})
- TopN           -> payload["keys"] (list of string) + payload["retrievedColumns"]
- Pivot          -> payload["explicitIdentifiers"] + payload["otherColumns"]
  (list of string)
- Prepare/Shaker -> payload["steps"][]["params"]["columns"|"column"] - beda
  dari docx, di sini ADA nama kolomnya per step (ColumnsSelector, dst)
- Evaluation/Prediction Scoring -> payload["keptInputColumns"] KALAU
  payload["filterInputColumns"] true, selain itu pass-through (semua kolom)
- Sampling/Split/Sync -> pass-through (nggak buang kolom, cuma filter/split/
  copy baris) - confirmed_columns diisi SEMUA kolom dataset input-nya
  (dari Dataset.columns, bukan dari payload recipe)

Yang TETAP "uncertain" (genuinely nggak ada info kolom terstruktur):
- Python/kode bebas lain - nggak ada cara pasti tau kolom yang dipakai
  tanpa baca/jalanin kodenya
- Prediction Training - config model ML, fitur-nya nested di preprocessing,
  terlalu spesifik/rumit buat di-generalisasi di sini
"""

from __future__ import annotations

import json
from dataclasses import dataclass, field


@dataclass
class Dataset:
    name: str
    type: str = ""
    connection: str = ""
    columns: list = field(default_factory=list)


@dataclass
class Recipe:
    name: str
    type: str = ""
    inputs: list = field(default_factory=list)
    outputs: list = field(default_factory=list)
    # dataset_name -> set kolom yang confirmed dipakai recipe ini
    confirmed_columns: dict = field(default_factory=dict)
    prepare_step_types: list = field(default_factory=list)

    def _add_confirmed(self, dataset_name: str, cols):
        bucket = self.confirmed_columns.setdefault(dataset_name, set())
        bucket.update(c for c in cols if c)


PREPARE_LIKE_TYPES = {"prepare", "shaker"}
PASS_THROUGH_TYPES = {"sampling", "split", "sync"}


def _first_input(recipe: Recipe) -> str | None:
    return recipe.inputs[0] if recipe.inputs else None


def _extract_prepare_columns(payload: dict, recipe: Recipe) -> None:
    """Ambil nama kolom dari SEMUA step Prepare yang nyebut kolom eksplisit
    di params-nya (ColumnsSelector, FilterOnValue, dst - banyak step type
    punya field "columns"/"column" di params, bukan cuma ColumnsSelector)."""
    target = _first_input(recipe)
    if not target:
        return
    cols: list[str] = []
    for step in payload.get("steps") or []:
        if not isinstance(step, dict):
            continue
        params = step.get("params") or {}
        multi = params.get("columns")
        if isinstance(multi, list):
            cols.extend(c for c in multi if isinstance(c, str))
        single = params.get("column")
        if isinstance(single, str):
            cols.append(single)
    if cols:
        recipe._add_confirmed(target, cols)


def _computed_alias_names_per_table(payload: dict) -> dict:
    """Map index_table -> set nama kolom yang HASIL FORMULA/GREL di dalam
    virtualInputs recipe ini (mis. "OWN_MDM_ID" = strval("MDM_ID")) - itu
    alias/kolom baru yang dibikin DALAM recipe ini, BUKAN kolom asli yang
    beneran ada di source table. Dipakai buat nge-skip nama2 ini supaya
    nggak salah dianggap "kolom DWH asli yang confirmed dipakai" padahal
    nama itu sendiri baru muncul gara-gara formula di recipe ini."""
    aliases: dict = {}
    for vi in payload.get("virtualInputs") or []:
        if not isinstance(vi, dict):
            continue
        idx = vi.get("index")
        if idx is None:
            continue
        names = {
            c.get("name")
            for c in (vi.get("computedColumns") or [])
            if isinstance(c, dict) and c.get("name")
        }
        if names:
            aliases[idx] = names
    return aliases


def _extract_join_columns(payload: dict, recipe: Recipe) -> None:
    """Join: selectedColumns per-item ada index "table" ke recipe.inputs.
    Ditambah kolom kondisi join (payload["joins"][]["on"][]) - kolom yang
    dipakai buat kondisi join tapi belum tentu ikut ke output/selectedColumns."""
    aliases = _computed_alias_names_per_table(payload)

    for c in payload.get("selectedColumns") or []:
        if not isinstance(c, dict):
            continue
        name, idx = c.get("name"), c.get("table")
        if name is None or idx is None or name in aliases.get(idx, ()):
            continue
        try:
            ds_name = recipe.inputs[idx]
        except (IndexError, TypeError):
            continue
        recipe._add_confirmed(ds_name, [name])

    for j in payload.get("joins") or []:
        if not isinstance(j, dict):
            continue
        for cond in j.get("on") or []:
            if not isinstance(cond, dict):
                continue
            for key in ("column1", "column2"):
                col = cond.get(key)
                if not isinstance(col, dict):
                    continue
                name, idx = col.get("name"), col.get("table")
                if name is None or idx is None or name in aliases.get(idx, ()):
                    continue
                try:
                    ds_name = recipe.inputs[idx]
                except (IndexError, TypeError):
                    continue
                recipe._add_confirmed(ds_name, [name])


def _extract_vstack_columns(payload: dict, recipe: Recipe) -> None:
    """VStack: selectedColumns FLAT (bukan table-indexed) - berlaku sama ke
    SEMUA input, soalnya vstack nyatuin baris dari input2 yang skemanya
    (dianggap) sama. Nama yang merupakan alias/computed (virtualInputs) di
    SALAH SATU input di-skip dari semua-nya, biar aman (nggak salah atribusi
    ke table manapun)."""
    cols = [c for c in (payload.get("selectedColumns") or []) if isinstance(c, str)]
    if not cols:
        return
    aliases = _computed_alias_names_per_table(payload)
    all_alias_names: set = set().union(*aliases.values()) if aliases else set()
    cols = [c for c in cols if c not in all_alias_names]
    if not cols:
        return
    for ds_name in recipe.inputs:
        recipe._add_confirmed(ds_name, cols)


def _extract_field_columns(payload: dict, recipe: Recipe, fields: tuple[str, ...]) -> None:
    """Generic: kumpulin nama kolom dari beberapa field payload sekaligus,
    tiap field isinya list of string ATAU list of {"column": "..."} - dikenain
    ke input PERTAMA recipe (Grouping/Distinct/TopN/Pivot cuma punya 1 input)."""
    target = _first_input(recipe)
    if not target:
        return
    cols: list[str] = []
    for field in fields:
        for item in payload.get(field) or []:
            if isinstance(item, str):
                cols.append(item)
            elif isinstance(item, dict) and item.get("column"):
                cols.append(item["column"])
    if cols:
        recipe._add_confirmed(target, cols)


def _extract_ml_kept_columns(payload: dict, recipe: Recipe) -> bool:
    """Evaluation/Prediction Scoring: kalau filterInputColumns true, pakai
    keptInputColumns. Return True kalau berhasil dapet info spesifik (kalau
    False, caller fallback ke pass-through)."""
    if not payload.get("filterInputColumns"):
        return False
    kept = payload.get("keptInputColumns")
    target = _first_input(recipe)
    if not (isinstance(kept, list) and target):
        return False
    cols = [c for c in kept if isinstance(c, str)]
    if not cols:
        return False
    recipe._add_confirmed(target, cols)
    return True


def _apply_pass_through(recipe: Recipe, datasets: dict) -> None:
    """Recipe yang nggak buang kolom (cuma filter/split/copy baris) - semua
    kolom dataset INPUT-nya (dari schema di `datasets`, kalau kepetakan)
    dianggap 'dipakai'."""
    for ds_name in recipe.inputs:
        ds = datasets.get(ds_name)
        if ds and ds.columns:
            recipe._add_confirmed(ds_name, ds.columns)


def _flatten_refs(io_roles) -> list:
    names = []
    for role in (io_roles or {}).values():
        for it in role.get("items", []):
            ref = it.get("ref")
            if ref:
                names.append(ref)
    return names


def _normalize_raw(raw: dict) -> dict:
    """Terima 2 kemungkinan bentuk file:
    - format baru (export_code.py): {"datasets": {...}, "recipes": {...}}
      dengan tiap recipe udah ada "inputs"/"outputs" flat.
    - format lama (recipe-only, dari script versi sebelumnya): {nama_recipe:
      {"type","definition","payload"/"code"}, ...} langsung di top-level,
      TANPA info dataset, dan inputs/outputs masih nested di
      definition.inputs/outputs (butuh di-flatten dulu).
    Return selalu dalam bentuk {"datasets": {...}, "recipes": {...}} yang flat."""
    if "recipes" in raw or "datasets" in raw:
        return raw

    recipes = {}
    for name, entry in raw.items():
        if not isinstance(entry, dict):
            continue
        definition = entry.get("definition") or {}
        new_entry = {
            "type": entry.get("type", "unknown"),
            "inputs": _flatten_refs(definition.get("inputs")),
            "outputs": _flatten_refs(definition.get("outputs")),
        }
        if isinstance(entry.get("payload"), dict):
            new_entry["payload"] = entry["payload"]
        recipes[name] = new_entry
    return {"datasets": {}, "recipes": recipes}


def parse_dataiku_json(file_or_path) -> tuple[dict, list]:
    """Parse file JSON hasil export_code.py (format baru, ada info dataset)
    ATAU format recipe-only lama (dideteksi & dinormalisasi otomatis lewat
    _normalize_raw, tapi TANPA info dataset Type/Connection/Schema karena
    memang tidak ada di format lama itu).

    Return (datasets, recipes).
    """
    if hasattr(file_or_path, "read"):
        raw = json.load(file_or_path)
    else:
        with open(file_or_path) as f:
            raw = json.load(f)
    raw = _normalize_raw(raw)

    datasets: dict[str, Dataset] = {}
    for name, entry in (raw.get("datasets") or {}).items():
        datasets[name] = Dataset(
            name=name,
            type=entry.get("type", ""),
            connection=entry.get("connection", ""),
            columns=list(entry.get("columns") or []),
        )

    recipes: list[Recipe] = []
    for name, entry in (raw.get("recipes") or {}).items():
        inputs, outputs = entry.get("inputs"), entry.get("outputs")
        if not inputs and not outputs and entry.get("definition"):
            # dump-nya belum nge-flatten inputs/outputs sendiri (masih nested
            # di definition.inputs/outputs.<role>.items[].ref) - flatten di
            # sini biar tetap kepakai tanpa perlu generate ulang dump-nya.
            definition = entry["definition"]
            inputs = _flatten_refs(definition.get("inputs"))
            outputs = _flatten_refs(definition.get("outputs"))
        recipe = Recipe(
            name=name,
            type=entry.get("type", ""),
            inputs=list(inputs or []),
            outputs=list(outputs or []),
        )
        rtype = (recipe.type or "").lower()
        payload = entry.get("payload")

        try:
            if rtype in PREPARE_LIKE_TYPES and isinstance(payload, dict):
                recipe.prepare_step_types = [
                    s.get("type", "unknown") for s in (payload.get("steps") or []) if isinstance(s, dict)
                ]
                _extract_prepare_columns(payload, recipe)
            elif rtype == "join" and isinstance(payload, dict):
                _extract_join_columns(payload, recipe)
            elif rtype == "vstack" and isinstance(payload, dict):
                _extract_vstack_columns(payload, recipe)
            elif rtype == "grouping" and isinstance(payload, dict):
                _extract_field_columns(payload, recipe, ("keys", "values"))
            elif rtype == "distinct" and isinstance(payload, dict):
                _extract_field_columns(payload, recipe, ("keys",))
            elif rtype == "topn" and isinstance(payload, dict):
                _extract_field_columns(payload, recipe, ("keys", "retrievedColumns"))
            elif rtype == "pivot" and isinstance(payload, dict):
                _extract_field_columns(payload, recipe, ("explicitIdentifiers", "otherColumns"))
            elif rtype in {"evaluation", "prediction_scoring"}:
                got_specific = isinstance(payload, dict) and _extract_ml_kept_columns(payload, recipe)
                if not got_specific:
                    _apply_pass_through(recipe, datasets)
            elif rtype in PASS_THROUGH_TYPES:
                _apply_pass_through(recipe, datasets)
            # tipe lain (python, prediction_training, dst) - dibiarkan tanpa
            # confirmed_columns, genuinely tidak ada info kolom terstruktur.
        except Exception:
            # 1 recipe gagal di-parse detailnya jangan bikin seluruh dump
            # gagal - recipe itu tetap masuk list, cuma confirmed_columns-nya
            # kosong (diperlakukan sama kayak recipe yang genuinely uncertain).
            pass

        recipes.append(recipe)

    return datasets, recipes
