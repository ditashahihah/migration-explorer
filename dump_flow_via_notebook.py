"""Export dataset schema + recipe definisi dari 1 project Dataiku DSS ke satu
file JSON, sebagai alternatif dari export "Dataiku Flow Documentation" (.docx)
buat app ini (dataiku_json.py yang baca hasilnya).

KENAPA ADA INI: export .docx Flow Documentation butuh library `python-docx`
buat di-parse, dan itu TIDAK tersedia di Snowflake Anaconda channel (dicek via
INFORMATION_SCHEMA.PACKAGES) - jadi kalau app ini di-deploy ke Streamlit in
Snowflake (warehouse runtime), upload .docx nggak akan bisa dipakai. Dump JSON
ini cuma butuh library `dataiku` bawaan notebook DSS, jadi portable ke mana
aja app-nya nanti di-hosting.

HOW TO RUN:
1. Buka project Dataiku yang mau di-dump di browser.
2. Buka menu "Notebooks" -> buat notebook Python baru (Jupyter, kernel
   Python 3) di project itu -- notebook DSS sudah otomatis punya akses
   `dataiku` API ke project ini, tidak perlu API key manual.
3. Copy-paste seluruh isi file ini ke satu cell, lalu Run.
4. Download file `flow_dump.json` dari file browser notebook, lalu upload
   file itu ke app Migration Progress Explorer (mode Seleksi Kolom).

Ini cuma MEMBACA definisi dataset & recipe (tidak mengubah apa pun di project).

Defensif dengan sengaja: kalau satu dataset/recipe/field gagal diambil,
errornya dicatat di key "*_error", bukan bikin seluruh dump gagal -- jadi
project yang isinya campur-campur (dataset belum di-build, tipe recipe yang
jarang dipakai, dst) tetap menghasilkan file yang bisa dipakai sebagian.
"""
import json

import dataiku

# ---- Sesuaikan kalau perlu -------------------------------------------
PROJECT_KEY = None          # None = project tempat notebook ini jalan
OUTPUT_PATH = "flow_dump.json"
# ------------------------------------------------------------------------


def _item_name(item):
    """list_datasets()/list_recipes() pernah mengembalikan dict polos atau
    objek ringan di versi DSS yang berbeda -- tangani dua-duanya."""
    return item["name"] if isinstance(item, dict) else item.name


def _dataset_obj(project, name):
    return project.get_dataset(name) if hasattr(project, "get_dataset") else dataiku.Dataset(name)


def dump_datasets(project):
    result = {}
    error_count = 0
    for item in project.list_datasets():
        name = _item_name(item)
        entry = {}
        try:
            ds = _dataset_obj(project, name)
            config = ds.get_config()
            entry["type"] = config.get("type", "unknown")
            # Nama key connection beda-beda antar tipe dataset (SQL: "connection",
            # filesystem: bisa nggak ada sama sekali) - ambil dari "params" kalau ada.
            entry["connection"] = (config.get("params") or {}).get("connection", "")
        except Exception as e:
            entry["config_error"] = str(e)

        try:
            schema = ds.read_schema()
            entry["columns"] = [c.get("name") for c in schema if c.get("name")]
        except Exception as e:
            entry["schema_error"] = str(e)
            error_count += 1

        result[name] = entry
    return result, error_count


def _flatten_refs(io_roles):
    """Struktur inputs/outputs recipe Dataiku: {"role_name": {"items": [{"ref":
    "dataset_name"}, ...]}, ...}. Gabung SEMUA role (bukan cuma "main") jadi 1
    list nama dataset, soalnya recipe kayak Join bisa punya role tambahan
    buat input kedua/ketiga."""
    names = []
    for role in (io_roles or {}).values():
        for it in role.get("items", []):
            ref = it.get("ref")
            if ref:
                names.append(ref)
    return names


CODE_TYPES = {
    "python", "sql_script", "shell", "r", "hive", "impala",
    "pyspark", "sparkr", "spark_scala",
}


def dump_recipes(project):
    result = {}
    error_count = 0
    for item in project.list_recipes():
        name = _item_name(item)
        rtype = item.get("type", "unknown") if isinstance(item, dict) else getattr(item, "type", "unknown")
        entry = {"type": rtype}

        try:
            settings = project.get_recipe(name).get_settings()
        except Exception as e:
            entry["fetch_error"] = str(e)
            result[name] = entry
            error_count += 1
            continue

        try:
            definition = settings.get_recipe_raw_definition()
            entry["inputs"] = _flatten_refs(definition.get("inputs"))
            entry["outputs"] = _flatten_refs(definition.get("outputs"))
        except Exception as e:
            entry["definition_error"] = str(e)
            error_count += 1

        if rtype not in CODE_TYPES:
            # Recipe visual (join/group/distinct/prepare/dst) - payload-nya yang
            # nyimpen join key / group key / step Prepare, bukan definisi mentah.
            try:
                entry["payload"] = settings.get_json_payload()
            except Exception as e:
                entry["payload_error"] = str(e)
                error_count += 1

        result[name] = entry
    return result, error_count


if __name__ == "__main__":
    client = dataiku.api_client()
    project = client.get_project(PROJECT_KEY) if PROJECT_KEY else dataiku.Project()

    datasets, ds_errors = dump_datasets(project)
    recipes, rc_errors = dump_recipes(project)

    dump = {"datasets": datasets, "recipes": recipes}
    with open(OUTPUT_PATH, "w") as f:
        json.dump(dump, f, indent=2, default=str)

    total_errors = ds_errors + rc_errors
    print(f"Dumped {len(datasets)} dataset & {len(recipes)} recipe ke {OUTPUT_PATH}.")
    if total_errors:
        print(
            f"{total_errors} field gagal diambil -- cek key '*_error' per dataset/"
            f"recipe di JSON-nya. Entry itu tetap ke-dump, cuma sebagian datanya kosong."
        )
    print("Download file ini dari file browser notebook, lalu upload ke app Migration Progress Explorer.")
