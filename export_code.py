# -*- coding: utf-8 -*-
"""Dump semua definisi recipe + info dataset (Type/Connection/Schema kolom)
satu project jadi JSON.

Dijalankan di notebook Dataiku (bukan recipe). Tidak perlu API key atau URL -
otomatis memakai sesi login yang sedang aktif.

Info dataset diambil cuma dari METADATA (get_config() + read_schema()) -
TIDAK baca/query isi data aslinya sama sekali, jadi aman & cepat walau
datasetnya besar.

Pilih cara pengambilan hasil lewat CARA di bawah:

  "unduh"  - tampilkan tautan download di output notebook. Tidak perlu bikin
             apa pun lebih dulu. Paling gampang.
  "folder" - tulis ke managed folder. Perlu dibuat dulu di Flow project yang
             SAMA dengan notebook: + DATASET -> Folder -> nama EXPORT_RECIPE.
  "cetak"  - cetak langsung ke output cell. Hanya masuk akal kalau FILTER diisi,
             karena dump lengkap all recipe dan tidak bisa
             disalin dari notebook.
"""
import base64
import json

import dataiku

CARA = "unduh"
NAMA_FOLDER = "EXPORT_RECIPE"
NAMA_BERKAS = "recipes.json"

# Kosongkan untuk mengambil semua recipe. Kalau diisi, hanya recipe yang dipilih
FILTER = [] # contoh: ["pinalti", "monthly_data"]

# Recipe berbasis kode menyimpan isinya di get_code(); sisanya di
# get_json_payload(). Dua-duanya perlu, jadi jenisnya dibedakan di sini.
CODE_RECIPE_TYPES = {
    "python", "sql_script", "shell", "r",
    "hive", "impala", "pyspark", "sparkr", "spark_scala",
}


def flatten_refs(io_roles):
    """Struktur inputs/outputs recipe Dataiku: {"role": {"items": [{"ref":
    "nama_dataset"}, ...]}, ...}. Gabung SEMUA role jadi 1 list nama dataset
    (recipe kayak Join bisa punya role tambahan buat input kedua/ketiga)."""
    names = []
    for role in (io_roles or {}).values():
        for it in role.get("items", []):
            ref = it.get("ref")
            if ref:
                names.append(ref)
    return names


client = dataiku.api_client()
project = client.get_project(dataiku.default_project_key())

recipes = project.list_recipes()
if FILTER:
    recipes = [r for r in recipes if any(f in r["name"] for f in FILTER)]
print(f"Mengambil {len(recipes)} recipe.\n")

hasil_recipe = {}
for item in recipes:
    nama = item["name"]
    jenis = item["type"]
    entry = {"type": jenis}
    try:
        settings = project.get_recipe(nama).get_settings()
        try:
            definition = settings.get_recipe_raw_definition()
            entry["definition"] = definition
            entry["inputs"] = flatten_refs(definition.get("inputs"))
            entry["outputs"] = flatten_refs(definition.get("outputs"))
        except Exception as e:
            entry["definition_error"] = str(e)
        if jenis in CODE_RECIPE_TYPES:
            try:
                entry["code"] = settings.get_code()
            except Exception as e:
                entry["code_error"] = str(e)
        else:
            try:
                entry["payload"] = settings.get_json_payload()
            except Exception as e:
                entry["payload_error"] = str(e)
        print(f"OK    - {nama} ({jenis})")
    except Exception as e:
        entry["fetch_error"] = str(e)
        print(f"GAGAL - {nama} ({jenis}): {e}")
    hasil_recipe[nama] = entry

# ------------------------------------------------------------------------
# Dataset: cuma metadata (type, connection, nama kolom schema) - TIDAK baca
# isi datanya sama sekali.
# ------------------------------------------------------------------------
datasets = project.list_datasets()
if FILTER:
    datasets = [d for d in datasets if any(f in d["name"] for f in FILTER)]
print(f"\nMengambil {len(datasets)} dataset (metadata doang, bukan isi data).\n")

hasil_dataset = {}
for item in datasets:
    nama = item["name"] if isinstance(item, dict) else item.name
    entry = {}
    try:
        ds = dataiku.Dataset(nama)
        config = ds.get_config()
        entry["type"] = config.get("type", "unknown")
        entry["connection"] = (config.get("params") or {}).get("connection", "")
        try:
            schema = ds.read_schema()
            entry["columns"] = [c.get("name") for c in schema if c.get("name")]
        except Exception as e:
            entry["schema_error"] = str(e)
        print(f"OK    - {nama} ({entry.get('type')})")
    except Exception as e:
        entry["fetch_error"] = str(e)
        print(f"GAGAL - {nama}: {e}")
    hasil_dataset[nama] = entry

hasil = {"datasets": hasil_dataset, "recipes": hasil_recipe}

teks = json.dumps(hasil, indent=2, ensure_ascii=False, default=str)
ukuran = len(teks.encode("utf-8")) / 1024 / 1024

gagal_recipe = [n for n, e in hasil_recipe.items() if "fetch_error" in e]
gagal_dataset = [n for n, e in hasil_dataset.items() if "fetch_error" in e]
print(f"\n[OK] {len(hasil_recipe)} recipe + {len(hasil_dataset)} dataset, {ukuran:.1f} MB")
if gagal_recipe:
    print(f"[!]  {len(gagal_recipe)} recipe gagal diambil: {gagal_recipe}")
if gagal_dataset:
    print(f"[!]  {len(gagal_dataset)} dataset gagal diambil: {gagal_dataset}")


if CARA == "unduh":
    from IPython.display import HTML, display

    b64 = base64.b64encode(teks.encode("utf-8")).decode("ascii")
    display(HTML(
        f'<a download="{NAMA_BERKAS}" '
        f'href="data:application/json;base64,{b64}" '
        f'style="font-size:16px">Unduh {NAMA_BERKAS} ({ukuran:.1f} MB)</a>'
    ))

elif CARA == "folder":
    folder = dataiku.Folder(NAMA_FOLDER)
    with folder.get_writer(NAMA_BERKAS) as w:
        w.write(teks.encode("utf-8"))
    print(f"[OK] ditulis ke {NAMA_FOLDER}/{NAMA_BERKAS}")

elif CARA == "cetak":
    if ukuran > 2 and not FILTER:
        print(f"\n[!]  {ukuran:.1f} MB terlalu besar untuk disalin dari notebook.")
        print("     Isi FILTER dulu, atau pakai CARA = 'unduh'.")
    else:
        print("\n\n===== JSON (salin dari sini ke bawah) =====\n")
        print(teks)

else:
    raise ValueError(f"CARA tidak dikenal: {CARA}")
