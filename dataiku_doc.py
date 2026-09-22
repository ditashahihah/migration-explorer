"""
Parser deterministik untuk "Dataiku Flow Documentation" (.docx export).

Tidak pakai AI/LLM sama sekali di sini - murni baca struktur dokumen yang
dihasilkan otomatis oleh Dataiku (Heading 1 "Datasets"/"Recipes", lalu
Heading 3 per dataset/recipe, dengan pola paragraf & tabel yang konsisten).

Yang bisa diambil PASTI (explicit di teks dokumen):
- Semua dataset & schema kolomnya (section "Datasets")
- Semua recipe & tipe-nya, input/output dataset-nya (section "Recipes")
- Kolom yang disebut eksplisit sebagai join key / group key / distinct key /
  row identifier / pivot / populate content

Yang TIDAK bisa diambil dari dokumen ini: nama kolom spesifik yang
diproses recipe tipe "Prepare" (dokumennya cuma nyebut jenis step-nya,
misal "ColumnsSelector", tanpa nama kolom). Recipe seperti ini ditandai di
`prepare_step_types` supaya pemanggil tahu ada ketidakpastian di situ.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

import docx
from docx.table import Table
from docx.text.paragraph import Paragraph


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


_DATASET_REF_RE = re.compile(r"^Dataset (.+?) \(")
_JOIN_PAIR_RE = re.compile(r"^(Left|Right|Inner|Full Outer|Cross)\s+(.+?)\s+-\s+(.+?)\s+And$")


def _iter_block_items(document):
    body = document.element.body
    for child in body.iterchildren():
        if child.tag.endswith("}p"):
            yield Paragraph(child, document)
        elif child.tag.endswith("}tbl"):
            yield Table(child, document)


def _collect_dataset_refs(items, start):
    """Baca beruntun paragraf 'List Paragraph' berisi 'Dataset X (...)' mulai
    dari index `start`. Return (list_nama_dataset, index_setelah_list)."""
    names, j = [], start
    n = len(items)
    while j < n and isinstance(items[j], Paragraph) and items[j].style.name == "List Paragraph":
        m = _DATASET_REF_RE.match(items[j].text.strip())
        if m:
            names.append(m.group(1))
        j += 1
    return names, j


def _skip_empty_paragraphs(items, j):
    """Word menyelipkan paragraf kosong di antara tabel satu-baris yang
    berurutan (batasan OOXML: 2 tabel tidak boleh langsung nempel tanpa
    paragraf di antaranya). Lewati paragraf kosong itu supaya loop
    pengumpul tabel di bawah tidak berhenti prematur."""
    n = len(items)
    while j < n and isinstance(items[j], Paragraph) and not items[j].text.strip():
        j += 1
    return j


def _collect_row_tables(items, start, header_first_cell):
    """Baca baris-baris tabel (Column Name/Type/dst) mulai `start`. Dataiku
    punya 2 varian format ekspor dokumen buat ini:
    - versi lama: beruntun tabel SATU-BARIS (boleh diselingi paragraf
      kosong), tabel pertama adalah header.
    - versi baru: SATU tabel multi-baris, baris pertamanya header.
    Deteksi otomatis dari jumlah baris tabel pertama yang ditemukan.
    Return (list_of_row_cells, index_setelahnya)."""
    j = _skip_empty_paragraphs(items, start)
    n = len(items)
    if j < n and isinstance(items[j], Table) and len(items[j].rows) > 1:
        # versi baru: satu tabel multi-baris, baris pertama = header
        rows = [
            [c.text.strip() for c in r.cells]
            for r in items[j].rows
            if r.cells and r.cells[0].text.strip() != header_first_cell
        ]
        return rows, j + 1

    rows = []
    if j < n and isinstance(items[j], Table):
        first_row = [c.text.strip() for c in items[j].rows[0].cells]
        if first_row[:1] == [header_first_cell]:
            j += 1
    while True:
        j = _skip_empty_paragraphs(items, j)
        if j >= n or not isinstance(items[j], Table):
            break
        cells = [c.text.strip() for c in items[j].rows[0].cells]
        if cells and cells[0] and cells[0] != header_first_cell:
            rows.append(cells)
        j += 1
    return rows, j


def _collect_step_type_tables(items, start):
    """Baca beruntun tabel 1x2 (Type/Comment). Return (list_step_type, index)."""
    rows, j = _collect_row_tables(items, start, "Type")
    return [r[0] for r in rows], j


def _collect_schema_columns(items, start):
    """Baca beruntun tabel 1x3 (Column Name/Column Type/Description).
    Return (list_nama_kolom, index)."""
    rows, j = _collect_row_tables(items, start, "Column Name")
    return [r[0] for r in rows], j


def _collect_join_conditions(items, start):
    """Baca beruntun tabel 1x3 (Left input/Condition/Right input).
    Return (list_of_(left_col,right_col), index_setelahnya)."""
    rows, j = _collect_row_tables(items, start, "Left input")
    pairs = [(r[0], r[2]) for r in rows if len(r) >= 3]
    return pairs, j


def parse_dataiku_doc(file_or_path) -> tuple[dict, list]:
    """Parse file .docx Dataiku Flow Documentation.

    Return (datasets, recipes):
      datasets: dict nama_dataset -> Dataset
      recipes: list[Recipe]
    """
    document = docx.Document(file_or_path)
    items = list(_iter_block_items(document))
    n = len(items)

    datasets: dict[str, Dataset] = {}
    recipes: list[Recipe] = []

    section = None
    current_dataset = None
    current_recipe = None

    i = 0
    while i < n:
        it = items[i]
        if not isinstance(it, Paragraph):
            i += 1
            continue

        style = it.style.name
        text = it.text.strip()

        if style == "Heading 1":
            section = "datasets" if text == "Datasets" else "recipes" if text == "Recipes" else None
            i += 1
            continue

        if section == "datasets" and style == "Heading 3":
            current_dataset = Dataset(name=text)
            datasets[text] = current_dataset
            i += 1
            continue

        if section == "datasets" and current_dataset is not None:
            if text.startswith("Type:"):
                current_dataset.type = text.split(":", 1)[1].strip()
            elif text.startswith("Connection:"):
                current_dataset.connection = text.split(":", 1)[1].strip()
            elif text == "Schema details of the dataset.":
                cols, j = _collect_schema_columns(items, i + 1)
                current_dataset.columns = cols
                i = j
                continue

        if section == "recipes" and style == "Heading 3":
            current_recipe = Recipe(name=text)
            recipes.append(current_recipe)
            i += 1
            continue

        if section == "recipes" and current_recipe is not None:
            if text.startswith("Type:"):
                current_recipe.type = text.split(":", 1)[1].strip()
            elif text == "List of the inputs of the recipe:":
                names, j = _collect_dataset_refs(items, i + 1)
                current_recipe.inputs = names
                i = j
                continue
            elif text == "List of the outputs of the recipe:":
                names, j = _collect_dataset_refs(items, i + 1)
                current_recipe.outputs = names
                i = j
                continue
            elif text == "Prepare Steps":
                steps, j = _collect_step_type_tables(items, i + 1)
                current_recipe.prepare_step_types = steps
                i = j
                continue
            elif style == "Heading 4" and text == "Join":
                # Join dgn >2 input didokumentasikan sbg beberapa pasang
                # "Left X - Y    And" berurutan, tiap pasang diikuti tabel
                # kondisinya sendiri. Nama dataset di teks ini yang jadi
                # acuan (bukan urutan `inputs`, karena urutannya bisa beda).
                j = i + 1
                while j < n and isinstance(items[j], Paragraph):
                    pair_m = _JOIN_PAIR_RE.match(items[j].text.strip())
                    if not pair_m:
                        break
                    left_name, right_name = pair_m.group(2).strip(), pair_m.group(3).strip()
                    pairs, j = _collect_join_conditions(items, j + 1)
                    for left_col, right_col in pairs:
                        current_recipe._add_confirmed(left_name, [left_col])
                        current_recipe._add_confirmed(right_name, [right_col])
                i = j
                continue
            elif text.startswith(("Group keys:", "Distinct keys:", "Row identifier:", "Pivots:", "Populate content:")):
                cols = [c.strip() for c in text.split(":", 1)[1].split(",") if c.strip()]
                for ds in current_recipe.inputs:
                    current_recipe._add_confirmed(ds, cols)

        i += 1

    return datasets, recipes


def build_project_column_report(datasets: dict, recipes: list, coretan_short_tables: set) -> dict:
    """Ringkas hasil parse jadi laporan per-table yang siap dipakai buat
    pre-fill Seleksi Kolom.

    Return dict: short_table -> {
        "dataiku_name": nama dataset asli di Dataiku,
        "confirmed_columns": set kolom yang confirmed dipakai (dari join/group/
            distinct/pivot key),
        "uncertain": bool - True kalau ada recipe Prepare yang nyentuh dataset
            ini (artinya mungkin ada kolom lain yang dipakai tapi tidak
            tercatat nama-nya di dokumen),
    }
    Cuma dataset yang namanya PERSIS cocok dengan salah satu Short Table di
    Coretan yang dimasukkan (supaya tidak salah mapping).
    """
    report = {}
    for name, ds in datasets.items():
        if name not in coretan_short_tables:
            continue
        confirmed = set()
        uncertain = False
        for r in recipes:
            if name in r.inputs or name in r.outputs:
                if name in r.confirmed_columns:
                    confirmed.update(r.confirmed_columns[name])
                if r.type == "Prepare":
                    uncertain = True
        report[name] = {
            "dataiku_name": name,
            "confirmed_columns": confirmed,
            "uncertain": uncertain,
        }
    return report
