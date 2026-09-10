# Migration Progress Explorer

Streamlit app kecil untuk menelusuri progress migrasi data AXA Mandiri
(source → Bronze → Silver Tier 1 → Silver Tier 2), dengan tiga mode:

1. **Cari by Table** — pilih satu table, lihat project mana saja yang
   memakainya, lalu lihat detail tiap kolomnya sampai ke Silver Tier 2.
2. **Cari by Project** — pilih satu project, lihat table & kolom apa saja
   yang dipakai, lalu lihat detail tiap kolomnya sampai ke Silver Tier 2.
3. **Seleksi Kolom** — PIC pilih project → pilih table relevan → centang
   kolom DWH yang bener-bener dipakai project itu. Hasilnya ditulis balik
   sebagai 3 sheet tambahan di `Checking_Progress_Migration.xlsx` sendiri
   (lihat bagian 4).

---

## 1. File yang dipakai

Cuma **satu file**: `Checking_Progress_Migration.xlsx`, dan cuma **satu
sheet** di dalamnya: **`Coretan Checking Migration`**.

Kenapa cukup satu sheet? Karena sheet ini sudah didesain sebagai satu baris
= satu kombinasi (Project, Table, Column DWH) yang membawa seluruh jejak
pipeline dari DWH sampai Silver Tier 2 dalam baris yang sama. Sheet-sheet
lain (`List Table Compile - DWH`, `List Project DA/BI/DS/DQ - DWH`, dll)
semuanya cuma agregasi/rekap dari sheet ini — jadi untuk kebutuhan "cari
detail per table/project", sumber aslinya (Coretan) sudah cukup dan paling
akurat (tidak lewat proses agregasi tambahan yang bisa basi).

### Struktur kolom yang dipakai

| Kolom di Excel | Nama di app | Keterangan |
|---|---|---|
| `Project` | Project | Nama project (dari 57 project) |
| `Data Source (Tables)` | Data Source | Path lengkap table sumber, misal `[DB_ANALYTICS].[AMFS_DWH].[dbo].[IB_BRANCH]` |
| `Phase 1` | Phase 1 | Yes/No — apakah table masuk Phase 1 |
| `Column DWH` | Column DWH | Nama kolom di level DWH |
| `Source Table in DataLake` | Bronze Table | Nama table di layer Bronze |
| `Status Column Source` | Status | `Table` / `Hardcode` / `Gap` (lihat penjelasan di bawah) |
| `Source Column In datalake` | Bronze Column | Nama kolom di layer Bronze |
| `Table Silver Tier 1 In Datalake` | Silver1 Table | Nama table di Silver Tier 1 |
| `Column Silver Tier 1 In Datalake` | Silver1 Column | Nama kolom di Silver Tier 1 |
| `Table Silver Tier 2 in datalake` | Silver2 Table | Nama table di Silver Tier 2 |
| `Silver Column Datalake Tier 2` | Silver2 Column | Nama kolom di Silver Tier 2 — kosong berarti gap di tahap ini |
| `Domain` | Domain | Application / Claim / FinAct / Party / Policy / Producer / Product |
| `Kategori Project` | Kategori Project | DA / BI / DS / DQ |
| `Formula Hardcode` | Formula Hardcode | Isi formula kalau Status = Hardcode |

**Kolom turunan** yang dibuat oleh app (bukan dari Excel langsung):

- **`Short Table`** — nama table pendek, diambil dari segmen terakhir
  `Data Source (Tables)` setelah titik terakhir, dibuang tanda kurung
  siku `[...]`-nya, lalu di-uppercase. Contoh:
  `[DB_ANALYTICS].[AMFS_DWH].[dbo].[IB_BRANCH]` → `IB_BRANCH`.
  Rule ini **persis sama** dengan yang dipakai waktu regenerasi sheet
  `List Table Compile - DWH`, supaya hasil pencarian di app ini konsisten
  dengan angka-angka di sheet itu.

### Arti status (kolom `Status`)

- 🟩 **Table** — source table & source column ketemu, kolom DWH berhasil
  dipetakan ke Bronze.
- 🟨 **Hardcode** — tidak ada source column, nilainya di-hardcode/derive
  (lihat isi `Formula Hardcode`).
- 🟥 **Gap** — source table ada tapi source column kosong, DAN tidak ada
  formula hardcode. Ini gap yang sebenarnya (lihat catatan di bawah).

> Catatan: definisi "Gap" ini penting — kolom dianggap gap **hanya** kalau
> source table+column kosong **dan** formula hardcode-nya juga kosong.
> Kalau ada formula hardcode-nya, statusnya Hardcode, bukan Gap.

---

## 2. Cara jalankan di VSCode

### a. Siapkan folder project

Taruh 3 file ini dalam satu folder:

```
migration-explorer/
├── app.py
├── requirements.txt
├── README.md
└── Checking_Progress_Migration.xlsx   <- taruh file Excel-nya di sini
```

(Kalau tidak mau taruh file Excel di folder yang sama, bisa juga upload
manual lewat tombol upload di sidebar app-nya — lihat bagian 3.)

### b. Buat virtual environment (disarankan)

Buka terminal di VSCode (`` Ctrl+` `` / `` Cmd+` ``), lalu:

```bash
python -m venv venv
source venv/bin/activate        # Windows: venv\Scripts\activate
pip install -r requirements.txt
```

### c. Jalankan app-nya

```bash
streamlit run app.py
```

Browser akan otomatis kebuka ke `http://localhost:8501`. Kalau tidak,
buka manual link yang muncul di terminal.

---

## 3. Cara pakai app-nya

**Sumber data (sidebar kiri):**
- App dikunci baca `Checking_Progress_Migration.xlsx` dari folder yang sama
  dengan `app.py` (variabel `DEFAULT_PATH` di `app.py`). Tidak ada opsi
  upload — kalau mau pakai file lain, ganti isi file ini (nama file harus
  tetap sama) atau ubah `DEFAULT_PATH` di kode.
- Tombol **"🔄 Reload data (clear cache)"** buat clear cache kalau file
  Excel-nya baru saja diupdate tapi app belum kebaca perubahannya.

**Filter (sidebar kiri):**
- Filter berdasarkan Kategori Project (DA/BI/DS/DQ) dan Status Column
  Source — berlaku ke kedua mode pencarian.

**Mode pencarian (di halaman utama):**
- **Cari by Table** — ketik sebagian nama table untuk mempersempit pilihan
  dropdown, pilih table-nya, lalu lihat daftar project pemakai + tabel
  detail kolomnya.
- **Cari by Project** — sama, tapi mulai dari nama project.
- **Seleksi Kolom** — lihat bagian 4 di bawah.

---

## 4. Mode "Seleksi Kolom" — cara kerja & output-nya

Mode ini untuk PIC tiap project menandai kolom DWH mana yang **bener-bener
dipakai** oleh project mereka (bukan sekadar yang ke-mapping di Coretan).

**Alur pemakaian:**
1. Pilih **Project**.
2. Pilih **Table** yang relevan untuk project itu (default: semua table
   yang dipakai project ini menurut Coretan, tinggal di-uncheck yang tidak
   relevan).
3. Untuk tiap table terpilih, muncul tabel kolom DWH-nya dengan checkbox
   **"Pilih"** (default semua tercentang) — uncheck kolom yang tidak
   dipakai project ini.
4. Klik **"✅ Go Selection"**.

**Output — 3 sheet/tab: `Selected Columns`, `Unique Columns per Table`,
`Compile per Table`.** Ke mana disimpannya tergantung backend yang aktif
(app otomatis pilih, lihat bagian 5):

| Sheet | Isi |
|---|---|
| `Selected Columns` | Sama seperti Coretan, tapi cuma baris (Project, Table, Column) yang sudah dicentang PIC. Numpuk dari semua project yang pernah diseleksi. |
| `Unique Columns per Table` | Union kolom unik per table dari **semua project** (digabung by table, bukan by project lagi) — kalau 2 project sama-sama pakai kolom yang sama di table yang sama, cuma muncul sekali, tapi kolom `Projects` menunjukkan project mana saja yang pakai. |
| `Compile per Table` | Hitungan per table: jumlah kolom DWH, berapa yang sampai Bronze, Silver Tier 1, Silver Tier 2, berapa Hardcode, berapa Gap — dihitung dari `Unique Columns per Table` (jadi tidak dobel-hitung kalau banyak project pakai kolom yang sama). |

**Penting — submit ulang untuk project yang sama akan MENGGANTIKAN seleksi
lama project itu**, bukan menumpuk duplikat. Jadi kalau PIC salah centang
dan mau perbaiki, tinggal ulangi seleksinya dari awal untuk project
tersebut dan klik Go Selection lagi — baris lama project itu di
`Selected Columns` otomatis kebuang, diganti yang baru. Project lain yang
sudah pernah disubmit tidak ikut kehapus. Ini berlaku sama di kedua
backend (Google Sheets maupun file lokal).

---

## 5. Backend penyimpanan Seleksi Kolom — lokal vs Google Sheets

App otomatis pilih salah satu, tergantung ada tidaknya secrets Google:

- **Ada secrets `gsheet_webapp_url` + `gsheet_webapp_token`** → pakai
  **Google Sheets** (lewat Apps Script Web App, lihat bagian 7). 3 sheet
  di atas jadi 3 tab di Google Sheet tujuan. Ini yang dipakai kalau app
  di-deploy ke hosting gratis karena storage container di hosting gratis
  itu sementara (ephemeral) — kalau hasil seleksi ditulis ke file Excel
  lokal di situ, hilang tiap container restart. Google Sheets hidup di
  luar container jadi aman.
- **Tidak ada secrets itu** → fallback ke file lokal
  `Checking_Progress_Migration.xlsx` (3 sheet tambahan + auto-backup ke
  `.bak`, seperti dijelaskan di bagian 4). Ini yang aktif kalau jalan
  lokal di laptop seperti biasa (bagian 2-3 di atas), tanpa perlu setup
  apa-apa.

Caption di mode Seleksi Kolom bakal bilang backend mana yang lagi aktif.

---

## 6. Struktur kode (`app.py`)

File `app.py` sengaja dibuat satu file (bukan dipecah banyak modul) supaya
gampang dibaca ulang. Bagian-bagiannya:

1. **Konfigurasi** — `SHEET_NAME`, `DEFAULT_PATH`, `COLUMN_MAP`,
   `DISPLAY_COLUMNS`. Kalau suatu saat header di Excel berubah nama, cukup
   update `COLUMN_MAP` di sini, tidak perlu ubah logika lain.
2. **`extract_short_table()`** — fungsi murni (tanpa Streamlit) untuk
   ekstrak nama table pendek. Bisa dites terpisah tanpa perlu jalanin app.
3. **`load_data()`** — baca Excel, rename kolom, tambah kolom `Short Table`,
   dibungkus `@st.cache_data` supaya tidak baca ulang file setiap kali
   ada interaksi di UI (cuma dibaca ulang kalau file/source berubah atau
   cache di-clear manual).
4. **`render_detail_table()`** — render tabel detail dengan warna
   berdasarkan `Status`.
5. **`build_unique_sheet()` / `build_compile_sheet()`** — logic murni
   (pandas doang, tanpa Streamlit) yang menghitung sheet "Unique Columns
   per Table" dan "Compile per Table" dari sheet "Selected Columns" —
   dipakai oleh kedua backend, jadi bisa dites terpisah.
6. **`gsheet_enabled()` / `_appscript_call()` / `load_selected_sheet_gsheet()`
   / `save_project_selection_gsheet()`** — backend Google Sheets, manggil
   Apps Script Web App (`appsscript/Code.gs`) lewat HTTP POST pakai
   `requests`.
7. **`load_selected_sheet_local()` / `save_project_selection_local()`** —
   backend file Excel lokal (fallback + backup otomatis).
8. **`save_project_selection()`** — dispatcher, dipanggil dari UI, milih
   backend gsheet atau lokal berdasarkan `gsheet_enabled()`.
9. **Bagian UI** — sidebar (sumber data + filter), lalu tiga mode di atas.

### Kalau mau extend

- **Tambah mode "Cari by Domain"**: pola-nya sama persis seperti mode
  Cari by Table/Project — filter `df_f` berdasarkan kolom `Domain`, lalu
  `render_detail_table()`.
- **Tambah nama/identitas PIC di hasil seleksi**: tinggal tambah
  `st.text_input("Nama PIC")` di mode Seleksi Kolom, masukkan ke
  `new_rows["PIC"]` sebelum dipanggil `save_project_selection()`, dan
  tambahkan `"PIC"` ke `SELECTED_SHEET_COLUMNS`.
- **Sumber data ganti ke database** (bukan Excel lagi): cukup ganti isi
  `load_data()` supaya query dari DB dan mengembalikan DataFrame dengan
  kolom yang sama seperti sekarang — bagian UI tidak perlu diubah sama
  sekali karena semuanya konsumsi dataframe `df`/`df_f` yang generik.

---

## 7. Deploy ke Streamlit Community Cloud (gratis)

Supaya orang lain bisa akses lewat URL publik. Ada 2 tahap: (a) setup
Google Sheets sebagai storage persisten buat mode Seleksi Kolom, (b) push
ke GitHub & connect ke Streamlit Community Cloud.

### a. Setup Google Sheets lewat Apps Script

Dipilih pakai **Google Apps Script** (bukan bikin project Google Cloud +
service account) karena Apps Script cuma butuh akun Google biasa — **tidak
perlu bikin project Google Cloud, tidak perlu kartu kredit sama sekali**.
Kalau sempat coba jalur Google Cloud Console dan stuck di halaman "Try
Google Cloud for free" yang minta kartu, itu memang jalur yang beda —
lewati saja, tidak perlu dipakai.

1. **Bikin Google Sheet kosong** (nama bebas, misal "Migration Explorer -
   Seleksi Kolom") di Google Drive kamu.
2. Di sheet itu, buka menu **Extensions → Apps Script**. Ini buka tab
   baru, editor script yang otomatis ke-bind ke sheet yang tadi dibuka
   (tidak perlu isi ID sheet manual).
3. Hapus isi default `Code.gs` di editor itu, ganti dengan **seluruh isi**
   file [`appsscript/Code.gs`](appsscript/Code.gs) di repo ini (copy-paste).
4. Di baris `const SECRET_TOKEN = "GANTI_DENGAN_TOKEN_RAHASIA_BEBAS_KAMU_SENDIRI";`,
   ganti dengan string rahasia bebas (semakin acak semakin aman, misal
   generate lewat `python -c "import secrets; print(secrets.token_hex(24))"`).
   Token ini yang jadi "password" biar cuma app kamu yang bisa nulis ke
   sheet ini.
5. **Deploy** → **New deployment** → klik ikon gear di sebelah "Select
   type" → pilih **Web app**. Isi:
   - Execute as: **Me**
   - Who has access: **Anyone**
     (aman meski "Anyone" — endpoint-nya tetap butuh `SECRET_TOKEN` yang
     benar buat baca/tulis, orang lain yang cuma tahu URL-nya tidak bisa
     ngapa-ngapain tanpa token itu)
6. Klik **Deploy**. Google minta "Authorize access" pertama kali (klik
   akun kamu sendiri → "Advanced" → "Go to ... (unsafe)" → "Allow" — ini
   normal buat script yang belum di-publish ke Marketplace, karena
   scriptnya punya kamu sendiri jadi aman).
7. Setelah deploy sukses, copy **Web app URL**-nya (bentuknya
   `https://script.google.com/macros/s/XXXXXXXXXXXX/exec`).

### b. Isi secrets lokal (buat tes sebelum deploy)

```bash
cd apps/streamlit-checking
cp .streamlit/secrets.toml.example .streamlit/secrets.toml
```

Buka `.streamlit/secrets.toml`, isi `gsheet_webapp_url` dengan Web app URL
dari langkah a.7, dan `gsheet_webapp_token` dengan `SECRET_TOKEN` yang sama
persis dengan yang diisi di langkah a.4. File ini sudah di-gitignore, aman
tidak ke-commit.

Jalankan `streamlit run app.py` lagi — kalau setup benar, caption di mode
Seleksi Kolom bakal bilang "Hasilnya disimpan ke **Google Sheets**". Coba
submit 1 seleksi, cek langsung muncul sebagai tab baru di Google Sheet-nya.

**Kalau nanti edit ulang `Code.gs`** (misal ganti `SECRET_TOKEN`): harus
**Deploy → Manage deployments → edit (ikon pensil) → New version → Deploy**
lagi setiap kali, karena Web app URL yang sudah di-deploy itu "beku" ke
versi script saat deploy terakhir — sekadar save di editor tidak otomatis
kepakai di URL yang sudah jalan.

### c. Push ke GitHub

Repo `riset` ini nge-gitignore `*.xlsx` secara default (karena biasanya
data mentah/besar tidak boleh di-commit). Tapi buat deploy, Streamlit
Cloud butuh `Checking_Progress_Migration.xlsx` (sumber data Coretan)
ada fisik di repo-nya. Tambahkan pengecualian di `.gitignore`:

```gitignore
!apps/streamlit-checking/Checking_Progress_Migration.xlsx
```

**Penting:** file `.streamlit/secrets.toml` (yang isinya kredensial asli)
JANGAN pernah di-commit — sudah otomatis ke-block oleh `.gitignore`, tapi
selalu double-check `git status` sebelum commit/push.

Push repo ke GitHub seperti biasa (bikin repo baru kalau belum ada, boleh
private).

### d. Connect & deploy di Streamlit Community Cloud

1. Buka [share.streamlit.io](https://share.streamlit.io/), login pakai
   akun GitHub.
2. "Create app" → pilih repo `riset` → branch `main` → **Main file path**
   isi `apps/streamlit-checking/app.py`.
3. Sebelum klik Deploy, buka "Advanced settings" → tab **Secrets** →
   tempel **isi lengkap** file `.streamlit/secrets.toml` lokal kamu
   (bukan file `.example`) ke situ. Ini cara aman naruh kredensial di
   cloud — tidak lewat git sama sekali.
4. Klik **Deploy**. Tunggu proses build (~1-2 menit), nanti dapat URL
   publik `https://<nama-app-random>.streamlit.app`.
5. **(Opsional, buat privasi)** Kalau tidak mau app-nya bisa diakses
   sembarang orang: di dashboard app → "Settings" → "Sharing" → ganti ke
   "Only specific people can view this app" dan daftarin email yang
   boleh akses. Tetap gratis.

### e. Yang perlu diingat

- Search mode (**Cari by Table/Project**) selalu baca dari
  `Checking_Progress_Migration.xlsx` yang ikut ke-push ke GitHub — kalau
  datanya update, harus di-push ulang (redeploy) biar app-nya kebaca yang
  terbaru.
- Mode **Seleksi Kolom** otomatis pakai Google Sheets di hosting ini
  (karena secrets sudah diisi) — hasil seleksi PIC aman, tidak hilang
  walau container restart/redeploy/sleep.
- Kalau nanti mau ganti/rotate `SECRET_TOKEN`: update di `Code.gs`
  (redeploy versi baru seperti dijelaskan di bagian a), lalu update juga
  `gsheet_webapp_token` di secrets Streamlit Cloud (Settings → Secrets) —
  app auto-restart dengan secrets baru, tidak perlu redeploy dari GitHub.
