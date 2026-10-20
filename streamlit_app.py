"""Gojek Splitter - versi web (Streamlit) dengan CRUD.

Upload PDF riwayat transaksi Gojek -> periksa & sesuaikan data (ubah / tambah / hapus baris)
-> pilih layanan -> unduh PDF/Excel per layanan.
File diproses di memori dan tidak disimpan oleh aplikasi ini.
"""
import hashlib
import io
import re
import zipfile
from collections import OrderedDict

import pandas as pd
import streamlit as st

import gojek_edit as ge
import gojek_split as gs

st.set_page_config(page_title="Gojek Splitter", page_icon="🧾", layout="wide", initial_sidebar_state="collapsed")

# --- UI/UX: clean, simple, responsive visual system ---
st.markdown("""
<style>
  :root { --ink:#172033; --muted:#64748b; --line:#e5eaf1; --brand:#087f5b; --brand-dark:#066649; --surface:#ffffff; --canvas:#f6f8fb; }
  .stApp { background:var(--canvas); color:var(--ink); }
  .block-container { max-width: 1440px; padding-top: 2rem; padding-bottom: 3rem; }
  header[data-testid="stHeader"] { background:transparent; }
  h1,h2,h3 { color:var(--ink); letter-spacing:-.025em; }
  h1 { font-size:2.15rem !important; font-weight:750 !important; margin-bottom:.2rem !important; }
  h2 { font-size:1.35rem !important; font-weight:700 !important; }
  h3 { font-size:1.05rem !important; font-weight:700 !important; }
  [data-testid="stCaptionContainer"] { color:var(--muted); }
  .hero { background:linear-gradient(120deg,#ffffff 0%,#f0faf6 100%); border:1px solid #dcefe7; border-radius:18px; padding:1.35rem 1.5rem; margin-bottom:1.25rem; }
  .hero-kicker { color:var(--brand); font-size:.76rem; font-weight:800; text-transform:uppercase; letter-spacing:.12em; margin-bottom:.45rem; }
  .hero-title { color:var(--ink); font-size:1.8rem; font-weight:800; letter-spacing:-.04em; line-height:1.2; }
  .hero-copy { color:#586579; font-size:.96rem; margin-top:.45rem; max-width:780px; line-height:1.6; }
  .step-row { display:flex; gap:.55rem; flex-wrap:wrap; margin:.75rem 0 1.25rem; }
  .step-chip { background:#fff; color:#5e6b7e; border:1px solid var(--line); border-radius:999px; padding:.45rem .8rem; font-size:.82rem; font-weight:650; }
  .step-chip.active { background:#e6f6ef; color:#066649; border-color:#b9e6d2; }
  div[data-testid="stFileUploader"] { background:#fff; border:1px dashed #b8c8d5; border-radius:14px; padding:1rem; }
  div[data-testid="stFileUploader"] section { border:0; }
  div[data-testid="stFileUploader"] button { border-radius:9px; }
  div[data-testid="stMetric"] { background:var(--surface); border:1px solid var(--line); border-radius:14px; padding:1rem 1.05rem; box-shadow:0 2px 7px rgba(15,23,42,.025); }
  div[data-testid="stMetricLabel"] { color:var(--muted); font-size:.82rem; }
  div[data-testid="stMetricValue"] { color:var(--ink); font-size:1.35rem; font-weight:750; }
  div[data-testid="stAlert"] { border-radius:12px; }
  div[data-testid="stExpander"] { background:#fff; border:1px solid var(--line); border-radius:12px; overflow:hidden; }
  div[data-testid="stExpander"] details summary { font-weight:650; }
  .section-intro { color:var(--muted); font-size:.9rem; margin-top:-.55rem; margin-bottom:.9rem; }
  div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] { border:1px solid var(--line); border-radius:12px; overflow:hidden; }
  .stButton button, .stDownloadButton button { border-radius:10px; font-weight:650; min-height:2.6rem; }
  .stButton button[kind="primary"], .stDownloadButton button[kind="primary"] { background:var(--brand); border-color:var(--brand); }
  .stButton button[kind="primary"]:hover, .stDownloadButton button[kind="primary"]:hover { background:var(--brand-dark); border-color:var(--brand-dark); }
  div[data-baseweb="select"] > div, div[data-baseweb="input"] > div { border-radius:9px; }
  hr { border-color:var(--line); }
  @media (max-width: 768px) {
    .block-container { padding:1rem .8rem 2rem; }
    h1 { font-size:1.7rem !important; }
    .hero { padding:1rem; border-radius:14px; }
    .hero-title { font-size:1.45rem; }
    div[data-testid="stMetric"] { padding:.8rem; }
    div[data-testid="stMetricValue"] { font-size:1.1rem; }
  }
</style>
""", unsafe_allow_html=True)


DEFAULT_ON = {"gofood", "gosend"}


@st.cache_data(show_spinner="Membaca PDF...")
def parse(pdf_bytes: bytes):
    return gs.parse_pdf(io.BytesIO(pdf_bytes))


st.markdown("""
<div class="hero">
  <div class="hero-kicker">Pengelolaan laporan transaksi</div>
  <div class="hero-title">Gojek Splitter</div>
  <div class="hero-copy">Pisahkan transaksi berdasarkan layanan, periksa data dengan mudah, lalu unduh laporan yang siap digunakan. File diproses di memori aplikasi.</div>
</div>
<div class="step-row">
  <span class="step-chip active">1 · Unggah PDF</span>
  <span class="step-chip">2 · Periksa data</span>
  <span class="step-chip">3 · Pilih layanan</span>
  <span class="step-chip">4 · Unduh laporan</span>
</div>
""", unsafe_allow_html=True)

with st.expander("Catatan penggunaan", expanded=False):
    st.markdown(
        "- Unggah PDF **Riwayat transaksi** yang diunduh dari aplikasi Gojek.\n"
        "- Tabel di bawah **bisa diedit**: klik sel untuk mengubah, tombol **+** di bawah tabel untuk menambah baris, "
        "centang kotak di kiri baris lalu tekan ikon tempat sampah untuk menghapus.\n"
        "- Jika data diubah, laporan PDF otomatis diberi **catatan penyesuaian manual** dan tersedia **log perubahan**.\n"
        "- PDF berisi data pribadi (nomor telepon, email, alamat). Jangan unggah di aplikasi yang dapat diakses umum "
        "tanpa persetujuan atasan/IT.\n"
        "- Hasil adalah PDF buatan ulang, **bukan dokumen resmi Gojek**. Untuk audit/klaim, gunakan bukti satuan dari aplikasi Gojek."
    )

st.markdown("### 1. Unggah riwayat transaksi")
st.markdown('<div class="section-intro">Pilih file PDF riwayat transaksi yang diunduh dari aplikasi Gojek.</div>', unsafe_allow_html=True)
up = st.file_uploader("Pilih file PDF", type=["pdf"], help="Format yang didukung: PDF riwayat transaksi Gojek.")
if up is None:
    st.info("Unggah PDF untuk mulai memeriksa dan memisahkan transaksi.")
    st.stop()

raw = up.getvalue()
try:
    info0, records = parse(raw)
except Exception as e:  # noqa: BLE001
    st.error(f"File tidak bisa dibaca sebagai riwayat transaksi Gojek. Detail: {e}")
    st.stop()
if not records:
    st.warning("Tidak ada transaksi yang terbaca dari PDF ini.")
    st.stop()

# --- state per file ---
fid = hashlib.md5(raw).hexdigest()[:12]
if st.session_state.get("fid") != fid:
    st.session_state.update(fid=fid, ver=0, orig=ge.records_to_df(records))
orig: pd.DataFrame = st.session_state["orig"]
header_total = info0.get("total_header")

# --- info akun & periode (bisa diedit) ---
with st.expander("Data akun & periode (bisa diedit)", expanded=False):
    a, b = st.columns(2)
    info = dict(info0)
    info["nama"] = a.text_input("Nama akun", info0.get("nama", ""), key=f"nama_{fid}")
    info["telp"] = b.text_input("No. telepon", info0.get("telp", ""), key=f"telp_{fid}")
    info["email"] = a.text_input("Email", info0.get("email", ""), key=f"email_{fid}")
    info["periode"] = b.text_input("Periode", info0.get("periode", ""), key=f"periode_{fid}")
info_changed = any(info[k] != info0.get(k, "") for k in ("nama", "telp", "email", "periode"))

# --- editor CRUD ---
st.markdown("### 2. Periksa dan sesuaikan transaksi")
cc1, cc2 = st.columns([5, 1])
cc1.markdown('<div class="section-intro">Klik sel untuk mengedit. Gunakan baris kosong di bagian bawah tabel untuk menambah transaksi.</div>', unsafe_allow_html=True)
if cc2.button("↺ Reset ke data asli", width="stretch"):
    st.session_state["ver"] += 1
    st.rerun()

service_opts = sorted(set(orig["Layanan"].dropna()) | set(ge.COMMON_SERVICES))
edited = st.data_editor(
    orig, key=f"editor_{fid}_{st.session_state['ver']}", num_rows="dynamic", hide_index=True, width="stretch",
    column_config={
        "Tanggal": st.column_config.TextColumn("Tanggal", help="Format dd/mm/yyyy", required=True, width="small"),
        "Jam": st.column_config.TextColumn("Jam", width="small"),
        "No. Transaksi": st.column_config.TextColumn("No. Transaksi", width="medium"),
        "Layanan": st.column_config.SelectboxColumn("Layanan", options=service_opts, required=True, width="small"),
        "Tier": st.column_config.TextColumn("Tier", width="small"),
        "Dari": st.column_config.TextColumn("Dari", width="medium"),
        "Alamat Dari": st.column_config.TextColumn("Alamat Dari", width="large"),
        "Tujuan": st.column_config.TextColumn("Tujuan", width="medium"),
        "Alamat Tujuan": st.column_config.TextColumn("Alamat Tujuan", width="large"),
        "Metode Bayar": st.column_config.TextColumn("Metode Bayar", width="small"),
        "Total": st.column_config.NumberColumn("Total (Rp)", min_value=0, step=1, format="%d", required=True),
    })

# --- validasi & ringkasan ---
errors, warnings = ge.validate(edited)
log = ge.change_log(orig, edited)
n_mod, n_add, n_del = ge.summarize(log)
edited_flag = (not log.empty) or info_changed
grand = int(pd.to_numeric(edited["Total"], errors="coerce").fillna(0).sum())

m1, m2, m3, m4 = st.columns(4)
m1.metric("Periode", info.get("periode") or "-")
m2.metric("Jumlah transaksi", len(edited), delta=len(edited) - len(orig) if len(edited) != len(orig) else None)
m3.metric("Total semua transaksi", gs.rp(grand))
if header_total is not None:
    diff = grand - header_total
    m4.metric("Selisih vs total PDF asli", gs.rp(abs(diff)) if diff else "Rp0",
              delta=None if diff == 0 else ("lebih besar" if diff > 0 else "lebih kecil"), delta_color="off")

if header_total is None:
    st.warning("Total di header PDF tidak ditemukan. Cek hasil secara manual.")
elif grand == header_total:
    st.success(f"Total {gs.rp(grand)} sama dengan total di header PDF asli.")
elif log.empty:
    st.error(f"Total terbaca {gs.rp(grand)} berbeda dari total header PDF {gs.rp(header_total)}. "
             f"Kemungkinan ada baris yang salah dibaca - periksa tabel di atas.")
else:
    st.info(f"Total sekarang {gs.rp(grand)}, berbeda dari total PDF asli {gs.rp(header_total)} "
            f"karena data sudah diubah.")

for m in errors:
    st.error(m)
for m in warnings:
    st.warning(m)

if edited_flag:
    parts = []
    if n_mod:
        parts.append(f"{n_mod} baris diubah")
    if n_add:
        parts.append(f"{n_add} baris ditambah")
    if n_del:
        parts.append(f"{n_del} baris dihapus")
    if info_changed:
        parts.append("data akun/periode diubah")
    st.info("Perubahan manual: " + ", ".join(parts) + ".")
    if not log.empty:
        with st.expander("Lihat log perubahan"):
            st.dataframe(log, hide_index=True, width="stretch")

if errors:
    st.stop()  # laporan tidak dibuat sebelum data valid
if edited.empty:
    st.info("Tabel kosong. Tambahkan transaksi atau reset ke data asli.")
    st.stop()

# --- pilih layanan & unduh ---
recs = ge.df_to_records(edited)
groups = OrderedDict()
for r in recs:
    groups.setdefault(r["layanan"], []).append(r)

st.markdown("---")
st.markdown("### 3. Pilih layanan untuk laporan")
st.markdown('<div class="section-intro">Pilih satu atau beberapa layanan. Setiap layanan akan dibuatkan file PDF tersendiri.</div>', unsafe_allow_html=True)
chosen = st.multiselect(
    "Layanan yang dibuat laporannya", options=list(groups.keys()),
    default=[l for l in groups if l.lower() in DEFAULT_ON],
    format_func=lambda l: f"{l} - {len(groups[l])} transaksi - {gs.rp(sum(r['total'] or 0 for r in groups[l]))}")
want_excel = st.checkbox("Sertakan file Excel", value=False)
if not chosen:
    st.info("Pilih minimal satu layanan untuk membuat laporan.")
    st.stop()

note = None
if edited_flag:
    note = "Data pada laporan ini telah disesuaikan manual dari PDF riwayat transaksi asli."

suffix = re.sub(r"^Riwayat_transaksi_Gojek_?", "", up.name.rsplit(".", 1)[0])
files = OrderedDict()
for lay in chosen:
    name = f"Riwayat_transaksi_Gojek_{lay}" + (f"_{suffix}" if suffix else "")
    files[f"{name}.pdf"] = gs.pdf_bytes(lay, groups[lay], info, note)
    if want_excel:
        files[f"{name}.xlsx"] = gs.excel_bytes(groups[lay])
if not log.empty:
    files[f"Log_perubahan_{suffix or 'hasil'}.csv"] = log.to_csv(index=False).encode("utf-8-sig")

st.markdown("---")
st.markdown("### 4. Unduh laporan")
st.markdown('<div class="section-intro">Unduh file satu per satu atau gunakan ZIP untuk mengambil semua hasil sekaligus.</div>', unsafe_allow_html=True)
MIME = {".pdf": "application/pdf", ".csv": "text/csv",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
for fname, data in files.items():
    st.download_button(f"Unduh {fname}", data=data, file_name=fname,
                       mime=MIME[fname[fname.rfind('.'):]], key=fname)

if len(files) > 1:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for fname, data in files.items():
            z.writestr(fname, data)
    st.download_button("📦 Unduh semua (ZIP)", data=buf.getvalue(),
                       file_name=f"Riwayat_transaksi_Gojek_{suffix or 'hasil'}.zip", mime="application/zip")