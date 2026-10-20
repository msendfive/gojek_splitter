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

st.set_page_config(page_title="Gojek Splitter", page_icon="🧾", layout="wide")

DEFAULT_ON = {"gofood", "gosend"}


@st.cache_data(show_spinner="Membaca PDF...")
def parse(pdf_bytes: bytes):
    return gs.parse_pdf(io.BytesIO(pdf_bytes))


st.title("🧾 Gojek Splitter")
st.caption("Pisahkan riwayat transaksi Gojek per layanan, periksa dan sesuaikan datanya, lalu unduh laporan PDF/Excel.")

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

up = st.file_uploader("Unggah PDF riwayat transaksi Gojek", type=["pdf"])
if up is None:
    st.info("Silakan unggah file PDF untuk memulai.")
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
st.subheader("Data transaksi")
cc1, cc2 = st.columns([5, 1])
cc1.caption("Edit langsung di tabel. Baris baru, hapus, dan ubah nilai akan tercatat di log perubahan.")
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

st.subheader("Pilih layanan")
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

st.subheader("Unduh hasil")
MIME = {".pdf": "application/pdf", ".csv": "text/csv",
        ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
for fname, data in files.items():
    st.download_button(f"⬇️ {fname}", data=data, file_name=fname,
                       mime=MIME[fname[fname.rfind('.'):]], key=fname)

if len(files) > 1:
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
        for fname, data in files.items():
            z.writestr(fname, data)
    st.download_button("📦 Unduh semua (ZIP)", data=buf.getvalue(),
                       file_name=f"Riwayat_transaksi_Gojek_{suffix or 'hasil'}.zip", mime="application/zip")