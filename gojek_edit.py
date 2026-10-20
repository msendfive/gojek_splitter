"""Helper CRUD untuk Gojek Splitter (versi web): konversi data <-> tabel, validasi, dan log perubahan."""
import re
from datetime import datetime

import pandas as pd

COLS = ["Tanggal", "Jam", "No. Transaksi", "Layanan", "Tier", "Dari", "Alamat Dari",
        "Tujuan", "Alamat Tujuan", "Metode Bayar", "Total"]
TEXT_COLS = [c for c in COLS if c != "Total"]
COMMON_SERVICES = ["GoFood", "GoSend", "GoRide", "GoCar", "GoShop", "GoMart", "GoBox"]


def _s(v) -> str:
    if v is None or (isinstance(v, float) and pd.isna(v)) or v is pd.NA:
        return ""
    return str(v).strip()


def _total(v):
    if v is None or v is pd.NA or (isinstance(v, float) and pd.isna(v)):
        return None
    try:
        return int(v)
    except (TypeError, ValueError):
        return None


def records_to_df(records) -> pd.DataFrame:
    df = pd.DataFrame([{
        "Tanggal": r["tanggal"], "Jam": r["jam"], "No. Transaksi": r["no"], "Layanan": r["layanan"],
        "Tier": r["tier"], "Dari": r["dari"][0], "Alamat Dari": r["dari"][1],
        "Tujuan": r["tujuan"][0], "Alamat Tujuan": r["tujuan"][1],
        "Metode Bayar": r["metode"], "Total": r["total"]} for r in records], columns=COLS)
    df["Total"] = df["Total"].astype("Int64")
    return df


def df_to_records(df: pd.DataFrame):
    out = []
    for _, row in df.iterrows():
        out.append({
            "tanggal": _s(row["Tanggal"]), "jam": _s(row["Jam"]), "no": _s(row["No. Transaksi"]),
            "layanan": _s(row["Layanan"]), "tier": _s(row["Tier"]),
            "dari": (_s(row["Dari"]), _s(row["Alamat Dari"])),
            "tujuan": (_s(row["Tujuan"]), _s(row["Alamat Tujuan"])),
            "metode": _s(row["Metode Bayar"]), "total": _total(row["Total"])})
    return out


def validate(df: pd.DataFrame):
    """Return (errors, warnings). Error memblokir pembuatan laporan; warning hanya peringatan."""
    errors, warnings = [], []
    seen = {}
    for pos, (_, row) in enumerate(df.iterrows(), start=1):
        tgl = _s(row["Tanggal"])
        try:
            if not re.fullmatch(r"\d{2}/\d{2}/\d{4}", tgl):
                raise ValueError
            datetime.strptime(tgl, "%d/%m/%Y")
        except ValueError:
            errors.append(f"Baris {pos}: Tanggal '{tgl}' tidak valid (format dd/mm/yyyy).")
        if not _s(row["Layanan"]):
            errors.append(f"Baris {pos}: Layanan kosong.")
        tot = _total(row["Total"])
        if tot is None:
            errors.append(f"Baris {pos}: Total kosong atau bukan angka.")
        elif tot < 0:
            errors.append(f"Baris {pos}: Total tidak boleh negatif.")
        no = _s(row["No. Transaksi"])
        if not no:
            warnings.append(f"Baris {pos}: No. Transaksi kosong.")
        elif no in seen:
            warnings.append(f"Baris {pos}: No. Transaksi '{no}' sama dengan baris {seen[no]} (kemungkinan duplikat).")
        else:
            seen[no] = pos
    return errors, warnings


def _norm(col, v):
    return str(_total(v)) if col == "Total" and _total(v) is not None else ("" if col == "Total" else _s(v))


def change_log(orig: pd.DataFrame, edited: pd.DataFrame) -> pd.DataFrame:
    """Daftar perubahan (Diubah / Ditambah / Dihapus) antara data asli dan hasil edit.
    Kolom Baris = nomor urut baris pada data asli (Diubah/Dihapus) atau pada data hasil edit (Ditambah)."""
    rows = []
    for idx in orig.index.difference(edited.index):
        o = orig.loc[idx]
        rows.append({"Aksi": "Dihapus", "Baris": orig.index.get_loc(idx) + 1, "No. Transaksi": _s(o["No. Transaksi"]), "Kolom": "(seluruh baris)",
                     "Nilai lama": f"{_s(o['Tanggal'])} | {_s(o['Layanan'])} | {_norm('Total', o['Total'])}",
                     "Nilai baru": ""})
    for idx in edited.index.difference(orig.index):
        n = edited.loc[idx]
        rows.append({"Aksi": "Ditambah", "Baris": edited.index.get_loc(idx) + 1, "No. Transaksi": _s(n["No. Transaksi"]), "Kolom": "(seluruh baris)",
                     "Nilai lama": "", "Nilai baru": f"{_s(n['Tanggal'])} | {_s(n['Layanan'])} | {_norm('Total', n['Total'])}"})
    for idx in edited.index.intersection(orig.index):
        o, n = orig.loc[idx], edited.loc[idx]
        for col in COLS:
            a, b = _norm(col, o[col]), _norm(col, n[col])
            if a != b:
                rows.append({"Aksi": "Diubah", "Baris": orig.index.get_loc(idx) + 1, "No. Transaksi": _s(o["No. Transaksi"]), "Kolom": col,
                             "Nilai lama": a, "Nilai baru": b})
    return pd.DataFrame(rows, columns=["Aksi", "Baris", "No. Transaksi", "Kolom", "Nilai lama", "Nilai baru"])


def summarize(log: pd.DataFrame):
    if log.empty:
        return 0, 0, 0
    mod = log[log["Aksi"] == "Diubah"]
    n_mod = mod["Baris"].nunique() if not mod.empty else 0
    return n_mod, int((log["Aksi"] == "Ditambah").sum()), int((log["Aksi"] == "Dihapus").sum())