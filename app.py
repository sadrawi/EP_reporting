"""
Analisis Aktivitas MBKM & Mata Kuliah Konversi - aplikasi Streamlit.

Jalankan:
    pip install streamlit pandas beautifulsoup4 lxml openpyxl xlrd reportlab matplotlib python-docx
    streamlit run app.py

Unggah file ekspor SIAKAD (.xls yang sebenarnya HTML, atau .xlsx). Aplikasi
menormalkan data (1 baris per aktivitas), memberi flag kualitas data, memeriksa
NIM, menyediakan rekap per Jenis Aktivitas / Program Studi, laporan naratif
bergrafik (Bahasa Indonesia / English; PDF & Word), serta unduhan PDF, Excel & CSV.

Normalisasi: ekspor SIAKAD berisi 1 baris per (aktivitas x MK konversi). Baris
dikelompokkan per NIM + atribut aktivitas. Kolom dosen TIDAK dipakai sebagai
kunci karena urutan nama dosen bisa berbeda antar baris MK pada aktivitas yang
sama; daftar dosen digabung (unik, urut ID dosen) per aktivitas.
"""

import io
import os
import re
import sys
import pandas as pd
import matplotlib
if "matplotlib.pyplot" not in sys.modules:
    matplotlib.use("Agg")  # hanya sekali; saat Streamlit rerun backend tidak diganti lagi
import matplotlib.pyplot as plt
from matplotlib.backends.backend_agg import FigureCanvasAgg
from matplotlib.ticker import MaxNLocator
import streamlit as st
from bs4 import BeautifulSoup

# ---------------------------------------------------------------- konstanta
JENIS_SHORT = {
    "Magang/Praktik Kerja (Kampus Merdeka)": "Magang",
    "Penelitian/Riset (Kampus Merdeka)": "Riset",
    "Kegiatan Wirausaha (Kampus Merdeka)": "Wirausaha",
    "Pertukaran Pelajar (Kampus Merdeka)": "Pertukaran",
    "Studi/Proyek Independen (Kampus Merdeka)": "Studi Independen",
}
JENIS_PERTUKARAN = "Pertukaran Pelajar (Kampus Merdeka)"
FLAG_COLS = [
    "F1 Tanpa MK", "F2 MK Berlebih", "F3 Tanpa Pembimbing",
    "F4 Tanpa Penguji", "F5 Mitra Tidak Valid", "F6 Tanpa MoU", "F7 NIM Ganda",
]
RECORD_FLAGS = ["F1 Tanpa MK", "F2 MK Berlebih", "F3 Tanpa Pembimbing", "F5 Mitra Tidak Valid"]

KEYS = [
    "NIM", "Nama", "Program Studi", "Periode Akademik", "Jenis Aktivitas",
    "Mitra", "Status Mitra", "Status Aktivitas", "Tanggal Mulai",
    "Tanggal Selesai", "Posisi", "Judul Aktivitas", "Dosen Pembimbing",
    "Dosen Penguji",
]
LECTURER_COLS = ["Dosen Pembimbing", "Dosen Penguji"]
GROUP_KEYS = []
for _key in KEYS:
    if _key not in LECTURER_COLS:
        GROUP_KEYS.append(_key)
MK_COLS = ["Kode MK Konversi", "Nama MK Konversi", "SKS"]
NIM_COLS = ["NIM", "Nama", "Program Studi", "Jml Aktivitas", "Aktivitas (Jenis - Status)", "MK Sama"]
NAVY = "#1F3864"
RED = "#7B2D26"

# Warna grafik laporan naratif (palet tervalidasi aman buta warna, urutan slot tetap).
JENIS_COLORS = {
    "Magang/Praktik Kerja (Kampus Merdeka)": "#2a78d6",
    "Penelitian/Riset (Kampus Merdeka)": "#eb6834",
    "Pertukaran Pelajar (Kampus Merdeka)": "#1baf7a",
    "Studi/Proyek Independen (Kampus Merdeka)": "#eda100",
    "Kegiatan Wirausaha (Kampus Merdeka)": "#e87ba4",
}
EXTRA_COLORS = ["#008300", "#4a3aa7", "#e34948"]
STATUS_ORDER = ["Selesai", "Evaluasi", "Diajukan", "Ditolak", "Dibatalkan"]
STATUS_COLORS = {"Selesai": "#184f95", "Evaluasi": "#3987e5", "Diajukan": "#86b6ef",
                 "Ditolak": "#898781", "Dibatalkan": "#c3c2b7"}
CAT_MOU = "Eksternal, sudah MoU"
CAT_NOMOU = "Eksternal, belum MoU"
CAT_INT = "Internal i3L"
CAT_BAD = "Mitra tidak valid"
PARTNER_ORDER = [CAT_MOU, CAT_NOMOU, CAT_INT, CAT_BAD]
PARTNER_COLORS = {CAT_MOU: "#008300", CAT_NOMOU: "#4a3aa7", CAT_INT: "#c3c2b7", CAT_BAD: "#e34948"}
INTERNAL_RE = r"\bi3l\b|indonesia international institute for life"
INK = "#0b0b0b"
INK2 = "#52514e"
MUTED = "#898781"
GRIDC = "#e1e0d9"
AXISC = "#c3c2b7"


def short_jenis(name):
    if name in JENIS_SHORT:
        return JENIS_SHORT[name]
    return str(name).replace("(Kampus Merdeka)", "").strip()


def flag_defs(mk_overload, mk_overload_px):
    return {
        "F1 Tanpa MK": "Aktivitas tanpa satu pun mata kuliah konversi (Jml MK = 0).",
        "F2 MK Berlebih": "Jml MK > %d (Pertukaran Pelajar: > %d) - diduga salah pilih / seluruh katalog terpilih."
                          % (mk_overload, mk_overload_px),
        "F3 Tanpa Pembimbing": "Kolom Dosen Pembimbing kosong atau '-'.",
        "F4 Tanpa Penguji": "Kolom Dosen Penguji kosong atau '-' (seluruh record).",
        "F5 Mitra Tidak Valid": "Mitra kosong/'-' atau berisi kode MK / teks bebas, bukan nama mitra.",
        "F6 Tanpa MoU": "Mitra berstatus 'Belum memiliki MoU kerjasama'.",
        "F7 NIM Ganda": "NIM tercatat di lebih dari 1 aktivitas (setelah filter) - cek duplikasi / pengajuan ulang.",
    }


# ---------------------------------------------------------------- parsing
def squash(text):
    return " ".join(str(text).split())


def clean_nim(value):
    text = re.sub(r"\s+", "", str(value))
    return re.sub(r"\.0$", "", text)


def sniff_format(data):
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "biff"
    if data[:2] == b"PK":
        return "xlsx"
    low = data[:2048].lower()
    if b"<html" in low or b"<!doctype html" in low or b"<table" in low:
        return "html"
    return "unknown"


def tidy_table(df):
    """Buang kolom duplikat / kolom 'No.' kedua, rapikan teks, normalkan NIM."""
    keep = []
    names = []
    seen_no = False
    for i, name in enumerate(df.columns):
        name = squash(name)
        if name == "" or name in names:
            continue
        if re.fullmatch(r"No\.?", name):
            if seen_no:
                continue
            seen_no = True
        keep.append(i)
        names.append(name)
    out = df.iloc[:, keep].copy()
    out.columns = names
    out = out.fillna("").astype(str)
    for col in out.columns:
        out[col] = out[col].map(squash)
    missing = []
    for col in GROUP_KEYS + LECTURER_COLS + MK_COLS:
        if col not in out.columns:
            missing.append(col)
    if missing:
        raise ValueError("Kolom tidak ditemukan: " + ", ".join(missing))
    out["NIM"] = out["NIM"].map(clean_nim)
    return out.reset_index(drop=True)


def read_html_report(text):
    soup = BeautifulSoup(text, "lxml")
    for tag in soup(["style", "script"]):
        tag.decompose()

    target = None
    for table in soup.find_all("table"):
        blob = table.get_text(" ", strip=True)
        if "NIM" in blob and "Judul Aktivitas" in blob:
            target = table
            break
    if target is None:
        raise ValueError("Tabel MBKM tidak ditemukan - apakah ini laporan yang benar?")

    trs = target.find_all("tr")
    header = None
    start = 0
    for i, tr in enumerate(trs):
        cells = []
        for c in tr.find_all(["td", "th"]):
            cells.append(squash(c.get_text(" ", strip=True)))
        if "NIM" in cells and "Nama" in cells:
            header = cells
            start = i + 1
            break
    if header is None:
        raise ValueError("Baris header tidak ditemukan")

    data = []
    for tr in trs[start:]:
        cells = []
        for c in tr.find_all(["td", "th"]):
            cells.append(squash(c.get_text(" ", strip=True)))
        if len(cells) == len(header):
            data.append(cells)
    return tidy_table(pd.DataFrame(data, columns=header))


def read_excel_report(data, engine):
    grid = pd.read_excel(io.BytesIO(data), engine=engine, header=None, dtype=str).fillna("")
    header_row = None
    for i in range(len(grid)):
        cells = []
        for value in grid.iloc[i]:
            cells.append(squash(value))
        if "NIM" in cells and "Nama" in cells:
            header_row = i
            break
    if header_row is None:
        raise ValueError("Baris header (NIM, Nama) tidak ditemukan")
    body = grid.iloc[header_row + 1:].copy()
    body.columns = list(grid.iloc[header_row])
    keep_rows = []
    for idx, row in body.iterrows():
        text = ""
        for value in row:
            text += squash(value)
        if text != "":
            keep_rows.append(idx)
    return tidy_table(body.loc[keep_rows])


def read_printed_by(text):
    soup = BeautifulSoup(text, "lxml")
    for tag in soup(["style", "script"]):
        tag.decompose()
    for line in soup.get_text("\n").split("\n"):
        if "Dicetak oleh" in line:
            return squash(line)
    return ""


def join_mk(codes, names, max_shown=5):
    pairs = []
    for code, name in zip(codes, names):
        if str(code).strip():
            pairs.append(code + " - " + name)
    if not pairs:
        return ""
    if len(pairs) > max_shown:
        return "; ".join(pairs[:max_shown]) + "; ...(+%d MK lain)" % (len(pairs) - max_shown)
    return "; ".join(pairs)


def split_people(text):
    """'ID - Nama ID - Nama' (atau dipisah ';') -> daftar entri, satu per dosen."""
    text = squash(text)
    if text in ("", "-"):
        return []
    people = []
    for part in re.split(r"\s*;\s*|\s+(?=\d{6,}\s*-\s)", text):
        part = part.strip(" ;")
        if part and part != "-":
            people.append(part)
    return people


def person_key(person):
    found = re.match(r"(\d{6,})\s*-", person)
    if found:
        return found.group(1)
    return person.lower()


def merge_people(values):
    """Gabung dosen dari semua baris MK satu aktivitas: unik per ID dosen, urut ID."""
    found = {}
    for value in values:
        for person in split_people(value):
            key = person_key(person)
            if key not in found or len(person) > len(found[key]):
                found[key] = person
    if not found:
        return "-"
    names = []
    for key in sorted(found):
        names.append(found[key])
    return "; ".join(names)


def to_float(value):
    number = pd.to_numeric(str(value).replace(",", "."), errors="coerce")
    if pd.isna(number):
        return 0.0
    return float(number)


def normalize_activities(df):
    """Gabungkan baris SIAKAD (aktivitas x MK) menjadi 1 baris per aktivitas.

    Kunci = NIM + atribut aktivitas (tanpa kolom dosen). Kolom dosen digabung
    per aktivitas; MK dihitung unik per kode.
    """
    work = df.copy()
    work["_i"] = range(len(work))
    records = []
    for _, group in work.groupby(GROUP_KEYS, sort=False, dropna=False):
        record = {}
        for key in GROUP_KEYS:
            record[key] = group.iloc[0][key]
        for col in LECTURER_COLS:
            record[col] = merge_people(group[col])
        codes = []
        names = []
        total_sks = 0.0
        for _, row in group.iterrows():
            code = str(row["Kode MK Konversi"]).strip()
            if code == "" or code in codes:
                continue
            codes.append(code)
            names.append(row["Nama MK Konversi"])
            total_sks += to_float(row["SKS"])
        record["Jml MK"] = len(codes)
        record["Total SKS"] = total_sks
        record["MK Konversi"] = join_mk(codes, names)
        record["_codes"] = ";".join(codes)
        record["_rows"] = len(group)
        record["_i"] = int(group["_i"].min())
        records.append(record)
    act = pd.DataFrame(records).sort_values("_i").reset_index(drop=True)
    act["No"] = range(1, len(act) + 1)
    return act[["No"] + KEYS + ["Jml MK", "Total SKS", "MK Konversi", "_codes", "_rows", "_i"]]


def count_split_records(raw):
    """Jumlah record yang akan terpecah jika kolom dosen ikut jadi kunci (cara lama)."""
    old = raw.groupby(KEYS, sort=False, dropna=False).ngroups
    new = raw.groupby(GROUP_KEYS, sort=False, dropna=False).ngroups
    return old - new


@st.cache_data(show_spinner=False)
def load(data, name):
    kind = sniff_format(data)
    note = ""
    if kind == "html":
        text = data.decode("utf-8", errors="replace")
        raw = read_html_report(text)
        note = read_printed_by(text)
    elif kind == "biff":
        raw = read_excel_report(data, "xlrd")
    elif kind == "xlsx":
        raw = read_excel_report(data, "openpyxl")
    else:
        raise ValueError("Format file tidak dikenali: " + kind)
    act = normalize_activities(raw)
    return raw, act, note, count_split_records(raw)


def flag_table(act, mk_overload, mk_overload_px):
    mitra = act["Mitra"].str.strip()
    bimb = act["Dosen Pembimbing"].str.strip()
    uji = act["Dosen Penguji"].str.strip()
    limit = pd.Series(mk_overload, index=act.index)
    limit[act["Jenis Aktivitas"] == JENIS_PERTUKARAN] = mk_overload_px
    out = pd.DataFrame(index=act.index)
    out["F1 Tanpa MK"] = act["Jml MK"] == 0
    out["F2 MK Berlebih"] = act["Jml MK"] > limit
    out["F3 Tanpa Pembimbing"] = bimb.isin(["-", ""])
    out["F4 Tanpa Penguji"] = uji.isin(["-", ""])
    out["F5 Mitra Tidak Valid"] = (mitra.isin(["-", ""])
                                   | mitra.str.contains("internship", case=False, regex=False)
                                   | mitra.str.contains(r"^[A-Za-z]{2,3}\d{3,6}\b", regex=True))
    out["F6 Tanpa MoU"] = act["Status Mitra"] == "Belum memiliki MoU kerjasama"
    out["F7 NIM Ganda"] = act["NIM"].duplicated(keep=False)
    return out


def prodi_label(name):
    return name.replace("S1 - ", "").strip()


def nim_check(act):
    """NIM dengan >1 aktivitas, NIM kosong/bukan angka, dan NIM dengan nama/prodi berbeda."""
    multi_rows = []
    conflict_rows = []
    for nim, group in act.groupby("NIM", sort=True):
        names = set()
        prodis = set()
        for value in group["Nama"]:
            names.add(value.upper())
        for value in group["Program Studi"]:
            prodis.add(value)
        if len(names) > 1 or len(prodis) > 1:
            conflict_rows.append({"NIM": nim, "Nama": " / ".join(sorted(names)),
                                  "Program Studi": " / ".join(sorted(prodis))})
        if len(group) < 2:
            continue
        items = []
        code_count = {}
        for _, row in group.iterrows():
            items.append("%s - %s" % (short_jenis(row["Jenis Aktivitas"]), row["Status Aktivitas"]))
            for code in row["_codes"].split(";"):
                if code:
                    code_count[code] = code_count.get(code, 0) + 1
        same = []
        for code in sorted(code_count):
            if code_count[code] > 1:
                same.append(code)
        multi_rows.append({
            "NIM": nim, "Nama": group.iloc[0]["Nama"],
            "Program Studi": prodi_label(group.iloc[0]["Program Studi"]),
            "Jml Aktivitas": len(group), "Aktivitas (Jenis - Status)": "; ".join(items),
            "MK Sama": ", ".join(same),
        })
    return {
        "multi": pd.DataFrame(multi_rows, columns=NIM_COLS),
        "conflict": pd.DataFrame(conflict_rows, columns=["NIM", "Nama", "Program Studi"]),
        "invalid": act[~act["NIM"].str.fullmatch(r"\d+")],
    }


def rekap(act, by):
    """Jumlah aktivitas & mahasiswa (NIM unik) per kategori."""
    table = pd.DataFrame({
        "Aktivitas": act.groupby(by).size(),
        "Mahasiswa": act.groupby(by)["NIM"].nunique(),
    })
    return table.sort_values(["Mahasiswa", "Aktivitas"], ascending=False)


def rekap_table(act, by):
    """Rekap + baris TOTAL (TOTAL mahasiswa = NIM unik, bukan jumlah baris)."""
    table = rekap(act, by).reset_index()
    total = pd.DataFrame([{by: "TOTAL", "Aktivitas": len(act), "Mahasiswa": act["NIM"].nunique()}])
    return pd.concat([table, total], ignore_index=True)


def student_matrix(act, lang="id"):
    """Matriks NIM unik Program Studi x Jenis; kolom/baris Total juga NIM unik."""
    piv = act.groupby(["Program Studi", "Jenis Aktivitas"])["NIM"].nunique().unstack(fill_value=0)
    piv["Total"] = act.groupby("Program Studi")["NIM"].nunique()
    total_row = act.groupby("Jenis Aktivitas")["NIM"].nunique()
    total_row["Total"] = act["NIM"].nunique()
    piv.loc["TOTAL"] = total_row
    piv = piv.fillna(0).astype(int)
    cols = []
    for c in piv.columns:
        cols.append(jenis_short(c, lang))
    rows = []
    for p in piv.index:
        rows.append(prodi_label(p))
    piv.columns = cols
    piv.index = rows
    return piv


def periode_label(act_all):
    periods = []
    for p in sorted(act_all["Periode Akademik"].unique()):
        if squash(p):
            periods.append(squash(p))
    return ", ".join(periods) if periods else "-"


def period_stem(act_all):
    parts = []
    for p in sorted(act_all["Periode Akademik"].unique()):
        if squash(p):
            parts.append(squash(p).replace(" ", "_"))
    if not parts:
        return "MBKM"
    if len(parts) > 2:
        return "MBKM_multi_periode"
    return "MBKM_" + "-".join(parts)


# ---------------------------------------------------------------- charts
def barh(series, color, title, xlabel="Jumlah"):
    fig, ax = plt.subplots(figsize=(6, 0.5 + 0.42 * len(series)))
    labels = list(series.index)[::-1]
    values = list(series.values)[::-1]
    ax.barh(labels, values, color=color)
    ax.set_title(title, fontsize=11)
    ax.set_xlabel(xlabel, fontsize=9)
    for i, v in enumerate(values):
        ax.text(v + max(values) * 0.01 + 0.05, i, str(int(v)), va="center", fontsize=8)
    ax.spines["top"].set_visible(False)
    ax.spines["right"].set_visible(False)
    plt.tight_layout()
    return fig


# ---------------------------------------------------------------- Excel, CSV & PDF
def export_table(act, flags, cols):
    table = act[cols].copy()
    for col in FLAG_COLS:
        table[col] = flags[col].map({True: "Ya", False: ""})
    return table


NOT_DONE_ORDER = ["Evaluasi", "Diajukan", "Ditolak", "Dibatalkan"]
NOT_DONE_COLS = {
    "id": ["NIM", "Nama", "Program Studi", "Jenis", "Status", "Mitra", "Aktivitas lain yang masih tercatat"],
    "en": ["NIM", "Name", "Study program", "Type", "Status", "Partner", "Other active activity"],
}


def status_rank(status):
    if status in NOT_DONE_ORDER:
        return NOT_DONE_ORDER.index(status)
    return len(NOT_DONE_ORDER)


def not_done_tables(act, act_all, drop_status, lang="id"):
    """(1) aktivitas dianalisis yang belum Selesai, (2) record yang dibuang filter (mis. Ditolak/Dibatalkan).

    Tabel (2) diberi kolom aktivitas lain milik NIM yang sama yang masih tercatat setelah filter.
    """
    cols = NOT_DONE_COLS[lang]
    pending_rows = []
    for _, r in act[act["Status Aktivitas"] != "Selesai"].iterrows():
        pending_rows.append([status_rank(r["Status Aktivitas"]), r["Program Studi"], r["Nama"],
                             r["NIM"], r["Nama"], prodi_label(r["Program Studi"]), jenis_short(r["Jenis Aktivitas"], lang),
                             status_name(r["Status Aktivitas"], lang), r["Mitra"]])
    pending_rows.sort(key=lambda x: (x[0], x[1], x[2]))
    pending = []
    for row in pending_rows:
        pending.append(row[3:])
    excluded_rows = []
    if act_all is not None and drop_status:
        for _, r in act_all[act_all["Status Aktivitas"].isin(drop_status)].iterrows():
            others = []
            for _, o in act[act["NIM"] == r["NIM"]].iterrows():
                others.append("%s - %s" % (jenis_short(o["Jenis Aktivitas"], lang),
                                           status_name(o["Status Aktivitas"], lang)))
            excluded_rows.append([status_rank(r["Status Aktivitas"]), r["Program Studi"], r["Nama"],
                                  r["NIM"], r["Nama"], prodi_label(r["Program Studi"]),
                                  jenis_short(r["Jenis Aktivitas"], lang), status_name(r["Status Aktivitas"], lang),
                                  r["Mitra"], "; ".join(others) if others else "-"])
    excluded_rows.sort(key=lambda x: (x[0], x[1], x[2]))
    excluded = []
    for row in excluded_rows:
        excluded.append(row[3:])
    return pd.DataFrame(pending, columns=cols[:6]), pd.DataFrame(excluded, columns=cols)


def build_excel(act, flags, nim_info, act_all=None, drop_status=None):
    buf = io.BytesIO()
    cols = ["No", "NIM", "Nama", "Program Studi", "Jenis Aktivitas", "Mitra",
            "Status Mitra", "Status Aktivitas", "Tanggal Mulai", "Tanggal Selesai",
            "Jml MK", "Total SKS", "MK Konversi", "Judul Aktivitas",
            "Dosen Pembimbing", "Dosen Penguji"]
    table = export_table(act, flags, cols)
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        table.to_excel(xl, sheet_name="Aktivitas", index=False)
        rekap_table(act, "Jenis Aktivitas").to_excel(xl, sheet_name="Per Jenis", index=False)
        rekap_table(act, "Program Studi").to_excel(xl, sheet_name="Per Prodi", index=False)
        nim_info["multi"].to_excel(xl, sheet_name="Cek NIM", index=False)
        pending, excluded = not_done_tables(act, act_all, drop_status, "id")
        pending.to_excel(xl, sheet_name="Belum Selesai", index=False)
        if len(excluded):
            excluded.to_excel(xl, sheet_name="Dikecualikan", index=False)
    buf.seek(0)
    return buf.getvalue()


def build_csv(act, flags):
    cols = ["No"] + KEYS + ["Jml MK", "Total SKS", "MK Konversi"]
    return export_table(act, flags, cols).to_csv(index=False).encode("utf-8")


def not_done_pdf_table(df, widths_mm, cell_style, header_font, table_style):
    """Tabel ReportLab untuk daftar mahasiswa belum Selesai / dikecualikan (sel dibungkus otomatis)."""
    from reportlab.lib import colors
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.lib.units import mm
    from reportlab.platypus import Paragraph, Table

    def esc(t):
        return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    small = ParagraphStyle("NDCELL", parent=cell_style, fontSize=7.6, leading=9.6)
    head = ParagraphStyle("NDHEAD", parent=small, fontName=header_font, textColor=colors.white)
    rows = [[]]
    for c in df.columns:
        rows[0].append(Paragraph(esc(c), head))
    for _, r in df.iterrows():
        row = []
        for c in df.columns:
            row.append(Paragraph(esc(r[c]), small))
        rows.append(row)
    widths = []
    for w in widths_mm:
        widths.append(w * mm)
    table = Table(rows, colWidths=widths, repeatRows=1)
    table.setStyle(table_style)
    return table


def build_pdf(act, raw, flags, nim_info, meta):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_LEFT
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                    TableStyle, HRFlowable, KeepTogether)

    navy = colors.HexColor(NAVY)
    red = colors.HexColor(RED)
    grey = colors.HexColor("#D9D9D9")
    light = colors.HexColor("#EEF1F7")
    mute = colors.HexColor("#666666")
    defs = flag_defs(meta["mk_overload"], meta["mk_overload_px"])

    def esc(t):
        return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    base = getSampleStyleSheet()
    H1 = ParagraphStyle("H1", parent=base["Title"], fontName="Helvetica-Bold", fontSize=17,
                        textColor=navy, alignment=TA_LEFT, spaceAfter=2, leading=20)
    SUB = ParagraphStyle("SUB", parent=base["Normal"], fontName="Helvetica", fontSize=8.5,
                         textColor=mute, spaceAfter=1, leading=11)
    H2 = ParagraphStyle("H2", parent=base["Heading2"], fontName="Helvetica-Bold", fontSize=11.5,
                        textColor=navy, spaceBefore=13, spaceAfter=5, leading=14)
    SMALL = ParagraphStyle("SMALL", parent=base["Normal"], fontName="Helvetica", fontSize=8,
                           textColor=mute, leading=10.5, spaceBefore=3)
    CELL = ParagraphStyle("CELL", parent=base["Normal"], fontName="Helvetica", fontSize=8.5, leading=11)
    CELLB = ParagraphStyle("CELLB", parent=CELL, fontName="Helvetica-Bold")

    def tstyle(header_bg=navy, total_row=False, numeric_cols=None, body_size=8.5):
        cmds = [("BACKGROUND", (0, 0), (-1, 0), header_bg),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), body_size),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING", (0, 0), (-1, -1), 6), ("RIGHTPADDING", (0, 0), (-1, -1), 6),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#BFBFBF")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, light])]
        for c in (numeric_cols or []):
            cmds.append(("ALIGN", (c, 1), (c, -1), "CENTER"))
        if total_row:
            cmds.append(("BACKGROUND", (0, -1), (-1, -1), grey))
            cmds.append(("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"))
        return TableStyle(cmds)

    story = []

    def section(title, flowables):
        story.append(KeepTogether([Paragraph(title, H2)] + flowables))

    n_act = len(act)
    n_stu = act["NIM"].nunique()
    multi = nim_info["multi"]
    multi_text = "0"
    if len(multi):
        multi_text = "%d NIM (%d aktivitas)" % (len(multi), int(multi["Jml Aktivitas"].sum()))
    n_issue = int(pd.concat([flags[k] for k in RECORD_FLAGS], axis=1).any(axis=1).sum())
    status_ct = act["Status Aktivitas"].value_counts()

    story.append(Paragraph("Laporan Analisis Aktivitas MBKM &amp; Mata Kuliah Konversi", H1))
    story.append(Paragraph("Indonesia International Institute for Life Sciences (i3L)", SUB))
    story.append(Paragraph("Periode Akademik: " + esc(meta["periode"]), SUB))
    if meta["source_note"]:
        story.append(Paragraph(esc(meta["source_note"]).replace("|", "&nbsp;|&nbsp;"), SUB))
    story.append(Spacer(1, 3))
    story.append(HRFlowable(width="100%", thickness=1.2, color=navy, spaceAfter=2))

    # 1. Ringkasan
    rows = [["Metrik", "Nilai"],
            ["Jumlah aktivitas (setelah filter)", str(n_act)],
            ["Jumlah mahasiswa (NIM unik)", str(n_stu)],
            ["NIM dengan >1 aktivitas", multi_text],
            ["Baris pada file ekspor asli", str(len(raw))],
            ["Record dikecualikan (%s)" % (", ".join(meta["drop_status"]) or "-"), str(meta["n_dropped"])],
            ["Aktivitas dengan isu data per-record", str(n_issue)]]
    t = Table(rows, colWidths=[120 * mm, 45 * mm])
    t.setStyle(tstyle(numeric_cols=[1]))
    if n_act == n_stu:
        note = "Setiap NIM tercatat tepat 1 aktivitas, sehingga jumlah aktivitas = jumlah mahasiswa."
    else:
        note = ("Jumlah aktivitas lebih besar dari jumlah mahasiswa karena %d NIM tercatat di lebih "
                "dari 1 aktivitas (lihat bagian 6)." % len(multi))
    section("1. Ringkasan Umum", [t, Paragraph(note, SMALL)])

    # 2. Per jenis
    per_jenis = rekap(act, "Jenis Aktivitas")
    rows = [["Jenis Aktivitas", "Aktivitas", "Mahasiswa"]]
    for j in per_jenis.index:
        rows.append([short_jenis(j), str(int(per_jenis.loc[j, "Aktivitas"])),
                     str(int(per_jenis.loc[j, "Mahasiswa"]))])
    rows.append(["TOTAL", str(n_act), str(n_stu)])
    t = Table(rows, colWidths=[105 * mm, 30 * mm, 30 * mm])
    t.setStyle(tstyle(total_row=True, numeric_cols=[1, 2]))
    items = [t]
    if int(per_jenis["Mahasiswa"].sum()) != n_stu:
        items.append(Paragraph("Mahasiswa yang memiliki aktivitas di lebih dari 1 jenis dihitung di tiap jenis; "
                               "baris TOTAL = NIM unik.", SMALL))
    section("2. Rekap per Jenis Aktivitas", items)

    # 3. Per prodi
    per_prodi = rekap(act, "Program Studi")
    rows = [["Program Studi", "Aktivitas", "Mahasiswa"]]
    for p in per_prodi.index:
        rows.append([Paragraph(esc(prodi_label(p)), CELL), str(int(per_prodi.loc[p, "Aktivitas"])),
                     str(int(per_prodi.loc[p, "Mahasiswa"]))])
    rows.append([Paragraph("TOTAL", CELLB), str(n_act), str(n_stu)])
    t = Table(rows, colWidths=[105 * mm, 30 * mm, 30 * mm])
    t.setStyle(tstyle(total_row=True, numeric_cols=[1, 2]))
    section("3. Rekap per Program Studi", [t])

    # 4. Status
    rows = [["Status", "Aktivitas"]]
    for s in status_ct.index:
        rows.append([s, str(int(status_ct[s]))])
    t = Table(rows, colWidths=[130 * mm, 35 * mm])
    t.setStyle(tstyle(numeric_cols=[1]))
    section("4. Status Aktivitas", [t])

    # 5. Flag
    rows = [["Flag", "Jml", "Definisi"]]
    for k in FLAG_COLS:
        rows.append([Paragraph(esc(k), CELLB), Paragraph(str(int(flags[k].sum())), CELL),
                     Paragraph(esc(defs[k]), CELL)])
    t = Table(rows, colWidths=[40 * mm, 14 * mm, 111 * mm])
    t.setStyle(tstyle(header_bg=red, numeric_cols=[1]))
    section("5. Temuan Kualitas Data (Flag)", [
        t, Paragraph("F1, F2, F3, F5 = isu input per-record. F4 &amp; F6 = isu sistemik. "
                     "F7 = perlu verifikasi (lihat bagian 6).", SMALL)])

    # 6. NIM
    items = []
    if len(multi) == 0:
        items.append(Paragraph("Tidak ada NIM ganda: setiap NIM tercatat tepat 1 aktivitas.", CELL))
    else:
        items.append(Paragraph(
            "%d NIM tercatat di lebih dari 1 aktivitas, sehingga jumlah aktivitas melebihi jumlah mahasiswa "
            "sebanyak %d. Kolom 'MK sama' berisi MK yang dikonversi di lebih dari 1 aktivitas "
            "(indikasi kuat duplikasi / pengajuan ulang)." % (len(multi), n_act - n_stu), CELL))
        items.append(Spacer(1, 4))
        rows = [["NIM", "Nama", "Program Studi", "Aktivitas (Jenis - Status)", "MK sama"]]
        for _, r in multi.iterrows():
            rows.append([r["NIM"], Paragraph(esc(r["Nama"]), CELL), Paragraph(esc(r["Program Studi"]), CELL),
                         Paragraph(esc(r["Aktivitas (Jenis - Status)"]), CELL),
                         Paragraph(esc(r["MK Sama"] or "-"), CELL)])
        t = Table(rows, colWidths=[20 * mm, 40 * mm, 37 * mm, 48 * mm, 20 * mm], repeatRows=1)
        t.setStyle(tstyle())
        items.append(t)
    if len(nim_info["invalid"]):
        items.append(Paragraph("NIM kosong / bukan angka: %d aktivitas." % len(nim_info["invalid"]), SMALL))
    if len(nim_info["conflict"]):
        items.append(Paragraph("NIM dengan nama / program studi berbeda: %d NIM." % len(nim_info["conflict"]), SMALL))
    if meta["n_split"]:
        items.append(Paragraph("%d record aktivitas yang terpecah karena urutan nama dosen berbeda di ekspor "
                               "SIAKAD telah digabung." % meta["n_split"], SMALL))
    section("6. Pemeriksaan NIM", items)

    # 7. Mahasiswa dengan status selain Selesai
    pending, excluded = not_done_tables(act, meta.get("act_all"), meta["drop_status"], "id")
    items = []
    if len(pending) == 0:
        items.append(Paragraph("Semua aktivitas yang dianalisis telah berstatus Selesai.", CELL))
    else:
        counts = pending["Status"].value_counts()
        parts = []
        for s in sorted(counts.index, key=status_rank):
            parts.append("%d %s" % (counts[s], s))
        items.append(Paragraph("%d aktivitas (%d mahasiswa) belum berstatus Selesai: %s."
                               % (len(pending), pending["NIM"].nunique(), ", ".join(parts)), CELL))
        items.append(Spacer(1, 4))
        items.append(not_done_pdf_table(pending, [18, 40, 36, 24, 18, 39], CELL, "Helvetica-Bold", tstyle()))
    if len(excluded):
        items.append(Spacer(1, 6))
        items.append(Paragraph("Record yang dikecualikan oleh filter (%s): %d record."
                               % (", ".join(meta["drop_status"]), len(excluded)), CELL))
        items.append(Spacer(1, 4))
        items.append(not_done_pdf_table(excluded, [17, 33, 29, 20, 18, 31, 27], CELL, "Helvetica-Bold", tstyle(header_bg=red)))
    section("7. Mahasiswa dengan Status Selain Selesai", items)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm,
                            topMargin=16 * mm, bottomMargin=15 * mm, title="Laporan Analisis MBKM")
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------- laporan naratif: bahasa
LANGS = {"Indonesia": "id", "English": "en"}
JENIS_EN = {
    "Magang/Praktik Kerja (Kampus Merdeka)": ("Internship/Work Placement", "Internship"),
    "Penelitian/Riset (Kampus Merdeka)": ("Research", "Research"),
    "Kegiatan Wirausaha (Kampus Merdeka)": ("Entrepreneurship", "Entrepreneurship"),
    "Pertukaran Pelajar (Kampus Merdeka)": ("Student Exchange", "Exchange"),
    "Studi/Proyek Independen (Kampus Merdeka)": ("Independent Study/Project", "Independent Study"),
}
STATUS_EN = {"Selesai": "Completed", "Evaluasi": "Under evaluation", "Diajukan": "Submitted",
             "Ditolak": "Rejected", "Dibatalkan": "Cancelled"}
STATUS_EN_STATE = {"Selesai": "completed", "Evaluasi": "under evaluation", "Diajukan": "awaiting approval",
                   "Ditolak": "rejected", "Dibatalkan": "cancelled"}
CAT_EN = {CAT_MOU: "External, with MoU", CAT_NOMOU: "External, no MoU", CAT_INT: "Internal (i3L)",
          CAT_BAD: "Invalid partner"}
NARR_TEXT = {
    "id": {
        "title": "Laporan Naratif Aktivitas MBKM",
        "institute": "Indonesia International Institute for Life Sciences (i3L)",
        "period": "Periode Akademik: %s",
        "s1": "1. Ringkasan Eksekutif",
        "s2": "2. Analisis per Program Studi",
        "s3": "3. Analisis per Jenis Aktivitas",
        "s_pending": "4. Mahasiswa dengan Status Selain Selesai",
        "pending_t1": "Tabel 1. Aktivitas yang belum berstatus Selesai.",
        "pending_t2": "Tabel 2. Record yang dikecualikan dari analisis, beserta aktivitas lain mahasiswa yang masih tercatat.",
        "s4": "5. Kualitas Data dan Tindak Lanjut",
        "s5": "6. Catatan Metodologi",
        "appendix": "Lampiran. Matriks Mahasiswa per Program Studi dan Jenis Aktivitas",
        "appendix_note": "Angka = jumlah mahasiswa (NIM unik); kolom dan baris Total juga dihitung sebagai NIM unik.",
        "actions": "Tindak lanjut yang disarankan:",
        "footer": "Laporan Naratif Aktivitas MBKM - Periode %s",
        "page": "Halaman ",
        "prodi_col": "Program Studi",
        "x_label": "Jumlah aktivitas",
        "students_paren": "%d (%d mahasiswa)",
        "k_students": "Mahasiswa", "k_acts": "Aktivitas", "k_prodi": "Program studi",
        "k_done": "Aktivitas selesai", "k_ext": "Mitra eksternal", "k_mou": "Eksternal ber-MoU",
        "prodi_intro": ("Gambar 1 memperlihatkan komposisi jenis aktivitas di setiap program studi, dan Gambar 2 "
                        "memperlihatkan kategori mitranya: eksternal yang sudah atau belum didukung MoU, internal "
                        "i3L, serta mitra yang tidak valid. Uraian berikut diurutkan dari program studi dengan "
                        "jumlah mahasiswa terbanyak."),
        "jenis_intro": ("Gambar 3 memperlihatkan status aktivitas untuk setiap jenis aktivitas, dan Gambar 4 "
                        "memperlihatkan kategori mitranya. Uraian berikut diurutkan dari jenis aktivitas dengan "
                        "peserta terbanyak."),
        "file": "laporan_naratif",
        "doc_lang": "id-ID",
    },
    "en": {
        "title": "MBKM Activity Narrative Report",
        "institute": "Indonesia International Institute for Life Sciences (i3L)",
        "period": "Academic period: %s",
        "s1": "1. Executive Summary",
        "s2": "2. Analysis by Study Program",
        "s3": "3. Analysis by Activity Type",
        "s_pending": "4. Students with a Status Other than Completed",
        "pending_t1": "Table 1. Activities not yet completed.",
        "pending_t2": "Table 2. Records excluded from the analysis, with the student's other activity still on record.",
        "s4": "5. Data Quality and Follow-up",
        "s5": "6. Methodological Notes",
        "appendix": "Appendix. Student Matrix by Study Program and Activity Type",
        "appendix_note": "Values = number of students (unique NIM); the Total row and column also count unique NIMs.",
        "actions": "Recommended follow-up:",
        "footer": "MBKM Activity Narrative Report - %s",
        "page": "Page ",
        "prodi_col": "Study program",
        "x_label": "Number of activities",
        "students_paren": "%d (%d students)",
        "k_students": "Students", "k_acts": "Activities", "k_prodi": "Study programs",
        "k_done": "Completed", "k_ext": "External partners", "k_mou": "External with MoU",
        "prodi_intro": ("Figure 1 shows the mix of activity types in each study program, and Figure 2 shows the "
                        "partner categories: external partners with or without an MoU, internal i3L activities, and "
                        "invalid partner entries. Study programs are listed from the largest number of students."),
        "jenis_intro": ("Figure 3 shows the status of activities for each activity type, and Figure 4 shows their "
                        "partner categories. Activity types are listed from the largest number of participants."),
        "file": "narrative_report",
        "doc_lang": "en-US",
    },
}
FIG_CAPTIONS = {
    "id": {
        "prodi_jenis": "Gambar 1. Jumlah aktivitas per program studi menurut jenis aktivitas. Angka di dalam batang = "
                       "jumlah aktivitas per jenis; angka di ujung batang = total aktivitas (jumlah mahasiswa "
                       "ditampilkan bila berbeda).",
        "prodi_mitra": "Gambar 2. Kategori mitra per program studi (jumlah aktivitas). Aktivitas internal i3L tidak "
                       "memerlukan MoU.",
        "jenis_status": "Gambar 3. Status aktivitas per jenis aktivitas. Angka di ujung batang = total aktivitas dan "
                        "persentasenya terhadap seluruh aktivitas.",
        "jenis_mitra": "Gambar 4. Kategori mitra per jenis aktivitas (jumlah aktivitas).",
    },
    "en": {
        "prodi_jenis": "Figure 1. Number of activities per study program by activity type. Numbers inside the bars = "
                       "activities per type; number at the end of each bar = total activities (number of students "
                       "shown when different).",
        "prodi_mitra": "Figure 2. Partner category per study program (number of activities). Internal i3L activities "
                       "do not require an MoU.",
        "jenis_status": "Figure 3. Activity status per activity type. Number at the end of each bar = total activities "
                        "and their share of all activities.",
        "jenis_mitra": "Figure 4. Partner category per activity type (number of activities).",
    },
}


def num(x, lang, nd=1):
    """Angka sesuai bahasa: id = 1.929 / 16,2 ; en = 1,929 / 16.2 (tanpa desimal nol)."""
    x = round(float(x), nd)
    if abs(x - round(x)) < 1e-9:
        text = "{:,}".format(int(round(x)))
    else:
        text = "{:,.{nd}f}".format(x, nd=nd)
    if lang == "id":
        return text.replace(",", "#").replace(".", ",").replace("#", ".")
    return text


def pct(part, whole, lang):
    if not whole:
        return "0%"
    return num(100.0 * part / whole, lang) + "%"


def join(items, lang):
    items = list(items)
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    word = "dan" if lang == "id" else "and"
    if len(items) == 2:
        if lang == "id" and (" dan " in items[0] or " dan " in items[1]):
            return items[0] + " serta " + items[1]
        return "%s %s %s" % (items[0], word, items[1])
    return ", ".join(items[:-1]) + ", %s %s" % (word, items[-1])


def join_or(items):
    """Daftar alternatif bahasa Inggris: 'A or B', 'A, B, or C'."""
    items = list(items)
    if len(items) <= 1:
        return "".join(items)
    if len(items) == 2:
        return "%s or %s" % (items[0], items[1])
    return ", ".join(items[:-1]) + ", or " + items[-1]


def plural(n, one, many):
    if n == 1:
        return one
    return many


def acts(n, lang):
    if lang == "id":
        return "%d aktivitas" % n
    return "%d %s" % (n, plural(n, "activity", "activities"))


def studs(n, lang):
    if lang == "id":
        return "%d mahasiswa" % n
    return "%d %s" % (n, plural(n, "student", "students"))


def was(n):
    return plural(n, "was", "were")


def ordinal(n):
    if 10 <= n % 100 <= 20:
        return "%dth" % n
    return "%d%s" % (n, {1: "st", 2: "nd", 3: "rd"}.get(n % 10, "th"))


def jenis_label(name):
    return str(name).replace("(Kampus Merdeka)", "").strip()


def jenis_long(name, lang):
    if lang == "en" and name in JENIS_EN:
        return JENIS_EN[name][0]
    return jenis_label(name)


def jenis_short(name, lang):
    if lang == "en" and name in JENIS_EN:
        return JENIS_EN[name][1]
    if lang == "en":
        return jenis_label(name)
    return short_jenis(name)


def status_name(name, lang):
    if lang == "en":
        return STATUS_EN.get(name, name)
    return name


def status_state(name):
    return STATUS_EN_STATE.get(name, "with status " + name)


def cat_name(name, lang):
    if lang == "en":
        return CAT_EN.get(name, name)
    return name


ID_MONTHS_EN = {"Januari": "January", "Februari": "February", "Maret": "March", "Mei": "May", "Juni": "June",
                "Juli": "July", "Agustus": "August", "Oktober": "October", "Desember": "December"}


def source_note_text(note, lang):
    """Keterangan cetak SIAKAD; untuk bahasa Inggris awalan dan nama bulan diterjemahkan."""
    if lang != "en" or not note:
        return note
    text = note.replace("Dicetak oleh:", "Printed by:").replace(", pada ", ", on ")
    for indo, eng in ID_MONTHS_EN.items():
        text = re.sub(r"\b%s\b" % indo, eng, text)
    return text


def period_in_text(periode, lang):
    if lang == "en":
        text = re.sub(r"\bGanjil\b", "Odd Semester", periode)
        return re.sub(r"\bGenap\b", "Even Semester", text)
    return periode


def period_display(periode, lang):
    text = period_in_text(periode, lang)
    if text != periode:
        return "%s (%s)" % (text, periode)
    return periode


def ranked(counts):
    """Hitungan (Series/dict) -> list (nama, n), urut n menurun lalu nama."""
    pairs = []
    for name, n in dict(counts).items():
        pairs.append((name, int(n)))
    pairs.sort(key=lambda p: (-p[1], str(p[0])))
    return pairs


def count_phrase(n, total, lang):
    if n == total and total > 1:
        return "seluruh aktivitas" if lang == "id" else "all activities"
    return acts(n, lang)


def en_issue(count, total, none_text, has_text, verb_one, verb_many):
    """Kalimat isu data bahasa Inggris: 'no activity has X' bila semua, selain itu '3 activities have no X'."""
    if count == total and total > 1:
        return none_text
    return "%s %s %s" % (acts(count, "en"), plural(count, verb_one, verb_many), has_text)


def md_escape(text):
    out = str(text)
    for ch in ["\\", "*", "_", "`", "$", "[", "]"]:
        out = out.replace(ch, "\\" + ch)
    return out


def partner_categories(act, flags):
    """Kategori mitra per aktivitas: eksternal (sudah/belum MoU), internal i3L, atau tidak valid."""
    cats = []
    for i in act.index:
        mitra = act.at[i, "Mitra"]
        if flags.at[i, "F5 Mitra Tidak Valid"]:
            cats.append(CAT_BAD)
        elif re.search(INTERNAL_RE, mitra, re.I):
            cats.append(CAT_INT)
        elif act.at[i, "Status Mitra"] == "Sudah memiliki MoU kerjasama":
            cats.append(CAT_MOU)
        else:
            cats.append(CAT_NOMOU)
    return pd.Series(cats, index=act.index)


def mk_name_map(raw):
    names = {}
    for code, name in zip(raw["Kode MK Konversi"], raw["Nama MK Konversi"]):
        code = str(code).strip()
        if code and code not in names:
            names[code] = name
    return names


def group_profile(sub, cats, flags):
    """Ringkasan angka untuk satu kelompok aktivitas (satu prodi / satu jenis / seluruh data)."""
    c = cats.loc[sub.index]
    f = flags.loc[sub.index]
    prof = {
        "n_act": len(sub), "n_stu": sub["NIM"].nunique(),
        "status": sub["Status Aktivitas"].value_counts(),
        "n_mou": int((c == CAT_MOU).sum()), "n_nomou": int((c == CAT_NOMOU).sum()),
        "n_int": int((c == CAT_INT).sum()), "n_bad": int((c == CAT_BAD).sum()),
    }
    prof["n_ext"] = prof["n_mou"] + prof["n_nomou"]
    ext = sub[c.isin([CAT_MOU, CAT_NOMOU])]
    prof["partners"] = ranked(ext["Mitra"].value_counts())
    prof["sks_total"] = float(sub["Total SKS"].sum())
    prof["sks_avg"] = float(sub["Total SKS"].mean()) if len(sub) else 0.0
    prof["mk_avg"] = float(sub["Jml MK"].mean()) if len(sub) else 0.0
    with_mk = sub[sub["Jml MK"] > 0]
    prof["sks_min"] = float(with_mk["Total SKS"].min()) if len(with_mk) else 0.0
    prof["sks_max"] = float(with_mk["Total SKS"].max()) if len(with_mk) else 0.0
    mk_count = {}
    for codes in sub["_codes"]:
        for code in codes.split(";"):
            if code:
                mk_count[code] = mk_count.get(code, 0) + 1
    prof["top_mk"] = ranked(mk_count)
    prof["flags"] = {}
    for k in FLAG_COLS:
        prof["flags"][k] = int(f[k].sum())
    prof["multi_nim"] = int(sub.loc[f["F7 NIM Ganda"], "NIM"].nunique())
    return prof


# ---------------------------------------------------------------- laporan naratif: kalimat
def all_phrase(n, lang):
    if lang == "id":
        if n == 1:
            return "Aktivitas tersebut"
        if n == 2:
            return "Kedua aktivitas"
        return "Seluruh %d aktivitas" % n
    if n == 1:
        return "This activity"
    if n == 2:
        return "Both activities"
    return "All %d activities" % n


def sebanyak(n, lang):
    if lang == "id":
        if n == 1:
            return "Satu aktivitas"
        return "Sebanyak %d aktivitas" % n
    if n == 1:
        return "One activity"
    return "A total of %d activities" % n


def status_sentence(status, n_act, lang):
    n_done = int(status.get("Selesai", 0))
    pairs = ranked(status)
    others = []
    for name, n in pairs:
        if name == "Selesai":
            continue
        if lang == "id":
            others.append("%d berstatus %s" % (n, name))
        else:
            others.append("%d %s %s" % (n, plural(n, "is", "are"), status_state(name)))
    if n_done == n_act:
        if lang == "id":
            if n_act == 1:
                return "Aktivitas tersebut telah berstatus Selesai."
            if n_act == 2:
                return "Kedua aktivitas telah berstatus Selesai."
            return "Seluruh %d aktivitas telah berstatus Selesai." % n_act
        if n_act == 1:
            return "This activity has been completed."
        if n_act == 2:
            return "Both activities have been completed."
        return "All %d activities have been completed." % n_act
    if n_done == 0:
        if lang == "id":
            if n_act == 1:
                return "Aktivitas tersebut masih berstatus %s." % pairs[0][0]
            return "Belum ada aktivitas yang berstatus Selesai; " + join(others, lang) + "."
        if n_act == 1:
            return "This activity is still %s." % status_state(pairs[0][0])
        return "None of the activities has been completed yet; " + join(others, lang) + "."
    if lang == "id":
        return ("Dari %d aktivitas, %d (%s) telah berstatus Selesai, sementara %s."
                % (n_act, n_done, pct(n_done, n_act, lang), join(others, lang)))
    return ("Of the %d activities, %d (%s) %s been completed, while %s."
            % (n_act, n_done, pct(n_done, n_act, lang), plural(n_done, "has", "have"), join(others, lang)))


def mou_sentence(prof, lang):
    n_mou = prof["n_mou"]
    n_ext = prof["n_ext"]
    if lang == "id":
        if n_mou == 0:
            return "Belum ada aktivitas bermitra eksternal yang didukung MoU kerja sama."
        if n_mou == n_ext:
            return "Seluruh aktivitas bermitra eksternal tersebut sudah didukung MoU kerja sama."
        return ("Dari aktivitas bermitra eksternal tersebut, %d (%s) sudah didukung MoU kerja sama."
                % (n_mou, pct(n_mou, n_ext, lang)))
    if n_mou == 0:
        if n_ext == 1:
            return "This externally partnered activity is not covered by an MoU."
        return "None of the externally partnered activities is covered by an MoU."
    if n_mou == n_ext:
        if n_ext == 1:
            return "This externally partnered activity is covered by an MoU."
        return "All of these externally partnered activities are covered by an MoU."
    return ("Of these externally partnered activities, %d (%s) %s covered by an MoU."
            % (n_mou, pct(n_mou, n_ext, lang), plural(n_mou, "is", "are")))


def describe_group(names, n, lang):
    shown = list(names)
    if len(shown) > 4:
        if lang == "id":
            shown = shown[:3] + ["%d mitra lain" % (len(names) - 3)]
        else:
            shown = shown[:3] + ["%d other partners" % (len(names) - 3)]
    if lang == "id":
        if len(names) == 1:
            return "%s dengan %d aktivitas" % (shown[0], n)
        return "%s dengan masing-masing %d aktivitas" % (join(shown, lang), n)
    if len(names) == 1:
        return "%s, with %s" % (shown[0], acts(n, lang))
    return "%s, with %s each" % (join(shown, lang), acts(n, lang))


def partner_sentences(prof, lang):
    out = []
    n_act = prof["n_act"]
    n_ext = prof["n_ext"]
    partners = prof["partners"]
    if n_ext:
        if n_ext == n_act:
            subject = all_phrase(n_act, lang)
        else:
            subject = "%s (%s)" % (sebanyak(n_ext, lang), pct(n_ext, n_act, lang))
        if lang == "id":
            head = subject + " dilaksanakan bersama "
        else:
            head = "%s %s carried out with " % (subject, was(n_ext))
        if len(partners) == 1:
            if lang == "id":
                out.append(head + "satu mitra eksternal, yaitu %s." % partners[0][0])
            else:
                out.append(head + "one external partner, %s." % partners[0][0])
        else:
            if lang == "id":
                out.append(head + "%d mitra eksternal." % len(partners))
            else:
                out.append(head + "%d external partners." % len(partners))
            groups = []
            for name, n in partners:
                if n < 2:
                    break
                if groups and groups[-1][1] == n:
                    groups[-1][0].append(name)
                elif len(groups) < 2:
                    groups.append(([name], n))
                else:
                    break
            if groups:
                listed = 0
                for names, n in groups:
                    listed += len(names)
                if lang == "id":
                    if listed == len(partners):
                        lead = "Mitra tersebut adalah "
                    else:
                        lead = "Mitra yang paling sering adalah "
                    follow = ", disusul "
                else:
                    if listed == len(partners):
                        lead = "These partners are "
                    elif len(groups[0][0]) == 1:
                        lead = "The most frequent partner is "
                    else:
                        lead = "The most frequent partners are "
                    follow = ", followed by "
                text = lead + describe_group(groups[0][0], groups[0][1], lang)
                if len(groups) > 1:
                    text += follow + describe_group(groups[1][0], groups[1][1], lang)
                out.append(text + ".")
            else:
                names = []
                for name, n in partners[:3]:
                    names.append(name)
                if lang == "id":
                    if len(partners) <= 3:
                        out.append("Mitra tersebut adalah %s." % join(names, lang))
                    else:
                        out.append("Mitra tersebut antara lain %s." % join(names, lang))
                else:
                    if len(partners) <= 3:
                        out.append("These partners are %s." % join(names, lang))
                    else:
                        out.append("They include %s." % join(names, lang))
        out.append(mou_sentence(prof, lang))
    n_int = prof["n_int"]
    if n_int:
        if n_int == n_act:
            subject = all_phrase(n_act, lang)
        else:
            subject = "%s (%s)" % (sebanyak(n_int, lang), pct(n_int, n_act, lang))
        if lang == "id":
            out.append(subject + " dilaksanakan secara internal di i3L.")
        else:
            out.append("%s %s carried out internally at i3L." % (subject, was(n_int)))
    return out


def conversion_sentence(prof, mk_names, lang):
    if prof["sks_total"] <= 0:
        if lang == "id":
            return "Belum ada SKS yang dikonversi."
        return "No credits (SKS) have been converted."
    mk = num(prof["mk_avg"], lang)
    total = num(prof["sks_total"], lang)
    if lang == "id":
        if prof["n_act"] == 1:
            return "Aktivitas tersebut mengonversi %s MK dengan total %s SKS." % (mk, total)
        text = "Rata-rata setiap aktivitas mengonversi %s MK atau %s SKS" % (mk, num(prof["sks_avg"], lang))
        if prof["sks_max"] > prof["sks_min"]:
            text += " (rentang %s–%s SKS)" % (num(prof["sks_min"], lang), num(prof["sks_max"], lang))
        text += ", dengan total %s SKS." % total
    else:
        course = "course" if mk == "1" else "courses"
        if prof["n_act"] == 1:
            return "This activity converted %s %s, totalling %s credits." % (mk, course, total)
        text = ("On average, each activity converted %s %s or %s credits"
                % (mk, course, num(prof["sks_avg"], lang)))
        if prof["sks_max"] > prof["sks_min"]:
            text += " (range %s–%s credits)" % (num(prof["sks_min"], lang), num(prof["sks_max"], lang))
        text += ", for a total of %s credits." % total
    if prof["top_mk"] and prof["top_mk"][0][1] > 1:
        code, n = prof["top_mk"][0]
        if lang == "id":
            text += (" MK yang paling sering dikonversi adalah %s – %s (%d aktivitas)."
                     % (code, mk_names.get(code, ""), n))
        else:
            text += (" The most frequently converted course is %s – %s (%s)."
                     % (code, mk_names.get(code, ""), acts(n, lang)))
    return text


def data_note(prof, scope, lang):
    f = prof["flags"]
    n = prof["n_act"]
    f1 = f["F1 Tanpa MK"]
    f2 = f["F2 MK Berlebih"]
    f3 = f["F3 Tanpa Pembimbing"]
    f4 = f["F4 Tanpa Penguji"]
    f5 = f["F5 Mitra Tidak Valid"]
    k = prof["multi_nim"]
    en = lang == "en"
    issues = []
    if n > 1 and f3 == n and f4 == n:
        if en:
            issues.append("no activity has a supervisor or an examiner recorded")
        else:
            issues.append("seluruh aktivitas belum mencantumkan dosen pembimbing maupun dosen penguji")
    else:
        if f3:
            if en:
                issues.append(en_issue(f3, n, "no activity has a supervisor recorded", "no supervisor recorded",
                                       "has", "have"))
            else:
                issues.append("%s belum mencantumkan dosen pembimbing" % count_phrase(f3, n, lang))
        if f4:
            if en:
                issues.append(en_issue(f4, n, "no activity has an examiner recorded", "no examiner recorded",
                                       "has", "have"))
            else:
                issues.append("%s belum mencantumkan dosen penguji" % count_phrase(f4, n, lang))
    if f1:
        if en:
            issues.append(en_issue(f1, n, "no activity has a converted course", "no converted course", "has", "have"))
        else:
            issues.append("%s belum memiliki MK konversi" % count_phrase(f1, n, lang))
    if f2:
        if en:
            issues.append(en_issue(f2, n, "all activities exceed the course-count threshold",
                                   "the course-count threshold", "exceeds", "exceed"))
        else:
            issues.append("%s memiliki jumlah MK di atas ambang" % count_phrase(f2, n, lang))
    if f5:
        if en:
            issues.append(en_issue(f5, n, "no activity has a valid partner name", "a valid partner name",
                                   "lacks", "lack"))
        else:
            issues.append("%s belum mencantumkan nama mitra yang valid" % count_phrase(f5, n, lang))
    if k:
        if en:
            issues.append("%s %s recorded in more than one activity" % (studs(k, lang), plural(k, "is", "are")))
        else:
            issues.append("%d mahasiswa tercatat pada lebih dari satu aktivitas" % k)
    if not issues:
        if en:
            return ("Data for this %s are complete: every activity has a supervisor, an examiner, a converted "
                    "course, and a valid partner." % ("study program" if scope == "prodi" else "activity type"))
        return ("Data %s ini lengkap: setiap aktivitas memiliki dosen pembimbing, dosen penguji, MK konversi, "
                "dan mitra yang valid." % ("program studi" if scope == "prodi" else "jenis aktivitas"))
    if en:
        return "Data notes: " + join(issues, lang) + "."
    return "Catatan data: " + join(issues, lang) + "."


def rank_clause(name, counts, kind, lang):
    """Posisi kelompok di antara kelompok lain (hanya bila ada >= 3 kelompok)."""
    if len(counts) < 3:
        return ""
    en = lang == "en"
    if kind == "prodi":
        noun = "study programs" if en else "program studi"
    else:
        noun = "activity types" if en else "jenis aktivitas"
    mine = 0
    for other, n in counts:
        if other == name:
            mine = n
    higher = 0
    lower = 0
    ties = []
    for other, n in counts:
        if other == name:
            continue
        if n > mine:
            higher += 1
        elif n < mine:
            lower += 1
        elif kind == "prodi":
            ties.append(prodi_label(other))
        else:
            ties.append(jenis_long(other, lang))
    if higher == 0:
        if en:
            best = "the highest" if kind == "prodi" else "the largest"
            if ties:
                return ", %s together with %s" % (best, join(ties, lang))
            return ", %s of all %s" % (best, noun)
        best = "tertinggi" if kind == "prodi" else "terbanyak"
        if ties:
            return ", %s bersama %s" % (best, join(ties, lang))
        return ", %s di antara seluruh %s" % (best, noun)
    if lower == 0:
        if en:
            if ties:
                return ", the fewest together with %s" % join(ties, lang)
            return ", the fewest among the %s" % noun
        if ties:
            return ", paling sedikit bersama %s" % join(ties, lang)
        return ", paling sedikit di antara %s yang ada" % noun
    if en:
        return " and ranked %s of %d %s" % (ordinal(higher + 1), len(counts), noun)
    return " dan menempati peringkat ke-%d dari %d %s" % (higher + 1, len(counts), noun)


def jenis_mix_sentence(jenis_counts, n_stu, lang):
    en = lang == "en"
    if len(jenis_counts) == 1:
        if en:
            return "All students took part in %s." % jenis_long(jenis_counts[0][0], lang)
        return "Seluruh mahasiswa mengikuti %s." % jenis_long(jenis_counts[0][0], lang)
    first_j, first_n = jenis_counts[0]
    rest = []
    for j, n in jenis_counts[1:]:
        if en:
            rest.append("%s (%s, %s)" % (jenis_long(j, lang), studs(n, lang), pct(n, n_stu, lang)))
        else:
            rest.append("%s (%d mahasiswa, %s)" % (jenis_long(j, lang), n, pct(n, n_stu, lang)))
    if en:
        return ("By activity type, %s was the most common choice with %s (%s), followed by %s."
                % (jenis_long(first_j, lang), studs(first_n, lang), pct(first_n, n_stu, lang), join(rest, lang)))
    return ("Berdasarkan jenis aktivitas, %s merupakan pilihan terbanyak dengan %d mahasiswa (%s), disusul %s."
            % (jenis_long(first_j, lang), first_n, pct(first_n, n_stu, lang), join(rest, lang)))


def participation_sentence(prodi_counts, n_stu, lang):
    en = lang == "en"
    if len(prodi_counts) == 1:
        if en:
            return "All participants came from the %s study program." % prodi_label(prodi_counts[0][0])
        return "Seluruh peserta berasal dari Program Studi %s." % prodi_label(prodi_counts[0][0])
    top_n = prodi_counts[0][1]
    tops = []
    for p, n in prodi_counts:
        if n == top_n:
            tops.append(p)
    top_labels = []
    for p in tops:
        top_labels.append(prodi_label(p))
    if len(tops) == 1:
        if en:
            text = ("By study program, participation was highest in %s (%s, %s)"
                    % (top_labels[0], studs(top_n, lang), pct(top_n, n_stu, lang)))
        else:
            text = ("Dari sisi program studi, partisipasi tertinggi berasal dari %s (%d mahasiswa, %s)"
                    % (top_labels[0], top_n, pct(top_n, n_stu, lang)))
    else:
        if en:
            text = ("By study program, participation was highest in %s (%s each)"
                    % (join(top_labels, lang), studs(top_n, lang)))
        else:
            text = ("Dari sisi program studi, partisipasi tertinggi berasal dari %s (masing-masing %d mahasiswa)"
                    % (join(top_labels, lang), top_n))
    low_n = prodi_counts[-1][1]
    lows = []
    if len(prodi_counts) >= 3 and low_n < top_n:
        for p, n in prodi_counts:
            if n == low_n:
                lows.append(p)
    middle = []
    for p, n in prodi_counts:
        if p not in tops and p not in lows and len(middle) < 2:
            middle.append("%s (%d)" % (prodi_label(p), n))
    if middle:
        text += (", followed by %s" if en else ", diikuti %s") % join(middle, lang)
    if lows:
        low_labels = []
        for p in lows:
            low_labels.append(prodi_label(p))
        if en:
            if len(lows) == 1:
                text += ", while participation was lowest in %s (%s)" % (low_labels[0], studs(low_n, lang))
            else:
                text += (", while participation was lowest in %s (%s each)"
                         % (join(low_labels, lang), studs(low_n, lang)))
        elif len(lows) == 1:
            text += ", sedangkan partisipasi terendah berasal dari %s (%d mahasiswa)" % (low_labels[0], low_n)
        else:
            text += (", sedangkan partisipasi terendah berasal dari %s (masing-masing %d mahasiswa)"
                     % (join(low_labels, lang), low_n))
    return text + "."


def mix_sentence(sub, n_stu, lang):
    """Komposisi jenis aktivitas di dalam satu program studi."""
    en = lang == "en"
    mix = ranked(sub.groupby("Jenis Aktivitas")["NIM"].nunique())
    if len(mix) == 1:
        if en:
            if n_stu == 1:
                return "This student took part in %s." % jenis_long(mix[0][0], lang)
            return "All of them took part in %s." % jenis_long(mix[0][0], lang)
        return "Seluruhnya mengikuti %s." % jenis_long(mix[0][0], lang)
    top_n = mix[0][1]
    tops = []
    rest = []
    for j, n in mix:
        if n == top_n:
            tops.append(jenis_long(j, lang))
        else:
            rest.append("%s (%d)" % (jenis_long(j, lang), n))
    if len(tops) > 1:
        if en:
            text = "The most common activity types were %s (%s each)" % (join(tops, lang), studs(top_n, lang))
        else:
            text = ("Jenis aktivitas yang paling banyak diikuti adalah %s (masing-masing %d mahasiswa)"
                    % (join(tops, lang), top_n))
    elif top_n * 2 > n_stu:
        if en:
            text = "Most took part in %s (%s, %s)" % (tops[0], studs(top_n, lang), pct(top_n, n_stu, lang))
        else:
            text = "Mayoritas mengikuti %s (%d mahasiswa, %s)" % (tops[0], top_n, pct(top_n, n_stu, lang))
    elif en:
        text = "The most common activity type was %s (%s)" % (tops[0], studs(top_n, lang))
    else:
        text = "Jenis aktivitas yang paling banyak diikuti adalah %s (%d mahasiswa)" % (tops[0], top_n)
    if rest:
        text += (", followed by %s" if en else ", diikuti %s") % join(rest, lang)
    return text + "."


def composition_sentence(comp, lang):
    """Asal program studi peserta satu jenis aktivitas. comp: list (prodi, n) urut menurun."""
    en = lang == "en"
    if len(comp) == 1:
        if en:
            return "All participants came from the %s study program." % prodi_label(comp[0][0])
        return "Seluruh pesertanya berasal dari Program Studi %s." % prodi_label(comp[0][0])
    all_equal = True
    for p, n in comp:
        if n != comp[0][1]:
            all_equal = False
    if all_equal and len(comp) <= 5:
        names = []
        for p, n in comp:
            names.append(prodi_label(p))
        if en:
            return ("Participants came from %d study programs, namely %s (%s each)."
                    % (len(comp), join(names, lang), studs(comp[0][1], lang)))
        return ("Peserta berasal dari %d program studi, yaitu %s (masing-masing %d mahasiswa)."
                % (len(comp), join(names, lang), comp[0][1]))
    shown = []
    cutoff = comp[min(2, len(comp) - 1)][1]
    for p, n in comp:
        if (len(comp) <= 4 or n >= cutoff) and len(shown) < 5:
            shown.append("%s (%d)" % (prodi_label(p), n))
    rest = len(comp) - len(shown)
    if rest == 0:
        if en:
            return "Participants came from %d study programs, namely %s." % (len(comp), join(shown, lang))
        return "Peserta berasal dari %d program studi, yaitu %s." % (len(comp), join(shown, lang))
    if en:
        return ("Participants came from %d study programs, mainly %s, plus %d other %s."
                % (len(comp), join(shown, lang), rest, plural(rest, "study program", "study programs")))
    return ("Peserta berasal dari %d program studi, terutama %s, serta %d program studi lainnya."
            % (len(comp), join(shown, lang), rest))


def summary_paragraphs(act, cats, flags, nim_info, meta, prodi_counts, jenis_counts, lang):
    en = lang == "en"
    n_act = len(act)
    n_stu = act["NIM"].nunique()
    n_prodi = len(prodi_counts)
    k_multi = len(nim_info["multi"])
    if en:
        p1 = ("In the %s, %s from %d %s took part in %s under the MBKM (Merdeka Belajar–Kampus Merdeka) scheme "
              "with course credit conversion."
              % (period_in_text(meta["periode"], lang), studs(n_stu, lang), n_prodi,
                 plural(n_prodi, "study program", "study programs"), acts(n_act, lang)))
        if meta["n_dropped"]:
            statuses = []
            for s in meta["drop_status"]:
                statuses.append(status_name(s, lang))
            p1 += (" These figures exclude %d %s with status %s."
                   % (meta["n_dropped"], plural(meta["n_dropped"], "record", "records"), join_or(statuses)))
        if k_multi == 0:
            p1 += " Each student is recorded in exactly one activity."
        else:
            p1 += (" The number of activities exceeds the number of students because %s %s recorded in more than "
                   "one activity (see Section 5)." % (studs(k_multi, lang), plural(k_multi, "is", "are")))
    else:
        p1 = ("Pada Periode Akademik %s, sebanyak %d mahasiswa dari %d program studi tercatat mengikuti %d "
              "aktivitas MBKM yang dikonversi ke mata kuliah." % (meta["periode"], n_stu, n_prodi, n_act))
        if meta["n_dropped"]:
            p1 += (" Angka ini sudah mengecualikan %d record berstatus %s."
                   % (meta["n_dropped"], join(meta["drop_status"], lang)))
        if k_multi == 0:
            p1 += " Setiap mahasiswa tercatat pada tepat satu aktivitas."
        else:
            p1 += (" Jumlah aktivitas lebih besar daripada jumlah mahasiswa karena %d mahasiswa tercatat pada lebih "
                   "dari satu aktivitas (lihat Bagian 5)." % k_multi)
    p2 = jenis_mix_sentence(jenis_counts, n_stu, lang) + " " + participation_sentence(prodi_counts, n_stu, lang)
    prof = group_profile(act, cats, flags)
    n_partners = len(prof["partners"])
    parts = []
    if prof["n_ext"]:
        if en:
            parts.append("%s (%s) %s carried out with %d external %s"
                         % (acts(prof["n_ext"], lang), pct(prof["n_ext"], n_act, lang), was(prof["n_ext"]),
                            n_partners, plural(n_partners, "partner", "partners")))
        else:
            parts.append("%d aktivitas (%s) dilaksanakan bersama %d mitra eksternal"
                         % (prof["n_ext"], pct(prof["n_ext"], n_act, lang), n_partners))
    if prof["n_int"]:
        if en:
            parts.append("%s (%s) %s carried out internally at i3L"
                         % (acts(prof["n_int"], lang), pct(prof["n_int"], n_act, lang), was(prof["n_int"])))
        else:
            parts.append("%d aktivitas (%s) dilaksanakan secara internal di i3L"
                         % (prof["n_int"], pct(prof["n_int"], n_act, lang)))
    if prof["n_bad"]:
        if en:
            parts.append("%s %s a valid partner name" % (acts(prof["n_bad"], lang), plural(prof["n_bad"], "lacks", "lack")))
        else:
            parts.append("%d aktivitas belum mencantumkan mitra yang valid" % prof["n_bad"])
    p3 = status_sentence(prof["status"], n_act, lang)
    if parts:
        p3 += (" A total of " if en else " Sebanyak ") + join(parts, lang) + "."
    if prof["n_ext"]:
        p3 += " " + mou_sentence(prof, lang)
    if en:
        p3 += (" In total, %s credits (SKS) were converted, an average of %s credits per activity."
               % (num(prof["sks_total"], lang), num(prof["sks_avg"], lang)))
    else:
        p3 += (" Secara total, %s SKS dikonversi, rata-rata %s SKS per aktivitas."
               % (num(prof["sks_total"], lang), num(prof["sks_avg"], lang)))
    return [p1, p2, p3], prof


def prodi_story(act, cats, flags, mk_names, prodi_counts, n_stu_all, lang):
    stories = []
    for p, n_stu in prodi_counts:
        sub = act[act["Program Studi"] == p]
        prof = group_profile(sub, cats, flags)
        rank = rank_clause(p, prodi_counts, "prodi", lang)
        if lang == "en":
            lead = "One student" if n_stu == 1 else "A total of %d students" % n_stu
            first = ("%s from the %s study program (%s of all participants) took part in %d MBKM %s%s."
                     % (lead, prodi_label(p), pct(n_stu, n_stu_all, lang), prof["n_act"],
                        plural(prof["n_act"], "activity", "activities"), rank))
        else:
            first = ("Sebanyak %d mahasiswa Program Studi %s (%s dari seluruh peserta) mengikuti %d aktivitas MBKM%s."
                     % (n_stu, prodi_label(p), pct(n_stu, n_stu_all, lang), prof["n_act"], rank))
        para1 = " ".join([first, mix_sentence(sub, n_stu, lang), status_sentence(prof["status"], prof["n_act"], lang)])
        para2 = " ".join(partner_sentences(prof, lang) + [conversion_sentence(prof, mk_names, lang),
                                                          data_note(prof, "prodi", lang)])
        stories.append((prodi_label(p), [para1, para2]))
    return stories


def jenis_story(act, cats, flags, mk_names, jenis_counts, n_stu_all, lang):
    stories = []
    for j, n_stu in jenis_counts:
        sub = act[act["Jenis Aktivitas"] == j]
        prof = group_profile(sub, cats, flags)
        rank = rank_clause(j, jenis_counts, "jenis", lang)
        via = ""
        if lang == "en":
            if prof["n_act"] != n_stu:
                via = " across %s" % acts(prof["n_act"], lang)
            first = ("%s involved %s (%s of all participants)%s%s."
                     % (jenis_long(j, lang), studs(n_stu, lang), pct(n_stu, n_stu_all, lang), via, rank))
        else:
            if prof["n_act"] != n_stu:
                via = " melalui %d aktivitas" % prof["n_act"]
            first = ("%s diikuti oleh %d mahasiswa (%s dari seluruh peserta)%s%s."
                     % (jenis_long(j, lang), n_stu, pct(n_stu, n_stu_all, lang), via, rank))
        comp = ranked(sub.groupby("Program Studi")["NIM"].nunique())
        para1 = " ".join([first, composition_sentence(comp, lang), status_sentence(prof["status"], prof["n_act"], lang)])
        para2 = " ".join(partner_sentences(prof, lang) + [conversion_sentence(prof, mk_names, lang),
                                                          data_note(prof, "jenis", lang)])
        stories.append((jenis_long(j, lang), [para1, para2]))
    return stories


def quality_section(act, cats, flags, nim_info, meta, lang):
    en = lang == "en"
    n_act = len(act)
    counts = {}
    for k in FLAG_COLS:
        counts[k] = int(flags[k].sum())
    n_issue = int(pd.concat([flags[k] for k in RECORD_FLAGS], axis=1).any(axis=1).sum())
    if en:
        wording = [("F3 Tanpa Pembimbing", "no supervisor recorded for %s"),
                   ("F1 Tanpa MK", "no converted course for %s"),
                   ("F2 MK Berlebih", "a course count above the threshold for %s"),
                   ("F5 Mitra Tidak Valid", "an invalid partner name for %s")]
    else:
        wording = [("F3 Tanpa Pembimbing", "dosen pembimbing kosong pada %s"),
                   ("F1 Tanpa MK", "MK konversi kosong pada %s"),
                   ("F2 MK Berlebih", "jumlah MK di atas ambang pada %s"),
                   ("F5 Mitra Tidak Valid", "nama mitra tidak valid pada %s")]
    details = []
    for k, text in wording:
        if counts[k]:
            details.append(text % acts(counts[k], lang))
    if n_issue and en:
        p1 = ("Automated checks found %s (%s) with at least one record-level data issue (F1, F2, F3, or F5): %s."
              % (acts(n_issue, lang), pct(n_issue, n_act, lang), join(details, lang)))
    elif n_issue:
        p1 = ("Pemeriksaan otomatis menemukan %d aktivitas (%s) dengan setidaknya satu isu input per-record "
              "(F1, F2, F3, atau F5): %s." % (n_issue, pct(n_issue, n_act, lang), join(details, lang)))
    elif en:
        p1 = "Automated checks found no record-level data issues (F1, F2, F3, or F5)."
    else:
        p1 = "Pemeriksaan otomatis tidak menemukan isu input per-record (F1, F2, F3, atau F5)."
    if counts["F4 Tanpa Penguji"]:
        if en:
            p1 += " The examiner field is also empty for %s." % acts(counts["F4 Tanpa Penguji"], lang)
        else:
            p1 += " Kolom dosen penguji juga masih kosong pada %d aktivitas." % counts["F4 Tanpa Penguji"]
    n6_ext = 0
    if counts["F6 Tanpa MoU"]:
        c6 = cats[flags["F6 Tanpa MoU"]]
        n6_int = int((c6 == CAT_INT).sum())
        n6_ext = int((c6 == CAT_NOMOU).sum())
        n6_bad = int((c6 == CAT_BAD).sum())
        parts = []
        if en:
            if n6_int:
                parts.append("%d %s" % (n6_int, plural(n6_int, "is an internal i3L activity (no MoU needed)",
                                                       "are internal i3L activities (no MoU needed)")))
            if n6_ext:
                parts.append("%d %s" % (n6_ext, plural(n6_ext, "involves an external institution",
                                                       "involve external institutions")))
            if n6_bad:
                parts.append("%d %s" % (n6_bad, plural(n6_bad, "lacks a valid partner name",
                                                       "lack a valid partner name")))
            p1 += (" A total of %s %s recorded with a partner that has no MoU; of these, %s."
                   % (acts(counts["F6 Tanpa MoU"], lang), plural(counts["F6 Tanpa MoU"], "is", "are"),
                      join(parts, lang)))
        else:
            if n6_int:
                parts.append("%d merupakan aktivitas internal i3L sehingga status MoU tidak relevan" % n6_int)
            if n6_ext:
                parts.append("%d bermitra dengan institusi eksternal" % n6_ext)
            if n6_bad:
                parts.append("%d tidak mencantumkan mitra yang valid" % n6_bad)
            p1 += (" Sebanyak %d aktivitas tercatat bermitra dengan institusi yang belum memiliki MoU; di antaranya, "
                   "%s." % (counts["F6 Tanpa MoU"], join(parts, lang)))

    multi = nim_info["multi"]
    if len(multi) == 0:
        if en:
            p2 = ("The student ID (NIM) check shows that every NIM appears in exactly one activity, so the number of "
                  "activities equals the number of students.")
        else:
            p2 = ("Pemeriksaan NIM menunjukkan bahwa setiap NIM tercatat pada tepat satu aktivitas, sehingga jumlah "
                  "aktivitas sama dengan jumlah mahasiswa.")
    else:
        items = []
        dup_mk = False
        for _, r in multi.head(5).iterrows():
            acts_list = []
            for _, a in act[act["NIM"] == r["NIM"]].iterrows():
                acts_list.append("%s (%s)" % (jenis_long(a["Jenis Aktivitas"], lang),
                                              status_name(a["Status Aktivitas"], lang)))
            if en:
                text = "%s (%s) in %s" % (r["Nama"], r["NIM"], join(acts_list, lang))
                if r["MK Sama"]:
                    text += ", converting the same course (%s)" % r["MK Sama"]
                    dup_mk = True
            else:
                text = "%s (%s) pada %s" % (r["Nama"], r["NIM"], join(acts_list, lang))
                if r["MK Sama"]:
                    text += " dengan MK yang sama (%s)" % r["MK Sama"]
                    dup_mk = True
            items.append(text)
        if en:
            p2 = ("The student ID (NIM) check found %s recorded in more than one activity: %s."
                  % (studs(len(multi), lang), "; ".join(items)))
            if dup_mk:
                p2 += (" Activities that convert the same course are most likely duplicates or resubmissions and "
                       "should be verified.")
            if len(multi) > 5:
                p2 += " The full list is available in the Cek NIM tab and the Excel file."
        else:
            p2 = ("Pemeriksaan NIM menemukan %d mahasiswa yang tercatat pada lebih dari satu aktivitas: %s."
                  % (len(multi), "; ".join(items)))
            if dup_mk:
                p2 += (" Aktivitas yang mengonversi MK yang sama kemungkinan besar merupakan duplikasi atau pengajuan "
                       "ulang sehingga perlu diverifikasi.")
            if len(multi) > 5:
                p2 += " Daftar lengkap tersedia pada tab Cek NIM dan file Excel."
    if meta["n_split"]:
        if en:
            lead = "One record" if meta["n_split"] == 1 else "A total of %d records" % meta["n_split"]
            p2 += (" %s that had been split because lecturer names appeared in a different order in the SIAKAD "
                   "export %s merged automatically." % (lead, was(meta["n_split"])))
        else:
            p2 += (" Sebanyak %d record yang terpecah karena urutan nama dosen berbeda di ekspor SIAKAD telah "
                   "digabung secara otomatis." % meta["n_split"])

    actions = []
    if en:
        if counts["F3 Tanpa Pembimbing"]:
            actions.append("Complete the supervisor field for %s." % acts(counts["F3 Tanpa Pembimbing"], lang))
        if counts["F4 Tanpa Penguji"]:
            actions.append("Complete the examiner field for %s." % acts(counts["F4 Tanpa Penguji"], lang))
        if counts["F1 Tanpa MK"]:
            actions.append("Check %s with no converted course." % acts(counts["F1 Tanpa MK"], lang))
        if counts["F2 MK Berlebih"]:
            actions.append("Verify %s with a course count above the threshold." % acts(counts["F2 MK Berlebih"], lang))
        if counts["F5 Mitra Tidak Valid"]:
            actions.append("Correct the partner name for %s." % acts(counts["F5 Mitra Tidak Valid"], lang))
        if len(multi):
            actions.append("Verify %d %s recorded in more than one activity." % (len(multi), plural(len(multi), "NIM", "NIMs")))
        if n6_ext:
            actions.append("Follow up on MoUs for %s with external partners that have no MoU." % acts(n6_ext, lang))
    else:
        if counts["F3 Tanpa Pembimbing"]:
            actions.append("Melengkapi dosen pembimbing pada %d aktivitas." % counts["F3 Tanpa Pembimbing"])
        if counts["F4 Tanpa Penguji"]:
            actions.append("Melengkapi dosen penguji pada %d aktivitas." % counts["F4 Tanpa Penguji"])
        if counts["F1 Tanpa MK"]:
            actions.append("Memeriksa %d aktivitas yang belum memiliki MK konversi." % counts["F1 Tanpa MK"])
        if counts["F2 MK Berlebih"]:
            actions.append("Memverifikasi %d aktivitas dengan jumlah MK di atas ambang." % counts["F2 MK Berlebih"])
        if counts["F5 Mitra Tidak Valid"]:
            actions.append("Memperbaiki nama mitra pada %d aktivitas." % counts["F5 Mitra Tidak Valid"])
        if len(multi):
            actions.append("Memverifikasi %d NIM yang tercatat pada lebih dari satu aktivitas." % len(multi))
        if n6_ext:
            actions.append("Menindaklanjuti MoU kerja sama untuk %d aktivitas bermitra eksternal yang belum "
                           "didukung MoU." % n6_ext)
    return [p1, p2], actions


def method_paragraph(meta, lang):
    if lang == "en":
        source = ""
        if meta["source_note"]:
            source = " (print details are shown on the first page)"
        if meta["drop_status"]:
            statuses = []
            for s in meta["drop_status"]:
                statuses.append(status_name(s, lang))
            drop = ("Records with status %s (%s) are excluded from the analysis."
                    % (join_or(statuses), ", ".join(meta["drop_status"])))
        else:
            drop = "No activity status is excluded."
        return ("Data source: the SIAKAD export \"Laporan Aktivitas MBKM dan Mata Kuliah Konversi\"%s. Each row of "
                "the export represents one activity–course pair; rows are merged into one activity per NIM based on "
                "activity type, partner, dates, title, and activity status. Students are counted as unique NIMs, "
                "while the charts count activities. %s Partners are classified as internal when the name refers to "
                "i3L; MoU status is taken from the Status Mitra column. Threshold F2: more than %d courses (Student "
                "Exchange: more than %d courses). Study program names are shown as recorded in SIAKAD. Status terms: "
                "Completed = Selesai, Under evaluation = Evaluasi, Submitted = Diajukan, Rejected = Ditolak, "
                "Cancelled = Dibatalkan." % (source, drop, meta["mk_overload"], meta["mk_overload_px"]))
    source = ""
    if meta["source_note"]:
        source = " (keterangan cetak tercantum di halaman pertama)"
    if meta["drop_status"]:
        drop = "Record berstatus %s dikecualikan dari analisis." % join(meta["drop_status"], lang)
    else:
        drop = "Tidak ada status aktivitas yang dikecualikan."
    return ("Data bersumber dari ekspor SIAKAD \"Laporan Aktivitas MBKM dan Mata Kuliah Konversi\"%s. "
            "Setiap baris ekspor mewakili satu pasangan aktivitas dan MK konversi; baris-baris tersebut digabung "
            "menjadi satu aktivitas per NIM berdasarkan jenis, mitra, tanggal, judul, dan status aktivitas. "
            "Jumlah mahasiswa dihitung sebagai NIM unik, sedangkan grafik menggunakan jumlah aktivitas. %s "
            "Mitra dikategorikan internal bila namanya merujuk ke i3L; status MoU diambil dari kolom Status Mitra. "
            "Ambang F2 adalah lebih dari %d MK (Pertukaran Pelajar: lebih dari %d MK)."
            % (source, drop, meta["mk_overload"], meta["mk_overload_px"]))


def pending_paragraphs(act, meta, lang):
    """Narasi + tabel mahasiswa yang aktivitasnya belum Selesai dan record yang dikecualikan filter."""
    en = lang == "en"
    pending, excluded = not_done_tables(act, meta.get("act_all"), meta["drop_status"], lang)
    paras = []
    if len(pending) == 0:
        paras.append("All analysed activities have been completed." if en
                     else "Semua aktivitas yang dianalisis telah berstatus Selesai.")
    else:
        raw_pending = act[act["Status Aktivitas"] != "Selesai"]
        n_rec = len(raw_pending)
        n_stu = raw_pending["NIM"].nunique()
        parts = []
        for s, n in sorted(raw_pending["Status Aktivitas"].value_counts().items(), key=lambda x: status_rank(x[0])):
            if en:
                parts.append("%d %s %s" % (n, plural(n, "is", "are"), status_state(s)))
            else:
                parts.append("%d berstatus %s" % (n, s))
        by_prodi = []
        for p, n in ranked(raw_pending.groupby("Program Studi").size()):
            by_prodi.append("%s (%d)" % (prodi_label(p), n))
        if en:
            who = acts(n_rec, lang) if n_rec == n_stu else "%s belonging to %s" % (acts(n_rec, lang), studs(n_stu, lang))
            text = ("A total of %s %s not yet been completed: %s. They are listed in Table 1 so that study programs "
                    "and supervisors can follow up. By study program: %s."
                    % (who, plural(n_rec, "has", "have"), join(parts, lang), join(by_prodi, lang)))
        else:
            if n_rec == n_stu:
                who = "%d aktivitas mahasiswa" % n_rec
            else:
                who = "%d aktivitas milik %d mahasiswa" % (n_rec, n_stu)
            text = ("Sebanyak %s belum berstatus Selesai: %s. Daftarnya tercantum pada Tabel 1 agar dapat "
                    "ditindaklanjuti oleh program studi dan dosen pembimbing. Menurut program studi: %s."
                    % (who, join(parts, lang), join(by_prodi, lang)))
        paras.append(text)
    if len(excluded):
        other_col = excluded.columns[-1]
        with_other = int((excluded[other_col] != "-").sum())
        no_other = excluded[excluded[other_col] == "-"]["NIM"].nunique()
        statuses = []
        for s in meta["drop_status"]:
            statuses.append(status_name(s, lang))
        if en:
            text = ("In addition, %d %s with status %s %s excluded from the analysis (Table 2). Of these, %d %s "
                    "to students who still have another activity on record, while %s %s no other MBKM activity "
                    "in this period."
                    % (len(excluded), plural(len(excluded), "record", "records"), join_or(statuses),
                       was(len(excluded)), with_other, plural(with_other, "belongs", "belong"),
                       studs(no_other, lang), plural(no_other, "has", "have")))
        else:
            text = ("Selain itu, %d record berstatus %s dikecualikan dari analisis (Tabel 2). Sebanyak %d di "
                    "antaranya milik mahasiswa yang masih memiliki aktivitas lain yang tercatat, sedangkan %d "
                    "mahasiswa tidak memiliki aktivitas MBKM lain pada periode ini."
                    % (len(excluded), join(meta["drop_status"], lang), with_other, no_other))
        paras.append(text)
    return paras, pending, excluded


def narrative_content(act, raw, flags, nim_info, meta, cats, lang):
    """Seluruh teks laporan naratif dalam satu bahasa (dipakai oleh PDF, Word, dan tab Narasi)."""
    T = NARR_TEXT[lang]
    mk_names = mk_name_map(raw)
    n_act = len(act)
    n_stu = act["NIM"].nunique()
    prodi_counts = ranked(act.groupby("Program Studi")["NIM"].nunique())
    jenis_counts = ranked(act.groupby("Jenis Aktivitas")["NIM"].nunique())
    summary, overall = summary_paragraphs(act, cats, flags, nim_info, meta, prodi_counts, jenis_counts, lang)
    quality, actions = quality_section(act, cats, flags, nim_info, meta, lang)
    n_done = int(overall["status"].get("Selesai", 0))
    mou_share = "-"
    if overall["n_ext"]:
        mou_share = pct(overall["n_mou"], overall["n_ext"], lang)
    kpi = [(T["k_students"], num(n_stu, lang)), (T["k_acts"], num(n_act, lang)),
           (T["k_prodi"], num(len(prodi_counts), lang)), (T["k_done"], pct(n_done, n_act, lang)),
           (T["k_ext"], num(len(overall["partners"]), lang)), (T["k_mou"], mou_share)]
    pending_text, pending_df, excluded_df = pending_paragraphs(act, meta, lang)
    return {
        "pending": pending_text,
        "pending_df": pending_df,
        "excluded_df": excluded_df,
        "lang": lang,
        "kpi": kpi,
        "summary": summary,
        "prodi_intro": T["prodi_intro"],
        "prodi": prodi_story(act, cats, flags, mk_names, prodi_counts, n_stu, lang),
        "jenis_intro": T["jenis_intro"],
        "jenis": jenis_story(act, cats, flags, mk_names, jenis_counts, n_stu, lang),
        "quality": quality,
        "actions": actions,
        "method": method_paragraph(meta, lang),
        "matrix": student_matrix(act, lang),
    }


# ---------------------------------------------------------------- laporan naratif: grafik
def contrast_text(hex_color):
    """Putih atau tinta, mana yang kontrasnya lebih tinggi di atas warna isian."""
    channels = []
    for i in (1, 3, 5):
        c = int(hex_color[i:i + 2], 16) / 255.0
        if c <= 0.03928:
            channels.append(c / 12.92)
        else:
            channels.append(((c + 0.055) / 1.055) ** 2.4)
    lum = 0.2126 * channels[0] + 0.7152 * channels[1] + 0.0722 * channels[2]
    if 1.05 / (lum + 0.05) >= (lum + 0.05) / 0.0534:
        return "#ffffff"
    return INK


def stacked_barh_png(table, colors, end_labels, xlabel):
    """Batang horizontal bertumpuk. table: index = label batang (atas ke bawah), kolom = segmen."""
    keep = []
    for col in table.columns:
        if table[col].sum() > 0:
            keep.append(col)
    table = table[keep]
    n = len(table.index)
    fig = plt.figure(figsize=(6.8, 0.95 + 0.36 * n))
    canvas = FigureCanvasAgg(fig)  # kanvas Agg eksplisit: aman saat Streamlit menjalankan ulang skrip
    ypos = []
    for i in range(n):
        ypos.append(n - 1 - i)
    lefts = [0.0] * n
    seg_labels = []
    for col in table.columns:
        vals = []
        for cat in table.index:
            vals.append(float(table.loc[cat, col]))
        plt.barh(ypos, vals, left=lefts, height=0.58, color=colors[col], edgecolor="white",
                 linewidth=1.4, label=col, zorder=3)
        ink = contrast_text(colors[col])
        for y, left, v in zip(ypos, lefts, vals):
            if v > 0:
                label = plt.text(left + v / 2.0, y, "%d" % v, ha="center", va="center",
                                 fontsize=7.5, color=ink, zorder=4)
                seg_labels.append((label, left, v, y))
        new_lefts = []
        for left, v in zip(lefts, vals):
            new_lefts.append(left + v)
        lefts = new_lefts
    xmax = max(lefts) if n else 1.0
    for y, total, cat in zip(ypos, lefts, table.index):
        plt.text(total + xmax * 0.012, y, end_labels[cat], ha="left", va="center",
                 fontsize=8, fontweight="bold", color=INK, zorder=4)
    plt.yticks(ypos, list(table.index), fontsize=8.5, color=INK2)
    plt.xticks(fontsize=7.5, color=MUTED)
    plt.xlim(0, xmax * 1.2)
    plt.ylim(-0.6, n - 0.4)
    plt.xlabel(xlabel, fontsize=8, color=MUTED, labelpad=4)
    plt.grid(axis="x", color=GRIDC, linewidth=0.6, zorder=0)
    ax = plt.gca()
    ax.set_axisbelow(True)
    ax.xaxis.set_major_locator(MaxNLocator(integer=True, nbins=6))
    for side in ["top", "right", "left"]:
        ax.spines[side].set_visible(False)
    ax.spines["bottom"].set_color(AXISC)
    ax.spines["bottom"].set_linewidth(0.8)
    ax.tick_params(axis="y", length=0, pad=6)
    ax.tick_params(axis="x", length=0, pad=3)
    plt.legend(loc="lower left", bbox_to_anchor=(0.0, 1.0), ncol=len(table.columns), frameon=False,
               fontsize=8, handlelength=0.9, handleheight=0.9, handletextpad=0.5, columnspacing=1.4,
               borderaxespad=0.3, labelcolor=INK2)
    plt.tight_layout()
    # label di dalam segmen hanya dipertahankan bila muat (diukur setelah tata letak final)
    canvas.draw()
    renderer = canvas.get_renderer()
    for label, left, v, y in seg_labels:
        x0 = ax.transData.transform((left, y))[0]
        x1 = ax.transData.transform((left + v, y))[0]
        if label.get_window_extent(renderer).width + 8 > (x1 - x0):
            label.remove()
    buf = io.BytesIO()
    plt.savefig(buf, format="png", dpi=200, bbox_inches="tight", pad_inches=0.06, facecolor="white")
    plt.close(fig)
    return buf.getvalue()


def ordered_jenis(values):
    order = []
    for j in JENIS_COLORS:
        if j in values:
            order.append(j)
    for j in sorted(values):
        if j not in order:
            order.append(j)
    return order


def jenis_colors(order, lang):
    colors = {}
    extra = 0
    for j in order:
        if j in JENIS_COLORS:
            colors[jenis_short(j, lang)] = JENIS_COLORS[j]
        else:
            colors[jenis_short(j, lang)] = EXTRA_COLORS[extra % len(EXTRA_COLORS)]
            extra += 1
    return colors


def count_table(rows, cols, row_order, col_order, row_label, col_label):
    table = pd.crosstab(rows, cols).reindex(index=row_order, columns=col_order, fill_value=0)
    names = []
    for r in table.index:
        names.append(row_label(r))
    table.index = names
    names = []
    for c in table.columns:
        names.append(col_label(c))
    table.columns = names
    return table


@st.cache_data(show_spinner=False)
def narrative_charts(act, cats, lang):
    """Empat grafik laporan naratif sebagai PNG (bytes), label sesuai bahasa."""
    T = NARR_TEXT[lang]

    def jenis_lbl(j):
        return jenis_short(j, lang)

    def status_lbl(s):
        return status_name(s, lang)

    def cat_lbl(c):
        return cat_name(c, lang)

    charts = {}
    prodi_order = []
    for p, n in ranked(act.groupby("Program Studi")["NIM"].nunique()):
        prodi_order.append(p)
    jenis_order = []
    for j, n in ranked(act.groupby("Jenis Aktivitas")["NIM"].nunique()):
        jenis_order.append(j)
    jenis_stack = ordered_jenis(set(act["Jenis Aktivitas"]))
    status_order = []
    for s in STATUS_ORDER:
        if s in set(act["Status Aktivitas"]):
            status_order.append(s)
    for s in sorted(set(act["Status Aktivitas"])):
        if s not in status_order:
            status_order.append(s)
    status_colors = {}
    for s in status_order:
        status_colors[status_lbl(s)] = STATUS_COLORS.get(s, "#c3c2b7")
    partner_colors = {}
    for c in PARTNER_ORDER:
        partner_colors[cat_lbl(c)] = PARTNER_COLORS[c]

    prodi_totals = {}
    prodi_plain = {}
    for p in prodi_order:
        sub = act[act["Program Studi"] == p]
        prodi_plain[prodi_label(p)] = "%d" % len(sub)
        if len(sub) != sub["NIM"].nunique():
            prodi_totals[prodi_label(p)] = T["students_paren"] % (len(sub), sub["NIM"].nunique())
        else:
            prodi_totals[prodi_label(p)] = "%d" % len(sub)
    jenis_totals = {}
    jenis_shares = {}
    for j in jenis_order:
        n = int((act["Jenis Aktivitas"] == j).sum())
        jenis_totals[jenis_lbl(j)] = "%d" % n
        jenis_shares[jenis_lbl(j)] = "%d (%s)" % (n, pct(n, len(act), lang))

    t = count_table(act["Program Studi"], act["Jenis Aktivitas"], prodi_order, jenis_stack, prodi_label, jenis_lbl)
    charts["prodi_jenis"] = stacked_barh_png(t, jenis_colors(jenis_stack, lang), prodi_totals, T["x_label"])
    t = count_table(act["Program Studi"], cats, prodi_order, PARTNER_ORDER, prodi_label, cat_lbl)
    charts["prodi_mitra"] = stacked_barh_png(t, partner_colors, prodi_plain, T["x_label"])
    t = count_table(act["Jenis Aktivitas"], act["Status Aktivitas"], jenis_order, status_order, jenis_lbl, status_lbl)
    charts["jenis_status"] = stacked_barh_png(t, status_colors, jenis_shares, T["x_label"])
    t = count_table(act["Jenis Aktivitas"], cats, jenis_order, PARTNER_ORDER, jenis_lbl, cat_lbl)
    charts["jenis_mitra"] = stacked_barh_png(t, partner_colors, jenis_totals, T["x_label"])
    return charts


# ---------------------------------------------------------------- laporan naratif: PDF
def register_fonts():
    """DejaVu Sans (terpasang bersama matplotlib) agar semua karakter tercetak; cadangan Helvetica."""
    from reportlab.pdfbase import pdfmetrics
    from reportlab.pdfbase.ttfonts import TTFont
    folder = os.path.join(matplotlib.get_data_path(), "fonts", "ttf")
    faces = {"DejaVu": "DejaVuSans.ttf", "DejaVu-Bold": "DejaVuSans-Bold.ttf",
             "DejaVu-Oblique": "DejaVuSans-Oblique.ttf", "DejaVu-BoldOblique": "DejaVuSans-BoldOblique.ttf"}
    try:
        registered = pdfmetrics.getRegisteredFontNames()
        for name, filename in faces.items():
            if name not in registered:
                pdfmetrics.registerFont(TTFont(name, os.path.join(folder, filename)))
        pdfmetrics.registerFontFamily("DejaVu", normal="DejaVu", bold="DejaVu-Bold",
                                      italic="DejaVu-Oblique", boldItalic="DejaVu-BoldOblique")
        return "DejaVu", "DejaVu-Bold", "DejaVu-Oblique"
    except Exception:  # noqa
        return "Helvetica", "Helvetica-Bold", "Helvetica-Oblique"


def build_narrative_pdf(content, charts, meta):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_LEFT, TA_CENTER, TA_JUSTIFY
    from reportlab.lib.utils import ImageReader
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle,
                                    HRFlowable, KeepTogether, Image, ListFlowable, ListItem)

    lang = content["lang"]
    T = NARR_TEXT[lang]
    captions = FIG_CAPTIONS[lang]
    period = period_display(meta["periode"], lang)
    font, font_b, font_i = register_fonts()
    navy = colors.HexColor(NAVY)
    light = colors.HexColor("#EEF1F7")
    grey = colors.HexColor("#D9D9D9")
    mute = colors.HexColor("#666666")
    ink = colors.HexColor(INK)
    page_w = A4[0] - 36 * mm

    def esc(t):
        return str(t).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

    base = getSampleStyleSheet()
    H1 = ParagraphStyle("NH1", parent=base["Title"], fontName=font_b, fontSize=16.5, leading=20,
                        textColor=navy, alignment=TA_LEFT, spaceAfter=2)
    SUB = ParagraphStyle("NSUB", parent=base["Normal"], fontName=font, fontSize=8.2, leading=11, textColor=mute)
    H2 = ParagraphStyle("NH2", parent=base["Heading2"], fontName=font_b, fontSize=12, leading=15,
                        textColor=navy, spaceBefore=14, spaceAfter=6)
    H3 = ParagraphStyle("NH3", parent=base["Heading3"], fontName=font_b, fontSize=9.8, leading=13,
                        textColor=ink, spaceBefore=9, spaceAfter=3)
    BODY = ParagraphStyle("NBODY", parent=base["Normal"], fontName=font, fontSize=8.8, leading=13,
                          textColor=ink, alignment=TA_JUSTIFY, spaceAfter=5)
    CAP = ParagraphStyle("NCAP", parent=base["Normal"], fontName=font_i, fontSize=7.5, leading=10,
                         textColor=mute, spaceBefore=2, spaceAfter=9)
    KPIV = ParagraphStyle("NKPIV", parent=base["Normal"], fontName=font_b, fontSize=14, leading=17,
                          textColor=navy, alignment=TA_CENTER)
    KPIL = ParagraphStyle("NKPIL", parent=base["Normal"], fontName=font, fontSize=7, leading=9,
                          textColor=mute, alignment=TA_CENTER)
    CELL = ParagraphStyle("NCELL", parent=base["Normal"], fontName=font, fontSize=8, leading=10.5)
    HEAD = ParagraphStyle("NHEAD", parent=CELL, fontName=font_b, fontSize=7.5, leading=9.5,
                          textColor=colors.white, alignment=TA_CENTER)

    def figure(key):
        png = charts[key]
        w, h = ImageReader(io.BytesIO(png)).getSize()
        img = Image(io.BytesIO(png), width=page_w, height=page_w * h / float(w))
        return KeepTogether([Spacer(1, 2), img, Paragraph(esc(captions[key]), CAP)])

    story = []
    story.append(Paragraph(esc(T["title"]), H1))
    story.append(Paragraph(esc(T["institute"]), SUB))
    story.append(Paragraph(esc(T["period"] % period), SUB))
    if meta["source_note"]:
        story.append(Paragraph(esc(source_note_text(meta["source_note"], lang)).replace("|", "&nbsp;|&nbsp;"), SUB))
    story.append(Spacer(1, 3))
    story.append(HRFlowable(width="100%", thickness=1.2, color=navy, spaceAfter=8))

    values = []
    labels = []
    for label, value in content["kpi"]:
        values.append(Paragraph(esc(value), KPIV))
        labels.append(Paragraph(esc(label), KPIL))
    widths = []
    for _ in content["kpi"]:
        widths.append(page_w / len(content["kpi"]))
    tiles = Table([values, labels], colWidths=widths)
    tiles.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), light),
        ("LINEAFTER", (0, 0), (-2, -1), 3, colors.white),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, 0), 8), ("BOTTOMPADDING", (0, 0), (-1, 0), 1),
        ("TOPPADDING", (0, 1), (-1, 1), 1), ("BOTTOMPADDING", (0, 1), (-1, 1), 8),
    ]))
    story.append(tiles)

    story.append(Paragraph(esc(T["s1"]), H2))
    for p in content["summary"]:
        story.append(Paragraph(esc(p), BODY))
    story.append(figure("prodi_jenis"))

    story.append(Paragraph(esc(T["s2"]), H2))
    story.append(Paragraph(esc(content["prodi_intro"]), BODY))
    story.append(figure("prodi_mitra"))
    for i, (name, paras) in enumerate(content["prodi"], start=1):
        block = [Paragraph("2.%d %s" % (i, esc(name)), H3)]
        for p in paras:
            block.append(Paragraph(esc(p), BODY))
        story.append(KeepTogether(block))

    story.append(Paragraph(esc(T["s3"]), H2))
    story.append(Paragraph(esc(content["jenis_intro"]), BODY))
    story.append(figure("jenis_status"))
    story.append(figure("jenis_mitra"))
    for i, (name, paras) in enumerate(content["jenis"], start=1):
        block = [Paragraph("3.%d %s" % (i, esc(name)), H3)]
        for p in paras:
            block.append(Paragraph(esc(p), BODY))
        story.append(KeepTogether(block))

    block = [Paragraph(esc(T["s_pending"]), H2)]
    for p in content["pending"]:
        block.append(Paragraph(esc(p), BODY))
    story.append(KeepTogether(block))
    list_style = TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#BFBFBF")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, light]),
    ])
    for key, caption, widths, bg in (("pending_df", "pending_t1", [18, 40, 36, 24, 18, 38], navy),
                                     ("excluded_df", "pending_t2", [17, 33, 29, 20, 18, 31, 26],
                                      colors.HexColor(RED))):
        if len(content[key]):
            table = not_done_pdf_table(content[key], widths, CELL, font_b, list_style)
            table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), bg)]))
            story.append(Paragraph(esc(T[caption]), CAP))
            story.append(table)
            story.append(Spacer(1, 8))

    story.append(Paragraph(esc(T["s4"]), H2))
    for p in content["quality"]:
        story.append(Paragraph(esc(p), BODY))
    if content["actions"]:
        items = []
        for a in content["actions"]:
            items.append(ListItem(Paragraph(esc(a), BODY), leftIndent=12))
        story.append(KeepTogether([Paragraph(esc(T["actions"]), BODY),
                                   ListFlowable(items, bulletType="bullet", start="•", leftIndent=12,
                                                bulletFontName=font, bulletFontSize=8)]))

    story.append(Paragraph(esc(T["s5"]), H2))
    story.append(Paragraph(esc(content["method"]), BODY))

    matrix = content["matrix"]
    header = [Paragraph(esc(T["prodi_col"]), HEAD)]
    for c in matrix.columns:
        header.append(Paragraph(esc(c), HEAD))
    rows = [header]
    for idx, r in matrix.iterrows():
        row = [Paragraph(esc(idx), CELL)]
        for c in matrix.columns:
            row.append(str(int(r[c])))
        rows.append(row)
    widths = [50 * mm]
    for _ in matrix.columns:
        widths.append((page_w - 50 * mm) / len(matrix.columns))
    table = Table(rows, colWidths=widths, repeatRows=1)
    table.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, 0), navy),
        ("LEFTPADDING", (0, 0), (-1, 0), 3), ("RIGHTPADDING", (0, 0), (-1, 0), 3),
        ("FONTNAME", (0, 1), (-1, -1), font), ("FONTSIZE", (0, 1), (-1, -1), 8),
        ("ALIGN", (1, 1), (-1, -1), "CENTER"), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#BFBFBF")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, light]),
        ("BACKGROUND", (0, -1), (-1, -1), grey), ("FONTNAME", (1, -1), (-1, -1), font_b),
    ]))
    story.append(KeepTogether([Paragraph(esc(T["appendix"]), H2), Paragraph(esc(T["appendix_note"]), CAP), table]))

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 7.2)
        canvas.setFillColor(mute)
        canvas.drawString(18 * mm, 9 * mm, T["footer"] % period)
        canvas.drawRightString(A4[0] - 18 * mm, 9 * mm, T["page"] + str(doc.page))
        canvas.restoreState()

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm,
                            topMargin=16 * mm, bottomMargin=17 * mm, title=T["title"])
    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    buf.seek(0)
    return buf.getvalue()


# ---------------------------------------------------------------- laporan naratif: Word (.docx)
def docx_set_font(element, name):
    """Font tetap pada style/run (hapus atribut tema agar tidak ditimpa font tema Word)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    rpr = element.get_or_add_rPr()
    fonts = rpr.find(qn("w:rFonts"))
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    for att in ("w:asciiTheme", "w:hAnsiTheme", "w:eastAsiaTheme", "w:cstheme"):
        if fonts.get(qn(att)) is not None:
            del fonts.attrib[qn(att)]
    for att in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
        fonts.set(qn(att), name)


def docx_shade(cell, fill):
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    shd = OxmlElement("w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)
    cell._tc.get_or_add_tcPr().append(shd)


def xml_insert(parent, child, successors):
    """Sisipkan child sebelum elemen penerus pertama yang ada (urutan elemen wajib skema OOXML)."""
    from docx.oxml.ns import qn
    for tag in successors:
        found = parent.find(qn(tag))
        if found is not None:
            found.addprevious(child)
            return child
    parent.append(child)
    return child


def twips(length):
    return int(round(int(length) / 635.0))


def docx_table_layout(table, widths, borders):
    """Lebar kolom tetap (grid + tiap sel, satuan DXA) dan garis tabel. borders: dict sisi -> (ukuran, warna)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    table.autofit = False
    tbl_pr = table._tbl.tblPr
    total = 0
    for w in widths:
        total += twips(w)
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = xml_insert(tbl_pr, OxmlElement("w:tblW"),
                           ["w:jc", "w:tblCellSpacing", "w:tblInd", "w:tblBorders", "w:shd", "w:tblLayout",
                            "w:tblCellMar", "w:tblLook"])
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")
    edges = OxmlElement("w:tblBorders")
    for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
        el = OxmlElement("w:" + side)
        size, color = borders.get(side, (0, "auto"))
        if size:
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), str(size))
            el.set(qn("w:color"), color)
        else:
            el.set(qn("w:val"), "nil")
        edges.append(el)
    xml_insert(tbl_pr, edges, ["w:shd", "w:tblLayout", "w:tblCellMar", "w:tblLook", "w:tblCaption",
                               "w:tblDescription"])
    grid_cols = table._tbl.tblGrid.findall(qn("w:gridCol"))
    for col, w in zip(grid_cols, widths):
        col.set(qn("w:w"), str(twips(w)))
    for row in table.rows:
        for i, cell in enumerate(row.cells):
            cell.width = widths[i]


def docx_page_field(paragraph):
    """Sisipkan field PAGE (nomor halaman otomatis)."""
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    for kind in ("begin", "instr", "separate", "text", "end"):
        run = paragraph.add_run()
        if kind == "instr":
            instr = OxmlElement("w:instrText")
            instr.set(qn("xml:space"), "preserve")
            instr.text = " PAGE "
            run._r.append(instr)
        elif kind == "text":
            run.text = "1"
        else:
            fld = OxmlElement("w:fldChar")
            fld.set(qn("w:fldCharType"), kind)
            run._r.append(fld)


def docx_list_table(doc, df, widths_cm, header_fill):
    """Tabel Word untuk daftar mahasiswa (header berwarna, baris berselang, lebar kolom tetap)."""
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.table import WD_TABLE_ALIGNMENT
    table = doc.add_table(rows=len(df.index) + 1, cols=len(df.columns))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = []
    for w in widths_cm:
        widths.append(Cm(w))
    grid = (4, "BFBFBF")
    docx_table_layout(table, widths, {"top": grid, "left": grid, "bottom": grid, "right": grid,
                                      "insideH": grid, "insideV": grid})
    for i, c in enumerate(df.columns):
        cell = table.cell(0, i)
        docx_shade(cell, header_fill)
        par = cell.paragraphs[0]
        par.paragraph_format.space_after = Pt(0)
        run = par.add_run(str(c))
        run.bold = True
        run.font.size = Pt(8)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for r, (_, values) in enumerate(df.iterrows(), start=1):
        fill = "EEF1F7" if r % 2 == 0 else "FFFFFF"
        for i, c in enumerate(df.columns):
            cell = table.cell(r, i)
            docx_shade(cell, fill)
            par = cell.paragraphs[0]
            par.paragraph_format.space_after = Pt(0)
            par.paragraph_format.line_spacing = 1.0
            run = par.add_run(str(values[c]))
            run.font.size = Pt(8)
    return table


def build_narrative_docx(content, charts, meta):
    from docx import Document
    from docx.shared import Pt, Cm, Emu, RGBColor
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn

    lang = content["lang"]
    T = NARR_TEXT[lang]
    captions = FIG_CAPTIONS[lang]
    period = period_display(meta["periode"], lang)
    navy = RGBColor(0x1F, 0x38, 0x64)
    ink = RGBColor(0x0B, 0x0B, 0x0B)
    mute = RGBColor(0x66, 0x66, 0x66)
    font = "Calibri"

    doc = Document()
    section = doc.sections[0]
    section.page_width = Cm(21.0)
    section.page_height = Cm(29.7)
    section.left_margin = Cm(1.8)
    section.right_margin = Cm(1.8)
    section.top_margin = Cm(1.6)
    section.bottom_margin = Cm(1.7)
    usable = section.page_width - section.left_margin - section.right_margin

    normal = doc.styles["Normal"]
    docx_set_font(normal.element, font)
    normal.font.size = Pt(10.5)
    normal.font.color.rgb = ink
    normal.paragraph_format.space_after = Pt(6)
    normal.paragraph_format.line_spacing = 1.15
    lang_el = OxmlElement("w:lang")
    lang_el.set(qn("w:val"), T["doc_lang"])
    normal.element.get_or_add_rPr().append(lang_el)
    for level, size, color, before, after in ((1, 14, navy, 16, 6), (2, 11.5, ink, 10, 3)):
        style = doc.styles["Heading %d" % level]
        docx_set_font(style.element, font)
        style.font.size = Pt(size)
        style.font.bold = True
        style.font.italic = False
        style.font.color.rgb = color
        style.paragraph_format.space_before = Pt(before)
        style.paragraph_format.space_after = Pt(after)
        style.paragraph_format.keep_with_next = True
    bullet = doc.styles["List Bullet"]
    docx_set_font(bullet.element, font)
    bullet.font.size = Pt(10.5)

    def small_line(text, size=9, color=mute, after=0):
        par = doc.add_paragraph()
        run = par.add_run(text)
        run.font.size = Pt(size)
        run.font.color.rgb = color
        par.paragraph_format.space_after = Pt(after)
        return par

    def body(text):
        par = doc.add_paragraph(text)
        par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        return par

    def figure(key):
        pic = doc.add_paragraph()
        pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic.paragraph_format.keep_with_next = True
        pic.paragraph_format.space_after = Pt(2)
        pic.add_run().add_picture(io.BytesIO(charts[key]), width=usable)
        cap = small_line(captions[key], size=8.5, after=10)
        for run in cap.runs:
            run.italic = True

    # judul
    title = doc.add_paragraph()
    run = title.add_run(T["title"])
    run.bold = True
    run.font.size = Pt(20)
    run.font.color.rgb = navy
    title.paragraph_format.space_after = Pt(2)
    small_line(T["institute"])
    last = small_line(T["period"] % period)
    if meta["source_note"]:
        last = small_line(source_note_text(meta["source_note"], lang))
    border = OxmlElement("w:pBdr")
    bottom = OxmlElement("w:bottom")
    bottom.set(qn("w:val"), "single")
    bottom.set(qn("w:sz"), "12")
    bottom.set(qn("w:space"), "6")
    bottom.set(qn("w:color"), NAVY.lstrip("#"))
    border.append(bottom)
    xml_insert(last._p.get_or_add_pPr(), border,
               ["w:shd", "w:tabs", "w:suppressAutoHyphens", "w:kinsoku", "w:wordWrap", "w:overflowPunct",
                "w:topLinePunct", "w:autoSpaceDE", "w:autoSpaceDN", "w:bidi", "w:adjustRightInd", "w:snapToGrid",
                "w:spacing", "w:ind", "w:contextualSpacing", "w:mirrorIndents", "w:suppressOverlap", "w:jc",
                "w:textDirection", "w:textAlignment", "w:outlineLvl", "w:rPr", "w:sectPr", "w:pPrChange"])
    last.paragraph_format.space_after = Pt(12)

    # kartu angka utama
    kpi = content["kpi"]
    tiles = doc.add_table(rows=2, cols=len(kpi))
    tiles.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = []
    for _ in kpi:
        widths.append(Emu(int(usable / len(kpi))))
    docx_table_layout(tiles, widths, {"insideV": (24, "FFFFFF")})
    for i, (label, value) in enumerate(kpi):
        top = tiles.cell(0, i)
        bot = tiles.cell(1, i)
        for cell in (top, bot):
            docx_shade(cell, "EEF1F7")
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        par = top.paragraphs[0]
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.space_before = Pt(6)
        par.paragraph_format.space_after = Pt(0)
        run = par.add_run(value)
        run.bold = True
        run.font.size = Pt(15)
        run.font.color.rgb = navy
        par = bot.paragraphs[0]
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.space_after = Pt(6)
        run = par.add_run(label)
        run.font.size = Pt(8)
        run.font.color.rgb = mute

    doc.add_heading(T["s1"], level=1)
    for p in content["summary"]:
        body(p)
    figure("prodi_jenis")

    doc.add_heading(T["s2"], level=1)
    body(content["prodi_intro"])
    figure("prodi_mitra")
    for i, (name, paras) in enumerate(content["prodi"], start=1):
        doc.add_heading("2.%d %s" % (i, name), level=2)
        for p in paras:
            body(p)

    doc.add_heading(T["s3"], level=1)
    body(content["jenis_intro"])
    figure("jenis_status")
    figure("jenis_mitra")
    for i, (name, paras) in enumerate(content["jenis"], start=1):
        doc.add_heading("3.%d %s" % (i, name), level=2)
        for p in paras:
            body(p)

    doc.add_heading(T["s_pending"], level=1)
    for p in content["pending"]:
        body(p)
    for key, caption, widths, fill in (("pending_df", "pending_t1", [2.0, 4.0, 3.6, 2.4, 2.4, 3.0], NAVY),
                                       ("excluded_df", "pending_t2", [1.8, 3.2, 2.9, 2.0, 1.9, 2.6, 3.0], RED)):
        if len(content[key]):
            cap = small_line(T[caption], size=8.5, after=4)
            cap.paragraph_format.keep_with_next = True
            for run in cap.runs:
                run.italic = True
            docx_list_table(doc, content[key], widths, fill.lstrip("#"))
            doc.add_paragraph().paragraph_format.space_after = Pt(2)

    doc.add_heading(T["s4"], level=1)
    for p in content["quality"]:
        body(p)
    if content["actions"]:
        lead = doc.add_paragraph(T["actions"])
        lead.paragraph_format.keep_with_next = True
        for a in content["actions"]:
            doc.add_paragraph(a, style="List Bullet")

    doc.add_heading(T["s5"], level=1)
    body(content["method"])

    # lampiran: matriks
    matrix = content["matrix"]
    doc.add_heading(T["appendix"], level=1)
    note = small_line(T["appendix_note"], size=8.5, after=6)
    for run in note.runs:
        run.italic = True
    note.paragraph_format.keep_with_next = True
    table = doc.add_table(rows=len(matrix.index) + 1, cols=len(matrix.columns) + 1)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    first_w = Cm(5.2)
    other_w = Emu(int((usable - first_w) / len(matrix.columns)))
    col_widths = [first_w]
    for _ in matrix.columns:
        col_widths.append(other_w)
    grid = (4, "BFBFBF")
    docx_table_layout(table, col_widths, {"top": grid, "left": grid, "bottom": grid, "right": grid,
                                          "insideH": grid, "insideV": grid})
    headers = [T["prodi_col"]]
    for c in matrix.columns:
        headers.append(str(c))
    for i, text in enumerate(headers):
        cell = table.cell(0, i)
        docx_shade(cell, NAVY.lstrip("#"))
        cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
        par = cell.paragraphs[0]
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        par.paragraph_format.space_after = Pt(0)
        run = par.add_run(text)
        run.bold = True
        run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    for r, (idx, values) in enumerate(matrix.iterrows(), start=1):
        is_total = r == len(matrix.index)
        fill = "D9D9D9" if is_total else ("EEF1F7" if r % 2 == 0 else "FFFFFF")
        cells_text = [str(idx)]
        for c in matrix.columns:
            cells_text.append(str(int(values[c])))
        for i, text in enumerate(cells_text):
            cell = table.cell(r, i)
            docx_shade(cell, fill)
            cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
            par = cell.paragraphs[0]
            par.paragraph_format.space_after = Pt(0)
            if i > 0:
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = par.add_run(text)
            run.font.size = Pt(9)
            run.bold = is_total

    # footer: judul + nomor halaman (tab stop bawaan template dihapus agar nomor rata kanan)
    foot = section.footer.paragraphs[0]
    if "Footer" in doc.styles:
        doc.styles["Footer"].paragraph_format.tab_stops.clear_all()
    foot.paragraph_format.tab_stops.add_tab_stop(usable, WD_TAB_ALIGNMENT.RIGHT)
    run = foot.add_run(T["footer"] % period + "\t" + T["page"])
    run.font.size = Pt(8)
    run.font.color.rgb = mute
    docx_page_field(foot)
    for run in foot.runs:
        run.font.size = Pt(8)
        run.font.color.rgb = mute

    doc.core_properties.title = "%s - %s" % (T["title"], period)
    doc.core_properties.language = T["doc_lang"]
    zoom = doc.settings.element.find(qn("w:zoom"))
    if zoom is not None and zoom.get(qn("w:percent")) is None:
        zoom.set(qn("w:percent"), "100")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ================================================================ UI
st.set_page_config(page_title="Analisis MBKM", page_icon="🎓", layout="wide")
st.title("Analisis Aktivitas MBKM & Mata Kuliah Konversi i3L University")
# st.caption("Unggah file ekspor SIAKAD (.xls / .xlsx). Data dinormalkan ke 1 baris per aktivitas, "
#            "lalu diberi flag kualitas data.")
st.caption("Supported by Claude Opus 4.8 (Anthropic PBC, San Francisco, California, U.S.) ")

up = st.file_uploader("File ekspor SIAKAD", type=["xls", "xlsx", "html", "htm"])
if up is None:
    st.info("Silakan unggah file laporan untuk memulai.")
    st.stop()

try:
    raw, act_all, source_note, n_split = load(up.getvalue(), up.name)
except Exception as exc:  # noqa
    st.error("Gagal membaca file: %s" % exc)
    st.stop()

with st.sidebar:
    st.header("Filter")
    all_status = sorted(act_all["Status Aktivitas"].unique())
    default_drop = []
    for s in ["Ditolak", "Dibatalkan"]:
        if s in all_status:
            default_drop.append(s)
    drop_status = st.multiselect("Buang status aktivitas", all_status, default=default_drop)
    mk_overload = st.number_input("Ambang MK berlebih (F2)", min_value=1, max_value=50, value=5)
    mk_overload_px = st.number_input("Ambang MK berlebih - Pertukaran Pelajar (F2)",
                                     min_value=1, max_value=60, value=12)
    lang_pick = st.radio("Bahasa laporan naratif", list(LANGS.keys()), horizontal=True)
    lang = LANGS[lang_pick]
    st.caption("File asli: %d baris (1 baris per aktivitas × MK konversi)." % len(raw))
    if n_split:
        st.caption("%d record terpecah (urutan nama dosen berbeda) sudah digabung." % n_split)
    if source_note:
        st.caption(source_note)

act = act_all[~act_all["Status Aktivitas"].isin(drop_status)].reset_index(drop=True)
act["No"] = range(1, len(act) + 1)
if len(act) == 0:
    st.warning("Tidak ada aktivitas tersisa setelah filter status.")
    st.stop()
flags = flag_table(act, mk_overload, mk_overload_px)
defs = flag_defs(mk_overload, mk_overload_px)
nim_info = nim_check(act)
periode = periode_label(act_all)
stem = period_stem(act_all)

n_act = len(act)
n_stu = act["NIM"].nunique()
n_dropped = len(act_all) - len(act)
n_issue = int(pd.concat([flags[k] for k in RECORD_FLAGS], axis=1).any(axis=1).sum())
meta = {"source_note": source_note, "drop_status": drop_status, "n_dropped": n_dropped,
        "mk_overload": mk_overload, "mk_overload_px": mk_overload_px, "n_split": n_split,
        "periode": periode, "act_all": act_all}

cats = partner_categories(act, flags)
narr = narrative_content(act, raw, flags, nim_info, meta, cats, lang)
charts = narrative_charts(act, cats, lang)
narr_name = "%s_%s" % (stem, NARR_TEXT[lang]["file"])
narr_errors = []
try:
    narr_pdf = build_narrative_pdf(narr, charts, meta)
except Exception as exc:  # noqa
    narr_pdf = None
    narr_errors.append("PDF: %s" % exc)
try:
    narr_docx = build_narrative_docx(narr, charts, meta)
except Exception as exc:  # noqa
    narr_docx = None
    narr_errors.append("Word (butuh paket python-docx): %s" % exc)
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Aktivitas", n_act)
# m2.metric("Mahasiswa", n_stu)
# m3.metric("Baris asli", len(raw))
# m4.metric("Dikecualikan", n_dropped)
# m5.metric("Isu per-record", n_issue)

PENDING_TAB = {"id": "Mahasiswa Status Selain Selesai", "en": "Students with a Status Other than Completed"}
tab_sum, tab_jenis, tab_prodi, tab_flag, tab_nim, tab_pending, tab_matrix, tab_narr, tab_dl = st.tabs(
    ["Ringkasan", "Per Jenis", "Per Program Studi", "Flag", "Cek NIM", PENDING_TAB[lang], "Matriks", "Narasi",
     "Unduh"])

with tab_sum:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Status Aktivitas")
        st.pyplot(barh(act["Status Aktivitas"].value_counts(), NAVY, "Status aktivitas"))
    with c2:
        st.subheader("Flag")
        fc = {}
        for k in FLAG_COLS:
            fc[k] = int(flags[k].sum())
        st.pyplot(barh(pd.Series(fc), RED, "Jumlah aktivitas ter-flag"))

with tab_jenis:
    st.subheader("Rekap mahasiswa per Jenis Aktivitas")
    stu_j = act.groupby("Jenis Aktivitas")["NIM"].nunique().sort_values(ascending=False)
    labels = []
    for j in stu_j.index:
        labels.append(short_jenis(j))
    stu_j.index = labels
    st.pyplot(barh(stu_j, NAVY, "Mahasiswa per Jenis Aktivitas"))
    jenis_pick = st.selectbox("Lihat daftar mahasiswa untuk Jenis Aktivitas",
                              act["Jenis Aktivitas"].value_counts().index)
    sub = act[act["Jenis Aktivitas"] == jenis_pick]
    st.caption("%d aktivitas, %d mahasiswa" % (len(sub), sub["NIM"].nunique()))
    st.dataframe(sub[["NIM", "Nama", "Program Studi", "Status Aktivitas", "Mitra", "Judul Aktivitas"]]
                 .reset_index(drop=True), use_container_width=True)

with tab_prodi:
    st.subheader("Rekap mahasiswa per Program Studi")
    stu_p = act.groupby("Program Studi")["NIM"].nunique().sort_values(ascending=False)

    # --- tabel rekap + unduh ---
    rekap_p = rekap_table(act, "Program Studi")
    labels = []
    for p in rekap_p["Program Studi"]:
        labels.append(prodi_label(p))
    rekap_p["Program Studi"] = labels

    stu_chart = stu_p.copy()
    labels = []
    for p in stu_chart.index:
        labels.append(prodi_label(p))
    stu_chart.index = labels
    st.pyplot(barh(stu_chart, NAVY, "Mahasiswa per Program Studi"))

    st.dataframe(rekap_p, use_container_width=True, hide_index=True)
    st.download_button("⬇️ Unduh rekap per Program Studi (CSV)",
                       rekap_p.to_csv(index=False).encode("utf-8"),
                       file_name="rekap_mahasiswa_per_prodi.csv", mime="text/csv",
                       key="dl_rekap_prodi")

    st.divider()

    # --- daftar mahasiswa per prodi (drill-down) + unduh ---
    prodi_pick = st.selectbox("Lihat daftar mahasiswa untuk Program Studi", stu_p.index)
    sub = act[act["Program Studi"] == prodi_pick].reset_index(drop=True)
    sub_show = sub[["NIM", "Nama", "Jenis Aktivitas", "Status Aktivitas", "Mitra", "Judul Aktivitas"]]
    st.caption("%d aktivitas, %d mahasiswa" % (len(sub), sub["NIM"].nunique()))
    st.dataframe(sub_show, use_container_width=True, hide_index=True)

    slug = prodi_label(prodi_pick).replace(" ", "_")
    st.download_button("⬇️ Unduh daftar mahasiswa (CSV)",
                       sub_show.to_csv(index=False).encode("utf-8"),
                       file_name="mahasiswa_%s.csv" % slug, mime="text/csv",
                       key="dl_students_prodi")

with tab_flag:
    st.subheader("Flag kualitas data")
    cols = st.columns(len(FLAG_COLS))
    for i, k in enumerate(FLAG_COLS):
        cols[i].metric(k, int(flags[k].sum()), help=defs[k])
    st.divider()
    flag_pick = st.selectbox("Lihat record untuk flag", FLAG_COLS)
    st.caption(defs[flag_pick])
    sub = act[flags[flag_pick]]
    show = ["No", "NIM", "Nama", "Program Studi", "Jenis Aktivitas", "Status Aktivitas",
            "Mitra", "Jml MK", "Total SKS", "Dosen Pembimbing", "Judul Aktivitas"]
    st.dataframe(sub[show].reset_index(drop=True), use_container_width=True)

with tab_nim:
    st.subheader("Pemeriksaan NIM")
    multi = nim_info["multi"]
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Aktivitas", n_act)
    c2.metric("Mahasiswa (NIM unik)", n_stu)
    c3.metric("Selisih", n_act - n_stu)
    c4.metric("NIM >1 aktivitas", len(multi))
    if n_split:
        st.caption("%d record aktivitas yang terpecah karena urutan nama dosen berbeda di ekspor SIAKAD "
                   "sudah digabung otomatis." % n_split)
    if len(multi) == 0:
        st.success("Setiap NIM tercatat tepat 1 aktivitas — jumlah aktivitas = jumlah mahasiswa.")
    else:
        st.warning("%d NIM tercatat di lebih dari 1 aktivitas, sehingga jumlah aktivitas (%d) > jumlah "
                   "mahasiswa (%d). 'MK Sama' = MK yang dikonversi di lebih dari 1 aktivitas "
                   "(indikasi kuat duplikasi / pengajuan ulang)." % (len(multi), n_act, n_stu))
        st.dataframe(multi, use_container_width=True, hide_index=True)
    if len(nim_info["invalid"]):
        st.error("%d aktivitas dengan NIM kosong / bukan angka." % len(nim_info["invalid"]))
        st.dataframe(nim_info["invalid"][["No", "NIM", "Nama", "Program Studi"]],
                     use_container_width=True, hide_index=True)
    if len(nim_info["conflict"]):
        st.error("%d NIM dengan nama / program studi berbeda." % len(nim_info["conflict"]))
        st.dataframe(nim_info["conflict"], use_container_width=True, hide_index=True)

    st.divider()
    nim_query = st.text_input("Cek satu NIM (semua status, termasuk yang dibuang filter)")
    if nim_query.strip():
        q = clean_nim(nim_query)
        hits = act_all[act_all["NIM"] == q]
        if len(hits) == 0:
            st.info("NIM %s tidak ditemukan di file." % q)
        else:
            st.caption("%s - %s: %d aktivitas, %d baris MK di file asli"
                       % (q, hits.iloc[0]["Nama"], len(hits), int(hits["_rows"].sum())))
            st.dataframe(hits[["Jenis Aktivitas", "Status Aktivitas", "Mitra", "Jml MK", "Total SKS",
                               "MK Konversi", "Dosen Pembimbing", "Dosen Penguji"]],
                         use_container_width=True, hide_index=True)

with tab_pending:
    T = NARR_TEXT[lang]
    en = lang == "en"
    pending_df = narr["pending_df"]
    excluded_df = narr["excluded_df"]
    st.subheader(PENDING_TAB[lang])
    raw_pending = act[act["Status Aktivitas"] != "Selesai"]
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("Activities" if en else "Aktivitas", len(raw_pending))
    p2.metric("Students" if en else "Mahasiswa", raw_pending["NIM"].nunique())
    p3.metric(status_name("Evaluasi", lang), int((raw_pending["Status Aktivitas"] == "Evaluasi").sum()))
    p4.metric(status_name("Diajukan", lang), int((raw_pending["Status Aktivitas"] == "Diajukan").sum()))
    for p in narr["pending"]:
        st.markdown(md_escape(p))
    st.divider()
    if len(pending_df):
        st.markdown("**" + T["pending_t1"] + "**")
        prodi_col = pending_df.columns[2]
        status_col = pending_df.columns[4]
        f1, f2 = st.columns(2)
        prodi_sel = f1.multiselect(prodi_col, sorted(pending_df[prodi_col].unique()), key="pend_prodi")
        status_sel = f2.multiselect(status_col, list(pending_df[status_col].unique()), key="pend_status")
        shown = pending_df
        if prodi_sel:
            shown = shown[shown[prodi_col].isin(prodi_sel)]
        if status_sel:
            shown = shown[shown[status_col].isin(status_sel)]
        st.dataframe(shown, use_container_width=True, hide_index=True)
        st.download_button("⬇️ " + ("Download list (CSV)" if en else "Unduh daftar (CSV)"),
                           shown.to_csv(index=False).encode("utf-8"),
                           file_name="%s_%s.csv" % (stem, "not_completed" if en else "belum_selesai"),
                           mime="text/csv", key="dl_pending")
    else:
        st.success("All analysed activities have been completed." if en
                   else "Semua aktivitas yang dianalisis telah berstatus Selesai.")
    if len(excluded_df):
        st.divider()
        st.markdown("**" + T["pending_t2"] + "**")
        st.dataframe(excluded_df, use_container_width=True, hide_index=True)
        st.download_button("⬇️ " + ("Download excluded records (CSV)" if en else "Unduh record dikecualikan (CSV)"),
                           excluded_df.to_csv(index=False).encode("utf-8"),
                           file_name="%s_%s.csv" % (stem, "excluded" if en else "dikecualikan"),
                           mime="text/csv", key="dl_excluded")

with tab_matrix:
    st.subheader("Matriks mahasiswa: Program Studi × Jenis Aktivitas")
    st.dataframe(student_matrix(act), use_container_width=True)
    st.caption("Angka = jumlah mahasiswa (NIM unik). Mahasiswa dengan aktivitas di lebih dari 1 jenis "
               "dihitung di tiap jenis, tetapi sekali di kolom/baris Total.")

with tab_narr:
    T = NARR_TEXT[lang]
    st.subheader("Laporan naratif (%s)" % lang_pick)
    st.caption("Narasi otomatis per program studi dan per jenis aktivitas, lengkap dengan grafik. Bahasa dipilih "
               "di sidebar; isi PDF dan Word sama dengan tampilan di bawah ini.")
    d1, d2 = st.columns(2)
    if narr_pdf:
        d1.download_button("📝 Unduh laporan naratif (PDF)", narr_pdf, file_name=narr_name + ".pdf",
                           mime="application/pdf", key="dl_narr_pdf_tab")
    if narr_docx:
        d2.download_button("📄 Unduh laporan naratif (Word)", narr_docx, file_name=narr_name + ".docx",
                           mime=DOCX_MIME, key="dl_narr_docx_tab")
    for err in narr_errors:
        st.warning("Laporan naratif gagal dibuat - %s" % err)
    kpi_cols = st.columns(len(narr["kpi"]))
    for i, (label, value) in enumerate(narr["kpi"]):
        kpi_cols[i].metric(label, value)
    st.markdown("#### " + T["s1"])
    for p in narr["summary"]:
        st.markdown(md_escape(p))
    st.image(charts["prodi_jenis"], caption=FIG_CAPTIONS[lang]["prodi_jenis"])
    st.markdown("#### " + T["s2"])
    st.markdown(md_escape(narr["prodi_intro"]))
    st.image(charts["prodi_mitra"], caption=FIG_CAPTIONS[lang]["prodi_mitra"])
    for i, (name, paras) in enumerate(narr["prodi"], start=1):
        st.markdown("**2.%d %s**" % (i, md_escape(name)))
        for p in paras:
            st.markdown(md_escape(p))
    st.markdown("#### " + T["s3"])
    st.markdown(md_escape(narr["jenis_intro"]))
    st.image(charts["jenis_status"], caption=FIG_CAPTIONS[lang]["jenis_status"])
    st.image(charts["jenis_mitra"], caption=FIG_CAPTIONS[lang]["jenis_mitra"])
    for i, (name, paras) in enumerate(narr["jenis"], start=1):
        st.markdown("**3.%d %s**" % (i, md_escape(name)))
        for p in paras:
            st.markdown(md_escape(p))
    st.markdown("#### " + T["s_pending"])
    for p in narr["pending"]:
        st.markdown(md_escape(p))
    if len(narr["pending_df"]):
        st.caption(T["pending_t1"])
        st.dataframe(narr["pending_df"], use_container_width=True, hide_index=True)
    if len(narr["excluded_df"]):
        st.caption(T["pending_t2"])
        st.dataframe(narr["excluded_df"], use_container_width=True, hide_index=True)
    st.markdown("#### " + T["s4"])
    for p in narr["quality"]:
        st.markdown(md_escape(p))
    if narr["actions"]:
        lines = [T["actions"]]
        for a in narr["actions"]:
            lines.append("- " + md_escape(a))
        st.markdown("\n".join(lines))
    st.markdown("#### " + T["s5"])
    st.markdown(md_escape(narr["method"]))

with tab_dl:
    st.subheader("Unduh hasil")
    try:
        pdf_bytes = build_pdf(act, raw, flags, nim_info, meta)
        st.download_button("📄 Unduh laporan PDF", pdf_bytes, file_name=stem + "_laporan.pdf",
                           mime="application/pdf")
    except Exception as exc:  # noqa
        st.warning("PDF gagal dibuat (butuh paket reportlab): %s" % exc)
    if narr_pdf:
        st.download_button("📝 Unduh laporan naratif - %s (PDF)" % lang_pick, narr_pdf,
                           file_name=narr_name + ".pdf", mime="application/pdf", key="dl_narr_pdf")
    if narr_docx:
        st.download_button("📝 Unduh laporan naratif - %s (Word)" % lang_pick, narr_docx,
                           file_name=narr_name + ".docx", mime=DOCX_MIME, key="dl_narr_docx")
    for err in narr_errors:
        st.warning("Laporan naratif gagal dibuat - %s" % err)
    xlsx_bytes = build_excel(act, flags, nim_info, act_all, drop_status)
    st.download_button("📊 Unduh data Excel (Aktivitas + rekap + cek NIM)", xlsx_bytes,
                       file_name=stem + "_bersih.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("📥 Unduh Aktivitas (CSV)", build_csv(act, flags),
                       file_name=stem + "_aktivitas.csv", mime="text/csv")
