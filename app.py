"""
Analisis Aktivitas MBKM & Mata Kuliah Konversi - aplikasi Streamlit.

Jalankan:
    pip install streamlit pandas beautifulsoup4 lxml openpyxl xlrd reportlab matplotlib python-docx
    streamlit run app.py

Unggah file ekspor SIAKAD (.xls yang sebenarnya HTML, atau .xlsx). Aplikasi
menormalkan data (1 baris per aktivitas), memberi flag kualitas data, memeriksa
NIM, menyediakan rekap per Jenis Aktivitas / Program Studi, laporan naratif
bergrafik (Bahasa Indonesia / English; PDF & Word), serta unduhan PDF, Excel & CSV.
Unggah file Semester Ganjil, Semester Genap, atau keduanya: dengan dua file, laporan
mencakup satu tahun akademik (per semester + total) beserta perbandingan semester.

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
# Semester: warna (satu hue, urut waktu: gelap = semester pertama), urutan istilah, dan flag lintas semester.
SEM_COLORS = {2: ["#1f3864", "#8fa4c8"], 3: ["#1f3864", "#57709f", "#8fa4c8"]}
TERM_ORDER = {"ganjil": 0, "genap": 1}
TERM_EN = {"ganjil": "odd", "genap": "even", "pendek": "short", "antara": "short"}
F8 = "F8 MK Ganda Lintas Semester"
CROSS_COLS = {"id": ["NIM", "Nama", "Program Studi"], "en": ["NIM", "Name", "Study program"]}
SAME_MK_COL = {"id": "MK sama", "en": "Same course"}
PT_PER_CM = 28.3465
DUP_SUFFIX = " (duplikat)"
DUP_BASE = ["Diajukan", "Evaluasi"]


def mark_superseded(act):
    """Pengajuan ganda: record Diajukan/Evaluasi yang mengonversi MK yang sama dengan aktivitas Selesai milik NIM
    yang sama pada semester yang sama diberi status '<status> (duplikat)', sehingga dapat dibuang seperti Ditolak."""
    out = act.reset_index(drop=True).copy()
    statuses = list(out["Status Aktivitas"])
    for i in range(len(out)):
        if statuses[i] not in DUP_BASE:
            continue
        codes = set(str(out.at[i, "_codes"]).split(";")) - {""}
        if not codes:
            continue
        for j in range(len(out)):
            if j == i or statuses[j] != "Selesai":
                continue
            if out.at[j, "NIM"] != out.at[i, "NIM"] or out.at[j, "Periode Akademik"] != out.at[i, "Periode Akademik"]:
                continue
            if codes & set(str(out.at[j, "_codes"]).split(";")):
                out.at[i, "Status Aktivitas"] = statuses[i] + DUP_SUFFIX
                break
    return out


def is_dup_status(name):
    return str(name).endswith(DUP_SUFFIX)


IN_PROGRESS = ["Evaluasi", "Diajukan"]


def uncounted_records(act_all, drop_status):
    """Aktivitas yang belum selesai (Evaluasi/Diajukan) tetapi tidak dihitung karena statusnya dibuang filter
    (mis. bila hanya status Selesai yang dihitung). Tetap dilaporkan sebagai aktivitas yang belum selesai."""
    if act_all is None:
        return None
    keep = []
    for s in drop_status or []:
        if s in IN_PROGRESS:
            keep.append(s)
    return act_all[act_all["Status Aktivitas"].isin(keep)]


def unc_part(unc, col=None, value=None):
    """Bagian dari aktivitas tak terhitung untuk satu kelompok (aman bila unc None)."""
    if unc is None or col is None:
        return unc
    return unc[unc[col] == value]


def unc_len(unc):
    return 0 if unc is None else len(unc)


def with_uncounted(act, unc):
    """Aktivitas terhitung + aktivitas belum selesai yang tidak dihitung (untuk tampilan status)."""
    if unc is None or len(unc) == 0:
        return act
    return pd.concat([act, unc], ignore_index=True)


DEFAULT_COHORTS = ["22", "23"]


def parse_cohorts(text):
    """'22, 23' -> ['22', '23'] (awal NIM yang dihitung)."""
    out = []
    for piece in re.split(r"[,;\s]+", str(text)):
        piece = re.sub(r"\D", "", piece)
        if piece and piece not in out:
            out.append(piece)
    return out


def cohort_counts(act, prefixes):
    """Jumlah mahasiswa (NIM unik) per awal NIM; NIM lain dikumpulkan per dua digit awal."""
    counts = {}
    for p in prefixes:
        counts[p] = 0
    others = {}
    for nim in act["NIM"].astype(str).unique():
        matched = False
        for p in prefixes:
            if nim.startswith(p):
                counts[p] += 1
                matched = True
                break
        if not matched:
            key = nim[:2] if nim[:2].isdigit() else "?"
            others[key] = others.get(key, 0) + 1
    return counts, others


def cohort_table(act, prefixes, periods):
    """Tabel mahasiswa per awal NIM: kolom per semester (bila >1) + Total."""
    groups = []
    if len(periods) > 1:
        for p in periods:
            groups.append((p, act[act["Periode Akademik"] == p]))
        groups.append(("Total", act))
    else:
        groups.append(("Mahasiswa", act))
    tallies = []
    other_keys = []
    for label, sub in groups:
        counts, others = cohort_counts(sub, prefixes)
        tallies.append((counts, others))
        for k in others:
            if k not in other_keys:
                other_keys.append(k)
    other_keys.sort()
    rows = []
    for p in prefixes:
        row = ["NIM %s..." % p]
        for counts, others in tallies:
            row.append(counts[p])
        rows.append(row)
    if other_keys:
        row = ["Lainnya (%s)" % ", ".join(other_keys)]
        for counts, others in tallies:
            row.append(sum(others.values()))
        rows.append(row)
    row = ["Total"]
    for label, sub in groups:
        row.append(int(sub["NIM"].nunique()))
    rows.append(row)
    cols = ["Awal NIM"]
    for label, sub in groups:
        cols.append(label)
    return pd.DataFrame(rows, columns=cols)


def any_dup(statuses):
    for s in statuses:
        if is_dup_status(s):
            return True
    return False


def short_jenis(name):
    if name in JENIS_SHORT:
        return JENIS_SHORT[name]
    return str(name).replace("(Kampus Merdeka)", "").strip()


def flag_defs(mk_overload, mk_overload_px, multi_sem=False):
    f7 = "NIM tercatat di lebih dari 1 aktivitas (setelah filter) - cek duplikasi / pengajuan ulang."
    if multi_sem:
        f7 = ("NIM tercatat di lebih dari 1 aktivitas dalam semester yang sama (setelah filter) - cek duplikasi / "
              "pengajuan ulang.")
    return {
        "F1 Tanpa MK": "Aktivitas tanpa satu pun mata kuliah konversi (Jml MK = 0).",
        "F2 MK Berlebih": "Jml MK > %d (Pertukaran Pelajar: > %d) - diduga salah pilih / seluruh katalog terpilih."
                          % (mk_overload, mk_overload_px),
        "F3 Tanpa Pembimbing": "Kolom Dosen Pembimbing kosong atau '-'.",
        "F4 Tanpa Penguji": "Kolom Dosen Penguji kosong atau '-' (seluruh record).",
        "F5 Mitra Tidak Valid": "Mitra kosong/'-' atau berisi kode MK / teks bebas, bukan nama mitra.",
        "F6 Tanpa MoU": "Mitra berstatus 'Belum memiliki MoU kerjasama'.",
        "F7 NIM Ganda": f7,
        F8: ("NIM yang sama mengonversi MK yang sama pada lebih dari satu semester (setelah filter) - cek "
             "konversi ganda."),
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


def cross_mk_flags(act):
    """True bila NIM yang sama mengonversi MK yang sama pada periode akademik (semester) lain."""
    periods_of = {}
    for i in act.index:
        for code in act.at[i, "_codes"].split(";"):
            if code:
                key = (act.at[i, "NIM"], code)
                if key not in periods_of:
                    periods_of[key] = set()
                periods_of[key].add(act.at[i, "Periode Akademik"])
    hits = []
    for i in act.index:
        hit = False
        for code in act.at[i, "_codes"].split(";"):
            if code and len(periods_of[(act.at[i, "NIM"], code)]) > 1:
                hit = True
        hits.append(hit)
    return pd.Series(hits, index=act.index, dtype=bool)


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
    # F7: duplikasi dalam semester yang sama; F8 (hanya bila data >1 semester): MK sama di semester lain
    out["F7 NIM Ganda"] = act.duplicated(["NIM", "Periode Akademik"], keep=False)
    if act["Periode Akademik"].nunique() > 1:
        out[F8] = cross_mk_flags(act)
    return out


def prodi_label(name):
    return name.replace("S1 - ", "").strip()


def nim_check(act):
    """NIM dengan >1 aktivitas (dalam semester yang sama), NIM di >1 semester, NIM kosong/bukan angka,
    dan NIM dengan nama/prodi berbeda."""
    multi_sem = act["Periode Akademik"].nunique() > 1
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
        parts = [(None, group)]
        if multi_sem:
            parts = []
            for period in sorted_periods(group["Periode Akademik"].unique()):
                parts.append((period, group[group["Periode Akademik"] == period]))
        for period, part in parts:
            if len(part) < 2:
                continue
            items = []
            code_count = {}
            for _, row in part.iterrows():
                items.append("%s - %s" % (short_jenis(row["Jenis Aktivitas"]), row["Status Aktivitas"]))
                for code in row["_codes"].split(";"):
                    if code:
                        code_count[code] = code_count.get(code, 0) + 1
            same = []
            for code in sorted(code_count):
                if code_count[code] > 1:
                    same.append(code)
            record = {"NIM": nim, "Nama": part.iloc[0]["Nama"],
                      "Program Studi": prodi_label(part.iloc[0]["Program Studi"]),
                      "Jml Aktivitas": len(part), "Aktivitas (Jenis - Status)": "; ".join(items),
                      "MK Sama": ", ".join(same)}
            if multi_sem:
                record["Semester"] = period
            multi_rows.append(record)
    cols = list(NIM_COLS)
    if multi_sem:
        cols = NIM_COLS[:3] + ["Semester"] + NIM_COLS[3:]
    return {
        "multi": pd.DataFrame(multi_rows, columns=cols),
        "cross": cross_semester_table(act, "id"),
        "conflict": pd.DataFrame(conflict_rows, columns=["NIM", "Nama", "Program Studi"]),
        "invalid": act[~act["NIM"].str.fullmatch(r"\d+")],
    }


def cross_semester_table(act, lang="id"):
    """Mahasiswa (NIM) yang tercatat pada lebih dari satu semester: aktivitas per semester dan MK yang sama."""
    periods = sorted_periods(act["Periode Akademik"].unique())
    ay = academic_year(periods)
    cols = list(CROSS_COLS[lang])
    for p in periods:
        cols.append(sem_label(p, lang, ay, "col"))
    cols.append(SAME_MK_COL[lang])
    rows = []
    if len(periods) < 2:
        return pd.DataFrame(rows, columns=cols)
    for nim, group in act.groupby("NIM", sort=True):
        if group["Periode Akademik"].nunique() < 2:
            continue
        row = [nim, group.iloc[0]["Nama"], prodi_label(group.iloc[0]["Program Studi"])]
        code_periods = {}
        for p in periods:
            items = []
            for _, r in group[group["Periode Akademik"] == p].iterrows():
                codes = []
                for code in r["_codes"].split(";"):
                    if code:
                        codes.append(code)
                        if code not in code_periods:
                            code_periods[code] = set()
                        code_periods[code].add(p)
                text = "%s - %s" % (jenis_short(r["Jenis Aktivitas"], lang), status_name(r["Status Aktivitas"], lang))
                if codes:
                    text += " (%s)" % ", ".join(codes)
                items.append(text)
            row.append("; ".join(items) if items else "-")
        same = []
        for code in sorted(code_periods):
            if len(code_periods[code]) > 1:
                same.append(code)
        row.append(", ".join(same) if same else "-")
        rows.append(row)
    return pd.DataFrame(rows, columns=cols)


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


def rekap_sem(act, by, periods):
    """Rekap per kategori untuk setiap semester + Total (mahasiswa = NIM unik)."""
    order = list(rekap(act, by).index)
    rows = []
    for cat in order + [None]:
        if cat is None:
            sub = act
        else:
            sub = act[act[by] == cat]
        row = {by: "TOTAL" if cat is None else cat}
        for p in periods:
            part = sub[sub["Periode Akademik"] == p]
            row["%s · Aktivitas" % p] = len(part)
            row["%s · Mahasiswa" % p] = int(part["NIM"].nunique())
        row["Total · Aktivitas"] = len(sub)
        row["Total · Mahasiswa"] = int(sub["NIM"].nunique())
        rows.append(row)
    return pd.DataFrame(rows)


def rekap_sem_display(act, by, periods, label):
    """rekap_sem dengan label kategori yang ringkas (baris TOTAL tetap)."""
    table = rekap_sem(act, by, periods)
    names = []
    for value in table[by]:
        if value == "TOTAL":
            names.append(value)
        else:
            names.append(label(value))
    table[by] = names
    return table


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


def period_parts(period):
    """'2025 Ganjil' -> (2025, 'Ganjil'); format lain -> (None, teks)."""
    found = re.match(r"\s*(\d{4})\s+(\S.*?)\s*$", str(period))
    if found:
        return int(found.group(1)), found.group(2)
    return None, squash(period)


def period_sort_key(period):
    year, term = period_parts(period)
    if year is None:
        return (9999, 9, term.lower())
    return (year, TERM_ORDER.get(term.lower(), 2), term.lower())


def sorted_periods(values):
    periods = []
    for value in values:
        value = squash(value)
        if value and value not in periods:
            periods.append(value)
    periods.sort(key=period_sort_key)
    return periods


def academic_year(periods):
    """Tahun akademik bersama, mis. ['2025 Ganjil', '2025 Genap'] -> '2025/2026'; bila berbeda -> ''."""
    years = set()
    for p in periods:
        year, term = period_parts(p)
        if year is None:
            return ""
        years.add(year)
    if len(years) != 1:
        return ""
    year = years.pop()
    return "%d/%d" % (year, year + 1)


def periode_text(periods):
    if not periods:
        return "-"
    if len(periods) == 1:
        return periods[0]
    text = join(periods, "id")
    ay = academic_year(periods)
    if ay:
        return "%s (Tahun Akademik %s)" % (text, ay)
    return text


def periode_label(act_all):
    return periode_text(sorted_periods(act_all["Periode Akademik"].unique()))


def period_stem(act_all):
    periods = sorted_periods(act_all["Periode Akademik"].unique())
    if not periods:
        return "MBKM"
    ay = academic_year(periods)
    if len(periods) > 1 and ay:
        return "MBKM_TA_" + ay.replace("/", "-")
    if len(periods) > 2:
        return "MBKM_multi_periode"
    parts = []
    for p in periods:
        parts.append(p.replace(" ", "_"))
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
    for col in flags.columns:
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
    Bila data mencakup lebih dari satu semester, kolom Semester ditambahkan di depan.
    """
    values = list(act["Periode Akademik"].unique())
    if act_all is not None:
        values += list(act_all["Periode Akademik"].unique())
    periods = sorted_periods(values)
    with_period = len(periods) > 1
    ay = academic_year(periods)
    cols = list(NOT_DONE_COLS[lang])
    if with_period:
        cols = ["Semester"] + cols

    def order_key(r):
        first = (0,)
        if with_period:
            first = period_sort_key(r["Periode Akademik"])
        return (first, status_rank(r["Status Aktivitas"]), r["Program Studi"], r["Nama"])

    def lead_cells(r):
        cells = [r["NIM"], r["Nama"], prodi_label(r["Program Studi"]), jenis_short(r["Jenis Aktivitas"], lang),
                 status_name(r["Status Aktivitas"], lang), r["Mitra"]]
        if with_period:
            cells = [sem_label(r["Periode Akademik"], lang, ay, "cell")] + cells
        return cells

    unc = uncounted_records(act_all, drop_status)
    pending_rows = []
    for _, r in act[act["Status Aktivitas"] != "Selesai"].iterrows():
        pending_rows.append((order_key(r), lead_cells(r)))
    if unc is not None:
        for _, r in unc.iterrows():
            pending_rows.append((order_key(r), lead_cells(r)))
    pending_rows.sort(key=lambda x: x[0])
    pending = []
    for key, cells in pending_rows:
        pending.append(cells)
    still = with_uncounted(act, unc)
    excluded_rows = []
    if act_all is not None and drop_status:
        gone = act_all[act_all["Status Aktivitas"].isin(drop_status) & ~act_all["Status Aktivitas"].isin(IN_PROGRESS)]
        for _, r in gone.iterrows():
            found = []
            for _, o in still[still["NIM"] == r["NIM"]].iterrows():
                text = "%s - %s" % (jenis_short(o["Jenis Aktivitas"], lang), status_name(o["Status Aktivitas"], lang))
                if with_period:
                    text += " (%s)" % sem_label(o["Periode Akademik"], lang, ay, "cell")
                found.append((period_sort_key(o["Periode Akademik"]), text))
            found.sort(key=lambda x: x[0])
            others = []
            for key, text in found:
                others.append(text)
            excluded_rows.append((order_key(r), lead_cells(r) + ["; ".join(others) if others else "-"]))
    excluded_rows.sort(key=lambda x: x[0])
    excluded = []
    for key, cells in excluded_rows:
        excluded.append(cells)
    return pd.DataFrame(pending, columns=cols[:-1]), pd.DataFrame(excluded, columns=cols)


def build_excel(act, flags, nim_info, act_all=None, drop_status=None):
    buf = io.BytesIO()
    periods = sorted_periods(act["Periode Akademik"].unique())
    multi_sem = len(periods) > 1
    cols = ["No", "NIM", "Nama", "Program Studi", "Jenis Aktivitas", "Mitra",
            "Status Mitra", "Status Aktivitas", "Tanggal Mulai", "Tanggal Selesai",
            "Jml MK", "Total SKS", "MK Konversi", "Judul Aktivitas",
            "Dosen Pembimbing", "Dosen Penguji"]
    if multi_sem:
        cols.insert(4, "Periode Akademik")
    table = export_table(act, flags, cols)
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        table.to_excel(xl, sheet_name="Aktivitas", index=False)
        if multi_sem:
            cats = partner_categories(act, flags)
            semester_compare_table(act, cats, flags, "id", uncounted_records(act_all, drop_status)).to_excel(xl, sheet_name="Per Semester", index=False)
            rekap_sem(act, "Jenis Aktivitas", periods).to_excel(xl, sheet_name="Per Jenis", index=False)
            rekap_sem(act, "Program Studi", periods).to_excel(xl, sheet_name="Per Prodi", index=False)
        else:
            rekap_table(act, "Jenis Aktivitas").to_excel(xl, sheet_name="Per Jenis", index=False)
            rekap_table(act, "Program Studi").to_excel(xl, sheet_name="Per Prodi", index=False)
        nim_info["multi"].to_excel(xl, sheet_name="Cek NIM", index=False)
        if multi_sem and len(nim_info["cross"]):
            nim_info["cross"].to_excel(xl, sheet_name="Lintas Semester", index=False)
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
    periods = sorted_periods(act["Periode Akademik"].unique())
    multi_sem = len(periods) > 1
    defs = flag_defs(meta["mk_overload"], meta["mk_overload_px"], multi_sem)
    flag_names = list(flags.columns)

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

    HEADC = ParagraphStyle("HEADC", parent=CELL, fontName="Helvetica-Bold", fontSize=8, leading=9.5,
                           textColor=colors.white, alignment=1)

    def tstyle(header_bg=navy, total_row=False, numeric_cols=None, body_size=8.5, pad=6):
        cmds = [("BACKGROUND", (0, 0), (-1, 0), header_bg),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (-1, -1), "Helvetica"),
                ("FONTSIZE", (0, 0), (-1, -1), body_size),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
                ("LEFTPADDING", (0, 0), (-1, -1), pad), ("RIGHTPADDING", (0, 0), (-1, -1), pad),
                ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#BFBFBF")),
                ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, light])]
        for c in (numeric_cols or []):
            cmds.append(("ALIGN", (c, 1), (c, -1), "CENTER"))
        if total_row:
            cmds.append(("BACKGROUND", (0, -1), (-1, -1), grey))
            cmds.append(("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold"))
        return TableStyle(cmds)

    def mm_list(values):
        out = []
        for v in values:
            out.append(v * mm)
        return out

    def value_cols(n):
        cols = []
        for c in range(1, n + 1):
            cols.append(c)
        return cols

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
    unc = uncounted_records(meta.get("act_all"), meta["drop_status"])
    reg = with_uncounted(act, unc)
    status_ct = reg["Status Aktivitas"].value_counts()
    counted_status = set(act["Status Aktivitas"])
    unc_note = []
    if unc_len(unc):
        unc_note = [Paragraph("(tidak dihitung) = aktivitas yang belum selesai; dilaporkan di bagian 7 tetapi tidak "
                              "termasuk angka laporan.", SMALL)]

    def status_cell(s):
        if unc_len(unc) and s not in counted_status:
            return s + " (tidak dihitung)"
        return s

    story.append(Paragraph("Laporan Analisis Aktivitas MBKM &amp; Mata Kuliah Konversi", H1))
    story.append(Paragraph("Indonesia International Institute for Life Sciences (i3L)", SUB))
    story.append(Paragraph("Periode Akademik: " + esc(meta["periode"]), SUB))
    if meta["source_note"]:
        for line in meta["source_note"].split("\n"):
            story.append(Paragraph(esc(line).replace("|", "&nbsp;|&nbsp;"), SUB))
    story.append(Spacer(1, 3))
    story.append(HRFlowable(width="100%", thickness=1.2, color=navy, spaceAfter=2))

    if multi_sem:
        act_all = meta.get("act_all")
        cross = nim_info["cross"]
        sum_sem = 0
        per = {"act": [], "stu": [], "dup": [], "raw": [], "drop": [], "issue": []}
        for p in periods:
            sub = act[act["Periode Akademik"] == p]
            sum_sem += sub["NIM"].nunique()
            per["act"].append(str(len(sub)))
            per["stu"].append(str(sub["NIM"].nunique()))
            per["dup"].append(str(int((multi["Semester"] == p).sum())))
            per["raw"].append(str(int((raw["Periode Akademik"] == p).sum())))
            dropped = 0
            if act_all is not None:
                dropped = int(((act_all["Periode Akademik"] == p)
                               & act_all["Status Aktivitas"].isin(meta["drop_status"])).sum())
            per["drop"].append(str(dropped))
            per["issue"].append(str(int(flags.loc[sub.index, RECORD_FLAGS].any(axis=1).sum())))
        dashes = []
        for p in periods:
            dashes.append("-")

        # 1. Ringkasan per semester
        rows = [["Metrik"] + periods + ["Total"]]
        for label, values, total in (
                ("Jumlah aktivitas (setelah filter)", per["act"], str(n_act)),
                ("Jumlah mahasiswa (NIM unik)", per["stu"], str(n_stu))) + cohort_pdf_rows(act, meta, periods) + (
                ("NIM dengan >1 aktivitas dalam semester yang sama", per["dup"], str(len(multi))),
                ("Mahasiswa di lebih dari satu semester", dashes, str(len(cross))),
                ("Baris pada file ekspor asli", per["raw"], str(len(raw))),
                ("Record dikecualikan (%s)" % (", ".join(meta["drop_status"]) or "-"), per["drop"],
                 str(meta["n_dropped"])),
                ("Aktivitas dengan isu data per-record", per["issue"], str(n_issue))):
            rows.append([Paragraph(esc(label), CELL)] + values + [total])
        value_w = 90.0 / (len(periods) + 1)
        widths = [75]
        for c in range(len(periods) + 1):
            widths.append(value_w)
        t = Table(rows, colWidths=mm_list(widths))
        style = tstyle(numeric_cols=value_cols(len(periods) + 1))
        style.add("ALIGN", (1, 0), (-1, 0), "CENTER")
        style.add("FONTNAME", (-1, 1), (-1, -1), "Helvetica-Bold")
        t.setStyle(style)
        if len(cross):
            note = ("Kolom Total menghitung mahasiswa sebagai NIM unik di seluruh semester: %d mahasiswa tercatat di "
                    "lebih dari satu semester, sehingga Total (%d) lebih kecil daripada penjumlahan per semester (%d)."
                    % (len(cross), n_stu, sum_sem))
        else:
            note = "Tidak ada mahasiswa yang tercatat di lebih dari satu semester."
        if len(multi):
            note += (" %d NIM tercatat di lebih dari 1 aktivitas dalam semester yang sama (lihat bagian 6)."
                     % multi["NIM"].nunique())
        section("1. Ringkasan Umum", [t, Paragraph(note, SMALL)])

        # 2 & 3. Rekap per jenis / per prodi, per semester
        def semester_rekap(by, header, label):
            df = rekap_sem(act, by, periods)
            groups = list(periods) + ["Total"]
            head1 = [header]
            head2 = [""]
            for g in groups:
                head1.append(g)
                head1.append("")
                head2.append("Akt.")
                head2.append("Mhs.")
            rows = [head1, head2]
            for _, r in df.iterrows():
                is_total = r[by] == "TOTAL"
                name = "TOTAL" if is_total else label(r[by])
                row = [Paragraph(esc(name), CELLB if is_total else CELL)]
                for p in periods:
                    row.append(str(int(r["%s · Aktivitas" % p])))
                    row.append(str(int(r["%s · Mahasiswa" % p])))
                row.append(str(int(r["Total · Aktivitas"])))
                row.append(str(int(r["Total · Mahasiswa"])))
                rows.append(row)
            label_w = 57
            each = (165.0 - label_w) / (2 * len(groups))
            widths = [label_w]
            for c in range(2 * len(groups)):
                widths.append(each)
            t = Table(rows, colWidths=mm_list(widths), repeatRows=2)
            cmds = [("BACKGROUND", (0, 0), (-1, 1), navy), ("TEXTCOLOR", (0, 0), (-1, 1), colors.white),
                    ("FONTNAME", (0, 0), (-1, 1), "Helvetica-Bold"), ("FONTNAME", (0, 2), (-1, -1), "Helvetica"),
                    ("FONTSIZE", (0, 0), (-1, -1), 8.2), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                    ("ALIGN", (1, 0), (-1, -1), "CENTER"),
                    ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
                    ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
                    ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#BFBFBF")),
                    ("ROWBACKGROUNDS", (0, 2), (-1, -1), [colors.white, light]),
                    ("SPAN", (0, 0), (0, 1)),
                    ("FONTNAME", (-2, 2), (-1, -1), "Helvetica-Bold"),
                    ("BACKGROUND", (0, -1), (-1, -1), grey), ("FONTNAME", (0, -1), (-1, -1), "Helvetica-Bold")]
            for gi in range(len(groups)):
                cmds.append(("SPAN", (1 + 2 * gi, 0), (2 + 2 * gi, 0)))
            t.setStyle(TableStyle(cmds))
            return t

        legend = Paragraph("Akt. = jumlah aktivitas; Mhs. = jumlah mahasiswa (NIM unik). Kolom Total menghitung NIM "
                           "unik di seluruh semester.", SMALL)
        section("2. Rekap per Jenis Aktivitas", [semester_rekap("Jenis Aktivitas", "Jenis Aktivitas", short_jenis),
                                                 legend])
        section("3. Rekap per Program Studi", [semester_rekap("Program Studi", "Program Studi", prodi_label), legend])

        # 4. Status per semester
        rows = [["Status"] + periods + ["Total"]]
        for s in status_ct.index:
            row = [status_cell(s)]
            for p in periods:
                row.append(str(int(((reg["Periode Akademik"] == p) & (reg["Status Aktivitas"] == s)).sum())))
            row.append(str(int(status_ct[s])))
            rows.append(row)
        t = Table(rows, colWidths=mm_list(widths_status(len(periods))))
        style = tstyle(numeric_cols=value_cols(len(periods) + 1))
        style.add("ALIGN", (1, 0), (-1, 0), "CENTER")
        style.add("FONTNAME", (-1, 1), (-1, -1), "Helvetica-Bold")
        t.setStyle(style)
        section("4. Status Aktivitas", [t] + unc_note)

        # 5. Flag per semester (header semester dibungkus agar muat di kolom sempit)
        head = ["Flag"]
        for p in periods:
            head.append(Paragraph(esc(p), HEADC))
        rows = [head + ["Total", "Definisi"]]
        for k in flag_names:
            row = [Paragraph(esc(k), CELLB)]
            for p in periods:
                row.append(str(int(flags.loc[act["Periode Akademik"] == p, k].sum())))
            row.append(str(int(flags[k].sum())))
            row.append(Paragraph(esc(defs[k]), CELL))
            rows.append(row)
        widths = [38]
        for c in range(len(periods) + 1):
            widths.append(14)
        widths.append(165 - 38 - 14 * (len(periods) + 1))
        t = Table(rows, colWidths=mm_list(widths))
        style = tstyle(header_bg=red, numeric_cols=value_cols(len(periods) + 1))
        style.add("ALIGN", (1, 0), (len(periods) + 1, 0), "CENTER")
        t.setStyle(style)
        section("5. Temuan Kualitas Data (Flag)", [
            t, Paragraph("F1, F2, F3, F5 = isu input per-record. F4 &amp; F6 = isu sistemik. "
                         "F7 &amp; F8 = perlu verifikasi (lihat bagian 6).", SMALL)])
    else:
        # 1. Ringkasan
        rows = [["Metrik", "Nilai"],
                ["Jumlah aktivitas (setelah filter)", str(n_act)],
                ["Jumlah mahasiswa (NIM unik)", str(n_stu)]]
        for label, values, total in cohort_pdf_rows(act, meta, periods):
            rows.append([label, total])
        rows += [["NIM dengan >1 aktivitas", multi_text],
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
            rows.append([status_cell(s), str(int(status_ct[s]))])
        t = Table(rows, colWidths=[130 * mm, 35 * mm])
        t.setStyle(tstyle(numeric_cols=[1]))
        section("4. Status Aktivitas", [t] + unc_note)

        # 5. Flag
        rows = [["Flag", "Jml", "Definisi"]]
        for k in flag_names:
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
        if multi_sem:
            items.append(Paragraph("Tidak ada NIM ganda dalam semester yang sama: dalam setiap semester, setiap NIM "
                                   "tercatat tepat 1 aktivitas.", CELL))
        else:
            items.append(Paragraph("Tidak ada NIM ganda: setiap NIM tercatat tepat 1 aktivitas.", CELL))
    else:
        if multi_sem:
            items.append(Paragraph(
                "%d NIM tercatat di lebih dari 1 aktivitas dalam semester yang sama. Kolom 'MK sama' berisi MK yang "
                "dikonversi di lebih dari 1 aktivitas (indikasi kuat duplikasi / pengajuan ulang)."
                % multi["NIM"].nunique(), CELL))
        else:
            items.append(Paragraph(
                "%d NIM tercatat di lebih dari 1 aktivitas, sehingga jumlah aktivitas melebihi jumlah mahasiswa "
                "sebanyak %d. Kolom 'MK sama' berisi MK yang dikonversi di lebih dari 1 aktivitas "
                "(indikasi kuat duplikasi / pengajuan ulang)." % (len(multi), n_act - n_stu), CELL))
        items.append(Spacer(1, 4))
        if multi_sem:
            rows = [["NIM", "Nama", "Program Studi", "Semester", "Aktivitas (Jenis - Status)", "MK sama"]]
            for _, r in multi.iterrows():
                rows.append([r["NIM"], Paragraph(esc(r["Nama"]), CELL), Paragraph(esc(r["Program Studi"]), CELL),
                             Paragraph(esc(r["Semester"]), CELL),
                             Paragraph(esc(r["Aktivitas (Jenis - Status)"]), CELL),
                             Paragraph(esc(r["MK Sama"] or "-"), CELL)])
            t = Table(rows, colWidths=mm_list([18, 34, 30, 20, 45, 18]), repeatRows=1)
        else:
            rows = [["NIM", "Nama", "Program Studi", "Aktivitas (Jenis - Status)", "MK sama"]]
            for _, r in multi.iterrows():
                rows.append([r["NIM"], Paragraph(esc(r["Nama"]), CELL), Paragraph(esc(r["Program Studi"]), CELL),
                             Paragraph(esc(r["Aktivitas (Jenis - Status)"]), CELL),
                             Paragraph(esc(r["MK Sama"] or "-"), CELL)])
            t = Table(rows, colWidths=[20 * mm, 40 * mm, 37 * mm, 48 * mm, 20 * mm], repeatRows=1)
        t.setStyle(tstyle())
        items.append(t)
    if multi_sem:
        cross = nim_info["cross"]
        items.append(Spacer(1, 6))
        if len(cross) == 0:
            items.append(Paragraph("Tidak ada mahasiswa yang tercatat di lebih dari satu semester.", CELL))
        else:
            n_same = int((cross[cross.columns[-1]] != "-").sum())
            items.append(Paragraph(
                "%d mahasiswa tercatat di lebih dari satu semester; %d di antaranya mengonversi MK yang sama pada "
                "lebih dari satu semester (F8 - cek konversi ganda)." % (len(cross), n_same), CELL))
            items.append(Spacer(1, 4))
            items.append(not_done_pdf_table(cross, fit_widths_mm(cross, 165, 7.6, 4), CELL, "Helvetica-Bold",
                                            tstyle(pad=4)))
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
        extra = ""
        if unc_len(unc) == len(pending):
            extra = " dan tidak dihitung dalam angka laporan ini"
        text = ("%d aktivitas (%d mahasiswa) belum berstatus Selesai%s: %s."
                % (len(pending), pending["NIM"].nunique(), extra, ", ".join(parts)))
        if multi_sem:
            sem_parts = []
            for p in periods:
                sem_parts.append("%s %d" % (p, int(((reg["Periode Akademik"] == p)
                                                      & (reg["Status Aktivitas"] != "Selesai")).sum())))
            text += " Per semester: %s." % ", ".join(sem_parts)
        items.append(Paragraph(text, CELL))
        items.append(Spacer(1, 4))
        items.append(not_done_pdf_table(pending, fit_widths_mm(pending, 165, 7.6, 4), CELL, "Helvetica-Bold",
                                        tstyle(pad=4)))
    if len(excluded):
        items.append(Spacer(1, 6))
        gone = []
        for s in meta["drop_status"]:
            if s not in IN_PROGRESS:
                gone.append(s)
        gone.sort(key=drop_order)
        items.append(Paragraph("Record yang dikecualikan oleh filter (%s): %d record."
                               % (", ".join(gone), len(excluded)), CELL))
        items.append(Spacer(1, 4))
        items.append(not_done_pdf_table(excluded, fit_widths_mm(excluded, 165, 7.6, 4), CELL, "Helvetica-Bold",
                                        tstyle(header_bg=red, pad=4)))
    section("7. Mahasiswa dengan Status Selain Selesai", items)

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm,
                            topMargin=16 * mm, bottomMargin=15 * mm, title="Laporan Analisis MBKM")
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


def widths_status(n_periods):
    """Lebar kolom (mm) tabel status per semester: label + tiap semester + Total = 165 mm."""
    value_w = 90.0 / (n_periods + 1)
    widths = [75]
    for c in range(n_periods + 1):
        widths.append(value_w)
    return widths


def fit_widths(df, total_pt, size, pad_pt, font="Helvetica", font_b="Helvetica-Bold"):
    """Lebar kolom (pt) sehingga kata terpanjang di setiap kolom (header + isi) muat tanpa terpotong;
    sisa lebar dibagi menurut kebutuhan teks penuh tiap kolom."""
    from reportlab.pdfbase.pdfmetrics import stringWidth
    mins = []
    wants = []
    for c in df.columns:
        head = str(c)
        min_w = 0.0
        for word in head.split():
            min_w = max(min_w, stringWidth(word, font_b, size))
        want = stringWidth(head, font_b, size)
        for value in df[c]:
            text = str(value)
            for word in text.split():
                min_w = max(min_w, stringWidth(word, font, size))
            want = max(want, stringWidth(text, font, size))
        mins.append(min_w + 2 * pad_pt + 2)
        wants.append(max(want, min_w) + 2 * pad_pt + 2)
    total_min = sum(mins)
    widths = []
    if total_min >= total_pt:
        for m in mins:
            widths.append(m * total_pt / total_min)
        return widths
    extra = total_pt - total_min
    need = []
    for m, w in zip(mins, wants):
        need.append(w - m)
    total_need = sum(need)
    if total_need <= extra:
        leftover = extra - total_need
        total_want = sum(wants)
        for w in wants:
            widths.append(w + leftover * w / total_want)
    else:
        for m, n in zip(mins, need):
            widths.append(m + extra * n / total_need)
    return widths


def fit_widths_mm(df, total_mm, size, pad_pt, font="Helvetica", font_b="Helvetica-Bold"):
    widths = []
    for w in fit_widths(df, total_mm * 72 / 25.4, size, pad_pt, font, font_b):
        widths.append(w * 25.4 / 72)
    return widths


def cohort_pdf_rows(act, meta, periods):
    """Baris 'Mahasiswa NIM 22...' untuk tabel ringkasan PDF: (label, nilai per semester, total)."""
    out = []
    prefixes = meta.get("cohorts", DEFAULT_COHORTS)
    totals, others = cohort_counts(act, prefixes)
    for p in prefixes:
        values = []
        if len(periods) > 1:
            for per in periods:
                counts, o = cohort_counts(act[act["Periode Akademik"] == per], [p])
                values.append(str(counts[p]))
        out.append(("Mahasiswa dengan NIM berawalan %s" % p, values, str(totals[p])))
    return tuple(out)


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
        "s1": "Ringkasan Eksekutif",
        "s_sem": "Perbandingan Semester",
        "s2": "Analisis per Program Studi",
        "s3": "Analisis per Jenis Aktivitas",
        "s_pending": "Mahasiswa dengan Status Selain Selesai",
        "s4": "Kualitas Data dan Tindak Lanjut",
        "s5": "Catatan Metodologi",
        "appendix": "Lampiran. Matriks Mahasiswa per Program Studi dan Jenis Aktivitas",
        "appendix_note": "Angka = jumlah mahasiswa (NIM unik); kolom dan baris Total juga dihitung sebagai NIM unik.",
        "actions": "Tindak lanjut yang disarankan:",
        "footer": "Laporan Naratif Aktivitas MBKM - Periode %s",
        "page": "Halaman ",
        "prodi_col": "Program Studi",
        "x_label": "Jumlah aktivitas",
        "x_label_stu": "Jumlah mahasiswa",
        "students_paren": "%d (%d mahasiswa)",
        "k_students": "Mahasiswa", "k_acts": "Aktivitas", "k_prodi": "Program studi",
        "k_done": "Aktivitas selesai", "k_ext": "Mitra eksternal", "k_mou": "Eksternal ber-MoU",
        "prodi_intro": ("Gambar {prodi_jenis} memperlihatkan komposisi jenis aktivitas di setiap program studi, dan "
                        "Gambar {prodi_mitra} memperlihatkan kategori mitranya: eksternal yang sudah atau belum "
                        "didukung MoU, internal i3L, serta mitra yang tidak valid. Uraian berikut diurutkan dari "
                        "program studi dengan jumlah mahasiswa terbanyak."),
        "jenis_intro": ("Gambar {jenis_status} memperlihatkan status aktivitas untuk setiap jenis aktivitas, dan "
                        "Gambar {jenis_mitra} memperlihatkan kategori mitranya. Uraian berikut diurutkan dari jenis "
                        "aktivitas dengan peserta terbanyak."),
        "sem_intro": ("Tabel {t_sem} merangkum indikator utama setiap semester, sedangkan Gambar {sem_prodi} dan "
                      "Gambar {sem_jenis} membandingkan jumlah mahasiswa per program studi dan per jenis aktivitas."),
        "file": "laporan_naratif",
        "doc_lang": "id-ID",
    },
    "en": {
        "title": "MBKM Activity Narrative Report",
        "institute": "Indonesia International Institute for Life Sciences (i3L)",
        "period": "Academic period: %s",
        "s1": "Executive Summary",
        "s_sem": "Semester Comparison",
        "s2": "Analysis by Study Program",
        "s3": "Analysis by Activity Type",
        "s_pending": "Students with a Status Other than Completed",
        "s4": "Data Quality and Follow-up",
        "s5": "Methodological Notes",
        "appendix": "Appendix. Student Matrix by Study Program and Activity Type",
        "appendix_note": "Values = number of students (unique NIM); the Total row and column also count unique NIMs.",
        "actions": "Recommended follow-up:",
        "footer": "MBKM Activity Narrative Report - %s",
        "page": "Page ",
        "prodi_col": "Study program",
        "x_label": "Number of activities",
        "x_label_stu": "Number of students",
        "students_paren": "%d (%d students)",
        "k_students": "Students", "k_acts": "Activities", "k_prodi": "Study programs",
        "k_done": "Completed", "k_ext": "External partners", "k_mou": "External with MoU",
        "prodi_intro": ("Figure {prodi_jenis} shows the mix of activity types in each study program, and Figure "
                        "{prodi_mitra} shows the partner categories: external partners with or without an MoU, "
                        "internal i3L activities, and invalid partner entries. Study programs are listed from the "
                        "largest number of students."),
        "jenis_intro": ("Figure {jenis_status} shows the status of activities for each activity type, and Figure "
                        "{jenis_mitra} shows their partner categories. Activity types are listed from the largest "
                        "number of participants."),
        "sem_intro": ("Table {t_sem} summarises the key indicators for each semester, while Figures {sem_prodi} and "
                      "{sem_jenis} compare the number of students by study program and by activity type."),
        "file": "narrative_report",
        "doc_lang": "en-US",
    },
}
FIG_CAPTIONS = {
    "id": {
        "prodi_jenis": "Gambar {n}. Jumlah aktivitas per program studi menurut jenis aktivitas. Angka di dalam batang "
                       "= jumlah aktivitas per jenis; angka di ujung batang = total aktivitas (jumlah mahasiswa "
                       "ditampilkan bila berbeda).",
        "sem_prodi": "Gambar {n}. Jumlah mahasiswa per program studi pada setiap semester (NIM unik dalam semester).",
        "sem_jenis": "Gambar {n}. Jumlah mahasiswa per jenis aktivitas pada setiap semester (NIM unik dalam semester).",
        "prodi_mitra": "Gambar {n}. Kategori mitra per program studi (jumlah aktivitas). Aktivitas internal i3L tidak "
                       "memerlukan MoU.",
        "jenis_status": "Gambar {n}. Status aktivitas per jenis aktivitas. Angka di ujung batang = total aktivitas "
                        "dan persentasenya terhadap seluruh aktivitas.",
        "jenis_mitra": "Gambar {n}. Kategori mitra per jenis aktivitas (jumlah aktivitas).",
    },
    "en": {
        "prodi_jenis": "Figure {n}. Number of activities per study program by activity type. Numbers inside the bars "
                       "= activities per type; number at the end of each bar = total activities (number of students "
                       "shown when different).",
        "sem_prodi": "Figure {n}. Number of students per study program in each semester (unique NIM per semester).",
        "sem_jenis": "Figure {n}. Number of students per activity type in each semester (unique NIM per semester).",
        "prodi_mitra": "Figure {n}. Partner category per study program (number of activities). Internal i3L "
                       "activities do not require an MoU.",
        "jenis_status": "Figure {n}. Activity status per activity type. Number at the end of each bar = total "
                        "activities and their share of all activities.",
        "jenis_mitra": "Figure {n}. Partner category per activity type (number of activities).",
    },
}
COMBINED_NOTE = {"id": " Data gabungan semua semester.", "en": " All semesters combined."}
TABLE_CAPTIONS = {
    "id": {
        "t_sem": "Tabel {n}. Indikator utama per semester.",
        "t_cross": "Tabel {n}. Mahasiswa yang mengikuti MBKM di lebih dari satu semester.",
        "t_pending": "Tabel {n}. Aktivitas yang belum berstatus Selesai.",
        "t_excluded": "Tabel {n}. Record yang dikecualikan dari analisis, beserta aktivitas lain mahasiswa yang masih "
                      "tercatat.",
    },
    "en": {
        "t_sem": "Table {n}. Key indicators by semester.",
        "t_cross": "Table {n}. Students who took part in MBKM in more than one semester.",
        "t_pending": "Table {n}. Activities not yet completed.",
        "t_excluded": "Table {n}. Records excluded from the analysis, with the student's other activity still on "
                      "record.",
    },
}
COMPARE_ROWS = {
    "id": ["Mahasiswa (NIM unik)", "Aktivitas", "Program studi", "Aktivitas selesai", "Aktivitas belum selesai",
           "Mitra eksternal", "Aktivitas bermitra eksternal", "– di antaranya ber-MoU", "Aktivitas internal i3L",
           "Total SKS dikonversi", "Rata-rata SKS per aktivitas", "Aktivitas dengan isu data per-record"],
    "en": ["Students (unique NIM)", "Activities", "Study programs", "Completed activities", "Not yet completed",
           "External partners", "Activities with external partners", "– of which with MoU", "Internal i3L activities",
           "Total credits (SKS) converted", "Average credits per activity", "Activities with record-level data issues"],
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
    if is_dup_status(name):
        base = str(name)[:-len(DUP_SUFFIX)]
        if lang == "en":
            return STATUS_EN.get(base, base) + " (duplicate)"
        return name
    if lang == "en":
        return STATUS_EN.get(name, name)
    return name


def status_state(name):
    if is_dup_status(name):
        base = str(name)[:-len(DUP_SUFFIX)]
        return STATUS_EN_STATE.get(base, base) + " (duplicate)"
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


def sem_label(period, lang, ay, form):
    """Label semester. form: 'cell' (Ganjil / Odd), 'col' (2025 Ganjil / Odd semester),
    'text' (Semester Ganjil / the odd semester). ay = tahun akademik bersama ('' bila berbeda)."""
    year, term = period_parts(period)
    if year is None:
        return period
    if lang == "en":
        word = TERM_EN.get(term.lower(), term.lower())
        if form == "cell":
            return word.capitalize() if ay else "%d %s" % (year, word.capitalize())
        if form == "col":
            return "%s semester" % word.capitalize() if ay else "%d %s semester" % (year, word)
        return "the %s semester" % word if ay else "the %d %s semester" % (year, word)
    if form == "cell":
        return term if ay else period
    if form == "col":
        return period
    return "Semester %s" % (term if ay else period)


def both_word(n_periods, lang):
    if lang == "en":
        return "both semesters" if n_periods == 2 else "more than one semester"
    return "kedua semester" if n_periods == 2 else "lebih dari satu semester"


def cap_first(text):
    return text[:1].upper() + text[1:]


def period_title(meta, lang):
    """(teks periode untuk kepala laporan, teks periode untuk footer)."""
    periods = meta.get("periods") or []
    if len(periods) <= 1:
        text = period_display(meta["periode"], lang)
        return text, text
    ay = academic_year(periods)
    if lang == "en":
        if ay:
            words = []
            for p in periods:
                term = period_parts(p)[1].lower()
                words.append(TERM_EN.get(term, term))
            return ("%s semesters of Academic Year %s (%s)" % (cap_first(join(words, "en")), ay, ", ".join(periods)),
                    "Academic Year %s" % ay)
        names = []
        for p in periods:
            names.append(period_in_text(p, "en"))
        return "%s (%s)" % (join(names, "en"), ", ".join(periods)), join(names, "en")
    if ay:
        return periode_text(periods), "Tahun Akademik %s" % ay
    return join(periods, "id"), join(periods, "id")


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
    """Ringkasan angka untuk satu kelompok aktivitas (satu prodi / satu jenis / satu semester / seluruh data)."""
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
    for k in f.columns:
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


def data_note(prof, scope, lang, multi_sem=False):
    f = prof["flags"]
    n = prof["n_act"]
    f1 = f["F1 Tanpa MK"]
    f2 = f["F2 MK Berlebih"]
    f3 = f["F3 Tanpa Pembimbing"]
    f4 = f["F4 Tanpa Penguji"]
    f5 = f["F5 Mitra Tidak Valid"]
    f8 = f.get(F8, 0)
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
    if f8:
        if en:
            issues.append(en_issue(f8, n, "every activity converts a course the same student also converted in "
                                          "another semester",
                                   "a course the same student also converted in another semester", "converts",
                                   "convert"))
        else:
            issues.append("%s mengonversi MK yang juga dikonversi mahasiswa yang sama pada semester lain"
                          % count_phrase(f8, n, lang))
    if k:
        if en:
            within = " within the same semester" if multi_sem else ""
            issues.append("%s %s recorded in more than one activity%s" % (studs(k, lang), plural(k, "is", "are"), within))
        else:
            within = " dalam semester yang sama" if multi_sem else ""
            issues.append("%d mahasiswa tercatat pada lebih dari satu aktivitas%s" % (k, within))
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


def drop_order(s):
    """Urutan tampil status yang dibuang: Ditolak, Dibatalkan, lalu duplikat."""
    base = str(s)[:-len(DUP_SUFFIX)] if is_dup_status(s) else s
    rank = STATUS_ORDER.index(base) if base in STATUS_ORDER else len(STATUS_ORDER)
    return (1 if is_dup_status(s) else 0, rank)


def split_drop(meta):
    """(status belum selesai yang tidak dihitung, status lain yang dikecualikan) dari daftar filter."""
    unc_status = []
    for s in IN_PROGRESS:
        if s in meta["drop_status"]:
            unc_status.append(s)
    others = []
    for s in meta["drop_status"]:
        if s not in IN_PROGRESS:
            others.append(s)
    others.sort(key=drop_order)
    return unc_status, others


def uncounted_sentence(meta, unc, lang):
    """Kalimat cakupan angka bila aktivitas belum selesai tidak dihitung (mis. hanya Selesai yang dihitung)."""
    unc_status, others = split_drop(meta)
    present = []
    for s in unc_status:
        if s in set(unc["Status Aktivitas"]):
            present.append(s)
    n_unc = len(unc)
    n_other = meta["n_dropped"] - n_unc
    if lang == "en":
        states = []
        for s in present:
            states.append(status_state(s))
        text = (" Activities not yet completed (%d %s %s) are not counted in these figures"
                % (n_unc, plural(n_unc, "activity", "activities"), join_or(states)))
        if n_other:
            names = []
            for s in others:
                names.append(status_name(s, lang))
            text += ", nor are %d %s with status %s" % (n_other, plural(n_other, "record", "records"), join_or(names))
        return text + "."
    text = (" Aktivitas yang belum selesai (%d aktivitas berstatus %s) tidak dihitung dalam angka ini"
            % (n_unc, join_or_id(present)))
    if n_other:
        text += ", demikian pula %d record berstatus %s" % (n_other, join_or_id(others))
    return text + "."


def status_text(status, n_act, unc, lang):
    """Kalimat status; bila ada aktivitas belum selesai yang tidak dihitung, keduanya disebut."""
    n_unc = unc_len(unc)
    if n_unc == 0:
        return status_sentence(status, n_act, lang)
    en = lang == "en"
    total = n_act + n_unc
    n_done = int(status.get("Selesai", 0))
    counted_other = []
    for name, n in ranked(status):
        if name != "Selesai":
            counted_other.append(("%d %s %s" % (n, plural(n, "is", "are"), status_state(name))) if en
                                 else ("%d berstatus %s" % (n, name)))
    unc_items = []
    for name, n in sorted(unc["Status Aktivitas"].value_counts().items(), key=lambda x: status_rank(x[0])):
        unc_items.append(("%d %s %s" % (n, plural(n, "is", "are"), status_state(name))) if en
                         else ("%d berstatus %s" % (n, name)))
    if en:
        text = ("Of the %d recorded activities, %d (%s) %s been completed"
                % (total, n_done, pct(n_done, total, lang), plural(n_done, "has", "have")))
        if counted_other:
            text += ", while " + join(counted_other, lang)
        return text + "; %s, so %s not counted in this report's figures." % (join(unc_items, lang),
                                                                            plural(n_unc, "it is", "they are"))
    text = ("Dari %d aktivitas yang tercatat, %d (%s) telah berstatus Selesai"
            % (total, n_done, pct(n_done, total, lang)))
    if counted_other:
        text += ", sementara " + join(counted_other, lang)
    return text + "; %s belum selesai sehingga tidak dihitung dalam angka laporan ini." % join(unc_items, lang)


def counting_scope(meta, lang):
    """Kalimat metode tentang status yang dihitung bila aktivitas belum selesai tidak dihitung."""
    unc_status, others = split_drop(meta)
    kept = []
    act_all = meta.get("act_all")
    if act_all is not None:
        present = set(act_all["Status Aktivitas"])
        for s in STATUS_ORDER:
            if s in present and s not in meta["drop_status"]:
                kept.append(s)
    en = lang == "en"
    if en:
        if kept == ["Selesai"]:
            text = "Only completed activities (Selesai) are counted."
        else:
            names = []
            for s in kept:
                names.append("%s (%s)" % (status_name(s, lang), s))
            text = "Activities with status %s are counted." % join_or(names)
        states = []
        for s in unc_status:
            states.append("%s (%s)" % (status_state(s), s))
        text += " Activities %s are reported separately as not yet completed" % join_or(states)
        if others:
            names = []
            for s in others:
                names.append(status_name(s, lang))
            text += ", while records with status %s are excluded from the analysis" % join_or(names)
        return text + "."
    if kept == ["Selesai"]:
        text = "Hanya aktivitas berstatus Selesai yang dihitung."
    else:
        text = "Aktivitas yang dihitung adalah yang berstatus %s." % join_or_id(kept)
    text += " Aktivitas berstatus %s dilaporkan terpisah sebagai aktivitas yang belum selesai" % join_or_id(unc_status)
    if others:
        text += ", sedangkan record berstatus %s dikecualikan dari analisis" % join_or_id(others)
    return text + "."


def dropped_sentence(meta, lang):
    if not meta["n_dropped"]:
        return ""
    unc = uncounted_records(meta.get("act_all"), meta["drop_status"])
    if unc_len(unc):
        return uncounted_sentence(meta, unc, lang)
    if lang == "en":
        statuses = []
        for s in meta["drop_status"]:
            statuses.append(status_name(s, lang))
        return (" These figures exclude %d %s with status %s."
                % (meta["n_dropped"], plural(meta["n_dropped"], "record", "records"), join_or(statuses)))
    return (" Angka ini sudah mengecualikan %d record berstatus %s."
            % (meta["n_dropped"], join(meta["drop_status"], lang)))


def cohort_sentence(act, meta, lang):
    """Kalimat jumlah mahasiswa per awal NIM (mis. NIM 22 dan 23)."""
    prefixes = meta.get("cohorts", DEFAULT_COHORTS)
    if not prefixes:
        return ""
    counts, others = cohort_counts(act, prefixes)
    en = lang == "en"
    parts = []
    for p in prefixes:
        n = counts[p]
        if en:
            if n == 0:
                parts.append("none have a NIM starting with %s" % p)
            else:
                parts.append("%s %s a NIM starting with %s" % (studs(n, lang), plural(n, "has", "have"), p))
        else:
            if n == 0:
                parts.append("tidak ada mahasiswa dengan NIM berawalan %s" % p)
            else:
                parts.append("%d mahasiswa memiliki NIM berawalan %s" % (n, p))
    if en:
        text = " By NIM, " + join(parts, lang)
    else:
        text = " Berdasarkan awal NIM, " + join(parts, lang)
    n_other = sum(others.values())
    if n_other:
        keys = sorted(others.keys())
        if en:
            text += "; the other %d %s a NIM starting with %s" % (n_other, plural(n_other, "has", "have"),
                                                                join_or(keys))
        else:
            text += "; %d mahasiswa lainnya memiliki NIM berawalan %s" % (n_other, join_or_id(keys))
    return text + "."


def join_or_id(items):
    items = list(items)
    if len(items) <= 1:
        return "".join(items)
    return ", ".join(items[:-1]) + " atau " + items[-1]


def summary_opening_multi(act, nim_info, meta, periods, n_prodi, lang, sec_no):
    """Paragraf pembuka ringkasan untuk laporan lebih dari satu semester (mis. satu tahun akademik)."""
    en = lang == "en"
    ay = academic_year(periods)
    n_act = len(act)
    n_stu = act["NIM"].nunique()
    per_sem = []
    for p in periods:
        n_p = int((act["Periode Akademik"] == p).sum())
        where = sem_label(p, lang, ay, "text")
        if en:
            per_sem.append("%s in %s (%s)" % (acts(n_p, lang), where, p))
        else:
            per_sem.append("%d aktivitas pada %s (%s)" % (n_p, where, p))
    if en:
        when = ("Academic Year %s" % ay) if ay else ("the period %s" % join(periods, lang))
        text = ("In %s, %s from %d %s took part in %s under the MBKM (Merdeka Belajar–Kampus Merdeka) scheme "
                "with course credit conversion: %s."
                % (when, studs(n_stu, lang), n_prodi, plural(n_prodi, "study program", "study programs"),
                   acts(n_act, lang), join(per_sem, lang)))
    else:
        when = ("Tahun Akademik %s" % ay) if ay else ("Periode Akademik %s" % join(periods, lang))
        text = ("Pada %s, sebanyak %d mahasiswa dari %d program studi tercatat mengikuti %d aktivitas MBKM yang "
                "dikonversi ke mata kuliah: %s." % (when, n_stu, n_prodi, n_act, join(per_sem, lang)))
    text += dropped_sentence(meta, lang)
    if n_act == n_stu:
        if en:
            return text + " Each student is recorded in exactly one activity."
        return text + " Setiap mahasiswa tercatat pada tepat satu aktivitas."
    per_nim = act.groupby("NIM")["Periode Akademik"].nunique()
    k_cross = int((per_nim > 1).sum())
    k_dup = 0
    if len(nim_info["multi"]):
        k_dup = int(nim_info["multi"]["NIM"].nunique())
    both = both_word(len(periods), lang)
    reasons = []
    if k_cross:
        if en:
            reasons.append("%s took part in MBKM in %s (see Section %d)" % (studs(k_cross, lang), both, sec_no["s_sem"]))
        else:
            reasons.append("%d mahasiswa mengikuti MBKM di %s (lihat Bagian %d)" % (k_cross, both, sec_no["s_sem"]))
    if k_dup:
        if en:
            reasons.append("%s %s recorded in more than one activity within the same semester (see Section %d)"
                           % (studs(k_dup, lang), plural(k_dup, "is", "are"), sec_no["s4"]))
        else:
            reasons.append("%d mahasiswa tercatat pada lebih dari satu aktivitas dalam semester yang sama "
                           "(lihat Bagian %d)" % (k_dup, sec_no["s4"]))
    if en:
        return text + " The number of activities exceeds the number of students because " + join(reasons, lang) + "."
    return text + " Jumlah aktivitas lebih besar daripada jumlah mahasiswa karena " + join(reasons, lang) + "."


def summary_paragraphs(act, cats, flags, nim_info, meta, prodi_counts, jenis_counts, lang, sec_no, unc=None):
    en = lang == "en"
    n_act = len(act)
    n_stu = act["NIM"].nunique()
    n_prodi = len(prodi_counts)
    k_multi = len(nim_info["multi"])
    periods = sorted_periods(act["Periode Akademik"].unique())
    if len(periods) > 1:
        p1 = summary_opening_multi(act, nim_info, meta, periods, n_prodi, lang, sec_no)
    elif en:
        p1 = ("In the %s, %s from %d %s took part in %s under the MBKM (Merdeka Belajar–Kampus Merdeka) scheme "
              "with course credit conversion."
              % (period_in_text(meta["periode"], lang), studs(n_stu, lang), n_prodi,
                 plural(n_prodi, "study program", "study programs"), acts(n_act, lang)))
        p1 += dropped_sentence(meta, lang)
        if k_multi == 0:
            p1 += " Each student is recorded in exactly one activity."
        else:
            p1 += (" The number of activities exceeds the number of students because %s %s recorded in more than "
                   "one activity (see Section %d)." % (studs(k_multi, lang), plural(k_multi, "is", "are"),
                                                       sec_no["s4"]))
    else:
        p1 = ("Pada Periode Akademik %s, sebanyak %d mahasiswa dari %d program studi tercatat mengikuti %d "
              "aktivitas MBKM yang dikonversi ke mata kuliah." % (meta["periode"], n_stu, n_prodi, n_act))
        p1 += dropped_sentence(meta, lang)
        if k_multi == 0:
            p1 += " Setiap mahasiswa tercatat pada tepat satu aktivitas."
        else:
            p1 += (" Jumlah aktivitas lebih besar daripada jumlah mahasiswa karena %d mahasiswa tercatat pada lebih "
                   "dari satu aktivitas (lihat Bagian %d)." % (k_multi, sec_no["s4"]))
    p1 += cohort_sentence(act, meta, lang)
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
    p3 = status_text(prof["status"], n_act, unc, lang)
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


def semester_split_sentence(sub, periods, lang):
    """Sebaran aktivitas satu kelompok (prodi / jenis) menurut semester."""
    en = lang == "en"
    ay = academic_year(periods)
    items = []
    present = []
    for p in periods:
        part = sub[sub["Periode Akademik"] == p]
        if len(part) == 0:
            continue
        present.append(p)
        text = acts(len(part), lang)
        n_s = part["NIM"].nunique()
        if n_s != len(part):
            text += " (%s)" % studs(n_s, lang)
        where = sem_label(p, lang, ay, "text")
        if not items:
            text += (" took place in %s" % where) if en else (" berlangsung pada %s" % where)
        else:
            text += (" in %s" % where) if en else (" pada %s" % where)
        items.append(text)
    if len(present) == 1:
        where = sem_label(present[0], lang, ay, "text")
        if en:
            if len(sub) == 1:
                return "This activity took place in %s." % where
            return "All activities took place in %s." % where
        if len(sub) == 1:
            return "Aktivitas tersebut berlangsung pada %s." % where
        return "Seluruh aktivitas berlangsung pada %s." % where
    if en:
        text = "Of these, " + join(items, lang) + "."
    else:
        text = "Sebanyak " + join(items, lang) + "."
    per_nim = sub.groupby("NIM")["Periode Akademik"].nunique()
    k = int((per_nim > 1).sum())
    if k:
        both = both_word(len(periods), lang)
        if en:
            lead = "One student" if k == 1 else "A total of %d students" % k
            text += " %s took part in %s." % (lead, both)
        else:
            lead = "satu mahasiswa" if k == 1 else "%d mahasiswa" % k
            text += " Di antaranya, %s mengikuti MBKM di %s." % (lead, both)
    return text


def prodi_story(act, cats, flags, mk_names, prodi_counts, n_stu_all, lang, periods=None, unc=None):
    multi = periods is not None and len(periods) > 1
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
        sentences = [first]
        if multi:
            sentences.append(semester_split_sentence(sub, periods, lang))
        sentences.append(mix_sentence(sub, n_stu, lang))
        sentences.append(status_text(prof["status"], prof["n_act"], unc_part(unc, "Program Studi", p), lang))
        para1 = " ".join(sentences)
        para2 = " ".join(partner_sentences(prof, lang) + [conversion_sentence(prof, mk_names, lang),
                                                          data_note(prof, "prodi", lang, multi)])
        stories.append((prodi_label(p), [para1, para2]))
    return stories


def jenis_story(act, cats, flags, mk_names, jenis_counts, n_stu_all, lang, periods=None, unc=None):
    multi = periods is not None and len(periods) > 1
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
        sentences = [first]
        if multi:
            sentences.append(semester_split_sentence(sub, periods, lang))
        sentences.append(composition_sentence(comp, lang))
        sentences.append(status_text(prof["status"], prof["n_act"], unc_part(unc, "Jenis Aktivitas", j), lang))
        para1 = " ".join(sentences)
        para2 = " ".join(partner_sentences(prof, lang) + [conversion_sentence(prof, mk_names, lang),
                                                          data_note(prof, "jenis", lang, multi)])
        stories.append((jenis_long(j, lang), [para1, para2]))
    return stories


def quality_section(act, cats, flags, nim_info, meta, lang, sec_no=None, periods=None):
    en = lang == "en"
    multi_sem = periods is not None and len(periods) > 1
    n_act = len(act)
    counts = {}
    for k in flags.columns:
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
    ay = ""
    if multi_sem:
        ay = academic_year(periods)
    within_en = " within the same semester" if multi_sem else ""
    within_id = " dalam semester yang sama" if multi_sem else ""
    if len(multi) == 0:
        if multi_sem and en:
            p2 = "Within each semester, every NIM appears in exactly one activity."
        elif multi_sem:
            p2 = "Dalam setiap semester, setiap NIM tercatat pada tepat satu aktivitas."
        elif en:
            p2 = ("The student ID (NIM) check shows that every NIM appears in exactly one activity, so the number of "
                  "activities equals the number of students.")
        else:
            p2 = ("Pemeriksaan NIM menunjukkan bahwa setiap NIM tercatat pada tepat satu aktivitas, sehingga jumlah "
                  "aktivitas sama dengan jumlah mahasiswa.")
    else:
        items = []
        dup_mk = False
        for _, r in multi.head(5).iterrows():
            rows = act[act["NIM"] == r["NIM"]]
            if multi_sem:
                rows = rows[rows["Periode Akademik"] == r["Semester"]]
            acts_list = []
            for _, a in rows.iterrows():
                acts_list.append("%s (%s)" % (jenis_long(a["Jenis Aktivitas"], lang),
                                              status_name(a["Status Aktivitas"], lang)))
            if en:
                text = "%s (%s) in %s" % (r["Nama"], r["NIM"], join(acts_list, lang))
                if multi_sem:
                    text += " in %s" % sem_label(r["Semester"], lang, ay, "text")
                if r["MK Sama"]:
                    text += ", converting the same course (%s)" % r["MK Sama"]
                    dup_mk = True
            else:
                text = "%s (%s) pada %s" % (r["Nama"], r["NIM"], join(acts_list, lang))
                if multi_sem:
                    text += " di %s" % sem_label(r["Semester"], lang, ay, "text")
                if r["MK Sama"]:
                    text += " dengan MK yang sama (%s)" % r["MK Sama"]
                    dup_mk = True
            items.append(text)
        if en:
            p2 = ("The student ID (NIM) check found %s recorded in more than one activity%s: %s."
                  % (studs(len(multi), lang), within_en, "; ".join(items)))
            if dup_mk:
                p2 += (" Activities that convert the same course are most likely duplicates or resubmissions and "
                       "should be verified.")
            if len(multi) > 5:
                p2 += " The full list is available in the Cek NIM tab and the Excel file."
        else:
            p2 = ("Pemeriksaan NIM menemukan %d mahasiswa yang tercatat pada lebih dari satu aktivitas%s: %s."
                  % (len(multi), within_id, "; ".join(items)))
            if dup_mk:
                p2 += (" Aktivitas yang mengonversi MK yang sama kemungkinan besar merupakan duplikasi atau pengajuan "
                       "ulang sehingga perlu diverifikasi.")
            if len(multi) > 5:
                p2 += " Daftar lengkap tersedia pada tab Cek NIM dan file Excel."
    n8 = counts.get(F8, 0)
    s8 = 0
    if F8 in flags.columns:
        s8 = int(act.loc[flags[F8], "NIM"].nunique())
    if multi_sem:
        cross = nim_info["cross"]
        both = both_word(len(periods), lang)
        if len(cross):
            if en:
                lead = "One student" if len(cross) == 1 else "A total of %d students" % len(cross)
                text = " %s took part in MBKM in %s (see Section %d)" % (lead, both, sec_no["s_sem"])
                if s8:
                    text += ("; %d of them converted the same course in more than one semester (F8, %s), which "
                             "should be verified so that no course is converted twice." % (s8, acts(n8, lang)))
                else:
                    text += "."
            else:
                lead = "Satu mahasiswa" if len(cross) == 1 else "Sebanyak %d mahasiswa" % len(cross)
                text = " %s mengikuti MBKM di %s (lihat Bagian %d)" % (lead, both, sec_no["s_sem"])
                if s8:
                    text += ("; %d di antaranya mengonversi MK yang sama pada lebih dari satu semester (F8, %d "
                             "aktivitas) sehingga perlu diverifikasi agar satu MK tidak dikonversi dua kali."
                             % (s8, n8))
                else:
                    text += "."
            p2 += text
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
            actions.append("Verify %d %s recorded in more than one activity%s."
                           % (len(multi), plural(len(multi), "NIM", "NIMs"), within_en))
        if n8:
            actions.append("Verify %s (%s) who converted the same course in more than one semester."
                           % (studs(s8, lang), acts(n8, lang)))
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
            actions.append("Memverifikasi %d NIM yang tercatat pada lebih dari satu aktivitas%s." % (len(multi), within_id))
        if n8:
            actions.append("Memverifikasi %d mahasiswa (%d aktivitas) yang mengonversi MK yang sama pada lebih dari "
                           "satu semester." % (s8, n8))
        if n6_ext:
            actions.append("Menindaklanjuti MoU kerja sama untuk %d aktivitas bermitra eksternal yang belum "
                           "didukung MoU." % n6_ext)
    return [p1, p2], actions


def method_paragraph(meta, lang, periods=None):
    multi = periods is not None and len(periods) > 1
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
        if split_drop(meta)[0]:
            drop = counting_scope(meta, lang)
        if any_dup(meta["drop_status"]):
            drop += (" \"Duplicate\" marks a submission (Diajukan/Evaluasi) that converts the same course as a completed "
                     "activity of the same student in the same semester.")
        extra = ""
        if multi:
            extra = (" Data for %s are combined in one report; students are counted as unique NIMs across all "
                     "semesters, so a student who took part in more than one semester is counted once in the totals "
                     "but in each semester in the semester comparison. Flag F7 compares activities within the same "
                     "semester only; F8 marks the same course converted in more than one semester."
                     % join(periods, "en"))
        return ("Data source: the SIAKAD export \"Laporan Aktivitas MBKM dan Mata Kuliah Konversi\"%s. Each row of "
                "the export represents one activity–course pair; rows are merged into one activity per NIM based on "
                "activity type, partner, dates, title, and activity status. Students are counted as unique NIMs, "
                "while the charts count activities. %s%s Partners are classified as internal when the name refers to "
                "i3L; MoU status is taken from the Status Mitra column. Threshold F2: more than %d courses (Student "
                "Exchange: more than %d courses). Study program names are shown as recorded in SIAKAD. Status terms: "
                "Completed = Selesai, Under evaluation = Evaluasi, Submitted = Diajukan, Rejected = Ditolak, "
                "Cancelled = Dibatalkan." % (source, drop, extra, meta["mk_overload"], meta["mk_overload_px"]))
    source = ""
    if meta["source_note"]:
        source = " (keterangan cetak tercantum di halaman pertama)"
    if meta["drop_status"]:
        drop = "Record berstatus %s dikecualikan dari analisis." % join(meta["drop_status"], lang)
    else:
        drop = "Tidak ada status aktivitas yang dikecualikan."
    if split_drop(meta)[0]:
        drop = counting_scope(meta, lang)
    if any_dup(meta["drop_status"]):
        drop += (" Status \"duplikat\" diberikan pada pengajuan (Diajukan/Evaluasi) yang mengonversi MK yang sama "
                 "dengan aktivitas Selesai milik mahasiswa yang sama pada semester yang sama.")
    extra = ""
    if multi:
        extra = (" Data %s digabung dalam satu laporan; jumlah mahasiswa dihitung sebagai NIM unik di semua semester, "
                 "sehingga mahasiswa yang mengikuti MBKM di lebih dari satu semester dihitung satu kali pada angka "
                 "total, tetapi dihitung di setiap semester pada perbandingan semester. Flag F7 hanya membandingkan "
                 "aktivitas dalam semester yang sama; F8 menandai MK yang sama yang dikonversi pada lebih dari satu "
                 "semester." % join(periods, "id"))
    return ("Data bersumber dari ekspor SIAKAD \"Laporan Aktivitas MBKM dan Mata Kuliah Konversi\"%s. "
            "Setiap baris ekspor mewakili satu pasangan aktivitas dan MK konversi; baris-baris tersebut digabung "
            "menjadi satu aktivitas per NIM berdasarkan jenis, mitra, tanggal, judul, dan status aktivitas. "
            "Jumlah mahasiswa dihitung sebagai NIM unik, sedangkan grafik menggunakan jumlah aktivitas. %s%s "
            "Mitra dikategorikan internal bila namanya merujuk ke i3L; status MoU diambil dari kolom Status Mitra. "
            "Ambang F2 adalah lebih dari %d MK (Pertukaran Pelajar: lebih dari %d MK)."
            % (source, drop, extra, meta["mk_overload"], meta["mk_overload_px"]))


def pending_paragraphs(act, meta, lang, pending, excluded, t_pending=1, t_excluded=2):
    """Narasi mahasiswa yang aktivitasnya belum Selesai dan record yang dikecualikan filter."""
    en = lang == "en"
    periods = sorted_periods(act["Periode Akademik"].unique())
    paras = []
    if len(pending) == 0:
        paras.append("All analysed activities have been completed." if en
                     else "Semua aktivitas yang dianalisis telah berstatus Selesai.")
    else:
        unc = uncounted_records(meta.get("act_all"), meta["drop_status"])
        raw_pending = with_uncounted(act[act["Status Aktivitas"] != "Selesai"], unc)
        n_rec = len(raw_pending)
        n_stu = raw_pending["NIM"].nunique()
        not_counted = unc_len(unc) == n_rec
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
            extra = (" and %s not counted in this report's figures" % plural(n_rec, "is", "are")) if not_counted else ""
            text = ("A total of %s %s not yet been completed%s: %s. They are listed in Table %d so that study "
                    "programs and supervisors can follow up. By study program: %s."
                    % (who, plural(n_rec, "has", "have"), extra, join(parts, lang), t_pending, join(by_prodi, lang)))
        else:
            if n_rec == n_stu:
                who = "%d aktivitas mahasiswa" % n_rec
            else:
                who = "%d aktivitas milik %d mahasiswa" % (n_rec, n_stu)
            extra = " dan tidak dihitung dalam angka laporan ini" if not_counted else ""
            text = ("Sebanyak %s belum berstatus Selesai%s: %s. Daftarnya tercantum pada Tabel %d agar dapat "
                    "ditindaklanjuti oleh program studi dan dosen pembimbing. Menurut program studi: %s."
                    % (who, extra, join(parts, lang), t_pending, join(by_prodi, lang)))
        if len(periods) > 1:
            ay = academic_year(periods)
            sem_items = []
            for p in periods:
                n_p = int((raw_pending["Periode Akademik"] == p).sum())
                if n_p:
                    sem_items.append(("%d in %s" if en else "%d pada %s") % (n_p, sem_label(p, lang, ay, "text")))
            text += (" By semester: %s." if en else " Menurut semester: %s.") % join(sem_items, lang)
        paras.append(text)
    if len(excluded):
        other_col = excluded.columns[-1]
        with_other = int((excluded[other_col] != "-").sum())
        no_other = excluded[excluded[other_col] == "-"]["NIM"].nunique()
        shown_status = []
        for s in meta["drop_status"]:
            if s not in IN_PROGRESS:
                shown_status.append(s)
        shown_status.sort(key=drop_order)
        statuses = []
        for s in shown_status:
            statuses.append(status_name(s, lang))
        scope_en = "in this period"
        scope_id = "pada periode ini"
        all_periods = meta.get("periods") or []
        if len(all_periods) > 1:
            if academic_year(all_periods):
                scope_en = "in this academic year"
                scope_id = "pada tahun akademik ini"
            else:
                scope_en = "in the analysed periods"
                scope_id = "pada periode yang dianalisis"
        if en:
            text = ("In addition, %d %s with status %s %s excluded from the analysis (Table %d). Of these, %d %s "
                    "to students who still have another activity on record, while %s %s no other MBKM activity "
                    "%s."
                    % (len(excluded), plural(len(excluded), "record", "records"), join_or(statuses),
                       was(len(excluded)), t_excluded, with_other, plural(with_other, "belongs", "belong"),
                       studs(no_other, lang), plural(no_other, "has", "have"), scope_en))
        else:
            text = ("Selain itu, %d record berstatus %s dikecualikan dari analisis (Tabel %d). Sebanyak %d di "
                    "antaranya milik mahasiswa yang masih memiliki aktivitas lain yang tercatat, sedangkan %d "
                    "mahasiswa tidak memiliki aktivitas MBKM lain %s."
                    % (len(excluded), join(shown_status, lang), t_excluded, with_other, no_other, scope_id))
        paras.append(text)
    return paras


def compare_values(sub, cats, flags, lang, unc_sub=None):
    """Nilai kolom tabel perbandingan semester (urutan sama dengan COMPARE_ROWS)."""
    prof = group_profile(sub, cats, flags)
    n_done = int(prof["status"].get("Selesai", 0))
    n_unc = unc_len(unc_sub)
    issues = 0
    if len(sub):
        issues = int(flags.loc[sub.index, RECORD_FLAGS].any(axis=1).sum())
    mou = "-"
    if prof["n_ext"]:
        mou = "%d (%s)" % (prof["n_mou"], pct(prof["n_mou"], prof["n_ext"], lang))
    return [num(prof["n_stu"], lang), num(prof["n_act"], lang), num(sub["Program Studi"].nunique(), lang),
            "%d (%s)" % (n_done, pct(n_done, prof["n_act"] + n_unc, lang)), num(prof["n_act"] - n_done + n_unc, lang),
            num(len(prof["partners"]), lang), num(prof["n_ext"], lang), mou, num(prof["n_int"], lang),
            num(prof["sks_total"], lang), num(prof["sks_avg"], lang), num(issues, lang)]


def semester_compare_table(act, cats, flags, lang, unc=None):
    """Tabel indikator utama: satu kolom per semester + kolom Total (mahasiswa = NIM unik)."""
    periods = sorted_periods(act["Periode Akademik"].unique())
    ay = academic_year(periods)
    columns = ["Indicator" if lang == "en" else "Indikator"]
    values = []
    for p in periods:
        columns.append(sem_label(p, lang, ay, "col"))
        values.append(compare_values(act[act["Periode Akademik"] == p], cats, flags, lang,
                                     unc_part(unc, "Periode Akademik", p)))
    total = "Total"
    if ay:
        total = ("Total %s" % ay) if lang == "en" else ("Total TA %s" % ay)
    columns.append(total)
    values.append(compare_values(act, cats, flags, lang, unc))
    rows = []
    for i, name in enumerate(COMPARE_ROWS[lang]):
        if unc_len(unc) and i == 3:
            name += " (% of recorded)" if lang == "en" else " (% dari yang tercatat)"
        if unc_len(unc) and i == 4:
            name += " (not counted)" if lang == "en" else " (tidak dihitung)"
        row = [name]
        for col_values in values:
            row.append(col_values[i])
        rows.append(row)
    return pd.DataFrame(rows, columns=columns)


def top_by_semester(subs, col, label, lang, ay, kind):
    """Kalimat kategori dengan peserta terbanyak di setiap semester. kind: 'prodi' atau 'jenis'."""
    en = lang == "en"
    tops = []
    for p, sub in subs:
        counts = ranked(sub.groupby(col)["NIM"].nunique())
        top_n = counts[0][1]
        names = []
        for name, n in counts:
            if n == top_n:
                names.append(label(name))
        tops.append((p, names, top_n))
    first_names = tops[0][1]
    same = len(first_names) == 1
    for p, names, n in tops:
        if names != first_names:
            same = False
    both = both_word(len(subs), lang)
    if same:
        values = []
        for p, names, n in tops:
            values.append(str(n))
        if en:
            if kind == "prodi":
                return "%s had the most participants in %s (%s students)." % (first_names[0], both, join(values, lang))
            return "%s was the most common activity type in %s (%s students)." % (first_names[0], both,
                                                                                   join(values, lang))
        if kind == "prodi":
            return ("%s menjadi program studi dengan peserta terbanyak pada %s (%s mahasiswa)."
                    % (first_names[0], both, join(values, lang)))
        return "%s menjadi jenis aktivitas terbanyak pada %s (%s mahasiswa)." % (first_names[0], both,
                                                                                  join(values, lang))
    items = []
    for p, names, n in tops:
        where = sem_label(p, lang, ay, "text")
        who = join(names, lang)
        if en:
            count = studs(n, lang) if len(names) == 1 else "%s each" % studs(n, lang)
            items.append("%s (%s) in %s" % (who, count, where))
        else:
            count = ("%d mahasiswa" % n) if len(names) == 1 else ("masing-masing %d mahasiswa" % n)
            items.append("pada %s adalah %s (%s)" % (where, who, count))
    if en:
        if kind == "prodi":
            return "The study program with the most participants was " + join(items, lang) + "."
        return "The most common activity type was " + join(items, lang) + "."
    lead = "Program studi dengan peserta terbanyak " if kind == "prodi" else "Jenis aktivitas terbanyak "
    if len(items) == 2:
        return lead + items[0] + ", sedangkan " + items[1] + "."
    return lead + join(items, lang) + "."


def semester_section(act, cats, flags, lang, numbers, unc=None):
    """Bagian perbandingan semester: 4 paragraf, tabel indikator, dan tabel mahasiswa lintas semester."""
    en = lang == "en"
    T = NARR_TEXT[lang]
    periods = sorted_periods(act["Periode Akademik"].unique())
    ay = academic_year(periods)
    both = both_word(len(periods), lang)
    if en:
        scope = "this academic year" if ay else "the analysed period"
    else:
        scope = "tahun akademik ini" if ay else "periode yang dianalisis"
    subs = []
    for p in periods:
        subs.append((p, act[act["Periode Akademik"] == p]))

    # 1. gambaran umum
    items = []
    for p, sub in subs:
        where = sem_label(p, lang, ay, "text")
        if en:
            items.append("%s recorded %s in %s" % (where, studs(sub["NIM"].nunique(), lang), acts(len(sub), lang)))
        else:
            items.append("%s mencatat %d mahasiswa dalam %d aktivitas" % (where, sub["NIM"].nunique(), len(sub)))
    if len(items) == 2:
        overview = items[0] + (", while " if en else ", sedangkan ") + items[1] + "."
    else:
        overview = join(items, lang) + "."
    top_p = subs[0][0]
    top_n = len(subs[0][1])
    for p, sub in subs:
        if len(sub) > top_n:
            top_p = p
            top_n = len(sub)
    share = pct(top_n, len(act), lang)
    if en:
        share_text = ("%s therefore accounts for %s of all activities in %s."
                      % (cap_first(sem_label(top_p, lang, ay, "text")), share, scope))
    else:
        share_text = ("Dengan demikian, %s menyumbang %s dari seluruh aktivitas pada %s."
                      % (sem_label(top_p, lang, ay, "text"), share, scope))
    change = ""
    if len(subs) == 2:
        pa, sa = subs[0]
        pb, sb = subs[1]
        na = sa["NIM"].nunique()
        nb = sb["NIM"].nunique()
        a = sem_label(pa, lang, ay, "text")
        b = sem_label(pb, lang, ay, "text")
        diff = nb - na
        if diff == 0:
            if en:
                change = " The number of students in %s was the same as in %s." % (b, a)
            else:
                change = " Jumlah mahasiswa pada %s sama dengan %s." % (b, a)
        else:
            rel = ""
            if na:
                rel = " (%s)" % pct(abs(diff), na, lang)
            if en:
                change = (" Compared with %s, the number of students in %s %s by %d%s."
                          % (a, b, "rose" if diff > 0 else "fell", abs(diff), rel))
            else:
                change = (" Dibandingkan %s, jumlah mahasiswa pada %s %s sebanyak %d mahasiswa%s."
                          % (a, b, "naik" if diff > 0 else "turun", abs(diff), rel))
    p_overview = T["sem_intro"].format(**numbers) + " " + cap_first(overview) + " " + share_text + change

    # 2. program studi dan jenis aktivitas
    prog_items = []
    for p, sub in subs:
        n_prog = sub["Program Studi"].nunique()
        where = sem_label(p, lang, ay, "text")
        if en:
            prog_items.append("%d %s in %s" % (n_prog, plural(n_prog, "study program", "study programs"), where))
        else:
            prog_items.append("%d program studi pada %s" % (n_prog, where))
    p_mix = ("Participants came from " if en else "Peserta berasal dari ") + join(prog_items, lang) + "."
    for p, sub in subs:
        others = set()
        for q, other in subs:
            if q != p:
                for name in other["Program Studi"].unique():
                    others.add(name)
        only = []
        for name in sorted(sub["Program Studi"].unique()):
            if name not in others:
                only.append(prodi_label(name))
        if only:
            where = sem_label(p, lang, ay, "text")
            if en:
                p_mix += " %s %s only in %s." % (join(only, lang), plural(len(only), "appears", "appear"), where)
            else:
                p_mix += " %s hanya tercatat pada %s." % (join(only, lang), where)

    def jenis_text(name):
        return jenis_long(name, lang)

    p_mix += " " + top_by_semester(subs, "Program Studi", prodi_label, lang, ay, "prodi")
    p_mix += " " + top_by_semester(subs, "Jenis Aktivitas", jenis_text, lang, ay, "jenis")

    # 3. penyelesaian, MoU, SKS
    done_items = []
    mou_items = []
    no_ext = []
    sks_items = []
    for p, sub in subs:
        prof = group_profile(sub, cats, flags)
        where = sem_label(p, lang, ay, "text")
        n_done = int(prof["status"].get("Selesai", 0))
        n_reg = prof["n_act"] + unc_len(unc_part(unc, "Periode Akademik", p))
        done_items.append(("%s in %s" if en else "%s pada %s") % (pct(n_done, n_reg, lang), where))
        if prof["n_ext"]:
            mou_items.append(("%s in %s" if en else "%s pada %s") % (pct(prof["n_mou"], prof["n_ext"], lang), where))
        else:
            no_ext.append(where)
        if en:
            sks_items.append("%s in %s" % (num(prof["sks_avg"], lang), where))
        else:
            sks_items.append("%s SKS pada %s" % (num(prof["sks_avg"], lang), where))
    if en:
        p_rates = "The completion rate was %s." % join(done_items, lang)
        if mou_items:
            p_rates += " Among externally partnered activities, the share covered by an MoU was %s." % join(mou_items, lang)
        if no_ext:
            p_rates += " There were no externally partnered activities in %s." % join(no_ext, lang)
        p_rates += " On average, the credits converted per activity were %s." % join(sks_items, lang)
    else:
        p_rates = "Tingkat penyelesaian aktivitas adalah %s." % join(done_items, lang)
        if mou_items:
            p_rates += (" Dari aktivitas bermitra eksternal, proporsi yang sudah didukung MoU kerja sama adalah %s."
                        % join(mou_items, lang))
        if no_ext:
            p_rates += " Tidak ada aktivitas bermitra eksternal pada %s." % join(no_ext, lang)
        p_rates += " Rata-rata SKS yang dikonversi per aktivitas adalah %s." % join(sks_items, lang)

    # 4. mahasiswa lintas semester
    cross = cross_semester_table(act, lang)
    sum_sem = 0
    for p, sub in subs:
        sum_sem += sub["NIM"].nunique()
    n_unique = act["NIM"].nunique()
    if len(cross) == 0:
        p_cross = ("No student took part in MBKM in %s." % both) if en else (
            "Tidak ada mahasiswa yang mengikuti MBKM di %s." % both)
    else:
        same_col = cross.columns[-1]
        name_col = cross.columns[1]
        same_rows = cross[cross[same_col] != "-"]
        if en:
            lead = "One student" if len(cross) == 1 else "A total of %d students" % len(cross)
            p_cross = ("%s took part in MBKM in %s, so %s counts %d unique students rather than the %d obtained by "
                       "adding the semesters. They are listed in Table %d."
                       % (lead, both, scope, n_unique, sum_sem, numbers["t_cross"]))
        else:
            lead = "Satu mahasiswa" if len(cross) == 1 else "Sebanyak %d mahasiswa" % len(cross)
            p_cross = ("%s mengikuti MBKM di %s, sehingga jumlah mahasiswa unik pada %s adalah %d, bukan %d seperti "
                       "hasil penjumlahan per semester. Daftarnya tercantum pada Tabel %d."
                       % (lead, both, scope, n_unique, sum_sem, numbers["t_cross"]))
        if len(same_rows):
            names = []
            for _, r in same_rows.head(8).iterrows():
                names.append("%s (%s)" % (r[name_col], r[same_col]))
            if len(same_rows) > 8:
                more = len(same_rows) - 8
                names.append(("%d other students" % more) if en else ("%d mahasiswa lainnya" % more))
            if en:
                p_cross += (" Of these, %d converted the same course in %s: %s. These activities are flagged F8 and "
                            "should be verified so that no course is converted twice."
                            % (len(same_rows), both, join(names, lang)))
            else:
                p_cross += (" Sebanyak %d di antaranya mengonversi MK yang sama pada %s: %s. Aktivitas tersebut "
                            "ditandai F8 dan perlu diverifikasi agar satu MK tidak dikonversi dua kali."
                            % (len(same_rows), both, join(names, lang)))
        elif en:
            p_cross += " None of them converted the same course in %s." % both
        else:
            p_cross += " Tidak ada di antaranya yang mengonversi MK yang sama pada %s." % both
    return {"paras": [p_overview, p_mix, p_rates, p_cross],
            "compare": semester_compare_table(act, cats, flags, lang, unc), "cross": cross}


UNC_FIG_NOTE = {
    "id": " Grafik ini juga memuat aktivitas yang belum selesai (Evaluasi/Diajukan), yang tidak dihitung pada angka lain.",
    "en": (" This chart also includes activities not yet completed (under evaluation or awaiting approval), which are "
           "not counted elsewhere."),
}


def narrative_content(act, raw, flags, nim_info, meta, cats, lang):
    """Seluruh teks laporan naratif dalam satu bahasa (dipakai oleh PDF, Word, dan tab Narasi)."""
    T = NARR_TEXT[lang]
    periods = sorted_periods(act["Periode Akademik"].unique())
    multi = len(periods) > 1
    mk_names = mk_name_map(raw)
    n_act = len(act)
    n_stu = act["NIM"].nunique()
    prodi_counts = ranked(act.groupby("Program Studi")["NIM"].nunique())
    jenis_counts = ranked(act.groupby("Jenis Aktivitas")["NIM"].nunique())
    pending_df, excluded_df = not_done_tables(act, meta.get("act_all"), meta["drop_status"], lang)
    unc = uncounted_records(meta.get("act_all"), meta["drop_status"])
    cross = cross_semester_table(act, lang)

    # tata letak: nomor bagian, gambar, dan tabel mengikuti urutan kemunculan
    sec_keys = ["s1"]
    fig_keys = ["prodi_jenis"]
    tab_keys = []
    if multi:
        sec_keys.append("s_sem")
        fig_keys += ["sem_prodi", "sem_jenis"]
        tab_keys.append("t_sem")
        if len(cross):
            tab_keys.append("t_cross")
    sec_keys += ["s2", "s3", "s_pending", "s4", "s5"]
    fig_keys += ["prodi_mitra", "jenis_status", "jenis_mitra"]
    if len(pending_df):
        tab_keys.append("t_pending")
    if len(excluded_df):
        tab_keys.append("t_excluded")
    sec_no = {}
    titles = {}
    for i, key in enumerate(sec_keys, start=1):
        sec_no[key] = i
        titles[key] = "%d. %s" % (i, T[key])
    titles["appendix"] = T["appendix"]
    numbers = {}
    captions = {}
    for i, key in enumerate(fig_keys, start=1):
        numbers[key] = i
        text = FIG_CAPTIONS[lang][key].format(n=i)
        if multi and key not in ("sem_prodi", "sem_jenis"):
            text += COMBINED_NOTE[lang]
        captions[key] = text
    for i, key in enumerate(tab_keys, start=1):
        numbers[key] = i
        captions[key] = TABLE_CAPTIONS[lang][key].format(n=i)
    if unc_len(unc):
        captions["jenis_status"] += UNC_FIG_NOTE[lang]

    summary, overall = summary_paragraphs(act, cats, flags, nim_info, meta, prodi_counts, jenis_counts, lang, sec_no,
                                          unc)
    quality, actions = quality_section(act, cats, flags, nim_info, meta, lang, sec_no, periods)
    n_done = int(overall["status"].get("Selesai", 0))
    mou_share = "-"
    if overall["n_ext"]:
        mou_share = pct(overall["n_mou"], overall["n_ext"], lang)
    kpi = [(T["k_students"], num(n_stu, lang)), (T["k_acts"], num(n_act, lang)),
           (T["k_prodi"], num(len(prodi_counts), lang)), (T["k_done"], pct(n_done, n_act + unc_len(unc), lang)),
           (T["k_ext"], num(len(overall["partners"]), lang)), (T["k_mou"], mou_share)]
    pending_text = pending_paragraphs(act, meta, lang, pending_df, excluded_df,
                                      numbers.get("t_pending", 1), numbers.get("t_excluded", 2))
    semester = None
    if multi:
        semester = semester_section(act, cats, flags, lang, numbers, unc)
    return {
        "lang": lang,
        "kpi": kpi,
        "summary": summary,
        "semester": semester,
        "prodi_intro": T["prodi_intro"].format(**numbers),
        "prodi": prodi_story(act, cats, flags, mk_names, prodi_counts, n_stu, lang, periods, unc),
        "jenis_intro": T["jenis_intro"].format(**numbers),
        "jenis": jenis_story(act, cats, flags, mk_names, jenis_counts, n_stu, lang, periods, unc),
        "pending": pending_text,
        "pending_df": pending_df,
        "excluded_df": excluded_df,
        "quality": quality,
        "actions": actions,
        "method": method_paragraph(meta, lang, periods),
        "matrix": student_matrix(act, lang),
        "titles": titles,
        "sec_no": sec_no,
        "captions": captions,
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


def sem_colors(n):
    """Warna semester: satu hue navy, gelap -> terang menurut urutan waktu (tervalidasi untuk 2 dan 3 langkah)."""
    if n in SEM_COLORS:
        return list(SEM_COLORS[n])
    dark = (0x1f, 0x38, 0x64)
    light = (0x8f, 0xa4, 0xc8)
    out = []
    for i in range(n):
        t = i / float(max(n - 1, 1))
        rgb = []
        for d, l in zip(dark, light):
            rgb.append(int(round(d + (l - d) * t)))
        out.append("#%02x%02x%02x" % (rgb[0], rgb[1], rgb[2]))
    return out


def grouped_barh_png(table, colors, xlabel):
    """Batang horizontal berkelompok. table: index = kategori (atas ke bawah), kolom = seri (mis. semester)."""
    n = len(table.index)
    k = len(table.columns)
    fig = plt.figure(figsize=(6.8, 0.95 + (0.22 * k + 0.16) * n))
    canvas = FigureCanvasAgg(fig)  # kanvas Agg eksplisit: aman saat Streamlit menjalankan ulang skrip
    group_h = 0.8
    bar_h = group_h / k
    base = []
    for i in range(n):
        base.append(n - 1 - i)
    xmax = 1.0
    for col in table.columns:
        for cat in table.index:
            xmax = max(xmax, float(table.loc[cat, col]))
    for j, col in enumerate(table.columns):
        ys = []
        vals = []
        for i, cat in enumerate(table.index):
            ys.append(base[i] + group_h / 2.0 - bar_h * (j + 0.5))
            vals.append(float(table.loc[cat, col]))
        plt.barh(ys, vals, height=bar_h, color=colors[col], edgecolor="white", linewidth=1.2, label=col, zorder=3)
        for y, v in zip(ys, vals):
            plt.text(v + xmax * 0.012, y, "%d" % v, ha="left", va="center", fontsize=7.5, fontweight="bold",
                     color=INK, zorder=4)
    plt.yticks(base, list(table.index), fontsize=8.5, color=INK2)
    plt.xticks(fontsize=7.5, color=MUTED)
    plt.xlim(0, xmax * 1.12)
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
    plt.legend(loc="lower left", bbox_to_anchor=(0.0, 1.0), ncol=k, frameon=False,
               fontsize=8, handlelength=0.9, handleheight=0.9, handletextpad=0.5, columnspacing=1.4,
               borderaxespad=0.3, labelcolor=INK2)
    plt.tight_layout()
    canvas.draw()
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
def narrative_charts(act, cats, lang, unc=None):
    """Grafik laporan naratif sebagai PNG (bytes), label sesuai bahasa; + 2 grafik semester bila >1 semester."""
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
    reg = with_uncounted(act, unc)
    status_order = []
    for s in STATUS_ORDER:
        if s in set(reg["Status Aktivitas"]):
            status_order.append(s)
    for s in sorted(set(reg["Status Aktivitas"])):
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
    reg_order = list(jenis_order)
    for j, n in ranked(reg.groupby("Jenis Aktivitas").size()):
        if j not in reg_order:
            reg_order.append(j)
    reg_shares = jenis_shares
    if len(reg) != len(act):
        reg_shares = {}
        for j in reg_order:
            n = int((reg["Jenis Aktivitas"] == j).sum())
            reg_shares[jenis_lbl(j)] = "%d (%s)" % (n, pct(n, len(reg), lang))
    t = count_table(reg["Jenis Aktivitas"], reg["Status Aktivitas"], reg_order, status_order, jenis_lbl, status_lbl)
    charts["jenis_status"] = stacked_barh_png(t, status_colors, reg_shares, T["x_label"])
    t = count_table(act["Jenis Aktivitas"], cats, jenis_order, PARTNER_ORDER, jenis_lbl, cat_lbl)
    charts["jenis_mitra"] = stacked_barh_png(t, partner_colors, jenis_totals, T["x_label"])

    periods = sorted_periods(act["Periode Akademik"].unique())
    if len(periods) > 1:
        ay = academic_year(periods)
        palette = sem_colors(len(periods))
        series = []
        sem_palette = {}
        for i, p in enumerate(periods):
            label = sem_label(p, lang, ay, "col")
            series.append(label)
            sem_palette[label] = palette[i]
        for key, by, order, label_of in (("sem_prodi", "Program Studi", prodi_order, prodi_label),
                                         ("sem_jenis", "Jenis Aktivitas", jenis_order, jenis_lbl)):
            rows = []
            index = []
            for cat in order:
                row = []
                for p in periods:
                    part = act[(act[by] == cat) & (act["Periode Akademik"] == p)]
                    row.append(int(part["NIM"].nunique()))
                rows.append(row)
                index.append(label_of(cat))
            table = pd.DataFrame(rows, index=index, columns=series)
            charts[key] = grouped_barh_png(table, sem_palette, T["x_label_stu"])
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
    titles = content["titles"]
    captions = content["captions"]
    sec_no = content["sec_no"]
    period_head, period_foot = period_title(meta, lang)
    font, font_b, font_i = register_fonts()
    navy = colors.HexColor(NAVY)
    red = colors.HexColor(RED)
    light = colors.HexColor("#EEF1F7")
    grey = colors.HexColor("#D9D9D9")
    mute = colors.HexColor("#666666")
    ink = colors.HexColor(INK)
    page_w = A4[0] - 36 * mm
    page_w_mm = 174

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
    CELLC = ParagraphStyle("NCELLC", parent=CELL, alignment=TA_CENTER)
    CELLCB = ParagraphStyle("NCELLCB", parent=CELLC, fontName=font_b)
    HEAD = ParagraphStyle("NHEAD", parent=CELL, fontName=font_b, fontSize=7.5, leading=9.5,
                          textColor=colors.white, alignment=TA_CENTER)

    def figure(key):
        png = charts[key]
        w, h = ImageReader(io.BytesIO(png)).getSize()
        img = Image(io.BytesIO(png), width=page_w, height=page_w * h / float(w))
        return KeepTogether([Spacer(1, 2), img, Paragraph(esc(captions[key]), CAP)])

    list_style = TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4),
        ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#BFBFBF")),
        ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, light]),
    ])

    def list_table(df, header_bg):
        widths_mm = fit_widths_mm(df, page_w_mm, 7.6, 4, font, font_b)
        table = not_done_pdf_table(df, widths_mm, CELL, font_b, list_style)
        table.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), header_bg)]))
        return table

    def compare_table(df):
        header = []
        for c in df.columns:
            header.append(Paragraph(esc(c), HEAD))
        rows = [header]
        last = len(df.columns) - 1
        for _, r in df.iterrows():
            row = []
            for i, c in enumerate(df.columns):
                if i == 0:
                    row.append(Paragraph(esc(r[c]), CELL))
                elif i == last:
                    row.append(Paragraph(esc(r[c]), CELLCB))
                else:
                    row.append(Paragraph(esc(r[c]), CELLC))
            rows.append(row)
        first = 64 * mm
        widths = [first]
        for c in range(last):
            widths.append((page_w - first) / last)
        table = Table(rows, colWidths=widths, repeatRows=1)
        table.setStyle(TableStyle([
            ("BACKGROUND", (0, 0), (-1, 0), navy), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
            ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
            ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#BFBFBF")),
            ("ROWBACKGROUNDS", (0, 1), (-1, -1), [colors.white, light]),
        ]))
        return table

    story = []
    story.append(Paragraph(esc(T["title"]), H1))
    story.append(Paragraph(esc(T["institute"]), SUB))
    story.append(Paragraph(esc(T["period"] % period_head), SUB))
    if meta["source_note"]:
        for line in meta["source_note"].split("\n"):
            story.append(Paragraph(esc(source_note_text(line, lang)).replace("|", "&nbsp;|&nbsp;"), SUB))
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

    story.append(Paragraph(esc(titles["s1"]), H2))
    for p in content["summary"]:
        story.append(Paragraph(esc(p), BODY))
    story.append(figure("prodi_jenis"))

    sem = content.get("semester")
    if sem:
        story.append(Paragraph(esc(titles["s_sem"]), H2))
        story.append(Paragraph(esc(sem["paras"][0]), BODY))
        story.append(KeepTogether([Paragraph(esc(captions["t_sem"]), CAP), compare_table(sem["compare"])]))
        story.append(Spacer(1, 8))
        story.append(Paragraph(esc(sem["paras"][1]), BODY))
        story.append(figure("sem_prodi"))
        story.append(figure("sem_jenis"))
        story.append(Paragraph(esc(sem["paras"][2]), BODY))
        story.append(Paragraph(esc(sem["paras"][3]), BODY))
        if len(sem["cross"]):
            story.append(Paragraph(esc(captions["t_cross"]), CAP))
            story.append(list_table(sem["cross"], navy))
            story.append(Spacer(1, 8))

    story.append(Paragraph(esc(titles["s2"]), H2))
    story.append(Paragraph(esc(content["prodi_intro"]), BODY))
    story.append(figure("prodi_mitra"))
    for i, (name, paras) in enumerate(content["prodi"], start=1):
        block = [Paragraph("%d.%d %s" % (sec_no["s2"], i, esc(name)), H3)]
        for p in paras:
            block.append(Paragraph(esc(p), BODY))
        story.append(KeepTogether(block))

    story.append(Paragraph(esc(titles["s3"]), H2))
    story.append(Paragraph(esc(content["jenis_intro"]), BODY))
    story.append(figure("jenis_status"))
    story.append(figure("jenis_mitra"))
    for i, (name, paras) in enumerate(content["jenis"], start=1):
        block = [Paragraph("%d.%d %s" % (sec_no["s3"], i, esc(name)), H3)]
        for p in paras:
            block.append(Paragraph(esc(p), BODY))
        story.append(KeepTogether(block))

    block = [Paragraph(esc(titles["s_pending"]), H2)]
    for p in content["pending"]:
        block.append(Paragraph(esc(p), BODY))
    story.append(KeepTogether(block))
    for key, cap_key, bg in (("pending_df", "t_pending", navy), ("excluded_df", "t_excluded", red)):
        df = content[key]
        if len(df):
            story.append(Paragraph(esc(captions[cap_key]), CAP))
            story.append(list_table(df, bg))
            story.append(Spacer(1, 8))

    story.append(Paragraph(esc(titles["s4"]), H2))
    for p in content["quality"]:
        story.append(Paragraph(esc(p), BODY))
    if content["actions"]:
        items = []
        for a in content["actions"]:
            items.append(ListItem(Paragraph(esc(a), BODY), leftIndent=12))
        story.append(KeepTogether([Paragraph(esc(T["actions"]), BODY),
                                   ListFlowable(items, bulletType="bullet", start="•", leftIndent=12,
                                                bulletFontName=font, bulletFontSize=8)]))

    story.append(Paragraph(esc(titles["s5"]), H2))
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
    story.append(KeepTogether([Paragraph(esc(titles["appendix"]), H2), Paragraph(esc(T["appendix_note"]), CAP),
                               table]))

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setFont(font, 7.2)
        canvas.setFillColor(mute)
        canvas.drawString(18 * mm, 9 * mm, T["footer"] % period_foot)
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


def docx_list_table(doc, df, widths_cm, header_fill, center_from=None, bold_last=False):
    """Tabel Word untuk daftar (header berwarna, baris berselang, lebar kolom tetap).

    center_from: indeks kolom pertama yang rata tengah; bold_last: kolom terakhir dicetak tebal (mis. Total).
    """
    from docx.shared import Pt, Cm, RGBColor
    from docx.enum.table import WD_TABLE_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    table = doc.add_table(rows=len(df.index) + 1, cols=len(df.columns))
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    widths = []
    for w in widths_cm:
        widths.append(Cm(w))
    grid = (4, "BFBFBF")
    docx_table_layout(table, widths, {"top": grid, "left": grid, "bottom": grid, "right": grid,
                                      "insideH": grid, "insideV": grid})
    last = len(df.columns) - 1
    for i, c in enumerate(df.columns):
        cell = table.cell(0, i)
        docx_shade(cell, header_fill)
        par = cell.paragraphs[0]
        par.paragraph_format.space_after = Pt(0)
        if center_from is not None and i >= center_from:
            par.alignment = WD_ALIGN_PARAGRAPH.CENTER
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
            if center_from is not None and i >= center_from:
                par.alignment = WD_ALIGN_PARAGRAPH.CENTER
            run = par.add_run(str(values[c]))
            run.font.size = Pt(8)
            if bold_last and i == last:
                run.bold = True
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
    titles = content["titles"]
    captions = content["captions"]
    sec_no = content["sec_no"]
    period_head, period_foot = period_title(meta, lang)
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

    def caption(text, after):
        cap = small_line(text, size=8.5, after=after)
        for run in cap.runs:
            run.italic = True
        return cap

    def figure(key):
        pic = doc.add_paragraph()
        pic.alignment = WD_ALIGN_PARAGRAPH.CENTER
        pic.paragraph_format.keep_with_next = True
        pic.paragraph_format.space_after = Pt(2)
        pic.add_run().add_picture(io.BytesIO(charts[key]), width=usable)
        caption(captions[key], 10)

    def fitted_cm(df):
        """Lebar kolom (cm) dari metrik huruf (Helvetica 8 pt, sedikit lebih lebar dari Calibri = aman)."""
        widths = []
        for w in fit_widths(df, 17.4 * PT_PER_CM, 8, 5.4):
            widths.append(w / PT_PER_CM)
        return widths

    def table_with_caption(key, df, widths_cm, fill, center_from=None, bold_last=False):
        cap = caption(captions[key], 4)
        cap.paragraph_format.keep_with_next = True
        docx_list_table(doc, df, widths_cm, fill.lstrip("#"), center_from, bold_last)
        doc.add_paragraph().paragraph_format.space_after = Pt(2)

    # judul
    title = doc.add_paragraph()
    run = title.add_run(T["title"])
    run.bold = True
    run.font.size = Pt(20)
    run.font.color.rgb = navy
    title.paragraph_format.space_after = Pt(2)
    small_line(T["institute"])
    last = small_line(T["period"] % period_head)
    if meta["source_note"]:
        for line in meta["source_note"].split("\n"):
            last = small_line(source_note_text(line, lang))
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

    doc.add_heading(titles["s1"], level=1)
    for p in content["summary"]:
        body(p)
    figure("prodi_jenis")

    sem = content.get("semester")
    if sem:
        doc.add_heading(titles["s_sem"], level=1)
        body(sem["paras"][0])
        compare = sem["compare"]
        n_val = len(compare.columns) - 1
        compare_widths = [6.4]
        for c in range(n_val):
            compare_widths.append(11.0 / n_val)
        table_with_caption("t_sem", compare, compare_widths, NAVY, center_from=1, bold_last=True)
        body(sem["paras"][1])
        figure("sem_prodi")
        figure("sem_jenis")
        body(sem["paras"][2])
        body(sem["paras"][3])
        if len(sem["cross"]):
            table_with_caption("t_cross", sem["cross"], fitted_cm(sem["cross"]), NAVY)

    doc.add_heading(titles["s2"], level=1)
    body(content["prodi_intro"])
    figure("prodi_mitra")
    for i, (name, paras) in enumerate(content["prodi"], start=1):
        doc.add_heading("%d.%d %s" % (sec_no["s2"], i, name), level=2)
        for p in paras:
            body(p)

    doc.add_heading(titles["s3"], level=1)
    body(content["jenis_intro"])
    figure("jenis_status")
    figure("jenis_mitra")
    for i, (name, paras) in enumerate(content["jenis"], start=1):
        doc.add_heading("%d.%d %s" % (sec_no["s3"], i, name), level=2)
        for p in paras:
            body(p)

    doc.add_heading(titles["s_pending"], level=1)
    for p in content["pending"]:
        body(p)
    for key, cap_key, fill in (("pending_df", "t_pending", NAVY), ("excluded_df", "t_excluded", RED)):
        df = content[key]
        if len(df):
            table_with_caption(cap_key, df, fitted_cm(df), fill)

    doc.add_heading(titles["s4"], level=1)
    for p in content["quality"]:
        body(p)
    if content["actions"]:
        lead = doc.add_paragraph(T["actions"])
        lead.paragraph_format.keep_with_next = True
        for a in content["actions"]:
            doc.add_paragraph(a, style="List Bullet")

    doc.add_heading(titles["s5"], level=1)
    body(content["method"])

    # lampiran: matriks
    matrix = content["matrix"]
    doc.add_heading(titles["appendix"], level=1)
    note = caption(T["appendix_note"], 6)
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
    run = foot.add_run(T["footer"] % period_foot + "\t" + T["page"])
    run.font.size = Pt(8)
    run.font.color.rgb = mute
    docx_page_field(foot)
    for run in foot.runs:
        run.font.size = Pt(8)
        run.font.color.rgb = mute

    doc.core_properties.title = "%s - %s" % (T["title"], period_foot)
    doc.core_properties.language = T["doc_lang"]
    zoom = doc.settings.element.find(qn("w:zoom"))
    if zoom is not None and zoom.get(qn("w:percent")) is None:
        zoom.set(qn("w:percent"), "100")
    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


# ================================================================ UI
st.set_page_config(page_title="Analisis MBKM", page_icon="🎓", layout="wide")
st.title("Analisis Aktivitas MBKM & Mata Kuliah Konversi")
# st.caption("Unggah file ekspor SIAKAD (.xls / .xlsx). Data dinormalkan ke 1 baris per aktivitas, "
#            "lalu diberi flag kualitas data.")
st.caption("Supported by Claude Opus 4.8 (Anthropic PBC, San Francisco, California, U.S.) ")

UPLOAD_TYPES = ["xls", "xlsx", "html", "htm"]
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
XLSX_MIME = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
PENDING_TAB = {"id": "Mahasiswa Status Selain Selesai", "en": "Students with a Status Other than Completed"}

st.markdown("**File ekspor SIAKAD** - unggah file Semester Ganjil, Semester Genap, atau keduanya "
            "(laporan satu tahun akademik dengan perbandingan semester).")
u1, u2 = st.columns(2)
up_odd = u1.file_uploader("Semester Ganjil (odd)", type=UPLOAD_TYPES, key="up_odd")
up_even = u2.file_uploader("Semester Genap (even)", type=UPLOAD_TYPES, key="up_even")
uploads = []
if up_odd is not None:
    uploads.append(("Ganjil", up_odd))
if up_even is not None:
    uploads.append(("Genap", up_even))
if not uploads:
    st.info("Silakan unggah file laporan untuk memulai.")
    st.stop()

files = []
for slot, up in uploads:
    try:
        f_raw, f_act, f_note, f_split = load(up.getvalue(), up.name)
    except Exception as exc:  # noqa
        st.error("Gagal membaca file Semester %s (%s): %s" % (slot, up.name, exc))
        st.stop()
    files.append({"slot": slot, "name": up.name, "raw": f_raw, "act": f_act, "note": f_note,
                  "periods": sorted_periods(f_act["Periode Akademik"].unique())})

owner = {}
for f in files:
    for p in f["periods"]:
        if p in owner:
            st.error("Periode %s ada di kedua file (%s dan %s). Unggah satu file untuk Semester Ganjil dan satu "
                     "file untuk Semester Genap." % (p, owner[p], f["name"]))
            st.stop()
        owner[p] = f["name"]
    terms = []
    for p in f["periods"]:
        terms.append(period_parts(p)[1].lower())
    if f["periods"] and f["slot"].lower() not in terms:
        st.warning("File pada kolom Semester %s (%s) berisi data %s. Aplikasi memakai periode yang tercatat di "
                   "dalam file." % (f["slot"], f["name"], join(f["periods"], "id")))
if len(files) > 1 and not academic_year(sorted_periods(owner.keys())):
    st.warning("Kedua file berasal dari tahun akademik yang berbeda (%s)." % join(sorted_periods(owner.keys()), "id"))
files.sort(key=lambda f: period_sort_key(f["periods"][0]) if f["periods"] else (9999, 9, ""))

raw_parts = []
act_parts = []
notes = []
for k, f in enumerate(files):
    part = f["act"].copy()
    aids = []
    for v in part["_i"]:
        aids.append("%d-%d" % (k, v))
    part["_aid"] = aids
    part["_file"] = k
    act_parts.append(part)
    raw_parts.append(f["raw"])
    if f["note"]:
        if len(files) > 1:
            notes.append("%s - %s" % (join(f["periods"], "id"), f["note"]))
        else:
            notes.append(f["note"])
raw_files = pd.concat(raw_parts, ignore_index=True)
act_files = pd.concat(act_parts, ignore_index=True)
order = sorted(range(len(act_files)), key=lambda i: (period_sort_key(act_files.at[i, "Periode Akademik"]),
                                                     act_files.at[i, "_file"], act_files.at[i, "_i"]))
act_files = act_files.iloc[order].reset_index(drop=True)
act_files["No"] = range(1, len(act_files) + 1)
act_files = mark_superseded(act_files)
source_note = "\n".join(notes)
periods_all = sorted_periods(act_files["Periode Akademik"].unique())

with st.sidebar:
    st.header("🔍 Cari mahasiswa")
    query = st.text_input("Nama atau NIM (boleh sebagian)", placeholder="contoh: ellen atau 22010173",
                          key="search_q", help="Mencari di semua record file, termasuk yang dibuang filter status.")
    st.divider()
    st.header("Filter")
    sem_choice = ""
    if len(periods_all) > 1:
        ay_all = academic_year(periods_all)
        all_label = ("Gabungan (Tahun Akademik %s)" % ay_all) if ay_all else "Gabungan (semua periode)"
        sem_pick = st.radio("Semester", [all_label] + periods_all, key="sem_pick",
                            help="Gabungan = laporan satu tahun akademik dengan perbandingan semester.")
        if sem_pick != all_label:
            sem_choice = sem_pick
    if sem_choice:
        act_all = act_files[act_files["Periode Akademik"] == sem_choice].reset_index(drop=True)
        raw = raw_files[raw_files["Periode Akademik"] == sem_choice].reset_index(drop=True)
    else:
        act_all = act_files
        raw = raw_files
    all_status = sorted(act_all["Status Aktivitas"].unique())
    default_drop = []
    for s in all_status:
        if s != "Selesai":
            default_drop.append(s)
    if "Selesai" not in all_status:
        default_drop = []
        for s in all_status:
            if s in ["Ditolak", "Dibatalkan"] or is_dup_status(s):
                default_drop.append(s)
    drop_status = st.multiselect("Buang status aktivitas", all_status, default=default_drop,
                                 help="Bawaan: hanya aktivitas Selesai yang dihitung. Evaluasi/Diajukan tetap "
                                      "dilaporkan sebagai 'belum selesai' tetapi tidak dihitung; hapus dari daftar ini "
                                      "untuk ikut menghitungnya. '(duplikat)' = pengajuan Diajukan/Evaluasi yang "
                                      "mengonversi MK yang sama dengan aktivitas Selesai milik mahasiswa yang sama "
                                      "pada semester yang sama.")
    mk_overload = st.number_input("Ambang MK berlebih (F2)", min_value=1, max_value=50, value=5)
    mk_overload_px = st.number_input("Ambang MK berlebih - Pertukaran Pelajar (F2)",
                                     min_value=1, max_value=60, value=12)
    cohort_text = st.text_input("Hitung mahasiswa per awal NIM", value=", ".join(DEFAULT_COHORTS),
                                help="Contoh: 22, 23. Dipakai di tab Ringkasan dan di laporan.")
    cohorts = parse_cohorts(cohort_text)
    lang_pick = st.radio("Bahasa laporan naratif", list(LANGS.keys()), horizontal=True)
    lang = LANGS[lang_pick]
    n_split = count_split_records(raw)
    if len(files) > 1:
        sizes = []
        for f in files:
            sizes.append("%d baris (%s)" % (len(f["raw"]), join(f["periods"], "id")))
        st.caption("File asli: %s; 1 baris per aktivitas × MK konversi." % " + ".join(sizes))
    else:
        st.caption("File asli: %d baris (1 baris per aktivitas × MK konversi)." % len(raw))
    if n_split:
        st.caption("%d record terpecah (urutan nama dosen berbeda) sudah digabung." % n_split)
    n_dup = 0
    for s in act_all["Status Aktivitas"]:
        if is_dup_status(s):
            n_dup += 1
    if n_dup:
        st.caption("%d pengajuan ganda (MK sama dengan aktivitas Selesai di semester yang sama) diberi status "
                   "'(duplikat)'." % n_dup)
    for line in notes:
        st.caption(line)

act = act_all[~act_all["Status Aktivitas"].isin(drop_status)].reset_index(drop=True)
act["No"] = range(1, len(act) + 1)
uncounted = uncounted_records(act_all, drop_status)
if len(act) == 0:
    st.warning("Tidak ada aktivitas tersisa setelah filter status.")
    st.stop()
periods = sorted_periods(act["Periode Akademik"].unique())
periods_sel = sorted_periods(act_all["Periode Akademik"].unique())
multi_sem = len(periods) > 1
flags = flag_table(act, mk_overload, mk_overload_px)
flag_names = list(flags.columns)
defs = flag_defs(mk_overload, mk_overload_px, multi_sem)
nim_info = nim_check(act)
periode = periode_label(act_all)
stem = period_stem(act_all)

n_act = len(act)
n_stu = act["NIM"].nunique()
n_dropped = len(act_all) - len(act)
n_issue = int(pd.concat([flags[k] for k in RECORD_FLAGS], axis=1).any(axis=1).sum())
meta = {"source_note": source_note, "drop_status": drop_status, "n_dropped": n_dropped,
        "mk_overload": mk_overload, "mk_overload_px": mk_overload_px, "n_split": n_split,
        "periode": periode, "act_all": act_all, "periods": periods_sel, "cohorts": cohorts}

cats = partner_categories(act, flags)
narr = narrative_content(act, raw, flags, nim_info, meta, cats, lang)
charts = narrative_charts(act, cats, lang, uncounted)
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

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Aktivitas", n_act)
# m2.metric("Mahasiswa", n_stu)
# m3.metric("Baris asli", len(raw))
# m4.metric("Dikecualikan", n_dropped)
# m5.metric("Isu per-record", n_issue)

if query.strip():
    try:
        search_box = st.container(border=True)
    except TypeError:  # Streamlit lama tanpa opsi border
        search_box = st.container()
    with search_box:
        st.markdown("#### 🔍 Hasil pencarian: \"%s\"" % md_escape(query.strip()))
        q = query.strip().lower()
        by_name = act_files["Nama"].str.lower().str.contains(q, regex=False)
        by_nim = act_files["NIM"].str.contains(re.sub(r"\s+", "", q), regex=False)
        hits = act_files[by_name | by_nim]
        if len(hits) == 0:
            st.info("Tidak ada mahasiswa dengan nama atau NIM yang mengandung \"%s\"." % query.strip())
        else:
            people = hits.drop_duplicates("NIM")
            options = []
            for _, r in people.iterrows():
                options.append("%s - %s (%s)" % (r["NIM"], r["Nama"], prodi_label(r["Program Studi"])))
            st.caption("%d mahasiswa ditemukan." % len(options))
            pick = options[0]
            if len(options) > 1:
                pick = st.selectbox("Pilih mahasiswa", options, key="search_pick")
            nim_pick = pick.split(" - ")[0]
            rec = act_files[act_files["NIM"] == nim_pick]
            kept = act[act["NIM"] == nim_pick]
            first = rec.iloc[0]
            st.markdown("### %s" % md_escape(first["Nama"]))
            st.caption("NIM %s · %s · Periode %s" % (nim_pick, prodi_label(first["Program Studi"]),
                                                    join(sorted_periods(rec["Periode Akademik"].unique()), "id")))
            s1, s2, s3 = st.columns(3)
            s1.metric("Aktivitas di file", len(rec))
            s2.metric("Masuk analisis", len(kept))
            s3.metric("Total SKS (dianalisis)", num(kept["Total SKS"].sum(), "id"))
            notes_by_aid = {}
            for i in kept.index:
                found = []
                for k in flag_names:
                    if flags.at[i, k]:
                        found.append(k)
                notes_by_aid[kept.at[i, "_aid"]] = found
            rows = []
            for _, r in rec.iterrows():
                if r["_aid"] in notes_by_aid:
                    included = "Ya"
                    found = notes_by_aid[r["_aid"]]
                    note = ", ".join(found) if found else "-"
                elif r["Periode Akademik"] not in periods_sel:
                    included = "Tidak (semester lain)"
                    note = "-"
                elif r["Status Aktivitas"] in IN_PROGRESS:
                    included = "Tidak (belum selesai)"
                    note = "-"
                else:
                    included = "Tidak (dibuang filter)"
                    note = "-"
                rows.append({
                    "Semester": r["Periode Akademik"],
                    "Jenis Aktivitas": short_jenis(r["Jenis Aktivitas"]), "Status": r["Status Aktivitas"],
                    "Masuk analisis": included, "Mitra": r["Mitra"], "Status Mitra": r["Status Mitra"],
                    "Tanggal": "%s – %s" % (r["Tanggal Mulai"], r["Tanggal Selesai"]),
                    "Jml MK": r["Jml MK"], "Total SKS": r["Total SKS"], "MK Konversi": r["MK Konversi"],
                    "Dosen Pembimbing": r["Dosen Pembimbing"], "Dosen Penguji": r["Dosen Penguji"],
                    "Judul Aktivitas": r["Judul Aktivitas"], "Catatan (flag)": note,
                })
            st.dataframe(pd.DataFrame(rows), use_container_width=True, hide_index=True)
            in_sel = rec[rec["Periode Akademik"].isin(periods_sel)]
            if len(kept) == 0 and len(in_sel) == 0:
                st.info("Mahasiswa ini tidak tercatat pada semester yang dipilih (%s)." % join(periods_sel, "id"))
            elif len(kept) == 0 and in_sel["Status Aktivitas"].isin(IN_PROGRESS).any():
                st.warning("Aktivitas mahasiswa ini belum berstatus Selesai, sehingga ia belum dihitung dalam jumlah "
                           "mahasiswa. Aktivitasnya tercantum di tab Mahasiswa Status Selain Selesai.")
            elif len(kept) == 0:
                st.warning("Semua record mahasiswa ini dibuang oleh filter status, sehingga ia tidak dihitung "
                           "dalam jumlah mahasiswa.")
            else:
                if kept.duplicated("Periode Akademik").any():
                    if multi_sem:
                        st.warning("Mahasiswa ini tercatat pada lebih dari satu aktivitas dalam semester yang sama "
                                   "(F7) - periksa kemungkinan duplikasi.")
                    else:
                        st.warning("Mahasiswa ini tercatat pada lebih dari satu aktivitas (F7) - periksa "
                                   "kemungkinan duplikasi.")
                if kept["Periode Akademik"].nunique() > 1:
                    if F8 in flag_names and bool(flags.loc[kept.index, F8].any()):
                        st.warning("Mahasiswa ini mengonversi MK yang sama pada lebih dari satu semester (F8) - "
                                   "periksa kemungkinan konversi ganda.")
                    else:
                        st.info("Mahasiswa ini mengikuti MBKM di lebih dari satu semester.")


tab_names = ["Ringkasan"]
if multi_sem:
    tab_names.append("Per Semester")
tab_names += ["Per Jenis", "Per Program Studi", "Flag", "Cek NIM", PENDING_TAB[lang], "Matriks", "Narasi", "Unduh"]
tabs = {}
for name, obj in zip(tab_names, st.tabs(tab_names)):
    tabs[name] = obj

with tabs["Ringkasan"]:
    if multi_sem:
        st.caption("Data gabungan %s. Rincian per semester ada di tab Per Semester." % join(periods, "id"))
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Status Aktivitas")
        st.pyplot(barh(with_uncounted(act, uncounted)["Status Aktivitas"].value_counts(), NAVY, "Status aktivitas"))
        if unc_len(uncounted):
            st.caption("Termasuk %d aktivitas belum selesai (Evaluasi/Diajukan) yang tidak dihitung pada angka lain."
                       % unc_len(uncounted))
    with c2:
        st.subheader("Flag")
        fc = {}
        for k in flag_names:
            fc[k] = int(flags[k].sum())
        st.pyplot(barh(pd.Series(fc), RED, "Jumlah aktivitas ter-flag"))
    if cohorts:
        st.subheader("Mahasiswa per awal NIM")
        st.dataframe(cohort_table(act, cohorts, periods), use_container_width=False, hide_index=True)
        st.caption("NIM unik setelah filter status.%s" % (" Kolom Total menghitung NIM unik di seluruh semester."
                                                          if multi_sem else ""))

if multi_sem:
    with tabs["Per Semester"]:
        st.subheader("Perbandingan semester")
        compare = semester_compare_table(act, cats, flags, "id", uncounted)
        st.dataframe(compare, use_container_width=True, hide_index=True, height=(len(compare) + 1) * 35 + 3)
        st.download_button("⬇️ Unduh perbandingan semester (CSV)", compare.to_csv(index=False).encode("utf-8"),
                           file_name="%s_perbandingan_semester.csv" % stem, mime="text/csv", key="dl_sem_compare")
        charts_id = narrative_charts(act, cats, "id", uncounted)
        st.image(charts_id["sem_prodi"],
                 caption="Jumlah mahasiswa per program studi pada setiap semester (NIM unik dalam semester).")
        st.image(charts_id["sem_jenis"],
                 caption="Jumlah mahasiswa per jenis aktivitas pada setiap semester (NIM unik dalam semester).")
        st.markdown("**Rekap per program studi dan semester**")
        st.dataframe(rekap_sem_display(act, "Program Studi", periods, prodi_label), use_container_width=True,
                     hide_index=True)
        st.caption("Kolom Total menghitung mahasiswa sebagai NIM unik di seluruh semester.")
        st.divider()
        cross = nim_info["cross"]
        st.markdown("**Mahasiswa yang mengikuti MBKM di lebih dari satu semester**")
        if len(cross) == 0:
            st.success("Tidak ada mahasiswa yang tercatat di lebih dari satu semester.")
        else:
            n_same = int((cross[SAME_MK_COL["id"]] != "-").sum())
            if n_same:
                st.warning("%d mahasiswa tercatat di lebih dari satu semester; %d di antaranya mengonversi MK yang "
                           "sama (F8) - periksa kemungkinan konversi ganda." % (len(cross), n_same))
            else:
                st.info("%d mahasiswa tercatat di lebih dari satu semester dengan MK yang berbeda." % len(cross))
            st.dataframe(cross, use_container_width=True, hide_index=True)

with tabs["Per Jenis"]:
    st.subheader("Rekap mahasiswa per Jenis Aktivitas")
    stu_j = act.groupby("Jenis Aktivitas")["NIM"].nunique().sort_values(ascending=False)
    labels = []
    for j in stu_j.index:
        labels.append(short_jenis(j))
    stu_j.index = labels
    st.pyplot(barh(stu_j, NAVY, "Mahasiswa per Jenis Aktivitas"))
    if multi_sem:
        st.markdown("**Per semester**")
        st.dataframe(rekap_sem_display(act, "Jenis Aktivitas", periods, short_jenis), use_container_width=True,
                     hide_index=True)
    jenis_pick = st.selectbox("Lihat daftar mahasiswa untuk Jenis Aktivitas",
                              act["Jenis Aktivitas"].value_counts().index)
    sub = act[act["Jenis Aktivitas"] == jenis_pick]
    st.caption("%d aktivitas, %d mahasiswa" % (len(sub), sub["NIM"].nunique()))
    show = ["NIM", "Nama", "Program Studi", "Status Aktivitas", "Mitra", "Judul Aktivitas"]
    if multi_sem:
        show.insert(2, "Periode Akademik")
    st.dataframe(sub[show].reset_index(drop=True), use_container_width=True)

with tabs["Per Program Studi"]:
    st.subheader("Rekap mahasiswa per Program Studi")
    stu_p = act.groupby("Program Studi")["NIM"].nunique().sort_values(ascending=False)

    # --- tabel rekap + unduh ---
    if multi_sem:
        rekap_p = rekap_sem_display(act, "Program Studi", periods, prodi_label)
    else:
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
    if multi_sem:
        st.caption("Kolom Total menghitung mahasiswa sebagai NIM unik di seluruh semester.")
    st.download_button("⬇️ Unduh rekap per Program Studi (CSV)",
                       rekap_p.to_csv(index=False).encode("utf-8"),
                       file_name="rekap_mahasiswa_per_prodi.csv", mime="text/csv",
                       key="dl_rekap_prodi")

    st.divider()

    # --- daftar mahasiswa per prodi (drill-down) + unduh ---
    prodi_pick = st.selectbox("Lihat daftar mahasiswa untuk Program Studi", stu_p.index)
    sub = act[act["Program Studi"] == prodi_pick].reset_index(drop=True)
    show = ["NIM", "Nama", "Jenis Aktivitas", "Status Aktivitas", "Mitra", "Judul Aktivitas"]
    if multi_sem:
        show.insert(2, "Periode Akademik")
    sub_show = sub[show]
    st.caption("%d aktivitas, %d mahasiswa" % (len(sub), sub["NIM"].nunique()))
    st.dataframe(sub_show, use_container_width=True, hide_index=True)

    slug = prodi_label(prodi_pick).replace(" ", "_")
    st.download_button("⬇️ Unduh daftar mahasiswa (CSV)",
                       sub_show.to_csv(index=False).encode("utf-8"),
                       file_name="mahasiswa_%s.csv" % slug, mime="text/csv",
                       key="dl_students_prodi")

with tabs["Flag"]:
    st.subheader("Flag kualitas data")
    grid = st.columns(4)
    for i, k in enumerate(flag_names):
        grid[i % 4].metric(k, int(flags[k].sum()), help=defs[k])
    st.divider()
    flag_pick = st.selectbox("Lihat record untuk flag", flag_names)
    st.caption(defs[flag_pick])
    sub = act[flags[flag_pick]]
    show = ["No", "NIM", "Nama", "Program Studi", "Jenis Aktivitas", "Status Aktivitas",
            "Mitra", "Jml MK", "Total SKS", "Dosen Pembimbing", "Judul Aktivitas"]
    if multi_sem:
        show.insert(1, "Periode Akademik")
        show.insert(9, "MK Konversi")
    st.dataframe(sub[show].reset_index(drop=True), use_container_width=True)

with tabs["Cek NIM"]:
    st.subheader("Pemeriksaan NIM")
    multi = nim_info["multi"]
    cross = nim_info["cross"]
    if multi_sem:
        c1, c2, c3, c4, c5 = st.columns(5)
        c1.metric("Aktivitas", n_act)
        c2.metric("Mahasiswa (NIM unik)", n_stu)
        c3.metric("Selisih", n_act - n_stu)
        c4.metric("NIM >1 aktivitas (semester sama)", int(multi["NIM"].nunique()) if len(multi) else 0)
        c5.metric("Mahasiswa di >1 semester", len(cross))
    else:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Aktivitas", n_act)
        c2.metric("Mahasiswa (NIM unik)", n_stu)
        c3.metric("Selisih", n_act - n_stu)
        c4.metric("NIM >1 aktivitas", len(multi))
    if n_split:
        st.caption("%d record aktivitas yang terpecah karena urutan nama dosen berbeda di ekspor SIAKAD "
                   "sudah digabung otomatis." % n_split)
    if len(multi) == 0:
        if multi_sem:
            st.success("Dalam setiap semester, setiap NIM tercatat tepat 1 aktivitas.")
        else:
            st.success("Setiap NIM tercatat tepat 1 aktivitas — jumlah aktivitas = jumlah mahasiswa.")
    else:
        if multi_sem:
            st.warning("%d NIM tercatat di lebih dari 1 aktivitas dalam semester yang sama. 'MK Sama' = MK yang "
                       "dikonversi di lebih dari 1 aktivitas (indikasi kuat duplikasi / pengajuan ulang)."
                       % multi["NIM"].nunique())
        else:
            st.warning("%d NIM tercatat di lebih dari 1 aktivitas, sehingga jumlah aktivitas (%d) > jumlah "
                       "mahasiswa (%d). 'MK Sama' = MK yang dikonversi di lebih dari 1 aktivitas "
                       "(indikasi kuat duplikasi / pengajuan ulang)." % (len(multi), n_act, n_stu))
        st.dataframe(multi, use_container_width=True, hide_index=True)
    if multi_sem:
        st.markdown("**Mahasiswa yang mengikuti MBKM di lebih dari satu semester**")
        if len(cross) == 0:
            st.success("Tidak ada mahasiswa yang tercatat di lebih dari satu semester.")
        else:
            n_same = int((cross[SAME_MK_COL["id"]] != "-").sum())
            if n_same:
                st.warning("%d mahasiswa tercatat di lebih dari satu semester; %d di antaranya mengonversi MK yang "
                           "sama (F8) - periksa kemungkinan konversi ganda." % (len(cross), n_same))
            else:
                st.info("%d mahasiswa tercatat di lebih dari satu semester dengan MK yang berbeda." % len(cross))
            st.dataframe(cross, use_container_width=True, hide_index=True)
    if len(nim_info["invalid"]):
        st.error("%d aktivitas dengan NIM kosong / bukan angka." % len(nim_info["invalid"]))
        st.dataframe(nim_info["invalid"][["No", "NIM", "Nama", "Program Studi"]],
                     use_container_width=True, hide_index=True)
    if len(nim_info["conflict"]):
        st.error("%d NIM dengan nama / program studi berbeda." % len(nim_info["conflict"]))
        st.dataframe(nim_info["conflict"], use_container_width=True, hide_index=True)

    st.divider()
    nim_query = st.text_input("Cek satu NIM (semua semester dan status, termasuk yang dibuang filter)")
    if nim_query.strip():
        q = clean_nim(nim_query)
        hits = act_files[act_files["NIM"] == q]
        if len(hits) == 0:
            st.info("NIM %s tidak ditemukan di file." % q)
        else:
            st.caption("%s - %s: %d aktivitas, %d baris MK di file asli"
                       % (q, hits.iloc[0]["Nama"], len(hits), int(hits["_rows"].sum())))
            st.dataframe(hits[["Periode Akademik", "Jenis Aktivitas", "Status Aktivitas", "Mitra", "Jml MK",
                               "Total SKS", "MK Konversi", "Dosen Pembimbing", "Dosen Penguji"]],
                         use_container_width=True, hide_index=True)

with tabs[PENDING_TAB[lang]]:
    en = lang == "en"
    caps = narr["captions"]
    pending_df = narr["pending_df"]
    excluded_df = narr["excluded_df"]
    st.subheader(PENDING_TAB[lang])
    raw_pending = with_uncounted(act[act["Status Aktivitas"] != "Selesai"], uncounted)
    p1, p2, p3, p4 = st.columns(4)
    p1.metric("Activities" if en else "Aktivitas", len(raw_pending))
    p2.metric("Students" if en else "Mahasiswa", raw_pending["NIM"].nunique())
    p3.metric(status_name("Evaluasi", lang), int((raw_pending["Status Aktivitas"] == "Evaluasi").sum()))
    p4.metric(status_name("Diajukan", lang), int((raw_pending["Status Aktivitas"] == "Diajukan").sum()))
    for p in narr["pending"]:
        st.markdown(md_escape(p))
    st.divider()
    if len(pending_df):
        st.markdown("**" + caps.get("t_pending", "") + "**")
        off = 1 if pending_df.columns[0] == "Semester" else 0
        prodi_col = pending_df.columns[2 + off]
        status_col = pending_df.columns[4 + off]
        if off:
            f0, f1, f2 = st.columns(3)
            sem_sel = f0.multiselect("Semester", list(pending_df["Semester"].unique()), key="pend_sem")
        else:
            f1, f2 = st.columns(2)
            sem_sel = []
        prodi_sel = f1.multiselect(prodi_col, sorted(pending_df[prodi_col].unique()), key="pend_prodi")
        status_sel = f2.multiselect(status_col, list(pending_df[status_col].unique()), key="pend_status")
        shown = pending_df
        if sem_sel:
            shown = shown[shown["Semester"].isin(sem_sel)]
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
        st.markdown("**" + caps.get("t_excluded", "") + "**")
        st.dataframe(excluded_df, use_container_width=True, hide_index=True)
        st.download_button("⬇️ " + ("Download excluded records (CSV)" if en else "Unduh record dikecualikan (CSV)"),
                           excluded_df.to_csv(index=False).encode("utf-8"),
                           file_name="%s_%s.csv" % (stem, "excluded" if en else "dikecualikan"),
                           mime="text/csv", key="dl_excluded")

with tabs["Matriks"]:
    st.subheader("Matriks mahasiswa: Program Studi × Jenis Aktivitas")
    st.dataframe(student_matrix(act), use_container_width=True)
    st.caption("Angka = jumlah mahasiswa (NIM unik). Mahasiswa dengan aktivitas di lebih dari 1 jenis "
               "dihitung di tiap jenis, tetapi sekali di kolom/baris Total.")
    if multi_sem:
        for p in periods:
            st.markdown("**%s**" % p)
            st.dataframe(student_matrix(act[act["Periode Akademik"] == p]), use_container_width=True)

with tabs["Narasi"]:
    T = NARR_TEXT[lang]
    titles = narr["titles"]
    caps = narr["captions"]
    sec_no = narr["sec_no"]
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
    st.markdown("#### " + titles["s1"])
    for p in narr["summary"]:
        st.markdown(md_escape(p))
    st.image(charts["prodi_jenis"], caption=caps["prodi_jenis"])
    sem = narr["semester"]
    if sem:
        st.markdown("#### " + titles["s_sem"])
        st.markdown(md_escape(sem["paras"][0]))
        st.caption(caps["t_sem"])
        st.dataframe(sem["compare"], use_container_width=True, hide_index=True)
        st.markdown(md_escape(sem["paras"][1]))
        st.image(charts["sem_prodi"], caption=caps["sem_prodi"])
        st.image(charts["sem_jenis"], caption=caps["sem_jenis"])
        st.markdown(md_escape(sem["paras"][2]))
        st.markdown(md_escape(sem["paras"][3]))
        if len(sem["cross"]):
            st.caption(caps["t_cross"])
            st.dataframe(sem["cross"], use_container_width=True, hide_index=True)
    st.markdown("#### " + titles["s2"])
    st.markdown(md_escape(narr["prodi_intro"]))
    st.image(charts["prodi_mitra"], caption=caps["prodi_mitra"])
    for i, (name, paras) in enumerate(narr["prodi"], start=1):
        st.markdown("**%d.%d %s**" % (sec_no["s2"], i, md_escape(name)))
        for p in paras:
            st.markdown(md_escape(p))
    st.markdown("#### " + titles["s3"])
    st.markdown(md_escape(narr["jenis_intro"]))
    st.image(charts["jenis_status"], caption=caps["jenis_status"])
    st.image(charts["jenis_mitra"], caption=caps["jenis_mitra"])
    for i, (name, paras) in enumerate(narr["jenis"], start=1):
        st.markdown("**%d.%d %s**" % (sec_no["s3"], i, md_escape(name)))
        for p in paras:
            st.markdown(md_escape(p))
    st.markdown("#### " + titles["s_pending"])
    for p in narr["pending"]:
        st.markdown(md_escape(p))
    if len(narr["pending_df"]):
        st.caption(caps["t_pending"])
        st.dataframe(narr["pending_df"], use_container_width=True, hide_index=True)
    if len(narr["excluded_df"]):
        st.caption(caps["t_excluded"])
        st.dataframe(narr["excluded_df"], use_container_width=True, hide_index=True)
    st.markdown("#### " + titles["s4"])
    for p in narr["quality"]:
        st.markdown(md_escape(p))
    if narr["actions"]:
        lines = [T["actions"]]
        for a in narr["actions"]:
            lines.append("- " + md_escape(a))
        st.markdown("\n".join(lines))
    st.markdown("#### " + titles["s5"])
    st.markdown(md_escape(narr["method"]))

with tabs["Unduh"]:
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
                       file_name=stem + "_bersih.xlsx", mime=XLSX_MIME)
    st.download_button("📥 Unduh Aktivitas (CSV)", build_csv(act, flags),
                       file_name=stem + "_aktivitas.csv", mime="text/csv")
