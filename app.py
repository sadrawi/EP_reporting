"""
Analisis Aktivitas MBKM & Mata Kuliah Konversi - aplikasi Streamlit.

Jalankan:
    pip install streamlit pandas beautifulsoup4 lxml openpyxl xlrd reportlab matplotlib
    streamlit run app.py

Unggah file ekspor SIAKAD (.xls yang sebenarnya HTML, atau .xlsx). Aplikasi
menormalkan data (1 baris per aktivitas), memberi flag kualitas data, memeriksa
NIM, dan menyediakan rekap per Jenis Aktivitas / Program Studi serta unduhan
PDF, Excel & CSV.

Normalisasi: ekspor SIAKAD berisi 1 baris per (aktivitas x MK konversi). Baris
dikelompokkan per NIM + atribut aktivitas. Kolom dosen TIDAK dipakai sebagai
kunci karena urutan nama dosen bisa berbeda antar baris MK pada aktivitas yang
sama; daftar dosen digabung (unik, urut ID dosen) per aktivitas.
"""

import io
import re
import pandas as pd
import matplotlib.pyplot as plt
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


def student_matrix(act):
    """Matriks NIM unik Program Studi x Jenis; kolom/baris Total juga NIM unik."""
    piv = act.groupby(["Program Studi", "Jenis Aktivitas"])["NIM"].nunique().unstack(fill_value=0)
    piv["Total"] = act.groupby("Program Studi")["NIM"].nunique()
    total_row = act.groupby("Jenis Aktivitas")["NIM"].nunique()
    total_row["Total"] = act["NIM"].nunique()
    piv.loc["TOTAL"] = total_row
    piv = piv.fillna(0).astype(int)
    cols = []
    for c in piv.columns:
        cols.append(short_jenis(c))
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


def build_excel(act, flags, nim_info):
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
    buf.seek(0)
    return buf.getvalue()


def build_csv(act, flags):
    cols = ["No"] + KEYS + ["Jml MK", "Total SKS", "MK Konversi"]
    return export_table(act, flags, cols).to_csv(index=False).encode("utf-8")


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

    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, leftMargin=18 * mm, rightMargin=18 * mm,
                            topMargin=16 * mm, bottomMargin=15 * mm, title="Laporan Analisis MBKM")
    doc.build(story)
    buf.seek(0)
    return buf.getvalue()


# ================================================================ UI
st.set_page_config(page_title="Analisis MBKM", page_icon="🎓", layout="wide")
st.title("Analisis Aktivitas MBKM & Mata Kuliah Konversi")
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

n_act = len(act)
n_stu = act["NIM"].nunique()
n_dropped = len(act_all) - len(act)
n_issue = int(pd.concat([flags[k] for k in RECORD_FLAGS], axis=1).any(axis=1).sum())

m1, m2, m3, m4, m5 = st.columns(5)
m1.metric("Aktivitas", n_act)
# m2.metric("Mahasiswa", n_stu)
# m3.metric("Baris asli", len(raw))
# m4.metric("Dikecualikan", n_dropped)
# m5.metric("Isu per-record", n_issue)

tab_sum, tab_jenis, tab_prodi, tab_flag, tab_nim, tab_matrix, tab_dl = st.tabs(
    ["Ringkasan", "Per Jenis", "Per Program Studi", "Flag", "Cek NIM", "Matriks", "Unduh"])

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

with tab_matrix:
    st.subheader("Matriks mahasiswa: Program Studi × Jenis Aktivitas")
    st.dataframe(student_matrix(act), use_container_width=True)
    st.caption("Angka = jumlah mahasiswa (NIM unik). Mahasiswa dengan aktivitas di lebih dari 1 jenis "
               "dihitung di tiap jenis, tetapi sekali di kolom/baris Total.")

with tab_dl:
    st.subheader("Unduh hasil")
    stem = period_stem(act_all)
    meta = {"source_note": source_note, "drop_status": drop_status, "n_dropped": n_dropped,
            "mk_overload": mk_overload, "mk_overload_px": mk_overload_px, "n_split": n_split,
            "periode": periode}
    try:
        pdf_bytes = build_pdf(act, raw, flags, nim_info, meta)
        st.download_button("📄 Unduh laporan PDF", pdf_bytes, file_name=stem + "_laporan.pdf",
                           mime="application/pdf")
    except Exception as exc:  # noqa
        st.warning("PDF gagal dibuat (butuh paket reportlab): %s" % exc)
    xlsx_bytes = build_excel(act, flags, nim_info)
    st.download_button("📊 Unduh data Excel (Aktivitas + rekap + cek NIM)", xlsx_bytes,
                       file_name=stem + "_bersih.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("📥 Unduh Aktivitas (CSV)", build_csv(act, flags),
                       file_name=stem + "_aktivitas.csv", mime="text/csv")
