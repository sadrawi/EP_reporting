"""
Analisis Aktivitas MBKM & Mata Kuliah Konversi - aplikasi Streamlit.

Jalankan:
    pip install streamlit pandas beautifulsoup4 lxml openpyxl reportlab matplotlib
    streamlit run app.py

Unggah file ekspor SIAKAD (.xls yang sebenarnya HTML, atau .xlsx). Aplikasi
menormalkan data (1 baris per aktivitas), memberi flag kualitas data, dan
menyediakan rekap per Jenis Aktivitas / Program Studi serta unduhan PDF & Excel.
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
}
FLAG_COLS = [
    "F1 Tanpa MK", "F2 MK Berlebih", "F3 Tanpa Pembimbing",
    "F4 Tanpa Penguji", "F5 Mitra Tidak Valid", "F6 Tanpa MoU",
]
RECORD_FLAGS = ["F1 Tanpa MK", "F2 MK Berlebih", "F3 Tanpa Pembimbing", "F5 Mitra Tidak Valid"]

KEYS = [
    "NIM", "Nama", "Program Studi", "Periode Akademik", "Jenis Aktivitas",
    "Mitra", "Status Mitra", "Status Aktivitas", "Tanggal Mulai",
    "Tanggal Selesai", "Posisi", "Judul Aktivitas", "Dosen Pembimbing",
    "Dosen Penguji",
]
NAVY = "#1F3864"
RED = "#7B2D26"


def flag_defs(mk_overload):
    return {
        "F1 Tanpa MK": "Aktivitas tanpa satu pun mata kuliah konversi (Jml MK = 0).",
        "F2 MK Berlebih": "Jml MK > %d - diduga salah pilih / seluruh katalog terpilih." % mk_overload,
        "F3 Tanpa Pembimbing": "Kolom Dosen Pembimbing kosong atau '-'.",
        "F4 Tanpa Penguji": "Kolom Dosen Penguji kosong atau '-' (seluruh record).",
        "F5 Mitra Tidak Valid": "Mitra kosong/'-' atau berisi kode MK / teks bebas, bukan nama mitra.",
        "F6 Tanpa MoU": "Mitra berstatus 'Belum memiliki MoU kerjasama'.",
    }


# ---------------------------------------------------------------- parsing
def squash(text):
    return " ".join(str(text).split())


def sniff_format(data):
    if data[:8] == b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1":
        return "biff"
    if data[:2] == b"PK":
        return "xlsx"
    low = data[:2048].lower()
    if b"<html" in low or b"<!doctype html" in low or b"<table" in low:
        return "html"
    return "unknown"


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
        cells = [squash(c.get_text(" ", strip=True)) for c in tr.find_all(["td", "th"])]
        if "NIM" in cells and "Nama" in cells:
            header = cells
            start = i + 1
            break
    if header is None:
        raise ValueError("Baris header tidak ditemukan")

    data = []
    for tr in trs[start:]:
        cells = [squash(c.get_text(" ", strip=True)) for c in tr.find_all(["td", "th"])]
        if len(cells) == len(header):
            data.append(cells)

    df = pd.DataFrame(data, columns=header)
    df = df.loc[:, ~df.columns.duplicated()]
    keep = []
    seen_no = False
    for name in df.columns:
        if re.fullmatch(r"No\.?", name):
            if seen_no:
                continue
            seen_no = True
        keep.append(name)
    df = df[keep]
    return df.fillna("").astype(str).apply(lambda s: s.str.strip())


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


def normalize_activities(df):
    work = df.copy()
    work["_i"] = range(len(work))
    records = []
    for _, group in work.groupby(KEYS, sort=False, dropna=False):
        record = {}
        for key in KEYS:
            record[key] = group.iloc[0][key]
        codes = list(group["Kode MK Konversi"])
        names = list(group["Nama MK Konversi"])
        record["Jml MK"] = sum(1 for c in codes if str(c).strip())
        record["Total SKS"] = float(pd.to_numeric(group["SKS"], errors="coerce").fillna(0).sum())
        record["MK Konversi"] = join_mk(codes, names)
        record["_i"] = int(group["_i"].min())
        records.append(record)
    act = pd.DataFrame(records).sort_values("_i").reset_index(drop=True)
    act["No"] = range(1, len(act) + 1)
    return act


@st.cache_data(show_spinner=False)
def load(data, name):
    kind = sniff_format(data)
    if kind == "html":
        text = data.decode("utf-8", errors="replace")
        raw = read_html_report(text)
        note = read_printed_by(text)
    elif kind == "biff":
        raw = pd.read_excel(io.BytesIO(data), engine="xlrd", dtype=str).fillna("")
        note = ""
    elif kind == "xlsx":
        raw = pd.read_excel(io.BytesIO(data), engine="openpyxl", dtype=str).fillna("")
        note = ""
    else:
        raise ValueError("Format file tidak dikenali: " + kind)
    act = normalize_activities(raw)
    return raw, act, note


def flag_table(act, mk_overload):
    mitra = act["Mitra"].str.strip()
    bimb = act["Dosen Pembimbing"].str.strip()
    uji = act["Dosen Penguji"].str.strip()
    out = pd.DataFrame(index=act.index)
    out["F1 Tanpa MK"] = act["Jml MK"] == 0
    out["F2 MK Berlebih"] = act["Jml MK"] > mk_overload
    out["F3 Tanpa Pembimbing"] = bimb.isin(["-", ""])
    out["F4 Tanpa Penguji"] = uji.isin(["-", ""])
    out["F5 Mitra Tidak Valid"] = mitra.isin(["-", ""]) | mitra.str.contains("internship", case=False, regex=False)
    out["F6 Tanpa MoU"] = act["Status Mitra"] == "Belum memiliki MoU kerjasama"
    return out


def prodi_label(name):
    return name.replace("S1 - ", "").strip()


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


# ---------------------------------------------------------------- Excel & PDF
def build_excel(act, flags):
    buf = io.BytesIO()
    table = act.copy()
    for col in FLAG_COLS:
        table[col] = flags[col].map({True: "Ya", False: ""})
    cols = ["No", "NIM", "Nama", "Program Studi", "Jenis Aktivitas", "Mitra",
            "Status Mitra", "Status Aktivitas", "Tanggal Mulai", "Tanggal Selesai",
            "Jml MK", "Total SKS", "MK Konversi", "Judul Aktivitas",
            "Dosen Pembimbing", "Dosen Penguji"] + FLAG_COLS
    table = table[cols]
    stu_jenis = act.groupby("Jenis Aktivitas")["NIM"].nunique().rename("Mahasiswa")
    stu_prodi = act.groupby("Program Studi")["NIM"].nunique().sort_values(ascending=False).rename("Mahasiswa")
    with pd.ExcelWriter(buf, engine="openpyxl") as xl:
        table.to_excel(xl, sheet_name="Aktivitas", index=False)
        stu_jenis.to_excel(xl, sheet_name="Per Jenis")
        stu_prodi.to_excel(xl, sheet_name="Per Prodi")
    buf.seek(0)
    return buf.getvalue()


def build_pdf(act, raw, flags, source_note, drop_status, n_dropped, mk_overload):
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import mm
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
    from reportlab.lib.enums import TA_LEFT
    from reportlab.platypus import (SimpleDocTemplate, Paragraph, Spacer, Table,
                                    TableStyle, HRFlowable)

    navy = colors.HexColor(NAVY)
    red = colors.HexColor(RED)
    grey = colors.HexColor("#D9D9D9")
    light = colors.HexColor("#EEF1F7")
    mute = colors.HexColor("#666666")
    defs = flag_defs(mk_overload)

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
                           textColor=mute, leading=10.5)
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

    n_act = len(act)
    n_stu = act["NIM"].nunique()
    n_issue = int(pd.concat([flags[k] for k in RECORD_FLAGS], axis=1).any(axis=1).sum())
    act_by_jenis = act["Jenis Aktivitas"].value_counts()
    stu_by_jenis = act.groupby("Jenis Aktivitas")["NIM"].nunique()
    stu_by_prodi = act.groupby("Program Studi")["NIM"].nunique().sort_values(ascending=False)
    status_ct = act["Status Aktivitas"].value_counts()

    story = []
    story.append(Paragraph("Laporan Analisis Aktivitas MBKM &amp; Mata Kuliah Konversi", H1))
    story.append(Paragraph("Indonesia International Institute for Life Sciences (i3L)", SUB))
    if source_note:
        story.append(Paragraph(esc(source_note).replace("|", "&nbsp;|&nbsp;"), SUB))
    story.append(Spacer(1, 3))
    story.append(HRFlowable(width="100%", thickness=1.2, color=navy, spaceAfter=2))

    story.append(Paragraph("1. Ringkasan Umum", H2))
    
    rows = [["Metrik", "Nilai"],
            ["Jumlah aktivitas (setelah filter)", str(n_act)],
            ["Jumlah mahasiswa (NIM unik)", str(n_stu)],
            ["Baris pada file ekspor asli", str(len(raw))],
            ["Record dikecualikan (%s)" % (", ".join(drop_status) or "-"), str(n_dropped)],
            ["Aktivitas dengan isu data per-record", str(n_issue)]]
    
    t = Table(rows, colWidths=[120 * mm, 45 * mm]); t.setStyle(tstyle(numeric_cols=[1])); story.append(t)

    story.append(Paragraph("2. Rekap per Jenis Aktivitas", H2))
    rows = [["Jenis Aktivitas", "Mahasiswa"]]
    for j in act_by_jenis.index:
        rows.append([JENIS_SHORT.get(j, j), str(int(stu_by_jenis[j]))])
    rows.append(["TOTAL", str(n_stu)])
    t = Table(rows, colWidths=[130 * mm, 35 * mm]); t.setStyle(tstyle(total_row=True, numeric_cols=[1])); story.append(t)

    story.append(Paragraph("3. Rekap per Program Studi", H2))
    rows = [["Program Studi", "Mahasiswa"]]
    for p in stu_by_prodi.index:
        rows.append([Paragraph(esc(prodi_label(p)), CELL), str(int(stu_by_prodi[p]))])
    rows.append([Paragraph("TOTAL", CELLB), str(n_stu)])
    t = Table(rows, colWidths=[130 * mm, 35 * mm]); t.setStyle(tstyle(total_row=True, numeric_cols=[1])); story.append(t)

    story.append(Paragraph("4. Status Aktivitas", H2))
    rows = [["Status", "Aktivitas"]] + [[s, str(int(status_ct[s]))] for s in status_ct.index]
    t = Table(rows, colWidths=[130 * mm, 35 * mm]); t.setStyle(tstyle(numeric_cols=[1])); story.append(t)

    story.append(Paragraph("5. Temuan Kualitas Data (Flag)", H2))
    rows = [["Flag", "Jml", "Definisi"]]
    for k in FLAG_COLS:
        rows.append([Paragraph(esc(k), CELLB), Paragraph(str(int(flags[k].sum())), CELL),
                     Paragraph(esc(defs[k]), CELL)])
    t = Table(rows, colWidths=[40 * mm, 14 * mm, 111 * mm]); t.setStyle(tstyle(header_bg=red, numeric_cols=[1])); story.append(t)
    story.append(Paragraph("F1, F2, F3, F5 = isu input per-record. F4 &amp; F6 = isu sistemik.", SMALL))

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
    raw, act_all, source_note = load(up.getvalue(), up.name)
except Exception as exc:  # noqa
    st.error("Gagal membaca file: %s" % exc)
    st.stop()

with st.sidebar:
    st.header("Filter")
    all_status = sorted(act_all["Status Aktivitas"].unique())
    default_drop = [s for s in ["Ditolak", "Dibatalkan"] if s in all_status]
    drop_status = st.multiselect("Buang status aktivitas", all_status, default=default_drop)
    mk_overload = st.number_input("Ambang MK berlebih (F2)", min_value=1, max_value=50, value=5)
    st.caption("File asli: %d baris (1 baris per aktivitas × MK konversi)." % len(raw))
    if source_note:
        st.caption(source_note)

act = act_all[~act_all["Status Aktivitas"].isin(drop_status)].reset_index(drop=True)
act["No"] = range(1, len(act) + 1)
flags = flag_table(act, mk_overload)
defs = flag_defs(mk_overload)

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

tab_sum, tab_jenis, tab_prodi, tab_flag, tab_matrix, tab_dl = st.tabs(
    ["Ringkasan", "Per Jenis", "Per Program Studi", "Flag", "Matriks", "Unduh"])

with tab_sum:
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Status Aktivitas")
        st.pyplot(barh(act["Status Aktivitas"].value_counts(), NAVY, "Status aktivitas"))
    with c2:
        st.subheader("Flag")
        fc = pd.Series({k: int(flags[k].sum()) for k in FLAG_COLS})
        st.pyplot(barh(fc, RED, "Jumlah aktivitas ter-flag"))

with tab_jenis:
    st.subheader("Rekap mahasiswa per Jenis Aktivitas")
    stu_j = act.groupby("Jenis Aktivitas")["NIM"].nunique().sort_values(ascending=False)
    stu_j.index = [JENIS_SHORT.get(j, j) for j in stu_j.index]
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

    # tabel rekap (bisa diunduh)
    rekap_p = stu_p.rename("Mahasiswa").reset_index()
    rekap_p["Program Studi"] = rekap_p["Program Studi"].map(prodi_label)
    rekap_p = pd.concat(
        [rekap_p, pd.DataFrame([{"Program Studi": "TOTAL", "Mahasiswa": int(stu_p.sum())}])],
        ignore_index=True)

    stu_chart = stu_p.copy()
    stu_chart.index = [prodi_label(p) for p in stu_chart.index]
    st.pyplot(barh(stu_chart, NAVY, "Mahasiswa per Program Studi"))

    st.dataframe(rekap_p, use_container_width=True, hide_index=True)
    st.download_button("⬇️ Unduh rekap per Program Studi (CSV)",
                       rekap_p.to_csv(index=False).encode("utf-8"),
                       file_name="rekap_mahasiswa_per_prodi.csv", mime="text/csv",
                       key="dl_rekap_prodi")

    # daftar mahasiswa per prodi (drill-down)
    prodi_pick = st.selectbox("Lihat daftar mahasiswa untuk Program Studi", stu_p.index)
    sub = act[act["Program Studi"] == prodi_pick]
    st.caption("%d aktivitas, %d mahasiswa" % (len(sub), sub["NIM"].nunique()))
    st.dataframe(sub[["NIM", "Nama", "Jenis Aktivitas", "Status Aktivitas", "Mitra", "Judul Aktivitas"]]
                 .reset_index(drop=True), use_container_width=True)

with tab_flag:
    st.subheader("Flag kualitas data")
    cols = st.columns(6)
    for i, k in enumerate(FLAG_COLS):
        cols[i % 6].metric(k, int(flags[k].sum()), help=defs[k])
    st.divider()
    flag_pick = st.selectbox("Lihat record untuk flag", FLAG_COLS)
    st.caption(defs[flag_pick])
    sub = act[flags[flag_pick]]
    show = ["No", "NIM", "Nama", "Program Studi", "Jenis Aktivitas", "Status Aktivitas",
            "Mitra", "Jml MK", "Total SKS", "Dosen Pembimbing", "Judul Aktivitas"]
    st.dataframe(sub[show].reset_index(drop=True), use_container_width=True)

with tab_matrix:
    st.subheader("Matriks mahasiswa: Program Studi × Jenis Aktivitas")
    piv = act.groupby(["Program Studi", "Jenis Aktivitas"])["NIM"].nunique().unstack(fill_value=0)
    piv.columns = [JENIS_SHORT.get(c, c) for c in piv.columns]
    piv.index = [prodi_label(p) for p in piv.index]
    piv["Total"] = piv.sum(axis=1)
    piv.loc["TOTAL"] = piv.sum(axis=0)
    st.dataframe(piv, use_container_width=True)
    st.caption("Angka = jumlah mahasiswa (NIM unik).")

with tab_dl:
    st.subheader("Unduh hasil")
    stem = "MBKM_2026_Ganjil"
    try:
        pdf_bytes = build_pdf(act, raw, flags, source_note, drop_status, n_dropped, mk_overload)
        st.download_button("📄 Unduh laporan PDF", pdf_bytes, file_name=stem + "_laporan.pdf",
                           mime="application/pdf")
    except Exception as exc:  # noqa
        st.warning("PDF gagal dibuat (butuh paket reportlab): %s" % exc)
    xlsx_bytes = build_excel(act, flags)
    st.download_button("📊 Unduh data Excel (Aktivitas + rekap)", xlsx_bytes,
                       file_name=stem + "_bersih.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.download_button("📥 Unduh Aktivitas (CSV)", act.to_csv(index=False).encode("utf-8"),
                       file_name=stem + "_aktivitas.csv", mime="text/csv")
