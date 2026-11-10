"""Asset & Maintenance Control (Head, Trailer, Genset)
Streamlit + Google Sheets. Semua input berupa dropdown / tanggal / angka (tanpa free text),
kecuali kode unit & no. polisi di menu Kelola Asset (admin) yang divalidasi format.
"""
from datetime import datetime, timedelta, timezone

import gspread
import pandas as pd
import streamlit as st
from google.oauth2.service_account import Credentials
from gspread.utils import rowcol_to_a1

st.set_page_config(page_title="Asset & Maintenance", page_icon="🚛", layout="wide")

# ----------------------------------------------------------------------------
# Konstanta (dipakai logika dashboard - jangan diubah tanpa mengubah template)
# ----------------------------------------------------------------------------
TAB_ASSET, TAB_WO, TAB_DAILY, TAB_LIST = "Master Asset", "Work Order", "Status Harian", "Daftar Pilihan"
TIPE = ["Head", "Trailer", "Genset"]
STATUS_HARIAN = ["Jalan", "Standby", "Maintenance", "Breakdown"]
STATUS_WO = ["Menunggu", "Dikerjakan", "Tunggu Sparepart", "Selesai", "Batal"]
WO_OPEN = ["Menunggu", "Dikerjakan", "Tunggu Sparepart"]
JENIS = ["PM / Servis Berkala", "Perbaikan", "Ban", "Overhaul", "Kecelakaan / Body"]
JENIS_PM = JENIS[0]
KONDISI = ["Tidak bisa jalan (Breakdown)", "Masih bisa jalan", "Servis terjadwal (PM)"]
K_BREAKDOWN, K_MASIH_JALAN, K_PM = KONDISI
PRIORITAS = ["Tinggi", "Sedang", "Rendah"]
STATUS_ASET = ["Aktif", "Nonaktif"]
INTERVAL_PM = ["30", "45", "60", "90", "120", "180"]
TAHUN = [str(y) for y in range(2005, 2028)]

A_COLS = ["Kode Unit", "Tipe", "No Polisi / Serial", "Merk", "Tahun", "Lokasi",
          "Interval PM (hari)", "Status Aset", "Dibuat Oleh", "Waktu Input"]
W_COLS = ["WO ID", "Tanggal Masuk", "Kode Unit", "Tipe", "Jenis Pekerjaan", "Kategori Kerusakan",
          "Kondisi Unit", "Prioritas", "Lokasi Perbaikan", "PIC Mekanik", "Odometer / Jam",
          "Estimasi Selesai", "Status WO", "Tanggal Selesai", "Biaya (Rp)", "Dibuat Oleh",
          "Waktu Input", "Diupdate Oleh", "Waktu Update"]
D_COLS = ["Tanggal", "Kode Unit", "Tipe", "Status", "Lokasi", "Diisi Oleh", "Waktu Input"]
TAB_PAIR = "Pairing Harian"
P_COLS = ["Tanggal", "Kode Head", "Kode Trailer", "Diisi Oleh", "Waktu Input", "Driver"]
NO_TRAILER = "(tanpa trailer)"
NO_DRIVER = "(tanpa driver)"
TAB_DRIVER = "Master Driver"
DR_COLS = ["Kode Driver", "Nama Driver", "Jenis SIM", "Masa Berlaku SIM", "Status Driver", "Dibuat Oleh", "Waktu Input"]
JENIS_SIM = ["B1", "B1 Umum", "B2", "B2 Umum"]
STATUS_DRIVER = ["Aktif", "Cuti", "Sakit", "Nonaktif"]

ROLES = {
    "owner": ["Dashboard"],
    "input": ["Dashboard", "Ready & Pairing", "Work Order Baru", "Update Work Order", "Status Harian"],
    "admin": ["Dashboard", "Ready & Pairing", "Work Order Baru", "Update Work Order", "Status Harian", "Kelola Asset", "Kelola Driver"],
}

WIB = timezone(timedelta(hours=7))


def now_wib():
    return datetime.now(WIB)


def today():
    return now_wib().date()


def stamp():
    return now_wib().strftime("%Y-%m-%d %H:%M")


# ----------------------------------------------------------------------------
# Login sederhana berbasis PIN (isi di st.secrets["pins"]; kosong = mode uji coba)
# ----------------------------------------------------------------------------
def login():
    try:
        pins = dict(st.secrets["pins"])
    except Exception:
        return "admin"
    if st.session_state.get("role") in ROLES:
        return st.session_state["role"]
    st.title("🚛 Asset & Maintenance")
    pin = st.text_input("PIN", type="password")
    if st.button("Masuk"):
        role = next((r for r, p in pins.items() if pin and pin == str(p) and r in ROLES), None)
        if role:
            st.session_state["role"] = role
            st.rerun()
        else:
            st.error("PIN salah")
    st.stop()


# ----------------------------------------------------------------------------
# Akses Google Sheets
# ----------------------------------------------------------------------------
@st.cache_resource
def get_sheet():
    creds = Credentials.from_service_account_info(
        dict(st.secrets["gcp_service_account"]),
        scopes=["https://www.googleapis.com/auth/spreadsheets"],
    )
    return gspread.authorize(creds).open_by_key(st.secrets["SPREADSHEET_ID"])


def ensure_tab_once(name, header):
    """Buat tab bila belum ada, dan tambahkan kolom yang belum ada (sekali per sesi)."""
    key = f"tab_ok_{name}"
    if st.session_state.get(key):
        return
    sh = get_sheet()
    try:
        ws = sh.worksheet(name)
        have = ws.row_values(1)
        missing = [h for h in header if h not in have]
        if missing:
            new_header = have + missing
            if ws.col_count < len(new_header):
                ws.resize(cols=len(new_header))
            ws.update(values=[new_header], range_name="A1")
            st.cache_data.clear()
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title=name, rows=2000, cols=len(header))
        ws.update(values=[header], range_name="A1")
        st.cache_data.clear()
    st.session_state[key] = True


@st.cache_data(ttl=30)
def read_tab(name: str) -> pd.DataFrame:
    values = get_sheet().worksheet(name).get_all_values()
    if not values:
        return pd.DataFrame()
    raw_header = values[0]
    keep = [i for i, h in enumerate(raw_header) if str(h).strip()]  # abaikan kolom tanpa judul
    header = [raw_header[i].strip() for i in keep]
    rows = []
    for r in values[1:]:
        r = list(r) + [""] * (len(raw_header) - len(r))
        row = [r[i] for i in keep]
        if any(str(x).strip() for x in row):
            rows.append(row)
    df = pd.DataFrame(rows, columns=header)
    return df.loc[:, ~df.columns.duplicated()]


def append_dicts(tab, dicts):
    ws = get_sheet().worksheet(tab)
    header = ws.row_values(1)
    ws.append_rows([[d.get(h, "") for h in header] for d in dicts], value_input_option="RAW")
    st.cache_data.clear()


def rewrite_tab(tab, header, rows):
    ws = get_sheet().worksheet(tab)
    ws.clear()
    ws.update(values=[header] + rows, range_name="A1")
    st.cache_data.clear()


def update_wo_row(wo_id, changes: dict):
    ws = get_sheet().worksheet(TAB_WO)
    header = ws.row_values(1)
    cell = ws.find(wo_id, in_column=header.index("WO ID") + 1)
    data = [{"range": rowcol_to_a1(cell.row, header.index(k) + 1), "values": [[v]]} for k, v in changes.items()]
    ws.batch_update(data)
    st.cache_data.clear()


def get_lists():
    df = read_tab(TAB_LIST)
    out = {}
    for c in df.columns:
        if c.startswith("[TETAP]"):
            continue
        out[c] = [x.strip() for x in df[c].tolist() if str(x).strip()]
    return out


def load_all():
    assets = read_tab(TAB_ASSET).reindex(columns=A_COLS, fill_value="")
    wo = read_tab(TAB_WO).reindex(columns=W_COLS, fill_value="")
    daily = read_tab(TAB_DAILY).reindex(columns=D_COLS, fill_value="")
    for col in ("Tanggal Masuk", "Estimasi Selesai", "Tanggal Selesai"):
        wo[col] = pd.to_datetime(wo[col], errors="coerce")
    wo["Biaya (Rp)"] = pd.to_numeric(wo["Biaya (Rp)"], errors="coerce")
    daily["Tanggal"] = pd.to_datetime(daily["Tanggal"], errors="coerce")
    return assets, wo, daily


# ----------------------------------------------------------------------------
# Logika dashboard
# ----------------------------------------------------------------------------
def streak_start(g, status):
    start = None
    for _, r in g.sort_values("Tanggal").iloc[::-1].iterrows():
        if r["Status"] == status:
            start = r["Tanggal"]
        else:
            break
    return start


def compute_units(assets, wo, daily):
    """Status efektif tiap unit aktif: WO terbuka menang atas status harian."""
    t = pd.Timestamp(today())
    open_wo = wo[wo["Status WO"].isin(WO_OPEN)]
    rows = []
    for _, a in assets.iterrows():
        k = a["Kode Unit"]
        g = daily[daily["Kode Unit"] == k].sort_values("Tanggal")
        w = open_wo[open_wo["Kode Unit"] == k]
        last_daily = g.iloc[-1]["Status"] if len(g) else None
        if (w["Kondisi Unit"] == K_BREAKDOWN).any():
            status = "Breakdown"
        elif (w["Kondisi Unit"] == K_PM).any():
            status = "Maintenance"
        else:
            status = last_daily or "Belum ada data"
        downtime = None
        if status in ("Breakdown", "Maintenance"):
            if len(w) and w["Tanggal Masuk"].notna().any():
                downtime = (t - w["Tanggal Masuk"].min()).days
            else:
                s = streak_start(g, status)
                downtime = (t - s).days if s is not None else None
        jalan = g[g["Status"] == "Jalan"]
        idle_days = (t - jalan["Tanggal"].max()).days if len(jalan) else None
        pm_done = wo[(wo["Kode Unit"] == k) & (wo["Jenis Pekerjaan"] == JENIS_PM) & (wo["Status WO"] == "Selesai")]
        last_pm = pm_done["Tanggal Selesai"].max()
        interval = pd.to_numeric(a["Interval PM (hari)"], errors="coerce")
        next_pm = last_pm + pd.Timedelta(days=int(interval)) if pd.notna(last_pm) and pd.notna(interval) else pd.NaT
        rows.append({
            "Kode Unit": k, "Tipe": a["Tipe"], "No Polisi / Serial": a["No Polisi / Serial"],
            "Lokasi": a["Lokasi"], "Status": status, "Hari Downtime": downtime,
            "Ada WO Terbuka": len(w) > 0, "Hari Sejak Terakhir Jalan": idle_days,
            "PM Terakhir": last_pm, "PM Berikutnya": next_pm,
        })
    return pd.DataFrame(rows)


def utilisasi(assets, daily, days):
    start = pd.Timestamp(today()) - pd.Timedelta(days=days - 1)
    d = daily[(daily["Tanggal"] >= start) & (daily["Kode Unit"].isin(assets["Kode Unit"]))]
    if d.empty:
        return pd.DataFrame()
    t = d.groupby(["Tipe", "Status"]).size().unstack(fill_value=0).reindex(columns=STATUS_HARIAN, fill_value=0)
    t["Total Hari Tercatat"] = t[STATUS_HARIAN].sum(axis=1)
    den = t["Jalan"] + t["Standby"]
    t["Utilisasi %"] = (t["Jalan"] / den.where(den > 0) * 100).round(1)
    t["Availability %"] = ((t["Jalan"] + t["Standby"]) / t["Total Hari Tercatat"] * 100).round(1)
    return t


# ----------------------------------------------------------------------------
# Halaman: Dashboard
# ----------------------------------------------------------------------------
def page_dashboard():
    st.header("Dashboard Asset & Maintenance")
    assets, wo, daily = load_all()
    assets = assets[assets["Status Aset"] == "Aktif"]
    if assets.empty:
        st.info("Belum ada asset aktif. Isi Master Asset dulu.")
        return
    tipe_f = st.multiselect("Tipe asset", TIPE, default=TIPE)
    assets = assets[assets["Tipe"].isin(tipe_f)]
    wo = wo[wo["Kode Unit"].isin(assets["Kode Unit"])]
    units = compute_units(assets, wo, daily)
    if units.empty:
        st.info("Tidak ada unit untuk filter ini.")
        return
    t = pd.Timestamp(today())
    open_wo = wo[wo["Status WO"].isin(WO_OPEN)]
    late = open_wo[open_wo["Estimasi Selesai"] < t]
    rusak = units[units["Status"].isin(["Breakdown", "Maintenance"])]
    rusak_no_wo = rusak[~rusak["Ada WO Terbuka"]]
    ready = units["Status"].isin(["Jalan", "Standby"]).sum()
    no_data = (units["Status"] == "Belum ada data").sum()

    c = st.columns(6)
    c[0].metric("Unit aktif", len(units))
    c[1].metric("Ready (Jalan + Standby)", f"{ready / len(units) * 100:.0f}%")
    c[2].metric("Breakdown", int((units["Status"] == "Breakdown").sum()))
    c[3].metric("Maintenance", int((units["Status"] == "Maintenance").sum()))
    c[4].metric("WO terbuka", len(open_wo), delta=f"{len(late)} terlambat" if len(late) else None, delta_color="inverse")
    c[5].metric("Rusak tanpa WO", len(rusak_no_wo), help="Unit berstatus rusak/maintenance tetapi tidak ada work order terbuka = perbaikan belum termonitor.")
    if no_data:
        st.caption(f"{no_data} unit belum punya data status harian.")

    st.subheader("Status unit per tipe")
    ct = pd.crosstab(units["Tipe"], units["Status"])
    st.bar_chart(ct)

    st.subheader("Unit rusak / dalam perbaikan")
    if rusak.empty:
        st.success("Tidak ada unit rusak saat ini.")
    else:
        show = rusak.sort_values("Hari Downtime", ascending=False, na_position="last")
        show = show[["Kode Unit", "Tipe", "No Polisi / Serial", "Status", "Hari Downtime", "Ada WO Terbuka"]]
        st.dataframe(show, hide_index=True, width="stretch")
        if len(rusak_no_wo):
            st.warning("Unit tanpa WO terbuka: " + ", ".join(rusak_no_wo["Kode Unit"]) + ". Buat Work Order agar perbaikannya termonitor.")

    st.subheader("Monitoring perbaikan")
    if open_wo.empty:
        st.caption("Tidak ada work order terbuka.")
    else:
        w = open_wo.copy()
        w["Hari Berjalan"] = (t - w["Tanggal Masuk"]).dt.days
        w["Terlambat"] = w["Estimasi Selesai"] < t
        w = w.sort_values(["Terlambat", "Hari Berjalan"], ascending=False)
        st.dataframe(
            w[["WO ID", "Kode Unit", "Kategori Kerusakan", "Prioritas", "PIC Mekanik", "Status WO",
               "Estimasi Selesai", "Hari Berjalan", "Terlambat"]],
            hide_index=True, width="stretch",
            column_config={"Estimasi Selesai": st.column_config.DateColumn(format="DD MMM YYYY")},
        )
        wait = open_wo[open_wo["Status WO"] == "Tunggu Sparepart"]
        if len(wait):
            st.caption(f"{len(wait)} WO tertahan menunggu sparepart.")

    st.subheader("Utilisasi")
    days = st.radio("Periode", [7, 30, 90], index=1, horizontal=True, format_func=lambda x: f"{x} hari")
    u = utilisasi(assets, daily, days)
    if u.empty:
        st.caption("Belum ada data status harian pada periode ini.")
    else:
        st.dataframe(u.reset_index(), hide_index=True, width="stretch")
        st.caption("Utilisasi = hari Jalan ÷ (Jalan + Standby). Availability = (Jalan + Standby) ÷ semua hari tercatat.")
    idle = units[units["Status"] == "Standby"].sort_values("Hari Sejak Terakhir Jalan", ascending=False, na_position="first")
    if len(idle):
        st.markdown("**Unit Standby (siap tapi tidak jalan)**")
        st.dataframe(idle[["Kode Unit", "Tipe", "No Polisi / Serial", "Lokasi", "Hari Sejak Terakhir Jalan"]],
                     hide_index=True, width="stretch")

    st.subheader("Jadwal PM")
    pm = units.copy()
    pm["Sisa Hari"] = (pm["PM Berikutnya"] - t).dt.days
    due = pm[pm["Sisa Hari"] <= 7].sort_values("Sisa Hari")
    if len(due):
        st.dataframe(due[["Kode Unit", "Tipe", "PM Terakhir", "PM Berikutnya", "Sisa Hari"]], hide_index=True, width="stretch",
                     column_config={"PM Terakhir": st.column_config.DateColumn(format="DD MMM YYYY"),
                                    "PM Berikutnya": st.column_config.DateColumn(format="DD MMM YYYY")})
    else:
        st.caption("Tidak ada PM jatuh tempo dalam 7 hari ke depan.")
    belum = pm[pm["PM Terakhir"].isna()]
    if len(belum):
        st.caption(f"{len(belum)} unit belum punya riwayat PM tercatat: " + ", ".join(belum["Kode Unit"]))

    cl, cr = st.columns(2)
    with cl:
        st.subheader("Kerusakan terbanyak (90 hari)")
        k = wo[(wo["Jenis Pekerjaan"] != JENIS_PM) & (wo["Tanggal Masuk"] >= t - pd.Timedelta(days=90))]
        if k.empty:
            st.caption("Belum ada data.")
        else:
            st.bar_chart(k["Kategori Kerusakan"].value_counts())
    with cr:
        st.subheader("Biaya perbaikan per bulan (Rp)")
        b = wo[(wo["Status WO"] == "Selesai") & wo["Biaya (Rp)"].notna() & wo["Tanggal Selesai"].notna()]
        if b.empty:
            st.caption("Belum ada data biaya.")
        else:
            st.bar_chart(b.groupby(b["Tanggal Selesai"].dt.to_period("M").astype(str))["Biaya (Rp)"].sum())

    st.subheader("Driver: SIM segera habis (60 hari)")
    al = sim_alerts(load_drivers())
    if al.empty:
        st.caption("Tidak ada SIM yang habis dalam 60 hari ke depan.")
    else:
        st.dataframe(al[["Nama Driver", "Jenis SIM", "Masa Berlaku SIM", "Sisa Hari", "Status Driver"]],
                     hide_index=True, width="stretch",
                     column_config={"Masa Berlaku SIM": st.column_config.DateColumn(format="DD MMM YYYY")})


# ----------------------------------------------------------------------------
# Halaman: Work Order Baru
# ----------------------------------------------------------------------------
def next_wo_id(wo):
    nums = pd.to_numeric(wo["WO ID"].str.replace("WO-", "", regex=False), errors="coerce").dropna()
    return f"WO-{(int(nums.max()) if len(nums) else 0) + 1:04d}"


def page_wo_baru():
    st.header("Work Order Baru")
    lists = get_lists()
    assets, wo, _ = load_all()
    assets = assets[assets["Status Aset"] == "Aktif"]
    if assets.empty:
        st.warning("Belum ada asset aktif.")
        return
    labels = {f"{r['Kode Unit']} — {r['Tipe']} — {r['No Polisi / Serial']}": r for _, r in assets.iterrows()}
    pick = st.selectbox("Unit", list(labels), index=None, placeholder="Pilih unit")
    if pick is None:
        return
    a = labels[pick]
    already = wo[(wo["Kode Unit"] == a["Kode Unit"]) & wo["Status WO"].isin(WO_OPEN)]
    if len(already):
        st.info(f"Unit ini sudah punya {len(already)} WO terbuka: " + ", ".join(already["WO ID"]) + ". Pastikan ini pekerjaan yang berbeda.")
    with st.form("wo_baru", clear_on_submit=True):
        c1, c2 = st.columns(2)
        tgl = c1.date_input("Tanggal masuk", value=today(), max_value=today())
        jenis = c2.selectbox("Jenis pekerjaan", JENIS, index=None, placeholder="Pilih")
        kategori = c1.selectbox("Kategori kerusakan", lists.get(f"Kategori {a['Tipe']}", []), index=None, placeholder="Pilih")
        kondisi = c2.selectbox("Kondisi unit", KONDISI, index=None, placeholder="Pilih")
        prio = c1.selectbox("Prioritas", PRIORITAS, index=None, placeholder="Pilih")
        lokasi = c2.selectbox("Lokasi perbaikan", lists.get("Lokasi Perbaikan", []), index=None, placeholder="Pilih")
        pic = c1.selectbox("PIC mekanik", lists.get("Mekanik", []), index=None, placeholder="Pilih")
        est = c2.date_input("Estimasi selesai", value=today() + timedelta(days=3))
        odo = c1.number_input("Odometer (km) / jam kerja genset", min_value=0, max_value=5_000_000, step=1, value=0, help="Isi 0 bila belum diketahui.")
        oleh = c2.selectbox("Diisi oleh", lists.get("Tim", []), index=None, placeholder="Pilih nama")
        ok = st.form_submit_button("Simpan Work Order", type="primary")
    if ok:
        missing = [n for n, v in {"Jenis pekerjaan": jenis, "Kategori": kategori, "Kondisi unit": kondisi,
                                  "Prioritas": prio, "Lokasi perbaikan": lokasi, "PIC mekanik": pic,
                                  "Diisi oleh": oleh}.items() if v is None]
        if missing:
            st.error("Belum dipilih: " + ", ".join(missing))
        elif est < tgl:
            st.error("Estimasi selesai tidak boleh sebelum tanggal masuk.")
        else:
            wo_id = next_wo_id(wo)
            append_dicts(TAB_WO, [{
                "WO ID": wo_id, "Tanggal Masuk": tgl.isoformat(), "Kode Unit": a["Kode Unit"], "Tipe": a["Tipe"],
                "Jenis Pekerjaan": jenis, "Kategori Kerusakan": kategori, "Kondisi Unit": kondisi,
                "Prioritas": prio, "Lokasi Perbaikan": lokasi, "PIC Mekanik": pic,
                "Odometer / Jam": int(odo) if odo else "", "Estimasi Selesai": est.isoformat(),
                "Status WO": "Menunggu", "Tanggal Selesai": "", "Biaya (Rp)": "",
                "Dibuat Oleh": oleh, "Waktu Input": stamp(), "Diupdate Oleh": "", "Waktu Update": "",
            }])
            st.success(f"Work order {wo_id} tersimpan untuk unit {a['Kode Unit']}.")


# ----------------------------------------------------------------------------
# Halaman: Update Work Order
# ----------------------------------------------------------------------------
def page_wo_update():
    st.header("Update Work Order")
    lists = get_lists()
    _, wo, _ = load_all()
    open_ = wo[wo["Status WO"].isin(WO_OPEN)]
    if open_.empty:
        st.success("Tidak ada work order terbuka.")
        return
    labels = {f"{r['WO ID']} — {r['Kode Unit']} — {r['Kategori Kerusakan']} — {r['Status WO']}": r for _, r in open_.iterrows()}
    pick = st.selectbox("Work order", list(labels), index=None, placeholder="Pilih WO")
    if pick is None:
        return
    r = labels[pick]
    est_default = r["Estimasi Selesai"].date() if pd.notna(r["Estimasi Selesai"]) else today()
    mekanik = lists.get("Mekanik", [])
    with st.form("wo_update"):
        c1, c2 = st.columns(2)
        status = c1.selectbox("Status WO", STATUS_WO, index=STATUS_WO.index(r["Status WO"]))
        est = c2.date_input("Estimasi selesai", value=est_default)
        pic = c1.selectbox("PIC mekanik", mekanik, index=mekanik.index(r["PIC Mekanik"]) if r["PIC Mekanik"] in mekanik else None)
        oleh = c2.selectbox("Diupdate oleh", lists.get("Tim", []), index=None, placeholder="Pilih nama")
        st.caption("Dua isian di bawah hanya dipakai bila status = Selesai.")
        tgl_selesai = c1.date_input("Tanggal selesai", value=today(), max_value=today())
        biaya = c2.number_input("Total biaya (Rp)", min_value=0, max_value=10_000_000_000, step=50_000, value=0)
        ok = st.form_submit_button("Simpan Update", type="primary")
    if ok:
        if oleh is None or pic is None:
            st.error("PIC mekanik dan 'Diupdate oleh' harus dipilih.")
        elif status == "Selesai" and pd.notna(r["Tanggal Masuk"]) and pd.Timestamp(tgl_selesai) < r["Tanggal Masuk"]:
            st.error("Tanggal selesai tidak boleh sebelum tanggal masuk.")
        else:
            ch = {"Status WO": status, "Estimasi Selesai": est.isoformat(), "PIC Mekanik": pic,
                  "Diupdate Oleh": oleh, "Waktu Update": stamp()}
            if status == "Selesai":
                ch["Tanggal Selesai"] = tgl_selesai.isoformat()
                ch["Biaya (Rp)"] = int(biaya)
            update_wo_row(r["WO ID"], ch)
            st.success(f"{r['WO ID']} diperbarui → {status}.")
            st.rerun() if status in ("Selesai", "Batal") else None


# ----------------------------------------------------------------------------
# Halaman: Status Harian (satu layar untuk semua unit)
# ----------------------------------------------------------------------------
def page_status_harian():
    st.header("Status Harian Unit")
    lists = get_lists()
    assets, _, daily = load_all()
    assets = assets[assets["Status Aset"] == "Aktif"]
    lokasi_opts = lists.get("Lokasi", [])
    if assets.empty or not lokasi_opts:
        st.warning("Master Asset atau daftar Lokasi masih kosong.")
        return
    c1, c2 = st.columns([1, 2])
    tgl = c1.date_input("Tanggal", value=today(), max_value=today())
    tipe_f = c2.multiselect("Tampilkan tipe", TIPE, default=TIPE)
    units = assets[assets["Tipe"].isin(tipe_f)]
    ts = pd.Timestamp(tgl)
    exist = daily[daily["Tanggal"] == ts]
    before = daily[daily["Tanggal"] < ts].sort_values("Tanggal").groupby("Kode Unit").last()
    if len(exist[exist["Kode Unit"].isin(units["Kode Unit"])]):
        st.info("Tanggal ini sudah pernah diisi. Menyimpan akan menimpa data unit yang tampil.")
    rows = []
    for _, a in units.iterrows():
        k = a["Kode Unit"]
        e = exist[exist["Kode Unit"] == k]
        if len(e):
            s, l = e.iloc[-1]["Status"], e.iloc[-1]["Lokasi"]
        elif k in before.index:
            s, l = before.loc[k, "Status"], before.loc[k, "Lokasi"]
        else:
            s, l = "Standby", a["Lokasi"]
        rows.append({"Kode Unit": k, "Tipe": a["Tipe"], "No Polisi / Serial": a["No Polisi / Serial"],
                     "Status": s if s in STATUS_HARIAN else "Standby",
                     "Lokasi": l if l in lokasi_opts else lokasi_opts[0]})
    df = pd.DataFrame(rows)
    oleh_opts = lists.get("Tim", [])
    with st.form("harian"):
        edited = st.data_editor(
            df, hide_index=True, width="stretch", num_rows="fixed",
            disabled=["Kode Unit", "Tipe", "No Polisi / Serial"],
            column_config={
                "Status": st.column_config.SelectboxColumn("Status", options=STATUS_HARIAN, required=True),
                "Lokasi": st.column_config.SelectboxColumn("Lokasi", options=lokasi_opts, required=True),
            },
            key=f"harian_{tgl}_{'-'.join(tipe_f)}",
        )
        oleh = st.selectbox("Diisi oleh", oleh_opts, index=None, placeholder="Pilih nama")
        ok = st.form_submit_button("Simpan Status Harian", type="primary")
    if ok:
        if oleh is None:
            st.error("Pilih nama pengisi.")
        elif edited.empty:
            st.error("Tidak ada unit untuk disimpan.")
        else:
            new_rows = [[tgl.isoformat(), r["Kode Unit"], r["Tipe"], r["Status"], r["Lokasi"], oleh, stamp()]
                        for _, r in edited.iterrows()]
            ws = get_sheet().worksheet(TAB_DAILY)
            values = ws.get_all_values()
            header = values[0] if values else D_COLS
            ix_t, ix_k = header.index("Tanggal"), header.index("Kode Unit")
            codes = set(edited["Kode Unit"])
            keep = [r for r in values[1:] if not (r[ix_t] == tgl.isoformat() and r[ix_k] in codes)]
            rewrite_tab(TAB_DAILY, header, keep + new_rows)
            n_rusak = int(edited["Status"].isin(["Breakdown", "Maintenance"]).sum())
            st.success(f"Tersimpan: {len(new_rows)} unit untuk {tgl:%d %b %Y}.")
            if n_rusak:
                st.warning(f"{n_rusak} unit berstatus Breakdown/Maintenance. Pastikan masing-masing punya Work Order.")


# ----------------------------------------------------------------------------
# Halaman: Kelola Asset (admin)
# ----------------------------------------------------------------------------
import re

KODE_RE = re.compile(r"^[A-Z0-9-]{3,12}$")
PLAT_RE = re.compile(r"^[A-Z0-9 -]{3,15}$")


def with_current(base, series):
    return list(dict.fromkeys(list(base) + [v for v in series if str(v).strip()]))


def page_kelola_asset():
    st.header("Kelola Asset")
    lists = get_lists()
    assets, _, _ = load_all()
    tab1, tab2 = st.tabs(["Tambah asset", "Ubah data asset"])
    with tab1:
        with st.form("asset_baru", clear_on_submit=True):
            c1, c2 = st.columns(2)
            kode = c1.text_input("Kode unit (huruf besar/angka/-, mis. H-003)", max_chars=12)
            plat = c2.text_input("No. polisi / serial (mis. N 1234 AB)", max_chars=15)
            tipe = c1.selectbox("Tipe", TIPE, index=None, placeholder="Pilih")
            merk = c2.selectbox("Merk", lists.get("Merk", []), index=None, placeholder="Pilih")
            tahun = c1.selectbox("Tahun", TAHUN, index=None, placeholder="Pilih")
            lokasi = c2.selectbox("Lokasi", lists.get("Lokasi", []), index=None, placeholder="Pilih")
            interval = c1.selectbox("Interval PM (hari)", INTERVAL_PM, index=None, placeholder="Pilih")
            oleh = c2.selectbox("Dibuat oleh", lists.get("Tim", []), index=None, placeholder="Pilih nama")
            ok = st.form_submit_button("Tambah asset", type="primary")
        if ok:
            kode, plat = kode.strip().upper(), plat.strip().upper()
            errs = []
            if not KODE_RE.match(kode):
                errs.append("Kode unit: 3-12 karakter, hanya huruf, angka, dan tanda -.")
            if not PLAT_RE.match(plat):
                errs.append("No. polisi/serial: 3-15 karakter, hanya huruf, angka, spasi, dan tanda -.")
            if kode in set(assets["Kode Unit"]):
                errs.append("Kode unit sudah ada.")
            if plat in set(assets["No Polisi / Serial"]):
                errs.append("No. polisi/serial sudah ada.")
            if None in (tipe, merk, tahun, lokasi, interval, oleh):
                errs.append("Semua dropdown harus dipilih.")
            if errs:
                for e in errs:
                    st.error(e)
            else:
                append_dicts(TAB_ASSET, [{
                    "Kode Unit": kode, "Tipe": tipe, "No Polisi / Serial": plat, "Merk": merk, "Tahun": tahun,
                    "Lokasi": lokasi, "Interval PM (hari)": interval, "Status Aset": "Aktif",
                    "Dibuat Oleh": oleh, "Waktu Input": stamp(),
                }])
                st.success(f"Asset {kode} ditambahkan.")
    with tab2:
        if assets.empty:
            st.caption("Belum ada asset.")
            return
        st.caption("Kode, tipe, dan no. polisi dikunci. Untuk asset yang dijual/tidak dipakai, ubah Status Aset menjadi Nonaktif (jangan dihapus agar riwayat tetap ada).")
        edited = st.data_editor(
            assets, hide_index=True, width="stretch", num_rows="fixed", key="edit_assets",
            disabled=["Kode Unit", "Tipe", "No Polisi / Serial", "Dibuat Oleh", "Waktu Input"],
            column_config={
                "Merk": st.column_config.SelectboxColumn(options=with_current(lists.get("Merk", []), assets["Merk"])),
                "Tahun": st.column_config.SelectboxColumn(options=with_current(TAHUN, assets["Tahun"])),
                "Lokasi": st.column_config.SelectboxColumn(options=with_current(lists.get("Lokasi", []), assets["Lokasi"])),
                "Interval PM (hari)": st.column_config.SelectboxColumn(options=with_current(INTERVAL_PM, assets["Interval PM (hari)"])),
                "Status Aset": st.column_config.SelectboxColumn(options=STATUS_ASET, required=True),
            },
        )
        if st.button("Simpan perubahan", type="primary"):
            rewrite_tab(TAB_ASSET, A_COLS, edited[A_COLS].astype(str).values.tolist())
            st.success("Perubahan disimpan.")



# ----------------------------------------------------------------------------
# Halaman: Ready & Pairing Head-Trailer-Driver
# ----------------------------------------------------------------------------
def pairing_defaults(pair, day, head_codes, trailer_codes, driver_names):
    """Isian awal: pairing tanggal itu bila ada, kalau tidak pairing hari terakhir.
    Hasil: {kode head: (trailer, driver)}"""
    cur = pair[pair["Tanggal"] == day]
    if cur.empty:
        prev = pair[pair["Tanggal"] < day]
        cur = prev[prev["Tanggal"] == prev["Tanggal"].max()] if len(prev) else prev
    tm = dict(zip(cur["Kode Head"], cur["Kode Trailer"]))
    dm = dict(zip(cur["Kode Head"], cur["Driver"]))
    return {h: (tm.get(h) if tm.get(h) in trailer_codes else NO_TRAILER,
                dm.get(h) if dm.get(h) in driver_names else NO_DRIVER) for h in head_codes}


def page_pairing():
    st.header("Unit Ready & Pairing Head–Trailer–Driver")
    if "flash" in st.session_state:
        st.success(st.session_state.pop("flash"))
    ensure_tab_once(TAB_PAIR, P_COLS)
    lists = get_lists()
    assets, wo, daily = load_all()
    pair = read_tab(TAB_PAIR).reindex(columns=P_COLS, fill_value="")
    pair["Tanggal"] = pd.to_datetime(pair["Tanggal"], errors="coerce")
    assets = assets[assets["Status Aset"] == "Aktif"]
    units = compute_units(assets, wo, daily)
    if units.empty:
        st.info("Belum ada asset aktif.")
        return
    info = units.set_index("Kode Unit")
    ready = units[units["Status"].isin(["Jalan", "Standby"])]
    heads = ready[ready["Tipe"] == "Head"].sort_values("Kode Unit")
    trailers = ready[ready["Tipe"] == "Trailer"].sort_values("Kode Unit")
    head_codes, trailer_codes = list(heads["Kode Unit"]), list(trailers["Kode Unit"])
    drv = load_drivers()
    driver_names = sorted(drv[drv["Status Driver"] == "Aktif"]["Nama Driver"])

    tgl = st.date_input("Tanggal pairing", value=today(), min_value=today(),
                        max_value=today() + timedelta(days=1),
                        help="Hari ini atau besok (untuk rencana). Kesiapan unit selalu berdasarkan status terbaru.")
    ts = pd.Timestamp(tgl)
    cur = pair[pair["Tanggal"] == ts]
    with_trailer = cur[cur["Kode Trailer"] != ""]
    with_driver = cur[cur["Driver"] != ""]

    c = st.columns(5)
    c[0].metric("Head Ready", len(head_codes))
    c[1].metric("Trailer Ready", len(trailer_codes))
    c[2].metric("Head + trailer", len(with_trailer))
    c[3].metric("Head + driver", len(with_driver))
    c[4].metric("Trailer Ready belum terpasang", len(set(trailer_codes) - set(with_trailer["Kode Trailer"])))

    st.subheader("Update pairing")
    if not driver_names:
        st.warning("Belum ada driver berstatus Aktif. Tambahkan driver di menu **Kelola Driver** (admin), lalu klik Refresh data.")
    if not head_codes:
        st.info("Belum ada Head berstatus Jalan/Standby. Isi atau perbarui Status Harian dulu.")
    elif not trailer_codes:
        st.info("Belum ada Trailer berstatus Jalan/Standby. Isi atau perbarui Status Harian dulu.")
    else:
        if cur.empty:
            st.caption("Belum disimpan untuk tanggal ini. Isian awal mengikuti pairing terakhir; periksa lalu simpan.")
        defaults = pairing_defaults(pair, ts, head_codes, trailer_codes, driver_names)
        df = pd.DataFrame([{"Kode Head": h, "No Polisi Head": info.loc[h, "No Polisi / Serial"],
                            "Status Head": info.loc[h, "Status"], "Trailer": defaults[h][0],
                            "Driver": defaults[h][1]} for h in head_codes])
        t_opts = [NO_TRAILER] + trailer_codes
        d_opts = [NO_DRIVER] + driver_names
        with st.form("pairing"):
            edited = st.data_editor(
                df, hide_index=True, width="stretch", num_rows="fixed",
                disabled=["Kode Head", "No Polisi Head", "Status Head"],
                column_config={
                    "Trailer": st.column_config.SelectboxColumn("Trailer", options=t_opts, required=True),
                    "Driver": st.column_config.SelectboxColumn("Driver", options=d_opts, required=True),
                },
                key=f"pair_{tgl}_{len(head_codes)}_{len(trailer_codes)}_{len(driver_names)}",
            )
            oleh = st.selectbox("Diisi oleh", lists.get("Tim", []), index=None, placeholder="Pilih nama")
            ok = st.form_submit_button("Simpan Pairing", type="primary")
        if ok:
            tr = edited[edited["Trailer"] != NO_TRAILER]["Trailer"]
            dr = edited[edited["Driver"] != NO_DRIVER]["Driver"]
            dup_t, dup_d = sorted(set(tr[tr.duplicated()])), sorted(set(dr[dr.duplicated()]))
            if oleh is None:
                st.error("Pilih nama pengisi.")
            elif dup_t or dup_d:
                if dup_t:
                    st.error("Trailer dipakai lebih dari satu Head: " + ", ".join(dup_t))
                if dup_d:
                    st.error("Driver dipasang di lebih dari satu Head: " + ", ".join(dup_d))
            else:
                ws = get_sheet().worksheet(TAB_PAIR)
                values = ws.get_all_values()
                header = values[0] if values else P_COLS
                ix = header.index("Tanggal")
                keep = [r + [""] * (len(header) - len(r)) for r in values[1:] if r[ix] != tgl.isoformat()]
                new = []
                for _, r in edited.iterrows():
                    rec = {"Tanggal": tgl.isoformat(), "Kode Head": r["Kode Head"],
                           "Kode Trailer": "" if r["Trailer"] == NO_TRAILER else r["Trailer"],
                           "Driver": "" if r["Driver"] == NO_DRIVER else r["Driver"],
                           "Diisi Oleh": oleh, "Waktu Input": stamp()}
                    new.append([rec.get(h, "") for h in header])
                rewrite_tab(TAB_PAIR, header, keep + new)
                st.session_state["flash"] = (f"Pairing {tgl:%d %b %Y} tersimpan "
                                             f"({len(tr)} trailer, {len(dr)} driver).")
                st.rerun()

    st.subheader("Pasangan tersimpan")
    shown = cur[(cur["Kode Trailer"] != "") | (cur["Driver"] != "")]
    if cur.empty:
        st.caption("Belum ada pairing tersimpan untuk tanggal ini.")
    else:
        rows, bad = [], []
        for _, r in shown.iterrows():
            h, t = r["Kode Head"], r["Kode Trailer"]
            hs = info["Status"].get(h, "-")
            ts_ = info["Status"].get(t, "-") if t else "-"
            rows.append({"Head": h, "Plat Head": info["No Polisi / Serial"].get(h, ""), "Status Head": hs,
                         "Trailer": t or "-", "Plat Trailer": info["No Polisi / Serial"].get(t, "") if t else "",
                         "Status Trailer": ts_, "Driver": r["Driver"] or "-"})
            bad += [f"{k} ({v})" for k, v in ((h, hs), (t, ts_)) if k and v not in ("Jalan", "Standby", "-")]
        if rows:
            st.dataframe(pd.DataFrame(rows), hide_index=True, width="stretch")
        if bad:
            st.warning("Pasangan ini memuat unit yang sudah tidak Ready: " + ", ".join(bad) + ". Perbarui pairing.")
        dinfo = drv.set_index("Nama Driver")
        tt = pd.Timestamp(today())
        dwarn = []
        for nm in shown["Driver"]:
            if not nm:
                continue
            stt = dinfo["Status Driver"].get(nm, "tidak ada di Master Driver")
            exp = dinfo["Masa Berlaku SIM"].get(nm, pd.NaT)
            if stt != "Aktif":
                dwarn.append(f"{nm} ({stt})")
            elif pd.notna(exp) and exp < tt:
                dwarn.append(f"{nm} (SIM habis {exp:%d %b %Y})")
        if dwarn:
            st.warning("Periksa driver: " + ", ".join(dwarn) + ".")
        no_trailer = [h for h in head_codes if h not in set(with_trailer["Kode Head"])]
        free_trailer = [t for t in trailer_codes if t not in set(with_trailer["Kode Trailer"])]
        no_driver = [h for h in head_codes if h not in set(with_driver["Kode Head"])]
        free_driver = [d for d in driver_names if d not in set(with_driver["Driver"])]
        a1, a2 = st.columns(2)
        a1.markdown("**Head Ready tanpa trailer**")
        a1.write(", ".join(no_trailer) if no_trailer else "—")
        a2.markdown("**Trailer Ready belum terpasang**")
        a2.write(", ".join(free_trailer) if free_trailer else "—")
        b1, b2 = st.columns(2)
        b1.markdown("**Head Ready tanpa driver**")
        b1.write(", ".join(no_driver) if no_driver else "—")
        b2.markdown("**Driver belum ditugaskan**")
        b2.write(", ".join(free_driver) if free_driver else "—")

    with st.expander("Riwayat pairing 7 hari terakhir"):
        h = pair[(pair["Tanggal"] >= ts - pd.Timedelta(days=7)) & ((pair["Kode Trailer"] != "") | (pair["Driver"] != ""))]
        if h.empty:
            st.caption("Belum ada riwayat.")
        else:
            st.dataframe(h.sort_values(["Tanggal", "Kode Head"], ascending=[False, True])[["Tanggal", "Kode Head", "Kode Trailer", "Driver", "Diisi Oleh"]],
                         hide_index=True, width="stretch",
                         column_config={"Tanggal": st.column_config.DateColumn(format="DD MMM YYYY")})


# ----------------------------------------------------------------------------
# Master Driver
# ----------------------------------------------------------------------------
def load_drivers():
    ensure_tab_once(TAB_DRIVER, DR_COLS)
    d = read_tab(TAB_DRIVER).reindex(columns=DR_COLS, fill_value="")
    d["Masa Berlaku SIM"] = pd.to_datetime(d["Masa Berlaku SIM"], errors="coerce")
    return d


def next_driver_code(drv):
    nums = pd.to_numeric(drv["Kode Driver"].str.replace("D-", "", regex=False), errors="coerce").dropna()
    return f"D-{(int(nums.max()) if len(nums) else 0) + 1:03d}"


def sim_alerts(drv, days=60):
    """Driver (selain Nonaktif) yang SIM-nya sudah habis atau habis dalam `days` hari."""
    x = drv[drv["Status Driver"] != "Nonaktif"].copy()
    x["Sisa Hari"] = (x["Masa Berlaku SIM"] - pd.Timestamp(today())).dt.days
    return x[x["Sisa Hari"] <= days].sort_values("Sisa Hari")


NAMA_RE = re.compile(r"^[A-Z .'-]{3,40}$")


def page_kelola_driver():
    st.header("Kelola Driver")
    lists = get_lists()
    drv = load_drivers()
    al = sim_alerts(drv)
    if len(al):
        st.warning(f"{len(al)} driver: SIM sudah habis atau habis dalam 60 hari.")
        st.dataframe(al[["Kode Driver", "Nama Driver", "Jenis SIM", "Masa Berlaku SIM", "Sisa Hari", "Status Driver"]],
                     hide_index=True, width="stretch",
                     column_config={"Masa Berlaku SIM": st.column_config.DateColumn(format="DD MMM YYYY")})
    tab1, tab2 = st.tabs(["Tambah driver", "Ubah data driver"])
    with tab1:
        with st.form("driver_baru", clear_on_submit=True):
            c1, c2 = st.columns(2)
            nama = c1.text_input("Nama driver (huruf saja)", max_chars=40)
            sim = c2.selectbox("Jenis SIM", JENIS_SIM, index=None, placeholder="Pilih")
            exp = c1.date_input("Masa berlaku SIM", value=today() + timedelta(days=365),
                                min_value=today() - timedelta(days=3650), max_value=today() + timedelta(days=3650))
            oleh = c2.selectbox("Dibuat oleh", lists.get("Tim", []), index=None, placeholder="Pilih nama")
            ok = st.form_submit_button("Tambah driver", type="primary")
        if ok:
            nama = " ".join(nama.upper().split())
            errs = []
            if not NAMA_RE.match(nama):
                errs.append("Nama: 3-40 karakter, hanya huruf, spasi, titik, apostrof, dan tanda -.")
            if nama in set(drv["Nama Driver"].str.upper()):
                errs.append("Nama driver sudah ada.")
            if sim is None or oleh is None:
                errs.append("Jenis SIM dan 'Dibuat oleh' harus dipilih.")
            if errs:
                for e in errs:
                    st.error(e)
            else:
                code = next_driver_code(drv)
                append_dicts(TAB_DRIVER, [{
                    "Kode Driver": code, "Nama Driver": nama, "Jenis SIM": sim,
                    "Masa Berlaku SIM": exp.isoformat(), "Status Driver": "Aktif",
                    "Dibuat Oleh": oleh, "Waktu Input": stamp(),
                }])
                st.success(f"Driver {code} — {nama} ditambahkan.")
    with tab2:
        if drv.empty:
            st.caption("Belum ada driver.")
            return
        st.caption("Kode dan nama dikunci. Driver yang berhenti: ubah Status menjadi Nonaktif (jangan dihapus agar riwayat pairing tetap terbaca). Driver berstatus Cuti/Sakit tidak muncul di pilihan pairing.")
        edited = st.data_editor(
            drv, hide_index=True, width="stretch", num_rows="fixed", key="edit_drivers",
            disabled=["Kode Driver", "Nama Driver", "Dibuat Oleh", "Waktu Input"],
            column_config={
                "Jenis SIM": st.column_config.SelectboxColumn(options=with_current(JENIS_SIM, drv["Jenis SIM"])),
                "Masa Berlaku SIM": st.column_config.DateColumn("Masa Berlaku SIM", format="DD MMM YYYY"),
                "Status Driver": st.column_config.SelectboxColumn(options=STATUS_DRIVER, required=True),
            },
        )
        if st.button("Simpan perubahan driver", type="primary"):
            out = edited.copy()
            out["Masa Berlaku SIM"] = pd.to_datetime(out["Masa Berlaku SIM"], errors="coerce").dt.strftime("%Y-%m-%d").fillna("")
            rewrite_tab(TAB_DRIVER, DR_COLS, out[DR_COLS].astype(str).values.tolist())
            st.success("Perubahan disimpan.")


# ----------------------------------------------------------------------------
# Main
# ----------------------------------------------------------------------------
PAGES = {
    "Dashboard": page_dashboard,
    "Ready & Pairing": page_pairing,
    "Work Order Baru": page_wo_baru,
    "Update Work Order": page_wo_update,
    "Status Harian": page_status_harian,
    "Kelola Asset": page_kelola_asset,
    "Kelola Driver": page_kelola_driver,
}


def main():
    role = login()
    st.sidebar.title("🚛 Asset & Maintenance")
    page = st.sidebar.radio("Menu", ROLES[role])
    if st.sidebar.button("Refresh data"):
        st.cache_data.clear()
        st.rerun()
    if "role" in st.session_state and st.sidebar.button("Keluar"):
        del st.session_state["role"]
        st.rerun()
    try:
        PAGES[page]()
    except gspread.exceptions.WorksheetNotFound as e:
        st.error(f"Tab tidak ditemukan di spreadsheet: {e}. Gunakan template yang disediakan dan jangan mengubah nama tab.")
    except gspread.exceptions.APIError as e:
        st.error("Gagal mengakses Google Sheets (mungkin terlalu sering diakses). Coba lagi sebentar lagi.")
        st.caption(str(e)[:200])


main()
