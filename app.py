"""
Glovebox: a shared car log for gas fill-ups, service records and receipts.

Data lives in a Google Sheet you own (tabs: Cars, Fuel, Service).
Receipts are saved to a "Glovebox Receipts" folder in your Google Drive
through a small Apps Script helper (see receipts_apps_script.gs).
"""

import base64
import hmac
import io
import uuid
from datetime import date, datetime
from pathlib import Path
from urllib.parse import quote

import gspread
import pandas as pd
import requests
import streamlit as st
from PIL import Image, ImageOps

try:  # lets Pillow read iPhone HEIC photos so they're saved as viewable JPEGs
    from pillow_heif import register_heif_opener
    register_heif_opener()
except Exception:
    pass

MIME_BY_EXT = {".pdf": "application/pdf", ".jpg": "image/jpeg", ".jpeg": "image/jpeg",
               ".png": "image/png", ".webp": "image/webp", ".heic": "image/heic",
               ".heif": "image/heif"}

_icon_path = Path(__file__).with_name("glovebox_icon.png")
st.set_page_config(page_title="Glovebox Car Log",
                   page_icon=Image.open(_icon_path) if _icon_path.exists() else "🚗",
                   layout="centered")

# ---------------------------------------------------------------------------
# Sheet layout. The app creates these tabs and header rows if they are missing.
# ---------------------------------------------------------------------------
TABS = {
    "Cars": ["car_id", "year", "make", "model", "start_odometer", "added_by", "added_on"],
    "Fuel": ["entry_id", "car_id", "date", "odometer", "gallons", "total_cost",
             "price_per_gallon", "full_tank", "missed_before", "entered_by", "logged_at"],
    "Service": ["entry_id", "car_id", "date", "odometer", "service_type", "cost", "shop",
                "notes", "receipt_url", "receipt_file_id", "entered_by", "logged_at",
                "visit_id"],
    "Schedule": ["car_id", "item", "every_miles", "every_months", "first_due_miles",
                 "first_due_months", "counts_as", "notes"],
}
NUMERIC = ("odometer", "gallons", "total_cost", "price_per_gallon", "cost", "start_odometer",
           "every_miles", "every_months", "first_due_miles", "first_due_months")
BOOLS = ("full_tank", "missed_before")

SERVICE_TYPES = [
    "Oil change", "Tire rotation", "New tires", "Brake inspection", "Brake pads",
    "Brake rotors", "Brake fluid", "Battery", "Engine air filter", "Cabin filter",
    "Wiper blades", "Alignment", "Wheel bearing", "Coolant flush", "Inverter coolant",
    "Transmission (CVT) fluid", "Rear differential fluid", "Hybrid battery filter", "Spark plugs",
    "State inspection", "Registration", "Car wash", "Other",
]

# Maintenance schedule templates: (item, every_miles, every_months, first_due_miles,
# first_due_months, counts_as [service types that reset it, ";"-separated], notes)
# Subaru values follow Subaru's normal-condition schedule for 2015-2016 models.
# Always confirm against your owner's manual.
SCHEDULE_TEMPLATES = {
    "Subaru (2015-2016, normal driving)": [
        ("Oil change", 6000, 6, None, None, "Oil change",
         "Severe driving (short trips, stop-and-go, towing): every 3,000 mi / 3 mo."),
        ("Tire rotation", 6000, 6, None, None, "Tire rotation;New tires", ""),
        ("Brake inspection", 12000, 12, None, None, "Brake inspection;Brake pads;Brake rotors",
         "Severe driving: every 6,000 mi / 6 mo."),
        ("Cabin filter", 12000, 12, None, None, "Cabin filter", ""),
        ("Wiper blades", None, 12, None, None, "Wiper blades",
         "Common practice; not part of Subaru's schedule."),
        ("Engine air filter", 30000, 30, None, None, "Engine air filter;Air filter",
         "Replace sooner in dusty areas."),
        ("Brake fluid", 30000, 30, None, None, "Brake fluid",
         "Hilly or humid areas: every 15,000 mi / 15 mo."),
        ("Transmission (CVT) fluid", 30000, 30, None, None, "Transmission (CVT) fluid",
         "Inspection under normal driving; replacement only for severe use or towing."),
        ("Spark plugs", 60000, 72, None, None, "Spark plugs",
         "Confirm this interval in your owner's manual."),
        ("Coolant flush", 75000, 72, 137500, 132, "Coolant flush",
         "First at 137,500 mi / 11 yrs, then every 75,000 mi / 6 yrs."),
    ],
    "Toyota RAV4 Hybrid (2019-2025, normal driving)": [
        ("Oil change", 10000, 12, None, None, "Oil change",
         "0W-16 synthetic. Severe driving (dust, towing, freezing, low-speed): 5,000 mi / 6 mo."),
        ("Tire rotation", 5000, 6, None, None, "Tire rotation;New tires", ""),
        ("Hybrid battery filter", 20000, None, None, None, "Hybrid battery filter",
         "Clean the hybrid battery cooling intake filter. Inspect every 5,000 mi / 6 mo."),
        ("Wiper blades", None, 12, None, None, "Wiper blades",
         "Common practice; not part of Toyota's schedule."),
        ("Cabin filter", 30000, 36, None, None, "Cabin filter", ""),
        ("Engine air filter", 30000, 36, None, None, "Engine air filter;Air filter",
         "Inspect every 5,000 mi on dusty or dirt roads."),
        ("Rear differential fluid", 30000, 36, None, None, "Rear differential fluid",
         "Inspection only under normal driving."),
        ("Spark plugs", 120000, 144, None, None, "Spark plugs", ""),
        ("Coolant flush", 50000, 60, 100000, 120, "Coolant flush",
         "Engine coolant: first at 100,000 mi / 10 yrs, then every 50,000 mi / 5 yrs."),
        ("Inverter coolant", 50000, 60, 150000, 180, "Inverter coolant",
         "Hybrid system coolant: first at 150,000 mi / 15 yrs, then every 50,000 mi / 5 yrs."),
    ],
    "Typical car (general guide)": [
        ("Oil change", 5000, 6, None, None, "Oil change", "Use your manual's interval if different."),
        ("Tire rotation", 6000, 6, None, None, "Tire rotation;New tires", ""),
        ("Brake inspection", 12000, 12, None, None, "Brake inspection;Brake pads;Brake rotors", ""),
        ("Cabin filter", 15000, 12, None, None, "Cabin filter", ""),
        ("Wiper blades", None, 12, None, None, "Wiper blades", ""),
        ("Engine air filter", 30000, 36, None, None, "Engine air filter;Air filter", ""),
        ("Brake fluid", 30000, 36, None, None, "Brake fluid", ""),
        ("Transmission fluid", 60000, 60, None, None, "Transmission (CVT) fluid", ""),
        ("Spark plugs", 100000, None, None, None, "Spark plugs", "Older engines: 30,000 mi."),
        ("Coolant flush", 100000, 60, None, None, "Coolant flush", ""),
    ],
}
DUE_SOON_MILES, DUE_SOON_DAYS = 1000, 30
RECEIPT_TYPES = ["jpg", "jpeg", "png", "heic", "heif", "webp", "pdf"]


# ---------------------------------------------------------------------------
# Small helpers
# ---------------------------------------------------------------------------
def now_stamp() -> str:
    return datetime.now().strftime("%Y-%m-%d %H:%M")


def fmt_date(x) -> str:
    return x.strftime("%b %d, %Y") if pd.notna(x) else "No date"


def to_bool(v, default=False) -> bool:
    s = str(v).strip().upper()
    if s in ("TRUE", "1", "YES", "Y"):
        return True
    if s in ("FALSE", "0", "NO", "N"):
        return False
    return default


def money(x) -> str:
    return f"${x:,.2f}" if pd.notna(x) else "--"


def form_key(name: str) -> str:
    """Widget keys that reset after a successful save (bump the version)."""
    v = st.session_state.get(f"_v_{name}", 0)
    return f"{name}_{v}"


def reset_form(name: str):
    st.session_state[f"_v_{name}"] = st.session_state.get(f"_v_{name}", 0) + 1


def flash(msg: str):
    st.session_state["_flash"] = msg


# ---------------------------------------------------------------------------
# Password gate (one shared household password stored in Streamlit secrets)
# ---------------------------------------------------------------------------
def password_ok() -> bool:
    """Ways in, set in secrets:
    - access_key: a personal link (…streamlit.app/?key=...) opens the app with no typing.
    - app_password: the password screen, for anyone without the link.
    - neither set: the app is open to anyone with the URL.
    """
    if st.session_state.get("authed"):
        return True
    key = str(st.secrets.get("access_key", "")).strip()
    pw_secret = str(st.secrets.get("app_password", "")).strip()
    if key and hmac.compare_digest(str(st.query_params.get("key", "")), key):
        st.session_state.authed = True
        return True
    if not key and not pw_secret:
        return True
    st.title("🚗 Glovebox")
    if not pw_secret:
        st.info("Open Glovebox from your personal link or home screen icon.")
        return False
    with st.form("login"):
        pw = st.text_input("Password", type="password")
        if st.form_submit_button("Open", width="stretch"):
            if hmac.compare_digest(pw, pw_secret):
                st.session_state.authed = True
                st.rerun()
            else:
                st.error("That password didn't match. Try again.")
    return False


# ---------------------------------------------------------------------------
# Google sign-in gate (used when [auth] and [allowed_users] are in secrets)
# ---------------------------------------------------------------------------
def allowed_users() -> dict:
    """{"email": "Display name"} from secrets, emails lowercased."""
    return {str(k).strip().lower(): str(v) for k, v in st.secrets.get("allowed_users", {}).items()}


def google_login_enabled() -> bool:
    return "auth" in st.secrets and bool(allowed_users())


def google_ok() -> str | None:
    """Returns the signed-in person's name, or None after drawing the sign-in screen."""
    if not st.user.get("is_logged_in", False):
        st.title("🚗 Glovebox")
        st.write("Sign in with your Google account to open your car log.")
        st.button("Sign in with Google", type="primary", width="stretch", on_click=st.login)
        return None
    email = str(st.user.get("email", "")).strip().lower()
    name = allowed_users().get(email)
    if not name:
        st.title("🚗 Glovebox")
        st.error(f"{email or 'This account'} isn't on the list for this app. "
                 "Ask Justin to add it, or sign in with a different Google account.")
        st.button("Sign out", on_click=st.logout)
        return None
    return name


# ---------------------------------------------------------------------------
# Google Sheets
# ---------------------------------------------------------------------------
@st.cache_resource
def get_book():
    gc = gspread.service_account_from_dict(dict(st.secrets["gcp_service_account"]))
    book = gc.open_by_key(st.secrets["sheet_id"])
    existing = {w.title: w for w in book.worksheets()}
    for name, headers in TABS.items():
        if name not in existing:
            w = book.add_worksheet(title=name, rows=1000, cols=len(headers))
            w.append_row(headers)
        else:
            w = existing[name]
            if w.row_values(1) != headers:  # add any new columns to an older sheet
                if w.col_count < len(headers):
                    w.resize(cols=len(headers))
                w.update(range_name="A1", values=[headers])
    return book


def ws(name: str):
    return get_book().worksheet(name)


@st.cache_data(ttl=60)
def load(name: str) -> pd.DataFrame:
    rows = ws(name).get_all_records()
    df = pd.DataFrame(rows)
    for col in TABS[name]:
        if col not in df.columns:
            df[col] = None
    df = df[TABS[name]]
    for col in NUMERIC:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col].replace("", None), errors="coerce")
    if "date" in df.columns:
        df["date"] = pd.to_datetime(df["date"], errors="coerce")
    if "full_tank" in df.columns:
        df["full_tank"] = df["full_tank"].map(lambda v: to_bool(v, default=True))
    if "missed_before" in df.columns:
        df["missed_before"] = df["missed_before"].map(lambda v: to_bool(v, default=False))
    for col in ("car_id", "entry_id", "receipt_url", "receipt_file_id", "notes", "shop",
                "item", "counts_as", "service_type", "visit_id"):
        if col in df.columns:
            df[col] = df[col].fillna("").astype(str)
    return df


def row_values(name: str, data: dict) -> list:
    out = []
    for col in TABS[name]:
        v = data.get(col, "")
        if isinstance(v, bool):
            v = "TRUE" if v else "FALSE"
        elif isinstance(v, pd.Timestamp):
            v = v.date().isoformat() if pd.notna(v) else ""
        elif v is None or (isinstance(v, float) and pd.isna(v)):
            v = ""
        out.append(v)
    return out


def append(name: str, data: dict):
    ws(name).append_row(row_values(name, data), value_input_option="USER_ENTERED")
    load.clear()


def append_many(name: str, rows: list[dict]):
    ws(name).append_rows([row_values(name, d) for d in rows], value_input_option="USER_ENTERED")
    load.clear()


def visit_key(r) -> str:
    """Rows from one shop visit share a key. Older rows without a visit_id are grouped
    by date + mileage + shop."""
    if r.get("visit_id"):
        return r["visit_id"]
    if pd.isna(r.get("date")):
        return r["entry_id"]
    odo = r.get("odometer")
    return f"{r['date']:%Y-%m-%d}|{'' if pd.isna(odo) else int(odo)}|{r.get('shop', '')}"


def receipt_shared(file_id: str, other_than: set) -> bool:
    """True if another service row still points at this receipt file."""
    if not file_id:
        return False
    rows = load("Service")
    return bool(((rows["receipt_file_id"] == file_id) & ~rows["entry_id"].isin(other_than)).any())


def monthly_spend(fuel_df: pd.DataFrame, svc_df: pd.DataFrame) -> pd.DataFrame:
    """Gas and service spending per calendar month, from the first record to this month.
    Months with nothing logged are included as $0."""
    sp = pd.concat([
        pd.DataFrame({"date": fuel_df["date"], "kind": "Gas", "amt": fuel_df["total_cost"]}),
        pd.DataFrame({"date": svc_df["date"], "kind": "Service", "amt": svc_df["cost"]}),
    ]).dropna(subset=["date"])
    if sp.empty:
        return pd.DataFrame(columns=["Gas", "Service"])
    sp["amt"] = sp["amt"].fillna(0)
    sp["month"] = sp["date"].dt.to_period("M")
    pv = sp.pivot_table(index="month", columns="kind", values="amt", aggfunc="sum")
    pv = pv.reindex(columns=["Gas", "Service"], fill_value=0)
    months = pd.period_range(pv.index.min(), pd.Timestamp(date.today()).to_period("M"), freq="M")
    return pv.reindex(months, fill_value=0).fillna(0)


def update(name: str, entry_id: str, data: dict):
    sheet = ws(name)
    cell = sheet.find(entry_id, in_column=1)
    if not cell:
        raise RuntimeError("Entry not found. Tap Refresh and try again.")
    sheet.update(range_name=f"A{cell.row}", values=[row_values(name, data)],
                 value_input_option="USER_ENTERED")
    load.clear()


def delete(name: str, entry_id: str):
    sheet = ws(name)
    cell = sheet.find(entry_id, in_column=1)
    if cell:
        sheet.delete_rows(cell.row)
    load.clear()


def save_schedule(car_id: str, items: list[dict]):
    """Replace one car's schedule rows, leaving other cars untouched."""
    sheet = ws("Schedule")
    keep = [r for r in sheet.get_all_values()[1:] if r and r[0] != car_id]
    rows = [TABS["Schedule"]] + keep + [row_values("Schedule", {**i, "car_id": car_id}) for i in items]
    sheet.clear()
    sheet.update(range_name="A1", values=rows, value_input_option="USER_ENTERED")
    load.clear()


# ---------------------------------------------------------------------------
# Maintenance schedule math
# ---------------------------------------------------------------------------
def schedule_status(sched: pd.DataFrame, svc: pd.DataFrame, cur_odo, today: date) -> pd.DataFrame:
    out = []
    ts_today = pd.Timestamp(today)
    for _, it in sched.iterrows():
        names = {n.strip().lower() for n in str(it["counts_as"] or it["item"]).split(";") if n.strip()}
        names.add(str(it["item"]).strip().lower())
        done = svc[svc["service_type"].str.strip().str.lower().isin(names)].sort_values(
            ["date", "odometer"], na_position="first")
        every_mi, every_mo = it["every_miles"], it["every_months"]
        last_date = last_odo = next_mi = next_date = None
        if not done.empty:
            last = done.iloc[-1]
            last_date = last["date"] if pd.notna(last["date"]) else None
            with_odo = done.dropna(subset=["odometer"])
            last_odo = with_odo["odometer"].iloc[-1] if not with_odo.empty else None
            if last_odo is not None and pd.notna(every_mi):
                next_mi = last_odo + every_mi
            if last_date is not None and pd.notna(every_mo):
                next_date = last_date + pd.DateOffset(months=int(every_mo))
        elif pd.notna(it["first_due_miles"]) and pd.notna(cur_odo) and cur_odo < it["first_due_miles"]:
            next_mi = it["first_due_miles"]  # e.g. Subaru's first coolant change

        miles_left = (next_mi - cur_odo) if (next_mi is not None and pd.notna(cur_odo)) else None
        days_left = (next_date - ts_today).days if next_date is not None else None

        if done.empty:
            status = "No record"
        elif (miles_left is not None and miles_left <= 0) or (days_left is not None and days_left <= 0):
            status = "Overdue"
        elif (miles_left is not None and miles_left <= DUE_SOON_MILES) or \
             (days_left is not None and days_left <= DUE_SOON_DAYS):
            status = "Due soon"
        else:
            status = "OK"

        progress = None
        if last_odo is not None and pd.notna(every_mi) and pd.notna(cur_odo) and every_mi:
            progress = (cur_odo - last_odo) / every_mi
        if days_left is not None and pd.notna(every_mo) and every_mo:
            p2 = 1 - days_left / (every_mo * 30.44)
            progress = max(progress or 0, p2)
        out.append({"item": it["item"], "status": status, "last_date": last_date,
                    "last_odo": last_odo, "next_mi": next_mi, "next_date": next_date,
                    "miles_left": miles_left, "days_left": days_left, "progress": progress,
                    "every_miles": every_mi, "every_months": every_mo, "notes": it["notes"]})
    return pd.DataFrame(out)


# ---------------------------------------------------------------------------
# Receipts (Apps Script helper saves into your Drive)
# ---------------------------------------------------------------------------
def receipts_enabled() -> bool:
    return bool(st.secrets.get("receipts_url")) and bool(st.secrets.get("receipts_token"))


def _call_receipts(payload: dict) -> dict:
    payload["token"] = st.secrets["receipts_token"]
    r = requests.post(st.secrets["receipts_url"], json=payload, timeout=120)
    r.raise_for_status()
    res = r.json()
    if not res.get("ok"):
        raise RuntimeError(res.get("error", "Upload failed"))
    return res


def upload_receipt(file, label: str) -> tuple[str, str]:
    data, name = file.getvalue(), file.name
    ext = Path(name).suffix.lower()
    # Phones don't always report the file type; go by the extension so Drive can preview it.
    mime = MIME_BY_EXT.get(ext) or file.type or "application/octet-stream"
    if mime.startswith("image/"):
        # Shrink big phone photos so uploads are quick; keep the original if Pillow can't read it.
        try:
            img = ImageOps.exif_transpose(Image.open(io.BytesIO(data)))
            img.thumbnail((2000, 2000))
            buf = io.BytesIO()
            img.convert("RGB").save(buf, "JPEG", quality=85)
            data, mime, name = buf.getvalue(), "image/jpeg", Path(name).stem + ".jpg"
        except Exception:
            pass
    res = _call_receipts({
        "action": "upload",
        "filename": f"{label} - {name}",
        "mimeType": mime,
        "data": base64.b64encode(data).decode(),
    })
    return res["url"], res["id"]


def receipt_link(url: str) -> str:
    """Open Drive links with the right Google account when several are signed in.
    Set `google_account = "you@gmail.com"` in secrets to enable."""
    acct = str(st.secrets.get("google_account", "")).strip()
    if not url or not acct or "authuser=" in url:
        return url
    return f"{url}{'&' if '?' in url else '?'}authuser={quote(acct)}"


def trash_receipt(file_id: str):
    if file_id and receipts_enabled():
        try:
            _call_receipts({"action": "trash", "id": file_id})
        except Exception:
            pass  # the entry is still removed; the file can be cleaned up in Drive


# ---------------------------------------------------------------------------
# Fuel math
# ---------------------------------------------------------------------------
def fuel_with_mpg(fuel: pd.DataFrame) -> pd.DataFrame:
    """MPG per full fill-up, handling partial fills and missed fill-ups.

    - Full tank: MPG = miles since the last full tank / all gallons bought since then.
    - Partial fill: gallons carry forward to the next full tank; no MPG on this row.
    - Missed a fill-up before this one: the chain restarts; no MPG on this row.
    - Gallons unknown (older records): the chain restarts at this fill-up.
    """
    f = fuel.assign(_partial=~fuel["full_tank"].astype(bool)).sort_values(
        ["odometer", "date", "_partial"]).drop(columns="_partial")
    mpg, seg_miles, seg_gal, seg_cost = [], [], [], []
    anchor, acc_gal, acc_cost = None, 0.0, 0.0
    for _, r in f.iterrows():
        m = sm = sg = sc = None
        if pd.isna(r["odometer"]):
            pass  # no mileage: can't place it, leave it out of MPG
        elif pd.isna(r["gallons"]) or not r["gallons"]:
            anchor, acc_gal, acc_cost = r["odometer"], 0.0, 0.0
        elif r["missed_before"]:
            anchor, acc_gal, acc_cost = (r["odometer"] if r["full_tank"] else None), 0.0, 0.0
        else:
            acc_gal += r["gallons"] or 0
            acc_cost += r["total_cost"] or 0
            if r["full_tank"]:
                if anchor is not None and r["odometer"] > anchor and acc_gal > 0:
                    sm, sg, sc = r["odometer"] - anchor, acc_gal, acc_cost
                    m = sm / sg
                anchor, acc_gal, acc_cost = r["odometer"], 0.0, 0.0
        mpg.append(m); seg_miles.append(sm); seg_gal.append(sg); seg_cost.append(sc)
    f["mpg"], f["seg_miles"], f["seg_gal"], f["seg_cost"] = mpg, seg_miles, seg_gal, seg_cost
    return f


def summary(fuel: pd.DataFrame, svc: pd.DataFrame) -> dict:
    out = {"mpg": None, "cpm": None, "fuel_total": fuel["total_cost"].sum(),
           "svc_total": svc["cost"].fillna(0).sum()}
    if not fuel.empty:
        f = fuel_with_mpg(fuel)
        miles, gal, cost = f["seg_miles"].sum(), f["seg_gal"].sum(), f["seg_cost"].sum()
        if miles and gal:
            out["mpg"], out["cpm"] = miles / gal, cost / miles
    return out


def order_warning(fuel: pd.DataFrame, d: date, odo: float, skip_id: str = "") -> str:
    """Catch typos: a later date with lower mileage (or the reverse)."""
    other = fuel[fuel["entry_id"] != skip_id].dropna(subset=["date", "odometer"])
    ts = pd.Timestamp(d)
    later_lower = other[(other["date"] > ts) & (other["odometer"] < odo)]
    earlier_higher = other[(other["date"] < ts) & (other["odometer"] > odo)]
    clash = pd.concat([later_lower, earlier_higher])
    if clash.empty:
        return ""
    r = clash.iloc[0]
    return (f"Mileage doesn't line up with your {fmt_date(r['date'])} fill-up "
            f"({r['odometer']:,.0f} mi). Check the date and mileage.")


# ---------------------------------------------------------------------------
# App
# ---------------------------------------------------------------------------
me = None
if google_login_enabled():
    me = google_ok()
    if not me:
        st.stop()
    drivers = list(dict.fromkeys(allowed_users().values()))
elif not password_ok():
    st.stop()
else:
    drivers = list(st.secrets.get("drivers", ["Me"]))

st.title("🚗 Glovebox")

if msg := st.session_state.pop("_flash", None):
    st.toast(msg)

top1, top2 = st.columns([2, 1])
with top2:
    driver = st.selectbox("Who's logging?", drivers, key="driver",
                          index=drivers.index(me) if me in drivers else 0)

cars = load("Cars")

# ---- Add a car -------------------------------------------------------------
with st.expander("➕ Add a car", expanded=cars.empty):
    with st.form("add_car", clear_on_submit=True):
        c1, c2, c3 = st.columns(3)
        year = c1.text_input("Year", placeholder="2015")
        make = c2.text_input("Make", placeholder="Subaru")
        model = c3.text_input("Model", placeholder="Outback")
        start_odo = st.number_input("Current mileage (optional)", min_value=0, step=1, value=None)
        if st.form_submit_button("Save car", width="stretch"):
            if not (year.isdigit() and len(year) == 4) or not make.strip():
                st.error("Enter a 4-digit year and a make, like 2015 Subaru.")
            else:
                append("Cars", {"car_id": uuid.uuid4().hex[:8], "year": year,
                                "make": make.strip(), "model": model.strip(),
                                "start_odometer": start_odo, "added_by": driver,
                                "added_on": now_stamp()})
                flash("Car added")
                st.rerun()

if cars.empty:
    st.info("Add your first car above to start logging.")
    st.stop()

labels = {r["car_id"]: " ".join(str(x) for x in (r["year"], r["make"], r["model"]) if str(x).strip())
          for _, r in cars.iterrows()}
with top1:
    car_id = st.selectbox("Car", list(labels), format_func=labels.get, key="car")
car_name = labels[car_id]

fuel_all, svc_all = load("Fuel"), load("Service")
fuel = fuel_all[fuel_all["car_id"] == car_id]
svc = svc_all[svc_all["car_id"] == car_id]
sched = load("Schedule")
sched = sched[sched["car_id"] == car_id]
car_row = cars[cars["car_id"] == car_id].iloc[0]

# Latest known mileage for this car
_readings = pd.concat([fuel["odometer"], svc["odometer"],
                       pd.Series([car_row["start_odometer"]], dtype="float")]).dropna()
cur_odo = _readings.max() if not _readings.empty else None

# Service dropdown: standard list plus this car's schedule items
service_choices = list(dict.fromkeys(
    [*SERVICE_TYPES[:-1], *[i for i in sched["item"] if i], "Other"]))

# ---- Stat tiles ------------------------------------------------------------
s = summary(fuel, svc)
m1, m2 = st.columns(2)
m1.metric("Average MPG", f"{s['mpg']:.1f}" if s["mpg"] else "--",
          help="Needs two full-tank fill-ups in a row.")
m2.metric("Fuel cost per mile", f"${s['cpm']:.3f}" if s["cpm"] else "--")
m3, m4 = st.columns(2)
m3.metric("Spent on gas", money(s["fuel_total"]))
m4.metric("Spent on service", money(s["svc_total"]))

tab_fuel, tab_svc, tab_sched, tab_time, tab_trend, tab_fix = st.tabs(
    ["⛽ Fill-up", "🔧 Service", "🗓️ Schedule", "🕒 Timeline", "📊 Trends", "✏️ Fix"])

# ---- Fill-up form ----------------------------------------------------------
with tab_fuel:
    k = form_key("fuel")
    with st.form(k):
        c1, c2 = st.columns(2)
        f_date = c1.date_input("Date", value=date.today(), max_value=date.today(),
                               key=k + "d", help="Forgot to log one? Pick the day you filled up.")
        f_odo = c2.number_input("Mileage (odometer)", min_value=0, step=1, value=None,
                                placeholder="98450", key=k + "o")
        c3, c4 = st.columns(2)
        f_cost = c3.number_input("Total cost ($)", min_value=0.0, step=0.01, value=None,
                                 format="%.2f", placeholder="42.18", key=k + "c")
        f_gal = c4.number_input("Gallons", min_value=0.0, step=0.001, value=None,
                                format="%.3f", placeholder="12.415", key=k + "g")
        f_full = st.checkbox("Filled the tank all the way", value=True, key=k + "f")
        f_missed = st.checkbox("I missed logging a fill-up before this one", key=k + "m",
                               help="Keeps your MPG accurate when a fill-up wasn't recorded.")
        f_anyway = False
        if st.session_state.get("fuel_warn"):
            st.warning(st.session_state["fuel_warn"])
            f_anyway = st.checkbox("It's correct, save anyway", key=k + "a")
        if st.form_submit_button("Log fill-up", width="stretch", type="primary"):
            if not f_odo or not f_cost or not f_gal:
                st.error("Fill in mileage, total cost and gallons.")
            else:
                warn = order_warning(fuel, f_date, f_odo)
                if warn and not f_anyway:
                    st.session_state["fuel_warn"] = warn
                    st.rerun()
                append("Fuel", {"entry_id": uuid.uuid4().hex[:10], "car_id": car_id,
                                "date": f_date.isoformat(), "odometer": int(f_odo),
                                "gallons": round(f_gal, 3), "total_cost": round(f_cost, 2),
                                "price_per_gallon": round(f_cost / f_gal, 3),
                                "full_tank": f_full, "missed_before": f_missed,
                                "entered_by": driver, "logged_at": now_stamp()})
                st.session_state.pop("fuel_warn", None)
                reset_form("fuel")
                flash(f"Fill-up saved: ${f_cost / f_gal:.3f}/gal")
                st.rerun()

# ---- Service form ----------------------------------------------------------
with tab_svc:
    # Plain widgets (not a form) so cost boxes appear as you pick services.
    k = form_key("svc")
    c1, c2 = st.columns(2)
    s_date = c1.date_input("Date", value=date.today(), max_value=date.today(), key=k + "d")
    s_odo = c2.number_input("Mileage (odometer)", min_value=0, step=1, value=None,
                            placeholder="98450", key=k + "o")
    s_types = st.multiselect("What was done? Pick everything from this visit",
                             [c for c in service_choices if c != "Other"], key=k + "t",
                             placeholder="Oil change, Tire rotation...")
    s_other = st.text_input("Anything not in the list? Separate items with commas",
                            placeholder="Rear wheel bearings, Headlight bulb", key=k + "x")
    jobs = list(dict.fromkeys(s_types + [x.strip() for x in s_other.split(",") if x.strip()]))

    per_item = False
    if len(jobs) > 1:
        per_item = st.radio("Cost", ["One total for the visit", "A cost for each item"],
                            horizontal=True, key=k + "cm") == "A cost for each item"
    item_costs, s_total = {}, None
    if per_item:
        cols = st.columns(2)
        for i, j in enumerate(jobs):
            item_costs[j] = cols[i % 2].number_input(f"{j} ($)", min_value=0.0, step=0.01,
                                                     value=None, format="%.2f", key=f"{k}c_{j}")
    else:
        s_total = st.number_input("Total cost ($)", min_value=0.0, step=0.01, value=None,
                                  format="%.2f", key=k + "c")
    s_shop = st.text_input("Shop", placeholder="Optional", key=k + "s")
    s_notes = st.text_area("Notes", placeholder="Parts used, next due, warranty...", key=k + "n")
    s_file = None
    if receipts_enabled():
        s_file = st.file_uploader("Receipt (photo or PDF)", type=RECEIPT_TYPES, key=k + "r")

    label = ("Log service" if len(jobs) <= 1 else f"Log {len(jobs)} services")
    if st.button(label, width="stretch", type="primary", key=k + "go"):
        if not jobs:
            st.error("Pick at least one service, or type what was done.")
        else:
            url = fid = ""
            if s_file:
                try:
                    with st.spinner("Uploading receipt..."):
                        url, fid = upload_receipt(
                            s_file, f"{s_date.isoformat()} {car_name} {' + '.join(jobs)}")
                except Exception as e:
                    st.error(f"Receipt didn't upload ({e}). Nothing was saved yet. Try again.")
                    st.stop()
            visit = uuid.uuid4().hex[:10]
            rows = []
            for i, j in enumerate(jobs):
                if per_item:
                    cost = item_costs.get(j)
                elif s_total is None:
                    cost = None
                else:
                    cost = s_total if i == 0 else 0
                note = s_notes.strip()
                if not per_item and i > 0 and s_total is not None:
                    note = (note + " · " if note else "") + f"Cost included with {jobs[0]}"
                rows.append({"entry_id": uuid.uuid4().hex[:10], "car_id": car_id,
                             "date": s_date.isoformat(),
                             "odometer": int(s_odo) if s_odo else "",
                             "service_type": j,
                             "cost": round(cost, 2) if cost is not None else "",
                             "shop": s_shop.strip(), "notes": note,
                             "receipt_url": url, "receipt_file_id": fid,
                             "entered_by": driver, "logged_at": now_stamp(),
                             "visit_id": visit})
            append_many("Service", rows)
            reset_form("svc")
            what = jobs[0] if len(jobs) == 1 else f"{len(jobs)} services"
            flash(f"Saved {what}" + (" with receipt" if url else ""))
            st.rerun()

# ---- Timeline: everything done to the car ----------------------------------
with tab_time:
    if st.button("🔄 Refresh to see the latest entries", width="stretch", key="ref1"):
        load.clear()
        st.rerun()

    if not svc.empty:
        st.markdown("**Last time each service was done**")
        last = (svc.dropna(subset=["date"]).sort_values("date")
                .groupby("service_type", as_index=False).last())
        cur_odo = pd.concat([fuel["odometer"], svc["odometer"]]).max()
        last["ago"] = cur_odo - last["odometer"]
        st.dataframe(
            last.sort_values("date", ascending=False)[["service_type", "date", "odometer", "ago"]],
            hide_index=True, width="stretch", column_config={
                "service_type": "Service",
                "date": st.column_config.DateColumn("Last done", format="MMM D, YYYY"),
                "odometer": st.column_config.NumberColumn("At miles", format="%d"),
                "ago": st.column_config.NumberColumn("Miles since", format="%d"),
            })

    fm = fuel_with_mpg(fuel) if not fuel.empty else fuel.assign(mpg=None)
    items = []
    for _, r in fm.iterrows():
        bits = ([f"{r['gallons']:.3f} gal", f"${r['price_per_gallon']:.3f}/gal"]
                if pd.notna(r["gallons"]) and r["gallons"] else ["gallons not recorded"])
        if pd.notna(r["mpg"]):
            bits.append(f"{r['mpg']:.1f} mpg")
        if not r["full_tank"]:
            bits.append("partial fill")
        if r["missed_before"]:
            bits.append("after a missed fill-up")
        items.append({"kind": "Fuel", "date": r["date"], "odo": r["odometer"],
                      "title": "⛽ Fill-up", "amount": r["total_cost"], "detail": " · ".join(bits),
                      "notes": "", "receipt": "", "by": r["entered_by"]})
    # One card per shop visit, listing everything done
    if not svc.empty:
        sv = svc.assign(_visit=[visit_key(r) for r in svc.to_dict("records")])
        for _, g in sv.groupby("_visit", sort=False):
            first = g.iloc[0]
            notes = [n for n in dict.fromkeys(
                        n.split(" · Cost included with")[0] for n in g["notes"])
                     if n and not n.startswith("Cost included with")]
            receipts = [u for u in g["receipt_url"] if u]
            items.append({"kind": "Service", "date": first["date"], "odo": first["odometer"],
                          "title": "🔧 " + " + ".join(dict.fromkeys(g["service_type"])),
                          "amount": g["cost"].sum(min_count=1),
                          "detail": first["shop"], "notes": "\n\n".join(notes),
                          "receipt": receipts[0] if receipts else "",
                          "by": first["entered_by"]})

    if not items:
        st.caption("Nothing logged yet for this car.")
    else:
        tl = pd.DataFrame(items).sort_values(["date", "odo"], ascending=False, na_position="last")
        c1, c2 = st.columns(2)
        show = c1.radio("Show", ["Everything", "Service", "Fill-ups"], horizontal=True, key="tl_kind")
        years = sorted({d.year for d in tl["date"].dropna()}, reverse=True)
        yr = c2.selectbox("Year", ["All years"] + years, key="tl_year")
        if show == "Service":
            tl = tl[tl["kind"] == "Service"]
        elif show == "Fill-ups":
            tl = tl[tl["kind"] == "Fuel"]
        if yr != "All years":
            tl = tl[tl["date"].dt.year == yr]
        st.caption(f"{len(tl)} entries")
        for _, r in tl.iterrows():
            with st.container(border=True):
                a, b = st.columns([3, 1])
                a.markdown(f"**{r['title']}**")
                b.markdown(f"**{money(r['amount'])}**")
                odo = f" · {r['odo']:,.0f} mi" if pd.notna(r["odo"]) else ""
                extra = f" · {r['detail']}" if r["detail"] else ""
                st.caption(f"{fmt_date(r['date'])}{odo}{extra} · by {r['by']}")
                if r["notes"]:
                    st.write(r["notes"])
                if r["receipt"]:
                    st.link_button("📄 View receipt", receipt_link(r["receipt"]))

# ---- Maintenance schedule vs. odometer -------------------------------------
STATUS_ICON = {"Overdue": "🔴", "Due soon": "🟠", "OK": "🟢", "No record": "⚪"}
STATUS_COLOR = {"Overdue": "#D64545", "Due soon": "#E09A2B", "OK": "#3A9D6A", "No record": "#9AA5AC"}
SCHED_COLS = ["item", "every_miles", "every_months", "first_due_miles", "first_due_months",
              "counts_as", "notes"]

with tab_sched:
    if sched.empty:
        st.markdown("**Set up a maintenance schedule for this car**")
        st.caption("Start from a template, then adjust any interval. "
                   "Your owner's manual is the final word.")
        tpl_names = list(SCHEDULE_TEMPLATES)
        make_l = str(car_row["make"]).lower()
        guess = next((i for i, n in enumerate(tpl_names) if n.split()[0].lower() in make_l),
                     len(tpl_names) - 1)
        tpl = st.radio("Template", tpl_names, key=f"tpl_{car_id}", index=guess)
        if st.button("Use this schedule", type="primary", width="stretch"):
            save_schedule(car_id, [dict(zip(SCHED_COLS, t)) for t in SCHEDULE_TEMPLATES[tpl]])
            flash("Schedule added")
            st.rerun()
    else:
        odo_now = st.number_input(
            "Current mileage", min_value=0, step=1, key=f"odo_now_{car_id}",
            value=int(cur_odo) if pd.notna(cur_odo) and cur_odo else None,
            help="Taken from your latest fill-up or service. Type a newer reading to preview.")
        stt = schedule_status(sched, svc, odo_now if odo_now else float("nan"), date.today())

        counts = stt["status"].value_counts()
        st.markdown("  ·  ".join(f"{STATUS_ICON[s]} **{counts[s]}** {s.lower()}"
                                 for s in STATUS_ICON if s in counts))

        # Odometer timeline: where each item is due relative to where you are now
        chart_df = stt.dropna(subset=["next_mi"]).copy()
        if odo_now and not chart_df.empty:
            import altair as alt
            chart_df["start"] = chart_df["last_odo"].fillna(chart_df["next_mi"])
            chart_df["label"] = chart_df.apply(
                lambda r: f"{r['next_mi']:,.0f} mi", axis=1)
            lo = min(chart_df["start"].min(), odo_now) - 2000
            hi = max(chart_df["next_mi"].max(), odo_now) + 2000
            x = alt.X("start:Q", title="Odometer (miles)", scale=alt.Scale(domain=[lo, hi]),
                      axis=alt.Axis(format=",d"))
            y = alt.Y("item:N", title=None, sort=alt.SortField("next_mi"))
            color = alt.Color("status:N", legend=None, scale=alt.Scale(
                domain=list(STATUS_COLOR), range=list(STATUS_COLOR.values())))
            span = alt.Chart(chart_df).mark_rule(strokeWidth=3, opacity=.35).encode(
                x=x, x2="next_mi:Q", y=y, color=color)
            last_pt = alt.Chart(chart_df.dropna(subset=["last_odo"])).mark_tick(
                thickness=2, size=14).encode(x="last_odo:Q", y=y, color=color)
            due_pt = alt.Chart(chart_df).mark_circle(size=120, opacity=1).encode(
                x="next_mi:Q", y=y, color=color,
                tooltip=["item", "status", alt.Tooltip("next_mi:Q", title="Due at", format=",d")])
            now_rule = alt.Chart(pd.DataFrame({"x": [odo_now]})).mark_rule(
                strokeDash=[4, 3], strokeWidth=2).encode(x="x:Q")
            now_txt = alt.Chart(pd.DataFrame({"x": [odo_now], "t": ["You are here"]})).mark_text(
                align="left", dx=4, dy=-6, fontWeight="bold").encode(x="x:Q", y=alt.value(0), text="t")
            st.altair_chart((span + last_pt + due_pt + now_rule + now_txt).properties(
                height=max(160, 34 * len(chart_df))), width="stretch")
            st.caption("Tick = last done · dot = next due · dashed line = your odometer")

        order = {"Overdue": 0, "Due soon": 1, "OK": 2, "No record": 3}
        stt["_o"] = stt["status"].map(order)
        for _, r in stt.sort_values(["_o", "miles_left"], na_position="last").iterrows():
            with st.container(border=True):
                a, b = st.columns([3, 1])
                a.markdown(f"**{STATUS_ICON[r['status']]} {r['item']}**")
                b.markdown(f"**{r['status']}**")
                if r["progress"] is not None and pd.notna(r["progress"]):
                    st.progress(float(min(max(r["progress"], 0), 1)))
                every = " or ".join(x for x in (
                    f"{r['every_miles']:,.0f} mi" if pd.notna(r["every_miles"]) else "",
                    f"{r['every_months']:.0f} months" if pd.notna(r["every_months"]) else "") if x)
                lines = [f"Every {every}"] if every else []
                if r["status"] == "No record":
                    lines.append(f"Not logged yet. In the Service tab, log the last time this was "
                                 f"done (pick “{r['item']}”), even if it was a while ago.")
                    if r["next_mi"] is not None and pd.notna(r["next_mi"]):
                        lines.append(f"First due at {r['next_mi']:,.0f} mi.")
                else:
                    last_bits = [fmt_date(r["last_date"]) if r["last_date"] is not None else ""]
                    if r["last_odo"] is not None and pd.notna(r["last_odo"]):
                        last_bits.append(f"at {r['last_odo']:,.0f} mi")
                    lines.append("Last done " + " ".join(b for b in last_bits if b))
                    nxt = []
                    if r["next_mi"] is not None and pd.notna(r["next_mi"]):
                        ml = r["miles_left"]
                        if ml is not None and pd.notna(ml):
                            left = (f"{ml:,.0f} mi left" if ml > 0 else f"{-ml:,.0f} mi overdue")
                            nxt.append(f"at {r['next_mi']:,.0f} mi ({left})")
                        else:
                            nxt.append(f"at {r['next_mi']:,.0f} mi")
                    if r["next_date"] is not None:
                        dl = int(r["days_left"])
                        when = (f"{dl} days left" if dl > 0 else f"{-dl} days overdue")
                        nxt.append(f"by {fmt_date(r['next_date'])} ({when})")
                    if nxt:
                        lines.append("Next due " + ", or ".join(nxt))
                st.caption("  \n".join(lines))
                if r["notes"]:
                    st.caption(f"ℹ️ {r['notes']}")

        with st.expander("⚙️ Edit this car's schedule"):
            st.caption("Change intervals, add or remove rows. “Counts as” lists the service "
                       "names (separated by ;) that reset an item.")
            edited = st.data_editor(
                sched[SCHED_COLS].reset_index(drop=True), num_rows="dynamic", hide_index=True,
                key=f"sched_edit_{car_id}", width="stretch", column_config={
                    "item": "Item",
                    "every_miles": st.column_config.NumberColumn("Every (mi)", format="%d"),
                    "every_months": st.column_config.NumberColumn("Every (months)", format="%d"),
                    "first_due_miles": st.column_config.NumberColumn("First due (mi)", format="%d"),
                    "first_due_months": st.column_config.NumberColumn("First due (months)", format="%d"),
                    "counts_as": "Counts as", "notes": "Notes",
                })
            c1, c2 = st.columns(2)
            if c1.button("Save schedule", type="primary", width="stretch"):
                rows = [r for r in edited.to_dict("records") if str(r.get("item") or "").strip()]
                save_schedule(car_id, rows)
                flash("Schedule saved")
                st.rerun()
            if c2.button("Start over from a template", width="stretch"):
                save_schedule(car_id, [])
                flash("Schedule cleared. Pick a template.")
                st.rerun()

# ---- Trends ----------------------------------------------------------------
with tab_trend:
    today_ts = pd.Timestamp(date.today())

    # ---- Spending: monthly and yearly averages ----
    st.subheader("💵 Spending")
    scope = st.radio("Include", [car_name, "Both cars" if len(cars) == 2 else "All cars"],
                     horizontal=True, key="spend_scope")
    both = scope != car_name
    pv = monthly_spend(fuel_all if both else fuel, svc_all if both else svc)
    if pv.empty:
        st.caption("Spending appears after your first fill-up or service.")
    else:
        pv["Total"] = pv["Gas"] + pv["Service"]
        n_months = len(pv)
        avg_month = pv["Total"].mean()
        a, b = st.columns(2)
        a.metric("Average per month", money(avg_month),
                 help=f"Gas {money(pv['Gas'].mean())} + service {money(pv['Service'].mean())}")
        b.metric("Average per year", money(avg_month * 12))
        c, d = st.columns(2)
        c.metric(f"{today_ts.year} so far", money(pv.loc[pv.index.year == today_ts.year, "Total"].sum()))
        d.metric("Last 12 months", money(pv["Total"].iloc[-12:].sum()))
        st.caption(f"Averages cover {n_months} months, {pv.index.min().strftime('%b %Y')} to now. "
                   "Months with nothing logged count as $0, so unlogged fill-ups pull them down.")

        yearly = pv.groupby(pv.index.year).agg(Gas=("Gas", "sum"), Service=("Service", "sum"),
                                                Total=("Total", "sum"), Months=("Total", "size"))
        yearly["Per month"] = yearly["Total"] / yearly["Months"]
        yearly.index = yearly.index.astype(str)
        st.markdown("**By year**")
        st.bar_chart(yearly[["Gas", "Service"]], height=220)
        st.dataframe(
            yearly.sort_index(ascending=False).reset_index(names="Year")
            [["Year", "Gas", "Service", "Total", "Per month"]],
            hide_index=True, width="stretch", column_config={
                c: st.column_config.NumberColumn(c, format="$%.2f")
                for c in ("Gas", "Service", "Total", "Per month")})
        st.caption("Per month for the current year counts only the months so far.")

        st.markdown("**Last 24 months**")
        recent = pv[["Gas", "Service"]].iloc[-24:]
        recent.index = recent.index.to_timestamp()
        st.bar_chart(recent, height=220)

    # ---- Fill-up frequency (this car) ----
    st.subheader("⛽ How often you fill up")
    fs = fuel.dropna(subset=["date"]).sort_values(["date", "odometer"])
    if len(fs) < 2:
        st.caption("Needs at least two fill-ups.")
    else:
        start = max(today_ts - pd.Timedelta(days=365), fs["date"].min())
        rec = fs[fs["date"] >= start]
        window = "last 12 months" if start > fs["date"].min() else "since your first fill-up"
        months_span = max((today_ts - start).days / 30.44, 1)
        gaps_d = rec["date"].diff().dt.days
        gaps_mi = rec["odometer"].diff()
        ok = rec["missed_before"].eq(False) & gaps_d.notna()
        typ_days = gaps_d[ok & (gaps_d > 0)].median()  # skip same-day top-ups
        typ_miles = gaps_mi[ok & (gaps_mi > 0)].median()
        a, b = st.columns(2)
        a.metric("Fill-ups per month", f"{len(rec) / months_span:.1f}")
        b.metric("Typical days between", f"{typ_days:.0f}" if pd.notna(typ_days) else "--")
        c, d = st.columns(2)
        c.metric("Typical miles between", f"{typ_miles:,.0f}" if pd.notna(typ_miles) else "--")
        c_days = (today_ts - fs["date"].max()).days
        d.metric("Days since last fill-up", f"{c_days}")
        st.caption(f"Based on {len(rec)} fill-ups, {window}. 'Typical' is the middle value, "
                   "so one long road trip doesn't skew it. Same-day top-ups and gaps after a missed fill-up are skipped.")
        per_month = fs.set_index("date").resample("MS").size()
        per_month = per_month.reindex(pd.date_range(per_month.index.min(),
                                                    today_ts.to_period("M").to_timestamp(),
                                                    freq="MS"), fill_value=0).iloc[-24:]
        st.markdown("**Fill-ups per month (last 24 months)**")
        st.bar_chart(per_month.rename("Fill-ups"), height=200)

    # ---- MPG and gas price ----
    if not fuel.empty:
        st.subheader("📈 MPG and gas price")
        f = fuel_with_mpg(fuel).sort_values("date")
        if f["mpg"].notna().sum() >= 2:
            st.markdown("**MPG over time**")
            st.line_chart(f.dropna(subset=["mpg"]).set_index("date")["mpg"], height=220)
        st.markdown("**Price per gallon**")
        st.line_chart(f.dropna(subset=["price_per_gallon"]).set_index("date")["price_per_gallon"],
                      height=220)

# ---- Fix or delete an entry ------------------------------------------------
with tab_fix:
    st.caption("Correct a typo, add a receipt to an older service, or delete an entry.")
    with st.expander("🩺 Check receipts connection"):
        if not receipts_enabled():
            st.warning("Receipts are off: `receipts_url` or `receipts_token` is missing from "
                       "your Streamlit secrets, so the upload box is hidden.")
        elif st.button("Run check", key="rcheck"):
            try:
                r = requests.post(st.secrets["receipts_url"], timeout=60,
                                  json={"action": "ping", "token": st.secrets["receipts_token"]})
                try:
                    res = r.json()
                except ValueError:
                    res = None
                if res is None:
                    st.error(f"The receipts script answered with a web page, not data "
                             f"(status {r.status_code}). Usually the Web app's 'Who has access' "
                             f"isn't set to Anyone, or the URL isn't the /exec one.")
                    st.code(r.text[:400])
                elif res.get("error") == "Not allowed":
                    st.error("Connected, but the secret phrase doesn't match. Make "
                             "`receipts_token` in secrets identical to TOKEN in the script.")
                elif res.get("error") == "Unknown action":
                    st.success("✅ Receipts connection works. The script is reachable and the "
                               "secret phrase matches.")
                else:
                    st.warning(f"Unexpected answer: {res}")
            except Exception as e:
                st.error(f"Couldn't reach the receipts script: {e}")
    options = {}
    for _, r in svc.sort_values("date", ascending=False).iterrows():
        tag = " 📄" if r["receipt_url"] else ""
        options[("Service", r["entry_id"])] = f"🔧 {fmt_date(r['date'])} · {r['service_type']}{tag}"
    for _, r in fuel.sort_values("date", ascending=False).iterrows():
        options[("Fuel", r["entry_id"])] = (
            f"⛽ {fmt_date(r['date'])} · {r['odometer']:,.0f} mi · {money(r['total_cost'])}")
    if not options:
        st.caption("Nothing to fix yet.")
    else:
        kind, eid = st.selectbox("Entry", list(options), format_func=options.get, key="fix_pick")
        src = fuel if kind == "Fuel" else svc
        r = src[src["entry_id"] == eid].iloc[0].to_dict()
        k = f"fix_{eid}_{st.session_state.get('_v_fix', 0)}"
        d0 = r["date"].date() if pd.notna(r["date"]) else date.today()

        with st.form(k):
            c1, c2 = st.columns(2)
            e_date = c1.date_input("Date", value=d0, max_value=date.today())
            e_odo = c2.number_input("Mileage", min_value=0, step=1,
                                    value=int(r["odometer"]) if pd.notna(r["odometer"]) else None)
            if kind == "Fuel":
                c3, c4 = st.columns(2)
                e_cost = c3.number_input("Total cost ($)", min_value=0.0, step=0.01, format="%.2f",
                                         value=float(r["total_cost"]) if pd.notna(r["total_cost"]) else None)
                e_gal = c4.number_input("Gallons", min_value=0.0, step=0.001, format="%.3f",
                                        value=float(r["gallons"]) if pd.notna(r["gallons"]) else None)
                e_full = st.checkbox("Filled the tank all the way", value=bool(r["full_tank"]))
                e_missed = st.checkbox("I missed logging a fill-up before this one",
                                       value=bool(r["missed_before"]))
            else:
                e_type = st.text_input("Service", value=r["service_type"])
                c3, c4 = st.columns(2)
                e_cost = c3.number_input("Cost ($)", min_value=0.0, step=0.01, format="%.2f",
                                         value=float(r["cost"]) if pd.notna(r["cost"]) else None)
                e_shop = c4.text_input("Shop", value=r["shop"])
                e_notes = st.text_area("Notes", value=r["notes"])
                e_file = None
                if r["receipt_url"]:
                    st.markdown(f"Current receipt: [open]({receipt_link(r['receipt_url'])})")
                mates_n = sum(1 for m in svc.to_dict("records")
                              if m["entry_id"] != eid and visit_key(m) == visit_key(r))
                if mates_n:
                    st.caption(f"Part of a visit with {mates_n} other item"
                               f"{'s' if mates_n > 1 else ''}. A new receipt is added to all of them.")
                if receipts_enabled():
                    e_file = st.file_uploader(
                        "Replace receipt" if r["receipt_url"] else "Add receipt",
                        type=RECEIPT_TYPES)
            saved = st.form_submit_button("Save changes", width="stretch", type="primary")

        if saved:
            try:
                if kind == "Fuel":
                    if not e_odo or not e_cost or not e_gal:
                        st.error("Mileage, total cost and gallons are all needed.")
                        st.stop()
                    update("Fuel", eid, {**r, "date": e_date.isoformat(), "odometer": int(e_odo),
                                         "gallons": round(e_gal, 3), "total_cost": round(e_cost, 2),
                                         "price_per_gallon": round(e_cost / e_gal, 3),
                                         "full_tank": e_full, "missed_before": e_missed})
                else:
                    url, fid = r["receipt_url"], r["receipt_file_id"]
                    if e_file:
                        with st.spinner("Uploading receipt..."):
                            url, fid = upload_receipt(e_file, f"{e_date.isoformat()} {car_name} {e_type}")
                        # Same receipt for everything from this visit
                        vk = visit_key(r)
                        mates = [m for m in svc.to_dict("records")
                                 if m["entry_id"] != eid and visit_key(m) == vk]
                        for m in mates:
                            update("Service", m["entry_id"], {**m, "receipt_url": url,
                                                              "receipt_file_id": fid})
                        changed = {eid, *[m["entry_id"] for m in mates]}
                        if not receipt_shared(r["receipt_file_id"], changed):
                            trash_receipt(r["receipt_file_id"])
                    update("Service", eid, {**r, "date": e_date.isoformat(),
                                            "odometer": int(e_odo) if e_odo else "",
                                            "service_type": e_type.strip() or r["service_type"],
                                            "cost": round(e_cost, 2) if e_cost is not None else "",
                                            "shop": e_shop.strip(), "notes": e_notes.strip(),
                                            "receipt_url": url, "receipt_file_id": fid})
                reset_form("fix")
                flash("Changes saved")
                st.rerun()
            except Exception as e:
                st.error(f"Couldn't save ({e}). Try again.")

        with st.expander("🗑️ Delete this entry"):
            sure = st.checkbox("Yes, delete it for good", key=f"del_{eid}")
            if st.button("Delete", disabled=not sure, key=f"delbtn_{eid}"):
                if kind == "Service" and not receipt_shared(r["receipt_file_id"], {eid}):
                    trash_receipt(r["receipt_file_id"])
                delete(kind, eid)
                flash("Entry deleted")
                st.rerun()

st.caption(f"Stored in your Google Sheet · loaded {datetime.now():%I:%M %p}")
if me:
    st.button(f"Sign out ({me})", on_click=st.logout, key="signout")
