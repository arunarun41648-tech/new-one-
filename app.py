# ╔══════════════════════════════════════════════════════════════════════════════╗
# ║  Garlic Order & Delivery Platform  —  app.py  FINAL v9                     ║
# ║  ROUTE FILE UPLOAD: Admin uploads CSV/Excel → preview → create trip        ║
# ║  GPS FIX: Single "Get My Location" button → auto-fills Lat/Long            ║
# ╚══════════════════════════════════════════════════════════════════════════════╝
import os, uuid, textwrap, io
from datetime import datetime, date

import streamlit as st
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials
from streamlit_js_eval import get_geolocation

# ── Page config ───────────────────────────────────────────────────────────────
st.set_page_config(
    page_title="Garlic Order & Delivery",
    page_icon="🧄", layout="wide",
    initial_sidebar_state="collapsed"
)

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Syne:wght@400;600;700;800&family=DM+Sans:wght@300;400;500&display=swap');
:root{
  --green:#1a7f4b; --amber:#854f0b; --blue:#185fa5;
  --border:#c8e6d4; --bg:#eef5f0; --text:#1a2e22; --muted:#5a7a65;
}
html,body,[class*="css"]{ font-family:'DM Sans',sans-serif; color:var(--text); }
h1,h2,h3{ font-family:'Syne',sans-serif; }
.stApp{ background:var(--bg); }
header[data-testid="stHeader"]{ background:transparent; }
.sl{
  font-family:'Syne',sans-serif; font-weight:700; font-size:.75rem;
  letter-spacing:.8px; text-transform:uppercase; color:var(--green);
  padding-bottom:.4rem; border-bottom:2px solid var(--border); margin-bottom:.9rem;
}
.sl-amber{ color:var(--amber); border-color:#f5d6a7; }
.sl-blue { color:var(--blue);  border-color:#b5d4f4; }
.pill{ display:inline-block; font-size:.75rem; padding:3px 12px; border-radius:20px; font-weight:600; }
.pill-pend{ background:#fff3cd; color:#856404; }
.pill-done{ background:#d4edda; color:#1a7f4b; }
.pill-fail{ background:#f8d7da; color:#842029; }
.pill-part{ background:#cce5ff; color:#004085; }
.pill-on  { background:#d4edda; color:#1a7f4b; }
.pill-off { background:#e2e3e5; color:#383d41; }
.map-frame{ border-radius:12px; overflow:hidden; border:2px solid var(--border); margin-top:.5rem; }
.total-box{
  background:#e1f5ee; border:2px solid #1a7f4b; border-radius:12px;
  padding:14px 18px; margin:10px 0; font-size:1.1rem; font-weight:700; color:#0d1f14;
}
.loc-box{
  background:#e6f1fb; border:1px solid #b5d4f4; border-radius:10px;
  padding:10px 14px; margin:6px 0; font-size:13px;
}
.gps-box{
  background:#e8f5e9; border:2px solid #1a7f4b; border-radius:12px;
  padding:12px 16px; margin:8px 0; font-size:14px; font-weight:600; color:#0d1f14;
}
.route-card{
  background:#fff; border:1.5px solid var(--border); border-radius:14px;
  padding:16px 20px; margin:10px 0;
  box-shadow:0 2px 8px rgba(26,127,75,.07);
}
.route-card-header{
  font-family:'Syne',sans-serif; font-weight:700; font-size:1rem;
  color:#0d1f14; margin-bottom:8px;
}
.route-row{
  display:flex; align-items:center; gap:10px;
  border-bottom:1px solid #eee; padding:6px 0; font-size:.88rem;
}
.route-row:last-child{ border-bottom:none; }
.stop-badge{
  background:#1a7f4b; color:#fff; border-radius:50%;
  width:24px; height:24px; display:inline-flex;
  align-items:center; justify-content:center;
  font-size:.75rem; font-weight:700; flex-shrink:0;
}
.upload-zone{
  background:#f4faf7; border:2px dashed #1a7f4b; border-radius:14px;
  padding:24px; text-align:center; margin:10px 0;
}
.template-box{
  background:#fffbf0; border:1.5px solid #f5d6a7; border-radius:10px;
  padding:12px 16px; margin:8px 0; font-size:.85rem; color:#6b4a00;
}
.warn-row{ background:#fff8e1; border-radius:8px; padding:6px 10px; margin:4px 0; font-size:.83rem; }
.ok-row  { background:#e8f5e9; border-radius:8px; padding:6px 10px; margin:4px 0; font-size:.83rem; }
div[data-testid="stTextInput"] input,
div[data-testid="stSelectbox"] select,
div[data-testid="stNumberInput"] input,
div[data-testid="stTextArea"] textarea{
  border-radius:10px !important; border-color:var(--border) !important;
}
.stButton>button{
  border-radius:12px !important; font-family:'Syne',sans-serif !important; font-weight:700 !important;
}
.stButton>button[kind="primary"]{
  background:var(--green) !important; border:none !important; color:#fff !important;
}
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════════════════════════════
#  SESSION STATE
# ═══════════════════════════════════════════════════════════════════════════════
DEFAULTS = {
    "logged_in": False, "user": None,
    "driver_id": None,  "driver_active": True,
    "active_stop": 0,   "cust_data": {},
    "task_done": False,  "_gps_requested": False,
    "route_parsed_df": None, "route_file_name": None,
}
for _k, _v in DEFAULTS.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v
if "_geo_lat" not in st.session_state: st.session_state["_geo_lat"] = ""
if "_geo_lng" not in st.session_state: st.session_state["_geo_lng"] = ""

try:
    ADMIN_REGISTER_PASSWORD = st.secrets.get("admin_register_password", "Admin@123")
except Exception:
    ADMIN_REGISTER_PASSWORD = "Admin@123"

# ═══════════════════════════════════════════════════════════════════════════════
#  GOOGLE AUTH
# ═══════════════════════════════════════════════════════════════════════════════
SCOPES = [
    "https://spreadsheets.google.com/feeds",
    "https://www.googleapis.com/auth/drive",
]
SPREADSHEET_NAME = "Garlic_Order & Delivery Project"

def _clean_private_key(raw: str) -> str:
    k = str(raw).strip()
    if (k.startswith('"') and k.endswith('"')) or (k.startswith("'") and k.endswith("'")):
        k = k[1:-1]
    k = (k.replace("\\r\\n","\n").replace("\\r","\n")
          .replace("\\n","\n").replace("\r\n","\n").replace("\r","\n"))
    h = "-----BEGIN PRIVATE KEY-----"; f = "-----END PRIVATE KEY-----"
    k = k.replace(h,"").replace(f,"").replace("\n","").replace(" ","").strip()
    if len(k) < 100: raise ValueError(f"Private key too short ({len(k)} chars).")
    return f"{h}\n" + "\n".join(textwrap.wrap(k, 64)) + f"\n{f}\n"

def _get_creds() -> Credentials:
    cp = os.path.join(os.path.dirname(os.path.abspath(__file__)), "credentials.json")
    if os.path.exists(cp):
        return Credentials.from_service_account_file(cp, scopes=SCOPES)
    try:
        raw = dict(st.secrets["gcp_service_account"])
        raw["private_key"] = _clean_private_key(str(raw["private_key"]))
        return Credentials.from_service_account_info(raw, scopes=SCOPES)
    except KeyError: pass
    except Exception as e: raise ValueError(f"Secrets error: {e}")
    raise ValueError("No credentials found. Add [gcp_service_account] to Streamlit Secrets.")

@st.cache_resource(show_spinner=False)
def _cached_client():
    return gspread.authorize(_get_creds())

def get_gspread_client():
    try: return _cached_client()
    except Exception:
        _cached_client.clear()
        return gspread.authorize(_get_creds())

# ═══════════════════════════════════════════════════════════════════════════════
#  SHEET CONFIG
# ═══════════════════════════════════════════════════════════════════════════════
TAB = {
    "base":             "Base",
    "customer_onboard": "Customer Onboard Data",
    "driver_onboard":   "Driver Onboard Data",
    "sales_exec":       "sales executive",
    "delivery_driver":  "delivery Driver",
    "user_registry":    "UserRegistry",
    "admin_log":        "Admin Log",
    "skus":             "SKU Master",
    "trips":            "Trips",
}

HEADERS = {
    "base": [
        "Order ID","SOID","City","ORDER DATE","DELIVERED DATE","ORDERED TIME",
        "CustomerId","Customer shop name","Customer Number","Customer_Classification",
        "sales executive","sales executive Number",
        "SKU","SKU Name","WeightType","Price","OrderedQty","OrderTotal",
        "ReturnQty","Reason","return_updated_role",
        "Tripid","Transport","ShopOpeningFrom","ShopReachTime","DeliveryCutOff",
        "Shop Location","Latitude","Longitude",
        "Delivery Status","EnteredBy_UID","Timestamp",
    ],
    "customer_onboard": [
        "CUST-ID","Full Name","Mobile","Email","Shop Name","Shop Address",
        "City","Classification","Latitude","Longitude",
        "Onboarded By","Onboard Date","Status",
    ],
    "driver_onboard": [
        "Driver ID","Full Name","Mobile","Email","Vehicle Type","Vehicle Number",
        "Bank Name","Account Number","IFSC Code","UPI ID",
        "Onboard Date","Active Status","Last Active",
    ],
    "user_registry":   ["UID","Full Name","Phone","Email","Role","Password","Created At","Status"],
    "sales_exec":      ["UID","Full Name","Phone","Email","Role","Password","Created At"],
    "delivery_driver": ["UID","Full Name","Phone","Email","Role","Password","Created At"],
    "admin_log": [
        "Log ID","Timestamp","Admin UID","Mail ID","Action Type",
        "Entity","Entity ID","Old Value","New Value","Notes",
    ],
    "skus":  ["SKU Code","SKU Name","Price","Weight Type","Category","Active","Created By","Created At"],
    "trips": ["Trip ID","Date","City","Shops","Driver UID","Driver Name","Status","Created By","Created At"],
}

# ── Sheet helpers ─────────────────────────────────────────────────────────────
def open_spreadsheet():
    client = get_gspread_client()
    try:
        return client.open(SPREADSHEET_NAME)
    except gspread.SpreadsheetNotFound:
        st.error(f'❌ Sheet not found: "{SPREADSHEET_NAME}"')
        st.stop()

def get_ws(key: str):
    sp = open_spreadsheet(); name = TAB[key]
    try:
        ws = sp.worksheet(name)
    except gspread.WorksheetNotFound:
        ws = sp.add_worksheet(title=name, rows=2000, cols=50)
        if key in HEADERS: ws.append_row(HEADERS[key])
        return ws
    if key in HEADERS:
        expected = HEADERS[key]
        current  = ws.row_values(1)
        if not current:
            ws.append_row(expected)
        else:
            missing = [(i,c) for i,c in enumerate(expected) if c not in current]
            for exp_idx, col_name in missing:
                ws.insert_cols([[col_name]], col=exp_idx+1)
                current = ws.row_values(1)
    return ws

def read_sheet(key: str) -> pd.DataFrame:
    try:
        rows = get_ws(key).get_all_records()
        return pd.DataFrame(rows) if rows else pd.DataFrame(columns=HEADERS.get(key,[]))
    except Exception as e:
        st.error(f"Sheet read error ({key}): {e}")
        return pd.DataFrame(columns=HEADERS.get(key,[]))

def append_row(key: str, row: list):
    ws = get_ws(key)
    if key not in HEADERS:
        ws.append_row(row, value_input_option="USER_ENTERED"); return
    expected = HEADERS[key]
    if len(row) != len(expected):
        ws.append_row(row, value_input_option="USER_ENTERED"); return
    data_dict    = dict(zip(expected, row))
    live_headers = ws.row_values(1)
    ordered_row  = [data_dict.get(h,"") for h in live_headers]
    ws.append_row(ordered_row, value_input_option="USER_ENTERED")

def update_row(key: str, id_col: str, id_val: str, updates: dict) -> bool:
    ws = get_ws(key); headers = ws.row_values(1)
    for i, row in enumerate(ws.get_all_records(), start=2):
        if str(row.get(id_col,"")).strip() == str(id_val).strip():
            for col, val in updates.items():
                if col in headers:
                    ws.update_cell(i, headers.index(col)+1, val)
            return True
    return False

def find_row(key: str, col: str, val: str):
    df = read_sheet(key)
    if df.empty or col not in df.columns: return None
    m = df[df[col].astype(str).str.strip() == str(val).strip()]
    return m.iloc[0].to_dict() if not m.empty else None

def col_exists(key: str, col: str, val: str) -> bool:
    return find_row(key, col, val) is not None

@st.cache_data(ttl=120)
def load_customers() -> pd.DataFrame: return read_sheet("customer_onboard")

@st.cache_data(ttl=60)
def load_skus() -> pd.DataFrame: return read_sheet("skus")

def active_skus() -> pd.DataFrame:
    df = load_skus()
    if df.empty: return df
    return df[df["Active"].astype(str).str.lower() == "true"]

def all_drivers() -> pd.DataFrame: return read_sheet("driver_onboard")

def active_drivers() -> pd.DataFrame:
    df = all_drivers()
    if df.empty: return df
    return df[df["Active Status"].astype(str).str.lower() == "active"]

def set_driver_status(driver_id: str, status: str):
    ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    update_row("driver_onboard","Driver ID",driver_id,{"Active Status":status,"Last Active":ts})

def get_driver_trip(driver_uid: str):
    df = read_sheet("trips")
    if df.empty: return None
    m = df[(df["Driver UID"].astype(str)==str(driver_uid)) &
           (df["Status"].astype(str).str.lower().isin(["assigned","in progress"]))]
    return m.iloc[0].to_dict() if not m.empty else None

def write_admin_log(admin_uid, mail_id, action, entity, entity_id, old="", new="", notes=""):
    lid = "LOG-"+uuid.uuid4().hex[:6].upper()
    ts  = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    append_row("admin_log",[lid,ts,admin_uid,mail_id,action,entity,
                             str(entity_id),str(old),str(new),notes])

# ═══════════════════════════════════════════════════════════════════════════════
#  ROUTE FILE HELPERS
#  Format: SaleOrderId | DeliveryDate | CustomerId | Customer | Slot | Driver |
#          TotalCrates | OrderKg | Latitude | Longitude | FC Latitude | FC Longitude
# ═══════════════════════════════════════════════════════════════════════════════

ROUTE_REQUIRED_COLS = [
    "SaleOrderId", "DeliveryDate", "CustomerId", "Customer",
    "Latitude", "Longitude",
]
ROUTE_OPTIONAL_COLS = ["Slot", "Driver", "TotalCrates", "OrderKg", "FC Latitude", "FC Longitude"]
ROUTE_ALL_COLS      = ROUTE_REQUIRED_COLS + ROUTE_OPTIONAL_COLS

ROUTE_TEMPLATE_CSV = (
    "SaleOrderId\tDeliveryDate\tCustomerId\tCustomer\tSlot\tDriver\t"
    "TotalCrates\tOrderKg\tLatitude\tLongitude\tFC Latitude\tFC Longitude\n"
    "SO-10001\t2025-06-01\tCUST-AABBCC\tSri Lakshmi Provision\tMorning\tRavi Kumar\t"
    "4\t12.5\t12.971599\t77.594566\t12.9352\t77.6245\n"
    "SO-10002\t2025-06-01\tCUST-DDEEFF\tHotel Majestic\tAfternoon\tRavi Kumar\t"
    "2\t6.0\t12.295810\t76.639380\t12.9352\t77.6245\n"
)


def parse_route_file(uploaded_file) -> tuple:
    """
    Parse uploaded CSV / TSV / Excel route file with the standard column format.
    Returns (df, errors_list, warnings_list)
    """
    errors = []; warnings = []
    fname  = uploaded_file.name.lower()
    try:
        if fname.endswith(".csv"):
            # Try tab-separated first (common for route exports), fall back to comma
            raw = uploaded_file.read()
            uploaded_file.seek(0)
            sample = raw[:2048].decode("utf-8", errors="replace")
            sep = "\t" if sample.count("\t") >= sample.count(",") else ","
            df = pd.read_csv(uploaded_file, sep=sep, dtype=str)
        elif fname.endswith(".tsv"):
            df = pd.read_csv(uploaded_file, sep="\t", dtype=str)
        elif fname.endswith((".xlsx", ".xls")):
            df = pd.read_excel(uploaded_file, dtype=str)
        else:
            return None, ["Unsupported file type. Upload a .csv, .tsv or .xlsx file."], []
    except Exception as e:
        return None, [f"Could not read file: {e}"], []

    # Normalise column names — strip whitespace
    df.columns = [str(c).strip() for c in df.columns]

    # Check required columns
    missing = [c for c in ROUTE_REQUIRED_COLS if c not in df.columns]
    if missing:
        errors.append(
            f"Missing required column(s): **{', '.join(missing)}**  |  "
            f"Found: {', '.join(df.columns.tolist())}"
        )
        return None, errors, warnings

    # Drop fully empty rows
    df = df.dropna(how="all").reset_index(drop=True)
    if df.empty:
        errors.append("File contains no data rows.")
        return None, errors, warnings

    # Normalise values — strip whitespace in all string cells
    for col in df.columns:
        df[col] = df[col].astype(str).str.strip()

    # Blank CustomerId rows
    blank_cid = df[df["CustomerId"].str.strip().isin(["", "nan", "None"])]
    if not blank_cid.empty:
        warnings.append(f"{len(blank_cid)} row(s) have a blank CustomerId and will be skipped.")
        df = df[~df["CustomerId"].str.strip().isin(["", "nan", "None"])].reset_index(drop=True)

    # Blank SaleOrderId rows
    blank_soid = df[df["SaleOrderId"].str.strip().isin(["", "nan", "None"])]
    if not blank_soid.empty:
        warnings.append(f"{len(blank_soid)} row(s) have a blank SaleOrderId and will be skipped.")
        df = df[~df["SaleOrderId"].str.strip().isin(["", "nan", "None"])].reset_index(drop=True)

    # Validate Latitude / Longitude are numeric
    for coord_col in ("Latitude", "Longitude"):
        try:
            df[coord_col] = pd.to_numeric(df[coord_col], errors="coerce")
            bad = df[df[coord_col].isna()]
            if not bad.empty:
                warnings.append(f"{len(bad)} row(s) have invalid {coord_col} values.")
        except Exception:
            warnings.append(f"Could not validate {coord_col} column.")

    # Numeric casts for optional numeric cols
    for num_col in ("TotalCrates", "OrderKg"):
        if num_col in df.columns:
            df[num_col] = pd.to_numeric(df[num_col], errors="coerce")

    # Duplicate SaleOrderId check
    dupes = df[df.duplicated("SaleOrderId", keep=False)]
    if not dupes.empty:
        warnings.append(
            f"{len(dupes)} row(s) share a duplicate SaleOrderId — "
            f"all will be included but please verify."
        )

    return df, errors, warnings


def validate_route_against_customers(route_df: pd.DataFrame, custs_df: pd.DataFrame):
    """
    Cross-check CustomerId in route file against customer_onboard sheet (CUST-ID column).
    Rows that DON'T match are still kept (route may contain IDs not yet onboarded),
    but we flag them so the admin can decide.
    Returns (df_with_match_flag, unmatched_ids_list)
    """
    if custs_df.empty:
        route_df = route_df.copy()
        route_df["_matched"] = False
        return route_df, route_df["CustomerId"].tolist()

    valid_ids = set(custs_df["CUST-ID"].astype(str).str.strip())
    route_ids = route_df["CustomerId"].astype(str).str.strip()
    route_df  = route_df.copy()
    route_df["_matched"] = route_ids.isin(valid_ids)
    unmatched = route_ids[~route_ids.isin(valid_ids)].unique().tolist()
    return route_df, unmatched


def enrich_route_row(row: pd.Series, custs_df: pd.DataFrame) -> dict:
    """Pull extra info (shop address, classification) from customer master if available."""
    out = {"address": "", "classification": "", "city": ""}
    if custs_df.empty: return out
    cid = str(row.get("CustomerId","")).strip()
    m   = custs_df[custs_df["CUST-ID"] == cid]
    if not m.empty:
        r = m.iloc[0]
        out["address"]        = str(r.get("Shop Address",""))
        out["classification"] = str(r.get("Classification",""))
        out["city"]           = str(r.get("City",""))
    return out


_ROUTE_TEMPLATE_ROWS = [
    {
        "SaleOrderId": "SO-10001", "DeliveryDate": "2025-06-01",
        "CustomerId": "CUST-AABBCC", "Customer": "Sri Lakshmi Provision",
        "Slot": "Morning", "Driver": "Ravi Kumar",
        "TotalCrates": 4, "OrderKg": 12.5,
        "Latitude": 12.971599, "Longitude": 77.594566,
        "FC Latitude": 12.9352, "FC Longitude": 77.6245,
    },
    {
        "SaleOrderId": "SO-10002", "DeliveryDate": "2025-06-01",
        "CustomerId": "CUST-DDEEFF", "Customer": "Hotel Majestic",
        "Slot": "Afternoon", "Driver": "Ravi Kumar",
        "TotalCrates": 2, "OrderKg": 6.0,
        "Latitude": 12.295810, "Longitude": 76.639380,
        "FC Latitude": 12.9352, "FC Longitude": 77.6245,
    },
]

def generate_route_template_excel() -> bytes | None:
    """
    Generate a downloadable Excel template for route uploads.
    Returns bytes on success, None if openpyxl is not available.
    """
    try:
        import openpyxl  # noqa: F401  — check availability first
        buf = io.BytesIO()
        df  = pd.DataFrame(_ROUTE_TEMPLATE_ROWS)
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False, sheet_name="Route")
        buf.seek(0)
        return buf.read()
    except Exception:
        return None

# ═══════════════════════════════════════════════════════════════════════════════
#  AUTH
# ═══════════════════════════════════════════════════════════════════════════════
def gen_uid(role):
    p={"admin":"ADMIN","sales executive":"SE","delivery Driver":"DD"}.get(role,"USR")
    return f"{p}-{uuid.uuid4().hex[:6].upper()}"
def gen_cust_id():   return f"CUST-{uuid.uuid4().hex[:6].upper()}"
def gen_driver_id(): return f"DD-{uuid.uuid4().hex[:6].upper()}"
def gen_order_id():  return f"ORD-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:4].upper()}"
def gen_trip_id():   return f"TRP-{datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:4].upper()}"

def register_user(name, phone, email, role, password):
    email = email.strip().lower()
    if col_exists("user_registry","Email",email):
        ex = find_row("user_registry","Email",email)
        return None, f"Email already registered. UID: {ex['UID']}"
    uid = gen_uid(role); ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    append_row("user_registry",[uid,name,phone,email,role,password,ts,"Active"])
    if role in ("sales executive","delivery Driver"):
        rk = "sales_exec" if role=="sales executive" else "delivery_driver"
        append_row(rk,[uid,name,phone,email,role,password,ts])
    return uid, None

def login_user(email, password):
    email = email.strip().lower()
    user  = find_row("user_registry","Email",email)
    if not user: return None,"Email not found."
    stored = str(user.get("Password","") or user.get("Password Hash",""))
    if stored != str(password): return None,"Incorrect password."
    if str(user.get("Status","")).lower() != "active": return None,"Account inactive."
    return {"uid":user["UID"],"name":user["Full Name"],"role":user["Role"],
            "phone":str(user.get("Phone","")),"email":str(user.get("Email",""))}, None

# ═══════════════════════════════════════════════════════════════════════════════
#  UI HELPERS
# ═══════════════════════════════════════════════════════════════════════════════
def sl(label, color=""):
    cls = f"sl sl-{color}" if color else "sl"
    return f'<div class="{cls}">{label}</div>'

def pill(text, cls="pill-pend"):
    return f'<span class="pill {cls}">{text}</span>'

def map_embed(lat, lng, height=260):
    if not lat or not lng: return ""
    return (f'<div class="map-frame"><iframe width="100%" height="{height}"'
            f' frameborder="0" style="border:0;display:block" allowfullscreen'
            f' src="https://maps.google.com/maps?q={lat},{lng}&output=embed&z=16">'
            f'</iframe></div>')

def gps_capture_component():
    cur_lat = st.session_state.get("_geo_lat", "")
    cur_lng = st.session_state.get("_geo_lng", "")
    if st.button("📍 Get My Current Location", type="primary", key="gps_loc_btn"):
        st.session_state["_gps_requested"] = True
        st.rerun()
    if st.session_state.get("_gps_requested", False) and not (cur_lat and cur_lng):
        st.info("⏳ Requesting location access from your browser…")
        loc = get_geolocation()
        if loc and isinstance(loc, dict) and "coords" in loc:
            c = loc["coords"]
            st.session_state["_geo_lat"] = f"{float(c['latitude']):.6f}"
            st.session_state["_geo_lng"] = f"{float(c['longitude']):.6f}"
            st.session_state["_gps_requested"] = False
            cur_lat = st.session_state["_geo_lat"]
            cur_lng = st.session_state["_geo_lng"]
            st.rerun()
    return cur_lat, cur_lng

def topbar(role_label, role_color="#1a7f4b"):
    user = st.session_state.user
    c1,c2,c3 = st.columns([5,3,2])
    with c1:
        st.markdown(
            '<div style="display:flex;align-items:center;gap:10px;padding:6px 0">'
            '<span style="font-size:1.6rem">🧄</span>'
            '<span style="font-family:Syne,sans-serif;font-weight:800;'
            'font-size:1.15rem;color:#1a7f4b">Garlic Order & Delivery</span>'
            '</div>', unsafe_allow_html=True)
    with c2:
        st.markdown(
            f'<div style="text-align:center;padding-top:8px">'
            f'<span style="background:{role_color};color:#fff;padding:4px 14px;'
            f'border-radius:20px;font-size:.8rem;font-weight:700">{role_label}</span>'
            f'&nbsp;<code style="font-size:.72rem;color:#5a7a65">{user["uid"]}</code>'
            f'</div>', unsafe_allow_html=True)
    with c3:
        if st.session_state.get("task_done", False):
            if st.button("🚪 Logout", key="topbar_logout"):
                if user["role"] == "delivery Driver":
                    dr = find_row("driver_onboard","Mobile",user["phone"])
                    if dr: set_driver_status(dr["Driver ID"],"Offline")
                for k in DEFAULTS: st.session_state[k] = DEFAULTS[k]
                st.rerun()
        else:
            st.markdown(
                '<div style="text-align:right;padding-top:10px">'
                '<span style="color:#5a7a65;font-size:.78rem">🔒 Complete tasks to logout</span>'
                '</div>', unsafe_allow_html=True)
    st.divider()

# ═══════════════════════════════════════════════════════════════════════════════
#  PAGE: LOGIN & REGISTER
# ═══════════════════════════════════════════════════════════════════════════════
def page_login():
    st.markdown("""
    <div style="display:flex;flex-direction:column;align-items:center;
                margin-top:2.5rem;margin-bottom:1.5rem">
      <div style="width:68px;height:68px;border-radius:18px;background:#1a7f4b;
                  display:flex;align-items:center;justify-content:center;
                  font-size:34px;margin-bottom:12px;
                  box-shadow:0 8px 24px rgba(26,127,75,.35)">🧄</div>
      <h1 style="font-size:1.8rem;color:#0d1f14;margin:0">Garlic Order & Delivery</h1>
      <p style="color:#5a7a65;font-size:.95rem;margin-top:4px">Field Operations Platform</p>
    </div>""", unsafe_allow_html=True)

    col = st.columns([1,2,1])[1]
    with col:
        tab_lg, tab_rg = st.tabs(["🔐  Login","📝  Register"])
        with tab_lg:
            email = st.text_input("Email (Login ID)", placeholder="you@example.com", key="lg_email")
            pw    = st.text_input("Password", type="password", key="lg_pw")
            if st.button("Login →", type="primary", use_container_width=True, key="lg_btn"):
                if not email or not pw:
                    st.error("Enter email and password.")
                else:
                    with st.spinner("Verifying…"):
                        user, err = login_user(email, pw)
                    if err:
                        st.error(f"❌ {err}")
                    else:
                        st.session_state.logged_in = True
                        st.session_state.user      = user
                        st.session_state.task_done = False
                        if user["role"] == "delivery Driver":
                            dr = find_row("driver_onboard","Mobile",user["phone"])
                            if dr:
                                set_driver_status(dr["Driver ID"],"Active")
                                st.session_state.driver_id = dr["Driver ID"]
                        st.rerun()

        with tab_rg:
            rn        = st.text_input("Full name *", key="rg_name")
            rph       = st.text_input("Phone number *", key="rg_ph")
            rem       = st.text_input("Email (Login ID) *", placeholder="you@example.com", key="rg_email")
            rrol      = st.selectbox("Role *", ["sales executive","delivery Driver","admin"], key="rg_role")
            rpw       = st.text_input("Password *", type="password", key="rg_pw")
            rpw2      = st.text_input("Confirm password *", type="password", key="rg_pw2")
            adm_gate  = st.text_input("Admin Registration Password *",
                                       type="password", key="rg_adm_gate",
                                       help="Required for all roles. Contact admin.")
            if st.button("Create account →", type="primary", use_container_width=True, key="rg_btn"):
                if not all([rn,rph,rem,rpw,rpw2]):
                    st.error("Fill all required fields.")
                elif "@" not in rem:
                    st.error("Enter a valid email address.")
                elif len(rpw) < 6:
                    st.error("Password min 6 characters.")
                elif rpw != rpw2:
                    st.error("Passwords do not match.")
                elif not adm_gate:
                    st.error("❌ Admin Registration Password required.")
                elif adm_gate != ADMIN_REGISTER_PASSWORD:
                    st.error("❌ Wrong Admin Registration Password.")
                else:
                    with st.spinner("Creating account…"):
                        uid, err = register_user(rn, rph, rem, rrol, rpw)
                    if err:
                        st.error(f"❌ {err}")
                    else:
                        st.success("✅ Account created!")
                        st.info(f"UID: **`{uid}`** — Login with email: **{rem.strip().lower()}**")

# ═══════════════════════════════════════════════════════════════════════════════
#  ROUTE PLANNER — fully inlined (no external modules)
# ═══════════════════════════════════════════════════════════════════════════════
import streamlit.components.v1 as _stc
import json as _json

_GARLIC_ROUTE_HTML = '<!DOCTYPE html>\n<html lang="en">\n<head>\n<meta charset="UTF-8">\n<meta name="viewport" content="width=device-width, initial-scale=1.0">\n<title>Garlic Route Planner</title>\n<link href="https://fonts.googleapis.com/css2?family=Space+Mono:wght@400;700&family=DM+Sans:wght@300;400;500;600;700&display=swap" rel="stylesheet">\n<link rel="stylesheet" href="https://unpkg.com/leaflet@1.9.4/dist/leaflet.css"/>\n<script src="https://unpkg.com/leaflet@1.9.4/dist/leaflet.js"></script>\n<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js"></script>\n<style>\n:root{\n  --bg:#08101a;--surf:#0d1a28;--card:#0f1f30;--card2:#122437;\n  --border:#1a2e42;--border2:#1e3550;\n  --cyan:#00d4ff;--orange:#ff6b2b;--green:#2ecc71;\n  --red:#e74c3c;--yellow:#f39c12;--purple:#9b59b6;\n  --text:#d4e8f5;--muted:#4a6a85;--dim:#2a4560;\n}\n*{box-sizing:border-box;margin:0;padding:0}\nhtml{scroll-behavior:smooth}\nbody{font-family:\'DM Sans\',sans-serif;background:var(--bg);color:var(--text);min-height:100vh;overflow-x:hidden}\nbody::before{content:\'\';position:fixed;inset:0;\n  background-image:linear-gradient(rgba(0,212,255,.02) 1px,transparent 1px),linear-gradient(90deg,rgba(0,212,255,.02) 1px,transparent 1px);\n  background-size:50px 50px;pointer-events:none;z-index:0}\n.wrap{position:relative;z-index:1;max-width:1600px;margin:0 auto;padding:14px 18px}\n\n/* HEADER */\n.hdr{display:flex;align-items:center;justify-content:space-between;flex-wrap:wrap;gap:10px;\n  background:var(--surf);border:1px solid var(--border);border-radius:12px;\n  padding:16px 22px;margin-bottom:14px;position:relative;overflow:hidden}\n.hdr::after{content:\'\';position:absolute;top:0;left:0;right:0;height:2px;\n  background:linear-gradient(90deg,var(--cyan),var(--orange),var(--green),var(--cyan))}\n.hdr-brand h1{font-family:\'Space Mono\',monospace;font-size:1rem;color:var(--cyan);letter-spacing:2px;text-transform:uppercase}\n.hdr-brand p{font-size:.68rem;color:var(--muted);font-family:\'Space Mono\',monospace;margin-top:2px}\n.hdr-r{display:flex;align-items:center;gap:8px;flex-wrap:wrap}\n.cron-badge{display:flex;align-items:center;gap:6px;background:rgba(0,212,255,.07);\n  border:1px solid rgba(0,212,255,.2);border-radius:20px;padding:5px 12px;\n  font-family:\'Space Mono\',monospace;font-size:.63rem;color:var(--cyan)}\n.dot{width:6px;height:6px;border-radius:50%;background:var(--green);animation:blink 1.4s infinite}\n@keyframes blink{0%,100%{opacity:1;transform:scale(1)}50%{opacity:.3;transform:scale(.6)}}\n.cdwn{font-family:\'Space Mono\',monospace;font-size:.63rem;color:var(--muted)}\n.cdwn b{color:var(--yellow)}\n\n/* BUTTONS */\n.btn{display:inline-flex;align-items:center;gap:5px;padding:7px 14px;border:none;border-radius:7px;\n  cursor:pointer;font-family:\'Space Mono\',monospace;font-size:.63rem;letter-spacing:.8px;\n  text-transform:uppercase;font-weight:700;transition:all .15s;white-space:nowrap}\n.bc{background:var(--cyan);color:#08101a}.bc:hover{background:#33ddff;transform:translateY(-1px)}\n.bo{background:linear-gradient(135deg,var(--orange),#e74c3c);color:#fff}.bo:hover{filter:brightness(1.1);transform:translateY(-1px)}\n.bg{background:rgba(46,204,113,.12);color:var(--green);border:1px solid rgba(46,204,113,.3)}.bg:hover{background:rgba(46,204,113,.2)}\n.bgh{background:rgba(255,255,255,.04);color:var(--muted);border:1px solid var(--border)}.bgh:hover{border-color:var(--cyan);color:var(--cyan)}\n.bsm{padding:4px 9px;font-size:.58rem}\n.bred{background:rgba(231,76,60,.12);color:var(--red);border:1px solid rgba(231,76,60,.25)}.bred:hover{background:rgba(231,76,60,.22)}\n.bpur{background:rgba(155,89,182,.12);color:var(--purple);border:1px solid rgba(155,89,182,.3)}.bpur:hover{background:rgba(155,89,182,.22)}\n\n/* STATS */\n.stats-bar{display:flex;gap:10px;flex-wrap:wrap;margin-bottom:14px}\n.sc{flex:1;min-width:120px;background:var(--surf);border:1px solid var(--border);border-radius:10px;padding:11px 14px;position:relative;overflow:hidden}\n.sc::before{content:\'\';position:absolute;top:0;left:0;width:3px;height:100%;background:var(--cyan)}\n.sc.o::before{background:var(--orange)}.sc.g::before{background:var(--green)}\n.sc.y::before{background:var(--yellow)}.sc.p::before{background:var(--purple)}.sc.r::before{background:var(--red)}\n.sl{font-family:\'Space Mono\',monospace;font-size:.55rem;color:var(--muted);text-transform:uppercase;letter-spacing:1px;margin-bottom:3px}\n.sv{font-family:\'Space Mono\',monospace;font-size:1.2rem;color:var(--text);font-weight:700}\n.ss{font-size:.62rem;color:var(--muted);margin-top:1px}\n\n/* TABS */\n.tabs{display:flex;gap:3px;background:var(--surf);border:1px solid var(--border);border-radius:10px;padding:3px;margin-bottom:14px;flex-wrap:wrap}\n.tab{flex:1;min-width:100px;text-align:center;padding:7px 10px;border-radius:7px;cursor:pointer;\n  font-family:\'Space Mono\',monospace;font-size:.6rem;letter-spacing:.8px;text-transform:uppercase;\n  color:var(--muted);transition:all .18s}\n.tab.active{background:var(--card2);color:var(--cyan);border:1px solid rgba(0,212,255,.2)}\n.tab-content{display:none}.tab-content.active{display:block}\n\n/* PANELS */\n.panel{background:var(--surf);border:1px solid var(--border);border-radius:12px;overflow:hidden;margin-bottom:14px}\n.ph{display:flex;align-items:center;justify-content:space-between;padding:11px 16px;\n  border-bottom:1px solid var(--border);background:rgba(0,0,0,.2);flex-wrap:wrap;gap:8px}\n.pt{font-family:\'Space Mono\',monospace;font-size:.62rem;color:var(--muted);text-transform:uppercase;letter-spacing:1.5px}\n.pb{padding:16px}\n\n/* FORM */\n.form-r{display:grid;grid-template-columns:1fr 1fr;gap:9px;margin-bottom:9px}\n.fg{display:flex;flex-direction:column;gap:4px}\n.fg.full{grid-column:1/-1}\n.fg.third{grid-column:span 1}\n.fg label{font-size:.58rem;color:var(--muted);text-transform:uppercase;letter-spacing:1px;font-family:\'Space Mono\',monospace}\n.fg input,.fg select,.fg textarea{background:var(--card2);border:1px solid var(--border);color:var(--text);\n  padding:7px 10px;border-radius:7px;font-family:\'DM Sans\',sans-serif;font-size:.8rem;outline:none;transition:border-color .18s}\n.fg input:focus,.fg select:focus,.fg textarea:focus{border-color:var(--cyan)}\n.fg input::placeholder{color:var(--dim)}\nselect option{background:var(--card2)}\n\n/* LAYOUT */\n.split{display:grid;grid-template-columns:1fr 1fr;gap:14px}\n.split-l{display:grid;grid-template-columns:400px 1fr;gap:14px;align-items:start}\n.split-3c{display:grid;grid-template-columns:1fr 1fr 1fr;gap:14px}\n\n/* UPLOAD */\n.drop-zone{border:2px dashed var(--border2);border-radius:10px;padding:28px 20px;text-align:center;cursor:pointer;transition:all .2s;position:relative}\n.drop-zone:hover,.drop-zone.over{border-color:var(--cyan);background:rgba(0,212,255,.03)}\n.drop-zone input{position:absolute;inset:0;opacity:0;cursor:pointer;width:100%;height:100%}\n.dz-ico{font-size:1.8rem;margin-bottom:7px;opacity:.4}\n.dz-txt{font-size:.8rem;color:var(--muted)}.dz-txt strong{color:var(--cyan)}\n.dz-hint{font-size:.63rem;color:var(--dim);margin-top:4px;font-family:\'Space Mono\',monospace}\n.file-pill{display:inline-flex;align-items:center;gap:7px;background:rgba(46,204,113,.08);\n  border:1px solid rgba(46,204,113,.25);border-radius:7px;padding:6px 12px;margin-top:8px;font-size:.75rem}\n.file-pill .ck{color:var(--green)}\n\n/* COLUMN MAPPER */\n.cm-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(190px,1fr));gap:9px;margin-top:9px}\n.cm-item{display:flex;flex-direction:column;gap:4px}\n.cm-item label{font-size:.58rem;color:var(--muted);font-family:\'Space Mono\',monospace;text-transform:uppercase;letter-spacing:1px}\n\n/* TABLE */\n.rt{width:100%;border-collapse:collapse;font-size:.74rem}\n.rt th{font-family:\'Space Mono\',monospace;font-size:.56rem;color:var(--muted);text-transform:uppercase;\n  letter-spacing:1px;padding:7px 9px;border-bottom:1px solid var(--border);text-align:left;white-space:nowrap}\n.rt td{padding:7px 9px;border-bottom:1px solid rgba(26,46,66,.4);vertical-align:middle}\n.rt tr:hover td{background:rgba(0,212,255,.02)}\n.rt tr:last-child td{border:none}\n.trip-tag{display:inline-block;padding:2px 7px;border-radius:4px;font-family:\'Space Mono\',monospace;font-size:.58rem;font-weight:700;color:#fff}\n\n/* TRIP COLORS */\n.tc0{background:linear-gradient(135deg,#c0392b,#e74c3c)}\n.tc1{background:linear-gradient(135deg,#d35400,#e67e22)}\n.tc2{background:linear-gradient(135deg,#b7950b,#f39c12)}\n.tc3{background:linear-gradient(135deg,#1a8c4e,#27ae60)}\n.tc4{background:linear-gradient(135deg,#148f77,#1abc9c)}\n.tc5{background:linear-gradient(135deg,#1f618d,#2980b9)}\n.tc6{background:linear-gradient(135deg,#7d3c98,#9b59b6)}\n.tc7{background:linear-gradient(135deg,#a93226,#e91e63)}\n.tc8{background:linear-gradient(135deg,#bf360c,#ff5722)}\n.tc9{background:linear-gradient(135deg,#4e342e,#795548)}\n.tc10{background:linear-gradient(135deg,#00695c,#00897b)}\n.tc11{background:linear-gradient(135deg,#283593,#3949ab)}\n\n/* ROUTEMAP CARDS */\n.rm-grid{display:flex;gap:8px;overflow-x:auto;padding-bottom:8px}\n.rm-grid::-webkit-scrollbar{height:5px}\n.rm-grid::-webkit-scrollbar-thumb{background:var(--border);border-radius:3px}\n.trip-col{min-width:155px;max-width:168px;flex-shrink:0;border-radius:9px;overflow:hidden;\n  border:1px solid var(--border);background:var(--card);animation:fup .3s ease;transition:opacity .2s}\n.trip-col.drag-over-col{border:2px dashed var(--cyan)!important;background:rgba(0,212,255,.05)}\n.trip-col.dragging-col{opacity:.4}\n@keyframes fup{from{opacity:0;transform:translateY(8px)}to{opacity:1;transform:translateY(0)}}\n.trip-hdr{padding:6px 8px 4px;font-size:.6rem;font-weight:700;color:#fff;line-height:1.3;cursor:pointer}\n.trip-hdr:hover{filter:brightness(1.1)}\n.trip-title{font-family:\'Space Mono\',monospace;letter-spacing:.4px;font-size:.62rem}\n.trip-cc{font-size:.54rem;opacity:.8;margin-top:1px}\n.trip-ton{background:rgba(0,0,0,.35);padding:3px 8px;font-family:\'Space Mono\',monospace;font-size:.63rem;\n  color:#fff;border-bottom:1px solid rgba(255,255,255,.08);display:flex;justify-content:space-between;align-items:center}\n.trip-customers{padding:4px 6px;display:flex;flex-direction:column;gap:3px;max-height:300px;overflow-y:auto;\n  min-height:40px;transition:background .15s}\n.trip-customers.drop-target{background:rgba(0,212,255,.08);border-radius:5px}\n.trip-customers::-webkit-scrollbar{width:3px}\n.trip-customers::-webkit-scrollbar-thumb{background:var(--dim)}\n.cust-card{background:rgba(0,0,0,.18);border-radius:5px;padding:5px 7px;border:1px solid rgba(255,255,255,.04);\n  cursor:grab;transition:border-color .12s,opacity .15s,transform .15s;user-select:none}\n.cust-card:active{cursor:grabbing}\n.cust-card:hover{border-color:rgba(0,212,255,.35)}\n.cust-card.dragging{opacity:.35;transform:scale(.97)}\n.cust-card.drag-over{border-top:2px solid var(--cyan)!important}\n.cust-name{font-size:.67rem;font-weight:600;color:var(--text);line-height:1.2;margin-bottom:2px}\n.cust-meta{font-family:\'Space Mono\',monospace;font-size:.55rem;color:var(--muted);line-height:1.7}\n.cust-meta .cr{color:var(--cyan)}.cust-meta .ti{color:var(--yellow)}.cust-meta .tn{color:var(--orange)}\n.snum{display:inline-flex;align-items:center;justify-content:center;width:14px;height:14px;\n  border-radius:50%;font-family:\'Space Mono\',monospace;font-size:.5rem;font-weight:700;\n  background:var(--cyan);color:#08101a;margin-right:3px;vertical-align:middle;flex-shrink:0}\n.trip-edit-btn{font-size:.5rem;padding:1px 5px;margin-left:4px}\n\n/* LEAFLET MAP + SIDEBAR */\n.map-layout{display:flex;gap:0;height:580px;border-radius:0 0 12px 12px;overflow:hidden}\n.map-sidebar{width:240px;flex-shrink:0;background:var(--card);border-right:1px solid var(--border);\n  display:flex;flex-direction:column;overflow:hidden}\n.map-sidebar-hdr{padding:10px 12px;border-bottom:1px solid var(--border);background:rgba(0,0,0,.2)}\n.map-sidebar-hdr-title{font-family:\'Space Mono\',monospace;font-size:.58rem;color:var(--muted);\n  text-transform:uppercase;letter-spacing:1.2px;margin-bottom:6px}\n.map-sidebar-list{flex:1;overflow-y:auto;padding:6px}\n.map-sidebar-list::-webkit-scrollbar{width:3px}\n.map-sidebar-list::-webkit-scrollbar-thumb{background:var(--border)}\n.msl-item{border-radius:8px;margin-bottom:5px;border:1px solid var(--border);cursor:pointer;\n  transition:all .15s;overflow:hidden}\n.msl-item:hover{border-color:rgba(0,212,255,.3);transform:translateX(2px)}\n.msl-item.active{border-color:var(--cyan)!important;box-shadow:0 0 10px rgba(0,212,255,.15)}\n.msl-item-hdr{padding:7px 10px;display:flex;align-items:center;justify-content:space-between}\n.msl-trip-name{font-family:\'Space Mono\',monospace;font-size:.65rem;font-weight:700;color:#fff}\n.msl-stops{font-family:\'Space Mono\',monospace;font-size:.55rem;color:rgba(255,255,255,.6)}\n.msl-item-body{padding:5px 10px 7px;background:rgba(0,0,0,.25)}\n.msl-stat{display:flex;justify-content:space-between;font-size:.62rem;margin-bottom:2px}\n.msl-stat-lbl{color:var(--muted);font-family:\'Space Mono\',monospace;font-size:.56rem}\n.msl-stat-val{font-family:\'Space Mono\',monospace;font-size:.6rem}\n.msl-all{border-radius:8px;margin-bottom:5px;border:1px solid var(--border);cursor:pointer;\n  transition:all .15s;padding:8px 10px;background:rgba(0,212,255,.05)}\n.msl-all:hover{border-color:var(--cyan)}\n.msl-all.active{border-color:var(--cyan);background:rgba(0,212,255,.1)}\n.msl-all-txt{font-family:\'Space Mono\',monospace;font-size:.63rem;color:var(--cyan);text-align:center}\n#leafMap{height:100%;border-radius:0;z-index:1;flex:1}\n\n/* CRON */\n.ibtn{padding:4px 11px;border-radius:5px;border:1px solid var(--border);background:transparent;\n  color:var(--muted);font-family:\'Space Mono\',monospace;font-size:.62rem;cursor:pointer;transition:all .13s}\n.ibtn.active,.ibtn:hover{border-color:var(--cyan);color:var(--cyan);background:rgba(0,212,255,.07)}\n.prog-bar{height:3px;background:var(--border);border-radius:2px;overflow:hidden;margin-top:6px}\n.prog-fill{height:100%;background:linear-gradient(90deg,var(--cyan),var(--green));border-radius:2px;transition:width .3s}\n\n/* LOG */\n.logbox{font-family:\'Space Mono\',monospace;font-size:.62rem;max-height:150px;overflow-y:auto;display:flex;flex-direction:column;gap:2px}\n.logbox::-webkit-scrollbar{width:3px}\n.logbox::-webkit-scrollbar-thumb{background:var(--border)}\n.le{display:flex;gap:8px;line-height:1.6;animation:fi .2s ease}\n@keyframes fi{from{opacity:0}to{opacity:1}}\n.lt2{color:var(--dim);flex-shrink:0}.lm{color:var(--text)}\n.lok .lm{color:var(--green)}.lwarn .lm{color:var(--yellow)}.lerr .lm{color:var(--red)}.linfo .lm{color:var(--cyan)}\n\n/* EMPTY */\n.empty{text-align:center;padding:32px;color:var(--muted);font-family:\'Space Mono\',monospace;font-size:.68rem}\n.eico{font-size:1.8rem;margin-bottom:8px;opacity:.3}\n\n/* MODAL */\n.modal-bg{position:fixed;inset:0;background:rgba(0,0,0,.75);z-index:9999;display:none;align-items:center;justify-content:center;padding:20px}\n.modal-bg.open{display:flex}\n.modal{background:var(--surf);border:1px solid var(--border);border-radius:14px;\n  width:100%;max-width:720px;max-height:90vh;overflow-y:auto;position:relative}\n.modal::-webkit-scrollbar{width:4px}\n.modal::-webkit-scrollbar-thumb{background:var(--border)}\n.modal-hdr{display:flex;align-items:center;justify-content:space-between;padding:16px 20px;\n  border-bottom:1px solid var(--border);background:rgba(0,0,0,.2);position:sticky;top:0;z-index:2;background:var(--surf)}\n.modal-title{font-family:\'Space Mono\',monospace;font-size:.8rem;color:var(--cyan);text-transform:uppercase;letter-spacing:1px}\n.modal-body{padding:18px}\n.modal-close{background:none;border:none;color:var(--muted);font-size:1.2rem;cursor:pointer;padding:2px 6px}\n.modal-close:hover{color:var(--red)}\n\n/* TRIP EDIT TABLE */\n.te-row{display:grid;grid-template-columns:32px 1fr 90px 80px 80px 80px 90px 80px 36px;gap:6px;align-items:center;\n  padding:5px 0;border-bottom:1px solid rgba(26,46,66,.4);font-size:.72rem}\n.te-row:last-child{border:none}\n.te-hdr{font-family:\'Space Mono\',monospace;font-size:.55rem;color:var(--muted);text-transform:uppercase;\n  padding:5px 0;border-bottom:1px solid var(--border);font-weight:700}\n.te-num{font-family:\'Space Mono\',monospace;font-size:.65rem;color:var(--cyan);text-align:center}\n.te-inp{background:var(--card2);border:1px solid var(--border);color:var(--text);\n  padding:4px 7px;border-radius:5px;font-size:.72rem;outline:none;width:100%;font-family:\'DM Sans\',sans-serif}\n.te-inp:focus{border-color:var(--cyan)}\n.drag-handle{cursor:grab;color:var(--dim);text-align:center;font-size:.9rem;user-select:none}\n.drag-handle:active{cursor:grabbing}\n\n/* COST PAGE */\n.cost-section{background:var(--card);border:1px solid var(--border);border-radius:10px;padding:16px;margin-bottom:12px}\n.cost-title{font-family:\'Space Mono\',monospace;font-size:.65rem;color:var(--cyan);text-transform:uppercase;letter-spacing:1.5px;margin-bottom:12px;display:flex;align-items:center;gap:8px}\n.cost-title::after{content:\'\';flex:1;height:1px;background:var(--border)}\n.cost-grid{display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:10px}\n.cost-row{display:flex;flex-direction:column;gap:4px}\n.cost-row label{font-size:.58rem;color:var(--muted);font-family:\'Space Mono\',monospace;text-transform:uppercase;letter-spacing:1px}\n.cost-row input{background:var(--bg);border:1px solid var(--border);color:var(--text);\n  padding:7px 10px;border-radius:7px;font-size:.82rem;outline:none;transition:border-color .18s;font-family:\'DM Sans\',sans-serif}\n.cost-row input:focus{border-color:var(--cyan)}\n\n/* COST TABLE */\n.cost-table{width:100%;border-collapse:collapse;font-size:.74rem}\n.cost-table th{font-family:\'Space Mono\',monospace;font-size:.55rem;color:var(--muted);text-transform:uppercase;\n  letter-spacing:1px;padding:7px 10px;border-bottom:1px solid var(--border);text-align:left;white-space:nowrap}\n.cost-table td{padding:7px 10px;border-bottom:1px solid rgba(26,46,66,.4);vertical-align:middle}\n.cost-table tr:hover td{background:rgba(0,212,255,.02)}\n.cost-table .total-row td{background:rgba(0,212,255,.05);font-weight:700;color:var(--cyan);border-top:1px solid var(--border)}\n.cost-val{font-family:\'Space Mono\',monospace;color:var(--green)}\n.cost-val.big{font-size:.95rem;color:var(--cyan)}\n\n/* VEHICLE MANAGER */\n.veh-list{display:flex;flex-direction:column;gap:6px;max-height:340px;overflow-y:auto}\n.veh-item{display:grid;grid-template-columns:80px 1fr 80px 80px 90px 36px;gap:8px;align-items:center;\n  background:var(--card2);border:1px solid var(--border);border-radius:8px;padding:8px 10px;font-size:.74rem}\n.veh-item:hover{border-color:rgba(0,212,255,.25)}\n\n/* MAP SIDEBAR */\n.map-trip-btn{padding:8px 10px;cursor:pointer;border-bottom:1px solid var(--border);\n  display:flex;align-items:flex-start;gap:8px;transition:background .15s}\n.map-trip-btn:hover{background:rgba(0,212,255,.04)}\n.active-trip-btn{background:rgba(0,212,255,.08)!important;border-left:3px solid var(--cyan)}\n.mst-dot{width:10px;height:10px;border-radius:50%;flex-shrink:0;margin-top:3px}\n.mst-info{flex:1;min-width:0}\n.mst-name{font-family:\'Space Mono\',monospace;font-size:.6rem;font-weight:700;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}\n.mst-km{font-family:\'Space Mono\',monospace;font-size:.55rem;color:var(--muted);margin-top:1px}\n.mst-cost{font-family:\'Space Mono\',monospace;font-size:.6rem;color:var(--green);font-weight:700}\n.mst-stops{font-size:.58rem;color:var(--muted);margin-top:1px}\n.mst-ton{font-family:\'Space Mono\',monospace;font-size:.55rem;color:var(--orange)}\n.pb-h{background:rgba(231,76,60,.15);color:#e74c3c;padding:2px 7px;border-radius:10px;font-family:\'Space Mono\',monospace;font-size:.56rem;font-weight:700}\n.pb-m{background:rgba(243,156,18,.15);color:#f39c12;padding:2px 7px;border-radius:10px;font-family:\'Space Mono\',monospace;font-size:.56rem;font-weight:700}\n.pb-l{background:rgba(46,204,113,.12);color:#2ecc71;padding:2px 7px;border-radius:10px;font-family:\'Space Mono\',monospace;font-size:.56rem;font-weight:700}\n\n@media(max-width:900px){.split,.split-l,.split-3c{grid-template-columns:1fr}}\n</style>\n</head>\n<body>\n<div class="wrap">\n\n<!-- HEADER -->\n<header class="hdr">\n  <div class="hdr-brand">\n    <h1>⬡ Garlic Route Planner</h1>\n    <p>// Garlic Order & Delivery · Route Optimization · Trip Creation</p>\n  </div>\n  <div class="hdr-r">\n    <div class="cron-badge"><span class="dot"></span>CRON</div>\n    <div class="cdwn">NEXT: <b id="nextRun">--</b></div>\n    <button class="btn bo" onclick="runCron()">▶ RUN CRON</button>\n    <button class="btn bg" onclick="exportXLSX()">⬇ XLSX</button>\n  </div>\n</header>\n\n<!-- STATS -->\n<div class="stats-bar">\n  <div class="sc"><div class="sl">Total Orders</div><div class="sv" id="s-ord">0</div><div class="ss">loaded</div></div>\n  <div class="sc o"><div class="sl">Trips</div><div class="sv" id="s-trips">0</div><div class="ss">active routes</div></div>\n  <div class="sc g"><div class="sl">Distance</div><div class="sv" id="s-dist">—</div><div class="ss">total km</div></div>\n  <div class="sc y"><div class="sl">Crates</div><div class="sv" id="s-crates">0</div><div class="ss">units</div></div>\n  <div class="sc p"><div class="sl">Tonnage</div><div class="sv" id="s-ton">0</div><div class="ss">kg</div></div>\n  <div class="sc r"><div class="sl">Est. Cost</div><div class="sv" id="s-cost">—</div><div class="ss">total</div></div>\n</div>\n\n<!-- TABS -->\n<div class="tabs">\n  <div class="tab active" onclick="showTab(\'upload\')">① Upload</div>\n  <div class="tab" onclick="showTab(\'routemap\')">② RouteMap</div>\n  <div class="tab" onclick="showTab(\'map\')">③ Live Map</div>\n  <div class="tab" onclick="showTab(\'table\')">④ Route Table</div>\n  <div class="tab" onclick="showTab(\'trips\')">⑤ Edit Trips</div>\n  <div class="tab" onclick="showTab(\'vehicles\')">⑥ Vehicles</div>\n  <div class="tab" onclick="showTab(\'cost\')">⑦ Cost Calc</div>\n  <div class="tab" onclick="showTab(\'cron\')">⑧ CRON</div>\n  <div class="tab" onclick="showTab(\'submit\')" id="tab-btn-submit" style="color:var(--green);border:1px solid rgba(46,204,113,.3)">✅ Submit</div>\n</div>\n\n<!-- ① UPLOAD -->\n<div id="tab-upload" class="tab-content active">\n<div class="split-l">\n  <div>\n    <div class="panel">\n      <div class="ph"><span class="pt">Upload Excel / CSV</span></div>\n      <div class="pb">\n        <div class="drop-zone" id="dropZone" ondragover="dzOver(event)" ondragleave="dzLeave()" ondrop="dzDrop(event)">\n          <input type="file" id="fileInput" accept=".xlsx,.xls,.csv" onchange="fileChg(event)">\n          <div class="dz-ico">📂</div>\n          <div class="dz-txt">Drop <strong>Excel / CSV</strong> here or click</div>\n          <div class="dz-hint">.xlsx · .xls · .csv supported</div>\n        </div>\n        <div id="filePill" style="display:none" class="file-pill">\n          <span class="ck">✓</span><span id="fName">—</span>\n          <span style="color:var(--muted);font-size:.65rem" id="fRows">—</span>\n        </div>\n      </div>\n    </div>\n    <div class="panel" id="cmPanel" style="display:none">\n      <div class="ph"><span class="pt">Map Columns</span><button class="btn bc bsm" onclick="applyMap()">✓ APPLY</button></div>\n      <div class="pb">\n        <p style="font-size:.72rem;color:var(--muted);margin-bottom:8px">Match file columns → required fields:</p>\n        <div class="cm-grid" id="cmGrid"></div>\n      </div>\n    </div>\n    <div class="panel">\n      <div class="ph"><span class="pt">Add Order Manually</span></div>\n      <div class="pb">\n        <div class="form-r">\n          <div class="fg"><label>Trip #</label><input id="m-trip" type="number" placeholder="2" min="1"></div>\n          <div class="fg"><label>Shift</label><input id="m-shift" type="number" placeholder="1" min="1"></div>\n          <div class="fg"><label>Vehicle</label>\n            <select id="m-veh"><option value="">-- select --</option></select>\n          </div>\n          <div class="fg"><label>Order ID</label><input id="m-oid" placeholder="ORD-001"></div>\n          <div class="fg full"><label>Customer Name</label><input id="m-cust" placeholder="SLV General Stores"></div>\n          <div class="fg full"><label>Address</label><input id="m-addr" placeholder="123 Main Road, Area"></div>\n          <div class="fg"><label>Lat</label><input id="m-lat" type="number" placeholder="12.9716" step="any"></div>\n          <div class="fg"><label>Lng</label><input id="m-lng" type="number" placeholder="77.5946" step="any"></div>\n          <div class="fg"><label>Crates</label><input id="m-cr" type="number" placeholder="19" step="any"></div>\n          <div class="fg"><label>Tonnage (kg)</label><input id="m-tn" type="number" placeholder="267" step="any"></div>\n          <div class="fg"><label>Time Window</label><input id="m-tw" placeholder="05:30-06:30"></div>\n          <div class="fg"><label>Priority</label>\n            <select id="m-pri"><option value="high">HIGH</option><option value="med" selected>MEDIUM</option><option value="low">LOW</option></select>\n          </div>\n        </div>\n        <button class="btn bc" style="width:100%;justify-content:center" onclick="addManual()">+ ADD ORDER</button>\n      </div>\n    </div>\n  </div>\n  <div>\n    <div class="panel">\n      <div class="ph">\n        <span class="pt">Orders (<span id="ordCount">0</span>)</span>\n        <div style="display:flex;gap:6px">\n          <button class="btn bgh bsm" onclick="loadDemo()">DEMO DATA</button>\n          <button class="btn bred bsm" onclick="clearAll()">✕ CLEAR</button>\n        </div>\n      </div>\n      <div style="overflow-x:auto">\n        <table class="rt"><thead><tr><th>Trip</th><th>Customer</th><th>Address</th><th>Crates</th><th>Tonnage</th><th>Window</th><th>Lat,Lng</th><th></th></tr></thead>\n        <tbody id="prevBody"><tr><td colspan="8"><div class="empty"><div class="eico">📋</div>No data yet</div></td></tr></tbody></table>\n      </div>\n    </div>\n  </div>\n</div>\n</div>\n\n<!-- ② ROUTEMAP -->\n<div id="tab-routemap" class="tab-content">\n  <div class="panel">\n    <div class="ph">\n      <span class="pt">RouteMap Cards</span>\n      <div style="display:flex;gap:8px;align-items:center">\n        <span style="font-family:\'Space Mono\',monospace;font-size:.6rem;color:var(--muted)" id="rmSub">Run CRON to populate</span>\n        <button class="btn bgh bsm" onclick="showTab(\'trips\')">✏ EDIT TRIPS</button>\n      </div>\n    </div>\n    <div class="pb">\n      <div class="rm-grid" id="rmGrid">\n        <div class="empty"><div class="eico">🗺️</div>Upload data and run CRON</div>\n      </div>\n    </div>\n  </div>\n</div>\n\n<!-- ③ LIVE MAP -->\n<div id="tab-map" class="tab-content">\n  <div class="panel" style="overflow:hidden">\n    <div class="ph">\n      <span class="pt">Live Map — OpenStreetMap</span>\n      <div style="display:flex;gap:8px;align-items:center;flex-wrap:wrap">\n        <span style="font-family:\'Space Mono\',monospace;font-size:.58rem;color:var(--muted)">Click trip → filter map</span>\n        <button class="btn bgh bsm" onclick="mapFitAll()">⊞ FIT ALL</button>\n      </div>\n    </div>\n    <div style="display:flex;height:580px">\n      <!-- LEFT SIDEBAR -->\n      <div id="mapSidebar" style="width:220px;flex-shrink:0;background:var(--bg);border-right:1px solid var(--border);overflow-y:auto;display:flex;flex-direction:column">\n        <div style="padding:8px 10px;font-family:\'Space Mono\',monospace;font-size:.56rem;color:var(--muted);text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid var(--border);background:rgba(0,0,0,.2)">\n          SELECT TRIP\n        </div>\n        <!-- ALL TRIPS button -->\n        <div id="mst-all" class="map-trip-btn active-trip-btn" onclick="mapSelectTrip(\'all\')"\n          style="padding:8px 10px;cursor:pointer;border-bottom:1px solid var(--border);display:flex;align-items:center;gap:8px;transition:background .15s">\n          <div style="width:10px;height:10px;border-radius:50%;background:var(--cyan);flex-shrink:0"></div>\n          <div>\n            <div style="font-family:\'Space Mono\',monospace;font-size:.6rem;font-weight:700;color:var(--cyan)">ALL TRIPS</div>\n            <div style="font-size:.6rem;color:var(--muted)" id="mst-all-sub">—</div>\n          </div>\n        </div>\n        <!-- Per-trip buttons injected here -->\n        <div id="mapTripList"></div>\n      </div>\n      <!-- MAP -->\n      <div style="flex:1;position:relative;min-width:0">\n        <div id="leafMap" style="height:100%;width:100%"></div>\n      </div>\n    </div>\n  </div>\n</div>\n\n<!-- ④ ROUTE TABLE -->\n<div id="tab-table" class="tab-content">\n  <div class="panel">\n    <div class="ph">\n      <span class="pt">Optimized Route Sequence</span>\n      <span style="font-family:\'Space Mono\',monospace;font-size:.6rem;color:var(--muted)" id="lRunLbl">LAST RUN: —</span>\n      <span style="font-family:\'Space Mono\',monospace;font-size:.55rem;color:var(--dim);background:rgba(0,212,255,.06);border:1px solid rgba(0,212,255,.15);padding:2px 8px;border-radius:4px">2-OPT + OR-OPT + PRIORITY</span>\n    </div>\n    <div style="overflow-x:auto">\n      <table class="rt"><thead><tr>\n        <th>Stop</th><th>Trip</th><th>Order ID</th><th>Customer</th><th>Address</th>\n        <th>Crates</th><th>Tonnage</th><th>Window</th><th>Priority</th>\n        <th>Leg km</th><th>Cum km</th><th>ETA</th>\n      </tr></thead>\n      <tbody id="routeBody"><tr><td colspan="12"><div class="empty"><div class="eico">📍</div>Run CRON to generate routes</div></td></tr></tbody>\n      </table>\n    </div>\n  </div>\n</div>\n\n<!-- ⑤ EDIT TRIPS -->\n<div id="tab-trips" class="tab-content">\n  <div class="split">\n    <div>\n      <div class="panel">\n        <div class="ph"><span class="pt">Create / Edit Trip</span></div>\n        <div class="pb">\n          <div class="form-r">\n            <div class="fg"><label>Trip # to Edit</label><input id="et-num" type="number" placeholder="2" min="1"></div>\n            <div class="fg"><label>Shift</label><input id="et-shift" type="number" placeholder="1" min="1"></div>\n            <div class="fg full"><label>Vehicle</label>\n              <select id="et-veh"><option value="">-- select vehicle --</option></select>\n            </div>\n          </div>\n          <div style="display:flex;gap:8px;margin-bottom:12px">\n            <button class="btn bc" onclick="loadTripForEdit()">LOAD TRIP</button>\n            <button class="btn bgh" onclick="newTripFromOrders()">CREATE NEW TRIP</button>\n          </div>\n          <div id="tripEditArea">\n            <div class="empty"><div class="eico">✏️</div>Enter trip number and click LOAD TRIP</div>\n          </div>\n          <div id="tripEditActions" style="display:none;margin-top:10px;gap:8px;display:none">\n            <button class="btn bc" onclick="saveTripEdit()">💾 SAVE CHANGES</button>\n            <button class="btn bred" onclick="deleteTrip()">🗑 DELETE TRIP</button>\n            <button class="btn bgh" onclick="reoptimizeTrip()">🔄 RE-OPTIMIZE</button>\n          </div>\n        </div>\n      </div>\n\n      <div class="panel">\n        <div class="ph"><span class="pt">Add Customer to Trip</span></div>\n        <div class="pb">\n          <div style="margin-bottom:8px">\n            <label style="font-size:.58rem;color:var(--muted);font-family:\'Space Mono\',monospace;text-transform:uppercase;display:block;margin-bottom:4px">Add from loaded orders to Trip #</label>\n            <div style="display:flex;gap:8px">\n              <input id="add-to-trip" type="number" placeholder="Trip #" style="background:var(--card2);border:1px solid var(--border);color:var(--text);padding:7px 10px;border-radius:7px;font-size:.8rem;outline:none;width:90px">\n              <select id="add-cust-sel" style="background:var(--card2);border:1px solid var(--border);color:var(--text);padding:7px 10px;border-radius:7px;font-size:.8rem;outline:none;flex:1" id="addCustSel"></select>\n              <button class="btn bc bsm" onclick="addCustToTrip()">ADD</button>\n            </div>\n          </div>\n        </div>\n      </div>\n    </div>\n\n    <div>\n      <div class="panel">\n        <div class="ph"><span class="pt">All Trips Overview</span></div>\n        <div style="overflow-x:auto">\n          <table class="rt"><thead><tr>\n            <th>Trip</th><th>Shift</th><th>Vehicle</th><th>Stops</th><th>Total Crates</th><th>Tonnage</th><th>Dist km</th><th>Actions</th>\n          </tr></thead>\n          <tbody id="tripsOverview"><tr><td colspan="8"><div class="empty"><div class="eico">🚚</div>No trips yet</div></td></tr></tbody>\n          </table>\n        </div>\n      </div>\n    </div>\n  </div>\n</div>\n\n<!-- ⑥ VEHICLES -->\n<div id="tab-vehicles" class="tab-content">\n  <div class="split">\n    <div>\n      <div class="panel">\n        <div class="ph"><span class="pt">Add / Edit Vehicle</span></div>\n        <div class="pb">\n          <div class="form-r">\n            <div class="fg"><label>Vehicle ID</label><input id="v-id" placeholder="VEH-01"></div>\n            <div class="fg"><label>Vehicle Name</label><input id="v-name" placeholder="Tata Ace"></div>\n            <div class="fg"><label>Type</label>\n              <select id="v-type">\n                <option>Mini Truck</option><option>Pickup</option><option>Van</option>\n                <option>Medium Truck</option><option>Heavy Truck</option><option>Bike</option><option>Other</option>\n              </select>\n            </div>\n            <div class="fg"><label>Capacity (kg)</label><input id="v-cap" type="number" placeholder="750"></div>\n            <div class="fg"><label>CC / Engine</label><input id="v-cc" placeholder="CC 58.1"></div>\n            <div class="fg"><label>Fuel Type</label>\n              <select id="v-fuel"><option>Diesel</option><option>Petrol</option><option>CNG</option><option>Electric</option></select>\n            </div>\n            <div class="fg"><label>Fixed Cost/day (₹)</label><input id="v-fc" type="number" placeholder="800" step="any"></div>\n            <div class="fg"><label>Cost/km (₹)</label><input id="v-kmc" type="number" placeholder="8" step="any"></div>\n            <div class="fg full"><label>Notes</label><input id="v-notes" placeholder="Optional notes"></div>\n          </div>\n          <div style="display:flex;gap:8px">\n            <button class="btn bc" style="flex:1;justify-content:center" onclick="addVehicle()">+ ADD VEHICLE</button>\n            <button class="btn bgh" onclick="clearVehicleForm()">CLEAR</button>\n          </div>\n        </div>\n      </div>\n    </div>\n    <div>\n      <div class="panel">\n        <div class="ph">\n          <span class="pt">Fleet (<span id="vehCount">0</span>)</span>\n          <button class="btn bgh bsm" onclick="loadDefaultVehicles()">LOAD DEFAULTS</button>\n        </div>\n        <div class="pb" style="padding:0;overflow-x:auto">\n          <table class="rt"><thead><tr>\n            <th>ID</th><th>Name</th><th>Type</th><th>Cap(kg)</th><th>CC</th><th>Fuel</th><th>Fixed₹</th><th>₹/km</th><th></th>\n          </tr></thead>\n          <tbody id="vehBody"><tr><td colspan="9"><div class="empty"><div class="eico">🚛</div>No vehicles — click LOAD DEFAULTS</div></td></tr></tbody>\n          </table>\n        </div>\n      </div>\n    </div>\n  </div>\n</div>\n\n<!-- ⑦ COST CALC -->\n<div id="tab-cost" class="tab-content">\n  <div class="split">\n    <div>\n      <div class="panel">\n        <div class="ph"><span class="pt">Global Cost Parameters</span></div>\n        <div class="pb">\n          <div class="cost-section">\n            <div class="cost-title">Fixed Costs (per trip)</div>\n            <div class="cost-grid">\n              <div class="cost-row"><label>Base KM Included</label><input id="cp-basekm" type="number" value="25" step="any" title="Trips up to this km use fixed rate only. Extra km beyond this are charged per-km."></div>\n              <div class="cost-row"><label>Driver Allowance (₹)</label><input id="cp-driver" type="number" value="350" step="any"></div>\n              <div class="cost-row"><label>Vehicle Fixed Cost (₹/day)</label><input id="cp-vfixed" type="number" value="800" step="any"></div>\n              <div class="cost-row"><label>Loading/Unloading (₹)</label><input id="cp-load" type="number" value="150" step="any"></div>\n              <div class="cost-row"><label>Toll Charges (₹)</label><input id="cp-toll" type="number" value="0" step="any"></div>\n              <div class="cost-row"><label>Other Fixed (₹)</label><input id="cp-other" type="number" value="0" step="any"></div>\n            </div>\n          </div>\n          <div class="cost-section">\n            <div class="cost-title">Variable Costs — Beyond Base KM only</div>\n            <p style="font-size:.68rem;color:var(--muted);margin-bottom:10px">These rates apply only to km <b style="color:var(--cyan)">exceeding the Base KM</b> (default 25 km). Below base km → fixed cost only.</p>\n            <div class="cost-grid">\n              <div class="cost-row"><label>Fuel Cost (₹/km)</label><input id="cp-fuel" type="number" value="6" step="any"></div>\n              <div class="cost-row"><label>Maintenance (₹/km)</label><input id="cp-maint" type="number" value="1.5" step="any"></div>\n              <div class="cost-row"><label>Other Variable (₹/km)</label><input id="cp-ovar" type="number" value="0.5" step="any"></div>\n            </div>\n          </div>\n          <div class="cost-section">\n            <div class="cost-title">Overrides per Vehicle Type</div>\n            <p style="font-size:.72rem;color:var(--muted);margin-bottom:8px">If a vehicle has its own cost set, it overrides the global values above.</p>\n            <div style="font-family:\'Space Mono\',monospace;font-size:.65rem;color:var(--cyan)">ℹ Vehicle-specific costs set in ⑥ Vehicles tab</div>\n          </div>\n          <button class="btn bc" style="width:100%;justify-content:center;margin-top:4px" onclick="calcCosts()">⚡ CALCULATE ALL COSTS</button>\n        </div>\n      </div>\n    </div>\n    <div>\n      <div class="panel">\n        <div class="ph">\n          <span class="pt">Trip Cost Breakdown</span>\n          <span style="font-family:\'Space Mono\',monospace;font-size:.6rem;color:var(--muted)" id="costRunLbl">—</span>\n        </div>\n        <div style="overflow-x:auto">\n          <table class="cost-table">\n            <thead><tr>\n              <th>Trip</th><th>Vehicle</th><th>Stops</th><th>Total km</th><th>Base km</th><th>Extra km</th>\n              <th>Fixed Cost</th><th>Var Cost</th><th>Total Cost</th><th>₹/Stop</th><th>₹/kg</th>\n            </tr></thead>\n            <tbody id="costBody"><tr><td colspan="11"><div class="empty"><div class="eico">₹</div>Calculate costs first</div></td></tr></tbody>\n          </table>\n        </div>\n        <div id="costSummaryBar" style="display:none;padding:14px 16px;border-top:1px solid var(--border);background:rgba(0,212,255,.04)">\n          <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:10px" id="costSumGrid"></div>\n        </div>\n      </div>\n    </div>\n  </div>\n</div>\n\n<!-- ⑧ CRON -->\n<div id="tab-cron" class="tab-content">\n  <div class="split">\n    <div>\n      <div class="panel">\n        <div class="ph"><span class="pt">CRON Interval</span></div>\n        <div class="pb">\n          <div style="display:flex;gap:6px;flex-wrap:wrap;margin-bottom:10px">\n            <button class="ibtn" onclick="setCI(10,this)">10s</button>\n            <button class="ibtn active" onclick="setCI(30,this)">30s</button>\n            <button class="ibtn" onclick="setCI(60,this)">1m</button>\n            <button class="ibtn" onclick="setCI(300,this)">5m</button>\n            <button class="ibtn" onclick="setCI(600,this)">10m</button>\n            <button class="ibtn" onclick="setCI(1800,this)">30m</button>\n          </div>\n          <div style="display:flex;justify-content:space-between;font-family:\'Space Mono\',monospace;font-size:.58rem;color:var(--muted)"><span>PROGRESS</span><span id="progPct">0%</span></div>\n          <div class="prog-bar"><div class="prog-fill" id="progFill" style="width:0%"></div></div>\n        </div>\n      </div>\n      <div class="panel">\n        <div class="ph"><span class="pt">Depot Location</span></div>\n        <div class="pb">\n          <div class="form-r">\n            <div class="fg full"><label>Depot Name</label><input id="depotName" value="Central Warehouse"></div>\n            <div class="fg"><label>Latitude</label><input id="depotLat" type="number" value="12.9716" step="any"></div>\n            <div class="fg"><label>Longitude</label><input id="depotLng" type="number" value="77.5946" step="any"></div>\n          </div>\n        </div>\n      </div>\n    </div>\n    <div>\n      <div class="panel">\n        <div class="ph"><span class="pt">CRON Log</span><button class="btn bgh bsm" onclick="clearLog()">CLEAR</button></div>\n        <div class="pb"><div class="logbox" id="cronLog">\n          <div class="le linfo"><span class="lt2">[init]</span><span class="lm">LML Route Estimator v3 — Ready</span></div>\n        </div></div>\n      </div>\n    </div>\n  </div>\n</div>\n\n\n<!-- ⑨ SUBMIT TRIPS -->\n<div id="tab-submit" class="tab-content">\n  <div class="panel">\n    <div class="ph">\n      <span class="pt" style="color:var(--green)">✅ Submit Trips to System</span>\n      <span id="submitStatusLbl" style="font-family:\'Space Mono\',monospace;font-size:.6rem;color:var(--muted)">Run CRON first to generate routes</span>\n    </div>\n    <div class="pb">\n      <!-- Summary cards -->\n      <div id="submitCards" style="display:grid;grid-template-columns:repeat(auto-fill,minmax(200px,1fr));gap:12px;margin-bottom:18px">\n        <div class="empty"><div class="eico">⚡</div>Run CRON to see trip summaries</div>\n      </div>\n\n      <!-- Driver assignment table -->\n      <div id="driverAssignSection" style="display:none">\n        <div style="font-family:\'Space Mono\',monospace;font-size:.65rem;color:var(--cyan);text-transform:uppercase;letter-spacing:1.5px;margin-bottom:10px;padding-bottom:6px;border-bottom:1px solid var(--border)">\n          🚚 Assign Drivers to Trips\n        </div>\n        <div id="driverAssignTable" style="margin-bottom:18px"></div>\n      </div>\n\n      <!-- Submit button -->\n      <div id="submitBtnWrap" style="display:none;text-align:center;padding:10px 0 4px">\n        <div style="margin-bottom:14px;font-family:\'Space Mono\',monospace;font-size:.6rem;color:var(--muted)">\n          Review trips above → assign drivers → click CREATE TRIPS\n        </div>\n        <button id="submitBtn" class="btn" onclick="submitTripsToSystem()"\n          style="background:linear-gradient(135deg,var(--green),#27ae60);color:#fff;font-size:.75rem;padding:11px 28px;border-radius:10px;box-shadow:0 4px 16px rgba(46,204,113,.3);letter-spacing:1px">\n          ✅ &nbsp; CREATE TRIPS IN SYSTEM\n        </button>\n        <div id="submitProg" style="display:none;margin-top:14px">\n          <div style="font-family:\'Space Mono\',monospace;font-size:.62rem;color:var(--cyan);margin-bottom:6px" id="submitMsg">Submitting...</div>\n          <div class="prog-bar"><div class="prog-fill" id="submitFill" style="width:0%"></div></div>\n        </div>\n      </div>\n    </div>\n  </div>\n\n  <!-- Activity log for submit -->\n  <div class="panel">\n    <div class="ph"><span class="pt">Submit Log</span><button class="btn bgh bsm" onclick="document.getElementById(\'submitLog\').innerHTML=\'\'">CLEAR</button></div>\n    <div class="pb" style="padding:10px 14px">\n      <div class="logbox" id="submitLog"></div>\n    </div>\n  </div>\n</div>\n\n</div><!-- /wrap -->\n\n<!-- TRIP EDIT MODAL -->\n<div class="modal-bg" id="tripModal">\n  <div class="modal">\n    <div class="modal-hdr">\n      <span class="modal-title" id="modalTitle">Edit Trip</span>\n      <button class="modal-close" onclick="closeModal()">✕</button>\n    </div>\n    <div class="modal-body" id="modalBody"></div>\n  </div>\n</div>\n\n<script>\n// ══════════════════════════════════════════════════════\n// CONSTANTS & STATE\n// ══════════════════════════════════════════════════════\nconst TCHEX=[\'#e74c3c\',\'#e67e22\',\'#f39c12\',\'#27ae60\',\'#1abc9c\',\'#2980b9\',\'#9b59b6\',\'#e91e63\',\'#ff5722\',\'#795548\',\'#00897b\',\'#3949ab\'];\nconst TCLS =[\'tc0\',\'tc1\',\'tc2\',\'tc3\',\'tc4\',\'tc5\',\'tc6\',\'tc7\',\'tc8\',\'tc9\',\'tc10\',\'tc11\'];\n\nlet orders=[], optimizedTrips={}, vehicles=[], cronSec=30;\nlet cronTimer=null, progTimer=null, cronProg=0, oc=1;\nlet leafMap=null, mapLayers=[];\nlet editingTrip=null;\nlet costData=[];\nlet rawHeaders=[], rawRows=[];\nlet selectedMapTrip=null; // for live map trip filter\n\n// ══════════════════════════════════════════════════════\n// TABS\n// ══════════════════════════════════════════════════════\nfunction showTab(t){\n  document.querySelectorAll(\'.tab-content\').forEach(e=>e.classList.remove(\'active\'));\n  document.querySelectorAll(\'.tab\').forEach(e=>e.classList.remove(\'active\'));\n  document.getElementById(\'tab-\'+t).classList.add(\'active\');\n  const keys=[\'upload\',\'routemap\',\'map\',\'table\',\'trips\',\'vehicles\',\'cost\',\'cron\'];\n  const idx=keys.indexOf(t);\n  if(idx>=0) document.querySelectorAll(\'.tab\')[idx].classList.add(\'active\');\n  if(t===\'map\'){renderMapSidebar();setTimeout(()=>{initLeafMap();},80);}\n  if(t===\'trips\'){renderTripsOverview();populateEditDropdowns();}\n  if(t===\'cost\') calcCosts();\n}\n\n// ══════════════════════════════════════════════════════\n// LEAFLET MAP — left sidebar trip selector\n// ══════════════════════════════════════════════════════\nfunction populateMapTripFilter(){\n  // Update old select (kept for focusStop compatibility)\n  const sel=document.getElementById(\'mapTripFilter\');\n  if(sel){\n    const keys=Object.keys(optimizedTrips).sort((a,b)=>+a-+b);\n    sel.innerHTML=\'<option value="all">ALL TRIPS</option>\'+keys.map((tk,i)=>`<option value="${tk}">TRIP ${tk}</option>`).join(\'\');\n  }\n  // Populate sidebar\n  renderMapSidebar();\n}\n\nfunction renderMapSidebar(){\n  const list=document.getElementById(\'mapTripList\');\n  if(!list)return;\n  const keys=Object.keys(optimizedTrips).sort((a,b)=>+a-+b);\n  const depot=getDepot();\n  // Build cost lookup\n  const costLookup={};costData.forEach(c=>costLookup[c.trip]={km:c.km,total:c.total,extraKm:c.extraKm});\n  // Update all-trips sub\n  const allSub=document.getElementById(\'mst-all-sub\');\n  if(allSub) allSub.textContent=`${keys.length} trips · ${orders.length} stops`;\n  // Update ALL button active state\n  const allBtn=document.getElementById(\'mst-all\');\n  if(allBtn){\n    allBtn.classList.toggle(\'active-trip-btn\',!selectedMapTrip);\n    allBtn.style.borderLeft=!selectedMapTrip?\'3px solid var(--cyan)\':\'3px solid transparent\';\n  }\n  if(!keys.length){list.innerHTML=\'<div style="padding:12px 10px;font-size:.65rem;color:var(--dim);font-family:Space Mono,monospace">No trips yet</div>\';return}\n  list.innerHTML=keys.map((tk,ci)=>{\n    const route=optimizedTrips[tk]||[];\n    const cd=costLookup[tk]||{};\n    const ton=route.reduce((s,o)=>s+(+o.tonnage||0),0).toFixed(1);\n    const hex=TCHEX[ci%12];\n    const isActive=selectedMapTrip===String(tk);\n    const km=cd.km?cd.km+\' km\':\'— km\';\n    const cost=cd.total!=null?\'₹\'+cd.total:\'—\';\n    const vinfo=orders.find(o=>o.trip==tk);\n    const veh=(vinfo?.vehicle||\'\').replace(/CC\\s*[\\d.]+/i,\'\').trim()||\'Tata Ace\';\n    return `<div class="map-trip-btn${isActive?\' active-trip-btn\':\'\'}"\n      id="mst-${tk}"\n      style="border-left:3px solid ${isActive?hex:\'transparent\'}"\n      onclick="mapSelectTrip(\'${tk}\')">\n      <div class="mst-dot" style="background:${hex}"></div>\n      <div class="mst-info">\n        <div class="mst-name" style="color:${hex}">Trip ${tk}</div>\n        <div style="font-size:.58rem;color:var(--muted);margin-top:1px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${veh}</div>\n        <div style="display:flex;justify-content:space-between;margin-top:3px">\n          <span class="mst-km">${km}</span>\n          <span class="mst-cost">${cost}</span>\n        </div>\n        <div style="display:flex;justify-content:space-between;margin-top:1px">\n          <span class="mst-stops">${route.length} stops</span>\n          <span class="mst-ton">${ton}kg</span>\n        </div>\n        ${cd.extraKm>0?`<div style="font-family:\'Space Mono\',monospace;font-size:.5rem;color:var(--orange);margin-top:2px">+${cd.extraKm}km extra</div>`:\'\'}\n      </div>\n    </div>`;\n  }).join(\'\');\n}\n\nfunction mapSelectTrip(tk){\n  selectedMapTrip=(tk===\'all\')?null:String(tk);\n  // Update sidebar active states\n  document.querySelectorAll(\'.map-trip-btn\').forEach(el=>{\n    el.classList.remove(\'active-trip-btn\');\n    el.style.borderLeft=\'3px solid transparent\';\n  });\n  const allBtn=document.getElementById(\'mst-all\');\n  if(!selectedMapTrip&&allBtn){\n    allBtn.classList.add(\'active-trip-btn\');\n    allBtn.style.borderLeft=\'3px solid var(--cyan)\';\n  }\n  if(selectedMapTrip){\n    const btn=document.getElementById(`mst-${selectedMapTrip}`);\n    if(btn){\n      btn.classList.add(\'active-trip-btn\');\n      const keys=Object.keys(optimizedTrips).sort((a,b)=>+a-+b);\n      const ci=keys.indexOf(selectedMapTrip);\n      btn.style.borderLeft=`3px solid ${TCHEX[ci%12]}`;\n    }\n  }\n  initLeafMap();\n}\n\nfunction onMapTripFilter(){\n  const sel=document.getElementById(\'mapTripFilter\');\n  selectedMapTrip=sel&&sel.value!==\'all\'?sel.value:null;\n  initLeafMap();\n}\n\nfunction initLeafMap(filterTk){\n  const container=document.getElementById(\'leafMap\');\n  if(!container)return;\n  if(!leafMap){\n    leafMap=L.map(\'leafMap\',{zoomControl:true,attributionControl:false});\n    L.tileLayer(\'https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png\',{maxZoom:19}).addTo(leafMap);\n    L.control.attribution({position:\'bottomright\'}).addAttribution(\'© OpenStreetMap\').addTo(leafMap);\n  }\n  // Trigger resize so map fills new container\n  setTimeout(()=>leafMap.invalidateSize(),80);\n  clearMapLayers();\n  const depot=getDepot();\n  const allTripKeys=Object.keys(optimizedTrips).sort((a,b)=>+a-+b);\n  const activeFk=filterTk!==undefined?filterTk:selectedMapTrip;\n  const focusKeys=activeFk?allTripKeys.filter(tk=>String(tk)===String(activeFk)):allTripKeys;\n\n  if(!allTripKeys.length){leafMap.setView([depot.lat,depot.lng],12);addDepotMarker(depot);return}\n\n  const showKeys=focusKeys.length?focusKeys:allTripKeys;\n  const allPts=[{lat:depot.lat,lng:depot.lng},...showKeys.flatMap(tk=>optimizedTrips[tk]||[])];\n  const validPts=allPts.filter(p=>p.lat&&p.lng&&Math.abs(p.lat)>.01);\n  if(validPts.length){\n    const lats=validPts.map(p=>p.lat),lngs=validPts.map(p=>p.lng);\n    leafMap.fitBounds([[Math.min(...lats)-.015,Math.min(...lngs)-.015],[Math.max(...lats)+.015,Math.max(...lngs)+.015]]);\n  }\n  addDepotMarker(depot);\n\n  // Draw faded context lines for non-selected trips\n  if(activeFk){\n    allTripKeys.filter(tk=>String(tk)!==String(activeFk)).forEach((tk,ti)=>{\n      const gi=allTripKeys.indexOf(String(tk));\n      const route=(optimizedTrips[tk]||[]).filter(o=>o.lat&&o.lng);\n      const pts=[[depot.lat,depot.lng],...route.map(o=>[o.lat,o.lng]),[depot.lat,depot.lng]];\n      if(pts.length>2){const pl=L.polyline(pts,{color:TCHEX[gi%12],weight:1,opacity:.18,dashArray:\'4 6\'}).addTo(leafMap);mapLayers.push(pl);}\n    });\n  }\n\n  // Draw active trips\n  showKeys.forEach((tk)=>{\n    const gi=allTripKeys.indexOf(String(tk));\n    const hex=TCHEX[gi%12];\n    const route=(optimizedTrips[tk]||[]).filter(o=>o.lat&&o.lng);\n    const pts=[[depot.lat,depot.lng],...route.map(o=>[o.lat,o.lng]),[depot.lat,depot.lng]];\n    if(pts.length>2){\n      // Glow\n      const glow=L.polyline(pts,{color:hex,weight:activeFk?10:6,opacity:.12}).addTo(leafMap);\n      mapLayers.push(glow);\n      const pl=L.polyline(pts,{color:hex,weight:activeFk?3:2,opacity:activeFk?.9:.65,dashArray:\'9 5\'}).addTo(leafMap);\n      mapLayers.push(pl);\n    }\n    // Get trip cost info\n    const cd=costData.find(c=>String(c.trip)===String(tk))||{};\n    route.forEach((o)=>{\n      const ic=L.divIcon({className:\'\',\n        html:`<div style="background:${hex};color:#fff;border-radius:50%;width:26px;height:26px;display:flex;align-items:center;justify-content:center;font-family:\'Space Mono\',monospace;font-size:10px;font-weight:700;border:2px solid rgba(255,255,255,.7);box-shadow:0 2px 8px ${hex}aa,0 0 0 3px ${hex}33">${o.stop}</div>`,\n        iconSize:[26,26],iconAnchor:[13,13]});\n      const mk=L.marker([o.lat,o.lng],{icon:ic}).addTo(leafMap);\n      mk.bindPopup(`<div style="font-family:\'DM Sans\',sans-serif;font-size:12px;min-width:200px;background:#0d1a28;color:#d4e8f5;padding:2px">\n        <div style="font-weight:700;color:${hex};margin-bottom:5px;font-size:13px;border-bottom:1px solid #1a2e42;padding-bottom:4px">Stop ${o.stop} — Trip ${tk}</div>\n        <div style="font-weight:600;margin-bottom:2px">${o.customer}</div>\n        <div style="color:#4a6a85;font-size:10px;margin-bottom:6px">${o.address}</div>\n        <div style="display:grid;grid-template-columns:1fr 1fr;gap:3px;font-size:11px;background:#08101a;padding:7px;border-radius:6px">\n          <span style="color:#4a6a85">Crates:</span><b style="color:#00d4ff;font-family:\'Space Mono\'">${o.crates}</b>\n          <span style="color:#4a6a85">Tonnage:</span><b style="color:#ff6b2b;font-family:\'Space Mono\'">${o.tonnage}kg</b>\n          <span style="color:#4a6a85">Window:</span><b style="color:#f39c12;font-family:\'Space Mono\'">${o.window}</b>\n          <span style="color:#4a6a85">ETA:</span><b style="color:#2ecc71;font-family:\'Space Mono\'">${o.eta}</b>\n          <span style="color:#4a6a85">Leg km:</span><b style="font-family:\'Space Mono\'">${o.legDist}km</b>\n          <span style="color:#4a6a85">Priority:</span><b style="font-family:\'Space Mono\'">${o.priority||\'med\'}</b>\n        </div>\n        ${cd.total!=null?`<div style="margin-top:6px;display:flex;justify-content:space-between;font-size:11px;background:#0a1520;padding:5px 7px;border-radius:5px"><span style="color:#4a6a85">Trip dist:</span><span style="color:#00d4ff;font-family:\'Space Mono\'">${cd.km}km</span><span style="color:#4a6a85">Cost:</span><span style="color:#2ecc71;font-family:\'Space Mono\';font-weight:700">₹${cd.total}</span></div>`:\'\'}\n      </div>`,{maxWidth:240,className:\'dark-popup\'});\n      mapLayers.push(mk);\n    });\n  });\n  renderMapSidebar();\n}\n\nfunction addDepotMarker(depot){\n  if(!depot.lat||!depot.lng) return;\n  const ic=L.divIcon({className:\'\',html:`<div style="background:#ff6b2b;color:#fff;border-radius:6px;padding:3px 8px;font-family:\'Space Mono\',monospace;font-size:10px;font-weight:700;border:2px solid rgba(255,255,255,0.5);box-shadow:0 0 8px #ff6b2b88;white-space:nowrap">🏭 DEPOT</div>`,iconSize:[80,26],iconAnchor:[40,13]});\n  const mk=L.marker([depot.lat,depot.lng],{icon:ic,zIndexOffset:1000}).addTo(leafMap);\n  mk.bindPopup(`<div style="font-family:\'DM Sans\',sans-serif;font-size:12px;background:#0d1a28;color:#d4e8f5;padding:4px"><b style="color:#ff6b2b">${depot.name}</b><br><span style="color:#4a6a85">${depot.lat.toFixed(4)}, ${depot.lng.toFixed(4)}</span></div>`);\n  mapLayers.push(mk);\n}\n\nfunction clearMapLayers(){mapLayers.forEach(l=>{try{leafMap.removeLayer(l)}catch(e){}});mapLayers=[]}\nfunction mapFitAll(){selectedMapTrip=null;renderMapSidebar();if(leafMap)initLeafMap()}\n\n// ══════════════════════════════════════════════════════\n// FILE UPLOAD\n// ══════════════════════════════════════════════════════\nfunction dzOver(e){e.preventDefault();document.getElementById(\'dropZone\').classList.add(\'over\')}\nfunction dzLeave(){document.getElementById(\'dropZone\').classList.remove(\'over\')}\nfunction dzDrop(e){e.preventDefault();dzLeave();const f=e.dataTransfer.files[0];if(f)procFile(f)}\nfunction fileChg(e){const f=e.target.files[0];if(f)procFile(f)}\n\nfunction procFile(file){\n  log(`📂 ${file.name}`,\'linfo\');\n  const r=new FileReader();\n  r.onload=e=>{\n    try{\n      const wb=XLSX.read(e.target.result,{type:\'binary\'});\n      const sh=wb.Sheets[wb.SheetNames[0]];\n      const j=XLSX.utils.sheet_to_json(sh,{header:1,defval:\'\'});\n      rawHeaders=j[0].map(h=>String(h).trim());\n      rawRows=j.slice(1).filter(r=>r.some(c=>c!==\'\'));\n      document.getElementById(\'filePill\').style.display=\'flex\';\n      document.getElementById(\'fName\').textContent=file.name;\n      document.getElementById(\'fRows\').textContent=rawRows.length+\' rows\';\n      buildCM(); log(`✓ Parsed: ${rawRows.length} rows, ${rawHeaders.length} cols`,\'lok\');\n    }catch(err){log(\'✗ \'+err.message,\'lerr\')}\n  };\n  r.readAsBinaryString(file);\n}\n\nconst FLDS=[\n  {k:\'trip\',l:\'Trip #\'},{k:\'shift\',l:\'Shift\'},{k:\'vehicle\',l:\'Vehicle\'},\n  {k:\'orderid\',l:\'Order ID\'},{k:\'customer\',l:\'Customer Name\'},{k:\'address\',l:\'Address\'},\n  {k:\'lat\',l:\'Latitude\'},{k:\'lng\',l:\'Longitude\'},{k:\'crates\',l:\'Crates\'},\n  {k:\'tonnage\',l:\'Tonnage (kg)\'},{k:\'window\',l:\'Time Window\'},{k:\'priority\',l:\'Priority\'}\n];\n\nfunction buildCM(){\n  const g=document.getElementById(\'cmGrid\');\n  g.innerHTML=\'\';\n  const autoMap={\n    trip:/trip|route/i,shift:/shift/i,vehicle:/vehicle|truck|cc/i,\n    orderid:/order.?id|ord|^id$/i,customer:/customer|name|client|shop|store/i,\n    address:/address|addr|location/i,lat:/^lat/i,lng:/^l(ng|on)/i,\n    crates:/crate|box|qty/i,tonnage:/ton|kg|weight/i,\n    window:/window|time|slot/i,priority:/priority|prio/i\n  };\n  FLDS.forEach(f=>{\n    const best=rawHeaders.findIndex(h=>autoMap[f.k]&&autoMap[f.k].test(h));\n    const d=document.createElement(\'div\');d.className=\'cm-item\';\n    d.innerHTML=`<label>${f.l}</label><select id="cm-${f.k}" style="background:var(--card2);border:1px solid var(--border);color:var(--text);padding:6px 9px;border-radius:6px;font-size:.78rem;outline:none">\n      <option value="">— skip —</option>\n      ${rawHeaders.map((h,i)=>`<option value="${i}"${i===best?\' selected\':\'\'}>${h||(i+1)}</option>`).join(\'\')}\n    </select>`;\n    g.appendChild(d);\n  });\n  document.getElementById(\'cmPanel\').style.display=\'block\';\n}\n\nfunction applyMap(){\n  const mp={};\n  FLDS.forEach(f=>{const v=document.getElementById(`cm-${f.k}`)?.value;if(v!==\'\')mp[f.k]=+v});\n  if(!mp.customer){log(\'⚠ Map Customer column\',\'lwarn\');return}\n  orders=[];\n  rawRows.forEach((row,ri)=>{\n    const g=k=>mp[k]!==undefined?String(row[mp[k]]||\'\').trim():\'\';\n    orders.push({\n      id:g(\'orderid\')||`ORD-${String(ri+1).padStart(3,\'0\')}`,\n      trip:+g(\'trip\')||1,shift:+g(\'shift\')||1,vehicle:g(\'vehicle\')||\'\',\n      customer:g(\'customer\')||`Cust ${ri+1}`,address:g(\'address\')||\'—\',\n      lat:+g(\'lat\')||0,lng:+g(\'lng\')||0,crates:+g(\'crates\')||0,\n      tonnage:+g(\'tonnage\')||0,window:g(\'window\')||\'07:00-08:00\',priority:g(\'priority\')||\'med\'\n    });\n  });\n  oc=orders.length+1;\n  renderPreview();updateStats();\n  log(`✓ Loaded ${orders.length} orders`,\'lok\');\n  runCron(); showTab(\'routemap\');\n}\n\n// ══════════════════════════════════════════════════════\n// MANUAL ORDER\n// ══════════════════════════════════════════════════════\nfunction addManual(){\n  const g=id=>document.getElementById(id)?.value?.trim()||\'\';\n  const gn=id=>+document.getElementById(id)?.value||0;\n  if(!g(\'m-cust\')){log(\'⚠ Customer name required\',\'lwarn\');return}\n  const vSel=document.getElementById(\'m-veh\');\n  const vName=vSel?vSel.options[vSel.selectedIndex]?.text:\'\';\n  orders.push({\n    id:g(\'m-oid\')||`ORD-${String(oc).padStart(3,\'0\')}`,\n    trip:gn(\'m-trip\')||1,shift:gn(\'m-shift\')||1,\n    vehicle:vName||g(\'m-veh\'),customer:g(\'m-cust\'),address:g(\'m-addr\')||\'—\',\n    lat:gn(\'m-lat\'),lng:gn(\'m-lng\'),crates:gn(\'m-cr\'),tonnage:gn(\'m-tn\'),\n    window:g(\'m-tw\')||\'07:00-08:00\',priority:g(\'m-pri\')||\'med\'\n  });\n  oc++;\n  renderPreview();updateStats();\n  log(`✓ Added: ${orders[orders.length-1].customer}`,\'lok\');\n  [\'m-trip\',\'m-shift\',\'m-oid\',\'m-cust\',\'m-addr\',\'m-lat\',\'m-lng\',\'m-cr\',\'m-tn\',\'m-tw\'].forEach(id=>{const el=document.getElementById(id);if(el)el.value=\'\'});\n}\n\nfunction removeOrder(i){const n=orders[i].customer;orders.splice(i,1);renderPreview();updateStats();log(\'✗ Removed: \'+n,\'lwarn\')}\nfunction clearAll(){orders=[];optimizedTrips={};renderPreview();updateStats();renderRoutemapCards({});renderRouteTable([]);if(leafMap)clearMapLayers();log(\'Cleared all\',\'lwarn\')}\n\n// ══════════════════════════════════════════════════════\n// RENDER PREVIEW\n// ══════════════════════════════════════════════════════\nfunction renderPreview(){\n  document.getElementById(\'ordCount\').textContent=orders.length;\n  const tb=document.getElementById(\'prevBody\');\n  const trips=[...new Set(orders.map(o=>o.trip))].sort((a,b)=>+a-+b);\n  const ti={}; trips.forEach((t,i)=>ti[t]=i);\n  if(!orders.length){tb.innerHTML=\'<tr><td colspan="8"><div class="empty"><div class="eico">📋</div>No data</div></td></tr>\';return}\n  tb.innerHTML=orders.map((o,i)=>`<tr>\n    <td><span class="trip-tag ${TCLS[ti[o.trip]%12]}">${o.trip}</span></td>\n    <td style="font-weight:600;max-width:120px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${o.customer}</td>\n    <td style="font-size:.68rem;color:var(--muted);max-width:100px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${o.address}</td>\n    <td style="font-family:\'Space Mono\',monospace;color:var(--cyan)">${o.crates}</td>\n    <td style="font-family:\'Space Mono\',monospace;color:var(--orange)">${o.tonnage}</td>\n    <td style="font-family:\'Space Mono\',monospace;font-size:.62rem;color:var(--yellow)">${o.window}</td>\n    <td style="font-family:\'Space Mono\',monospace;font-size:.58rem;color:var(--muted)">${o.lat?o.lat.toFixed(4):\'—\'},${o.lng?o.lng.toFixed(4):\'—\'}</td>\n    <td><button class="btn bred bsm" onclick="removeOrder(${i})">✕</button></td>\n  </tr>`).join(\'\');\n}\n\n// ══════════════════════════════════════════════════════\n// STATS\n// ══════════════════════════════════════════════════════\nfunction updateStats(){\n  document.getElementById(\'s-ord\').textContent=orders.length;\n  document.getElementById(\'s-trips\').textContent=[...new Set(orders.map(o=>o.trip))].length;\n  document.getElementById(\'s-crates\').textContent=orders.reduce((s,o)=>s+o.crates,0).toFixed(0);\n  document.getElementById(\'s-ton\').textContent=orders.reduce((s,o)=>s+o.tonnage,0).toFixed(1);\n}\n\n// ══════════════════════════════════════════════════════\n// ROUTE OPTIMIZATION\n// ══════════════════════════════════════════════════════\nfunction getDepot(){return{lat:+(document.getElementById(\'depotLat\')?.value)||12.9716,lng:+(document.getElementById(\'depotLng\')?.value)||77.5946,name:document.getElementById(\'depotName\')?.value||\'Warehouse\'}}\n\nfunction hav(a,b){\n  if(!a.lat||!b.lat||!a.lng||!b.lng)return 5;\n  const R=6371,dLa=(b.lat-a.lat)*Math.PI/180,dLn=(b.lng-a.lng)*Math.PI/180;\n  const s=Math.sin(dLa/2)**2+Math.cos(a.lat*Math.PI/180)*Math.cos(b.lat*Math.PI/180)*Math.sin(dLn/2)**2;\n  return R*2*Math.atan2(Math.sqrt(s),Math.sqrt(1-s));\n}\n\n// ══════════════════════════════════════════════════════\n// ROUTE OPTIMIZATION ENGINE — OR-Quality TSP\n// Stages: 1) Time-window sort  2) Priority sort\n//         3) Nearest-Neighbor seed  4) 2-opt improvement\n//         5) Or-opt (1-move & 2-move)  6) Return distance\n// ══════════════════════════════════════════════════════\n\n// Parse "HH:MM-HH:MM" → start minutes from midnight\nfunction parseWindowStart(win){\n  if(!win||typeof win!==\'string\')return 480;\n  const m=win.match(/(\\d{1,2}):(\\d{2})/);\n  return m?+m[1]*60+(+m[2]):480;\n}\n\n// Priority weight: high = visit early\nconst PW={high:0,med:2,low:4};\n\n// Total route distance (depot→stops→depot)\nfunction routeDist(stops,depot){\n  if(!stops.length)return 0;\n  let d=hav(depot,stops[0]);\n  for(let i=1;i<stops.length;i++)d+=hav(stops[i-1],stops[i]);\n  d+=hav(stops[stops.length-1],depot);\n  return d;\n}\n\n// ── Stage 1: Smart greedy seed ──\n// Visits high-priority & earliest time-window stops first,\n// then nearest neighbour with combined cost function\nfunction greedySeed(stops,depot){\n  if(!stops.length)return[];\n  // Pre-sort by priority + window start as a first pass guide\n  const sorted=[...stops].sort((a,b)=>{\n    const pa=(PW[a.priority]||2),pb=(PW[b.priority]||2);\n    if(pa!==pb)return pa-pb;\n    return parseWindowStart(a.window)-parseWindowStart(b.window);\n  });\n  let unvis=[...sorted],route=[],cur=depot;\n  while(unvis.length){\n    let best=null,bs=Infinity;\n    unvis.forEach(o=>{\n      const d=hav(cur,o);\n      const ws=parseWindowStart(o.window);\n      const pw=PW[o.priority]||2;\n      // Combined score: distance + priority penalty + time-window urgency\n      const score=d+(pw*1.5)+(ws/60)*0.3;\n      if(score<bs){bs=score;best=o}\n    });\n    unvis=unvis.filter(o=>o!==best);\n    route.push(best);cur=best;\n  }\n  return route;\n}\n\n// ── Stage 2: 2-opt improvement ──\n// Repeatedly reverse sub-segments if it shortens total distance\nfunction twoOpt(route,depot){\n  if(route.length<4)return route;\n  let improved=true,best=[...route],bestD=routeDist(best,depot);\n  let iters=0;\n  while(improved&&iters<200){\n    improved=false;iters++;\n    for(let i=0;i<best.length-1;i++){\n      for(let j=i+2;j<best.length;j++){\n        // Reverse segment [i+1..j]\n        const candidate=[...best.slice(0,i+1),...best.slice(i+1,j+1).reverse(),...best.slice(j+1)];\n        const d=routeDist(candidate,depot);\n        if(d<bestD-0.0001){bestD=d;best=candidate;improved=true;}\n      }\n    }\n  }\n  return best;\n}\n\n// ── Stage 3: Or-opt (relocate single stops) ──\n// Try moving each stop to every other position; keep best improvement\nfunction orOpt1(route,depot){\n  if(route.length<3)return route;\n  let best=[...route],bestD=routeDist(best,depot),improved=true,iters=0;\n  while(improved&&iters<150){\n    improved=false;iters++;\n    for(let i=0;i<best.length;i++){\n      const node=best[i];\n      const without=best.filter((_,idx)=>idx!==i);\n      for(let j=0;j<=without.length;j++){\n        const cand=[...without.slice(0,j),node,...without.slice(j)];\n        const d=routeDist(cand,depot);\n        if(d<bestD-0.0001){bestD=d;best=cand;improved=true;break;}\n      }\n      if(improved)break;\n    }\n  }\n  return best;\n}\n\n// ── Stage 4: Or-opt-2 (relocate consecutive pairs) ──\nfunction orOpt2(route,depot){\n  if(route.length<4)return route;\n  let best=[...route],bestD=routeDist(best,depot),improved=true,iters=0;\n  while(improved&&iters<100){\n    improved=false;iters++;\n    for(let i=0;i<best.length-1;i++){\n      const pair=[best[i],best[i+1]];\n      const without=best.filter((_,idx)=>idx!==i&&idx!==i+1);\n      for(let j=0;j<=without.length;j++){\n        const cand=[...without.slice(0,j),...pair,...without.slice(j)];\n        const d=routeDist(cand,depot);\n        if(d<bestD-0.0001){bestD=d;best=cand;improved=true;break;}\n      }\n      if(improved)break;\n    }\n  }\n  return best;\n}\n\n// ── Stage 5: Re-apply priority constraint ──\n// After geometry is optimal, if a HIGH-priority stop is too far back,\n// pull it forward (soft enforcement: only if it costs < 10% more distance)\nfunction enforcePriority(route,depot){\n  let r=[...route];\n  for(let i=r.length-1;i>0;i--){\n    if((PW[r[i].priority]||2)<(PW[r[i-1].priority]||2)){\n      // Higher priority stop is behind lower — try swap\n      const swapped=[...r];[swapped[i],swapped[i-1]]=[swapped[i-1],swapped[i]];\n      if(routeDist(swapped,depot)<=routeDist(r,depot)*1.08){\n        r=swapped;\n      }\n    }\n  }\n  return r;\n}\n\n// ── Master optimizer: runs all stages ──\nfunction optTrip(stops){\n  if(!stops.length)return[];\n  const depot=getDepot();\n  // Filter stops with no coords — place them at end\n  const withCoords=stops.filter(o=>o.lat&&o.lng&&Math.abs(o.lat)>0.001);\n  const noCoords=stops.filter(o=>!o.lat||!o.lng||Math.abs(o.lat)<=0.001);\n\n  let route=greedySeed(withCoords,depot);\n  if(route.length>=4) route=twoOpt(route,depot);\n  if(route.length>=3) route=orOpt1(route,depot);\n  if(route.length>=4) route=orOpt2(route,depot);\n  route=enforcePriority(route,depot);\n  // Append no-coord stops at end\n  route=[...route,...noCoords];\n\n  const before=routeDist(greedySeed(withCoords,depot),depot).toFixed(2);\n  const after=routeDist(route.filter(o=>o.lat&&o.lng),depot).toFixed(2);\n  log(`Trip opt: seed ${before}km → optimized ${after}km (${((before-after)/before*100).toFixed(1)}% saved)`,\'lok\');\n  return route;\n}\n\nfunction buildRouteMeta(route){\n  const depot=getDepot();\n  let cum=0,prev=depot;\n  // Start time from depot at 07:00\n  let elapsedMin=0;\n  return route.map((o,i)=>{\n    const d=hav(prev,o);cum+=d;\n    elapsedMin+=Math.round(d/35*60);\n    // Add ~5min stop time per customer\n    elapsedMin+=5;\n    const hh=7+Math.floor(elapsedMin/60), mm=elapsedMin%60;\n    const eta=`${String(hh).padStart(2,\'0\')}:${String(mm).padStart(2,\'0\')}`;\n    prev=o;\n    return{...o,stop:i+1,legDist:d.toFixed(2),cumDist:cum.toFixed(2),eta,_cumKm:cum};\n  });\n}\n\nfunction runCron(){\n  log(`▶ CRON — ${orders.length} orders`,\'linfo\');\n  if(!orders.length){log(\'⚠ No orders loaded\',\'lwarn\');return}\n  const depot=getDepot();\n\n  // ── FIX 4: Re-number trips ascending 1,2,3... ──\n  const rawTripKeys=[...new Set(orders.map(o=>+o.trip))].sort((a,b)=>a-b);\n  const tripRemap={};rawTripKeys.forEach((tk,i)=>{tripRemap[tk]=i+1});\n  orders.forEach(o=>{o.trip=tripRemap[o.trip]||o.trip});\n\n  const tripKeys=[...new Set(orders.map(o=>o.trip))].sort((a,b)=>a-b);\n  let grand=0; optimizedTrips={};\n  tripKeys.forEach(tk=>{\n    const sorted=optTrip(orders.filter(o=>o.trip==tk));\n    const meta=buildRouteMeta(sorted);\n    optimizedTrips[tk]=meta;\n    const lastStop=meta[meta.length-1];\n    grand+=(lastStop?+lastStop.cumDist+hav(lastStop,depot):0);\n    log(`✓ Trip ${tk}: ${meta.length} stops`,\'lok\');\n  });\n  const gkm=grand.toFixed(1),gm=Math.round(grand/35*60);\n  document.getElementById(\'s-dist\').textContent=gkm+\'km\';\n  const now=new Date().toLocaleTimeString(\'en-US\',{hour12:false});\n  document.getElementById(\'lRunLbl\').textContent=\'LAST RUN: \'+now;\n  document.getElementById(\'rmSub\').textContent=`${tripKeys.length} trips · ${orders.length} stops · ${gkm}km · ${now} · 2-OPT+OR-OPT`;\n  renderRoutemapCards(optimizedTrips);\n  renderRouteTable(Object.values(optimizedTrips).flat());\n  renderTripsOverview();\n  populateMapTripFilter();\n  if(leafMap&&document.getElementById(\'tab-map\').classList.contains(\'active\'))initLeafMap();\n  calcCosts();\n  resetProg();\n  log(`✓ Done: ${tripKeys.length} trips, ${gkm}km`,\'lok\');\n  updateStats();\n}\n\n// ══════════════════════════════════════════════════════\n// ROUTEMAP CARDS — full drag-and-drop within trip & between trips\n// ══════════════════════════════════════════════════════\nlet rmDrag={srcTrip:null,srcIdx:null};\n\nfunction renderRoutemapCards(trips){\n  const g=document.getElementById(\'rmGrid\');\n  const keys=Object.keys(trips).sort((a,b)=>+a-+b);\n  if(!keys.length){g.innerHTML=\'<div class="empty"><div class="eico">🗺️</div>Upload data and run CRON</div>\';return}\n  const vinfo={};orders.forEach(o=>{vinfo[o.trip]=vinfo[o.trip]||{shift:o.shift,vehicle:o.vehicle}});\n  // Build cost lookup from costData\n  const costLookup={};costData.forEach(c=>costLookup[c.trip]={km:c.km,total:c.total,extraKm:c.extraKm});\n  const depot=getDepot();\n  g.innerHTML=keys.map((tk,ci)=>{\n    const custs=trips[tk],vi=vinfo[tk]||{},ton=custs.reduce((s,c)=>s+(+c.tonnage||0),0).toFixed(1);\n    const vn=(vi.vehicle||\'Tata Ace\').replace(/CC\\s*[\\d.]+/i,\'\').trim();\n    const ccm=(vi.vehicle||\'\').match(/CC\\s*([\\d.]+)/i);\n    const cc=ccm?ccm[1]:\'—\';\n    // Get km and cost for this trip\n    const cd=costLookup[tk]||{};\n    const km=cd.km?cd.km+\' km\':\'—\';\n    const cost=cd.total!=null?\'₹\'+cd.total:\'—\';\n    const extraTag=cd.extraKm>0?`<span style="font-size:.5rem;background:rgba(255,107,43,.25);color:#ff6b2b;padding:1px 4px;border-radius:3px;margin-left:3px">+${cd.extraKm}km</span>`:\'\';\n    return `<div class="trip-col" id="tripcol-${tk}"\n        ondragover="rmColOver(event,\'${tk}\')"\n        ondragleave="rmColLeave(event,\'${tk}\')"\n        ondrop="rmColDrop(event,\'${tk}\')">\n      <div class="trip-hdr ${TCLS[ci%12]}" onclick="openTripModal(${tk})">\n        <div class="trip-title">Trip ${tk} | Shift ${vi.shift||1}</div>\n        <div class="trip-cc">${vn}</div>\n        <div class="trip-cc">CC ${cc} <span style="opacity:.45;font-size:.42rem">⠿ drag stops</span></div>\n      </div>\n      <div style="background:rgba(0,0,0,.45);padding:3px 8px;display:flex;justify-content:space-between;align-items:center;border-bottom:1px solid rgba(255,255,255,.06);font-family:\'Space Mono\',monospace;font-size:.58rem">\n        <span style="color:#00d4ff">${km}${extraTag}</span>\n        <span style="color:#2ecc71;font-weight:700">${cost}</span>\n      </div>\n      <div class="trip-ton" id="tripton-${tk}"><span>Tonnage: ${ton}</span><span style="opacity:.6;font-size:.55rem">${custs.length} stops</span></div>\n      <div class="trip-customers" id="tripcusts-${tk}">\n        ${custs.map((c,i)=>`\n          <div class="cust-card"\n            id="custcard-${tk}-${i}"\n            draggable="true"\n            ondragstart="rmDragStart(event,\'${tk}\',${i})"\n            ondragend="rmDragEnd()"\n            ondragover="rmCardOver(event,\'${tk}\',${i})"\n            ondragleave="rmCardLeave(\'${tk}\',${i})"\n            ondrop="rmCardDrop(event,\'${tk}\',${i})">\n            <div class="cust-name"><span class="snum">${c.stop}</span>${c.customer}</div>\n            <div class="cust-meta">\n              <span class="cr">Crates: ${c.crates}</span>&nbsp;\n              <span class="ti">${c.window}</span><br>\n              <span class="tn">Tonnage:${c.tonnage}</span>\n            </div>\n          </div>`).join(\'\')}\n      </div>\n    </div>`;\n  }).join(\'\');\n}\n\nfunction rmDragStart(e,tk,idx){\n  rmDrag={srcTrip:String(tk),srcIdx:idx};\n  e.dataTransfer.effectAllowed=\'move\';\n  e.dataTransfer.setData(\'text/plain\',`${tk}:${idx}`);\n  setTimeout(()=>{const el=document.getElementById(`custcard-${tk}-${idx}`);if(el)el.classList.add(\'dragging\');},0);\n}\nfunction rmDragEnd(){\n  document.querySelectorAll(\'.cust-card\').forEach(el=>el.classList.remove(\'dragging\',\'drag-over\'));\n  document.querySelectorAll(\'.trip-customers\').forEach(el=>el.classList.remove(\'drop-target\'));\n  document.querySelectorAll(\'.trip-col\').forEach(el=>el.classList.remove(\'drag-over-col\'));\n  rmDrag={srcTrip:null,srcIdx:null};\n}\nfunction rmCardOver(e,tk,idx){\n  e.preventDefault();e.stopPropagation();\n  e.dataTransfer.dropEffect=\'move\';\n  if(rmDrag.srcTrip===String(tk)&&rmDrag.srcIdx===idx)return;\n  document.querySelectorAll(\'.cust-card\').forEach(el=>el.classList.remove(\'drag-over\'));\n  document.getElementById(`custcard-${tk}-${idx}`)?.classList.add(\'drag-over\');\n}\nfunction rmCardLeave(tk,idx){document.getElementById(`custcard-${tk}-${idx}`)?.classList.remove(\'drag-over\')}\nfunction rmCardDrop(e,destTrip,destIdx){\n  e.preventDefault();e.stopPropagation();\n  const {srcTrip,srcIdx}=rmDrag;\n  if(srcTrip===null)return;\n  rmMoveStop(srcTrip,srcIdx,String(destTrip),destIdx);\n}\nfunction rmColOver(e,tk){\n  e.preventDefault();\n  if(rmDrag.srcTrip&&rmDrag.srcTrip!==String(tk)){\n    document.getElementById(`tripcol-${tk}`)?.classList.add(\'drag-over-col\');\n    document.getElementById(`tripcusts-${tk}`)?.classList.add(\'drop-target\');\n  }\n}\nfunction rmColLeave(e,tk){\n  document.getElementById(`tripcol-${tk}`)?.classList.remove(\'drag-over-col\');\n  document.getElementById(`tripcusts-${tk}`)?.classList.remove(\'drop-target\');\n}\nfunction rmColDrop(e,destTrip){\n  e.preventDefault();\n  const {srcTrip,srcIdx}=rmDrag;\n  if(!srcTrip||srcTrip===String(destTrip))return;\n  const destLen=(optimizedTrips[destTrip]||[]).length;\n  rmMoveStop(srcTrip,srcIdx,String(destTrip),destLen);\n}\nfunction rmMoveStop(srcTrip,srcIdx,destTrip,destIdx){\n  if(!optimizedTrips[srcTrip])return;\n  const srcArr=[...optimizedTrips[srcTrip]];\n  const [moved]=srcArr.splice(srcIdx,1);\n  if(srcTrip===destTrip){\n    const ins=destIdx>srcIdx?destIdx-1:destIdx;\n    srcArr.splice(Math.max(0,Math.min(ins,srcArr.length)),0,moved);\n    optimizedTrips[srcTrip]=srcArr.map((o,i)=>({...o,stop:i+1}));\n    log(`Trip ${srcTrip}: moved "${moved.customer}" → stop ${ins+1}`,\'linfo\');\n  }else{\n    moved.trip=+destTrip;\n    const oi=orders.findIndex(o=>o.id===moved.id);\n    if(oi>=0)orders[oi].trip=+destTrip;\n    optimizedTrips[srcTrip]=srcArr.map((o,i)=>({...o,stop:i+1}));\n    if(!optimizedTrips[destTrip])optimizedTrips[destTrip]=[];\n    const destArr=[...optimizedTrips[destTrip]];\n    destArr.splice(Math.min(destIdx,destArr.length),0,moved);\n    optimizedTrips[destTrip]=destArr.map((o,i)=>({...o,stop:i+1}));\n    if(!optimizedTrips[srcTrip].length){delete optimizedTrips[srcTrip];orders=orders.filter(o=>o.trip!=+srcTrip);}\n    log(`Moved "${moved.customer}": Trip ${srcTrip} → Trip ${destTrip}`,\'lok\');\n  }\n  renderRoutemapCards(optimizedTrips);\n  renderRouteTable(Object.values(optimizedTrips).flat());\n  renderTripsOverview();updateStats();\n}\nfunction updateTripTonnage(tk){\n  const el=document.getElementById(`tripton-${tk}`);\n  if(!el||!optimizedTrips[tk])return;\n  const ton=optimizedTrips[tk].reduce((s,o)=>s+(+o.tonnage||0),0).toFixed(1);\n  el.innerHTML=`<span>Tonnage: ${ton}</span><span style="opacity:.6;font-size:.55rem">${optimizedTrips[tk].length} stops</span>`;\n}\n\nfunction focusStop(tk,idx){\n  selectedMapTrip=String(tk);\n  showTab(\'map\');\n  setTimeout(()=>{\n    renderMapSidebar();\n    initLeafMap(String(tk));\n    const route=optimizedTrips[tk];\n    if(route&&route[idx]&&leafMap)setTimeout(()=>leafMap.setView([route[idx].lat,route[idx].lng],15),350);\n  },180);\n}\n\n\n// ══════════════════════════════════════════════════════\n// ROUTE TABLE\n// ══════════════════════════════════════════════════════\nfunction renderRouteTable(rows){\n  const tb=document.getElementById(\'routeBody\');\n  if(!rows.length){tb.innerHTML=\'<tr><td colspan="12"><div class="empty"><div class="eico">📍</div>Run CRON</div></td></tr>\';return}\n  const trips=[...new Set(rows.map(r=>r.trip))].sort((a,b)=>+a-+b);\n  const ti={};trips.forEach((t,i)=>ti[t]=i);\n  tb.innerHTML=rows.map(r=>`<tr>\n    <td><div style="display:inline-flex;align-items:center;justify-content:center;width:20px;height:20px;border-radius:50%;background:var(--cyan);color:#08101a;font-family:\'Space Mono\',monospace;font-size:.58rem;font-weight:700">${r.stop}</div></td>\n    <td><span class="trip-tag ${TCLS[ti[r.trip]%12]}">${r.trip}</span></td>\n    <td style="font-family:\'Space Mono\',monospace;font-size:.6rem;color:var(--cyan)">${r.id}</td>\n    <td style="font-weight:600">${r.customer}</td>\n    <td style="font-size:.68rem;color:var(--muted)">${r.address}</td>\n    <td style="font-family:\'Space Mono\',monospace;color:var(--cyan)">${r.crates}</td>\n    <td style="font-family:\'Space Mono\',monospace;color:var(--orange)">${r.tonnage}</td>\n    <td style="font-family:\'Space Mono\',monospace;font-size:.62rem;color:var(--yellow)">${r.window}</td>\n    <td><span class="${r.priority===\'high\'?\'pb-h\':r.priority===\'low\'?\'pb-l\':\'pb-m\'}">${r.priority?.toUpperCase()||\'MED\'}</span></td>\n    <td style="font-family:\'Space Mono\',monospace;font-size:.65rem">${r.legDist} km</td>\n    <td style="font-family:\'Space Mono\',monospace;font-size:.65rem;color:var(--muted)">${r.cumDist} km</td>\n    <td style="font-family:\'Space Mono\',monospace;color:var(--green)">${r.eta}</td>\n  </tr>`).join(\'\');\n}\n\n// ══════════════════════════════════════════════════════\n// TRIP EDITOR\n// ══════════════════════════════════════════════════════\nfunction populateEditDropdowns(){\n  const trips=[...new Set(orders.map(o=>o.trip))].sort((a,b)=>+a-+b);\n  // Vehicle selects\n  [\'m-veh\',\'et-veh\'].forEach(id=>{\n    const sel=document.getElementById(id);\n    if(!sel)return;\n    sel.innerHTML=\'<option value="">-- select --</option>\'+vehicles.map(v=>`<option value="${v.id}">${v.name} (${v.cc||v.type})</option>`).join(\'\');\n  });\n  // Add-to-trip customer select\n  const csel=document.getElementById(\'add-cust-sel\');\n  if(csel) csel.innerHTML=\'<option value="">-- select customer --</option>\'+orders.map((o,i)=>`<option value="${i}">${o.customer} (Trip ${o.trip})</option>`).join(\'\');\n}\n\nfunction loadTripForEdit(){\n  const tk=+document.getElementById(\'et-num\').value;\n  if(!tk){log(\'⚠ Enter trip number\',\'lwarn\');return}\n  const tripOrders=orders.filter(o=>o.trip==tk);\n  if(!tripOrders.length){log(`⚠ No orders in trip ${tk}`,\'lwarn\');return}\n  editingTrip=tk;\n  renderTripEditArea(tk);\n  document.getElementById(\'tripEditActions\').style.display=\'flex\';\n  log(`Loaded Trip ${tk} for editing`,\'linfo\');\n}\n\nfunction renderTripEditArea(tk){\n  const tripOrders=optimizedTrips[tk]||orders.filter(o=>o.trip==tk);\n  const area=document.getElementById(\'tripEditArea\');\n  area.innerHTML=`\n    <div style="margin-bottom:8px;font-family:\'Space Mono\',monospace;font-size:.62rem;color:var(--cyan)">TRIP ${tk} — ${tripOrders.length} customers (drag to reorder)</div>\n    <div class="te-row te-hdr">\n      <div style="text-align:center">⠿</div><div>Customer</div><div>Window</div>\n      <div>Crates</div><div>Tonnage</div><div>Priority</div><div>Address</div><div>Lat,Lng</div><div></div>\n    </div>\n    <div id="teList">\n      ${tripOrders.map((o,i)=>`\n        <div class="te-row" draggable="true" data-idx="${i}" ondragstart="teDragStart(event,${i})" ondragover="teDragOver(event,${i})" ondrop="teDrop(event,${i})">\n          <div class="drag-handle te-num">⠿${o.stop||i+1}</div>\n          <div><input class="te-inp" value="${o.customer}" onchange="updateTripCell(${i},\'customer\',this.value)"></div>\n          <div><input class="te-inp" value="${o.window}" onchange="updateTripCell(${i},\'window\',this.value)" style="width:75px"></div>\n          <div><input class="te-inp" type="number" value="${o.crates}" onchange="updateTripCell(${i},\'crates\',+this.value)" style="width:65px"></div>\n          <div><input class="te-inp" type="number" value="${o.tonnage}" onchange="updateTripCell(${i},\'tonnage\',+this.value)" style="width:65px"></div>\n          <div><select class="te-inp" onchange="updateTripCell(${i},\'priority\',this.value)" style="width:72px">\n            <option${o.priority===\'high\'?\' selected\':\'\'}>high</option>\n            <option${!o.priority||o.priority===\'med\'?\' selected\':\'\'}>med</option>\n            <option${o.priority===\'low\'?\' selected\':\'\'}>low</option>\n          </select></div>\n          <div><input class="te-inp" value="${o.address}" onchange="updateTripCell(${i},\'address\',this.value)" style="width:100%"></div>\n          <div style="font-family:\'Space Mono\',monospace;font-size:.55rem;color:var(--muted)">${o.lat?o.lat.toFixed(4):\'0\'},${o.lng?o.lng.toFixed(4):\'0\'}</div>\n          <div><button class="btn bred bsm" onclick="removeFromTrip(${tk},${i})">✕</button></div>\n        </div>`).join(\'\')}\n    </div>`;\n}\n\nlet dragSrcIdx=null;\nfunction teDragStart(e,i){dragSrcIdx=i;e.currentTarget.style.opacity=\'.4\'}\nfunction teDragOver(e,i){e.preventDefault()}\nfunction teDrop(e,i){\n  e.preventDefault();\n  if(dragSrcIdx===null||dragSrcIdx===i)return;\n  const tk=editingTrip;\n  const src=optimizedTrips[tk]||orders.filter(o=>o.trip==tk);\n  const arr=[...src];\n  const [moved]=arr.splice(dragSrcIdx,1);\n  arr.splice(i,0,moved);\n  if(optimizedTrips[tk]) optimizedTrips[tk]=arr.map((o,idx)=>({...o,stop:idx+1}));\n  else orders.filter(o=>o.trip==tk).forEach((o,idx)=>{o._manualOrder=idx});\n  dragSrcIdx=null;\n  renderTripEditArea(tk);\n  log(`Trip ${tk}: reordered stop ${dragSrcIdx} → ${i}`,\'linfo\');\n}\n\nfunction updateTripCell(i,field,val){\n  if(!editingTrip)return;\n  const src=optimizedTrips[editingTrip];\n  if(src&&src[i]){src[i][field]=val;const oi=orders.findIndex(o=>o.id===src[i].id);if(oi>=0)orders[oi][field]=val;}\n}\n\nfunction removeFromTrip(tk,i){\n  const src=optimizedTrips[tk];\n  if(!src)return;\n  const oid=src[i].id;\n  optimizedTrips[tk]=src.filter((_,idx)=>idx!==i).map((o,j)=>({...o,stop:j+1}));\n  const oi=orders.findIndex(o=>o.id===oid);\n  if(oi>=0)orders.splice(oi,1);\n  renderTripEditArea(tk);renderPreview();updateStats();renderTripsOverview();\n  log(`Removed stop from Trip ${tk}`,\'lwarn\');\n}\n\nfunction saveTripEdit(){\n  if(!editingTrip)return;\n  renderRoutemapCards(optimizedTrips);\n  renderRouteTable(Object.values(optimizedTrips).flat());\n  renderTripsOverview();\n  if(leafMap&&document.getElementById(\'tab-map\').classList.contains(\'active\'))initLeafMap();\n  log(`✓ Trip ${editingTrip} saved`,\'lok\');\n}\n\nfunction deleteTrip(){\n  if(!editingTrip)return;\n  if(!confirm(`Delete Trip ${editingTrip} and remove all its orders?`))return;\n  orders=orders.filter(o=>o.trip!=editingTrip);\n  delete optimizedTrips[editingTrip];\n  editingTrip=null;\n  document.getElementById(\'tripEditArea\').innerHTML=\'<div class="empty"><div class="eico">✏️</div>Trip deleted</div>\';\n  document.getElementById(\'tripEditActions\').style.display=\'none\';\n  renderPreview();updateStats();renderRoutemapCards(optimizedTrips);renderTripsOverview();\n  log(\'Trip deleted\',\'lwarn\');\n}\n\nfunction reoptimizeTrip(){\n  if(!editingTrip)return;\n  const tripOrders=orders.filter(o=>o.trip==editingTrip);\n  const meta=buildRouteMeta(optTrip(tripOrders));\n  optimizedTrips[editingTrip]=meta;\n  renderTripEditArea(editingTrip);\n  log(`Re-optimized Trip ${editingTrip}`,\'lok\');\n}\n\nfunction newTripFromOrders(){\n  const maxTrip=orders.length?Math.max(...orders.map(o=>+o.trip||0)):0;\n  const newTk=maxTrip+1;\n  document.getElementById(\'et-num\').value=newTk;\n  log(`New trip number: ${newTk} — add orders and assign trip #${newTk}`,\'linfo\');\n}\n\nfunction addCustToTrip(){\n  const tk=+document.getElementById(\'add-to-trip\').value;\n  const idx=+document.getElementById(\'add-cust-sel\').value;\n  if(!tk||isNaN(idx)){log(\'⚠ Select trip and customer\',\'lwarn\');return}\n  const o=orders[idx];\n  if(!o)return;\n  orders[idx].trip=tk;\n  if(optimizedTrips[tk]){\n    optimizedTrips[tk]=[...optimizedTrips[tk],{...o,stop:optimizedTrips[tk].length+1,legDist:\'—\',cumDist:\'—\',eta:\'—\'}];\n  }else{\n    optimizedTrips[tk]=buildRouteMeta(orders.filter(o=>o.trip==tk));\n  }\n  renderPreview();renderTripsOverview();renderRoutemapCards(optimizedTrips);\n  log(`Added ${o.customer} → Trip ${tk}`,\'lok\');\n}\n\nfunction renderTripsOverview(){\n  const tb=document.getElementById(\'tripsOverview\');\n  const keys=Object.keys(optimizedTrips).sort((a,b)=>+a-+b);\n  if(!keys.length){tb.innerHTML=\'<tr><td colspan="8"><div class="empty"><div class="eico">🚚</div>No trips yet</div></td></tr>\';return}\n  const ti={};keys.forEach((t,i)=>ti[t]=i);\n  tb.innerHTML=keys.map(tk=>{\n    const route=optimizedTrips[tk];\n    const lastR=route[route.length-1];\n    const km=lastR?lastR.cumDist:\'—\';\n    const vinfo=orders.find(o=>o.trip==tk);\n    return `<tr>\n      <td><span class="trip-tag ${TCLS[ti[tk]%12]}">${tk}</span></td>\n      <td>${vinfo?.shift||1}</td>\n      <td style="font-size:.7rem">${vinfo?.vehicle||\'—\'}</td>\n      <td style="font-family:\'Space Mono\',monospace;color:var(--cyan)">${route.length}</td>\n      <td style="font-family:\'Space Mono\',monospace">${route.reduce((s,o)=>s+o.crates,0).toFixed(1)}</td>\n      <td style="font-family:\'Space Mono\',monospace;color:var(--orange)">${route.reduce((s,o)=>s+o.tonnage,0).toFixed(1)}</td>\n      <td style="font-family:\'Space Mono\',monospace">${km} km</td>\n      <td>\n        <button class="btn bgh bsm" onclick="document.getElementById(\'et-num\').value=${tk};loadTripForEdit();showTab(\'trips\')">✏ Edit</button>\n        <button class="btn bc bsm" onclick="focusStop(${tk},0)">🗺</button>\n      </td>\n    </tr>`;\n  }).join(\'\');\n}\n\n// ══════════════════════════════════════════════════════\n// TRIP MODAL\n// ══════════════════════════════════════════════════════\nfunction openTripModal(tk){\n  const route=optimizedTrips[tk];\n  if(!route)return;\n  const vinfo=orders.find(o=>o.trip==tk);\n  document.getElementById(\'modalTitle\').textContent=`Edit Trip ${tk}`;\n  document.getElementById(\'modalBody\').innerHTML=`\n    <div style="margin-bottom:12px">\n      <div class="form-r">\n        <div class="fg"><label>Shift</label><input class="te-inp" id="mo-shift" value="${vinfo?.shift||1}"></div>\n        <div class="fg"><label>Vehicle</label><select class="te-inp" id="mo-veh">\n          ${vehicles.map(v=>`<option value="${v.name}" ${vinfo?.vehicle===v.name?\'selected\':\'\'}>${v.name}</option>`).join(\'\')}\n          <option value="${vinfo?.vehicle||\'\'}" selected>${vinfo?.vehicle||\'Custom\'}</option>\n        </select></div>\n      </div>\n      <div style="font-family:\'Space Mono\',monospace;font-size:.6rem;color:var(--muted);margin-bottom:6px">STOP SEQUENCE (drag to reorder)</div>\n    </div>\n    <div>\n      ${route.map((o,i)=>`<div style="display:flex;align-items:center;gap:8px;padding:6px 8px;background:var(--card2);border-radius:6px;margin-bottom:4px;font-size:.74rem">\n        <span style="font-family:\'Space Mono\',monospace;color:var(--cyan);font-size:.65rem;min-width:20px">${o.stop}</span>\n        <span style="flex:1;font-weight:600">${o.customer}</span>\n        <span style="color:var(--muted);font-size:.65rem">${o.window}</span>\n        <span style="font-family:\'Space Mono\',monospace;font-size:.65rem;color:var(--orange)">${o.tonnage}kg</span>\n        <span style="font-family:\'Space Mono\',monospace;font-size:.62rem;color:var(--green)">${o.eta}</span>\n      </div>`).join(\'\')}\n    </div>\n    <div style="margin-top:12px;display:flex;gap:8px">\n      <button class="btn bc" onclick="saveModalTrip(${tk})">💾 SAVE</button>\n      <button class="btn bgh" onclick="closeModal()">CANCEL</button>\n    </div>`;\n  document.getElementById(\'tripModal\').classList.add(\'open\');\n}\n\nfunction saveModalTrip(tk){\n  const shift=document.getElementById(\'mo-shift\')?.value;\n  const veh=document.getElementById(\'mo-veh\')?.value;\n  orders.filter(o=>o.trip==tk).forEach(o=>{if(shift)o.shift=+shift;if(veh)o.vehicle=veh});\n  renderRoutemapCards(optimizedTrips);renderTripsOverview();\n  closeModal();log(`✓ Trip ${tk} updated`,\'lok\');\n}\nfunction closeModal(){document.getElementById(\'tripModal\').classList.remove(\'open\')}\n\n// ══════════════════════════════════════════════════════\n// VEHICLES\n// ══════════════════════════════════════════════════════\nfunction addVehicle(){\n  const g=id=>document.getElementById(id)?.value?.trim()||\'\';\n  const gn=id=>+document.getElementById(id)?.value||0;\n  if(!g(\'v-name\')){log(\'⚠ Vehicle name required\',\'lwarn\');return}\n  const v={id:g(\'v-id\')||`VEH-${String(vehicles.length+1).padStart(2,\'0\')}`,\n    name:g(\'v-name\'),type:g(\'v-type\'),capacity:gn(\'v-cap\'),cc:g(\'v-cc\'),\n    fuel:g(\'v-fuel\'),fixedCost:gn(\'v-fc\'),kmCost:gn(\'v-kmc\'),notes:g(\'v-notes\')};\n  vehicles.push(v);\n  renderVehicles();populateEditDropdowns();\n  log(`✓ Vehicle added: ${v.name}`,\'lok\');\n  clearVehicleForm();\n}\n\nfunction clearVehicleForm(){[\'v-id\',\'v-name\',\'v-cc\',\'v-cap\',\'v-fc\',\'v-kmc\',\'v-notes\'].forEach(id=>{const el=document.getElementById(id);if(el)el.value=\'\'})}\n\nfunction removeVehicle(i){const n=vehicles[i].name;vehicles.splice(i,1);renderVehicles();populateEditDropdowns();log(\'✗ Removed: \'+n,\'lwarn\')}\n\nfunction renderVehicles(){\n  document.getElementById(\'vehCount\').textContent=vehicles.length;\n  const tb=document.getElementById(\'vehBody\');\n  if(!vehicles.length){tb.innerHTML=\'<tr><td colspan="9"><div class="empty"><div class="eico">🚛</div>No vehicles</div></td></tr>\';return}\n  tb.innerHTML=vehicles.map((v,i)=>`<tr>\n    <td style="font-family:\'Space Mono\',monospace;font-size:.65rem;color:var(--cyan)">${v.id}</td>\n    <td style="font-weight:600">${v.name}</td>\n    <td style="font-size:.7rem;color:var(--muted)">${v.type}</td>\n    <td style="font-family:\'Space Mono\',monospace">${v.capacity}</td>\n    <td style="font-family:\'Space Mono\',monospace;font-size:.65rem">${v.cc}</td>\n    <td style="font-size:.7rem;color:var(--muted)">${v.fuel}</td>\n    <td style="font-family:\'Space Mono\',monospace;color:var(--green)">₹${v.fixedCost}</td>\n    <td style="font-family:\'Space Mono\',monospace;color:var(--orange)">₹${v.kmCost}</td>\n    <td><button class="btn bred bsm" onclick="removeVehicle(${i})">✕</button></td>\n  </tr>`).join(\'\');\n}\n\nfunction loadDefaultVehicles(){\n  vehicles=[\n    {id:\'VEH-01\',name:\'Tata Ace\',type:\'Mini Truck\',capacity:750,cc:\'CC 58\',fuel:\'Diesel\',fixedCost:800,kmCost:8,notes:\'Standard mini\'},\n    {id:\'VEH-02\',name:\'Tata Ace HD\',type:\'Mini Truck\',capacity:1000,cc:\'CC 98\',fuel:\'Diesel\',fixedCost:900,kmCost:9,notes:\'Heavy duty\'},\n    {id:\'VEH-03\',name:\'Mahindra Jeeto\',type:\'Mini Truck\',capacity:600,cc:\'CC 48\',fuel:\'Diesel\',fixedCost:700,kmCost:7,notes:\'\'},\n    {id:\'VEH-04\',name:\'Ashok Leyland Dost\',type:\'Medium Truck\',capacity:1500,cc:\'CC 101\',fuel:\'Diesel\',fixedCost:1200,kmCost:11,notes:\'\'},\n    {id:\'VEH-05\',name:\'Bajaj RE\',type:\'Pickup\',capacity:350,cc:\'CC 24\',fuel:\'CNG\',fixedCost:450,kmCost:4,notes:\'3-wheeler\'},\n  ];\n  renderVehicles();populateEditDropdowns();\n  log(\'Default fleet loaded\',\'lok\');\n}\n\n// ══════════════════════════════════════════════════════\n// COST CALCULATION — base km + extra km model\n// ══════════════════════════════════════════════════════\nfunction gcp(id){return +document.getElementById(id)?.value||0}\n\nfunction getVehicleCosts(vehicleName){\n  const v=vehicles.find(v=>vehicleName&&vehicleName.includes(v.name));\n  if(v&&(v.fixedCost||v.kmCost)) return{fixed:v.fixedCost,kmRate:v.kmCost};\n  const fixedTotal=gcp(\'cp-driver\')+gcp(\'cp-vfixed\')+gcp(\'cp-load\')+gcp(\'cp-toll\')+gcp(\'cp-other\');\n  const kmRate=gcp(\'cp-fuel\')+gcp(\'cp-maint\')+gcp(\'cp-ovar\');\n  return{fixed:fixedTotal,kmRate};\n}\n\nfunction calcCosts(){\n  const keys=Object.keys(optimizedTrips).sort((a,b)=>+a-+b);\n  if(!keys.length) return;\n  const depot=getDepot();\n  const baseKm=gcp(\'cp-basekm\')||25;\n  costData=[];\n  keys.forEach(tk=>{\n    const route=optimizedTrips[tk];\n    if(!route.length)return;\n    const lastStop=route[route.length-1];\n    const rtrn=hav(lastStop,depot);\n    const km=+(+lastStop.cumDist+rtrn).toFixed(2);\n    const extraKm=Math.max(0,+(km-baseKm).toFixed(2));\n    const vinfo=orders.find(o=>o.trip==tk);\n    const {fixed,kmRate}=getVehicleCosts(vinfo?.vehicle||\'\');\n    const varCost=+(extraKm*kmRate).toFixed(2);\n    const total=+(fixed+varCost).toFixed(2);\n    const tons=route.reduce((s,o)=>s+(+o.tonnage||0),0);\n    costData.push({\n      trip:tk,vehicle:vinfo?.vehicle||\'—\',stops:route.length,\n      km,baseKm:Math.min(km,baseKm).toFixed(2),extraKm,\n      fixed,varCost,total,\n      perStop:+(total/route.length).toFixed(2),\n      perKg:tons>0?+(total/tons).toFixed(2):0\n    });\n  });\n  renderCostTable();\n  const totalCost=costData.reduce((s,c)=>s+c.total,0).toFixed(2);\n  document.getElementById(\'s-cost\').textContent=\'\\u20B9\'+totalCost;\n  document.getElementById(\'costRunLbl\').textContent=\'Calc: \'+new Date().toLocaleTimeString(\'en-US\',{hour12:false});\n  // Refresh routemap cards so km/cost bars update\n  if(Object.keys(optimizedTrips).length) renderRoutemapCards(optimizedTrips);\n  renderMapSidebar();\n}\n\nfunction renderCostTable(){\n  const tb=document.getElementById(\'costBody\');\n  if(!costData.length){tb.innerHTML=\'<tr><td colspan="11"><div class="empty"><div class="eico">\\u20B9</div>No route data</div></td></tr>\';return}\n  const ti={};costData.forEach((c,i)=>ti[c.trip]=i);\n  const baseKm=gcp(\'cp-basekm\')||25;\n  tb.innerHTML=costData.map(c=>`<tr>\n    <td><span class="trip-tag ${TCLS[ti[c.trip]%12]}">${c.trip}</span></td>\n    <td style="font-size:.68rem;max-width:95px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${c.vehicle}</td>\n    <td style="font-family:\'Space Mono\',monospace;text-align:center">${c.stops}</td>\n    <td style="font-family:\'Space Mono\',monospace">${c.km} km</td>\n    <td style="font-family:\'Space Mono\',monospace;color:var(--green)">${c.baseKm} km</td>\n    <td style="font-family:\'Space Mono\',monospace;color:${c.extraKm>0?\'var(--orange)\':\'var(--muted)\'}">${c.extraKm>0?\'+\'+c.extraKm+\' km\':\'—\'}</td>\n    <td style="font-family:\'Space Mono\',monospace;color:var(--yellow)">\\u20B9${c.fixed}</td>\n    <td style="font-family:\'Space Mono\',monospace;color:var(--orange)">\\u20B9${c.varCost}<span style="font-size:.52rem;color:var(--muted);display:block">${c.extraKm>0?\'on \'+c.extraKm+\'km\':\'no extra\'}</span></td>\n    <td style="font-family:\'Space Mono\',monospace;color:var(--cyan);font-weight:700">\\u20B9${c.total}</td>\n    <td style="font-family:\'Space Mono\',monospace;font-size:.62rem;color:var(--muted)">\\u20B9${c.perStop}</td>\n    <td style="font-family:\'Space Mono\',monospace;font-size:.62rem;color:var(--muted)">\\u20B9${c.perKg}</td>\n  </tr>`).join(\'\')+`\n  <tr class="total-row">\n    <td colspan="3" style="font-family:\'Space Mono\',monospace">TOTAL (${costData.length} trips)</td>\n    <td style="font-family:\'Space Mono\',monospace">${costData.reduce((s,c)=>s+c.km,0).toFixed(1)} km</td>\n    <td style="font-family:\'Space Mono\',monospace;color:var(--green)">${costData.reduce((s,c)=>s+(+c.baseKm||0),0).toFixed(1)} km</td>\n    <td style="font-family:\'Space Mono\',monospace;color:var(--orange)">${costData.reduce((s,c)=>s+c.extraKm,0).toFixed(1)} km extra</td>\n    <td style="font-family:\'Space Mono\',monospace">\\u20B9${costData.reduce((s,c)=>s+c.fixed,0).toFixed(2)}</td>\n    <td style="font-family:\'Space Mono\',monospace">\\u20B9${costData.reduce((s,c)=>s+c.varCost,0).toFixed(2)}</td>\n    <td style="font-family:\'Space Mono\',monospace;color:var(--cyan)">\\u20B9${costData.reduce((s,c)=>s+c.total,0).toFixed(2)}</td>\n    <td style="font-family:\'Space Mono\',monospace">\\u20B9${(costData.reduce((s,c)=>s+c.total,0)/Math.max(1,costData.reduce((s,c)=>s+c.stops,0))).toFixed(2)}</td>\n    <td>—</td>\n  </tr>`;\n  const sb=document.getElementById(\'costSummaryBar\');\n  sb.style.display=\'block\';\n  document.getElementById(\'costSumGrid\').innerHTML=`\n    <div class="sc" style="background:var(--bg)"><div class="sl">Total Cost</div><div class="sv" style="color:var(--cyan)">\\u20B9${costData.reduce((s,c)=>s+c.total,0).toFixed(2)}</div><div class="ss">all trips</div></div>\n    <div class="sc g" style="background:var(--bg)"><div class="sl">Base KM Limit</div><div class="sv">${baseKm} km</div><div class="ss">no var charge below</div></div>\n    <div class="sc y" style="background:var(--bg)"><div class="sl">Fixed Costs</div><div class="sv">\\u20B9${costData.reduce((s,c)=>s+c.fixed,0).toFixed(2)}</div><div class="ss">all trips</div></div>\n    <div class="sc o" style="background:var(--bg)"><div class="sl">Variable Costs</div><div class="sv">\\u20B9${costData.reduce((s,c)=>s+c.varCost,0).toFixed(2)}</div><div class="ss">${costData.reduce((s,c)=>s+c.extraKm,0).toFixed(1)} extra km</div></div>`;\n}\n\n// ══════════════════════════════════════════════════════\n// EXCEL EXPORT\n// ══════════════════════════════════════════════════════\nfunction exportXLSX(){\n  if(!Object.keys(optimizedTrips).length){log(\'⚠ Run CRON first\',\'lwarn\');return}\n  const wb=XLSX.utils.book_new();\n  const depot=getDepot();\n\n  // Summary\n  const sumData=[\n    [\'LML ROUTE ESTIMATOR — EXPORT\'],[\'Generated:\',new Date().toLocaleString()],\n    [\'Algorithm:\',\'Nearest Neighbor TSP + Priority Weighting\'],[],\n    [\'Orders\',orders.length],[\'Trips\',Object.keys(optimizedTrips).length],\n    [\'Total Crates\',orders.reduce((s,o)=>s+o.crates,0)],\n    [\'Total Tonnage (kg)\',orders.reduce((s,o)=>s+o.tonnage,0).toFixed(1)],\n    [\'Est. Total Cost (₹)\',costData.reduce((s,c)=>s+c.total,0).toFixed(2)],[],\n    [\'COST SUMMARY\'],[\'Trip\',\'Vehicle\',\'Stops\',\'Total km\',\'Base km\',\'Extra km\',\'Fixed ₹\',\'Var ₹\',\'Total ₹\',\'₹/Stop\',\'₹/kg\'],\n    ...costData.map(c=>[c.trip,c.vehicle,c.stops,c.km,c.baseKm,c.extraKm,c.fixed,c.varCost,c.total,c.perStop,c.perKg])\n  ];\n  const sumWs=XLSX.utils.aoa_to_sheet(sumData);\n  sumWs[\'!cols\']=[{wch:20},{wch:25}];\n  XLSX.utils.book_append_sheet(wb,sumWs,\'Summary\');\n\n  // Full route\n  const rh=[\'Stop\',\'Trip\',\'Shift\',\'Vehicle\',\'Order ID\',\'Customer\',\'Address\',\'Lat\',\'Lng\',\'Crates\',\'Tonnage\',\'Window\',\'Priority\',\'Leg km\',\'Cum km\',\'ETA\'];\n  const rb=[rh];\n  const keys=Object.keys(optimizedTrips).sort((a,b)=>+a-+b);\n  keys.forEach(tk=>{\n    optimizedTrips[tk].forEach(r=>rb.push([r.stop,r.trip,r.shift||1,r.vehicle||\'\',r.id,r.customer,r.address,r.lat,r.lng,r.crates,r.tonnage,r.window,r.priority,+r.legDist,+r.cumDist,r.eta]));\n    const t=optimizedTrips[tk];\n    rb.push([\'\',`TRIP ${tk} TOTAL`,\'\',\'\',`${t.length} stops`,\'\',\'\',\'\',\'\',t.reduce((s,o)=>s+o.crates,0),t.reduce((s,o)=>s+o.tonnage,0).toFixed(1),\'\',\'\',\'\',t[t.length-1]?.cumDist||\'\',\'\']);\n    rb.push([]);\n  });\n  const rws=XLSX.utils.aoa_to_sheet(rb);\n  rws[\'!cols\']=[{wch:6},{wch:6},{wch:6},{wch:16},{wch:10},{wch:28},{wch:30},{wch:10},{wch:10},{wch:7},{wch:11},{wch:12},{wch:8},{wch:9},{wch:9},{wch:8}];\n  XLSX.utils.book_append_sheet(wb,rws,\'Full Route\');\n\n  // Cost sheet\n  const ch=[\'Trip\',\'Vehicle\',\'Stops\',\'Distance (km)\',\'Fixed Cost (₹)\',\'Variable Cost (₹)\',\'Total Cost (₹)\',\'Cost/Stop (₹)\',\'Cost/kg (₹)\'];\n  const cd=[ch,...costData.map(c=>[c.trip,c.vehicle,c.stops,c.km,c.fixed,c.varCost,c.total,c.perStop,c.perKg]),\n    [\'TOTAL\',\'\',costData.reduce((s,c)=>s+c.stops,0),costData.reduce((s,c)=>s+c.km,0).toFixed(1),\n     costData.reduce((s,c)=>s+c.fixed,0).toFixed(2),costData.reduce((s,c)=>s+c.varCost,0).toFixed(2),\n     costData.reduce((s,c)=>s+c.total,0).toFixed(2),\'\',\'\']\n  ];\n  const cws=XLSX.utils.aoa_to_sheet(cd);\n  XLSX.utils.book_append_sheet(wb,cws,\'Cost Analysis\');\n\n  // Per-trip sheets\n  keys.forEach(tk=>{\n    const sh=[[`TRIP ${tk}`],[`Vehicle: ${orders.find(o=>o.trip==tk)?.vehicle||\'—\'}`],[`Tonnage: ${optimizedTrips[tk].reduce((s,o)=>s+o.tonnage,0).toFixed(1)} kg`],[],\n      [\'Stop\',\'Order ID\',\'Customer\',\'Address\',\'Crates\',\'Tonnage\',\'Window\',\'Leg km\',\'Cum km\',\'ETA\'],\n      ...optimizedTrips[tk].map(r=>[r.stop,r.id,r.customer,r.address,r.crates,r.tonnage,r.window,+r.legDist,+r.cumDist,r.eta])\n    ];\n    const ws=XLSX.utils.aoa_to_sheet(sh);\n    ws[\'!cols\']=[{wch:6},{wch:10},{wch:28},{wch:30},{wch:7},{wch:10},{wch:12},{wch:8},{wch:8},{wch:8}];\n    XLSX.utils.book_append_sheet(wb,ws,`Trip ${tk}`);\n  });\n\n  const fname=`LML_Route_${new Date().toISOString().slice(0,10)}.xlsx`;\n  XLSX.writeFile(wb,fname);\n  log(`⬇ Exported: ${fname}`,\'lok\');\n}\n\n// ══════════════════════════════════════════════════════\n// CRON\n// ══════════════════════════════════════════════════════\nfunction setCI(sec,btn){\n  cronSec=sec;\n  document.querySelectorAll(\'.ibtn\').forEach(b=>b.classList.remove(\'active\'));\n  btn.classList.add(\'active\');\n  restartCron();log(`CRON → ${sec}s`,\'linfo\');\n}\n\nfunction resetProg(){\n  clearInterval(progTimer);cronProg=0;\n  const fill=document.getElementById(\'progFill\'),pct=document.getElementById(\'progPct\');\n  fill.style.width=\'0%\';pct.textContent=\'0%\';\n  const step=100/(cronSec*10);\n  progTimer=setInterval(()=>{cronProg=Math.min(100,cronProg+step);fill.style.width=cronProg+\'%\';pct.textContent=Math.round(cronProg)+\'%\';},100);\n}\n\nfunction restartCron(){if(cronTimer)clearInterval(cronTimer);resetProg();cronTimer=setInterval(runCron,cronSec*1000)}\n\nsetInterval(()=>{const r=Math.round((1-cronProg/100)*cronSec);document.getElementById(\'nextRun\').textContent=\'T-\'+r+\'s\';},500);\n\n// ══════════════════════════════════════════════════════\n// LOG\n// ══════════════════════════════════════════════════════\nfunction log(msg,cls=\'\'){\n  const el=document.getElementById(\'cronLog\');\n  const d=document.createElement(\'div\');\n  d.className=\'le \'+cls;\n  d.innerHTML=`<span class="lt2">[${new Date().toLocaleTimeString(\'en-US\',{hour12:false})}]</span><span class="lm">${msg}</span>`;\n  el.prepend(d);\n  while(el.children.length>100)el.removeChild(el.lastChild);\n}\nfunction clearLog(){document.getElementById(\'cronLog\').innerHTML=\'\';log(\'Log cleared\',\'linfo\')}\n\n// ══════════════════════════════════════════════════════\n// DEMO DATA\n// ══════════════════════════════════════════════════════\nfunction loadDemo(){\n  orders=[\n    {id:\'ORD-001\',trip:2,shift:1,vehicle:\'Tata Ace CC 58.1\',customer:\'SLV GENERAL STORES\',address:\'KR Market, Bengaluru\',lat:12.985,lng:77.597,crates:19,tonnage:267,window:\'05:30-06:30\',priority:\'high\'},\n    {id:\'ORD-002\',trip:2,shift:1,vehicle:\'Tata Ace CC 58.1\',customer:\'Kiran vegetables\',address:\'Shivajinagar, Bengaluru\',lat:12.964,lng:77.578,crates:19.1,tonnage:255,window:\'07:00-08:00\',priority:\'med\'},\n    {id:\'ORD-003\',trip:2,shift:1,vehicle:\'Tata Ace CC 58.1\',customer:"GROCER\'S CHOICE",address:\'Koramangala 4th Block\',lat:12.936,lng:77.624,crates:5,tonnage:110,window:\'07:30-08:30\',priority:\'low\'},\n    {id:\'ORD-004\',trip:2,shift:1,vehicle:\'Tata Ace CC 58.1\',customer:\'The Fresh Harvest Hub\',address:\'HSR Layout, Bengaluru\',lat:12.912,lng:77.647,crates:15,tonnage:330,window:\'08:30-09:30\',priority:\'med\'},\n    {id:\'ORD-005\',trip:3,shift:1,vehicle:\'Tata Ace CC 98.1\',customer:\'haritha vegetables\',address:\'Jayanagar 4th Block\',lat:12.929,lng:77.583,crates:50.6,tonnage:529,window:\'05:00-06:00\',priority:\'high\'},\n    {id:\'ORD-006\',trip:3,shift:1,vehicle:\'Tata Ace CC 98.1\',customer:\'Aditya mart\',address:\'BTM 2nd Stage, Bengaluru\',lat:12.916,lng:77.610,crates:11.3,tonnage:144.5,window:\'07:00-08:00\',priority:\'med\'},\n    {id:\'ORD-007\',trip:3,shift:1,vehicle:\'Tata Ace CC 98.1\',customer:\'Raith Bazar\',address:\'Electronic City Phase 1\',lat:12.845,lng:77.660,crates:21.1,tonnage:252,window:\'07:00-08:00\',priority:\'med\'},\n    {id:\'ORD-008\',trip:3,shift:1,vehicle:\'Tata Ace CC 98.1\',customer:\'Mohammed asif ulla\',address:\'Silk Board, Bengaluru\',lat:12.917,lng:77.622,crates:15.1,tonnage:262.5,window:\'09:00-10:00\',priority:\'low\'},\n    {id:\'ORD-009\',trip:4,shift:1,vehicle:\'Tata Ace CC 97.1\',customer:\'S V T vegetables\',address:\'Marathahalli Bridge\',lat:12.959,lng:77.697,crates:34.9,tonnage:431,window:\'05:00-06:00\',priority:\'high\'},\n    {id:\'ORD-010\',trip:4,shift:1,vehicle:\'Tata Ace CC 97.1\',customer:\'Hi cups\',address:\'Whitefield Main Rd\',lat:12.969,lng:77.749,crates:17.3,tonnage:250.5,window:\'07:00-08:00\',priority:\'med\'},\n    {id:\'ORD-011\',trip:4,shift:1,vehicle:\'Tata Ace CC 97.1\',customer:\'Nammoora thota\',address:\'Indiranagar 100ft Rd\',lat:12.978,lng:77.640,crates:23.7,tonnage:388.5,window:\'09:00-10:00\',priority:\'low\'},\n    {id:\'ORD-012\',trip:8,shift:1,vehicle:\'Tata Ace CC 79.0\',customer:\'Anjan Vegitables\',address:\'Malleshwaram Circle\',lat:13.003,lng:77.566,crates:30,tonnage:501,window:\'05:00-06:00\',priority:\'high\'},\n    {id:\'ORD-013\',trip:8,shift:1,vehicle:\'Tata Ace CC 79.0\',customer:\'Ganesh fruits and vegetables\',address:\'Rajajinagar 1st Block\',lat:12.992,lng:77.552,crates:17,tonnage:351,window:\'06:00-07:00\',priority:\'med\'},\n    {id:\'ORD-014\',trip:8,shift:1,vehicle:\'Tata Ace CC 79.0\',customer:\'Bangalore fruits\',address:\'Nagarbhavi Cross\',lat:12.951,lng:77.501,crates:17,tonnage:274.2,window:\'07:00-08:00\',priority:\'low\'},\n    {id:\'ORD-015\',trip:9,shift:1,vehicle:\'Tata Ace CC 46.9\',customer:\'Farm Fresh\',address:\'JP Nagar 6th Phase\',lat:12.893,lng:77.597,crates:16.8,tonnage:260.5,window:\'06:00-07:00\',priority:\'high\'},\n    {id:\'ORD-016\',trip:9,shift:1,vehicle:\'Tata Ace CC 46.9\',customer:\'Green Fresh\',address:\'Banashankari 2nd Stage\',lat:12.924,lng:77.546,crates:18.1,tonnage:299,window:\'06:00-07:00\',priority:\'med\'},\n    {id:\'ORD-017\',trip:17,shift:1,vehicle:\'Tata Ace CC 23.9\',customer:\'Mayur fresh\',address:\'Hebbal Flyover\',lat:13.035,lng:77.597,crates:23.9,tonnage:260.5,window:\'06:00-07:00\',priority:\'high\'},\n    {id:\'ORD-018\',trip:17,shift:1,vehicle:\'Tata Ace CC 23.9\',customer:\'Surjodhaya vegetables\',address:\'Yelahanka New Town\',lat:13.100,lng:77.596,crates:18.8,tonnage:266,window:\'06:30-07:30\',priority:\'med\'},\n    {id:\'ORD-019\',trip:40,shift:1,vehicle:\'Tata Ace CC 80.3\',customer:\'hulimavu Halli Thota\',address:\'Hulimavu Lake Rd\',lat:12.876,lng:77.620,crates:21.1,tonnage:265,window:\'05:00-06:00\',priority:\'high\'},\n    {id:\'ORD-020\',trip:40,shift:1,vehicle:\'Tata Ace CC 80.3\',customer:\'SRI MALAIMAHADESH\',address:\'Uttarahalli Main Rd\',lat:12.893,lng:77.544,crates:13,tonnage:266,window:\'06:00-07:00\',priority:\'med\'},\n    {id:\'ORD-021\',trip:42,shift:1,vehicle:\'Tata Ace CC 90.8\',customer:\'Lalitha Enterprises\',address:\'Vijayanagar Main\',lat:12.971,lng:77.534,crates:15.7,tonnage:258.5,window:\'05:30-06:30\',priority:\'high\'},\n    {id:\'ORD-022\',trip:42,shift:1,vehicle:\'Tata Ace CC 90.8\',customer:\'Green Garden fruit\',address:\'Chord Road, Bengaluru\',lat:12.982,lng:77.525,crates:15.5,tonnage:254.5,window:\'06:00-07:00\',priority:\'med\'},\n    {id:\'ORD-023\',trip:45,shift:1,vehicle:\'Tata Ace CC 88.7\',customer:\'Etharth vegetables\',address:\'Banaswadi Main Rd\',lat:13.021,lng:77.642,crates:32.8,tonnage:511,window:\'06:00-07:00\',priority:\'high\'},\n    {id:\'ORD-024\',trip:45,shift:1,vehicle:\'Tata Ace CC 88.7\',customer:\'Kabbalamma jigani\',address:\'Jigani Main Rd\',lat:12.800,lng:77.635,crates:27.3,tonnage:439,window:\'07:00-08:00\',priority:\'med\'},\n  ];\n  oc=orders.length+1;\n  loadDefaultVehicles();\n  renderPreview();updateStats();\n  log(`Demo: ${orders.length} orders, 10 trips loaded`,\'lok\');\n  runCron();showTab(\'routemap\');\n}\n\n// ══════════════════════════════════════════════════════\n// INIT\n// ══════════════════════════════════════════════════════\nwindow.addEventListener(\'load\',()=>{\n  loadDefaultVehicles();\n  restartCron();\n  log(\'✓ System ready — load demo or upload your Excel file\',\'lok\');\n});\ndocument.getElementById(\'tripModal\').addEventListener(\'click\',e=>{if(e.target===document.getElementById(\'tripModal\'))closeModal()});\n\n\n// ══════════════════════════════════════════════════════\n// STREAMLIT DATA BRIDGE\n// Called by Streamlit to inject orders + drivers\n// ══════════════════════════════════════════════════════\nlet _stDrivers = [];   // drivers loaded from Streamlit\nlet _tripDriverMap = {};  // { tripKey: {uid, name} }\n\nfunction receiveFromStreamlit(jsonStr) {\n  try {\n    const d = JSON.parse(jsonStr);\n    if (d.orders && d.orders.length) {\n      loadOrdersFromStreamlit(d.orders);\n    }\n    if (d.drivers) {\n      _stDrivers = d.drivers;\n      slog(`✓ ${d.drivers.length} drivers loaded`, \'lok\');\n    }\n    if (d.date) {\n      slog(`📅 Date: ${d.date}`, \'linfo\');\n    }\n  } catch(e) { slog(\'Error parsing data: \'+e.message,\'lerr\'); }\n}\n\nfunction loadOrdersFromStreamlit(raw) {\n  orders = raw.map((r, i) => ({\n    id:       r.id       || r[\'Order ID\']          || `ORD-${String(i+1).padStart(3,\'0\')}`,\n    trip:     parseInt(r.trip || r[\'Tripid\'] || 1) || 1,\n    shift:    parseInt(r.shift || 1) || 1,\n    vehicle:  r.vehicle  || r[\'Vehicle\']           || \'Tata Ace\',\n    customer: r.customer || r[\'Customer shop name\']|| r[\'Customer\'] || `Customer ${i+1}`,\n    address:  r.address  || r[\'Shop Location\']     || \'—\',\n    lat:      parseFloat(r.lat  || r[\'Latitude\']   || 0) || 0,\n    lng:      parseFloat(r.lng  || r[\'Longitude\']  || 0) || 0,\n    crates:   parseFloat(r.crates  || r[\'OrderedQty\']  || r[\'TotalCrates\'] || 0) || 0,\n    tonnage:  parseFloat(r.tonnage || r[\'OrderKg\']     || r[\'OrderTotal\']  || 0) || 0,\n    window:   r.window   || r[\'Slot\']              || r[\'DeliveryCutOff\'] || \'07:00-08:00\',\n    priority: r.priority || \'med\',\n    custId:   r.custId   || r[\'CustomerId\']        || \'\',\n  }));\n  // Re-number trips 1,2,3… from whatever they were\n  const rawKeys = [...new Set(orders.map(o=>o.trip))].sort((a,b)=>a-b);\n  const remap   = {}; rawKeys.forEach((k,i)=>remap[k]=i+1);\n  orders.forEach(o=>{ o.trip = remap[o.trip]||o.trip; });\n  oc = orders.length+1;\n  renderPreview(); updateStats();\n  slog(`✓ ${orders.length} orders loaded across ${rawKeys.length} trip group(s)`, \'lok\');\n  runCron();\n  showTab(\'routemap\');\n}\n\n// ══════════════════════════════════════════════════════\n// SUBMIT TAB RENDERING\n// ══════════════════════════════════════════════════════\nfunction renderSubmitTab() {\n  const keys = Object.keys(optimizedTrips).sort((a,b)=>+a-+b);\n  if (!keys.length) return;\n\n  document.getElementById(\'submitStatusLbl\').textContent = `${keys.length} trips ready`;\n\n  // Summary cards\n  const cards = document.getElementById(\'submitCards\');\n  cards.innerHTML = keys.map((tk,ci) => {\n    const route  = optimizedTrips[tk];\n    const km     = route.length ? route[route.length-1].cumDist : 0;\n    const ton    = route.reduce((s,o)=>s+(+o.tonnage||0),0).toFixed(1);\n    const crates = route.reduce((s,o)=>s+(+o.crates||0),0).toFixed(0);\n    const vinfo  = orders.find(o=>o.trip==tk);\n    const vname  = (vinfo?.vehicle||\'Tata Ace\').replace(/CC\\s*[\\d.]+/i,\'\').trim();\n    const cd     = costData.find(c=>String(c.trip)===String(tk))||{};\n    const cost   = cd.total!=null?`₹${cd.total}`:\'—\';\n    return `<div style="background:var(--card2);border:1.5px solid var(--border);border-radius:11px;padding:13px;border-top:3px solid ${TCHEX[ci%12]}">\n      <div style="font-family:\'Space Mono\',monospace;font-size:.7rem;font-weight:700;color:${TCHEX[ci%12]};margin-bottom:7px">TRIP ${tk}</div>\n      <div style="font-size:.76rem;font-weight:600;margin-bottom:5px">${vname}</div>\n      <div style="font-family:\'Space Mono\',monospace;font-size:.6rem;color:var(--muted);line-height:2">\n        <span style="color:var(--text)">${route.length}</span> stops &nbsp;·&nbsp;\n        <span style="color:var(--cyan)">${km} km</span><br>\n        <span style="color:var(--orange)">${ton} kg</span> &nbsp;·&nbsp;\n        <span style="color:var(--yellow)">${crates} crates</span><br>\n        <span style="color:var(--green);font-weight:700">${cost}</span>\n      </div>\n    </div>`;\n  }).join(\'\');\n\n  // Driver assign table\n  const dat = document.getElementById(\'driverAssignTable\');\n  if (_stDrivers.length) {\n    const drvOpts = _stDrivers.map(d =>\n      `<option value="${d.uid}">${d.status===\'active\'?\'🟢\':\'⚫\'} ${d.name} | ${d.uid} | ${d.vehicle||\'\'}</option>`\n    ).join(\'\');\n    dat.innerHTML = `<div style="display:grid;grid-template-columns:120px 1fr 260px;gap:10px;font-family:\'Space Mono\',monospace;font-size:.55rem;color:var(--muted);text-transform:uppercase;padding:6px 10px;background:rgba(0,0,0,.2);border-radius:6px;margin-bottom:6px">\n      <div>Trip</div><div>Route</div><div>Driver</div></div>` +\n      keys.map((tk,ci) => {\n        const route = optimizedTrips[tk];\n        const km    = route.length ? route[route.length-1].cumDist : 0;\n        const cur   = _tripDriverMap[tk];\n        return `<div style="display:grid;grid-template-columns:120px 1fr 260px;gap:10px;align-items:center;padding:8px 10px;border-bottom:1px solid rgba(26,46,66,.5);font-size:.76rem">\n          <div>\n            <span class="trip-tag ${TCLS[ci%12]}">Trip ${tk}</span>\n            <div style="font-family:\'Space Mono\',monospace;font-size:.54rem;color:var(--muted);margin-top:3px">${km} km · ${route.length} stops</div>\n          </div>\n          <div style="font-size:.68rem;color:var(--muted)">\n            ${route.slice(0,3).map(o=>`<div style="overflow:hidden;text-overflow:ellipsis;white-space:nowrap;max-width:180px">${o.stop}. ${o.customer}</div>`).join(\'\')}\n            ${route.length>3?`<div style="color:var(--dim);font-size:.6rem">+${route.length-3} more</div>`:\'\'}\n          </div>\n          <select id="drv-${tk}" onchange="assignDriver(\'${tk}\',this)"\n            style="background:var(--card2);border:1px solid var(--border);color:var(--text);padding:6px 8px;border-radius:7px;font-size:.72rem;outline:none;width:100%">\n            <option value="">⬜ Assign later</option>\n            ${drvOpts}\n          </select>\n        </div>`;\n      }).join(\'\');\n    document.getElementById(\'driverAssignSection\').style.display = \'block\';\n  } else {\n    dat.innerHTML = `<div style="font-family:\'Space Mono\',monospace;font-size:.65rem;color:var(--muted);padding:10px">No drivers loaded — load orders from Streamlit first.</div>`;\n    document.getElementById(\'driverAssignSection\').style.display = \'block\';\n  }\n\n  document.getElementById(\'submitBtnWrap\').style.display = \'block\';\n}\n\nfunction assignDriver(tk, sel) {\n  if (!sel.value) { delete _tripDriverMap[tk]; return; }\n  const parts = sel.options[sel.selectedIndex].text.replace(/[🟢⚫]/g,\'\').trim().split(\'|\');\n  _tripDriverMap[tk] = { uid: sel.value.trim(), name: (parts[0]||\'\').trim() };\n  slog(`Trip ${tk} → Driver: ${_tripDriverMap[tk].name}`, \'lok\');\n}\n\nfunction slog(msg, cls=\'\') {\n  const el = document.getElementById(\'submitLog\');\n  if (!el) return;\n  const d = document.createElement(\'div\');\n  d.className = \'le \'+(cls||\'\');\n  d.innerHTML = `<span class="lt2">[${new Date().toLocaleTimeString(\'en-US\',{hour12:false})}]</span><span class="lm">${msg}</span>`;\n  el.prepend(d);\n  while(el.children.length > 60) el.removeChild(el.lastChild);\n}\n\n// Override showTab to auto-render submit tab\nconst _origShowTab = showTab;\nfunction showTab(id) {\n  _origShowTab(id);\n  if (id === \'submit\' && Object.keys(optimizedTrips).length) renderSubmitTab();\n}\n\n// ══════════════════════════════════════════════════════\n// SUBMIT to Streamlit via postMessage\n// ══════════════════════════════════════════════════════\nfunction submitTripsToSystem() {\n  const keys = Object.keys(optimizedTrips).sort((a,b)=>+a-+b);\n  if (!keys.length) { slog(\'⚠ No trips to submit — run CRON first\',\'lwarn\'); return; }\n  const btn = document.getElementById(\'submitBtn\');\n  btn.disabled = true;\n  document.getElementById(\'submitProg\').style.display = \'block\';\n  let prog = 0;\n  const tick = setInterval(()=>{ prog=Math.min(90,prog+12); document.getElementById(\'submitFill\').style.width=prog+\'%\'; }, 100);\n  const depot = getDepot();\n  const payload = keys.map(tk => {\n    const route = optimizedTrips[tk];\n    const vinfo = orders.find(o=>o.trip==tk)||{};\n    const last  = route[route.length-1]||{};\n    const retKm = last.lat&&last.lng ? hav(last,depot) : 0;\n    const cd    = costData.find(c=>String(c.trip)===String(tk))||{};\n    return {\n      tripKey:    String(tk),\n      vehicle:    vinfo.vehicle||\'\',\n      shift:      vinfo.shift||1,\n      stops:      route.map(o=>({\n        custId:   o.custId||\'\', orderId: o.id,\n        customer: o.customer,  address: o.address,\n        lat:      o.lat,       lng:     o.lng,\n        stop:     o.stop,      crates:  o.crates,\n        tonnage:  o.tonnage,   window:  o.window,\n        eta:      o.eta,       legKm:   o.legDist,\n        cumKm:    o.cumDist,   priority:o.priority||\'med\',\n      })),\n      driverUid:   (_tripDriverMap[tk]||{}).uid  || \'\',\n      driverName:  (_tripDriverMap[tk]||{}).name || \'\',\n      totalKm:     (+last.cumDist||0 + retKm).toFixed(2),\n      totalStops:  route.length,\n      totalCrates: route.reduce((s,o)=>s+(+o.crates||0),0),\n      totalTonnage:route.reduce((s,o)=>s+(+o.tonnage||0),0).toFixed(1),\n      estimatedCost: cd.total||0,\n    };\n  });\n  setTimeout(()=>{\n    clearInterval(tick);\n    document.getElementById(\'submitFill\').style.width = \'100%\';\n    document.getElementById(\'submitMsg\').textContent  = `✅ ${payload.length} trip(s) ready — confirm in Streamlit below`;\n    slog(`✅ Payload ready: ${payload.length} trip(s)`, \'lok\');\n    payload.forEach(t => slog(`  Trip ${t.tripKey}: ${t.totalStops} stops · ${t.totalKm}km · ${t.driverName||\'unassigned\'}`, \'linfo\'));\n    // Send to Streamlit parent\n    window.parent.postMessage({ type: \'GARLIC_TRIPS_SUBMITTED\', trips: payload }, \'*\');\n    // Also set on window for polling\n    window._submittedTrips = payload;\n    window._submittedAt    = Date.now();\n    btn.disabled = false;\n    btn.textContent = \'✅  SUBMITTED — Click again to resubmit\';\n    btn.style.background = \'linear-gradient(135deg,#1a7f4b,#27ae60)\';\n  }, 1200);\n}\n\n// ══════════════════════════════════════════════════════\n// LISTEN for incoming data from Streamlit\n// ══════════════════════════════════════════════════════\nwindow.addEventListener(\'message\', e => {\n  if (e.data && e.data.type === \'GARLIC_LOAD_DATA\') {\n    receiveFromStreamlit(JSON.stringify(e.data));\n  }\n});\n\n</script>\n</body>\n</html>\n'

def _safe_float(v):
    try:
        f = float(str(v).replace("\u20b9","").replace(",","").strip())
        return f if f == f else 0.0
    except Exception:
        return 0.0

_CITY_KW = {
    "Bengaluru": ["bengaluru","bangalore","blr","koramangala","indiranagar","whitefield","hebbal","jayanagar","btm","hsr","malleshwaram","rajajinagar","nagarbhavi"],
    "Mysuru":    ["mysuru","mysore","sayyaji"],
    "Hubli":     ["hubli","dharwad"],
    "Mangaluru": ["mangaluru","mangalore"],
    "Hassan":    ["hassan"],
    "Tumkur":    ["tumkur","tumakuru"],
}

def _infer_city_rp(stops):
    text = " ".join((o.get("address","")+" "+o.get("customer","")).lower() for o in stops)
    best, bs = "Bengaluru", 0
    for city, kws in _CITY_KW.items():
        s = sum(1 for kw in kws if kw in text)
        if s > bs: best, bs = city, s
    return best

def _prep_orders_rp(df_orders, sel_date):
    if df_orders is None or df_orders.empty: return []
    if "ORDER DATE" in df_orders.columns:
        df = df_orders[df_orders["ORDER DATE"].astype(str) == sel_date].copy()
    else:
        df = df_orders.copy()
    out = []
    for _, r in df.iterrows():
        trip_raw = str(r.get("Tripid","") or "").strip()
        if not trip_raw or trip_raw.lower() in ("nan","none",""):
            trip_raw = str(r.get("CustomerId","1")).strip()[-4:]
        try:    trip_num = int(trip_raw)
        except: trip_num = (abs(hash(trip_raw)) % 900) + 1
        out.append({
            "id":       str(r.get("Order ID", r.get("SaleOrderId",""))),
            "customer": str(r.get("Customer shop name", r.get("Customer",""))),
            "address":  str(r.get("Shop Location", r.get("address","\u2014"))),
            "lat":      _safe_float(r.get("Latitude",  0)),
            "lng":      _safe_float(r.get("Longitude", 0)),
            "crates":   _safe_float(r.get("OrderedQty", r.get("TotalCrates",0))),
            "tonnage":  _safe_float(r.get("OrderTotal", r.get("OrderKg",0))),
            "window":   str(r.get("DeliveryCutOff", r.get("Slot","07:00-08:00"))),
            "custId":   str(r.get("CustomerId","")),
            "trip":     trip_num, "priority": "med",
            "vehicle":  str(r.get("Driver", "Tata Ace")), "shift": 1,
        })
    return out

def _prep_drivers_rp(df_drivers):
    if df_drivers is None or df_drivers.empty: return []
    return [{
        "uid":    str(r.get("Driver ID","")),
        "name":   str(r.get("Full Name","")),
        "vehicle":str(r.get("Vehicle Type","")),
        "vnum":   str(r.get("Vehicle Number","")),
        "status": str(r.get("Active Status","Offline")).lower(),
    } for _, r in df_drivers.iterrows()]


def _page_route_planner_inline(user):
    import uuid as _uuid2
    from datetime import date as _d2, datetime as _dt2

    def _sl2(label, color=""):
        cls = f"sl sl-{color}" if color else "sl"
        return f'<div class="{cls}">{label}</div>'

    st.markdown(_sl2("\U0001f6e3\ufe0f Route Planner \u2014 Optimize & Dispatch"), unsafe_allow_html=True)

    dc1, dc2, dc3 = st.columns([2,2,4])
    with dc1:
        sel_date = st.date_input("Delivery date", value=_d2.today(), key="rp_date_filter")
    with dc2:
        st.write(""); st.write("")
        load_btn = st.button("\U0001f4e5 Load Orders", type="primary", key="rp_load_btn", use_container_width=True)
    with dc3:
        st.write(""); st.write("")
        st.caption("Loads from **Base** sheet filtered by ORDER DATE, then injects into the optimizer below.")

    sel_date_str = str(sel_date)
    if load_btn:
        with st.spinner(f"Loading orders for {sel_date_str}\u2026"):
            df_o = read_sheet("base")
            df_d = all_drivers()
            orders_js  = _prep_orders_rp(df_o, sel_date_str)
            drivers_js = _prep_drivers_rp(df_d)
        if not orders_js:
            st.warning(f"\u26a0\ufe0f No orders found for **{sel_date_str}** in the Base sheet.")
        else:
            n_trips = len(set(o["trip"] for o in orders_js))
            st.success(f"\u2705 **{len(orders_js)} order(s)** \u00b7 **{n_trips} trip group(s)**")
        st.session_state["rp_oj"]  = orders_js
        st.session_state["rp_dj"]  = drivers_js
        st.session_state["rp_dt"]  = sel_date_str
        st.session_state["rp_sub"] = False

    orders_js  = st.session_state.get("rp_oj",  [])
    drivers_js = st.session_state.get("rp_dj",  [])
    cached_dt  = st.session_state.get("rp_dt",  sel_date_str)

    if orders_js:
        m1,m2,m3,m4 = st.columns(4)
        m1.metric("Orders",      len(orders_js))
        m2.metric("Trip Groups", len(set(o["trip"] for o in orders_js)))
        m3.metric("Crates",      f'{sum(o.get("crates",0) for o in orders_js):.0f}')
        m4.metric("Kg",          f'{sum(o.get("tonnage",0) for o in orders_js):,.1f}')
    st.divider()

    # Inject data into the HTML component via inline script
    init_js = f"""<script>
window.addEventListener('load', function() {{
  var payload = {_json.dumps({"type":"GARLIC_LOAD_DATA","orders":orders_js,"drivers":drivers_js,"date":cached_dt})};
  setTimeout(function() {{
    if (typeof receiveFromStreamlit==='function') receiveFromStreamlit(JSON.stringify(payload));
  }}, 500);
}});
</script>"""
    html_out = _GARLIC_ROUTE_HTML.replace("</body>", init_js+"\n</body>", 1)
    _stc.html(html_out, height=1100, scrolling=True)

    st.divider()
    st.markdown("### \u2705 Confirm & Create Trips in System")
    st.caption("After reviewing optimized routes in the planner above (\u2192 **\u2461 RouteMap**, **\u2462 Live Map**, **\u2705 Submit** tabs), confirm here to write trips to Google Sheets.")

    if not orders_js:
        st.info("\U0001f4e5 Load orders using the date filter above.")
        return

    tg = {}
    for o in orders_js:
        tg.setdefault(str(o["trip"]), []).append(o)
    tkeys = sorted(tg.keys(), key=lambda x: int(x))

    df_d2     = all_drivers()
    do_opts   = ["\u25a1 Assign later"]
    do_ids    = [""]
    do_names  = [""]
    if not df_d2.empty:
        for _, dr in df_d2.iterrows():
            s   = str(dr.get("Active Status","Offline"))
            ico = "\U0001f7e2" if s.lower()=="active" else "\u26ab"
            do_opts.append(f"{ico} {dr['Full Name']} | {dr['Driver ID']} | {s}")
            do_ids.append(str(dr["Driver ID"]))
            do_names.append(str(dr["Full Name"]))

    with st.form("rp_cf"):
        st.markdown(f"**{len(tkeys)} trip(s) ready for {cached_dt}**")
        tcfgs = []
        for i, tk in enumerate(tkeys):
            stops = tg[tk]
            with st.expander(
                f"Trip {tk} \u2014 {len(stops)} stop(s) \u00b7 "
                f"{sum(o.get('crates',0) for o in stops):.0f} crates \u00b7 "
                f"{sum(o.get('tonnage',0) for o in stops):.1f} kg",
                expanded=(i==0),
            ):
                c1,c2,c3 = st.columns([2,2,3])
                with c1:
                    tid = st.text_input("Trip ID *",
                        value=f"TRP-{cached_dt.replace('-','')} -{_uuid2.uuid4().hex[:4].upper()}",
                        key=f"rp_tid_{tk}_{i}")
                with c2:
                    try:    tdt = st.date_input("Date", value=_d2.fromisoformat(cached_dt), key=f"rp_tdt_{tk}_{i}")
                    except: tdt = _d2.today()
                with c3:
                    dsel  = st.selectbox("Driver", do_opts, key=f"rp_drv_{tk}_{i}")
                    didx  = do_opts.index(dsel)
                    duid  = do_ids[didx]; dname = do_names[didx]
                prev = " \u2192 ".join(f"`{o['customer'][:16]}`" for o in stops[:4])
                if len(stops)>4: prev += f" +{len(stops)-4} more"
                st.markdown(f"**Stops:** {prev}")
                tcfgs.append({"tk":tk,"tid":tid,"date":str(tdt),"stops":stops,
                               "duid":duid,"dname":dname})
        st.divider()
        sub = st.form_submit_button("\u2705 Create All Trips in System",
                                     type="primary", use_container_width=True)

    if sub:
        errs, created = [], []
        pb = st.progress(0, text="Creating trips\u2026")
        for idx, cfg in enumerate(tcfgs):
            pb.progress(idx/len(tcfgs), text=f"Creating {cfg['tid']}\u2026")
            tid2 = cfg["tid"].strip()
            if not tid2: errs.append(f"Trip {cfg['tk']}: empty ID."); continue
            if col_exists("trips","Trip ID",tid2): errs.append(f"**{tid2}** already exists."); continue
            sids = ",".join(str(o.get("custId") or o.get("id","")).strip() for o in cfg["stops"])
            city = _infer_city_rp(cfg["stops"])
            stat = "Assigned" if cfg["duid"] else "Unassigned"
            append_row("trips",[tid2,cfg["date"],city,sids,cfg["duid"],cfg["dname"],
                                 stat,user["uid"],_dt2.now().strftime("%Y-%m-%d %H:%M:%S")])
            write_admin_log(user["uid"],user.get("email",""),"CREATE TRIP (ROUTE PLANNER)",
                            "Trip",tid2,"",cfg["dname"] or "Unassigned",
                            f"{len(cfg['stops'])} stops \u00b7 {cfg['date']} \u00b7 Route Planner")
            created.append(tid2)
        pb.progress(1.0, text="Done!")
        if created:
            st.success(f"\u2705 **{len(created)} trip(s) created:** "+" ".join(f"`{t}`" for t in created))
            st.balloons()
            st.session_state["task_done"]=True
            st.session_state["rp_oj"]=[]
            st.session_state["rp_sub"]=True
        for e in errs:
            st.error(f"\u274c {e}")


# ═══════════════════════════════════════════════════════════════════════════════
#  PAGE: ADMIN
# ═══════════════════════════════════════════════════════════════════════════════
def page_admin():
    user = st.session_state.user
    topbar("🛡️ Admin","#185fa5")
    tabs = st.tabs(["📦 SKUs","🗺️ Trips","🚚 Assign Drivers",
                    "👤 Customers","🚗 Driver Onboard","📋 Orders","📝 Audit Log","🛣️ Route Planner"])

    # ──────────────────────────────────────────────────────────────────────────
    # TAB 0 — SKUs
    # ──────────────────────────────────────────────────────────────────────────
    with tabs[0]:
        st.markdown(sl("📦 SKU Master"), unsafe_allow_html=True)
        df_sku = read_sheet("skus")
        if not df_sku.empty:
            act = len(df_sku[df_sku["Active"].astype(str).str.lower()=="true"])
            avg = df_sku["Price"].apply(lambda x: float(str(x).replace("₹","").replace(",","") or 0)).mean()
            c1,c2,c3,c4 = st.columns(4)
            c1.metric("Total SKUs",len(df_sku)); c2.metric("Active",act)
            c3.metric("Disabled",len(df_sku)-act); c4.metric("Avg price",f"₹{avg:,.2f}")
        with st.expander("➕ Add new SKU"):
            sc1,sc2,sc3 = st.columns(3)
            with sc1:
                sk_code  = st.text_input("SKU code *",  placeholder="GRLIC-1KG",      key="sk_c")
                sk_name  = st.text_input("SKU name *",  placeholder="Garlic 1KG Pack", key="sk_n")
            with sc2:
                sk_price = st.number_input("Price ₹ *", min_value=0.0, step=1.0,       key="sk_p")
                sk_wt    = st.selectbox("Weight type",  ["KG","Gram","Box","Piece","Dozen"], key="sk_w")
            with sc3:
                sk_cat   = st.text_input("Category",    placeholder="Garlic",          key="sk_cat")
            if st.button("➕ Add SKU", type="primary", key="sk_add"):
                if not sk_code or not sk_name or sk_price<=0:
                    st.error("Code, name and price required.")
                elif col_exists("skus","SKU Code",sk_code):
                    st.error("SKU code already exists.")
                else:
                    append_row("skus",[sk_code,sk_name,sk_price,sk_wt,
                                       sk_cat or "General","true",
                                       user["uid"],str(date.today())])
                    load_skus.clear()
                    write_admin_log(user["uid"],user.get("email",""),"ADD SKU","SKU",sk_code,"","",sk_name)
                    st.success(f"✅ SKU **{sk_code}** added!")
                    st.session_state.task_done = True; st.rerun()
        df_sku = read_sheet("skus")
        if df_sku.empty:
            st.info("No SKUs yet.")
        else:
            for idx, row in df_sku.iterrows():
                c1,c2,c3,c4,c5,c6 = st.columns([1.5,2.5,1.2,1.5,0.8,1.5])
                c1.markdown(f"**`{row['SKU Code']}`**")
                c2.write(str(row["SKU Name"])); c3.write(str(row.get("Weight Type","")))
                cur_p = float(str(row["Price"]).replace("₹","").replace(",","") or 0)
                new_p = c4.number_input("₹",value=cur_p,step=1.0,key=f"skp{idx}",label_visibility="collapsed")
                is_act = str(row.get("Active","")).lower()=="true"
                c5.markdown(pill("ON","pill-on") if is_act else pill("OFF","pill-off"),unsafe_allow_html=True)
                if c6.button("Disable" if is_act else "Enable", key=f"skt{idx}"):
                    update_row("skus","SKU Code",row["SKU Code"],
                               {"Active":"false" if is_act else "true","Price":new_p})
                    write_admin_log(user["uid"],user.get("email",""),
                                    ("Disable" if is_act else "Enable")+" SKU",
                                    "SKU",row["SKU Code"],cur_p,new_p,"")
                    load_skus.clear(); st.session_state.task_done=True; st.rerun()
                st.divider()

    # ──────────────────────────────────────────────────────────────────────────
    # TAB 1 — Trips  (UPDATED: Route File Upload)
    # ──────────────────────────────────────────────────────────────────────────
    with tabs[1]:
        st.markdown(sl("🗺️ Trips & Routes"), unsafe_allow_html=True)

        custs_df = load_customers()

        # ── Sub-tabs inside Trips ────────────────────────────────────────────
        trip_sub = st.tabs(["📂 Upload Route File", "✏️ Manual Trip", "📋 All Trips"])

        # ════════════════════════════════════════
        # SUB-TAB A: Upload Route File
        # ════════════════════════════════════════
        with trip_sub[0]:
            st.markdown("#### 📂 Upload a Route File to Create a Trip")

            # ── Column spec box ──────────────────────────────────────────────
            st.markdown(
                '<div class="template-box">'
                '📄 <strong>Required columns:</strong> &nbsp;'
                '<code>SaleOrderId</code> &nbsp;'
                '<code>DeliveryDate</code> &nbsp;'
                '<code>CustomerId</code> &nbsp;'
                '<code>Customer</code> &nbsp;'
                '<code>Latitude</code> &nbsp;'
                '<code>Longitude</code>'
                '<br><span style="color:#854f0b">Optional:</span> &nbsp;'
                '<code>Slot</code> &nbsp;'
                '<code>Driver</code> &nbsp;'
                '<code>TotalCrates</code> &nbsp;'
                '<code>OrderKg</code> &nbsp;'
                '<code>FC Latitude</code> &nbsp;'
                '<code>FC Longitude</code>'
                '<br><span style="color:#5a7a65;font-size:.82rem">'
                '💡 Accepts comma-separated .csv, tab-separated .csv/.tsv, and .xlsx'
                '</span>'
                '</div>', unsafe_allow_html=True)

            # ── Template downloads — wrapped so a missing openpyxl never crashes the page ──
            try:
                _xlsx_bytes = generate_route_template_excel()
            except Exception:
                _xlsx_bytes = None

            dl_col1, dl_col2 = st.columns(2)
            with dl_col1:
                if _xlsx_bytes:
                    st.download_button(
                        "⬇️ Download Excel Template",
                        data=_xlsx_bytes,
                        file_name="route_template.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="dl_tpl_xlsx",
                    )
                else:
                    st.caption("📋 Excel template unavailable (openpyxl missing in env)")
            with dl_col2:
                st.download_button(
                    "⬇️ Download CSV Template",
                    data=ROUTE_TEMPLATE_CSV,
                    file_name="route_template.csv",
                    mime="text/csv",
                    key="dl_tpl_csv",
                )

            st.divider()

            # ── File uploader ────────────────────────────────────────────────
            uploaded = st.file_uploader(
                "Upload your Route File (.csv, .tsv or .xlsx)",
                type=["csv","tsv","xlsx","xls"],
                key="route_uploader",
                help="Each row = one delivery stop. CustomerId should match Customer Onboard Data.",
            )

            if uploaded:
                with st.spinner("Parsing route file…"):
                    df_route, parse_errors, parse_warnings = parse_route_file(uploaded)

                if parse_errors:
                    for e in parse_errors:
                        st.error(f"❌ {e}")
                else:
                    # Cross-check against customer master (non-blocking)
                    df_flagged, unmatched_ids = validate_route_against_customers(df_route, custs_df)

                    # Show parse warnings
                    for w in parse_warnings:
                        st.warning(f"⚠️ {w}")

                    # Unmatched customer IDs — warn but keep all rows
                    if unmatched_ids:
                        st.warning(
                            f"⚠️ **{len(unmatched_ids)} CustomerId(s)** not found in Customer Onboard Data — "
                            f"rows kept but marked. IDs: `{', '.join(str(x) for x in unmatched_ids[:10])}`"
                            + (" …" if len(unmatched_ids) > 10 else "")
                        )

                    total_rows = len(df_flagged)
                    matched_rows   = int(df_flagged["_matched"].sum())
                    unmatched_rows = total_rows - matched_rows

                    # ── File-level summary metrics ────────────────────────────
                    sm1, sm2, sm3, sm4 = st.columns(4)
                    sm1.metric("Total Stops",   total_rows)
                    sm2.metric("✅ Matched",     matched_rows)
                    sm3.metric("⚠️ Unmatched",  unmatched_rows)
                    # Aggregate TotalCrates and OrderKg if present
                    if "TotalCrates" in df_flagged.columns:
                        try:
                            total_crates = pd.to_numeric(df_flagged["TotalCrates"], errors="coerce").sum()
                            sm4.metric("📦 Total Crates", f"{total_crates:,.0f}")
                        except Exception:
                            sm4.metric("📦 Total Crates", "—")

                    st.divider()

                    # ── Route Preview Table ───────────────────────────────────
                    st.markdown(
                        f'<div class="route-card">'
                        f'<div class="route-card-header">'
                        f'🗺️ Route Preview — {total_rows} Stop(s) from <em>{uploaded.name}</em>'
                        f'</div>',
                        unsafe_allow_html=True)

                    for i, row in df_flagged.iterrows():
                        cid        = str(row.get("CustomerId", "")).strip()
                        customer   = str(row.get("Customer", "")).strip()
                        sale_oid   = str(row.get("SaleOrderId", "")).strip()
                        slot       = str(row.get("Slot", "")).strip()
                        lat        = row.get("Latitude", "")
                        lng        = row.get("Longitude", "")
                        crates     = row.get("TotalCrates", "")
                        kg         = row.get("OrderKg", "")
                        is_matched = bool(row.get("_matched", False))

                        # Enrich address / city from customer master
                        enrich     = enrich_route_row(row, custs_df)
                        city_label = enrich["city"] or ""
                        addr_label = enrich["address"] or ""

                        match_ico  = "✅" if is_matched else "⚠️"
                        lat_str    = f"{float(lat):.5f}" if pd.notna(lat) and str(lat) not in ("","nan") else "?"
                        lng_str    = f"{float(lng):.5f}" if pd.notna(lng) and str(lng) not in ("","nan") else "?"

                        details = []
                        if slot and slot != "nan":       details.append(f"🕐 {slot}")
                        if crates and str(crates) not in ("","nan"): details.append(f"📦 {crates} crates")
                        if kg    and str(kg)     not in ("","nan"): details.append(f"⚖️ {kg} kg")
                        if city_label: details.append(f"📍 {city_label}")
                        detail_str = " &nbsp;·&nbsp; ".join(details)

                        st.markdown(
                            f'<div class="route-row">'
                            f'<span class="stop-badge">{i+1}</span>'
                            f'<span style="min-width:14px">{match_ico}</span>'
                            f'<div style="flex:1">'
                            f'<strong>{customer}</strong>'
                            f'&nbsp;<span style="color:#5a7a65;font-size:.8rem">({cid})</span>'
                            f'&nbsp;·&nbsp;<code style="font-size:.78rem">{sale_oid}</code>'
                            + (f'<br><span style="color:#5a7a65;font-size:.8rem">{detail_str}</span>' if detail_str else "")
                            + f'</div>'
                            f'<span style="color:#185fa5;font-size:.78rem;white-space:nowrap">'
                            f'{lat_str}, {lng_str}</span>'
                            f'</div>',
                            unsafe_allow_html=True)

                    st.markdown('</div>', unsafe_allow_html=True)

                    # ── Full raw data expander ────────────────────────────────
                    with st.expander("📊 View full raw data table"):
                        display_cols = [c for c in ROUTE_ALL_COLS if c in df_flagged.columns]
                        st.dataframe(
                            df_flagged[display_cols],
                            use_container_width=True, hide_index=True)

                    # ── Trip creation settings ────────────────────────────────
                    st.divider()
                    st.markdown("#### ⚙️ Trip Settings")

                    rf1, rf2, rf3 = st.columns(3)
                    with rf1:
                        # Auto-fill Trip ID from file's DeliveryDate if parseable
                        auto_date_str = ""
                        if "DeliveryDate" in df_flagged.columns:
                            try:
                                first_date = pd.to_datetime(
                                    df_flagged["DeliveryDate"].dropna().iloc[0], dayfirst=False
                                )
                                auto_date_str = first_date.strftime("%Y%m%d")
                            except Exception:
                                pass
                        auto_trip_id = f"TRP-{auto_date_str or datetime.now().strftime('%Y%m%d')}-{uuid.uuid4().hex[:4].upper()}"
                        rf_trip_id   = st.text_input(
                            "Trip ID *", value=auto_trip_id, key="rf_trip_id",
                            help="Auto-generated from DeliveryDate — edit if needed")

                    with rf2:
                        # Auto-fill date from file
                        rf_default_date = date.today()
                        if "DeliveryDate" in df_flagged.columns:
                            try:
                                rf_default_date = pd.to_datetime(
                                    df_flagged["DeliveryDate"].dropna().iloc[0], dayfirst=False
                                ).date()
                            except Exception:
                                pass
                        rf_date = st.date_input("Trip Date *", value=rf_default_date, key="rf_date")

                    with rf3:
                        allowed_cities = ["Bengaluru","Mysuru","Hubli","Mangaluru","Hassan","Tumkur"]
                        rf_city = st.selectbox("City *", allowed_cities, key="rf_city")

                    # ── Driver auto-suggest from file ─────────────────────────
                    file_driver_name = ""
                    if "Driver" in df_flagged.columns:
                        drv_vals = df_flagged["Driver"].dropna()
                        drv_vals = drv_vals[~drv_vals.isin(["","nan","None"])]
                        if not drv_vals.empty:
                            file_driver_name = drv_vals.mode().iloc[0]

                    all_d = all_drivers()
                    rf_drv_uid = ""; rf_drv_name = ""
                    if not all_d.empty:
                        # Try to pre-select driver whose name matches file's Driver column
                        pre_idx = 0
                        if file_driver_name:
                            for _di, _dr in enumerate(all_d["Full Name"].tolist()):
                                if file_driver_name.strip().lower() in _dr.lower() or \
                                   _dr.lower() in file_driver_name.strip().lower():
                                    pre_idx = _di + 1   # +1 because index 0 = "Assign later"
                                    break

                        drv_opts  = ["⬜ Assign later"] + [
                            f"{'🟢' if str(r.get('Active Status','')).lower()=='active' else '⚫'} "
                            f"{r['Full Name']}  ({r['Driver ID']})"
                            for _, r in all_d.iterrows()
                        ]
                        drv_ids   = [""] + all_d["Driver ID"].tolist()
                        drv_names = [""] + all_d["Full Name"].tolist()

                        if file_driver_name:
                            st.info(f"💡 Route file suggests driver: **{file_driver_name}**")

                        sel_drv = st.selectbox(
                            "Assign Driver (optional — can reassign later)",
                            drv_opts, index=min(pre_idx, len(drv_opts)-1),
                            key="rf_driver_sel")
                        drv_idx     = drv_opts.index(sel_drv)
                        rf_drv_uid  = drv_ids[drv_idx]
                        rf_drv_name = drv_names[drv_idx]
                    else:
                        st.info("No drivers onboarded yet — you can assign one after creating the trip.")

                    st.divider()

                    # ── Trip summary card ─────────────────────────────────────
                    total_kg = ""
                    total_cr = ""
                    if "OrderKg" in df_flagged.columns:
                        try: total_kg = f"{pd.to_numeric(df_flagged['OrderKg'], errors='coerce').sum():,.1f} kg"
                        except Exception: pass
                    if "TotalCrates" in df_flagged.columns:
                        try: total_cr = f"{pd.to_numeric(df_flagged['TotalCrates'], errors='coerce').sum():,.0f} crates"
                        except Exception: pass

                    st.markdown(
                        f'<div class="route-card">'
                        f'<div class="route-card-header">📋 Trip Summary — confirm before saving</div>'
                        f'<div class="route-row"><span style="width:140px;color:#5a7a65">Trip ID</span>'
                        f'<strong>{rf_trip_id}</strong></div>'
                        f'<div class="route-row"><span style="width:140px;color:#5a7a65">Delivery Date</span>'
                        f'<strong>{rf_date}</strong></div>'
                        f'<div class="route-row"><span style="width:140px;color:#5a7a65">City</span>'
                        f'<strong>{rf_city}</strong></div>'
                        f'<div class="route-row"><span style="width:140px;color:#5a7a65">Total Stops</span>'
                        f'<strong>{total_rows}</strong>'
                        f'&nbsp;({matched_rows} matched &nbsp;/&nbsp; {unmatched_rows} unmatched)</div>'
                        + (f'<div class="route-row"><span style="width:140px;color:#5a7a65">Total Kg</span>'
                           f'<strong>{total_kg}</strong></div>' if total_kg else "")
                        + (f'<div class="route-row"><span style="width:140px;color:#5a7a65">Total Crates</span>'
                           f'<strong>{total_cr}</strong></div>' if total_cr else "")
                        + f'<div class="route-row"><span style="width:140px;color:#5a7a65">Driver</span>'
                        f'<strong>{rf_drv_name or "To be assigned"}</strong></div>'
                        f'<div class="route-row"><span style="width:140px;color:#5a7a65">Source file</span>'
                        f'<span style="color:#5a7a65">{uploaded.name}</span></div>'
                        f'</div>', unsafe_allow_html=True)

                    if st.button("✅ Create Trip from Route File",
                                 type="primary", use_container_width=True,
                                 key="rf_create_trip"):
                        if not rf_trip_id.strip():
                            st.error("Trip ID is required.")
                        elif col_exists("trips", "Trip ID", rf_trip_id.strip()):
                            st.error(f"Trip ID **{rf_trip_id}** already exists. Please change it.")
                        else:
                            # Store CustomerId list as the shops string
                            shop_ids_str = ",".join(
                                df_flagged["CustomerId"].astype(str).str.strip().tolist()
                            )
                            status = "Assigned" if rf_drv_uid else "Unassigned"
                            append_row("trips", [
                                rf_trip_id.strip(),
                                str(rf_date),
                                rf_city,
                                shop_ids_str,
                                rf_drv_uid,
                                rf_drv_name,
                                status,
                                user["uid"],
                                datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                            ])
                            write_admin_log(
                                user["uid"], user.get("email", ""),
                                "CREATE TRIP (ROUTE FILE)", "Trip",
                                rf_trip_id.strip(), "",
                                rf_drv_name or "Unassigned",
                                f"{total_rows} stops | {total_kg or '?'} | {total_cr or '?'} | {uploaded.name}"
                            )
                            drv_msg = f" · Driver: **{rf_drv_name}**" if rf_drv_name else " · Driver: **to be assigned**"
                            st.success(
                                f"✅ Trip **{rf_trip_id}** created — "
                                f"**{total_rows} stop(s)**{drv_msg}!"
                            )
                            st.session_state.task_done = True
                            st.balloons()
                            st.rerun()

            else:
                st.markdown(
                    '<div class="upload-zone">'
                    '<span style="font-size:2.2rem">📂</span><br>'
                    '<strong style="font-family:Syne,sans-serif;color:#1a7f4b">'
                    'Drop your route file here</strong><br>'
                    '<span style="color:#5a7a65;font-size:.88rem">'
                    'Accepts .csv (comma or tab), .tsv, .xlsx — '
                    'download the template above to get started'
                    '</span>'
                    '</div>', unsafe_allow_html=True)

        # ════════════════════════════════════════
        # SUB-TAB B: Manual Trip
        # ════════════════════════════════════════
        with trip_sub[1]:
            st.markdown("#### ✏️ Create Trip Manually")
            with st.container():
                tc1,tc2 = st.columns(2)
                with tc1:
                    tr_id   = st.text_input("Trip ID *", placeholder="TRP-001", key="tr_id")
                    tr_date = st.date_input("Date *", value=date.today(), key="tr_date")
                with tc2:
                    tr_city = st.selectbox("City",
                        ["Bengaluru","Mysuru","Hubli","Mangaluru","Hassan","Tumkur"],key="tr_city")

                if not custs_df.empty:
                    shop_opts = custs_df.apply(
                        lambda r: f"{r['CUST-ID']} — {r['Shop Name']} ({r['City']})",axis=1).tolist()
                    cust_ids  = custs_df["CUST-ID"].tolist()
                    sel_shops = st.multiselect("Select shops * (multiple allowed)",shop_opts,key="tr_shops")
                    sel_ids   = [cust_ids[shop_opts.index(s)] for s in sel_shops]
                    if sel_ids:
                        st.info(f"✅ {len(sel_ids)} shop(s): {', '.join(sel_ids)}")
                else:
                    st.warning("No customers onboarded yet. Use the **Customer Onboard** tab.")
                    sel_ids = []

                if st.button("✅ Create Trip", type="primary", key="tr_btn"):
                    if not tr_id: st.error("Trip ID required.")
                    elif not sel_ids: st.error("Select at least one shop.")
                    elif col_exists("trips","Trip ID",tr_id): st.error("Trip ID already exists.")
                    else:
                        append_row("trips",[tr_id,str(tr_date),tr_city,",".join(sel_ids),
                                            "","","Unassigned",user["uid"],
                                            datetime.now().strftime("%Y-%m-%d %H:%M:%S")])
                        write_admin_log(user["uid"],user.get("email",""),
                                        "CREATE TRIP","Trip",tr_id,"","",f"{len(sel_ids)} shops")
                        st.success(f"✅ Trip **{tr_id}** created with {len(sel_ids)} shop(s)!")
                        st.session_state.task_done=True; st.rerun()

        # ════════════════════════════════════════
        # SUB-TAB C: All Trips (with inline driver assign)
        # ════════════════════════════════════════
        with trip_sub[2]:
            st.markdown("#### 📋 All Trips")
            trips_df = read_sheet("trips")

            if trips_df.empty:
                st.info("No trips yet. Create one using Upload or Manual tab above.")
            else:
                # Metrics
                total_t = len(trips_df)
                unassigned_t = len(trips_df[trips_df["Driver Name"].astype(str).str.strip().isin(["","nan"])])
                assigned_t   = total_t - unassigned_t
                completed_t  = len(trips_df[trips_df["Status"].astype(str).str.lower()=="completed"])
                tm1,tm2,tm3,tm4 = st.columns(4)
                tm1.metric("Total Trips", total_t)
                tm2.metric("🟢 Driver Assigned", assigned_t)
                tm3.metric("⚠️ Unassigned", unassigned_t)
                tm4.metric("✅ Completed", completed_t)
                st.divider()

                # Filter
                flt_status = st.selectbox(
                    "Filter by status",
                    ["All","Unassigned","Assigned","In Progress","Completed"],
                    key="trip_list_filter")
                disp_trips = trips_df.copy()
                if flt_status != "All":
                    disp_trips = disp_trips[
                        disp_trips["Status"].astype(str).str.lower() == flt_status.lower()
                    ]

                all_d_for_assign = all_drivers()

                for _, t in disp_trips.iterrows():
                    shop_ids = [s.strip() for s in str(t.get("Shops","")).split(",") if s.strip()]
                    drv_name = str(t.get("Driver Name","")).strip()
                    drv_uid  = str(t.get("Driver UID","")).strip()
                    t_status = str(t.get("Status","")).strip()

                    status_cls = {
                        "completed": "pill-done",
                        "in progress": "pill-part",
                        "assigned": "pill-on",
                        "unassigned": "pill-pend",
                    }.get(t_status.lower(), "pill-pend")

                    with st.container():
                        hc1, hc2, hc3, hc4 = st.columns([3, 2, 2, 2])
                        with hc1:
                            st.markdown(
                                f'<strong style="font-family:Syne,sans-serif">{t["Trip ID"]}</strong>'
                                f'&nbsp;{pill(t_status, status_cls)}',
                                unsafe_allow_html=True)
                            st.caption(f"📅 {t.get('Date','')} · 📍 {t.get('City','')} · {len(shop_ids)} stop(s)")
                        with hc2:
                            st.caption("Driver")
                            st.markdown(
                                f'<strong>{drv_name if drv_name and drv_name!="nan" else "—"}</strong>'
                                + (f'<br><code style="font-size:.72rem">{drv_uid}</code>'
                                   if drv_uid and drv_uid!="nan" else ""),
                                unsafe_allow_html=True)
                        with hc3:
                            st.caption("Created by")
                            st.write(str(t.get("Created By","")))
                        with hc4:
                            st.caption("Created at")
                            st.write(str(t.get("Created At","")))

                    # Expandable stops list
                    with st.expander(f"🔍 View {len(shop_ids)} Stop(s) for {t['Trip ID']}"):
                        if custs_df.empty:
                            st.write(", ".join(shop_ids))
                        else:
                            for si, sid in enumerate(shop_ids):
                                cm = custs_df[custs_df["CUST-ID"]==sid]
                                if not cm.empty:
                                    r = cm.iloc[0]
                                    st.markdown(
                                        f'<div class="route-row">'
                                        f'<span class="stop-badge">{si+1}</span>'
                                        f'<strong>{r.get("Shop Name","")}</strong>'
                                        f'&nbsp;<span style="color:#5a7a65;font-size:.83rem">({sid})</span>'
                                        f'&nbsp;·&nbsp;<span style="color:#185fa5">{r.get("City","")}</span>'
                                        f'&nbsp;·&nbsp;<span style="color:#854f0b;font-size:.8rem">'
                                        f'{r.get("Shop Address","")}</span>'
                                        f'</div>', unsafe_allow_html=True)
                                else:
                                    st.markdown(
                                        f'<div class="route-row">'
                                        f'<span class="stop-badge">{si+1}</span>'
                                        f'<span style="color:#842029">{sid} (not found in master)</span>'
                                        f'</div>', unsafe_allow_html=True)

                        # ── Inline driver assign ──────────────────────────────
                        if t_status.lower() not in ("completed",) and not all_d_for_assign.empty:
                            st.divider()
                            st.markdown("**Assign / Change Driver**")
                            ia1, ia2 = st.columns([4, 2])
                            with ia1:
                                drv_labels_ia = []
                                drv_ids_ia    = []
                                drv_names_ia  = []
                                for _, dr in all_d_for_assign.iterrows():
                                    s   = str(dr.get("Active Status","Offline"))
                                    ico = "🟢" if s.lower()=="active" else "⚫"
                                    drv_labels_ia.append(
                                        f"{ico} {dr['Full Name']} | {dr['Driver ID']} | {s}")
                                    drv_ids_ia.append(dr["Driver ID"])
                                    drv_names_ia.append(dr["Full Name"])

                                # Pre-select current driver if already assigned
                                cur_idx = 0
                                if drv_uid and drv_uid in drv_ids_ia:
                                    cur_idx = drv_ids_ia.index(drv_uid)

                                sel_ia = st.selectbox(
                                    "Select driver",
                                    drv_labels_ia,
                                    index=cur_idx,
                                    key=f"ia_drv_{t['Trip ID']}",
                                    label_visibility="collapsed")
                                ia_idx      = drv_labels_ia.index(sel_ia)
                                ia_drv_id   = drv_ids_ia[ia_idx]
                                ia_drv_name = drv_names_ia[ia_idx]

                            with ia2:
                                st.write("")
                                if st.button(
                                    "✅ Assign",
                                    key=f"ia_btn_{t['Trip ID']}",
                                    type="primary",
                                    use_container_width=True):
                                    update_row("trips","Trip ID",t["Trip ID"],{
                                        "Driver UID":  ia_drv_id,
                                        "Driver Name": ia_drv_name,
                                        "Status":      "Assigned",
                                    })
                                    write_admin_log(
                                        user["uid"],user.get("email",""),
                                        "ASSIGN DRIVER","Trip",t["Trip ID"],
                                        drv_name,ia_drv_id,ia_drv_name)
                                    st.success(
                                        f"✅ **{ia_drv_name}** assigned to **{t['Trip ID']}**!")
                                    st.session_state.task_done = True
                                    st.rerun()

                    st.divider()

    # ──────────────────────────────────────────────────────────────────────────
    # TAB 2 — Assign Drivers
    # ──────────────────────────────────────────────────────────────────────────
    with tabs[2]:
        st.markdown(sl("🚚 Assign Drivers to Trips"), unsafe_allow_html=True)
        all_d    = all_drivers()
        act_d    = active_drivers()
        trips_df = read_sheet("trips")
        st.markdown("#### All registered drivers")
        if all_d.empty:
            st.info("No drivers onboarded yet. Use the **Driver Onboard** tab.")
        else:
            st.success(f"🟢 {len(act_d)} active  ⚫ {len(all_d)-len(act_d)} offline")
            for _,r in all_d.iterrows():
                is_on=str(r.get("Active Status","")).lower()=="active"
                c1,c2,c3,c4=st.columns([2,2,2,1])
                c1.markdown(f"**{r['Full Name']}**")
                c2.write(f"`{r['Driver ID']}` · {r.get('Vehicle Type','')} {r.get('Vehicle Number','')}")
                c3.write(f"Last active: {r.get('Last Active','—')}")
                c4.markdown(pill("Active","pill-on") if is_on else pill("Offline","pill-off"),unsafe_allow_html=True)
        st.divider()
        st.markdown("#### Assign / Reassign driver to trip")
        if trips_df.empty:
            st.info("No trips yet.")
        elif all_d.empty:
            st.info("No drivers yet.")
        else:
            custs=load_customers()
            ac1,ac2=st.columns(2)
            with ac1:
                trip_labels=[]
                for _,t in trips_df.iterrows():
                    n=len([s for s in str(t.get("Shops","")).split(",") if s.strip()])
                    drv=str(t.get("Driver Name","")).strip()
                    trip_labels.append(f"{t['Trip ID']}  ({n} shops, {t['City']}, {t['Date']}) → {drv or 'Unassigned'}")
                trip_ids=trips_df["Trip ID"].tolist()
                sel_lbl_t=st.selectbox("Trip",trip_labels,key="asgn_trip_sel")
                sel_trip=trip_ids[trip_labels.index(sel_lbl_t)]
                t_row=trips_df[trips_df["Trip ID"]==sel_trip].iloc[0]
                shop_ids=[s.strip() for s in str(t_row.get("Shops","")).split(",") if s.strip()]
                st.info(f"📦 {len(shop_ids)} shop(s) in this trip:")
                for sid in shop_ids:
                    if not custs.empty:
                        m=custs[custs["CUST-ID"]==sid]
                        st.caption(f"  • {m.iloc[0]['Shop Name']} — {m.iloc[0]['City']} — `{sid}`" if not m.empty else f"  • `{sid}`")
                    else: st.caption(f"  • `{sid}`")
            with ac2:
                drv_labels=[]; drv_ids=[]; drv_names=[]
                for _,r in all_d.iterrows():
                    s=str(r.get("Active Status","Offline"))
                    ico="🟢" if s.lower()=="active" else "⚫"
                    drv_labels.append(f"{ico} {r['Full Name']}  |  {r['Driver ID']}  |  {r.get('Vehicle Type','')}  |  {s}")
                    drv_ids.append(r["Driver ID"]); drv_names.append(r["Full Name"])
                sel_lbl_d=st.selectbox("Driver",drv_labels,key="asgn_drv_sel")
                idx_d=drv_labels.index(sel_lbl_d)
                sel_drv_id=drv_ids[idx_d]; sel_drv_name=drv_names[idx_d]
            if st.button("✅ Assign Driver to Trip",type="primary",use_container_width=True,key="asgn_btn"):
                update_row("trips","Trip ID",sel_trip,
                           {"Driver UID":sel_drv_id,"Driver Name":sel_drv_name,"Status":"Assigned"})
                write_admin_log(user["uid"],user.get("email",""),
                                "ASSIGN DRIVER","Trip",sel_trip,"",sel_drv_id,sel_drv_name)
                st.success(f"✅ **{sel_drv_name}** assigned to **{sel_trip}**!")
                st.session_state.task_done=True; st.rerun()
            st.divider()
            st.markdown("#### All trip assignments")
            disp_t=trips_df[["Trip ID","Date","City","Driver Name","Driver UID","Status"]].copy()
            disp_t["Driver Name"]=disp_t["Driver Name"].apply(lambda v:v if str(v).strip() else "⚠️ Unassigned")
            st.dataframe(disp_t,use_container_width=True,hide_index=True)

    # ──────────────────────────────────────────────────────────────────────────
    # TAB 3 — Customers
    # ──────────────────────────────────────────────────────────────────────────
    with tabs[3]:
        st.markdown(sl("👤 Customer Onboard Data"), unsafe_allow_html=True)
        df_c=load_customers()
        if df_c.empty:
            st.info("No customers onboarded yet.")
        else:
            c1,c2,c3=st.columns(3)
            c1.metric("Total",len(df_c)); c2.metric("Active",len(df_c[df_c["Status"]=="Active"]))
            c3.metric("Cities",df_c["City"].nunique())
            st.dataframe(df_c,use_container_width=True,hide_index=True)

    # ──────────────────────────────────────────────────────────────────────────
    # TAB 4 — Driver Onboard
    # ──────────────────────────────────────────────────────────────────────────
    with tabs[4]:
        st.markdown(sl("🚗 Driver Onboard","amber"), unsafe_allow_html=True)
        all_d2=all_drivers()
        if not all_d2.empty:
            m1,m2,m3=st.columns(3)
            m1.metric("Total drivers",len(all_d2))
            m2.metric("Active now",len(all_d2[all_d2["Active Status"]=="Active"]))
            m3.metric("Offline",len(all_d2[all_d2["Active Status"]!="Active"]))
            st.markdown("#### All drivers")
            disp_d2=all_d2.copy()
            if "Account Number" in disp_d2.columns:
                disp_d2["Account Number"]=disp_d2["Account Number"].apply(
                    lambda v:("*"*(len(str(v))-4)+str(v)[-4:]) if len(str(v))>4 else "****")
            st.dataframe(disp_d2,use_container_width=True,hide_index=True)
            st.divider()
        ds1,ds2=st.columns([3,1])
        with ds1: do_sv=st.text_input("Search driver by mobile",key="adm_do_search")
        with ds2:
            st.write(""); st.write("")
            do_dos=st.button("🔍 Search",key="adm_do_search_btn")
        if do_dos and do_sv:
            ex=find_row("driver_onboard","Mobile",do_sv.strip())
            if ex:
                acct=str(ex.get("Account Number",""))
                masked=("*"*(len(acct)-4)+acct[-4:]) if len(acct)>4 else "****"
                st.success(f"✅ Found — Driver ID: **{ex['Driver ID']}**")
                c1,c2,c3=st.columns(3)
                c1.write(f"**Name:** {ex.get('Full Name')}"); c1.write(f"**Mobile:** {ex.get('Mobile')}")
                c2.write(f"**Vehicle:** {ex.get('Vehicle Type')} {ex.get('Vehicle Number','')}"); c2.write(f"**Status:** {ex.get('Active Status')}")
                c3.write(f"**Bank:** {ex.get('Bank Name')}"); c3.write(f"**Account:** {masked}")
            else:
                st.info("Not found — fill form below.")
        st.divider()
        st.markdown("#### Onboard new driver")
        dn1,dn2,dn3=st.columns(3)
        with dn1:
            do_name    = st.text_input("Full name *",        key="adm_do_name")
            do_mob     = st.text_input("Mobile *",           placeholder="10-digit", key="adm_do_mob")
            do_email   = st.text_input("Email",              key="adm_do_email")
        with dn2:
            do_veh     = st.selectbox("Vehicle type",
                ["Bike","Auto","Van","Truck","Mini-Truck"],  key="adm_do_veh")
            do_veh_num = st.text_input("Vehicle number *",   placeholder="e.g. KA-01-AB-1234", key="adm_do_veh_num")
            do_bank    = st.text_input("Bank name *",        key="adm_do_bank")
        with dn3:
            do_acct    = st.text_input("Account number *",   key="adm_do_acct")
            do_ifsc    = st.text_input("IFSC code *",        key="adm_do_ifsc")
            do_upi     = st.text_input("UPI ID",             placeholder="mobile@upi", key="adm_do_upi")
        st.caption("🔒 Bank details visible to admin only.")
        if st.button("✅ Onboard Driver", type="primary", use_container_width=True, key="adm_do_btn"):
            if not all([do_name,do_mob,do_veh_num,do_bank,do_acct,do_ifsc]):
                st.error("Fill all required (*) fields including Vehicle Number.")
            else:
                with st.spinner("Checking duplicates…"):
                    ex=find_row("driver_onboard","Mobile",do_mob.strip())
                if ex:
                    st.warning(f"⚠️ Mobile already registered — Driver ID: **{ex['Driver ID']}**")
                else:
                    did=gen_driver_id()
                    append_row("driver_onboard",[
                        did,do_name,do_mob,do_email,do_veh,do_veh_num,
                        do_bank,do_acct,do_ifsc,do_upi,
                        str(date.today()),"Offline",""])
                    write_admin_log(user["uid"],user.get("email",""),
                                    "ONBOARD DRIVER","Driver",did,"","",do_name)
                    st.success(f"✅ Driver onboarded! Permanent Driver ID: **`{did}`**")
                    st.session_state.task_done=True; st.rerun()

    # ──────────────────────────────────────────────────────────────────────────
    # TAB 5 — Orders
    # ──────────────────────────────────────────────────────────────────────────
    with tabs[5]:
        st.markdown(sl("📋 All Orders"), unsafe_allow_html=True)
        df_o=read_sheet("base")
        if df_o.empty:
            st.info("No orders yet.")
        else:
            today_str=str(date.today())
            show_all=st.checkbox("Show all orders (default: today only)", key="admin_show_all")
            if not show_all and "ORDER DATE" in df_o.columns:
                df_show=df_o[df_o["ORDER DATE"].astype(str)==today_str]
                st.caption(f"📅 Showing orders for **{today_str}** — {len(df_show)} order(s)")
            else:
                df_show=df_o
                st.caption(f"Showing all {len(df_show)} order(s)")
            c1,c2,c3,c4=st.columns(4)
            c1.metric("Shown",len(df_show))
            c2.metric("Pending",len(df_show[df_show["Delivery Status"]=="Pending"]) if not df_show.empty else 0)
            c3.metric("Delivered",len(df_show[df_show["Delivery Status"]=="Delivered"]) if not df_show.empty else 0)
            c4.metric("Failed/Partial",len(df_show[df_show["Delivery Status"].isin(["Failed","Partial"])]) if not df_show.empty else 0)
            if df_show.empty:
                st.info(f"No orders for {today_str}. Check 'Show all orders' to see history.")
            else:
                st.dataframe(df_show,use_container_width=True,hide_index=True)

    # ──────────────────────────────────────────────────────────────────────────
    # TAB 6 — Audit Log
    # ──────────────────────────────────────────────────────────────────────────
    with tabs[6]:
        st.markdown(sl("📝 Admin Audit Log","blue"), unsafe_allow_html=True)
        df_l=read_sheet("admin_log")
        if df_l.empty:
            st.info("No admin actions logged yet.")
        else:
            st.dataframe(df_l.sort_values("Timestamp",ascending=False),
                         use_container_width=True,hide_index=True)

    # ──────────────────────────────────────────────────────────────────────────
    # TAB 7 — Route Planner  (fully inlined — no external modules needed)
    # ──────────────────────────────────────────────────────────────────────────
    with tabs[7]:
        _page_route_planner_inline(user)



    st.divider()
    if not st.session_state.get("task_done",False):
        if st.button("✅ Mark All Tasks Done (enables Logout)",key="admin_tasks_done"):
            st.session_state.task_done=True; st.rerun()
    else:
        st.success("✅ Tasks marked complete — Logout button is now active above.")


# ═══════════════════════════════════════════════════════════════════════════════
#  PAGE: SALES EXECUTIVE (T1)
# ═══════════════════════════════════════════════════════════════════════════════
def page_sales():
    user = st.session_state.user
    topbar("🧑‍💼 Sales Executive · T1")
    tabs = st.tabs(["➕ New Order","👤 Onboard Customer","📋 My Orders"])

    with tabs[0]:
        st.markdown(sl("🔍 Customer Lookup"), unsafe_allow_html=True)
        lc1,lc2,lc3=st.columns([2,2,1])
        with lc1: lk_id =st.text_input("Customer ID",placeholder="CUST-XXXXXX",key="lk_id")
        with lc2: lk_mob=st.text_input("OR mobile number",placeholder="10-digit",key="lk_mob")
        with lc3:
            st.write(""); st.write("")
            do_lk=st.button("Fetch →",key="lk_btn")
        if do_lk:
            with st.spinner("Looking up…"):
                cust=(find_row("customer_onboard","CUST-ID",lk_id.strip()) if lk_id.strip()
                      else find_row("customer_onboard","Mobile",lk_mob.strip()))
            if cust:
                st.session_state.cust_data=cust
                st.success(f"✅ Found: **{cust.get('Full Name')}** — {cust.get('Shop Name')}")
            else:
                st.error("❌ Customer not found. Onboard them first.")
                st.session_state.cust_data={}
        cust=st.session_state.get("cust_data",{})
        if cust:
            auto_lat=str(cust.get("Latitude","")).strip()
            auto_lng=str(cust.get("Longitude","")).strip()
            auto_addr=str(cust.get("Shop Address","")).strip()
            st.markdown(
                f'<div class="loc-box">'
                f'📌 <strong>Location auto-loaded from Customer Onboard Data</strong><br>'
                f'Address: {auto_addr}<br>GPS: {auto_lat}, {auto_lng}'
                f'</div>', unsafe_allow_html=True)
        st.divider()
        st.markdown(sl("📦 Order Details"), unsafe_allow_html=True)
        oc1,oc2,oc3=st.columns(3)
        with oc1:
            o_id  =st.text_input("Order ID (auto)",value=gen_order_id(),disabled=True,key="o_id")
            o_date=st.date_input("Order date",value=date.today(),key="o_date")
        with oc2:
            cities=["Bengaluru","Mysuru","Hubli","Mangaluru","Hassan","Tumkur"]
            auto_city=cust.get("City","Bengaluru") if cust else "Bengaluru"
            ci=cities.index(auto_city) if auto_city in cities else 0
            o_city=st.selectbox("City *",cities,index=ci,key="o_city")
            st.text_input("⏰ Ordered time (auto on submit)",
                           value=datetime.now().strftime("%H:%M:%S"),
                           disabled=True,key="o_time_disp")
        with oc3:
            o_dcoff=st.time_input("Delivery cut-off",key="o_dcoff")
            o_sopen=st.time_input("Shop opens at",key="o_sopen")
        st.markdown(sl("👤 Customer Details"), unsafe_allow_html=True)
        cc1,cc2,cc3=st.columns(3)
        with cc1:
            st.text_input("Customer ID",    value=cust.get("CUST-ID",""),        disabled=True,key="c_id")
            st.text_input("Shop name",      value=cust.get("Shop Name",""),      disabled=True,key="c_shop")
        with cc2:
            st.text_input("Mobile",         value=cust.get("Mobile",""),         disabled=True,key="c_mob")
            st.text_input("Classification", value=cust.get("Classification",""), disabled=True,key="c_cls")
        with cc3:
            st.text_input("Sales executive",value=user["name"],disabled=True,key="c_se")
            st.text_input("SE UID",         value=user["uid"], disabled=True,key="c_seuid")
        st.markdown(sl("🛒 SKU / Product"), unsafe_allow_html=True)
        df_sku=active_skus()
        if df_sku.empty:
            st.warning("⚠️ No active SKUs. Ask admin to add SKUs first.")
        else:
            sc1,sc2,sc3=st.columns(3)
            with sc1:
                sku_display=[f"{r['SKU Name']}  ({r['SKU Code']})" for _,r in df_sku.iterrows()]
                sel_disp=st.selectbox("SKU Name *",sku_display,key="o_sku_disp")
                sel_idx=sku_display.index(sel_disp); sku_row=df_sku.iloc[sel_idx]
                sel_sku_code=sku_row["SKU Code"]; sel_sku_name=sku_row["SKU Name"]; sku_wt=str(sku_row["Weight Type"])
            with sc2:
                sku_price=float(str(sku_row["Price"]).replace("₹","").replace(",","") or 0)
                st.text_input("SKU Code",value=sel_sku_code,disabled=True,key="o_sku_code")
                st.text_input("Unit price ₹ (admin rate)",value=f"₹{sku_price:.2f}",disabled=True,key="o_price")
            with sc3:
                o_qty=st.number_input("Ordered qty *",min_value=0.0,step=0.5,key="o_qty")
                st.text_input("Weight type",value=sku_wt,disabled=True,key="o_wt")
            o_total=sku_price*o_qty
            if o_qty>0:
                st.markdown(
                    f'<div class="total-box">🧮 Order Total = {o_qty} × ₹{sku_price:.2f} = '
                    f'<span style="color:#1a7f4b;font-size:1.35rem">₹{o_total:,.2f}</span></div>',
                    unsafe_allow_html=True)
            else:
                st.info("Enter quantity — order total will appear here.")
            st.markdown(sl("📍 Shop Location (auto from customer record)"), unsafe_allow_html=True)
            auto_addr=str(cust.get("Shop Address","")).strip() if cust else ""
            auto_lat =str(cust.get("Latitude","")).strip()     if cust else ""
            auto_lng =str(cust.get("Longitude","")).strip()    if cust else ""
            o_addr=st.text_input("Shop address",value=auto_addr,key="o_addr")
            c_lat_col,c_lng_col=st.columns(2)
            with c_lat_col:
                o_lat=st.text_input("Latitude",value=auto_lat,key="o_lat",
                                     disabled=bool(auto_lat))
            with c_lng_col:
                o_lng=st.text_input("Longitude",value=auto_lng,key="o_lng",
                                     disabled=bool(auto_lng))
            if auto_lat and auto_lng:
                st.markdown(map_embed(auto_lat, auto_lng, 240), unsafe_allow_html=True)
                st.caption("📍 Customer's registered shop location.")
            st.divider()
            if st.button("✅ Submit Order",type="primary",use_container_width=True,key="o_submit"):
                if not cust:
                    st.error("Look up a customer first.")
                elif o_qty<=0:
                    st.error("Enter ordered quantity.")
                elif not o_addr:
                    st.error("Shop address required.")
                else:
                    soid="SO-"+o_id.replace("ORD-","")
                    ordered_time=datetime.now().strftime("%H:%M:%S")
                    append_row("base",[
                        o_id,soid,o_city,str(o_date),"",ordered_time,
                        cust.get("CUST-ID",""),cust.get("Shop Name",""),
                        cust.get("Mobile",""),cust.get("Classification",""),
                        user["name"],user["uid"],
                        sel_sku_code,sel_sku_name,sku_wt,
                        sku_price,o_qty,o_total,0,"","","","",
                        str(o_sopen),"",str(o_dcoff),
                        o_addr,auto_lat,auto_lng,
                        "Pending",user["uid"],
                        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    ])
                    st.success(f"✅ Order **{o_id}** submitted!  Total: **₹{o_total:,.2f}**")
                    st.session_state.cust_data={}; st.session_state.task_done=True; st.balloons()

    with tabs[1]:
        st.markdown(sl("👤 Customer Onboarding"), unsafe_allow_html=True)
        sc1,sc2=st.columns([3,1])
        with sc1: co_search=st.text_input("Search existing customer by mobile",key="co_search")
        with sc2:
            st.write(""); st.write("")
            do_cos=st.button("🔍 Search",key="co_search_btn")
        if do_cos and co_search:
            ex=find_row("customer_onboard","Mobile",co_search.strip())
            if ex:
                st.success(f"✅ Already onboarded — CUST-ID: **{ex['CUST-ID']}**")
                st.json({"Name":ex.get("Full Name"),"Shop":ex.get("Shop Name"),
                         "City":ex.get("City"),"Lat":ex.get("Latitude",""),
                         "Lng":ex.get("Longitude",""),"Status":ex.get("Status")})
            else:
                st.info("Not found — fill form below to onboard.")
        st.divider()
        st.markdown("#### New customer details")
        nc1,nc2=st.columns(2)
        with nc1:
            co_name =st.text_input("Full name *",           key="co_name")
            co_mob  =st.text_input("Mobile * (unique key)", placeholder="10-digit",key="co_mob")
            co_email=st.text_input("Email",                 key="co_email")
            co_shop =st.text_input("Shop name *",           key="co_shop")
        with nc2:
            co_city=st.selectbox("City *",
                ["Bengaluru","Mysuru","Hubli","Mangaluru","Hassan","Tumkur"],key="co_city")
            co_cls =st.selectbox("Classification",
                ["Restaurants","PG","Pubs","Premium Hotels","Wholesale","Retail","Others"],key="co_cls")
            co_addr=st.text_input("Shop address *",
                placeholder="e.g. 12/3 MG Road, Bengaluru",key="co_addr")
        st.markdown(sl("📍 Shop GPS Location"), unsafe_allow_html=True)
        st.markdown("**Tap the button below to capture your current location:**")
        captured_lat, captured_lng = gps_capture_component()
        if captured_lat and captured_lng:
            st.markdown(
                f'<div class="gps-box">'
                f'✅ &nbsp;Location captured &nbsp;|&nbsp; '
                f'<span style="color:#185fa5">Lat: {captured_lat}</span>'
                f' &nbsp;|&nbsp; '
                f'<span style="color:#185fa5">Lng: {captured_lng}</span>'
                f'</div>', unsafe_allow_html=True)
            st.markdown(map_embed(captured_lat, captured_lng, 260), unsafe_allow_html=True)
            st.caption(f"📍 Verify the pin is on the correct shop location. GPS: {captured_lat}, {captured_lng}")
        else:
            st.info("📍 Press the button above to capture your location.")
        st.divider()
        if st.button("✅ Onboard Customer", type="primary", use_container_width=True, key="co_btn"):
            if not all([co_name, co_mob, co_shop, co_addr]):
                st.error("Fill all required (*) fields.")
            elif not captured_lat or not captured_lng:
                st.error("📍 Location required — press 'Get My Current Location' button above first.")
            else:
                with st.spinner("Checking for duplicates…"):
                    ex = find_row("customer_onboard", "Mobile", co_mob.strip())
                if ex:
                    st.warning(f"⚠️ Mobile already registered — CUST-ID: **{ex['CUST-ID']}**")
                else:
                    cid = gen_cust_id()
                    append_row("customer_onboard", [
                        cid, co_name, co_mob, co_email, co_shop,
                        co_addr, co_city, co_cls,
                        captured_lat, captured_lng,
                        user["uid"], str(date.today()), "Active",
                    ])
                    load_customers.clear()
                    st.session_state["_geo_lat"] = ""
                    st.session_state["_geo_lng"] = ""
                    st.success(f"✅ Customer onboarded! CUST-ID: **`{cid}`** | GPS: {captured_lat}, {captured_lng}")
                    st.session_state.task_done = True
                    st.balloons()

    with tabs[2]:
        st.markdown(sl("📋 My Orders"), unsafe_allow_html=True)
        df_o=read_sheet("base")
        if df_o.empty:
            st.info("No orders yet.")
        else:
            my=df_o[df_o["sales executive Number"].astype(str)==user["uid"]]
            if my.empty:
                st.info("You haven't submitted any orders yet.")
            else:
                today_str=str(date.today())
                show_all_my=st.checkbox("Show all my orders (default: today only)",key="se_show_all")
                if not show_all_my and "ORDER DATE" in my.columns:
                    my_show=my[my["ORDER DATE"].astype(str)==today_str]
                    st.caption(f"📅 Today ({today_str}): {len(my_show)} order(s)")
                else:
                    my_show=my; st.caption(f"All orders: {len(my_show)}")
                tot_val=my_show["OrderTotal"].apply(
                    lambda x:float(str(x).replace("₹","").replace(",","") or 0)).sum()
                mc1,mc2,mc3,mc4=st.columns(4)
                mc1.metric("Shown",len(my_show))
                mc2.metric("Pending",len(my_show[my_show["Delivery Status"]=="Pending"]) if not my_show.empty else 0)
                mc3.metric("Delivered",len(my_show[my_show["Delivery Status"]=="Delivered"]) if not my_show.empty else 0)
                mc4.metric("Total value",f"₹{tot_val:,.0f}")
                if my_show.empty:
                    st.info(f"No orders for today. Check the box above to see all orders.")
                else:
                    show_cols=[c for c in ["Order ID","Customer shop name","SKU","SKU Name",
                                           "OrderedQty","OrderTotal","ORDER DATE","ORDERED TIME",
                                           "Delivery Status"] if c in my_show.columns]
                    st.dataframe(my_show[show_cols].sort_values("ORDER DATE",ascending=False),
                                 use_container_width=True,hide_index=True)

    st.divider()
    if not st.session_state.get("task_done",False):
        if st.button("✅ Mark All Tasks Done (enables Logout)",key="sales_tasks_done"):
            st.session_state.task_done=True; st.rerun()
    else:
        st.success("✅ Tasks marked complete — Logout button is now active above.")


# ═══════════════════════════════════════════════════════════════════════════════
#  PAGE: DELIVERY DRIVER (T2)
# ═══════════════════════════════════════════════════════════════════════════════
def page_delivery():
    user=st.session_state.user
    topbar("🚚 Delivery Driver · T2","#854f0b")
    tabs=st.tabs(["🗺️ My Route","📦 History"])

    with tabs[0]:
        st.markdown(sl("🗺️ Route & Deliveries","amber"),unsafe_allow_html=True)
        is_active=st.session_state.get("driver_active",True)
        drv_id   =st.session_state.get("driver_id","")
        tog_lbl  ="🟢 Active — admin can see you" if is_active else "⚫ Offline — tap to go active"
        if st.button(tog_lbl,key="active_tog"):
            is_active=not is_active; st.session_state.driver_active=is_active
            if drv_id: set_driver_status(drv_id,"Active" if is_active else "Offline")
            st.rerun()
        if not is_active:
            st.warning("You are offline. Tap the button above to go active."); return
        trip=get_driver_trip(user["uid"])
        if not trip:
            st.info("📋 No trip assigned. Contact admin."); return
        shop_ids=[s.strip() for s in str(trip.get("Shops","")).split(",") if s.strip()]
        st.info(f"**Trip:** {trip['Trip ID']} · **{trip['City']}** · **{trip['Date']}** · {len(shop_ids)} stop(s)")
        df_orders=read_sheet("base"); trip_ord={}
        if not df_orders.empty:
            for _,r in df_orders[df_orders["Tripid"].astype(str)==str(trip["Trip ID"])].iterrows():
                trip_ord[str(r["CustomerId"])]=r.to_dict()
        if str(trip.get("Status","")).lower()=="assigned":
            st.warning("Trip not started yet.")
            if st.button("▶️ Start Trip",type="primary",key="start_trip"):
                update_row("trips","Trip ID",trip["Trip ID"],{"Status":"In Progress"}); st.rerun()
            return
        if "active_stop" not in st.session_state: st.session_state.active_stop=0
        done_count=sum(1 for sid in shop_ids
                       if trip_ord.get(sid) and
                          str(trip_ord[sid].get("Delivery Status","")) not in ("Pending",""))
        if done_count>st.session_state.active_stop: st.session_state.active_stop=done_count
        active_idx=st.session_state.active_stop
        if len(shop_ids)>0:
            st.progress(active_idx/len(shop_ids),
                        text=f"Progress: {active_idx}/{len(shop_ids)} stops completed")
        for i,sid in enumerate(shop_ids):
            shop =find_row("customer_onboard","CUST-ID",sid)
            order=trip_ord.get(sid)
            is_done=bool(order and str(order.get("Delivery Status","")) not in ("Pending",""))
            is_cur =(i==active_idx) and not is_done
            is_lock=i>active_idx
            icon="✅" if is_done else ("📍" if is_cur else "🔒")
            stat=order.get("Delivery Status","Pending") if order else "No order"
            p_cls="pill-done" if is_done else ("pill-pend" if is_cur else "pill-off")
            with st.container():
                r1,r2=st.columns([9,2])
                with r1:
                    st.markdown(f"**{icon} Stop {i+1}** — **{shop.get('Shop Name','') if shop else sid}**")
                    if shop:
                        lat=str(shop.get("Latitude","")); lng=str(shop.get("Longitude",""))
                        st.caption(f"📍 {shop.get('Shop Address','')}  |  GPS: {lat}, {lng}")
                    if order:
                        st.caption(f"SKU: {order.get('SKU Name',order.get('SKU',''))} · "
                                   f"Qty: {order.get('OrderedQty','')} · ₹{order.get('OrderTotal','')}")
                with r2:
                    st.markdown(pill(stat,p_cls),unsafe_allow_html=True)
            if is_cur and not is_lock:
                if shop:
                    lat=str(shop.get("Latitude","")); lng=str(shop.get("Longitude",""))
                    if lat.strip() and lng.strip():
                        st.markdown(map_embed(lat, lng, 200), unsafe_allow_html=True)
                with st.form(key=f"del_form_{i}"):
                    st.markdown("##### ✍️ Update delivery")
                    df1,df2,df3=st.columns(3)
                    with df1:
                        d_reach  =st.time_input("Reach time *",value=datetime.now().time())
                        d_ddate  =st.date_input("Delivered date",value=date.today())
                    with df2:
                        d_status=st.selectbox("Status *",["Delivered","Partial","Failed","Rescheduled"])
                    with df3:
                        d_rqty  =st.number_input("Return qty",min_value=0.0,step=0.5)
                        d_rreason=st.text_input("Return reason")
                    d_notes=st.text_input("Notes (optional)")
                    submitted=st.form_submit_button("✅ Submit & Unlock Next Stop",
                                                    type="primary",use_container_width=True)
                if submitted:
                    if order:
                        update_row("base","Order ID",order["Order ID"],{
                            "Delivery Status":    d_status,
                            "ShopReachTime":      str(d_reach),
                            "DELIVERED DATE":     str(d_ddate),
                            "ReturnQty":          d_rqty,
                            "Reason":             d_rreason,
                            "return_updated_role":"delivery Driver",
                        })
                    st.session_state.active_stop=i+1; st.session_state.task_done=True
                    st.success(f"✅ Stop {i+1} marked **{d_status}**. "
                               f"{'Next stop unlocked! 🔓' if i+1<len(shop_ids) else '🎉 All done!'}")
                    st.rerun()
            st.divider()
        if active_idx>=len(shop_ids) and len(shop_ids)>0:
            st.success("🎉 All stops completed! Trip finished.")
            update_row("trips","Trip ID",trip["Trip ID"],{"Status":"Completed"})
            st.session_state.task_done=True

    with tabs[1]:
        st.markdown(sl("📦 My Delivery History","amber"),unsafe_allow_html=True)
        df_h=read_sheet("base")
        if df_h.empty:
            st.info("No records.")
        else:
            my_h=df_h[df_h["Delivery Status"]!="Pending"]
            if "return_updated_role" in my_h.columns:
                my_h=my_h[my_h["return_updated_role"].astype(str)=="delivery Driver"]
            today_str=str(date.today())
            show_all_h=st.checkbox("Show all history (default: today only)",key="dd_show_all")
            if not show_all_h and "DELIVERED DATE" in my_h.columns:
                my_h_show=my_h[my_h["DELIVERED DATE"].astype(str)==today_str]
            else:
                my_h_show=my_h
            if my_h_show.empty:
                st.info("No completed deliveries for today. Check the box to see all history.")
            else:
                show_cols=[c for c in ["Order ID","Customer shop name","SKU","SKU Name",
                           "OrderedQty","OrderTotal","Delivery Status",
                           "DELIVERED DATE","ShopReachTime","ReturnQty","Reason"]
                          if c in my_h_show.columns]
                st.dataframe(my_h_show[show_cols].sort_values("DELIVERED DATE",ascending=False),
                             use_container_width=True,hide_index=True)

    st.divider()
    if not st.session_state.get("task_done",False):
        if st.button("✅ Mark All Tasks Done (enables Logout)",key="driver_tasks_done"):
            st.session_state.task_done=True; st.rerun()
    else:
        st.success("✅ Tasks marked complete — Logout button is now active above.")


# ═══════════════════════════════════════════════════════════════════════════════
#  MAIN ROUTER
# ═══════════════════════════════════════════════════════════════════════════════
creds_path=os.path.join(os.path.dirname(os.path.abspath(__file__)),"credentials.json")
if not os.path.exists(creds_path):
    try:
        get_gspread_client()
    except Exception as e:
        st.error(f"❌ Cannot connect to Google Sheets: {e}")
        st.info("Add your Google credentials to Streamlit Cloud **Secrets** under `[gcp_service_account]`.")
        if st.button("🔄 Retry connection"):
            _cached_client.clear(); st.rerun()
        st.stop()

if not st.session_state.logged_in:
    page_login()
else:
    role=st.session_state.user["role"]
    if   role=="admin":           page_admin()
    elif role=="sales executive": page_sales()
    elif role=="delivery Driver": page_delivery()
    else:
        st.error(f"Unknown role: '{role}'")
        if st.button("Logout"):
            for k in DEFAULTS: st.session_state[k]=DEFAULTS[k]
            st.rerun()
