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


def generate_route_template_excel():
    """Generate a downloadable Excel template for route uploads."""
    buf = io.BytesIO()
    df  = pd.DataFrame([
        {
            "SaleOrderId": "SO-10001",
            "DeliveryDate": "2025-06-01",
            "CustomerId": "CUST-AABBCC",
            "Customer": "Sri Lakshmi Provision",
            "Slot": "Morning",
            "Driver": "Ravi Kumar",
            "TotalCrates": 4,
            "OrderKg": 12.5,
            "Latitude": 12.971599,
            "Longitude": 77.594566,
            "FC Latitude": 12.9352,
            "FC Longitude": 77.6245,
        },
        {
            "SaleOrderId": "SO-10002",
            "DeliveryDate": "2025-06-01",
            "CustomerId": "CUST-DDEEFF",
            "Customer": "Hotel Majestic",
            "Slot": "Afternoon",
            "Driver": "Ravi Kumar",
            "TotalCrates": 2,
            "OrderKg": 6.0,
            "Latitude": 12.295810,
            "Longitude": 76.639380,
            "FC Latitude": 12.9352,
            "FC Longitude": 77.6245,
        },
    ])
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Route")
    buf.seek(0)
    return buf.read()

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
#  PAGE: ADMIN
# ═══════════════════════════════════════════════════════════════════════════════
def page_admin():
    user = st.session_state.user
    topbar("🛡️ Admin","#185fa5")
    tabs = st.tabs(["📦 SKUs","🗺️ Trips","🚚 Assign Drivers",
                    "👤 Customers","🚗 Driver Onboard","📋 Orders","📝 Audit Log"])

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

            dl_col1, dl_col2 = st.columns(2)
            with dl_col1:
                st.download_button(
                    "⬇️ Download Excel Template",
                    data=generate_route_template_excel(),
                    file_name="route_template.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key="dl_tpl_xlsx",
                )
            with dl_col2:
                st.download_button(
                    "⬇️ Download TSV Template",
                    data=ROUTE_TEMPLATE_CSV,
                    file_name="route_template.tsv",
                    mime="text/tab-separated-values",
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
