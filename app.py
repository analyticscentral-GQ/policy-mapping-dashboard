import os
import hmac
import streamlit as st
import pandas as pd
import requests
import io
from dotenv import load_dotenv
from openpyxl.styles import PatternFill, Font, Alignment

load_dotenv()

# ─── CONFIG ───────────────────────────────────────────────────────────────────
METABASE_URL = os.environ.get("METABASE_URL", "https://metabase.grayquest.com")
API_KEY = os.environ["METABASE_API_KEY"]
DATABASE_ID = int(os.environ.get("METABASE_DATABASE_ID", "2"))
APP_PASSWORD = os.environ["APP_PASSWORD"]

st.set_page_config(
    page_title="Policy Mapping Dashboard",
    page_icon="🏦",
    layout="wide",
)

st.markdown("""
<style>
    .stTabs [data-baseweb="tab-list"] { gap: 6px; }
    .stTabs [data-baseweb="tab"] {
        height: 48px; padding: 0 22px;
        background-color: #eef2f7;
        border-radius: 8px 8px 0 0;
        font-weight: 600; font-size: 15px;
    }
    .stTabs [aria-selected="true"] {
        background-color: #1a56db; color: white !important;
    }
    div[data-testid="metric-container"] {
        background: #f8fafc; border: 1px solid #e2e8f0;
        border-radius: 10px; padding: 10px;
    }
    .step-label {
        font-size: 16px; font-weight: 700;
        color: #1a56db; margin-bottom: 6px;
    }
</style>
""", unsafe_allow_html=True)


# ─── AUTH GATE ─────────────────────────────────────────────────────────────────
def check_password() -> bool:
    """Simple shared-password gate. Returns True once authenticated."""
    if st.session_state.get("authenticated"):
        return True

    st.title("🔒 Policy Mapping Dashboard")
    with st.form("login_form"):
        pwd = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Login")
    if submitted:
        if hmac.compare_digest(pwd, APP_PASSWORD):
            st.session_state["authenticated"] = True
            st.rerun()
        else:
            st.error("❌ Incorrect password.")
    return False


if not check_password():
    st.stop()


# ─── METABASE HELPER ──────────────────────────────────────────────────────────
def run_query(sql: str):
    """Execute SQL via Metabase API. Returns (DataFrame, error_string)."""
    url = f"{METABASE_URL}/api/dataset"
    headers = {"x-api-key": API_KEY, "Content-Type": "application/json"}
    payload = {"database": DATABASE_ID, "type": "native", "native": {"query": sql}}
    try:
        resp = requests.post(url, json=payload, headers=headers, timeout=90)
        resp.raise_for_status()
        data = resp.json()
        if "data" in data and "cols" in data["data"]:
            cols = [c["name"] for c in data["data"]["cols"]]
            return pd.DataFrame(data["data"]["rows"], columns=cols), None
        if "error" in data:
            return pd.DataFrame(), data["error"]
        return pd.DataFrame(), "Unexpected response from Metabase."
    except requests.exceptions.Timeout:
        return pd.DataFrame(), "Query timed out (>90s). Try with fewer IDs."
    except Exception as exc:
        return pd.DataFrame(), str(exc)


# ─── EXCEL BUILDER ────────────────────────────────────────────────────────────
def to_excel(df: pd.DataFrame) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Output")
        ws = writer.sheets["Output"]

        # Header style
        hdr_fill = PatternFill(start_color="1A56DB", end_color="1A56DB", fill_type="solid")
        for cell in ws[1]:
            cell.fill = hdr_fill
            cell.font = Font(color="FFFFFF", bold=True)
            cell.alignment = Alignment(horizontal="center")

        # Colour-code action column
        action_col_idx = None
        for idx, cell in enumerate(ws[1], 1):
            if str(cell.value).lower() == "action":
                action_col_idx = idx
                break
        if action_col_idx:
            ins_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
            del_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            for row in ws.iter_rows(min_row=2, min_col=action_col_idx, max_col=action_col_idx):
                for cell in row:
                    if cell.value == "INSERT":
                        cell.fill = ins_fill
                    elif cell.value == "DELETE":
                        cell.fill = del_fill

        # Auto-width
        for col in ws.columns:
            max_len = max((len(str(c.value)) if c.value else 0) for c in col)
            ws.column_dimensions[col[0].column_letter].width = max(12, min(max_len + 4, 40))

    output.seek(0)
    return output.read()


# ─── FILE READER ──────────────────────────────────────────────────────────────
def read_institute_ids(uploaded_file) -> list:
    try:
        if uploaded_file.name.lower().endswith(".csv"):
            df = pd.read_csv(uploaded_file)
        else:
            df = pd.read_excel(uploaded_file)
        col_map = {c.lower().strip(): c for c in df.columns}
        if "institute_id" in col_map:
            ids = df[col_map["institute_id"]].dropna().astype(int).tolist()
            return ids
        st.error("❌ Column 'institute_id' not found in uploaded file.")
        return []
    except Exception as e:
        st.error(f"❌ Could not read file: {e}")
        return []


def fmt_ids(ids: list) -> str:
    return ", ".join(str(i) for i in ids)


# ─── PAIR INPUT COMPONENT ─────────────────────────────────────────────────────
def pair_input_ui(prefix: str) -> list:
    """Reusable lender-policy pair selector with INSERT/DELETE toggle."""
    key = f"{prefix}_pairs"
    if key not in st.session_state:
        st.session_state[key] = []

    st.markdown('<p class="step-label">➕ Add Lender-Policy Pair</p>', unsafe_allow_html=True)
    c1, c2, c3, c4 = st.columns([2, 2, 2, 1])
    lid = c1.number_input("Lender ID", min_value=1, step=1, key=f"{prefix}_lid")
    pid = c2.number_input("Policy ID", min_value=1, step=1, key=f"{prefix}_pid")
    act = c3.selectbox("Action", ["INSERT", "DELETE"], key=f"{prefix}_act")
    c4.markdown("<br><br>", unsafe_allow_html=True)
    add_clicked = c4.button("Add ➕", key=f"{prefix}_add")

    if add_clicked:
        dup = any(
            p["lender_id"] == int(lid) and p["policy_id"] == int(pid)
            for p in st.session_state[key]
        )
        if dup:
            st.warning("⚠️ This lender-policy pair already exists. Remove it first to change action.")
        else:
            st.session_state[key].append(
                {"lender_id": int(lid), "policy_id": int(pid), "action": act}
            )

    pairs = st.session_state[key]
    if pairs:
        st.markdown("**Current pairs:**")
        h0, h1, h2, h3 = st.columns([2, 2, 3, 1])
        h0.markdown("**Lender ID**"); h1.markdown("**Policy ID**")
        h2.markdown("**Action**");    h3.markdown("**Del**")
        to_remove = None
        for i, p in enumerate(pairs):
            r0, r1, r2, r3 = st.columns([2, 2, 3, 1])
            r0.write(p["lender_id"]); r1.write(p["policy_id"])
            badge = "🟢 INSERT" if p["action"] == "INSERT" else "🔴 DELETE"
            r2.write(badge)
            if r3.button("🗑", key=f"{prefix}_del_{i}"):
                to_remove = i
        if to_remove is not None:
            st.session_state[key].pop(to_remove)
            st.rerun()
        if st.button("🧹 Clear All", key=f"{prefix}_clr"):
            st.session_state[key] = []
            st.rerun()

    return pairs


# ─── SUBTABLE 4 COMPONENT ─────────────────────────────────────────────────────
def subtable4_ui(prefix: str):
    """Available Lender-Policy reference table with live search."""
    with st.expander("📋 Sub-table 4 — Available Lender-Policy in System (Reference)", expanded=False):
        srch = st.text_input(
            "🔍 Search — type anything (e.g. 'bureau', 'banking', lender name, policy name…)",
            key=f"{prefix}_st4_srch",
        )
        load = st.button("🔄 Load / Refresh", key=f"{prefix}_st4_load")

        SQL_ST4 = """
SELECT gilp.lender_id,
       gilp.id          AS policy_id,
       gilm.name        AS lender_name,
       gilp.name        AS policy_name,
       gilp.`type`
FROM   gq_internal_lender_policies gilp
LEFT JOIN gq_internal_lender_master gilm ON gilm.id = gilp.lender_id
WHERE  (gilp.`type` LIKE '%BANKING%' OR gilp.`type` LIKE '%BUREAU%')
  AND  gilp.is_active  = 1
  AND  gilp.deleted_on IS NULL
ORDER BY gilp.lender_id, gilp.id
"""
        if load:
            with st.spinner("Fetching policy list…"):
                df, err = run_query(SQL_ST4)
            if err:
                st.error(f"❌ {err}")
            else:
                st.session_state[f"{prefix}_st4_df"] = df

        if f"{prefix}_st4_df" in st.session_state:
            df = st.session_state[f"{prefix}_st4_df"].copy()
            if srch.strip():
                mask = df.apply(
                    lambda row: row.astype(str).str.contains(srch.strip(), case=False).any(), axis=1
                )
                df = df[mask]
            st.dataframe(df, use_container_width=True, height=280)
            st.caption(f"Showing {len(df)} of {len(st.session_state[f'{prefix}_st4_df'])} policies")


# ─── OUTPUT BUILDER ───────────────────────────────────────────────────────────
def build_output(institute_ids: list, pairs: list) -> pd.DataFrame:
    rows = [
        {"institute_id": iid, "policy_id": p["policy_id"],
         "action": p["action"], "lender_id": p["lender_id"]}
        for iid in institute_ids
        for p in pairs
    ]
    return pd.DataFrame(rows, columns=["institute_id", "policy_id", "action", "lender_id"])


# ══════════════════════════════════════════════════════════════════════════════
# PAGE HEADER
# ══════════════════════════════════════════════════════════════════════════════
st.title("🏦 Policy Mapping Dashboard")
st.caption("Credit Policy Team · Loan Product Management · GrayQuest")
st.divider()

tab1, tab2, tab3 = st.tabs([
    "🏦  New Institute Onboarding",
    "🔄  New Product – Existing Institute",
    "✏️  Policy Edit – Active Products",
])


# ══════════════════════════════════════════════════════════════════════════════
# TAB 1 — New Institute Onboarding
# ══════════════════════════════════════════════════════════════════════════════
with tab1:
    st.header("🏦 New Institute Onboarding")
    st.info(
        "**Scenario:** First-time institute onboarding. "
        "Upload product file → add lender-policy pairs → download INSERT-ready mapping Excel."
    )

    left, right = st.columns([1, 1], gap="large")

    with left:
        st.markdown('<p class="step-label">📁 Step 1 — Upload Product File</p>', unsafe_allow_html=True)
        t1_file = st.file_uploader(
            "Excel / CSV with column **institute_id**",
            type=["xlsx", "csv"], key="t1_file"
        )
        t1_ids = []
        if t1_file:
            t1_ids = read_institute_ids(t1_file)
            if t1_ids:
                st.success(f"✅ {len(t1_ids)} product(s) loaded")
                st.dataframe(pd.DataFrame({"institute_id": t1_ids}), height=200, use_container_width=True)

    with right:
        st.markdown('<p class="step-label">🔗 Step 2 — Define Lender-Policy Pairs</p>', unsafe_allow_html=True)
        t1_pairs = pair_input_ui("t1")

    st.divider()
    subtable4_ui("t1")
    st.divider()

    st.markdown('<p class="step-label">📥 Step 3 — Generate & Download Output</p>', unsafe_allow_html=True)
    if t1_ids and t1_pairs:
        out = build_output(t1_ids, t1_pairs)
        st.dataframe(out, use_container_width=True)
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Rows", len(out))
        m2.metric("🟢 INSERT", len(out[out.action == "INSERT"]))
        m3.metric("🔴 DELETE", len(out[out.action == "DELETE"]))
        st.download_button(
            "⬇️  Download Mapping Excel",
            data=to_excel(out),
            file_name="new_institute_policy_mapping.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        missing = (["product file"] if not t1_ids else []) + (["lender-policy pairs"] if not t1_pairs else [])
        st.warning(f"⚠️ Please provide: {' and '.join(missing)}")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 2 — New Product – Existing Institute
# ══════════════════════════════════════════════════════════════════════════════
with tab2:
    st.header("🔄 New Product – Existing Institute")
    st.info(
        "**Scenario:** Institute already exists, new product created. "
        "Upload product IDs → system finds reference product from same group/inst/location → "
        "copies bureau policies → download INSERT-ready Excel."
    )

    st.markdown('<p class="step-label">📁 Step 1 — Upload Product File</p>', unsafe_allow_html=True)
    t2_file = st.file_uploader(
        "Excel / CSV with column **institute_id**",
        type=["xlsx", "csv"], key="t2_file"
    )
    t2_ids = []
    if t2_file:
        t2_ids = read_institute_ids(t2_file)
        if t2_ids:
            st.success(f"✅ {len(t2_ids)} product(s) loaded: {t2_ids}")

    st.divider()
    st.markdown('<p class="step-label">🔍 Step 2 — Fetch Reference Policy Mappings</p>', unsafe_allow_html=True)

    if t2_ids:
        if st.button("🚀 Run Query", key="t2_run"):
            SQL_T2 = f"""
WITH cte AS (
    SELECT i.*, i2.institute_id AS required_product
    FROM   institutes i
    INNER JOIN institutes i2
           ON  i.group_id    = i2.group_id
           AND i.inst_id     = i2.inst_id
           AND i.location_id = i2.location_id
           AND i.board_id    = i2.board_id
    WHERE  i2.institute_id IN ({fmt_ids(t2_ids)})
),
cte1 AS (
    SELECT cte.institute_id
    FROM   cte
    LEFT JOIN gq_old_product_lender_policy_mapping goplpm ON cte.institute_id = goplpm.product_id
    LEFT JOIN gq_internal_lender_policies          gilp   ON gilp.id          = goplpm.policy_id
    WHERE  goplpm.is_active  = 1
      AND  (gilp.`type` LIKE '%BUREAU%' OR gilp.`type` LIKE '%BANKING%')
      AND  goplpm.deleted_on IS NULL
    GROUP  BY goplpm.product_id
    HAVING COUNT(DISTINCT goplpm.lender_id) > 2
),
cte2 AS (
    SELECT cte.*,
           ROW_NUMBER() OVER (
               PARTITION BY cte.group_id, cte.inst_id, cte.location_id
               ORDER BY cte.institute_id DESC
           ) AS rnk
    FROM   cte
    WHERE  cte.institute_id IN (SELECT institute_id FROM cte1)
)
SELECT
    cte2.institute_id   AS reference_id,
    cte2.required_product,
    cte2.status,
    goplpm.lender_id,
    goplpm.policy_id,
    gilp.`type`,
    gilp.name
FROM   cte2
LEFT JOIN gq_old_product_lender_policy_mapping goplpm ON cte2.institute_id = goplpm.product_id
LEFT JOIN gq_internal_lender_policies          gilp   ON gilp.id           = goplpm.policy_id
WHERE  cte2.rnk          = 1
  AND  (gilp.`type` LIKE '%BUREAU%' OR gilp.`type` LIKE '%BANKING%')
  AND  goplpm.is_active  = 1
  AND  goplpm.deleted_on IS NULL
"""
            with st.spinner("Running query on Metabase…"):
                df, err = run_query(SQL_T2)
            if err:
                st.error(f"❌ Query error: {err}")
            else:
                st.session_state["t2_raw"] = df
                st.success(f"✅ {len(df)} rows fetched")

    if "t2_raw" in st.session_state:
        raw = st.session_state["t2_raw"]
        with st.expander("🔍 Raw Query Result (all columns)", expanded=False):
            st.dataframe(raw, use_container_width=True, height=300)

        st.divider()
        st.markdown('<p class="step-label">📥 Step 3 — Generate & Download Output</p>', unsafe_allow_html=True)

        if not raw.empty:
            out = raw[["required_product", "policy_id", "lender_id"]].copy()
            out.rename(columns={"required_product": "institute_id"}, inplace=True)
            out["action"] = "INSERT"
            out = out[["institute_id", "policy_id", "action", "lender_id"]].drop_duplicates()

            st.dataframe(out, use_container_width=True)
            st.metric("Total Rows in Output", len(out))
            st.download_button(
                "⬇️  Download Mapping Excel",
                data=to_excel(out),
                file_name="existing_institute_new_product_mapping.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.warning("⚠️ Query returned no results. Check if these product IDs have qualifying reference products.")
    elif not t2_ids:
        st.warning("⚠️ Upload a product file to continue.")


# ══════════════════════════════════════════════════════════════════════════════
# TAB 3 — Policy Edit – Active Products
# ══════════════════════════════════════════════════════════════════════════════
with tab3:
    st.header("✏️ Policy Edit – Active Products")
    st.info(
        "**Scenario:** Edit policies on existing active products. "
        "Use the lookup tables to identify current mappings and relevant IDs, "
        "then build your INSERT / DELETE output Excel below."
    )

    # ── Sub-table 1 ──────────────────────────────────────────────────────────
    with st.expander("📊 Sub-table 1 — Active Policies for Uploaded Products", expanded=True):
        st.markdown("Upload a product file to see all currently active BUREAU / BANKING policies.")
        t3_f1 = st.file_uploader(
            "Excel / CSV with column **institute_id**",
            type=["xlsx", "csv"], key="t3_f1"
        )
        t3_st1_manual = st.text_input(
            "Or enter Institute ID(s) directly (comma-separated)",
            placeholder="e.g. 299612, 311785",
            key="t3_st1_manual"
        )

        t3_ids_st1 = []
        if t3_f1:
            t3_ids_st1 = read_institute_ids(t3_f1)
            if t3_ids_st1:
                st.success(f"✅ {len(t3_ids_st1)} product(s) loaded from file")
        if t3_st1_manual.strip():
            try:
                manual_ids = [int(x.strip()) for x in t3_st1_manual.split(",") if x.strip()]
                # Merge with file ids, deduplicate
                t3_ids_st1 = list(dict.fromkeys(t3_ids_st1 + manual_ids))
                st.info(f"ℹ️ Using {len(t3_ids_st1)} institute ID(s): {t3_ids_st1}")
            except ValueError:
                st.error("❌ Invalid input — please enter numeric IDs only, comma-separated.")

        if t3_ids_st1 and st.button("🔍 Fetch Active Policies", key="t3_st1_run"):
            SQL_ST1 = f"""
SELECT goplpm.product_id,
       gigm.group_name,
       giim.inst_name,
       goplpm.lender_id,
       goplpm.policy_id,
       gilp.name,
       gilp.`type`,
       gigm.group_id,
       giim.inst_id,
       goplpm.is_active,
       goplpm.updated_on,
       goplpm.created_on,
       i.institute_name,
       i.status,
       i.platform,
       goplpm.deleted_on
FROM   gq_old_product_lender_policy_mapping goplpm
LEFT JOIN gq_internal_lender_policies   gilp  ON gilp.id         = goplpm.policy_id
LEFT JOIN institutes                    i     ON i.institute_id  = goplpm.product_id
LEFT JOIN gq_internal_group_master      gigm  ON gigm.group_id   = i.group_id
LEFT JOIN gq_internal_institute_master  giim  ON i.inst_id       = giim.inst_id
WHERE  i.status            = 1
  AND  goplpm.product_id   IN ({fmt_ids(t3_ids_st1)})
  AND  (gilp.`type` LIKE '%BANKING%' OR gilp.`type` LIKE '%BUREAU%')
  AND  goplpm.is_active    = 1
  AND  goplpm.deleted_on   IS NULL
GROUP  BY 1, 4, 5
"""
            with st.spinner("Fetching…"):
                df, err = run_query(SQL_ST1)
            if err:
                st.error(f"❌ {err}")
            else:
                st.session_state["t3_st1"] = df
                st.success(f"✅ {len(df)} rows")

        if "t3_st1" in st.session_state:
            st.dataframe(st.session_state["t3_st1"], use_container_width=True, height=350)
            st.caption(f"{len(st.session_state['t3_st1'])} active policy-product combinations")

    # ── Sub-table 2 ──────────────────────────────────────────────────────────
    with st.expander("🏢 Sub-table 2 — Policies by Group / Institute / Location", expanded=False):
        st.markdown(
            "See all **distinct lender-policy combinations** already mapped at group, "
            "institute, or location level. Fill one or more filters."
        )
        fc1, fc2, fc3 = st.columns(3)
        t3_gids = fc1.text_input("Group ID(s)    (comma-separated)", placeholder="e.g. 10, 20",  key="t3_gids")
        t3_iids = fc2.text_input("Institute ID(s) (comma-separated)", placeholder="e.g. 5, 6",   key="t3_iids")
        t3_lids = fc3.text_input("Location ID(s)  (comma-separated)", placeholder="e.g. 1, 2",   key="t3_lids")

        if st.button("🔍 Fetch Group-Level Policies", key="t3_st2_run"):
            if not any([t3_gids.strip(), t3_iids.strip(), t3_lids.strip()]):
                st.warning("⚠️ Enter at least one filter value.")
            else:
                conds = [
                    "i.status = 1",
                    "(gilp.`type` LIKE '%BANKING%' OR gilp.`type` LIKE '%BUREAU%')",
                    "goplpm.is_active = 1",
                    "goplpm.deleted_on IS NULL",
                ]
                if t3_gids.strip():
                    conds.append(f"gigm.group_id IN ({t3_gids.strip()})")
                if t3_iids.strip():
                    conds.append(f"giim.inst_id IN ({t3_iids.strip()})")
                if t3_lids.strip():
                    conds.append(f"gilm.location_id IN ({t3_lids.strip()})")

                SQL_ST2 = f"""
SELECT goplpm.product_id,
       gigm.group_name,
       giim.inst_name,
       goplpm.lender_id,
       goplpm.policy_id,
       gilp.name,
       gilp.`type`,
       gigm.group_id,
       giim.inst_id,
       goplpm.is_active,
       goplpm.updated_on,
       goplpm.created_on,
       i.institute_name,
       i.status,
       i.platform,
       goplpm.deleted_on
FROM   gq_old_product_lender_policy_mapping goplpm
LEFT JOIN gq_internal_lender_policies   gilp  ON gilp.id         = goplpm.policy_id
LEFT JOIN institutes                    i     ON i.institute_id  = goplpm.product_id
LEFT JOIN gq_internal_group_master      gigm  ON gigm.group_id   = i.group_id
LEFT JOIN gq_internal_institute_master  giim  ON i.inst_id       = giim.inst_id
LEFT JOIN gq_internal_location_master   gilm  ON gilm.location_id = i.location_id
WHERE  {chr(10) + '  AND '.join(conds)}
GROUP  BY goplpm.lender_id, goplpm.policy_id
"""
                with st.spinner("Fetching…"):
                    df, err = run_query(SQL_ST2)
                if err:
                    st.error(f"❌ {err}")
                else:
                    st.session_state["t3_st2"] = df
                    st.success(f"✅ {len(df)} distinct lender-policy combinations")

        if "t3_st2" in st.session_state:
            st.dataframe(st.session_state["t3_st2"], use_container_width=True, height=320)

    # ── Sub-table 3 ──────────────────────────────────────────────────────────
    with st.expander("🔎 Sub-table 3 — Search Group / Institute / Location", expanded=False):
        st.markdown(
            "Search by **name** (partial match) to look up `group_id`, `inst_id`, `location_id`. "
            "Use those IDs in Sub-table 2 above."
        )
        st.info(
            "💡 Fill **only the field(s) you want to search by**. "
            "Each filled field uses partial (LIKE) match; if more than one field is filled, "
            "results must match all of them (AND)."
        )
        s1, s2, s3 = st.columns(3)
        t3_sg = s1.text_input("Group Name",    placeholder="e.g. Shri", key="t3_sg")
        t3_si = s2.text_input("Institute Name",placeholder="e.g. Avantika", key="t3_si")
        t3_sl = s3.text_input("Location Name", placeholder="e.g. worli", key="t3_sl")

        if st.button("🔍 Search", key="t3_st3_run"):
            if not any([t3_sg.strip(), t3_si.strip(), t3_sl.strip()]):
                st.warning("⚠️ Enter at least one search term.")
            else:
                base_conds = []
                # Each filled field is AND-ed in as its own LIKE condition
                if t3_si.strip():
                    base_conds.append(f"giim.inst_name LIKE '%{t3_si.strip()}%'")
                if t3_sg.strip():
                    base_conds.append(f"gigm.group_name LIKE '%{t3_sg.strip()}%'")
                if t3_sl.strip():
                    base_conds.append(f"gilm.location_name LIKE '%{t3_sl.strip()}%'")

                SQL_ST3 = f"""
SELECT  gigm.group_id,
        giim.inst_id,
        gilm.location_id,
        giim.inst_name, 
        gigm.group_name,
        gilm.location_name,
        i.platform           AS inst_type,
        giim.created_on,
        giim.updated_on,
        i.institute_name,
        i.institute_id,
        i.status,
        i.tags,
        i.retention_percent
FROM    institutes i
LEFT JOIN gq_internal_institute_master  giim ON i.inst_id       = giim.inst_id
LEFT JOIN gq_internal_group_master      gigm ON gigm.group_id   = i.group_id
LEFT JOIN gq_internal_location_master   gilm ON gilm.location_id = i.location_id
WHERE   {chr(10) + '  AND '.join(base_conds)}
GROUP   BY 1, 2, 3
"""
                with st.spinner("Searching…"):
                    df, err = run_query(SQL_ST3)
                if err:
                    st.error(f"❌ {err}")
                else:
                    st.session_state["t3_st3"] = df
                    st.success(f"✅ {len(df)} result(s) found")

        if "t3_st3" in st.session_state:
            st.dataframe(st.session_state["t3_st3"], use_container_width=True, height=300)

    # ── Sub-table 4 ──────────────────────────────────────────────────────────
    subtable4_ui("t3")

    # ── Sub-table 5 ──────────────────────────────────────────────────────────
    with st.expander("🗂️ Sub-table 5 — Active Products in System", expanded=False):
        st.markdown(
            "Search active products by **Group**, **Group + Institute**, "
            "**Group + Institute + Location**, **Institute Name**, or **Application ID (bi.id)**. "
            "Fill only the boxes you need — unused ones are ignored."
        )

        st.info("💡 Fill only the boxes you need — all filled filters are applied as AND conditions.")

        sa, sb, sc = st.columns(3)
        t3_s5_group   = sa.text_input("Group ID(s)",                        placeholder="e.g. 10, 20",         key="t3_s5_group")
        t3_s5_inst    = sb.text_input("Inst ID(s)  — institution-level ID", placeholder="e.g. 1786, 200",      key="t3_s5_inst")
        t3_s5_loc     = sc.text_input("Location ID(s)",                     placeholder="e.g. 1, 2",           key="t3_s5_loc")

        sb2_a, sb2_b, sb2_c = st.columns(3)
        t3_s5_prodid  = sb2_a.text_input("Product ID(s)  — institute_id",   placeholder="e.g. 316415, 299612", key="t3_s5_prodid")
        t3_s5_iname   = sb2_b.text_input("Institute Name  (partial match)", placeholder="e.g. Ryan",           key="t3_s5_iname")
        t3_s5_platform= sb2_c.text_input("Platform / Inst Type",            placeholder="e.g. school, edtech", key="t3_s5_platform")

        if st.button("🔍 Fetch Active Products", key="t3_st5_run"):
            g   = t3_s5_group.strip()
            ins = t3_s5_inst.strip()
            loc = t3_s5_loc.strip()
            pid = t3_s5_prodid.strip()
            nm  = t3_s5_iname.strip()
            plt = t3_s5_platform.strip()

            if not any([g, ins, loc, pid, nm, plt]):
                st.warning("⚠️ Please fill at least one filter.")
            else:
                conds = ["i.status = 1"]
                if g:
                    conds.append(f"i.group_id IN ({g})")
                if ins:
                    conds.append(f"i.inst_id IN ({ins})")
                if loc:
                    conds.append(f"i.location_id IN ({loc})")
                if pid:
                    conds.append(f"i.institute_id IN ({pid})")
                if nm:
                    conds.append(f"i.institute_name LIKE '%{nm}%'")
                if plt:
                    conds.append(f"i.platform LIKE '%{plt}%'")

                SQL_ST5 = f"""
SELECT  i.institute_id   AS product_id,
        i.inst_id        AS inst_id,
        i.group_id,
        i.location_id,
        i.board_id,
        i.status,
        i.platform       AS inst_type,
        i.institute_name AS product_name,
        gigm.group_name,
        giim.inst_name,
        i.tags,
        IF(i.tags LIKE '%split%', 1, 0) AS split_disbursement_tagk
FROM    institutes i
LEFT JOIN gq_internal_group_master     gigm ON gigm.group_id = i.group_id
LEFT JOIN gq_internal_institute_master giim ON giim.inst_id  = i.inst_id
WHERE   {chr(10) + '  AND '.join(conds)}
GROUP BY i.institute_id
"""
                with st.spinner("Fetching active products…"):
                    df, err = run_query(SQL_ST5)
                if err:
                    st.error(f"❌ {err}")
                else:
                    st.session_state["t3_st5"] = df
                    st.success(f"✅ {len(df)} active product(s) found")

        if "t3_st5" in st.session_state:
            df5 = st.session_state["t3_st5"]
            st.info(
                "📌 **product_id** = `institute_id` — the unique product/product row ID (used for policy mapping uploads)  \n"
                "📌 **inst_id** = institution-level ID — links to the institute master (group/inst/location)"
            )
            st.dataframe(df5, use_container_width=True, height=350)
            st.caption(f"{len(df5)} unique products found")

            # Export options — product_id (institute_id) only
            df5_export = df5[["product_id"]].rename(columns={"product_id": "institute_id"})
            ex1, ex2 = st.columns(2)
            ex1.download_button(
                "⬇️  Export as Excel",
                data=to_excel(df5_export),
                file_name="active_products.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="t3_st5_xl"
            )
            csv_buf = io.BytesIO(df5_export.to_csv(index=False).encode("utf-8"))
            ex2.download_button(
                "⬇️  Export as CSV",
                data=csv_buf.getvalue(),
                file_name="active_products.csv",
                mime="text/csv",
                key="t3_st5_csv"
            )

    st.divider()

    # ── Sub-table 6 ──────────────────────────────────────────────────────────
    with st.expander("🔁 Sub-table 6 — Application ID ↔ Institute ID Lookup", expanded=False):
        st.markdown(
            "Enter **Application ID(s)** to get their Institute IDs, "
            "or enter **Institute ID(s)** to get all their Application IDs. "
            "Both boxes together work as OR."
        )
        sb6_c1, sb6_c2 = st.columns(2)
        t3_s6_appid = sb6_c1.text_input("Application ID(s)  (bi.id)", placeholder="e.g. 1836859, 1234567", key="t3_s6_appid")
        t3_s6_instid = sb6_c2.text_input("Institute ID(s)", placeholder="e.g. 316415, 299612", key="t3_s6_instid")

        if st.button("🔍 Lookup", key="t3_st6_run"):
            app6 = t3_s6_appid.strip()
            inst6 = t3_s6_instid.strip()

            if not any([app6, inst6]):
                st.warning("⚠️ Please enter at least one Application ID or Institute ID.")
            else:
                or_conds = []
                if app6:
                    or_conds.append(f"bi.id IN ({app6})")
                if inst6:
                    or_conds.append(f"i.institute_id IN ({inst6})")

                SQL_ST6 = f"""
SELECT  bi.id          AS application_id,
        i.institute_id
FROM    basic_info bi
LEFT JOIN course_details cd ON cd.basic_info_id = bi.id
LEFT JOIN institutes     i  ON i.institute_id   = cd.institute_id
WHERE   {' OR '.join(or_conds)}
"""
                with st.spinner("Looking up…"):
                    df, err = run_query(SQL_ST6)
                if err:
                    st.error(f"❌ {err}")
                else:
                    st.session_state["t3_st6"] = df
                    st.success(f"✅ {len(df)} row(s) found")

        if "t3_st6" in st.session_state:
            df6 = st.session_state["t3_st6"]
            st.dataframe(df6, use_container_width=True, height=300)
            st.caption(f"{len(df6)} rows — {df6['application_id'].nunique()} unique applications, {df6['institute_id'].nunique()} unique institute IDs")

    st.divider()

    # ── Final Output Section ─────────────────────────────────────────────────
    st.markdown("## 🎯 Generate Policy Edit Output")
    st.info(
        "After reviewing the tables above, upload the product file and "
        "define which lender-policy pairs to INSERT or DELETE."
    )

    out_left, out_right = st.columns([1, 1], gap="large")

    with out_left:
        st.markdown('<p class="step-label">📁 Upload Product File</p>', unsafe_allow_html=True)
        t3_fout = st.file_uploader(
            "Excel / CSV with column **institute_id**",
            type=["xlsx", "csv"], key="t3_fout"
        )
        t3_out_ids = []
        if t3_fout:
            t3_out_ids = read_institute_ids(t3_fout)
            if t3_out_ids:
                st.success(f"✅ {len(t3_out_ids)} product(s) loaded")
                st.dataframe(pd.DataFrame({"institute_id": t3_out_ids}), height=200, use_container_width=True)

    with out_right:
        st.markdown('<p class="step-label">🔗 Define Lender-Policy Pairs</p>', unsafe_allow_html=True)
        t3_pairs = pair_input_ui("t3_out")

    st.divider()

    if t3_out_ids and t3_pairs:
        out = build_output(t3_out_ids, t3_pairs)
        st.markdown("#### 📊 Output Preview")
        st.dataframe(out, use_container_width=True)
        m1, m2, m3 = st.columns(3)
        m1.metric("Total Rows", len(out))
        m2.metric("🟢 INSERT", len(out[out.action == "INSERT"]))
        m3.metric("🔴 DELETE", len(out[out.action == "DELETE"]))
        st.download_button(
            "⬇️  Download Policy Edit Excel",
            data=to_excel(out),
            file_name="policy_edit_active_products.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    else:
        missing = (["product file"] if not t3_out_ids else []) + \
                  (["lender-policy pairs"] if not t3_pairs else [])
        if missing:
            st.warning(f"⚠️ To generate output, please provide: {' and '.join(missing)}")
