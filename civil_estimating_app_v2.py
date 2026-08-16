import streamlit as st
import pandas as pd
from pathlib import Path

st.set_page_config(page_title="Civil Estimating Software", page_icon="🏗️")
st.title("🏗️ Civil Estimating Software")
st.caption("Version 2 — BNI Productivity Database")

DB = Path(__file__).parent / "construction_costs_pages_16_to_36.xlsx"

@st.cache_data
def load_db(source):
    df = pd.read_excel(source)
    df.columns = [str(c).strip() for c in df.columns]
    required = ["CSI Code","Item Code","Description","Unit","Manhr/Unit","Page"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Missing columns: " + ", ".join(missing))
    df = df[required].copy()
    for c in ["CSI Code","Item Code","Description","Unit"]:
        df[c] = df[c].fillna("").astype(str).str.strip()
    df["Manhr/Unit"] = pd.to_numeric(df["Manhr/Unit"], errors="coerce")
    return df

with st.sidebar:
    st.header("BNI Database")
    uploaded = st.file_uploader("Upload/replace productivity Excel", type=["xlsx"])

try:
    if uploaded:
        df = load_db(uploaded)
        source = "Uploaded Excel"
    elif DB.exists():
        df = load_db(DB)
        source = DB.name
    else:
        st.warning("Upload the BNI Excel database in the sidebar.")
        st.stop()
except Exception as e:
    st.error(f"Could not load database: {e}")
    st.stop()

st.success(f"Database loaded: {len(df):,} rows — {source}")
st.divider()

st.subheader("1. Find a BNI Item")
q = st.text_input("Search description, CSI code, item code, or unit",
                  placeholder="Example: concrete, asphalt, pipe, 03-30")
matches = df.copy()
if q.strip():
    q = q.lower().strip()
    matches = matches[
        matches["Description"].str.lower().str.contains(q, na=False) |
        matches["CSI Code"].str.lower().str.contains(q, na=False) |
        matches["Item Code"].str.lower().str.contains(q, na=False) |
        matches["Unit"].str.lower().str.contains(q, na=False)
    ]

matches = matches[matches["Manhr/Unit"].notna()]
st.write(f"Matching calculable items: **{len(matches):,}**")

if matches.empty:
    st.info("No matching items with a Manhr/Unit value were found.")
    st.stop()

matches = matches.head(200)
options = {
    idx: f'{row["Description"]} | {row["Unit"]} | {row["Manhr/Unit"]:.4f} MH/{row["Unit"]} | {row["CSI Code"]} | {row["Item Code"]}'
    for idx, row in matches.iterrows()
}
selected = st.selectbox("Select the item", list(options), format_func=lambda x: options[x])
item = df.loc[selected]

st.divider()
st.subheader("2. Quantity")
quantity = st.number_input("Quantity", min_value=0.0, value=500.0, step=1.0)

st.subheader("3. BNI Productivity")
c1, c2 = st.columns(2)
c1.metric("Manhr / Unit", f'{item["Manhr/Unit"]:.4f}')
c2.metric("BNI Page", str(item["Page"]))
st.write(f'**Description:** {item["Description"]}')
st.caption(f'CSI: {item["CSI Code"]}  •  Item Code: {item["Item Code"]}')

if st.button("CALCULATE MAN-HOURS", type="primary", use_container_width=True):
    total = quantity * float(item["Manhr/Unit"])
    st.divider()
    st.subheader("Result")
    st.success(f"Estimated Man-Hours: **{total:,.2f} MH**")
    st.code(
        f"Man-Hours = Quantity × Manhr/Unit\n\n"
        f"= {quantity:,.2f} {item['Unit']} × {item['Manhr/Unit']:.4f} MH/{item['Unit']}\n\n"
        f"= {total:,.2f} MH"
    )
    st.caption("Verify the applicable productivity rate against your licensed BNI Costbook before bidding.")
