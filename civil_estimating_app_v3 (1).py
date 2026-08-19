import streamlit as st
import pandas as pd
from pathlib import Path
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Civil Estimating Software", page_icon="🏗️")
st.title("🏗️ Civil Estimating Software")
st.caption("Version 3 — BNI Productivity + Crew + Excel/PDF Output")

DB = Path(__file__).parent / "construction_costs_pages_16_to_36.xlsx"

@st.cache_data
def load_db(source):
    df = pd.read_excel(source)
    df.columns = [str(c).strip() for c in df.columns]
    required = ["CSI Code", "Item Code", "Description", "Unit", "Manhr/Unit", "Page"]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Missing columns: " + ", ".join(missing))
    df = df[required].copy()
    for c in ["CSI Code", "Item Code", "Description", "Unit"]:
        df[c] = df[c].fillna("").astype(str).str.strip()
    df["Manhr/Unit"] = pd.to_numeric(df["Manhr/Unit"], errors="coerce")
    return df

def make_excel(project, item, quantity, crew, total_mh, crew_hours, days):
    rows = [
        ["CIVIL ESTIMATING SOFTWARE — ESTIMATE OUTPUT", ""],
        ["Project", project],
        ["Date", datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["", ""],
        ["SELECTED BNI ITEM", ""],
        ["CSI Code", item["CSI Code"]],
        ["Item Code", item["Item Code"]],
        ["Description", item["Description"]],
        ["Unit", item["Unit"]],
        ["BNI Manhr/Unit", float(item["Manhr/Unit"])],
        ["BNI Page", item["Page"]],
        ["Quantity", quantity],
        ["Total Man-Hours", total_mh],
        ["Crew Hours", crew_hours],
        ["Estimated Working Days", days],
        ["", ""],
        ["CREW", ""],
        ["Position", "Quantity"],
    ]
    rows += [[position, qty] for position, qty in crew.items()]
    out = BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        pd.DataFrame(rows).to_excel(writer, index=False, header=False, sheet_name="Estimate")
    out.seek(0)
    return out.getvalue()

def make_pdf(project, item, quantity, crew, total_mh, crew_hours, days):
    from reportlab.lib.pagesizes import letter
    from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
    from reportlab.lib import colors
    from reportlab.lib.styles import getSampleStyleSheet
    from reportlab.lib.units import inch
    out = BytesIO()
    doc = SimpleDocTemplate(out, pagesize=letter, rightMargin=.6*inch, leftMargin=.6*inch)
    styles = getSampleStyleSheet()
    story = [
        Paragraph("CIVIL ESTIMATING SOFTWARE", styles["Title"]),
        Paragraph("Estimate / Productivity Report", styles["Heading2"]),
        Paragraph(f"<b>Project:</b> {project}", styles["Normal"]),
        Paragraph(f"<b>Date:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles["Normal"]),
        Spacer(1, 12),
        Paragraph("BNI PRODUCTIVITY", styles["Heading2"])
    ]
    data = [
        ["CSI Code", str(item["CSI Code"])],
        ["Item Code", str(item["Item Code"])],
        ["Description", str(item["Description"])],
        ["Unit", str(item["Unit"])],
        ["Quantity", f'{quantity:,.2f} {item["Unit"]}'],
        ["BNI Manhr/Unit", f'{float(item["Manhr/Unit"]):.4f} MH/{item["Unit"]}'],
        ["BNI Page", str(item["Page"])],
        ["Total Man-Hours", f"{total_mh:,.2f} MH"],
        ["Crew Hours", f"{crew_hours:,.2f} hours"],
        ["Estimated Working Days", f"{days:,.2f} days"],
    ]
    table = Table(data, colWidths=[1.8*inch, 4.8*inch])
    table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.5,colors.grey),
                               ("BACKGROUND",(0,0),(0,-1),colors.whitesmoke),
                               ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold")]))
    story += [table, Spacer(1,12), Paragraph("CREW", styles["Heading2"])]
    crew_table = Table([["Position","Quantity"]] + [[p,str(q)] for p,q in crew.items()])
    crew_table.setStyle(TableStyle([("GRID",(0,0),(-1,-1),.5,colors.grey),
                                    ("BACKGROUND",(0,0),(-1,0),colors.whitesmoke)]))
    story += [crew_table, Spacer(1,12),
              Paragraph("Total Man-Hours = Quantity × BNI Manhr/Unit", styles["Normal"]),
              Paragraph("Crew Hours = Total Man-Hours ÷ Total Crew Members", styles["Normal"]),
              Paragraph("Working Days = Crew Hours ÷ Hours per Workday", styles["Normal"]),
              Spacer(1,10),
              Paragraph("Verify the applicable productivity rate against your licensed BNI Costbook before bidding.", styles["Italic"])]
    doc.build(story)
    out.seek(0)
    return out.getvalue()

with st.sidebar:
    st.header("BNI Database")
    uploaded = st.file_uploader("Upload/replace productivity Excel", type=["xlsx"])
    st.divider()
    project = st.text_input("Project name", "My Civil Construction Project")
    workday_hours = st.number_input("Hours per workday", min_value=1.0, value=8.0, step=.5)

try:
    if uploaded:
        df = load_db(uploaded)
    elif DB.exists():
        df = load_db(DB)
    else:
        st.warning("Upload the BNI Excel database in the sidebar.")
        st.stop()
except Exception as e:
    st.error(f"Could not load database: {e}")
    st.stop()

st.success(f"Database loaded: {len(df):,} rows")
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
        matches["Unit"].str.lower().str.contains(q, na=False)]
matches = matches[matches["Manhr/Unit"].notna()]
if matches.empty:
    st.info("No matching items with a Manhr/Unit value were found.")
    st.stop()

matches = matches.head(200)
options = {idx: f'{r["Description"]} | {r["Unit"]} | {r["Manhr/Unit"]:.4f} MH/{r["Unit"]} | {r["CSI Code"]} | {r["Item Code"]}'
           for idx,r in matches.iterrows()}
selected = st.selectbox("Select the item", list(options), format_func=lambda x: options[x])
item = df.loc[selected]

st.divider()
st.subheader("2. Quantity")
quantity = st.number_input("Quantity", min_value=0.0, value=500.0, step=1.0)

st.subheader("3. BNI Productivity")
c1,c2 = st.columns(2)
c1.metric("Manhr / Unit", f'{item["Manhr/Unit"]:.4f}')
c2.metric("BNI Page", str(item["Page"]))
st.write(f'**Description:** {item["Description"]}')
st.caption(f'CSI: {item["CSI Code"]}  •  Item Code: {item["Item Code"]}')

st.divider()
st.subheader("4. Crew")
crew = {}
for position, default in [("Foreman",1),("Laborer",2),("Equipment Operator",1)]:
    crew[position] = st.number_input(position, min_value=0, value=default, step=1)
total_crew = sum(crew.values())

if st.button("CALCULATE ESTIMATE", type="primary", use_container_width=True):
    if total_crew <= 0:
        st.error("Enter at least one crew member.")
        st.stop()
    total_mh = quantity * float(item["Manhr/Unit"])
    crew_hours = total_mh / total_crew
    days = crew_hours / workday_hours

    st.divider()
    st.subheader("Result")
    r1,r2,r3 = st.columns(3)
    r1.metric("Total Man-Hours", f"{total_mh:,.2f} MH")
    r2.metric("Crew Hours", f"{crew_hours:,.2f} hr")
    r3.metric("Working Days", f"{days:,.2f}")

    st.code(f"Man-Hours = {quantity:,.2f} × {item['Manhr/Unit']:.4f} = {total_mh:,.2f} MH\n"
            f"Crew Hours = {total_mh:,.2f} ÷ {total_crew} workers = {crew_hours:,.2f} hr\n"
            f"Working Days = {crew_hours:,.2f} ÷ {workday_hours:.2f} hr/day = {days:,.2f} days")

    st.divider()
    st.subheader("5. Export")
    excel_bytes = make_excel(project,item,quantity,crew,total_mh,crew_hours,days)
    pdf_bytes = make_pdf(project,item,quantity,crew,total_mh,crew_hours,days)
    e1,e2 = st.columns(2)
    with e1:
        st.download_button("📊 Download Excel Estimate", excel_bytes,
                           "civil_estimate_output.xlsx",
                           "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                           use_container_width=True)
    with e2:
        st.download_button("📄 Download PDF Report", pdf_bytes,
                           "civil_estimate_report.pdf", "application/pdf",
                           use_container_width=True)

st.divider()
st.caption("Next: multiple estimate line items, labor/equipment/material pricing, markups, and full bid-sheet export.")
