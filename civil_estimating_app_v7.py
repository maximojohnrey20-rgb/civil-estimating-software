import streamlit as st
import pandas as pd
from pathlib import Path
from io import BytesIO
from datetime import datetime

st.set_page_config(
    page_title="Civil Estimating Software",
    page_icon="🏗️",
    layout="wide"
)

st.title("🏗️ Civil Estimating Software")
st.caption(
    "Version 7 — Transparent Estimate / BNI Productivity / Cost Assembly"
)

DB = Path(__file__).parent / "construction_costs_pages_16_to_36.xlsx"


# ============================================================
# BNI DATABASE
# ============================================================

@st.cache_data
def load_bni(source):

    df = pd.read_excel(source)
    df.columns = [str(c).strip() for c in df.columns]

    required = [
        "CSI Code",
        "Item Code",
        "Description",
        "Unit",
        "Manhr/Unit",
        "Page"
    ]

    missing = [c for c in required if c not in df.columns]

    if missing:
        raise ValueError(
            "Missing columns: " + ", ".join(missing)
        )

    df = df[required].copy()

    for c in [
        "CSI Code",
        "Item Code",
        "Description",
        "Unit"
    ]:
        df[c] = (
            df[c]
            .fillna("")
            .astype(str)
            .str.strip()
        )

    df["Manhr/Unit"] = pd.to_numeric(
        df["Manhr/Unit"],
        errors="coerce"
    )

    return df


# ============================================================
# DEFAULT COST DATABASES
# ============================================================

DEFAULT_LABOR = pd.DataFrame([
    ["Foreman", "HR", 55.00, "User", ""],
    ["Laborer", "HR", 42.00, "User", ""],
    ["Equipment Operator", "HR", 48.00, "User", ""],
], columns=[
    "Description",
    "Unit",
    "Rate",
    "Source",
    "Date Updated"
])


DEFAULT_EQUIPMENT = pd.DataFrame([
    ["Excavator", "HR", 125.00, "User", ""],
    ["CTL", "HR", 110.00, "User", ""],
    ["Tri-Axle", "HR", 95.00, "User", ""],
    ["Loader", "HR", 110.00, "User", ""],
    ["Roller", "HR", 100.00, "User", ""],
], columns=[
    "Description",
    "Unit",
    "Rate",
    "Source",
    "Date Updated"
])


DEFAULT_MATERIAL = pd.DataFrame([
    ["8\" PVC Sanitary Pipe", "LF", 0.00, "User", ""],
    ["Crushed Gravel", "CY", 32.00, "User", ""],
    ["Fill Material - Natural", "CY", 0.00, "User", ""],
    ["Fill Material - Import", "CY", 0.00, "User", ""],
    ["Concrete", "CY", 165.00, "User", ""],
    ["Asphalt", "TON", 95.00, "User", ""],
], columns=[
    "Description",
    "Unit",
    "Rate",
    "Source",
    "Date Updated"
])


# ============================================================
# SESSION STATE
# ============================================================

if "labor_db" not in st.session_state:
    st.session_state.labor_db = DEFAULT_LABOR.copy()

if "equipment_db" not in st.session_state:
    st.session_state.equipment_db = DEFAULT_EQUIPMENT.copy()

if "material_db" not in st.session_state:
    st.session_state.material_db = DEFAULT_MATERIAL.copy()

if "estimate_items" not in st.session_state:
    st.session_state.estimate_items = []

if "equipment_lines" not in st.session_state:
    st.session_state.equipment_lines = []

if "material_lines" not in st.session_state:
    st.session_state.material_lines = []

if "labor_lines" not in st.session_state:
    st.session_state.labor_lines = []

if "current_item" not in st.session_state:
    st.session_state.current_item = None


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("PROJECT")

    project = st.text_input(
        "Project Name",
        "My Civil Construction Project"
    )

    estimator = st.text_input(
        "Estimator",
        ""
    )

    workday_hours = st.number_input(
        "Hours per Workday",
        min_value=1.0,
        value=8.0,
        step=0.5
    )

    st.divider()

    st.header("MARKUPS")

    overhead_pct = st.number_input(
        "Overhead %",
        min_value=0.0,
        value=15.0,
        step=0.5
    )

    profit_pct = st.number_input(
        "Profit %",
        min_value=0.0,
        value=10.0,
        step=0.5
    )

    st.divider()

    st.header("BNI DATABASE")

    uploaded = st.file_uploader(
        "Upload BNI productivity Excel",
        type=["xlsx"]
    )


# ============================================================
# LOAD BNI
# ============================================================

try:

    if uploaded:
        df = load_bni(uploaded)

    elif DB.exists():
        df = load_bni(DB)

    else:
        st.warning(
            "Upload the BNI Excel database."
        )
        st.stop()

except Exception as e:

    st.error(
        f"Could not load BNI database: {e}"
    )
    st.stop()


st.success(
    f"BNI database loaded: {len(df):,} rows"
)


# ============================================================
# COST DATABASE
# ============================================================

st.divider()

st.header("⚙️ COST DATABASE")

cost_tabs = st.tabs([
    "👷 Labor",
    "🚜 Equipment",
    "🧱 Materials"
])


with cost_tabs[0]:

    st.subheader("Labor Rates")

    st.session_state.labor_db = st.data_editor(
        st.session_state.labor_db,
        num_rows="dynamic",
        use_container_width=True,
        key="labor_editor_v7"
    )


with cost_tabs[1]:

    st.subheader("Equipment Rates")

    st.session_state.equipment_db = st.data_editor(
        st.session_state.equipment_db,
        num_rows="dynamic",
        use_container_width=True,
        key="equipment_editor_v7"
    )


with cost_tabs[2]:

    st.subheader("Material Prices")

    st.session_state.material_db = st.data_editor(
        st.session_state.material_db,
        num_rows="dynamic",
        use_container_width=True,
        key="material_editor_v7"
    )


# ============================================================
# BUILD ITEM
# ============================================================

st.divider()

st.header("🏗️ BUILD BID ITEM")

st.subheader("1. Scope of Work")

scope = st.text_area(
    "Describe the scope exactly as you want it to appear in the estimate.",
    placeholder=(
        "Example: 2,050 LF of 8\" PVC Sanitary Line "
        "(0'-6') average depth under paved, unpaved, "
        "and repaired areas."
    ),
    height=100
)


st.subheader("2. Find BNI Item")

search = st.text_input(
    "Search BNI description, CSI code, item code, or unit",
    placeholder="Example: sanitary, PVC, pipe, asphalt"
)

matches = df.copy()

if search.strip():

    q = search.lower().strip()

    matches = matches[
        matches["Description"]
        .str.lower()
        .str.contains(q, na=False)

        |

        matches["CSI Code"]
        .str.lower()
        .str.contains(q, na=False)

        |

        matches["Item Code"]
        .str.lower()
        .str.contains(q, na=False)

        |

        matches["Unit"]
        .str.lower()
        .str.contains(q, na=False)
    ]


matches = matches[
    matches["Manhr/Unit"].notna()
].head(200)


if matches.empty:

    st.info("No BNI items found.")
    st.stop()


options = {

    index:
        f'{row["Description"]} | '
        f'{row["Unit"]} | '
        f'{row["Manhr/Unit"]:.4f} MH/{row["Unit"]} | '
        f'{row["CSI Code"]} | '
        f'{row["Item Code"]}'

    for index, row in matches.iterrows()
}


selected = st.selectbox(
    "Select BNI productivity item",
    list(options.keys()),
    format_func=lambda x: options[x]
)


item = df.loc[selected]


# ============================================================
# QUANTITY
# ============================================================

st.subheader("3. Quantity")

quantity = st.number_input(
    f'Quantity ({item["Unit"]})',
    min_value=0.0,
    value=2050.0,
    step=1.0
)


# ============================================================
# BNI REFERENCE
# ============================================================

st.subheader("4. BNI Productivity Reference")

b1, b2, b3, b4 = st.columns(4)

b1.metric(
    "BNI Man-Hours / Unit",
    f'{float(item["Manhr/Unit"]):.4f}'
)

b2.metric(
    "Quantity",
    f'{quantity:,.2f} {item["Unit"]}'
)

b3.metric(
    "BNI Page",
    str(item["Page"])
)

b4.metric(
    "Unit",
    str(item["Unit"])
)


st.write(
    f'**BNI Description:** {item["Description"]}'
)

st.write(
    f'**CSI Code:** {item["CSI Code"]}'
)

st.write(
    f'**Item Code:** {item["Item Code"]}'
)


# ============================================================
# BNI CALCULATION
# ============================================================

bni_mh_per_unit = float(
    item["Manhr/Unit"]
)

total_bni_mh = (
    quantity *
    bni_mh_per_unit
)


st.markdown("### 📐 BNI Calculation")

st.code(
    f"BNI Man-Hours\n"
    f"= Quantity × BNI Man-Hours/Unit\n\n"
    f"= {quantity:,.2f} {item['Unit']} × "
    f"{bni_mh_per_unit:.4f} MH/{item['Unit']}\n\n"
    f"= {total_bni_mh:,.2f} MH"
)


# ============================================================
# CREW
# ============================================================

st.subheader("5. Crew Composition")

c1, c2, c3 = st.columns(3)

with c1:

    foreman = st.number_input(
        "Foreman",
        min_value=0,
        value=1,
        step=1
    )

with c2:

    laborer = st.number_input(
        "Laborer",
        min_value=0,
        value=2,
        step=1
    )

with c3:

    operator = st.number_input(
        "Equipment Operator",
        min_value=0,
        value=1,
        step=1
    )


total_crew = (
    foreman +
    laborer +
    operator
)


if total_crew > 0:

    crew_hours = (
        total_bni_mh /
        total_crew
    )

    production_days = (
        crew_hours /
        workday_hours
    )

    production_per_day = (
        quantity /
        production_days
        if production_days > 0
        else 0
    )

else:

    crew_hours = 0
    production_days = 0
    production_per_day = 0


# ============================================================
# PRODUCTION
# ============================================================

st.subheader("6. Production / Duration")

p1, p2, p3, p4 = st.columns(4)

p1.metric(
    "Total BNI MH",
    f"{total_bni_mh:,.2f}"
)

p2.metric(
    "Crew Members",
    f"{total_crew}"
)

p3.metric(
    "Crew Hours",
    f"{crew_hours:,.2f}"
)

p4.metric(
    "Working Days",
    f"{production_days:,.2f}"
)


st.markdown("### 📐 Production Derivation")

st.code(
    f"Crew Hours\n"
    f"= Total BNI Man-Hours ÷ Total Crew\n\n"
    f"= {total_bni_mh:,.2f} ÷ {total_crew}\n\n"
    f"= {crew_hours:,.2f} hours\n\n"
    f"Working Days\n"
    f"= Crew Hours ÷ Hours per Workday\n\n"
    f"= {crew_hours:,.2f} ÷ {workday_hours:.2f}\n\n"
    f"= {production_days:,.2f} days\n\n"
    f"Average Production\n"
    f"= {quantity:,.2f} ÷ {production_days:,.2f}\n\n"
    f"= {production_per_day:,.2f} {item['Unit']}/day"
)


# ============================================================
# QUANTITY DERIVATION
# ============================================================

st.divider()

st.subheader(
    "7. Quantity / Excavation Derivation"
)

use_excavation = st.checkbox(
    "Include excavation / trench calculation",
    value=False
)


excavation_data = {}


if use_excavation:

    q1, q2, q3 = st.columns(3)

    with q1:

        trench_width = st.number_input(
            "Trench Width (ft)",
            min_value=0.0,
            value=3.0,
            step=0.1
        )

    with q2:

        trench_depth = st.number_input(
            "Average Depth (ft)",
            min_value=0.0,
            value=6.0,
            step=0.1
        )

    with q3:

        overexcavation = st.number_input(
            "Overexcavation %",
            min_value=0.0,
            value=0.0,
            step=0.5
        )


    base_excavation = (
        quantity *
        trench_width *
        trench_depth /
        27
    )


    excavation_qty = (
        base_excavation *
        (1 + overexcavation / 100)
    )


    st.markdown("### 📐 Excavation Formula")

    st.code(
        f"Excavation Volume\n"
        f"= Length × Width × Depth ÷ 27\n\n"
        f"= {quantity:,.2f} × "
        f"{trench_width:.2f} × "
        f"{trench_depth:.2f} ÷ 27\n\n"
        f"= {base_excavation:,.2f} CY\n\n"
        f"With {overexcavation:.1f}% overexcavation:\n"
        f"= {excavation_qty:,.2f} CY"
    )


    excavation_data = {

        "Length LF":
            quantity,

        "Trench Width FT":
            trench_width,

        "Average Depth FT":
            trench_depth,

        "Overexcavation %":
            overexcavation,

        "Base Excavation CY":
            base_excavation,

        "Total Excavation CY":
            excavation_qty
    }


# ============================================================
# LABOR
# ============================================================

st.divider()

st.subheader(
    "8. Labor Breakdown"
)


labor_db = st.session_state.labor_db

labor_names = (
    labor_db["Description"]
    .dropna()
    .astype(str)
    .tolist()
)


labor_lines = []


for position, workers in [
    ("Foreman", foreman),
    ("Laborer", laborer),
    ("Equipment Operator", operator)
]:

    if workers <= 0:
        continue


    default_index = (
        labor_names.index(position)
        if position in labor_names
        else 0
    )


    selected_labor = st.selectbox(
        f"{position} Rate",
        labor_names,
        index=default_index,
        key=f"labor_v7_{position}"
    )


    rate_row = labor_db[
        labor_db["Description"].astype(str)
        == selected_labor
    ].iloc[0]


    rate = float(
        rate_row["Rate"]
    )


    hours = crew_hours

    cost = (
        workers *
        hours *
        rate
    )


    labor_lines.append({

        "Position": position,
        "Workers": workers,
        "Hours": hours,
        "Rate": rate,
        "Cost": cost
    })


labor_df = pd.DataFrame(
    labor_lines
)


if not labor_df.empty:

    st.dataframe(
        labor_df,
        use_container_width=True,
        hide_index=True
    )


labor_total = sum(
    x["Cost"]
    for x in labor_lines
)


st.metric(
    "TOTAL LABOR",
    f"${labor_total:,.2f}"
)


# ============================================================
# EQUIPMENT
# ============================================================

st.divider()

st.subheader(
    "9. Equipment Breakdown"
)


equipment_db = st.session_state.equipment_db

equipment_names = (
    equipment_db["Description"]
    .dropna()
    .astype(str)
    .tolist()
)


e1, e2, e3 = st.columns(3)


with e1:

    selected_equipment = st.selectbox(
        "Equipment",
        ["None"] + equipment_names,
        key="equipment_select_v7"
    )


with e2:

    equipment_hours = st.number_input(
        "Equipment Hours",
        min_value=0.0,
        value=0.0,
        step=1.0,
        key="equipment_hours_v7"
    )


with e3:

    if selected_equipment != "None":

        eq_row = equipment_db[
            equipment_db["Description"].astype(str)
            == selected_equipment
        ].iloc[0]

        equipment_rate = float(
            eq_row["Rate"]
        )

        equipment_unit = str(
            eq_row["Unit"]
        )

    else:

        equipment_rate = 0.0
        equipment_unit = "HR"


    st.metric(
        "Equipment Rate",
        f"${equipment_rate:,.2f}/{equipment_unit}"
    )


if st.button(
    "➕ ADD EQUIPMENT",
    key="add_equipment_v7"
):

    if (
        selected_equipment != "None"
        and equipment_hours > 0
    ):

        st.session_state.equipment_lines.append({

            "Description":
                selected_equipment,

            "Hours":
                equipment_hours,

            "Rate":
                equipment_rate,

            "Cost":
                equipment_hours *
                equipment_rate
        })


if st.session_state.equipment_lines:

    equipment_df = pd.DataFrame(
        st.session_state.equipment_lines
    )

    st.dataframe(
        equipment_df,
        use_container_width=True,
        hide_index=True
    )


equipment_total = sum(
    x["Cost"]
    for x in st.session_state.equipment_lines
)


st.metric(
    "TOTAL EQUIPMENT",
    f"${equipment_total:,.2f}"
)


# ============================================================
# MATERIALS
# ============================================================

st.divider()

st.subheader(
    "10. Material Breakdown"
)


material_db = st.session_state.material_db

material_names = (
    material_db["Description"]
    .dropna()
    .astype(str)
    .tolist()
)


m1, m2, m3 = st.columns(3)


with m1:

    selected_material = st.selectbox(
        "Material",
        ["None"] + material_names,
        key="material_select_v7"
    )


with m2:

    material_quantity = st.number_input(
        "Material Quantity",
        min_value=0.0,
        value=0.0,
        step=1.0,
        key="material_quantity_v7"
    )


with m3:

    if selected_material != "None":

        mat_row = material_db[
            material_db["Description"].astype(str)
            == selected_material
        ].iloc[0]

        material_rate = float(
            mat_row["Rate"]
        )

        material_unit = str(
            mat_row["Unit"]
        )

    else:

        material_rate = 0.0
        material_unit = ""


    st.metric(
        "Material Price",
        f"${material_rate:,.2f}/{material_unit}"
        if material_unit
        else "$0.00"
    )


if st.button(
    "➕ ADD MATERIAL",
    key="add_material_v7"
):

    if (
        selected_material != "None"
        and material_quantity > 0
    ):

        st.session_state.material_lines.append({

            "Description":
                selected_material,

            "Quantity":
                material_quantity,

            "Unit":
                material_unit,

            "Rate":
                material_rate,

            "Cost":
                material_quantity *
                material_rate
        })


if st.session_state.material_lines:

    material_df = pd.DataFrame(
        st.session_state.material_lines
    )

    st.dataframe(
        material_df,
        use_container_width=True,
        hide_index=True
    )


material_total = sum(
    x["Cost"]
    for x in st.session_state.material_lines
)


st.metric(
    "TOTAL MATERIAL",
    f"${material_total:,.2f}"
)


# ============================================================
# FILL / HAULING
# ============================================================

st.divider()

st.subheader(
    "11. Fill / Hauling"
)


f1, f2 = st.columns(2)


with f1:

    natural_fill = st.number_input(
        "Natural Fill Quantity (CY)",
        min_value=0.0,
        value=0.0,
        step=1.0
    )


with f2:

    import_fill = st.number_input(
        "Import Fill Quantity (CY)",
        min_value=0.0,
        value=0.0,
        step=1.0
    )


fill_total = (
    natural_fill +
    import_fill
)


st.write(
    f"**Total Fill Material:** "
    f"{fill_total:,.2f} CY"
)


# ============================================================
# SUBCONTRACT
# ============================================================

st.divider()

st.subheader(
    "12. Subcontract"
)


subcontract_cost = st.number_input(
    "Subcontract Cost",
    min_value=0.0,
    value=0.0,
    step=100.0
)


# ============================================================
# COST SUMMARY
# ============================================================

direct_cost = (
    labor_total
    +
    equipment_total
    +
    material_total
    +
    subcontract_cost
)


overhead_amount = (
    direct_cost *
    overhead_pct /
    100
)


cost_after_overhead = (
    direct_cost +
    overhead_amount
)


profit_amount = (
    cost_after_overhead *
    profit_pct /
    100
)


final_bid = (
    cost_after_overhead +
    profit_amount
)


cost_per_unit = (
    final_bid / quantity
    if quantity > 0
    else 0
)


st.divider()

st.header(
    "💰 13. COST SUMMARY"
)


a, b, c, d = st.columns(4)


a.metric(
    "Labor",
    f"${labor_total:,.2f}"
)

b.metric(
    "Equipment",
    f"${equipment_total:,.2f}"
)

c.metric(
    "Materials",
    f"${material_total:,.2f}"
)

d.metric(
    "Subcontract",
    f"${subcontract_cost:,.2f}"
)


st.divider()


a, b, c = st.columns(3)


a.metric(
    "DIRECT COST",
    f"${direct_cost:,.2f}"
)

b.metric(
    "OVERHEAD",
    f"${overhead_amount:,.2f}"
)

c.metric(
    "PROFIT",
    f"${profit_amount:,.2f}"
)


st.success(
    f"🏆 FINAL BID: ${final_bid:,.2f}"
)


st.metric(
    f"FINAL COST / {item['Unit']}",
    f"${cost_per_unit:,.2f}"
)


# ============================================================
# ASSUMPTIONS
# ============================================================

st.divider()

st.subheader(
    "📝 14. Estimating Assumptions / Notes"
)


assumptions = st.text_area(
    "Notes that will appear in the final estimate",
    value=(
        "BNI productivity rate is used as the estimating "
        "productivity reference. Actual field production may "
        "vary depending on site conditions, crew composition, "
        "access, weather, soil conditions, equipment, and "
        "project-specific requirements."
    ),
    height=140
)


# ============================================================
# ADD ITEM TO ESTIMATE
# ============================================================

st.divider()


if st.button(
    "➕ ADD THIS COMPLETE ITEM TO ESTIMATE",
    type="primary",
    use_container_width=True
):

    complete_item = {

        "Scope":
            scope,

        "Description":
            str(item["Description"]),

        "CSI Code":
            str(item["CSI Code"]),

        "Item Code":
            str(item["Item Code"]),

        "BNI Page":
            str(item["Page"]),

        "Quantity":
            quantity,

        "Unit":
            str(item["Unit"]),

        "BNI MH/Unit":
            bni_mh_per_unit,

        "Total BNI MH":
            total_bni_mh,

        "Crew":
            total_crew,

        "Crew Hours":
            crew_hours,

        "Working Days":
            production_days,

        "Production / Day":
            production_per_day,

        "Labor":
            labor_total,

        "Equipment":
            equipment_total,

        "Materials":
            material_total,

        "Subcontract":
            subcontract_cost,

        "Direct Cost":
            direct_cost,

        "Overhead":
            overhead_amount,

        "Profit":
            profit_amount,

        "Final Bid":
            final_bid,

        "Cost / Unit":
            cost_per_unit,

        "Labor Details":
            labor_lines.copy(),

        "Equipment Details":
            st.session_state.equipment_lines.copy(),

        "Material Details":
            st.session_state.material_lines.copy(),

        "Excavation":
            excavation_data.copy(),

        "Natural Fill":
            natural_fill,

        "Import Fill":
            import_fill,

        "Assumptions":
            assumptions
    }


    st.session_state.estimate_items.append(
        complete_item
    )


    st.session_state.equipment_lines = []
    st.session_state.material_lines = []


    st.success(
        "Complete bid item added to the estimate."
    )


# ============================================================
# FINAL ESTIMATE
# ============================================================

st.divider()

st.header(
    "📋 15. FINAL ESTIMATE"
)


items = st.session_state.estimate_items


if items:

    summary_rows = []


    for number, x in enumerate(
        items,
        start=1
    ):

        summary_rows.append({

            "Line":
                number,

            "Scope":
                x["Scope"],

            "Qty":
                x["Quantity"],

            "Unit":
                x["Unit"],

            "Labor":
                x["Labor"],

            "Equipment":
                x["Equipment"],

            "Materials":
                x["Materials"],

            "Direct Cost":
                x["Direct Cost"],

            "Overhead":
                x["Overhead"],

            "Profit":
                x["Profit"],

            "FINAL BID":
                x["Final Bid"],

            "Cost / Unit":
                x["Cost / Unit"],

            "Days":
                x["Working Days"]

        })


    summary_df = pd.DataFrame(
        summary_rows
    )


    st.dataframe(
        summary_df,
        use_container_width=True,
        hide_index=True
    )


    total_labor = sum(
        x["Labor"]
        for x in items
    )

    total_equipment = sum(
        x["Equipment"]
        for x in items
    )

    total_material = sum(
        x["Materials"]
        for x in items
    )

    total_direct = sum(
        x["Direct Cost"]
        for x in items
    )

    total_overhead = sum(
        x["Overhead"]
        for x in items
    )

    total_profit = sum(
        x["Profit"]
        for x in items
    )

    total_bid = sum(
        x["Final Bid"]
        for x in items
    )


    st.success(
        f"🏆 TOTAL PROJECT BID: ${total_bid:,.2f}"
    )


else:

    st.info(
        "Your final estimate will appear here."
    )


# ============================================================
# EXCEL EXPORT
# ============================================================

st.divider()

st.header(
    "📊 16. CLIENT ESTIMATE EXPORT"
)


if items:

    output = BytesIO()


    # --------------------------------------------------------
    # SUMMARY
    # --------------------------------------------------------

    project_summary = pd.DataFrame([{

        "Project":
            project,

        "Estimator":
            estimator,

        "Date":
            datetime.now().strftime(
                "%Y-%m-%d %H:%M"
            ),

        "Total Items":
            len(items),

        "Labor":
            total_labor,

        "Equipment":
            total_equipment,

        "Materials":
            total_material,

        "Direct Cost":
            total_direct,

        "Overhead %":
            overhead_pct,

        "Overhead":
            total_overhead,

        "Profit %":
            profit_pct,

        "Profit":
            total_profit,

        "FINAL BID":
            total_bid
    }])


    # --------------------------------------------------------
    # PRODUCTIVITY
    # --------------------------------------------------------

    productivity_rows = []


    for number, x in enumerate(
        items,
        start=1
    ):

        productivity_rows.append({

            "Line":
                number,

            "Scope":
                x["Scope"],

            "BNI Description":
                x["Description"],

            "CSI Code":
                x["CSI Code"],

            "Item Code":
                x["Item Code"],

            "BNI Page":
                x["BNI Page"],

            "Quantity":
                x["Quantity"],

            "Unit":
                x["Unit"],

            "BNI MH/Unit":
                x["BNI MH/Unit"],

            "Total BNI MH":
                x["Total BNI MH"],

            "Crew":
                x["Crew"],

            "Crew Hours":
                x["Crew Hours"],

            "Working Days":
                x["Working Days"],

            "Production / Day":
                x["Production / Day"]
        })


    productivity_df = pd.DataFrame(
        productivity_rows
    )


    # --------------------------------------------------------
    # RESOURCE BREAKDOWN
    # --------------------------------------------------------

    resource_rows = []


    for number, x in enumerate(
        items,
        start=1
    ):

        for labor in x["Labor Details"]:

            resource_rows.append({

                "Line":
                    number,

                "Scope":
                    x["Scope"],

                "Type":
                    "Labor",

                "Resource":
                    labor["Position"],

                "Quantity":
                    labor["Workers"],

                "Hours":
                    labor["Hours"],

                "Unit":
                    "HR",

                "Rate":
                    labor["Rate"],

                "Cost":
                    labor["Cost"]
            })


        for equipment in x["Equipment Details"]:

            resource_rows.append({

                "Line":
                    number,

                "Scope":
                    x["Scope"],

                "Type":
                    "Equipment",

                "Resource":
                    equipment["Description"],

                "Quantity":
                    equipment["Hours"],

                "Hours":
                    equipment["Hours"],

                "Unit":
                    "HR",

                "Rate":
                    equipment["Rate"],

                "Cost":
                    equipment["Cost"]
            })


        for material in x["Material Details"]:

            resource_rows.append({

                "Line":
                    number,

                "Scope":
                    x["Scope"],

                "Type":
                    "Material",

                "Resource":
                    material["Description"],

                "Quantity":
                    material["Quantity"],

                "Hours":
                    "",

                "Unit":
                    material["Unit"],

                "Rate":
                    material["Rate"],

                "Cost":
                    material["Cost"]
            })


    resource_df = pd.DataFrame(
        resource_rows
    )


    # --------------------------------------------------------
    # EXCAVATION
    # --------------------------------------------------------

    excavation_rows = []


    for number, x in enumerate(
        items,
        start=1
    ):

        ex = x["Excavation"]


        if ex:

            excavation_rows.append({

                "Line":
                    number,

                "Scope":
                    x["Scope"],

                "Length LF":
                    ex["Length LF"],

                "Trench Width FT":
                    ex["Trench Width FT"],

                "Average Depth FT":
                    ex["Average Depth FT"],

                "Overexcavation %":
                    ex["Overexcavation %"],

                "Base Excavation CY":
                    ex["Base Excavation CY"],

                "Total Excavation CY":
                    ex["Total Excavation CY"],

                "Natural Fill CY":
                    x["Natural Fill"],

                "Import Fill CY":
                    x["Import Fill"]

            })


    excavation_df = pd.DataFrame(
        excavation_rows
    )


    # --------------------------------------------------------
    # ASSUMPTIONS
    # --------------------------------------------------------

    assumption_rows = []


    for number, x in enumerate(
        items,
        start=1
    ):

        assumption_rows.append({

            "Line":
                number,

            "Scope":
                x["Scope"],

            "BNI Reference":
                f'BNI Costbook - Page {x["BNI Page"]}',

            "Assumptions":
                x["Assumptions"]

        })


    assumptions_df = pd.DataFrame(
        assumption_rows
    )


    # --------------------------------------------------------
    # WRITE EXCEL
    # --------------------------------------------------------

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        project_summary.to_excel(
            writer,
            index=False,
            sheet_name="Bid Summary"
        )

        summary_df.to_excel(
            writer,
            index=False,
            sheet_name="Estimate"
        )

        productivity_df.to_excel(
            writer,
            index=False,
            sheet_name="BNI Productivity"
        )

        resource_df.to_excel(
            writer,
            index=False,
            sheet_name="Resource Breakdown"
        )

        excavation_df.to_excel(
            writer,
            index=False,
            sheet_name="Excavation"
        )

        assumptions_df.to_excel(
            writer,
            index=False,
            sheet_name="Assumptions"
        )

        st.session_state.labor_db.to_excel(
            writer,
            index=False,
            sheet_name="Labor Database"
        )

        st.session_state.equipment_db.to_excel(
            writer,
            index=False,
            sheet_name="Equipment Database"
        )

        st.session_state.material_db.to_excel(
            writer,
            index=False,
            sheet_name="Material Database"
        )


        # Basic formatting
        workbook = writer.book

        for worksheet in workbook.worksheets:

            worksheet.freeze_panes = "A2"

            for column in worksheet.columns:

                max_length = 0

                column_letter = column[0].column_letter

                for cell in column:

                    try:

                        length = len(
                            str(cell.value)
                        )

                        if length > max_length:
                            max_length = length

                    except Exception:
                        pass

                worksheet.column_dimensions[
                    column_letter
                ].width = min(
                    max_length + 2,
                    45
                )


    output.seek(0)


    st.download_button(

        "📥 DOWNLOAD CLIENT ESTIMATE — EXCEL",

        output.getvalue(),

        "civil_estimate_v7.xlsx",

        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

        use_container_width=True
    )


else:

    st.info(
        "Add at least one complete bid item to enable the client estimate."
    )


# ============================================================
# FOOTER
# ============================================================

st.divider()

st.caption(
    "Version 7 — Transparent Estimating: "
    "Scope → Quantity → BNI Productivity → Crew → "
    "Production → Labor → Equipment → Materials → "
    "Excavation → Direct Cost → Markup → Final Bid"
)
