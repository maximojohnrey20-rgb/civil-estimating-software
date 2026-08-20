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
st.caption("Version 6 — BNI Productivity + Labor + Equipment + Material Cost Database")

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
    {
        "Description": "Foreman",
        "Unit": "HR",
        "Rate": 55.00,
        "Source": "User",
        "Date Updated": ""
    },
    {
        "Description": "Laborer",
        "Unit": "HR",
        "Rate": 42.00,
        "Source": "User",
        "Date Updated": ""
    },
    {
        "Description": "Equipment Operator",
        "Unit": "HR",
        "Rate": 48.00,
        "Source": "User",
        "Date Updated": ""
    }
])


DEFAULT_EQUIPMENT = pd.DataFrame([
    {
        "Description": "Excavator",
        "Unit": "HR",
        "Rate": 125.00,
        "Source": "User",
        "Date Updated": ""
    },
    {
        "Description": "Loader",
        "Unit": "HR",
        "Rate": 110.00,
        "Source": "User",
        "Date Updated": ""
    },
    {
        "Description": "Dump Truck",
        "Unit": "HR",
        "Rate": 95.00,
        "Source": "User",
        "Date Updated": ""
    },
    {
        "Description": "Roller",
        "Unit": "HR",
        "Rate": 100.00,
        "Source": "User",
        "Date Updated": ""
    }
])


DEFAULT_MATERIAL = pd.DataFrame([
    {
        "Description": "Crushed Gravel",
        "Unit": "CY",
        "Rate": 32.00,
        "Source": "User",
        "Date Updated": ""
    },
    {
        "Description": "Asphalt",
        "Unit": "TON",
        "Rate": 95.00,
        "Source": "User",
        "Date Updated": ""
    },
    {
        "Description": "Concrete",
        "Unit": "CY",
        "Rate": 165.00,
        "Source": "User",
        "Date Updated": ""
    },
    {
        "Description": "HDPE Pipe",
        "Unit": "LF",
        "Rate": 12.00,
        "Source": "User",
        "Date Updated": ""
    }
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


# ============================================================
# SIDEBAR
# ============================================================

with st.sidebar:

    st.header("PROJECT")

    project = st.text_input(
        "Project Name",
        "My Civil Construction Project"
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
# COST DATABASE MANAGEMENT
# ============================================================

st.divider()

st.header("⚙️ COST DATABASE")

st.write(
    "These are your editable labor, equipment, and material rates."
)


tabs = st.tabs([
    "👷 Labor",
    "🚜 Equipment",
    "🧱 Materials"
])


# ============================================================
# LABOR DATABASE
# ============================================================

with tabs[0]:

    st.subheader("Labor Rates")

    edited_labor = st.data_editor(
        st.session_state.labor_db,
        num_rows="dynamic",
        use_container_width=True,
        key="labor_editor"
    )

    st.session_state.labor_db = edited_labor

    st.info(
        "Example rates are placeholders. Replace them with your actual estimating rates."
    )


# ============================================================
# EQUIPMENT DATABASE
# ============================================================

with tabs[1]:

    st.subheader("Equipment Rates")

    edited_equipment = st.data_editor(
        st.session_state.equipment_db,
        num_rows="dynamic",
        use_container_width=True,
        key="equipment_editor"
    )

    st.session_state.equipment_db = edited_equipment

    st.info(
        "Enter your equipment ownership/rental or estimating rates."
    )


# ============================================================
# MATERIAL DATABASE
# ============================================================

with tabs[2]:

    st.subheader("Material Prices")

    edited_material = st.data_editor(
        st.session_state.material_db,
        num_rows="dynamic",
        use_container_width=True,
        key="material_editor"
    )

    st.session_state.material_db = edited_material

    st.info(
        "Enter your actual material prices and update them when needed."
    )


# ============================================================
# SEARCH BNI ITEM
# ============================================================

st.divider()

st.header("🏗️ BUILD ESTIMATE")

st.subheader("1. Find a BNI Item")

search = st.text_input(
    "Search description, CSI code, item code, or unit",
    placeholder="Example: asphalt, concrete, pipe"
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

    st.info(
        "No BNI items found."
    )

    st.stop()


options = {

    index:
        f'{row["Description"]} | '
        f'{row["Unit"]} | '
        f'{row["Manhr/Unit"]:.4f} MH/{row["Unit"]} | '
        f'{row["CSI Code"]} | '
        f'{row["Item Code"]}'

    for index, row
    in matches.iterrows()
}


selected = st.selectbox(
    "Select BNI Item",
    list(options.keys()),
    format_func=lambda x: options[x]
)


item = df.loc[selected]


# ============================================================
# QUANTITY
# ============================================================

st.subheader("2. Quantity")

quantity = st.number_input(
    f'Quantity ({item["Unit"]})',
    min_value=0.0,
    value=500.0,
    step=1.0
)


# ============================================================
# PRODUCTIVITY
# ============================================================

st.subheader("3. BNI Productivity")

c1, c2, c3 = st.columns(3)

c1.metric(
    "BNI Manhr / Unit",
    f'{item["Manhr/Unit"]:.4f}'
)

c2.metric(
    "BNI Page",
    str(item["Page"])
)

c3.metric(
    "Unit",
    str(item["Unit"])
)

st.write(
    f'**Description:** {item["Description"]}'
)

st.caption(
    f'CSI: {item["CSI Code"]} • '
    f'Item Code: {item["Item Code"]}'
)


# ============================================================
# CREW
# ============================================================

st.subheader("4. Crew")

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


# ============================================================
# CALCULATE PRODUCTIVITY
# ============================================================

if st.button(
    "CALCULATE ITEM",
    type="primary",
    use_container_width=True
):

    if total_crew <= 0:

        st.error(
            "Enter at least one crew member."
        )

    elif quantity <= 0:

        st.error(
            "Enter a quantity greater than zero."
        )

    else:

        total_mh = (
            quantity *
            float(item["Manhr/Unit"])
        )

        crew_hours = (
            total_mh /
            total_crew
        )

        days = (
            crew_hours /
            workday_hours
        )

        st.session_state.current_item = {

            "CSI Code":
                str(item["CSI Code"]),

            "Item Code":
                str(item["Item Code"]),

            "Description":
                str(item["Description"]),

            "Quantity":
                float(quantity),

            "Unit":
                str(item["Unit"]),

            "Manhr/Unit":
                float(item["Manhr/Unit"]),

            "Page":
                str(item["Page"]),

            "Foreman":
                int(foreman),

            "Laborer":
                int(laborer),

            "Equipment Operator":
                int(operator),

            "Total Crew":
                int(total_crew),

            "Total MH":
                float(total_mh),

            "Crew Hours":
                float(crew_hours),

            "Days":
                float(days)
        }


# ============================================================
# CURRENT ITEM
# ============================================================

if "current_item" in st.session_state:

    current = st.session_state.current_item

    st.divider()

    st.subheader("5. Labor Cost")

    labor_db = st.session_state.labor_db

    labor_names = labor_db[
        "Description"
    ].dropna().astype(str).tolist()


    labor_cost = 0.0

    labor_rows = []


    for position, workers in [
        ("Foreman", current["Foreman"]),
        ("Laborer", current["Laborer"]),
        ("Equipment Operator", current["Equipment Operator"])
    ]:

        if workers <= 0:
            continue

        default_name = position

        if default_name in labor_names:
            default_index = labor_names.index(
                default_name
            )
        else:
            default_index = 0


        selected_labor = st.selectbox(
            f"{position} Rate",
            labor_names,
            index=default_index,
            key=f"labor_{position}"
        )


        rate_row = labor_db[
            labor_db["Description"].astype(str)
            == selected_labor
        ].iloc[0]


        rate = float(
            rate_row["Rate"]
        )


        extension = (
            workers *
            current["Crew Hours"] *
            rate
        )


        labor_cost += extension


        labor_rows.append({

            "Position": position,

            "Workers": workers,

            "Hours": current["Crew Hours"],

            "Rate": rate,

            "Cost": extension

        })


    st.metric(
        "TOTAL LABOR",
        f"${labor_cost:,.2f}"
    )


    # ========================================================
    # EQUIPMENT
    # ========================================================

    st.divider()

    st.subheader(
        "6. Equipment"
    )


    equipment_db = (
        st.session_state.equipment_db
    )


    equipment_names = (
        equipment_db["Description"]
        .dropna()
        .astype(str)
        .tolist()
    )


    if equipment_names:

        eq_col1, eq_col2, eq_col3 = st.columns(3)

        with eq_col1:

            selected_equipment = st.selectbox(
                "Equipment",
                ["None"] + equipment_names,
                key="selected_equipment"
            )

        with eq_col2:

            equipment_hours = st.number_input(
                "Hours",
                min_value=0.0,
                value=0.0,
                step=1.0,
                key="equipment_hours"
            )

        with eq_col3:

            if selected_equipment != "None":

                eq_row = equipment_db[
                    equipment_db["Description"].astype(str)
                    == selected_equipment
                ].iloc[0]

                equipment_rate = float(
                    eq_row["Rate"]
                )

                st.metric(
                    "Rate",
                    f"${equipment_rate:,.2f}/HR"
                )

            else:

                equipment_rate = 0.0

                st.metric(
                    "Rate",
                    "$0.00/HR"
                )


        if st.button(
            "➕ ADD EQUIPMENT",
            key="add_equipment"
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

                st.success(
                    "Equipment added."
                )


    if st.session_state.equipment_lines:

        equipment_table = pd.DataFrame(
            st.session_state.equipment_lines
        )

        st.dataframe(
            equipment_table,
            use_container_width=True,
            hide_index=True
        )


    equipment_cost = sum(
        x["Cost"]
        for x in st.session_state.equipment_lines
    )


    st.metric(
        "TOTAL EQUIPMENT",
        f"${equipment_cost:,.2f}"
    )


    # ========================================================
    # MATERIALS
    # ========================================================

    st.divider()

    st.subheader(
        "7. Materials"
    )


    material_db = (
        st.session_state.material_db
    )


    material_names = (
        material_db["Description"]
        .dropna()
        .astype(str)
        .tolist()
    )


    if material_names:

        mat_col1, mat_col2, mat_col3 = st.columns(3)

        with mat_col1:

            selected_material = st.selectbox(
                "Material",
                ["None"] + material_names,
                key="selected_material"
            )

        with mat_col2:

            material_quantity = st.number_input(
                "Quantity",
                min_value=0.0,
                value=0.0,
                step=1.0,
                key="material_quantity"
            )

        with mat_col3:

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

                st.metric(
                    "Unit Price",
                    f"${material_rate:,.2f}/{material_unit}"
                )

            else:

                material_rate = 0.0
                material_unit = ""

                st.metric(
                    "Unit Price",
                    "$0.00"
                )


        if st.button(
            "➕ ADD MATERIAL",
            key="add_material"
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

                st.success(
                    "Material added."
                )


    if st.session_state.material_lines:

        material_table = pd.DataFrame(
            st.session_state.material_lines
        )

        st.dataframe(
            material_table,
            use_container_width=True,
            hide_index=True
        )


    material_cost = sum(
        x["Cost"]
        for x in st.session_state.material_lines
    )


    st.metric(
        "TOTAL MATERIAL",
        f"${material_cost:,.2f}"
    )


    # ========================================================
    # SUBCONTRACT
    # ========================================================

    st.divider()

    st.subheader(
        "8. Subcontract"
    )


    subcontract_cost = st.number_input(
        "Subcontract Cost",
        min_value=0.0,
        value=0.0,
        step=100.0,
        key="subcontract_cost"
    )


    # ========================================================
    # FINAL COST
    # ========================================================

    direct_cost = (
        labor_cost
        +
        equipment_cost
        +
        material_cost
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


    st.divider()

    st.subheader(
        "9. COST SUMMARY"
    )


    c1, c2, c3, c4 = st.columns(4)


    c1.metric(
        "Labor",
        f"${labor_cost:,.2f}"
    )

    c2.metric(
        "Equipment",
        f"${equipment_cost:,.2f}"
    )

    c3.metric(
        "Material",
        f"${material_cost:,.2f}"
    )

    c4.metric(
        "Subcontract",
        f"${subcontract_cost:,.2f}"
    )


    st.divider()


    c1, c2, c3 = st.columns(3)


    c1.metric(
        "DIRECT COST",
        f"${direct_cost:,.2f}"
    )

    c2.metric(
        f"OVERHEAD ({overhead_pct:.1f}%)",
        f"${overhead_amount:,.2f}"
    )

    c3.metric(
        f"PROFIT ({profit_pct:.1f}%)",
        f"${profit_amount:,.2f}"
    )


    st.success(
        f"🏆 FINAL BID PRICE: ${final_bid:,.2f}"
    )


    # ========================================================
    # ADD TO ESTIMATE
    # ========================================================

    if st.button(
        "➕ ADD PRICED ITEM TO ESTIMATE",
        type="primary",
        use_container_width=True
    ):

        current["Labor Cost"] = labor_cost
        current["Equipment Cost"] = equipment_cost
        current["Material Cost"] = material_cost
        current["Subcontract Cost"] = subcontract_cost
        current["Direct Cost"] = direct_cost
        current["Overhead"] = overhead_amount
        current["Profit"] = profit_amount
        current["Final Bid"] = final_bid

        current["Labor Details"] = labor_rows
        current["Equipment Details"] = (
            st.session_state.equipment_lines.copy()
        )
        current["Material Details"] = (
            st.session_state.material_lines.copy()
        )

        st.session_state.estimate_items.append(
            current.copy()
        )

        st.session_state.equipment_lines = []
        st.session_state.material_lines = []

        st.success(
            f'{current["Description"]} added to estimate.'
        )


# ============================================================
# ESTIMATE BUILDER
# ============================================================

st.divider()

st.header(
    "📋 FINAL ESTIMATE"
)


estimate_items = (
    st.session_state.estimate_items
)


if estimate_items:

    estimate_table = pd.DataFrame([

        {

            "Line":
                n,

            "Item":
                x["Description"],

            "Qty":
                x["Quantity"],

            "Unit":
                x["Unit"],

            "Labor":
                x.get("Labor Cost", 0),

            "Equipment":
                x.get("Equipment Cost", 0),

            "Material":
                x.get("Material Cost", 0),

            "Subcontract":
                x.get("Subcontract Cost", 0),

            "Direct Cost":
                x.get("Direct Cost", 0),

            "Overhead":
                x.get("Overhead", 0),

            "Profit":
                x.get("Profit", 0),

            "FINAL BID":
                x.get("Final Bid", 0)

        }

        for n, x
        in enumerate(
            estimate_items,
            start=1
        )

    ])


    st.dataframe(
        estimate_table,
        use_container_width=True,
        hide_index=True
    )


    total_labor = sum(
        x.get("Labor Cost", 0)
        for x in estimate_items
    )

    total_equipment = sum(
        x.get("Equipment Cost", 0)
        for x in estimate_items
    )

    total_material = sum(
        x.get("Material Cost", 0)
        for x in estimate_items
    )

    total_subcontract = sum(
        x.get("Subcontract Cost", 0)
        for x in estimate_items
    )

    total_direct = sum(
        x.get("Direct Cost", 0)
        for x in estimate_items
    )

    total_overhead = sum(
        x.get("Overhead", 0)
        for x in estimate_items
    )

    total_profit = sum(
        x.get("Profit", 0)
        for x in estimate_items
    )

    total_bid = sum(
        x.get("Final Bid", 0)
        for x in estimate_items
    )


    st.divider()

    a, b, c, d = st.columns(4)

    a.metric(
        "LABOR",
        f"${total_labor:,.2f}"
    )

    b.metric(
        "EQUIPMENT",
        f"${total_equipment:,.2f}"
    )

    c.metric(
        "MATERIAL",
        f"${total_material:,.2f}"
    )

    d.metric(
        "SUBCONTRACT",
        f"${total_subcontract:,.2f}"
    )


    st.divider()

    a, b, c = st.columns(3)

    a.metric(
        "DIRECT COST",
        f"${total_direct:,.2f}"
    )

    b.metric(
        "OVERHEAD",
        f"${total_overhead:,.2f}"
    )

    c.metric(
        "PROFIT",
        f"${total_profit:,.2f}"
    )


    st.success(
        f"🏆 TOTAL FINAL BID: ${total_bid:,.2f}"
    )


    if st.button(
        "🗑️ CLEAR ENTIRE ESTIMATE",
        use_container_width=True
    ):

        st.session_state.estimate_items = []
        st.session_state.equipment_lines = []
        st.session_state.material_lines = []
        st.session_state.pop(
            "current_item",
            None
        )

        st.rerun()


else:

    st.info(
        "No priced items have been added yet."
    )


# ============================================================
# EXCEL EXPORT
# ============================================================

st.divider()

st.header(
    "📊 EXPORT"
)


if estimate_items:

    summary_df = pd.DataFrame([

        {
            "Project": project,
            "Date": datetime.now().strftime(
                "%Y-%m-%d %H:%M"
            ),
            "Items": len(estimate_items),
            "Labor": total_labor,
            "Equipment": total_equipment,
            "Material": total_material,
            "Subcontract": total_subcontract,
            "Direct Cost": total_direct,
            "Overhead %": overhead_pct,
            "Overhead": total_overhead,
            "Profit %": profit_pct,
            "Profit": total_profit,
            "FINAL BID": total_bid
        }

    ])


    pricing_df = estimate_table.copy()


    bni_df = pd.DataFrame([

        {

            "Line": n,

            "CSI Code":
                x["CSI Code"],

            "Item Code":
                x["Item Code"],

            "Description":
                x["Description"],

            "Quantity":
                x["Quantity"],

            "Unit":
                x["Unit"],

            "BNI MH/Unit":
                x["Manhr/Unit"],

            "BNI Page":
                x["Page"],

            "Total Man-Hours":
                x["Total MH"]

        }

        for n, x
        in enumerate(
            estimate_items,
            start=1
        )

    ])


    labor_details = []

    equipment_details = []

    material_details = []


    for n, x in enumerate(
        estimate_items,
        start=1
    ):

        for labor in x.get(
            "Labor Details",
            []
        ):

            labor_details.append({

                "Line":
                    n,

                "Estimate Item":
                    x["Description"],

                "Position":
                    labor["Position"],

                "Workers":
                    labor["Workers"],

                "Hours":
                    labor["Hours"],

                "Rate":
                    labor["Rate"],

                "Cost":
                    labor["Cost"]

            })


        for equipment in x.get(
            "Equipment Details",
            []
        ):

            equipment_details.append({

                "Line":
                    n,

                "Estimate Item":
                    x["Description"],

                "Equipment":
                    equipment["Description"],

                "Hours":
                    equipment["Hours"],

                "Rate":
                    equipment["Rate"],

                "Cost":
                    equipment["Cost"]

            })


        for material in x.get(
            "Material Details",
            []
        ):

            material_details.append({

                "Line":
                    n,

                "Estimate Item":
                    x["Description"],

                "Material":
                    material["Description"],

                "Quantity":
                    material["Quantity"],

                "Unit":
                    material["Unit"],

                "Rate":
                    material["Rate"],

                "Cost":
                    material["Cost"]

            })


    labor_details_df = pd.DataFrame(
        labor_details
    )

    equipment_details_df = pd.DataFrame(
        equipment_details
    )

    material_details_df = pd.DataFrame(
        material_details
    )


    output = BytesIO()


    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        summary_df.to_excel(
            writer,
            index=False,
            sheet_name="Summary"
        )

        pricing_df.to_excel(
            writer,
            index=False,
            sheet_name="Estimate"
        )

        bni_df.to_excel(
            writer,
            index=False,
            sheet_name="BNI Productivity"
        )

        labor_details_df.to_excel(
            writer,
            index=False,
            sheet_name="Labor"
        )

        equipment_details_df.to_excel(
            writer,
            index=False,
            sheet_name="Equipment"
        )

        material_details_df.to_excel(
            writer,
            index=False,
            sheet_name="Materials"
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


    output.seek(0)


    st.download_button(

        "📥 DOWNLOAD FINAL ESTIMATE EXCEL",

        output.getvalue(),

        "civil_estimate_v6.xlsx",

        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

        use_container_width=True

    )


else:

    st.info(
        "Add priced items to enable Excel export."
    )


st.divider()

st.caption(
    "Version 6 — Reusable Cost Database + "
    "BNI Productivity + Multi-Item Estimating"
)
