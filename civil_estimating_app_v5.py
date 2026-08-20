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
st.caption("Version 5 — BNI Productivity + Cost & Pricing Engine")

DB = Path(__file__).parent / "construction_costs_pages_16_to_36.xlsx"


# ============================================================
# DATABASE
# ============================================================

@st.cache_data
def load_db(source):

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

    missing = [
        c for c in required
        if c not in df.columns
    ]

    if missing:
        raise ValueError(
            "Missing columns: " +
            ", ".join(missing)
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
# EXCEL EXPORT
# ============================================================

def make_excel(
    project,
    estimate_items,
    workday_hours,
    overhead_pct,
    profit_pct
):

    output = BytesIO()

    estimate_rows = []

    for number, item in enumerate(
        estimate_items,
        start=1
    ):

        estimate_rows.append({

            "Line":
                number,

            "CSI Code":
                item["CSI Code"],

            "Item Code":
                item["Item Code"],

            "Description":
                item["Description"],

            "Quantity":
                item["Quantity"],

            "Unit":
                item["Unit"],

            "BNI MH/Unit":
                item["Manhr/Unit"],

            "BNI Page":
                item["Page"],

            "Total Man-Hours":
                item["Total MH"],

            "Labor Cost":
                item["Labor Cost"],

            "Equipment Cost":
                item["Equipment Cost"],

            "Material Cost":
                item["Material Cost"],

            "Subcontract Cost":
                item["Subcontract Cost"],

            "Direct Cost":
                item["Direct Cost"],

            "Overhead":
                item["Overhead"],

            "Profit":
                item["Profit"],

            "Final Bid":
                item["Final Bid"]

        })

    estimate_df = pd.DataFrame(
        estimate_rows
    )

    total_direct = sum(
        x["Direct Cost"]
        for x in estimate_items
    )

    total_overhead = sum(
        x["Overhead"]
        for x in estimate_items
    )

    total_profit = sum(
        x["Profit"]
        for x in estimate_items
    )

    total_bid = sum(
        x["Final Bid"]
        for x in estimate_items
    )

    summary_df = pd.DataFrame({

        "Field": [

            "Project",

            "Date",

            "Number of Items",

            "Total Direct Cost",

            "Overhead %",

            "Total Overhead",

            "Profit %",

            "Total Profit",

            "FINAL BID"

        ],

        "Value": [

            project,

            datetime.now().strftime(
                "%Y-%m-%d %H:%M"
            ),

            len(estimate_items),

            total_direct,

            overhead_pct,

            total_overhead,

            profit_pct,

            total_profit,

            total_bid

        ]

    })

    pricing_df = pd.DataFrame([

        {

            "Line":
                number,

            "Description":
                item["Description"],

            "Quantity":
                item["Quantity"],

            "Unit":
                item["Unit"],

            "Labor Cost":
                item["Labor Cost"],

            "Equipment Cost":
                item["Equipment Cost"],

            "Material Cost":
                item["Material Cost"],

            "Subcontract Cost":
                item["Subcontract Cost"],

            "Direct Cost":
                item["Direct Cost"],

            "Overhead":
                item["Overhead"],

            "Profit":
                item["Profit"],

            "Final Bid":
                item["Final Bid"]

        }

        for number, item
        in enumerate(
            estimate_items,
            start=1
        )

    ])

    bni_df = pd.DataFrame([

        {

            "Line":
                number,

            "CSI Code":
                item["CSI Code"],

            "Item Code":
                item["Item Code"],

            "Description":
                item["Description"],

            "Quantity":
                item["Quantity"],

            "Unit":
                item["Unit"],

            "BNI MH/Unit":
                item["Manhr/Unit"],

            "BNI Page":
                item["Page"]

        }

        for number, item
        in enumerate(
            estimate_items,
            start=1
        )

    ])

    notes_df = pd.DataFrame({

        "Notes": [

            "BNI productivity is used as the productivity source.",

            "Labor, equipment, material, and subcontract rates are user-entered.",

            "Direct Cost = Labor + Equipment + Material + Subcontract.",

            "Overhead = Direct Cost × Overhead %.",

            "Profit is calculated after overhead.",

            "Final Bid = Direct Cost + Overhead + Profit.",

            "Verify all productivity rates and prices before bidding."

        ]

    })

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        summary_df.to_excel(
            writer,
            index=False,
            sheet_name="Summary"
        )

        estimate_df.to_excel(
            writer,
            index=False,
            sheet_name="Estimate"
        )

        pricing_df.to_excel(
            writer,
            index=False,
            sheet_name="Pricing"
        )

        bni_df.to_excel(
            writer,
            index=False,
            sheet_name="BNI Productivity"
        )

        notes_df.to_excel(
            writer,
            index=False,
            sheet_name="Notes"
        )

    output.seek(0)

    return output.getvalue()


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
# LOAD DATABASE
# ============================================================

try:

    if uploaded:

        df = load_db(uploaded)

    elif DB.exists():

        df = load_db(DB)

    else:

        st.warning(
            "Upload the BNI Excel database."
        )

        st.stop()

except Exception as e:

    st.error(
        f"Could not load database: {e}"
    )

    st.stop()


st.success(
    f"Database loaded: {len(df):,} rows"
)


# ============================================================
# ESTIMATE MEMORY
# ============================================================

if "estimate_items" not in st.session_state:

    st.session_state.estimate_items = []


# ============================================================
# SEARCH
# ============================================================

st.divider()

st.subheader(
    "1. Find a BNI Item"
)

search = st.text_input(

    "Search description, CSI code, item code, or unit",

    placeholder=
    "Example: asphalt, concrete, pipe"

)

matches = df.copy()


if search.strip():

    q = search.lower().strip()

    matches = matches[

        matches["Description"]
        .str.lower()
        .str.contains(
            q,
            na=False
        )

        |

        matches["CSI Code"]
        .str.lower()
        .str.contains(
            q,
            na=False
        )

        |

        matches["Item Code"]
        .str.lower()
        .str.contains(
            q,
            na=False
        )

        |

        matches["Unit"]
        .str.lower()
        .str.contains(
            q,
            na=False
        )

    ]


matches = matches[
    matches["Manhr/Unit"].notna()
].head(200)


if matches.empty:

    st.info(
        "No matching BNI items found."
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

    format_func=
    lambda x: options[x]

)


item = df.loc[selected]


# ============================================================
# QUANTITY
# ============================================================

st.divider()

st.subheader(
    "2. Quantity"
)

quantity = st.number_input(

    f'Quantity ({item["Unit"]})',

    min_value=0.0,

    value=500.0,

    step=1.0

)


# ============================================================
# PRODUCTIVITY
# ============================================================

st.subheader(
    "3. BNI Productivity"
)

c1, c2, c3 = st.columns(3)

c1.metric(
    "Manhr / Unit",
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
    f'**Description:** '
    f'{item["Description"]}'
)

st.caption(
    f'CSI: {item["CSI Code"]} '
    f'• Item Code: {item["Item Code"]}'
)


# ============================================================
# CREW
# ============================================================

st.divider()

st.subheader(
    "4. CREW"
)

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
# COST INPUTS
# ============================================================

if "current_item" in st.session_state:

    current = (
        st.session_state.current_item
    )

    st.divider()

    st.subheader(
        "5. COST & PRICING"
    )

    st.info(
        "Enter your own current rates. "
        "BNI productivity and pricing are kept separate."
    )

    st.markdown(
        "### Labor Pricing"
    )

    l1, l2, l3 = st.columns(3)

    with l1:

        foreman_rate = st.number_input(
            "Foreman $/HR",
            min_value=0.0,
            value=55.0,
            step=1.0
        )

    with l2:

        laborer_rate = st.number_input(
            "Laborer $/HR",
            min_value=0.0,
            value=42.0,
            step=1.0
        )

    with l3:

        operator_rate = st.number_input(
            "Equipment Operator $/HR",
            min_value=0.0,
            value=48.0,
            step=1.0
        )

    labor_cost = (

        current["Foreman"] *
        current["Crew Hours"] *
        foreman_rate

        +

        current["Laborer"] *
        current["Crew Hours"] *
        laborer_rate

        +

        current["Equipment Operator"] *
        current["Crew Hours"] *
        operator_rate

    )

    st.metric(
        "Calculated Labor Cost",
        f"${labor_cost:,.2f}"
    )


    st.markdown(
        "### Equipment Pricing"
    )

    equipment_hours = st.number_input(
        "Equipment Hours",
        min_value=0.0,
        value=0.0,
        step=1.0
    )

    equipment_rate = st.number_input(
        "Equipment Rate $/HR",
        min_value=0.0,
        value=0.0,
        step=5.0
    )

    equipment_cost = (
        equipment_hours *
        equipment_rate
    )

    st.metric(
        "Equipment Cost",
        f"${equipment_cost:,.2f}"
    )


    st.markdown(
        "### Material Pricing"
    )

    material_quantity = st.number_input(
        f"Material Quantity ({current['Unit']})",
        min_value=0.0,
        value=0.0,
        step=1.0
    )

    material_unit_price = st.number_input(
        f"Material Price $/{current['Unit']}",
        min_value=0.0,
        value=0.0,
        step=1.0
    )

    material_cost = (
        material_quantity *
        material_unit_price
    )

    st.metric(
        "Material Cost",
        f"${material_cost:,.2f}"
    )


    st.markdown(
        "### Subcontract"
    )

    subcontract_cost = st.number_input(
        "Subcontract Cost",
        min_value=0.0,
        value=0.0,
        step=100.0
    )


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

    st.markdown(
        "### PRICE SUMMARY"
    )

    p1, p2, p3, p4 = st.columns(4)

    p1.metric(
        "Labor",
        f"${labor_cost:,.2f}"
    )

    p2.metric(
        "Equipment",
        f"${equipment_cost:,.2f}"
    )

    p3.metric(
        "Material",
        f"${material_cost:,.2f}"
    )

    p4.metric(
        "Subcontract",
        f"${subcontract_cost:,.2f}"
    )


    st.divider()

    s1, s2, s3 = st.columns(3)

    s1.metric(
        "DIRECT COST",
        f"${direct_cost:,.2f}"
    )

    s2.metric(
        f"OVERHEAD ({overhead_pct:.1f}%)",
        f"${overhead_amount:,.2f}"
    )

    s3.metric(
        f"PROFIT ({profit_pct:.1f}%)",
        f"${profit_amount:,.2f}"
    )


    st.success(
        f"FINAL BID PRICE: ${final_bid:,.2f}"
    )


    # Store pricing temporarily for adding to estimate
    current["Labor Cost"] = labor_cost
    current["Equipment Cost"] = equipment_cost
    current["Material Cost"] = material_cost
    current["Subcontract Cost"] = subcontract_cost
    current["Direct Cost"] = direct_cost
    current["Overhead"] = overhead_amount
    current["Profit"] = profit_amount
    current["Final Bid"] = final_bid


    if st.button(
        "➕ ADD PRICED ITEM TO ESTIMATE",
        use_container_width=True
    ):

        st.session_state.estimate_items.append(
            current.copy()
        )

        st.success(
            f'{current["Description"]} '
            'was added to the priced estimate.'
        )


# ============================================================
# ESTIMATE BUILDER
# ============================================================

st.divider()

st.subheader(
    "6. ESTIMATE BUILDER"
)

estimate_items = (
    st.session_state.estimate_items
)


if estimate_items:

    table = pd.DataFrame([

        {

            "Line":
                n,

            "Item Code":
                x["Item Code"],

            "Description":
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

            "Direct Cost":
                x.get("Direct Cost", 0),

            "Overhead":
                x.get("Overhead", 0),

            "Profit":
                x.get("Profit", 0),

            "Final Bid":
                x.get("Final Bid", 0)

        }

        for n, x
        in enumerate(
            estimate_items,
            start=1
        )

    ])


    st.dataframe(
        table,
        use_container_width=True,
        hide_index=True
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

    d.metric(
        "FINAL BID",
        f"${total_bid:,.2f}"
    )


    if st.button(
        "🗑️ CLEAR ESTIMATE",
        use_container_width=True
    ):

        st.session_state.estimate_items = []

        st.session_state.pop(
            "current_item",
            None
        )

        st.rerun()


else:

    st.info(
        "Your estimate is empty."
    )


# ============================================================
# EXPORT
# ============================================================

st.divider()

st.subheader(
    "7. EXPORT"
)

if estimate_items:

    excel = make_excel(
        project,
        estimate_items,
        workday_hours,
        overhead_pct,
        profit_pct
    )

    st.download_button(

        "📊 DOWNLOAD FINAL EXCEL ESTIMATE",

        excel,

        "civil_final_estimate.xlsx",

        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

        use_container_width=True

    )

else:

    st.info(
        "Add a priced item to the estimate before exporting."
    )


st.divider()

st.caption(
    "Version 5 — BNI Productivity + Labor + Equipment + "
    "Material + Subcontract + Overhead + Profit"
)
