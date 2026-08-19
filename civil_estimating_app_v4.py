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
st.caption("Version 4 — BNI Productivity + Estimate Builder")

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
# EXCEL EXPORT
# ============================================================

def make_excel(project, estimate_items, workday_hours):

    output = BytesIO()

    rows = []

    for number, item in enumerate(
        estimate_items,
        start=1
    ):

        rows.append({

            "Line": number,

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

            "Foreman":
                item["Foreman"],

            "Laborer":
                item["Laborer"],

            "Equipment Operator":
                item["Equipment Operator"],

            "Total Crew":
                item["Total Crew"],

            "Total Man-Hours":
                item["Total MH"],

            "Crew Hours":
                item["Crew Hours"],

            "Working Days":
                item["Days"]
        })

    estimate_df = pd.DataFrame(rows)

    summary_df = pd.DataFrame({

        "Project": [project],

        "Date": [
            datetime.now().strftime(
                "%Y-%m-%d %H:%M"
            )
        ],

        "Number of Items": [
            len(estimate_items)
        ],

        "Total Man-Hours": [
            sum(
                x["Total MH"]
                for x in estimate_items
            )
        ],

        "Total Crew Hours": [
            sum(
                x["Crew Hours"]
                for x in estimate_items
            )
        ],

        "Total Working Days": [
            sum(
                x["Days"]
                for x in estimate_items
            )
        ]
    })

    bni_df = pd.DataFrame([

        {

            "Line": number,

            "CSI Code":
                item["CSI Code"],

            "Item Code":
                item["Item Code"],

            "Description":
                item["Description"],

            "Unit":
                item["Unit"],

            "Quantity":
                item["Quantity"],

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

    crew_df = pd.DataFrame([

        {

            "Line": number,

            "Description":
                item["Description"],

            "Foreman":
                item["Foreman"],

            "Laborer":
                item["Laborer"],

            "Equipment Operator":
                item["Equipment Operator"],

            "Total Crew":
                item["Total Crew"]

        }

        for number, item
        in enumerate(
            estimate_items,
            start=1
        )

    ])

    notes_df = pd.DataFrame({

        "Notes": [

            "BNI productivity source used.",

            "Total Man-Hours = Quantity × BNI Manhr/Unit.",

            "Crew Hours = Total Man-Hours ÷ Total Crew.",

            "Working Days = Crew Hours ÷ Hours per Workday.",

            "Verify BNI productivity against the applicable licensed BNI Costbook before bidding."

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

        bni_df.to_excel(
            writer,
            index=False,
            sheet_name="BNI Productivity"
        )

        crew_df.to_excel(
            writer,
            index=False,
            sheet_name="Crew"
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

    st.header("Project")

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

    st.header("BNI Database")

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
# SEARCH BNI
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
]


matches = matches.head(200)


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
    "4. Crew"
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
# CALCULATE
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
            float(
                item["Manhr/Unit"]
            )

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
# CURRENT RESULT
# ============================================================

if "current_item" in st.session_state:

    current = (
        st.session_state.current_item
    )


    st.divider()

    st.subheader(
        "Result"
    )


    r1, r2, r3 = st.columns(3)


    r1.metric(
        "Total Man-Hours",
        f'{current["Total MH"]:,.2f} MH'
    )


    r2.metric(
        "Crew Hours",
        f'{current["Crew Hours"]:,.2f} hr'
    )


    r3.metric(
        "Working Days",
        f'{current["Days"]:,.2f}'
    )


    st.code(

        f'Man-Hours = '
        f'{current["Quantity"]:,.2f} × '
        f'{current["Manhr/Unit"]:.4f} = '
        f'{current["Total MH"]:,.2f} MH\n\n'

        f'Crew Hours = '
        f'{current["Total MH"]:,.2f} ÷ '
        f'{current["Total Crew"]} workers = '
        f'{current["Crew Hours"]:,.2f} hr\n\n'

        f'Working Days = '
        f'{current["Crew Hours"]:,.2f} ÷ '
        f'{workday_hours:.2f} hr/day = '
        f'{current["Days"]:,.2f} days'

    )


    if st.button(

        "➕ ADD ITEM TO ESTIMATE",

        use_container_width=True

    ):

        st.session_state.estimate_items.append(
            current.copy()
        )

        st.success(
            f'{current["Description"]} '
            'was added to the estimate.'
        )


# ============================================================
# ESTIMATE BUILDER
# ============================================================

st.divider()

st.subheader(
    "5. Estimate Builder"
)


estimate_items = (
    st.session_state.estimate_items
)


if estimate_items:

    table = pd.DataFrame([

        {

            "Line":
                number,

            "Item Code":
                x["Item Code"],

            "Description":
                x["Description"],

            "Qty":
                x["Quantity"],

            "Unit":
                x["Unit"],

            "BNI MH/Unit":
                x["Manhr/Unit"],

            "Total MH":
                x["Total MH"],

            "Crew Hrs":
                x["Crew Hours"],

            "Days":
                x["Days"]

        }

        for number, x
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


    total_mh = sum(
        x["Total MH"]
        for x in estimate_items
    )


    total_hours = sum(
        x["Crew Hours"]
        for x in estimate_items
    )


    total_days = sum(
        x["Days"]
        for x in estimate_items
    )


    c1, c2, c3 = st.columns(3)


    c1.metric(
        "Items",
        len(estimate_items)
    )


    c2.metric(
        "Total Man-Hours",
        f"{total_mh:,.2f}"
    )


    c3.metric(
        "Total Crew Hours",
        f"{total_hours:,.2f}"
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

        "Your estimate is empty. "
        "Calculate an item above and "
        "click ADD ITEM TO ESTIMATE."

    )


# ============================================================
# EXPORT
# ============================================================

st.divider()

st.subheader(
    "6. Export"
)


if estimate_items:

    excel = make_excel(

        project,

        estimate_items,

        workday_hours

    )


    st.download_button(

        "📊 DOWNLOAD EXCEL ESTIMATE",

        excel,

        "civil_estimate.xlsx",

        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",

        use_container_width=True

    )


else:

    st.info(
        "Add items to the estimate before exporting."
    )


st.divider()

st.caption(
    "Next: labor pricing + equipment pricing + material pricing + "
    "overhead + profit + final bid price."
)
