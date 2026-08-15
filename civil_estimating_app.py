import streamlit as st

st.set_page_config(
    page_title="Civil Estimating Software",
    page_icon="🏗️",
    layout="centered"
)

st.title("🏗️ Civil Estimating Software")
st.caption("Version 1 — Productivity & Man-Hour Calculator")

st.divider()

# Item selection
item = st.selectbox(
    "Item",
    [
        "Asphalt Paving",
        "Concrete Paving",
        "Excavation",
        "Gravel Base",
        "Pipe Installation"
    ]
)

# Quantity and unit
col1, col2 = st.columns([2, 1])

with col1:
    quantity = st.number_input(
        "Quantity",
        min_value=0.0,
        value=500.0,
        step=1.0
    )

with col2:
    unit = st.selectbox(
        "Unit",
        ["CY", "TON", "LF", "SF", "SY", "EA"]
    )

# Productivity source
source = st.selectbox(
    "Productivity Source",
    ["BNI General Construction Costbook", "Custom"]
)

st.info(
    "Enter the productivity rate from your licensed BNI Costbook. "
    "The example rate below is only a demonstration and is not presented "
    "as an actual BNI rate."
)

# Demonstration productivity rate
default_rate = 2.50 if item == "Asphalt Paving" and unit == "CY" else 1.00

productivity = st.number_input(
    f"Productivity Rate ({unit}/Man-Hour)",
    min_value=0.0001,
    value=default_rate,
    step=0.10,
    format="%.4f"
)

st.divider()

if st.button("CALCULATE", type="primary", use_container_width=True):

    man_hours = quantity / productivity

    st.subheader("Result")

    c1, c2 = st.columns(2)

    with c1:
        st.metric("Quantity", f"{quantity:,.2f} {unit}")

    with c2:
        st.metric("Productivity", f"{productivity:,.4f} {unit}/MH")

    st.success(f"Estimated Man-Hours: **{man_hours:,.2f} MH**")

    st.write("### Calculation")

    st.code(
        f"Man-Hours = Quantity ÷ Productivity\n\n"
        f"          = {quantity:,.2f} {unit} ÷ "
        f"{productivity:,.4f} {unit}/MH\n\n"
        f"          = {man_hours:,.2f} MH"
    )

    st.caption(
        f"Productivity source: {source}. "
        "Verify the applicable productivity rate against your licensed costbook."
    )

st.divider()

st.caption(
    "Future versions can add labor, equipment, material pricing, "
    "vendor quotes, project locations, markups, and Excel/PDF export."
)
