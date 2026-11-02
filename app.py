from pathlib import Path
import re
import pandas as pd
import plotly.graph_objects as go
import streamlit as st

# ------------------------------------------------------------
# Blue Economy and Lean Construction Decision Support Tool
# Expected Excel columns:
# Impact of Marine Construction
# Blue Economy Practice
# Challenges to Implement Blue Economy Practice
# Lean Principles
# ------------------------------------------------------------

st.set_page_config(
    page_title="Blue Economy & Lean Decision Support",
    page_icon="🌊",
    layout="wide",
    initial_sidebar_state="expanded",
)

EXPECTED_COLUMNS = [
    "Impact of Marine Construction",
    "Blue Economy Practice",
    "Challenges to Implement Blue Economy Practice",
    "Lean Principles",
]

LEAN_ORDER = [
    "Define Value",
    "Map the Value Stream",
    "Create Flow",
    "Establish Pull",
    "Pursue Perfection",
]

st.markdown(
    """
    <style>
    .main-header {
        padding: 1.25rem 1.5rem;
        border-radius: 14px;
        background: linear-gradient(110deg, #073b4c, #087e8b);
        color: white;
        margin-bottom: 1rem;
    }
    .main-header h1 { color: white; margin: 0; font-size: 2rem; }
    .main-header p { color: #e8f6f7; margin: .45rem 0 0 0; }
    .stage-label {
        color: #087e8b;
        font-size: .78rem;
        font-weight: 700;
        letter-spacing: .08em;
        text-transform: uppercase;
        margin-bottom: .2rem;
    }
    div[data-testid="stButton"] button {
        width: 100%;
        min-height: 3.1rem;
        text-align: left;
        white-space: normal;
        border-radius: 9px;
        border: 1px solid #b8d7dc;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="main-header">
      <h1>Blue Economy & Lean Construction</h1>
      <p>Interactive decision-support tool for connecting marine construction impacts,
      Blue Economy practices, implementation challenges and Lean principles.</p>
    </div>
    """,
    unsafe_allow_html=True,
)


def normalise_column_name(value):
    return re.sub(r"\s+", " ", str(value).strip()).casefold()


def find_excel_file():
    candidates = [
        p for p in Path(".").glob("*.xlsx")
        if not p.name.startswith("~$")
    ]
    # Prefer a workbook whose name suggests the subject.
    candidates.sort(key=lambda p: (
        "blue" not in p.name.casefold(),
        "econom" not in p.name.casefold(),
        p.name.casefold(),
    ))
    return candidates[0] if candidates else None


def split_values(value):
    """Split cells containing multiple mapped values on common separators."""
    if pd.isna(value):
        return []
    text = str(value).strip()
    if not text or text.casefold() in {"nan", "none", "n/a", "na", "-"}:
        return []
    parts = re.split(r"\s*(?:\n|;|\|)\s*", text)
    # Commas are also accepted as separators for compact multi-value cells.
    expanded = []
    for part in parts:
        expanded.extend(re.split(r"\s*,\s*", part))
    return [p.strip() for p in expanded if p.strip()]


def load_workbook(uploaded_file):
    if uploaded_file is not None:
        return pd.read_excel(uploaded_file, engine="openpyxl")
    workbook = find_excel_file()
    if workbook is not None:
        return pd.read_excel(workbook, engine="openpyxl")
    return None


def prepare_data(raw):
    if raw is None:
        return None, None

    raw.columns = [str(c).strip() for c in raw.columns]
    lookup = {normalise_column_name(c): c for c in raw.columns}
    missing = [c for c in EXPECTED_COLUMNS if normalise_column_name(c) not in lookup]
    if missing:
        return None, (
            "The Excel workbook is missing the following required column(s): "
            + ", ".join(missing)
            + ". Please use the exact column headings shown in the instructions."
        )

    data = raw[[lookup[normalise_column_name(c)] for c in EXPECTED_COLUMNS]].copy()
    data.columns = EXPECTED_COLUMNS

    for col in EXPECTED_COLUMNS:
        data[col] = data[col].fillna("").astype(str).str.strip()

    data = data[
        (data[EXPECTED_COLUMNS[0]] != "")
        & (data[EXPECTED_COLUMNS[1]] != "")
        & (data[EXPECTED_COLUMNS[2]] != "")
    ].copy()

    # Expand rows where a cell contains multiple mapped practices/challenges/
    # Lean principles. Each resulting row represents one relationship.
    expanded_rows = []
    for _, row in data.iterrows():
        impacts = split_values(row[EXPECTED_COLUMNS[0]])
        practices = split_values(row[EXPECTED_COLUMNS[1]])
        challenges = split_values(row[EXPECTED_COLUMNS[2]])
        principles = split_values(row[EXPECTED_COLUMNS[3]])

        # A row should ideally contain one value in each column. If a cell
        # contains several values, all combinations are treated as mappings.
        if not principles:
            principles = [""]
        for impact in impacts:
            for practice in practices:
                for challenge in challenges:
                    for principle in principles:
                        expanded_rows.append({
                            EXPECTED_COLUMNS[0]: impact,
                            EXPECTED_COLUMNS[1]: practice,
                            EXPECTED_COLUMNS[2]: challenge,
                            EXPECTED_COLUMNS[3]: principle,
                        })

    result = pd.DataFrame(expanded_rows, columns=EXPECTED_COLUMNS).drop_duplicates()
    if result.empty:
        return None, "No complete mapping rows were found in the workbook."
    return result, None


def unique_values(frame, column):
    return sorted(frame[column].dropna().astype(str).unique().tolist(), key=str.casefold)


def make_sankey(frame, selected_impact=None, selected_practice=None, selected_challenge=None):
    """Create a four-stage Sankey diagram from the workbook's mapping rows."""
    stages = EXPECTED_COLUMNS
    links = []
    for i in range(3):
        pair = frame[[stages[i], stages[i + 1]]].copy()
        pair = pair[(pair[stages[i]] != "") & (pair[stages[i + 1]] != "")]
        counts = pair.value_counts().reset_index(name="value")
        for _, r in counts.iterrows():
            links.append((i, str(r[stages[i]]), i + 1, str(r[stages[i + 1]]), int(r["value"])))

    labels_by_stage = []
    for col in stages:
        labels_by_stage.append(unique_values(frame[frame[col] != ""], col))

    labels = []
    node_x = []
    node_y = []
    node_lookup = {}
    for stage_index, stage_labels in enumerate(labels_by_stage):
        for j, label in enumerate(stage_labels):
            node_lookup[(stage_index, label)] = len(labels)
            labels.append(label)
            node_x.append([0.02, 0.35, 0.68, 0.98][stage_index])
            node_y.append((j + 1) / (len(stage_labels) + 1))

    sources, targets, values, link_colours = [], [], [], []
    for s_stage, s_label, t_stage, t_label, value in links:
        sources.append(node_lookup[(s_stage, s_label)])
        targets.append(node_lookup[(t_stage, t_label)])
        values.append(value)
        # Highlight links that are consistent with the current selection.
        is_selected = True
        if selected_impact and s_stage == 0 and s_label != selected_impact:
            is_selected = False
        if selected_practice and t_stage == 1 and t_label != selected_practice:
            is_selected = False
        if selected_challenge and t_stage == 2 and t_label != selected_challenge:
            is_selected = False
        link_colours.append("rgba(8,126,139,0.58)" if is_selected else "rgba(160,170,175,0.16)")

    fig = go.Figure(data=[go.Sankey(
        arrangement="snap",
        node=dict(
            pad=17,
            thickness=17,
            line=dict(color="rgba(70,90,100,0.35)", width=0.5),
            label=labels,
            x=node_x,
            y=node_y,
            color=["#087e8b" if i in (0, 1) else "#5b9ea6" if i in (2, 3) else "#8bbfc3"
                   for i in range(len(labels))],
            hovertemplate="%{label}<extra></extra>",
        ),
        link=dict(
            source=sources,
            target=targets,
            value=values,
            color=link_colours,
            hovertemplate="%{source.label} → %{target.label}<br>Mapping count: %{value}<extra></extra>",
        ),
    )])
    fig.update_layout(
        height=max(500, min(850, 190 + 24 * max(map(len, labels_by_stage)))),
        margin=dict(l=12, r=12, t=12, b=12),
        font=dict(size=11, color="#20343a"),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)",
    )
    return fig


with st.sidebar:
    st.header("Data source")
    uploaded = st.file_uploader(
        "Upload your Excel workbook",
        type=["xlsx"],
        help="Alternatively, place the Excel workbook in the same GitHub repository folder as app.py.",
    )
    st.divider()
    st.caption("Required Excel columns")
    for column in EXPECTED_COLUMNS:
        st.write("• " + column)
    st.divider()
    st.caption("The workbook is read-only. Uploading it does not change the original file.")

try:
    raw_data = load_workbook(uploaded)
except Exception as exc:
    st.error(f"The Excel workbook could not be opened: {exc}")
    st.stop()

if raw_data is None:
    st.info(
        "Upload your Excel workbook using the panel on the left, or add the .xlsx "
        "file to the GitHub repository alongside app.py."
    )
    st.stop()

data, error = prepare_data(raw_data)
if error:
    st.error(error)
    st.write("Columns detected in your workbook:", list(raw_data.columns))
    st.stop()

st.caption(f"Loaded {len(data):,} mapping record(s) from the Excel workbook.")

# Initialise selection state.
for key in ["impact", "practice", "challenge"]:
    if key not in st.session_state:
        st.session_state[key] = None

impacts = unique_values(data, EXPECTED_COLUMNS[0])
if st.session_state.impact not in impacts:
    st.session_state.impact = None

st.markdown("### Explore the relationships")
st.write("Select an impact, then a related practice and challenge to reveal the corresponding Lean principles.")

# Stage 1
st.markdown('<div class="stage-label">Stage 1</div>', unsafe_allow_html=True)
st.subheader("1. Impact of Marine Construction")
impact_cols = st.columns(2)
for idx, impact in enumerate(impacts):
    with impact_cols[idx % 2]:
        if st.button(
            ("✓  " if st.session_state.impact == impact else "") + impact,
            key=f"impact_{idx}",
            type="primary" if st.session_state.impact == impact else "secondary",
        ):
            st.session_state.impact = impact
            st.session_state.practice = None
            st.session_state.challenge = None
            st.rerun()

if st.session_state.impact is None:
    st.info("Select an impact to view its associated Blue Economy practices.")
    st.markdown("### Overall mapping")
    st.plotly_chart(make_sankey(data), use_container_width=True, config={"displayModeBar": False})
    st.stop()

impact_data = data[data[EXPECTED_COLUMNS[0]] == st.session_state.impact]
st.success(f"Selected impact: {st.session_state.impact}")

# Stage 2
st.markdown('<div class="stage-label">Stage 2</div>', unsafe_allow_html=True)
st.subheader("2. Blue Economy Practice")
practices = unique_values(impact_data, EXPECTED_COLUMNS[1])
practice_cols = st.columns(2)
for idx, practice in enumerate(practices):
    with practice_cols[idx % 2]:
        if st.button(
            ("✓  " if st.session_state.practice == practice else "") + practice,
            key=f"practice_{idx}",
            type="primary" if st.session_state.practice == practice else "secondary",
        ):
            st.session_state.practice = practice
            st.session_state.challenge = None
            st.rerun()

if st.session_state.practice not in practices:
    st.session_state.practice = None
    st.info("Select a Blue Economy practice to view its associated implementation challenges.")
    st.markdown("### Mapping for the selected impact")
    st.plotly_chart(
        make_sankey(impact_data, selected_impact=st.session_state.impact),
        use_container_width=True,
        config={"displayModeBar": False},
    )
    st.stop()

practice_data = impact_data[impact_data[EXPECTED_COLUMNS[1]] == st.session_state.practice]
st.success(f"Selected practice: {st.session_state.practice}")

# Stage 3
st.markdown('<div class="stage-label">Stage 3</div>', unsafe_allow_html=True)
st.subheader("3. Challenges to Implement Blue Economy Practice")
challenges = unique_values(practice_data, EXPECTED_COLUMNS[2])
challenge_cols = st.columns(2)
for idx, challenge in enumerate(challenges):
    with challenge_cols[idx % 2]:
        if st.button(
            ("✓  " if st.session_state.challenge == challenge else "") + challenge,
            key=f"challenge_{idx}",
            type="primary" if st.session_state.challenge == challenge else "secondary",
        ):
            st.session_state.challenge = challenge
            st.rerun()

if st.session_state.challenge not in challenges:
    st.session_state.challenge = None
    st.info("Select a challenge to view the Lean principles linked to it.")
    st.markdown("### Mapping for the selected practice")
    st.plotly_chart(
        make_sankey(
            practice_data,
            selected_impact=st.session_state.impact,
            selected_practice=st.session_state.practice,
        ),
        use_container_width=True,
        config={"displayModeBar": False},
    )
    st.stop()

challenge_data = practice_data[practice_data[EXPECTED_COLUMNS[2]] == st.session_state.challenge]
mapped_principles = {
    p.casefold()
    for value in challenge_data[EXPECTED_COLUMNS[3]].tolist()
    for p in split_values(value)
}

# Stage 4
st.markdown('<div class="stage-label">Stage 4</div>', unsafe_allow_html=True)
st.subheader("4. Lean Principles")
st.write(f"**Selected challenge:** {st.session_state.challenge}")
st.caption("All five Lean principles are displayed below in the prescribed order. Principles mapped to the selected challenge are highlighted.")

for number, principle in enumerate(LEAN_ORDER, start=1):
    is_mapped = principle.casefold() in mapped_principles
    col1, col2 = st.columns([0.12, 0.88])
    with col1:
        st.markdown(f"### {number}")
    with col2:
        if is_mapped:
            st.success(f"**{principle}**  ·  Mapped to this challenge")
        else:
            st.markdown(f"**{principle}**")
            st.caption("No direct mapping recorded for this challenge in the workbook.")

# Also display any workbook principle wording that does not match the five standard labels.
other_principles = sorted(
    {p for value in challenge_data[EXPECTED_COLUMNS[3]].tolist() for p in split_values(value)
     if p.casefold() not in {x.casefold() for x in LEAN_ORDER}},
    key=str.casefold,
)
if other_principles:
    with st.expander("Additional Lean principle labels found in the Excel workbook"):
        st.write(other_principles)

st.markdown("### Sankey mapping")
st.plotly_chart(
    make_sankey(
        data,
        selected_impact=st.session_state.impact,
        selected_practice=st.session_state.practice,
        selected_challenge=st.session_state.challenge,
    ),
    use_container_width=True,
    config={"displayModeBar": False},
)

with st.expander("View the records behind this selection"):
    st.dataframe(challenge_data, use_container_width=True, hide_index=True)

st.caption("The relationships and mapped Lean principles are based on the records in the uploaded Excel workbook.")
