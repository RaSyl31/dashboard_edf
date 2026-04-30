from textwrap import dedent
import pandas as pd
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO

param_url = "https://docs.google.com/spreadsheets/d/e/2PACX-1vSPtJ4Ipch22BEPIjkL4U466elod-K3yegtgOiYKAcaXjMmcqpsM6g8zuA2F5VWWaZdrXavEIP3AbY2/pub?output=csv"

OBJECTIF_JOURNALIER_PAR_AGENT = 3

st.set_page_config(page_title="Dashboard Hellopro", layout="wide")

st.markdown("""
<style>
.block-container { padding-top: 1.5rem; background-color: #dfe6f5; }

.header-card {
    background: linear-gradient(90deg, #ffffff 0%, #eef3ff 100%);
    padding: 24px;
    border-radius: 18px;
    border: 1px solid #E3E8F0;
    box-shadow: 0px 4px 14px rgba(0,01,0,0.05);
    margin-bottom: 15px;
}

.title {
    font-size: 35px;
    font-weight: 900;
    color: #3c5bff;
    text-align: center;
}

.subtitle {
    font-size: 20px;
    color: #2842d4;
    margin-top: 10px;
    text-align: center;
}

.section-title {
    font-size: 22px;
    font-weight: 800;
    margin-top: 24px;
    margin-bottom: 16px;
    color: #3c5bff;
}

.kpi-card {
    background-color: white;
    padding: 12px 16px;
    border-radius: 10px;
    border: 1px solid #E5EAF2;
    box-shadow: 0px 5px 8px rgba(0,0,0,0.04);
    text-align: center;
    min-height: 60px;
}

.kpi-label {
    font-size: 15px;
    font-weight: 650;
    color: #374151;
    text-align: center;
}

.kpi-value {
    font-size: 22px;
    font-weight: 800;
    margin-top: 4px;
    text-align: center;
}

.blue { color: #2563EB; }
.green { color: #16A34A; }
.lightgreen { color: #22C55E; }
.orange { color: #F97316; }
.red { color: #DC2626; }
.black { color: #111111; }

[data-testid="stSidebar"] {
    background-color: #111827;
}

[data-testid="stSidebar"] h1,
[data-testid="stSidebar"] h2,
[data-testid="stSidebar"] h3,
[data-testid="stSidebar"] label,
[data-testid="stSidebar"] p {
    color: white !important;
}

[data-baseweb="select"] * {
    color: #111111 !important;
}

.selection-card {
    background-color: #1F2937;
    padding: 14px;
    border-radius: 14px;
    border: 1px solid #374151;
    margin-top: 10px;
    margin-bottom: 18px;
}

.selection-title {
    color: #93C5FD;
    font-size: 13px;
    font-weight: 700;
    margin-bottom: 8px;
}

.selection-value {
    color: white;
    font-size: 15px;
    font-weight: 800;
    line-height: 1.7;
}

.alert-card {
    background-color: #FFF7ED;
    border: 1px solid #FDBA74;
    padding: 14px;
    border-radius: 14px;
    color: #9A3412;
    font-weight: 700;
    margin-bottom: 10px;
}

.ok-card {
    background-color: #ECFDF5;
    border: 1px solid #86EFAC;
    padding: 15px;
    border-radius: 40px;
    color: #166534;
    font-weight: 700;
}

.row-spacing {
    margin-bottom: 10px;
}

section[data-testid="stSidebar"] button {
    background-color: #2563EB !important;
    color: white !important;
    font-weight: 700;
    border-radius: 10px;
    padding: 10px;
}

.kpi-flow-card {
    background: white;
    border-radius: 12px;
    padding: 14px 18px;
    border: 1px solid #E5EAF2;
    box-shadow: 0px 3px 8px rgba(0,0,0,0.05);
    display: flex;
    align-items: center;
    gap: 14px;
    min-height: 85px;
}

.kpi-icon {
    font-size: 25px;
    color: #64748B;
}

.kpi-flow-label {
    font-size: 15px;
    font-weight: 700;
    color: #182466;
}

.kpi-flow-value {
    font-size: 28px;
    font-weight: 800;
    color: #182466;
    line-height: 1.1;

}

.gauge-card {
    background: white;
    border-radius: 15px;
    padding: 10px;
    border: 1px solid #E5EAF2;
    box-shadow: 0px 3px 8px rgba(0,0,0,0.05);
    text-align: center;
}

.block-container {
    max-width: 100% !important;
    width: 100% !important;
    padding-left: 2rem !important;
    padding-right: 2rem !important;
}

.agent-table th, .agent-table td {
    white-space: nowrap;
}

.table-scroll {
    width: 100%;
    max-width: 100%;
    overflow-x: auto;
    overflow-y: hidden;
    padding-bottom: 8px;
}

.agent-table {
    width: max-content;
    min-width: 100%;
}

.agent-table td:first-child {
    background-color: #d8deff !important;  /* 👉 couleur ici */
    font-weight: 700;
}

</style>
""", unsafe_allow_html=True)


@st.cache_data(ttl=10)
def charger_donnees():
    param_df = pd.read_csv(param_url, dtype=str, keep_default_na=False, on_bad_lines="skip")
    all_data = []

    for _, row in param_df.iterrows():
        if row["Actif"].strip().lower() != "oui":
            continue

        df = pd.read_csv(
            row["Lien Google Sheet"].strip(),
            dtype=str,
            keep_default_na=False,
            on_bad_lines="skip"
        )

        df["Operateur"] = row["Opérateur"].strip()
        all_data.append(df)

    data = pd.concat(all_data, ignore_index=True)
    data.columns = data.columns.str.strip()

    data["Statut"] = data["Statut"].fillna("").astype(str).str.strip()
    data["Qualification"] = data["Qualification"].fillna("").astype(str).str.strip()
    data["Semaine"] = data["Semaine"].fillna("").astype(str).str.strip()
    data["Cible / code NAF"] = data["Cible / code NAF"].fillna("").astype(str).str.strip()

    data["Cible"] = (
    data["Cible"]
    .fillna("")
    .astype(str)
    .str.strip()
    .str.replace("\u00a0", " ", regex=False)
)

    data["Date appel"] = pd.to_datetime(
        data["Date de l'appel"],
        errors="coerce",
        dayfirst=True
    )

    return data


def calcul_kpi(df):
    total = len(df)
    invalide = len(df[df["Statut"] == "Invalide"])
    non_joint = len(df[df["Statut"] == "Non joint"])
    a_rappeler = len(df[df["Statut"] == "À rappeler"])
    joint = len(df[df["Statut"] == "Joint"])
    contacte = len(df[(df["Statut"] != "") & (df["Statut"] != "Invalide")])
    non_traite = total - contacte - invalide
    ok = len(df[df["Qualification"] == "OK"])
    ko = len(df[df["Qualification"] == "KO"])

    return {
        "total": total,
        "non_traite": non_traite,
        "invalide": invalide,
        "non_joint": non_joint,
        "a_rappeler": a_rappeler,
        "joint": joint,
        "contacte": contacte,
        "ok": ok,
        "ko": ko,
        "taux_traitement": contacte / total if total else 0,
        "taux_joint": joint / contacte if contacte else 0,
        "taux_tc": ok / contacte if contacte else 0,
        "taux_tj": ok / joint if joint else 0,
        "taux_invalide": invalide / total if total else 0
    }


def card(label, value, color="blue"):
    st.markdown(f"""
    <div class="kpi-card">
        <div class="kpi-label">{label}</div>
        <div class="kpi-value {color}">{value}</div>
    </div>
    """, unsafe_allow_html=True)


def build_agent_table(df):
    rows = []

    for op, g in df.groupby("Operateur"):
        k = calcul_kpi(g)

        nb_jours = g["Date appel"].dropna().dt.date.nunique()
        objectif = OBJECTIF_JOURNALIER_PAR_AGENT * nb_jours if nb_jours > 0 else OBJECTIF_JOURNALIER_PAR_AGENT

        atteinte = k["ok"] / objectif if objectif else 0

        rows.append({
            "Opérateur": op,
            "Objectif": objectif,
            "Réalisé": k["ok"],
            "Atteinte objectif %": f"{atteinte * 100:.2f}"" %",
            "Total fiches": k["total"],
            "Contacté": k["contacte"],
            "Joint": k["joint"],
            "Non traité": k["non_traite"],
            "Invalide": k["invalide"],
            "Non joint": k["non_joint"],
            "À rappeler": k["a_rappeler"],
            "Taux joint %": f"{k['taux_joint'] * 100:.2f}"" %",
            "Transfo/contacté %": f"{k['taux_tc'] * 100:.2f}"" %",
            "Transfo/joint %": f"{k['taux_tj'] * 100:.2f}"" %"
        })

    if not rows:
        return pd.DataFrame()

    return pd.DataFrame(rows).sort_values("Réalisé", ascending=False)


def exporter_excel(df_data, df_agents):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df_data.to_excel(writer, index=False, sheet_name="Données filtrées")
        df_agents.to_excel(writer, index=False, sheet_name="Agents")

    return output.getvalue()

def flow_card(icon, label, value):
    st.markdown(f"""
    <div class="kpi-flow-card">
        <div class="kpi-icon">{icon}</div>
        <div>
            <div class="kpi-flow-label">{label}</div>
            <div class="kpi-flow-value">{value}</div>
        </div>
    </div>
    """, unsafe_allow_html=True)


def gauge(title, value, max_value):

    st.markdown(f"""
    <div style="text-align:center; margin-bottom:6px;">
        <div style="
            font-size:18px;
            font-weight:850;
            color:#182466;
        ">
            {title}
        </div>
        <div style="
            width:40px;
            height:3px;
            background:#3B82F6;
            margin:4px auto;
            border-radius:2px;
        "></div>
    </div>
    """, unsafe_allow_html=True)

    fig = go.Figure(go.Indicator(
        mode="gauge",
        value=value,

        gauge={
            "axis": {
                "range": [0, max_value],
                "tickvals": [0, max_value],
                "ticktext": ["0%", f"{max_value}%"],
                "tickfont": {
                    "size": 10,
                    "color": "#182466"
                }
            },
            "bar": {
                "color": "#3c5bff",
                "thickness": 0.28
            },
            "bgcolor": "#E5E7EB",
            "borderwidth": 0
        }
    ))

    fig.update_layout(
        height=120,
        margin=dict(l=10, r=10, t=15, b=15),
        paper_bgcolor="white"
    )

    st.plotly_chart(fig, use_container_width=True)

    st.markdown(f"""
    <div style="
        text-align:center;
        font-size:22px;
        font-weight:700;
        color:#182466;
        margin-top:-72px;
        margin-bottom:45px;
    ">
        {value:.2f}%
    </div>
    """, unsafe_allow_html=True)

data_all = charger_donnees()

st.sidebar.image("logo_hellopro.png", width=250)

st.sidebar.markdown("<br>", unsafe_allow_html=True)

if st.sidebar.button("🔄 Rafraîchir les données", use_container_width=True):
    st.sidebar.success("Actualisation en cours...")
    st.cache_data.clear()
    st.cache_resource.clear()
    st.rerun()

st.sidebar.title("Filtres")

dates_valides = data_all["Date appel"].dropna()
periode = None

if not dates_valides.empty:
    date_min = dates_valides.min().date()
    date_max = dates_valides.max().date()
    periode = st.sidebar.date_input("Période", value=(date_min, date_max))
    
op_list = ["Tous"] + sorted(data_all["Operateur"].dropna().unique().tolist())
op_selected = st.sidebar.selectbox("Opérateur", op_list)

cible_list = ["Toutes"] + sorted([x for x in data_all["Cible"].dropna().unique().tolist() if x])
cible_selected = st.sidebar.selectbox("Cible", cible_list)

semaines = ["Toutes"] + sorted([x for x in data_all["Semaine"].unique().tolist() if x])
semaine_selected = st.sidebar.selectbox("Semaine", semaines)

nafs = ["Tous"] + sorted([x for x in data_all["Cible / code NAF"].unique().tolist() if x])
naf_selected = st.sidebar.selectbox("Cible / code NAF", nafs)

st.sidebar.markdown(f"""
<div class="selection-card">
    <div class="selection-title">Sélection actuelle</div>
    <div class="selection-value">Opérateur : {op_selected}</div>
    <div class="selection-value">Cible : {cible_selected}</div>
    <div class="selection-value">Semaine : {semaine_selected}</div>
    <div class="selection-value">NAF : {naf_selected}</div>
</div>
""", unsafe_allow_html=True)

data = data_all.copy()

if op_selected != "Tous":
    data = data[data["Operateur"] == op_selected]

if cible_selected != "Toutes":
    data = data[data["Cible"] == cible_selected]

if semaine_selected != "Toutes":
    data = data[data["Semaine"] == semaine_selected]

if naf_selected != "Tous":
    data = data[data["Cible / code NAF"] == naf_selected]

if periode and len(periode) == 2:
    debut = pd.to_datetime(periode[0])
    fin = pd.to_datetime(periode[1])
    data = data[(data["Date appel"] >= debut) & (data["Date appel"] <= fin)]

kpi = calcul_kpi(data)
df_agents = build_agent_table(data)

st.markdown("""
<div class="header-card">
    <div class="title"> Dashboard Production - EDF Solutions Solaires ⚡</div>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="section-title">►  Résultat global</div>', unsafe_allow_html=True)

c1, c2, c3, c4 = st.columns(4)

with c1:
    flow_card("📄", "Total fiches", kpi["total"])

with c2:
    flow_card("📞", "Contacté", kpi["contacte"])

with c3:
    flow_card("🤝", "Joint", kpi["joint"])

with c4:
    flow_card("✔️", "Qualification OK", kpi["ok"])

st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)

g1, g2, g3, g4 = st.columns(4)

with g1:
    gauge("Taux de traitement", round(kpi["taux_traitement"] * 100, 2), 100)

with g2:
    gauge("Taux joint", round(kpi["taux_joint"] * 100, 2), 50)

with g3:
    gauge("Taux OK/joint", round(kpi["taux_tj"] * 100, 2), 10)

with g4:
    gauge("Taux OK/contacté", round(kpi["taux_tc"] * 100, 2), 5)
    
st.markdown('<div class="section-title">►  Résultat par agent</div>', unsafe_allow_html=True)

def pct_to_float(value):
    try:
        return float(str(value).replace("%", "").strip())
    except:
        return 0

if not df_agents.empty:
    rows_html = ""

    for _, row in df_agents.iterrows():
        atteinte = pct_to_float(row["Atteinte objectif %"])

        rows_html += f"""\
<tr>
    <td class="op-name">{row["Opérateur"]}</td>
    <td>{row["Objectif"]}</td>
    <td>{row["Réalisé"]}</td>
    <td>
        <div class="progress-wrap">
            <span>{row["Atteinte objectif %"]}</span>
            <div class="progress-bg">
                <div class="progress-bar" style="width:{min(atteinte,100)}%;"></div>
            </div>
        </div>
    </td>
    <td>{row["Total fiches"]}</td>
    <td>{row["Contacté"]}</td>
    <td>{row["Joint"]}</td>
    <td>{row["Non traité"]}</td>
    <td>{row["Invalide"]}</td>
    <td>{row["Non joint"]}</td>
    <td>{row["À rappeler"]}</td>
    <td>{row["Taux joint %"]}</td>
    <td>{row["Transfo/contacté %"]}</td>
    <td>{row["Transfo/joint %"]}</td>
</tr>
"""

    st.markdown(dedent(f"""
<style>
.agent-table {{
    width: 100%;
    border-collapse: separate;
    border-spacing: 0;
    background: white;
    border-radius: 14px;
    overflow: hidden;
    box-shadow: 0px 4px 12px rgba(0,0,0,0.06);
    font-size: 13px;
}}

.agent-table th {{
    background: #121b4c;
    color: white;
    padding: 12px 10px;
    text-align: center;
    font-weight: 800;
    white-space: nowrap;
}}

.agent-table td {{
    padding: 11px 10px;
    text-align: center;
    border-bottom: 1px solid #E5E7EB;
    color: #111827;
    white-space: nowrap;
}}

.agent-table tr:nth-child(even) td {{
    background: #F8FAFC;
}}

.agent-table tr:hover td {{
    background: #EEF3FF;
}}

.op-name {{
    font-weight: 800;
    text-align: left !important;
}}

.progress-wrap {{
    display: flex;
    align-items: center;
    gap: 8px;
    justify-content: center;
}}

.progress-bg {{
    width: 70px;
    height: 8px;
    background: #E5E7EB;
    border-radius: 999px;
    overflow: hidden;
}}

.progress-bar {{
    height: 100%;
    background: #0F766E;
    border-radius: 999px;
}}
</style>

<div class="table-scroll">
<table class="agent-table">
    <thead>
        <tr>
            <th>Opérateur</th>
            <th>Objectif</th>
            <th>Réalisé</th>
            <th>Atteinte objectif</th>
            <th>Total fiches</th>
            <th>Contacté</th>
            <th>Joint</th>
            <th>Non traité</th>
            <th>Invalide</th>
            <th>Non joint</th>
            <th>À rappeler</th>
            <th>Taux joint</th>
            <th>Transfo/contacté</th>
            <th>Transfo/joint</th>
        </tr>
    </thead>
        <tbody>
        {rows_html.strip()}
        </tbody>
</table>
</div>
"""), unsafe_allow_html=True)

else:
    st.info("Aucune donnée agent disponible.")
    
st.markdown('<div class="section-title">►  Top opérateur</div>', unsafe_allow_html=True)

if not df_agents.empty:
    top = df_agents.iloc[0]
    st.markdown(f"""
<div style="
    background-color: #f2fff7;
    border: 1px solid #9eadff;
    padding: 16px;
    border-radius: 14px;
    font-weight: 700;
    color: #10571b;
">
🏆 {top['Opérateur']} | 
Réalisé : {top['Réalisé']} / Objectif : {top['Objectif']} 
({top['Atteinte objectif %']}%)
</div>
""", unsafe_allow_html=True)
else:
    st.info("Aucune donnée disponible.")

st.markdown('<div class="section-title">📊 Funnel production</div>', unsafe_allow_html=True)

fig_funnel = go.Figure(go.Funnel(
    y=["Total fiches", "Contacté", "Joint", "Qualification OK"],
    x=[kpi["total"], kpi["contacte"], kpi["joint"], kpi["ok"]],
    textinfo="value+percent initial",
    textfont=dict(
        size=15,
        color="white"
    ),
    marker=dict(
        color=["#0F172A", "#2563EB", "#38BDF8", "#16A34A"]
    ),
    connector=dict(
        fillcolor="#E5E7EB"
    ),
    opacity=0.95
))

fig_funnel.update_layout(
    height=420,
    paper_bgcolor="#FFFFFF",
    plot_bgcolor="#FFFFFF",
    font=dict(size=16, color="#111827"),
    margin=dict(l=120, r=20, t=30, b=20)
)

fig_funnel.update_yaxes(
    tickfont=dict(
        size=12,
        color="#182466"  # 👉 change ici la couleur
    ),
    ticklabelposition="outside"
)

fig_funnel.update_yaxes(
    tickfont=dict(
        size=12,
        color="#111827"
    )
)

st.plotly_chart(fig_funnel, use_container_width=True)

st.markdown('<div class="section-title">🎯 Réalisé vs Objectif par agent</div>', unsafe_allow_html=True)

if not df_agents.empty:
    fig_obj = go.Figure()

    fig_obj.add_trace(go.Bar(
        x=df_agents["Opérateur"],
        y=df_agents["Objectif"],
        name="Objectif",
        marker_color="#ddf0de",
        text=df_agents["Objectif"],
        textposition="outside",
        width=0.55,
        opacity=0.85
    ))

    fig_obj.add_trace(go.Bar(
        x=df_agents["Opérateur"],
        y=df_agents["Réalisé"],
        name="Réalisé",
        marker_color="#3ead46",
        text=df_agents["Réalisé"],
        textposition="inside",
        width=0.35
    ))

    fig_obj.update_layout(
        barmode="overlay",
        height=430,
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        font=dict(color="#111111"),
        legend_title_text="",
        yaxis_title="Volume",
        xaxis_title="",
        bargap=0.35,
        margin=dict(l=30, r=30, t=30, b=30)
    )

    fig_obj.update_yaxes(showgrid=True, gridcolor="#E5E7EB")
    fig_obj.update_xaxes(showgrid=False)

    st.plotly_chart(fig_obj, use_container_width=True)

st.markdown('<div class="section-title">Performance par agent</div>', unsafe_allow_html=True)

if not df_agents.empty:
    fig_perf = px.bar(
        df_agents,
        x="Opérateur",
        y=["Contacté", "Joint", "Réalisé"],
        barmode="group",
        text_auto=True,
        color_discrete_sequence=["#2563EB", "#38BDF8", "#16A34A"]
    )

    fig_perf.update_layout(
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        font=dict(color="#111111"),
        legend_title_text="Indicateurs",
        yaxis_title="Volume",
        xaxis_title=""
    )

    st.plotly_chart(fig_perf, use_container_width=True)

st.markdown('<div class="section-title">📈 Comparaison semaine / semaine</div>', unsafe_allow_html=True)

df_week = data[data["Semaine"] != ""].copy()

if not df_week.empty:
    comparaison = df_week.groupby("Semaine").agg(
        Contacte=("Statut", lambda x: ((x != "") & (x != "Invalide")).sum()),
        Joint=("Statut", lambda x: (x == "Joint").sum()),
        Realise=("Qualification", lambda x: (x == "OK").sum())
    ).reset_index()

    comparaison = comparaison.rename(columns={"Realise": "Réalisé"})

    fig_week = px.bar(
        comparaison,
        x="Semaine",
        y=["Contacte", "Joint", "Réalisé"],
        barmode="group",
        text_auto=True,
        color_discrete_sequence=["#2563EB", "#38BDF8", "#16A34A"]
    )

    fig_week.update_layout(
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        font=dict(color="#111111"),
        legend_title_text="Indicateurs",
        yaxis_title="Volume",
        xaxis_title="Semaine"
    )

    st.plotly_chart(fig_week, use_container_width=True)
else:
    st.info("Aucune semaine disponible.")

st.markdown('<div class="section-title">📈 Évolution dans le temps</div>', unsafe_allow_html=True)

df_time = data.dropna(subset=["Date appel"]).copy()

if not df_time.empty:
    evolution = df_time.groupby("Date appel").agg(
        Contacte=("Statut", lambda x: ((x != "") & (x != "Invalide")).sum()),
        Joint=("Statut", lambda x: (x == "Joint").sum()),
        Realise=("Qualification", lambda x: (x == "OK").sum())
    ).reset_index()

    evolution = evolution.rename(columns={"Realise": "Réalisé"})

    fig_time = px.line(
        evolution,
        x="Date appel",
        y=["Contacte", "Joint", "Réalisé"],
        markers=True,
        color_discrete_sequence=["#2563EB", "#38BDF8", "#16A34A"]
    )

    fig_time.update_layout(
        plot_bgcolor="#FFFFFF",
        paper_bgcolor="#FFFFFF",
        font=dict(color="#111111"),
        legend_title_text="Indicateurs",
        yaxis_title="Volume",
        xaxis_title="Date"
    )

    st.plotly_chart(fig_time, use_container_width=True)
else:
    st.info("Aucune date exploitable pour afficher l’évolution.")

st.markdown('<div class="section-title">🚨 Alertes</div>', unsafe_allow_html=True)

alertes = []

if kpi["contacte"] > 0 and kpi["taux_joint"] < 0.40:
    alertes.append("Taux de joint global inférieur à 40%.")

if kpi["taux_invalide"] > 0.15:
    alertes.append("Taux d'invalidité supérieur à 15%.")

if kpi["a_rappeler"] > 0:
    alertes.append(f"{kpi['a_rappeler']} fiche(s) à rappeler.")

if not df_agents.empty:
    faibles = df_agents[
    df_agents["Taux joint %"].str.replace(" %", "").astype(float) < 40]["Opérateur"].tolist()
    if faibles:
        alertes.append("Taux de joint faible pour : " + ", ".join(faibles))

    retard_obj = df_agents[
    df_agents["Atteinte objectif %"].str.replace(" %", "").astype(float) < 50]["Opérateur"].tolist()
    if retard_obj:
        alertes.append("Objectif atteint à moins de 50% pour : " + ", ".join(retard_obj))

if alertes:
    for alerte in alertes:
        st.markdown(f'<div class="alert-card">⚠️ {alerte}</div>', unsafe_allow_html=True)
else:
    st.markdown('<div class="ok-card">✅ Aucune alerte critique détectée.</div>', unsafe_allow_html=True)

st.markdown('<div class="section-title">📤 Export Excel</div>', unsafe_allow_html=True)

excel_file = exporter_excel(data, df_agents)

st.download_button(
    label="Télécharger l'export Excel",
    data=excel_file,
    file_name="export_dashboard_hellopro.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
