import streamlit as st
import pandas as pd
import numpy as np
from io import BytesIO
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.pagesizes import A4
import altair as alt
from datetime import datetime

# -------------------------------------------------------------------------
# CONFIGURAÇÃO
# -------------------------------------------------------------------------
DEFAULT_ARQUIVO_A = "POSTO.xlsx"
DEFAULT_ARQUIVO_B = "PMM.xlsx"

EXPECTED_COLS = {
    "data": ["data", "Data"],
    "nota": ["numero_nota", "Numero_Nota", "nota"],
    "tipo": ["tipo_combustivel", "Tipo_Combustivel", "combustivel", "Tipo"],
    "valor": ["litragem", "Litragem", "litros", "Valor"],
    "setor": ["setor", "Setor", "departamento", "secretaria"]
}

# -------------------------------------------------------------------------
# FUNÇÕES AUXILIARES
# -------------------------------------------------------------------------
def guess_column(df, keys):
    df_cols = {c.lower(): c for c in df.columns}
    for key in keys:
        if key.lower() in df_cols:
            return df_cols[key.lower()]
    for k in keys:
        for c in df.columns:
            if k.lower() in c.lower():
                return c
    return None

def normalize_df(df):
    df = df.copy()
    new_cols = {}
    for key, variations in EXPECTED_COLS.items():
        found = guess_column(df, variations)
        if found:
            new_cols[found] = key.upper()
    if new_cols:
        df.rename(columns=new_cols, inplace=True)

    for key in EXPECTED_COLS:
        if key.upper() not in df.columns:
            df[key.upper()] = np.nan

    return df

# -------------------------------------------------------------------------
# FUNÇÃO PRINCIPAL DE ANÁLISE
# -------------------------------------------------------------------------
def analisar(df_a, df_b, setor):
    df_a = normalize_df(df_a)
    df_b = normalize_df(df_b)

    if setor and setor != "Todos":
        df_a = df_a[df_a["SETOR"].astype(str).str.upper() == setor.upper()]
        df_b = df_b[df_b["SETOR"].astype(str).str.upper() == setor.upper()]

    df_a["NOTA"] = df_a["NOTA"].astype(str).str.zfill(4)
    df_b["NOTA"] = df_b["NOTA"].astype(str).str.zfill(4)

    df_a["VALOR"] = pd.to_numeric(df_a["VALOR"], errors="coerce")
    df_b["VALOR"] = pd.to_numeric(df_b["VALOR"], errors="coerce")

    df_a["DATA_PARSED"] = pd.to_datetime(df_a["DATA"], errors="coerce")
    df_b["DATA_PARSED"] = pd.to_datetime(df_b["DATA"], errors="coerce")

    results = {}

    # Divergências internas
    results["duplicadas_a"] = df_a[df_a["NOTA"].duplicated(keep=False)]
    results["duplicadas_b"] = df_b[df_b["NOTA"].duplicated(keep=False)]

    results["negativos_a"] = df_a[df_a["VALOR"] < 0]
    results["negativos_b"] = df_b[df_b["VALOR"] < 0]

    def rep_day(df):
        t = df.copy()
        t["DATA_FMT"] = t["DATA_PARSED"].dt.strftime("%d/%m/%Y")
        return t[t.duplicated(subset=["DATA_FMT", "NOTA", "VALOR"], keep=False)]

    results["rep_mesmo_dia_a"] = rep_day(df_a)
    results["rep_mesmo_dia_b"] = rep_day(df_b)

    def many_days(df):
        t = df.copy()
        t["DATA_FMT"] = t["DATA_PARSED"].dt.strftime("%d/%m/%Y")
        g = t.groupby("NOTA")["DATA_FMT"].nunique()
        notas = g[g > 1].index
        return t[t["NOTA"].isin(notas)]

    results["nota_diff_a"] = many_days(df_a)
    results["nota_diff_b"] = many_days(df_b)

    # Comparação entre planilhas
    merged = pd.merge(df_a, df_b, on="NOTA", suffixes=("_A", "_B"), how="outer")
    results["merged"] = merged

    results["notas_apenas_em_a"] = merged[merged["VALOR_B"].isnull()]
    results["notas_apenas_em_b"] = merged[merged["VALOR_A"].isnull()]

    merged["DATA_A_PARSED"] = pd.to_datetime(merged.get("DATA_PARSED_A"), errors="coerce")
    merged["DATA_B_PARSED"] = pd.to_datetime(merged.get("DATA_PARSED_B"), errors="coerce")

    results["datas_divergentes"] = merged[merged["DATA_A_PARSED"] != merged["DATA_B_PARSED"]]

    results["tipos_divergentes"] = merged[
        (merged["TIPO_A"].astype(str) != merged["TIPO_B"].astype(str))
        & (~merged["TIPO_A"].isna() | ~merged["TIPO_B"].isna())
    ]

    results["valores_divergentes"] = merged[
        ~(np.isclose(merged["VALOR_A"].fillna(-999999), merged["VALOR_B"].fillna(-888888)))
    ]

    # Totais por combustível
    def sum_by_fuel(df, pats):
        mask = pd.Series(False, index=df.index)
        for p in pats:
            mask = mask | df["TIPO"].astype(str).str.contains(p, case=False, na=False)
        return df.loc[mask, "VALOR"].sum()

    gasolina = ["GAS", "GASOLINA"]
    diesel = ["DIE", "DIESEL"]

    gas_a = sum_by_fuel(df_a, gasolina)
    gas_b = sum_by_fuel(df_b, gasolina)
    die_a = sum_by_fuel(df_a, diesel)
    die_b = sum_by_fuel(df_b, diesel)

    # Totais gerais
    resumo = {
        "Total Registros A": len(df_a),
        "Total Registros B": len(df_b),
        "Litros Totais A": df_a["VALOR"].sum(),
        "Litros Totais B": df_b["VALOR"].sum(),
        "Gasolina A": gas_a,
        "Gasolina B": gas_b,
        "Diesel A": die_a,
        "Diesel B": die_b,
        "Duplicadas A": len(results["duplicadas_a"]),
        "Duplicadas B": len(results["duplicadas_b"]),
        "Negativos A": len(results["negativos_a"]),
        "Negativos B": len(results["negativos_b"]),
        "Repetidas A": len(results["rep_mesmo_dia_a"]),
        "Repetidas B": len(results["rep_mesmo_dia_b"]),
        "Nota dias diferentes A": len(results["nota_diff_a"]),
        "Nota dias diferentes B": len(results["nota_diff_b"]),
        "Só em A": len(results["notas_apenas_em_a"]),
        "Só em B": len(results["notas_apenas_em_b"]),
        "Datas divergentes": len(results["datas_divergentes"]),
        "Tipos divergentes": len(results["tipos_divergentes"]),
        "Valores divergentes": len(results["valores_divergentes"]),
    }

    results["resumo"] = resumo
    results["df_a"] = df_a
    results["df_b"] = df_b

    return results

# -------------------------------------------------------------------------
# EXCEL
# -------------------------------------------------------------------------
def gerar_excel(results):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        pd.DataFrame([results["resumo"]]).to_excel(w, "Resumo", index=False)
        results["df_a"].to_excel(w, "POSTO", index=False)
        results["df_b"].to_excel(w, "PMM", index=False)
        results["merged"].to_excel(w, "Comparacao", index=False)

        for k, df in results.items():
            if isinstance(df, pd.DataFrame) and not df.empty:
                name = k[:31]
                if name not in ["Resumo", "POSTO", "PMM", "Comparacao"]:
                    df.to_excel(w, name, index=False)

    buf.seek(0)
    return buf

# -------------------------------------------------------------------------
# PDF
# -------------------------------------------------------------------------
def gerar_pdf(results):
    buf = BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4)
    styles = getSampleStyleSheet()
    story = []

    story.append(Paragraph("<b>Relatório de Auditoria</b>", styles["Title"]))
    story.append(Spacer(1, 8))
    story.append(Paragraph(f"Gerado em {datetime.now()}", styles["BodyText"]))
    story.append(Spacer(1, 8))

    for k, v in results["resumo"].items():
        if isinstance(v, float):
            story.append(Paragraph(f"<b>{k}:</b> {v:.2f}", styles["BodyText"]))
        else:
            story.append(Paragraph(f"<b>{k}:</b> {v}", styles["BodyText"]))
        story.append(Spacer(1, 4))

    doc.build(story)
    buf.seek(0)
    return buf

# -------------------------------------------------------------------------
# INTERFACE STREAMLIT
# -------------------------------------------------------------------------
st.set_page_config(layout="wide")
st.title("🚚 Dashboard de Auditoria de Abastecimentos")

st.sidebar.header("Configurações")
file_a = st.sidebar.file_uploader("Planilha A (POSTO)", type=["xlsx"])
file_b = st.sidebar.file_uploader("Planilha B (PMM)", type=["xlsx"])
use_default = st.sidebar.checkbox("Usar arquivos padrão (POSTO.xlsx / PMM.xlsx)", True)
processar = st.sidebar.button("Processar")

if processar:
    df_a = pd.read_excel(file_a) if file_a else pd.read_excel(DEFAULT_ARQUIVO_A)
    df_b = pd.read_excel(file_b) if file_b else pd.read_excel(DEFAULT_ARQUIVO_B)

    col_a = guess_column(df_a, EXPECTED_COLS["setor"])
    col_b = guess_column(df_b, EXPECTED_COLS["setor"])

    setores = set()
    if col_a: setores.update(df_a[col_a].dropna().astype(str).unique())
    if col_b: setores.update(df_b[col_b].dropna().astype(str).unique())
    setores = sorted(list(setores))

    setor = st.sidebar.selectbox("Filtrar por Setor", ["Todos"] + setores)

    results = analisar(df_a, df_b, setor)
    st.session_state["results"] = results

# -------------------------------------------------------------------------
# EXIBIÇÃO
# -------------------------------------------------------------------------
if "results" not in st.session_state:
    st.info("Carregue e processe as planilhas.")
else:
    results = st.session_state["results"]
    resumo = results["resumo"]

    #------------------------ RESUMO -----------------------
    st.subheader("📊 Resumo")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Gasolina A", f"{resumo['Gasolina A']:.2f} L")
    c2.metric("Gasolina B", f"{resumo['Gasolina B']:.2f} L")
    c3.metric("Diesel A", f"{resumo['Diesel A']:.2f} L")
    c4.metric("Diesel B", f"{resumo['Diesel B']:.2f} L")

    colA, colB = st.columns(2)
    colA.metric("Litros Totais POSTO", f"{resumo['Litros Totais A']:.2f} L")
    colB.metric("Litros Totais PMM", f"{resumo['Litros Totais B']:.2f} L")

    st.markdown("---")

    #------------------------ GRÁFICO POR SECRETARIA -----------------------
    st.subheader("🏛️ Consumo por Secretaria (Gasolina / Diesel / POSTO / PMM)")

    df_a = results["df_a"]
    df_b = results["df_b"]

    # Se não houver setor, evita erro
    if "SETOR" not in df_a.columns:
        st.warning("Nenhuma coluna de setor encontrada.")
    else:
        def sum_by_sector(df, label):
            df2 = df.copy()
            df2["combustivel"] = df2["TIPO"].astype(str).str.upper()
            df2["gasolina"] = df2["combustivel"].str.contains("GAS")
            df2["diesel"] = df2["combustivel"].str.contains("DIE")

            return pd.DataFrame({
                "Setor": df2["SETOR"],
                "Gasolina": df2[df2["gasolina"]]["VALOR"].groupby(df2["SETOR"]).sum(),
                "Diesel": df2[df2["diesel"]]["VALOR"].groupby(df2["SETOR"]).sum()
            }).fillna(0).assign(Origem=label)

        tabela_a = sum_by_sector(df_a, "POSTO")
        tabela_b = sum_by_sector(df_b, "PMM")

        consumo = pd.concat([tabela_a, tabela_b]).reset_index()

        # Converter para formato longo
        consumo_long = consumo.melt(
            id_vars=["Setor", "Origem"],
            value_vars=["Gasolina", "Diesel"],
            var_name="Combustível",
            value_name="Litros"
        )

        chart = alt.Chart(consumo_long).mark_bar().encode(
            x=alt.X("Litros:Q", title="Litros Abastecidos"),
            y=alt.Y("Setor:N", title="Setor", sort="-x"),
            color="Origem:N",
            column="Combustível:N",
            tooltip=["Setor", "Origem", "Combustível", "Litros"]
        ).properties(height=350)

        st.altair_chart(chart, use_container_width=True)

    st.markdown("---")

    #------------------------ TABELAS -----------------------
    def show(df, title):
        if df is not None and not df.empty:
            with st.expander(f"{title} ({len(df)})"):
                st.dataframe(df)

    show(results["duplicadas_a"], "Duplicadas A")
    show(results["duplicadas_b"], "Duplicadas B")
    show(results["negativos_a"], "Negativos A")
    show(results["negativos_b"], "Negativos B")
    show(results["rep_mesmo_dia_a"], "Repetidas Mesmo Dia A")
    show(results["rep_mesmo_dia_b"], "Repetidas Mesmo Dia B")
    show(results["nota_diff_a"], "Not a em Dias Diferentes A")
    show(results["nota_diff_b"], "Nota em Dias Diferentes B")
    show(results["notas_apenas_em_a"], "Somente na A")
    show(results["notas_apenas_em_b"], "Somente na B")
    show(results["datas_divergentes"], "Datas Divergentes")
    show(results["tipos_divergentes"], "Tipos Divergentes")
    show(results["valores_divergentes"], "Valores Divergentes")

    st.markdown("---")

    #------------------------ DOWNLOADS -----------------------
    st.subheader("📥 Exportar Relatórios")

    col1, col2 = st.columns(2)
    with col1:
        st.download_button(
            "⬇️ Baixar Excel",
            gerar_excel(results),
            file_name="relatorio_abastecimentos.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    with col2:
        st.download_button(
            "⬇️ Baixar PDF",
            gerar_pdf(results),
            file_name="relatorio_abastecimentos.pdf",
            mime="application/pdf"
        )
