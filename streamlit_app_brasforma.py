
# streamlit_app_brasforma.py
# Dashboard Comercial – Brasforma | Streamlit Cloud ready
# Usage on Streamlit Cloud:
# 1) Upload this file and the Excel: "Dashboard - Comite Semanal - Brasforma (1).xlsx"
# 2) Set main file to this script. Done.

import streamlit as st
import pandas as pd
import numpy as np

st.set_page_config(page_title="Brasforma – Dashboard Comercial", layout="wide")

@st.cache_data(show_spinner=False)
def load_data(path: str):
        try:
        xls = pd.ExcelFile(path, engine="openpyxl")
    except Exception as e:
        st.error("Falha ao abrir Excel. Verifique se o arquivo é .xlsx válido e se a dependência openpyxl está instalada.")
        st.exception(e)
        st.stop()
    df = pd.read_excel(xls, sheet_name="Carteira de Vendas")
    # Normalize columns
    df.columns = [c.strip() for c in df.columns]
    # Safe dtype handling
    # Dates
    for col in ["Data / Mês", "Data Final", "Data do Pedido", "Data da Entrega", "Data Inserção"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    # Monetary and numeric fields (accept comma decimal formats)
    def to_num(x):
        if pd.isna(x): 
            return np.nan
        if isinstance(x, (int, float, np.integer, np.floating)): 
            return x
        s = str(x).replace(".", "").replace(",", ".")
        try:
            return float(s)
        except:
            return np.nan
    if "Valor Pedido R$" in df.columns:
        df["Valor Pedido R$"] = df["Valor Pedido R$"].apply(to_num)
    if "TICKET MÉDIO" in df.columns:
        df["TICKET MÉDIO"] = df["TICKET MÉDIO"].apply(to_num)
    if "Quant. Pedidos" in df.columns:
        df["Quant. Pedidos"] = pd.to_numeric(df["Quant. Pedidos"], errors="coerce")
    # Derivations
    if "Data / Mês" in df.columns:
        df["Ano"] = df["Data / Mês"].dt.year
        df["Mes"] = df["Data / Mês"].dt.month
        df["Ano-Mes"] = df["Data / Mês"].dt.to_period("M").astype(str)
    if "Data do Pedido" in df.columns and "Data da Entrega" in df.columns:
        df["LeadTime (dias)"] = (df["Data da Entrega"] - df["Data do Pedido"]).dt.days
    if "Atrasado / No prazo" in df.columns:
        df["AtrasadoFlag"] = df["Atrasado / No prazo"].astype(str).str.contains("Atras", case=False, na=False)
    return df

DATA_PATH = "Dashboard - Comite Semanal - Brasforma (1).xlsx"
st.sidebar.title("Filtros")
uploaded = st.sidebar.file_uploader("Base: Dashboard - Comitê Semanal - Brasforma (1).xlsx", type=["xlsx"], accept_multiple_files=False)
if uploaded is not None:
    DATA_PATH = uploaded
df = load_data(DATA_PATH)

# Filter widgets
col_f1, col_f2 = st.sidebar.columns(2)
# Período
if "Data / Mês" in df.columns:
    min_date = pd.to_datetime(df["Data / Mês"]).min()
    max_date = pd.to_datetime(df["Data / Mês"]).max()
    d_ini, d_fim = st.sidebar.date_input("Período (Data / Mês)", value=(min_date, max_date))
else:
    d_ini = d_fim = None

# Chained filters
reg = st.sidebar.multiselect("Regional", sorted(df["Regional"].dropna().unique()) if "Regional" in df.columns else [])
rep = st.sidebar.multiselect("Representante", sorted(df["Representante"].dropna().unique()) if "Representante" in df.columns else [])
uf  = st.sidebar.multiselect("UF", sorted(df["UF"].dropna().unique()) if "UF" in df.columns else [])
stat = st.sidebar.multiselect("Status Produção/Faturamento", sorted(df["Status de Produção / Faturamento"].dropna().unique()) if "Status de Produção / Faturamento" in df.columns else [])
cliente = st.sidebar.text_input("Cliente (contém)")
item = st.sidebar.text_input("SKU/Item (contém)")

# Apply filters
flt = df.copy()
if "Data / Mês" in df.columns and d_ini is not None:
    flt = flt[(flt["Data / Mês"] >= pd.to_datetime(d_ini)) & (flt["Data / Mês"] <= pd.to_datetime(d_fim))]
if reg:
    flt = flt[flt["Regional"].isin(reg)]
if rep:
    flt = flt[flt["Representante"].isin(rep)]
if uf:
    flt = flt[flt["UF"].isin(uf)]
if stat:
    flt = flt[flt["Status de Produção / Faturamento"].isin(stat)]
if cliente:
    flt = flt[flt["Nome Cliente"].astype(str).str.contains(cliente, case=False, na=False)]
if item:
    flt = flt[flt["ITEM"].astype(str).str.contains(item, case=False, na=False)]

# KPIs block
def fmt_money(v): 
    if pd.isna(v): return "R$ 0"
    return f"R$ {v:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")

def kpi_card(label, value, delta=None):
    st.metric(label, value, delta)

col1, col2, col3, col4, col5, col6 = st.columns(6)
# Faturamento (sum Valor Pedido R$)
fat = flt["Valor Pedido R$"].sum() if "Valor Pedido R$" in flt.columns else np.nan
# Pedidos
n_ped = flt["Pedido"].nunique() if "Pedido" in flt.columns else len(flt)
# Clientes
n_cli = flt["Nome Cliente"].nunique() if "Nome Cliente" in flt.columns else np.nan
# Ticket Médio (ponderado)
ticket = (fat / n_ped) if (n_ped and n_ped>0) else np.nan
# Itens (SKUs)
n_sku = flt["ITEM"].nunique() if "ITEM" in flt.columns else np.nan
# % Atraso (entre pedidos com status/flag)
pct_atraso = np.nan
if "AtrasadoFlag" in flt.columns:
    base = flt["AtrasadoFlag"].notna().sum()
    if base>0:
        pct_atraso = 100*flt["AtrasadoFlag"].mean()

col1.metric("Faturamento", fmt_money(fat))
col2.metric("Pedidos", f"{n_ped:,}".replace(",", "."))
col3.metric("Clientes", f"{n_cli:,}".replace(",", ".") if pd.notna(n_cli) else "-")
col4.metric("Ticket Médio", fmt_money(ticket) if pd.notna(ticket) else "-")
col5.metric("SKUs Vendidos", f"{n_sku:,}".replace(",", ".") if pd.notna(n_sku) else "-")
col6.metric("% Pedidos Atrasados", f"{pct_atraso:,.1f}%".replace(",", "X").replace(".", ",").replace("X", ".") if pd.notna(pct_atraso) else "-")

st.markdown("---")

# Charts
import altair as alt

# 1) Série temporal por mês (Faturamento)
if "Ano-Mes" in flt.columns and "Valor Pedido R$" in flt.columns:
    serie = flt.groupby("Ano-Mes", as_index=False)["Valor Pedido R$"].sum().sort_values("Ano-Mes")
    st.subheader("Faturamento por Mês")
    chart = alt.Chart(serie).mark_bar().encode(
        x=alt.X("Ano-Mes:N", sort=None, title="Ano-Mês"),
        y=alt.Y("Valor Pedido R$:Q", title="Faturamento (R$)"),
        tooltip=["Ano-Mes", alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
    ).properties(height=300)
    st.altair_chart(chart, use_container_width=True)

# 2) Top 15 Clientes
if "Nome Cliente" in flt.columns and "Valor Pedido R$" in flt.columns:
    top_cli = flt.groupby("Nome Cliente", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False).head(15)
    st.subheader("Top 15 Clientes (Faturamento)")
    chart2 = alt.Chart(top_cli).mark_bar().encode(
        x=alt.X("Valor Pedido R$:Q", title="Faturamento (R$)"),
        y=alt.Y("Nome Cliente:N", sort="-x", title="Cliente"),
        tooltip=["Nome Cliente", alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
    ).properties(height=450)
    st.altair_chart(chart2, use_container_width=True)

# 3) Top 15 SKUs
if "ITEM" in flt.columns and "Valor Pedido R$" in flt.columns:
    top_sku = flt.groupby("ITEM", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False).head(15)
    st.subheader("Top 15 SKUs (Faturamento)")
    chart3 = alt.Chart(top_sku).mark_bar().encode(
        x=alt.X("Valor Pedido R$:Q", title="Faturamento (R$)"),
        y=alt.Y("ITEM:N", sort="-x", title="SKU"),
        tooltip=["ITEM", alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
    ).properties(height=450)
    st.altair_chart(chart3, use_container_width=True)

# 4) Por UF
if "UF" in flt.columns and "Valor Pedido R$" in flt.columns:
    por_uf = flt.groupby("UF", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
    st.subheader("Faturamento por UF")
    chart4 = alt.Chart(por_uf).mark_bar().encode(
        x=alt.X("UF:N", sort="-y"),
        y=alt.Y("Valor Pedido R$:Q", title="Faturamento (R$)"),
        tooltip=["UF", alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
    ).properties(height=300)
    st.altair_chart(chart4, use_container_width=True)

# 5) Por Representante
if "Representante" in flt.columns and "Valor Pedido R$" in flt.columns:
    por_rep = flt.groupby("Representante", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False).head(20)
    st.subheader("Faturamento por Representante (Top 20)")
    chart5 = alt.Chart(por_rep).mark_bar().encode(
        x=alt.X("Valor Pedido R$:Q", title="Faturamento (R$)"),
        y=alt.Y("Representante:N", sort="-x"),
        tooltip=["Representante", alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
    ).properties(height=500)
    st.altair_chart(chart5, use_container_width=True)

# 6) SLA / Atraso
if "AtrasadoFlag" in flt.columns and "LeadTime (dias)" in flt.columns:
    st.subheader("Lead Time e Atraso")
    c1, c2 = st.columns(2)
    with c1:
        lt = flt["LeadTime (dias)"].dropna()
        if len(lt)>0:
            lt_desc = pd.Series(lt).describe()[["count","mean","50%","min","max"]]
            st.dataframe(lt_desc.to_frame("LeadTime (dias)").rename(index={"50%":"mediana"}))
    with c2:
        atrasos = flt.groupby("Atrasado / No prazo", as_index=False)["Pedido"].nunique() if "Atrasado / No prazo" in flt.columns else None
        if atrasos is not None and len(atrasos)>0:
            st.dataframe(atrasos.rename(columns={"Pedido":"Qtde Pedidos"}))

# 7) Tabela detalhe
with st.expander("Tabela detalhada (aplica filtros)"):
    st.dataframe(flt)

st.caption("Fonte: Carteira de Vendas – Dashboard Comitê Semanal – Brasforma")
