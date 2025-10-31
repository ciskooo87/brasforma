
# streamlit_app_brasforma_v5_1.py
# Brasforma – Dashboard Comercial v5.1
# Incrementos: KPIs GRÁFICOS na Visão Executiva (sparklines e donut 100%)
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pathlib import Path

st.set_page_config(page_title="Brasforma – Dashboard Comercial v5.1", layout="wide")

# ---------------- Utils ----------------
def to_num(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return float(x)
    s = str(x).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return np.nan

def fmt_money(v):
    if pd.isna(v): return "-"
    return ("R$ " + f"{v:,.2f}").replace(",", "X").replace(".", ",").replace("X", ".")

def fmt_int(v):
    if pd.isna(v): return "-"
    return f"{int(v):,}".replace(",", ".")

def fmt_pct(v, decimals=1):
    if pd.isna(v): return "-"
    return f"{v:.{decimals}f}%".replace(".", ",")

def display_table(df, money_cols=None, pct_cols=None, int_cols=None, max_rows=500):
    money_cols = money_cols or []
    pct_cols = pct_cols or []
    int_cols = int_cols or []
    view = df.copy().head(max_rows)
    for c in view.columns:
        if c in money_cols:
            view[c] = view[c].apply(fmt_money)
        elif c in pct_cols:
            view[c] = view[c].apply(lambda x: fmt_pct(x, 1))
        elif c in int_cols:
            view[c] = view[c].apply(fmt_int)
    st.dataframe(view, use_container_width=True)

# ---------------- Load & prep ----------------
@st.cache_data(show_spinner=False)
def load_data(path: str, sheet_name="Carteira de Vendas"):
    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
    except Exception as e:
        st.error("Falha ao abrir Excel. Verifique .xlsx e dependência openpyxl.")
        st.exception(e)
        st.stop()
    df = pd.read_excel(xls, sheet_name=sheet_name)
    df.columns = [c.strip() for c in df.columns]

    for col in ["Data / Mês","Data Final","Data do Pedido","Data da Entrega","Data Inserção"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    for col in ["Valor Pedido R$","TICKET MÉDIO","Quant. Pedidos","Custo"]:
        if col in df.columns:
            if col == "Quant. Pedidos":
                df[col] = pd.to_numeric(df[col], errors="coerce")
            else:
                df[col] = df[col].apply(to_num)

    if "Data / Mês" in df.columns:
        df["Ano"] = df["Data / Mês"].dt.year
        df["Mes"] = df["Data / Mês"].dt.month
        df["Ano-Mes"] = df["Data / Mês"].dt.to_period("M").astype(str)

    if "Data do Pedido" in df.columns and "Data da Entrega" in df.columns:
        df["LeadTime (dias)"] = (df["Data da Entrega"] - df["Data do Pedido"]).dt.days

    if "Atrasado / No prazo" in df.columns:
        df["AtrasadoFlag"] = df["Atrasado / No prazo"].astype(str).str.contains("Atras", case=False, na=False)

    # COST: unitário x quantidade (auto-detect; fallback para coluna M)
    qty_candidates = ["Qtde","QTDE","Quantidade","Quantidade Pedido","Qtd","QTD","Quant.","Quant","Qde","QTD.","QTD PEDIDA","QTD PEDIDO","QTD SOLICITADA","QTD Solicitada"]
    qty_col = None
    for c in qty_candidates:
        if c in df.columns:
            qty_col = c; break
    if qty_col is None:
        try:
            qty_col = df.columns[12]  # fallback para coluna M
        except Exception:
            qty_col = None

    if "Custo" in df.columns:
        if qty_col is not None:
            df[qty_col] = df[qty_col].apply(to_num)
            df["Custo Total"] = df["Custo"].apply(to_num) * df[qty_col]
        else:
            df["Custo Total"] = df["Custo"].apply(to_num)
    else:
        df["Custo Total"] = np.nan

    if "Valor Pedido R$" in df.columns:
        df["Lucro Bruto"] = df["Valor Pedido R$"] - df["Custo Total"]
        df["Margem %"] = np.where(df["Valor Pedido R$"]>0, 100*df["Lucro Bruto"]/df["Valor Pedido R$"], np.nan)

    if "Pedido" in df.columns and "ITEM" in df.columns:
        df["PedidoItemKey"] = df["Pedido"].astype(str) + "||" + df["ITEM"].astype(str)

    return df, qty_col

DEFAULT_DATA = "Dashboard - Comite Semanal - Brasforma IA (1).xlsx"
ALT_DATA = "Dashboard - Comite Semanal - Brasforma (1).xlsx"

st.sidebar.title("Fonte de dados")
uploaded = st.sidebar.file_uploader("Envie a base (.xlsx)", type=["xlsx"], accept_multiple_files=False)
data_path = uploaded if uploaded is not None else (DEFAULT_DATA if Path(DEFAULT_DATA).exists() else ALT_DATA)
st.sidebar.caption(f"Arquivo em uso: **{data_path}**")

df, qty_col = load_data(data_path)

st.sidebar.title("Filtros")
if "Data / Mês" in df.columns:
    min_date = pd.to_datetime(df["Data / Mês"]).min()
    max_date = pd.to_datetime(df["Data / Mês"]).max()
    d_ini, d_fim = st.sidebar.date_input("Período (Data / Mês)", value=(min_date, max_date))
else:
    d_ini = d_fim = None

reg = st.sidebar.multiselect("Regional", sorted(df["Regional"].dropna().unique()) if "Regional" in df.columns else [])
rep = st.sidebar.multiselect("Representante", sorted(df["Representante"].dropna().unique()) if "Representante" in df.columns else [])
uf  = st.sidebar.multiselect("UF", sorted(df["UF"].dropna().unique()) if "UF" in df.columns else [])
stat = st.sidebar.multiselect("Status Prod./Fat.", sorted(df["Status de Produção / Faturamento"].dropna().unique()) if "Status de Produção / Faturamento" in df.columns else [])
cliente = st.sidebar.text_input("Cliente (contém)")
item = st.sidebar.text_input("SKU/Item (contém)")
show_neg = st.sidebar.checkbox("Mostrar apenas linhas com margem negativa", value=False)

def apply_filters(_df):
    flt = _df.copy()
    if "Data / Mês" in flt.columns and d_ini is not None:
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
    if show_neg and "Lucro Bruto" in flt.columns:
        flt = flt[flt["Lucro Bruto"] < 0]
    return flt

flt = apply_filters(df)

# --------------- KPIs base ---------------
def calc_kpis(_df):
    fat = _df["Valor Pedido R$"].sum() if "Valor Pedido R$" in _df.columns else np.nan
    n_ped = _df["Pedido"].nunique() if "Pedido" in _df.columns else len(_df)
    n_cli = _df["Nome Cliente"].nunique() if "Nome Cliente" in _df.columns else np.nan
    n_sku = _df["ITEM"].nunique() if "ITEM" in _df.columns else np.nan
    ticket = (fat / n_ped) if (n_ped and n_ped>0) else np.nan
    lucro = _df["Lucro Bruto"].sum() if "Lucro Bruto" in _df.columns else np.nan
    margem_w = 100*(lucro/fat) if (pd.notna(lucro) and fat and fat>0) else np.nan
    pct_rentavel = 100.0*(_df["Lucro Bruto"]>0).mean() if "Lucro Bruto" in _df.columns and len(_df)>0 else np.nan
    return fat, n_ped, n_cli, n_sku, ticket, lucro, margem_w, pct_rentavel

fat, n_ped, n_cli, n_sku, ticket, lucro, margem_w, pct_rentavel = calc_kpis(flt)

# =============== Layout ===============
tabs = st.tabs([
    "Visão Executiva","Clientes – RFM","Rentabilidade","Clientes","Produtos","Representantes","Geografia","Operacional","Pareto/ABC","Exportar"
])
tab_exec, tab_rfm, tab_profit, tab_cli, tab_sku, tab_rep, tab_geo, tab_ops, tab_pareto, tab_export = tabs

with tab_exec:
    st.subheader("KPIs Executivos")
    c1, c2, c3 = st.columns(3)
    c1.metric("Faturamento", fmt_money(fat))
    c2.metric("Pedidos", fmt_int(n_ped))
    c3.metric("Ticket Médio", fmt_money(ticket) if pd.notna(ticket) else "-")
    c4, c5, c6 = st.columns(3)
    c4.metric("Lucro Bruto", fmt_money(lucro))
    c5.metric("Margem Bruta (pond.)", fmt_pct(margem_w) if pd.notna(margem_w) else "-")
    c6.metric("% Itens Rentáveis", fmt_pct(pct_rentavel) if pd.notna(pct_rentavel) else "-")

    st.markdown("### KPI gráficos")
    # Séries mensais (últimos 12)
    if {"Ano-Mes","Valor Pedido R$","Lucro Bruto","Margem %"} .issubset(flt.columns):
        serie = flt.groupby("Ano-Mes", as_index=False).agg({
            "Valor Pedido R$":"sum",
            "Lucro Bruto":"sum"
        }).sort_values("Ano-Mes")
        # Margem mensal (ponderada)
        mg = flt.groupby("Ano-Mes", as_index=False).apply(lambda d: pd.Series({
            "Margem %": (100*d["Lucro Bruto"].sum()/d["Valor Pedido R$"].sum()) if d["Valor Pedido R$"].sum()>0 else np.nan
        })).reset_index(drop=True)
        serie = serie.merge(mg, on="Ano-Mes", how="left")
        if len(serie)>12:
            serie = serie.tail(12)

        k1, k2, k3 = st.columns(3)
        with k1:
            st.caption("Faturamento – últimos 12 meses")
            st.altair_chart(
                alt.Chart(serie).mark_area(opacity=0.4).encode(
                    x=alt.X("Ano-Mes:N", sort=None, title=None),
                    y=alt.Y("Valor Pedido R$:Q", title=None),
                    tooltip=[alt.Tooltip("Ano-Mes:N"), alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
                ),
                use_container_width=True
            )
        with k2:
            st.caption("Lucro Bruto – últimos 12 meses")
            st.altair_chart(
                alt.Chart(serie).mark_area(opacity=0.4).encode(
                    x=alt.X("Ano-Mes:N", sort=None, title=None),
                    y=alt.Y("Lucro Bruto:Q", title=None),
                    tooltip=[alt.Tooltip("Ano-Mes:N"), alt.Tooltip("Lucro Bruto:Q", format=",.0f")]
                ),
                use_container_width=True
            )
        with k3:
            st.caption("Margem Bruta (%) – últimos 12 meses")
            st.altair_chart(
                alt.Chart(serie).mark_line(point=True).encode(
                    x=alt.X("Ano-Mes:N", sort=None, title=None),
                    y=alt.Y("Margem %:Q", title=None),
                    tooltip=[alt.Tooltip("Ano-Mes:N"), alt.Tooltip("Margem %:Q", format=",.1f")]
                ),
                use_container_width=True
            )

    # Donut 100%: Itens rentáveis vs negativos
    if "Lucro Bruto" in flt.columns and len(flt)>0:
        pos = int((flt["Lucro Bruto"]>0).sum())
        neg = int((flt["Lucro Bruto"]<0).sum())
        tot = pos + neg
        donut_df = pd.DataFrame({
            "Categoria": ["Rentáveis","Negativos"],
            "Qtd": [pos, neg]
        })
        cdon1, cdon2 = st.columns([2,1])
        with cdon1:
            st.caption("Composição de linhas – rentáveis vs negativas")
            st.altair_chart(
                alt.Chart(donut_df).mark_arc(innerRadius=60).encode(
                    theta="Qtd:Q",
                    color="Categoria:N",
                    tooltip=["Categoria","Qtd"]
                ).properties(height=300),
                use_container_width=True
            )
        with cdon2:
            if tot>0:
                pct_pos = 100*pos/tot
                st.metric("% Linhas Rentáveis", fmt_pct(pct_pos))
            else:
                st.metric("% Linhas Rentáveis", "-")

# ---- As demais abas: reaproveitar v5 (RFM, Rentabilidade, etc.) ----
# Para evitar duplicação e manter foco do pedido, manteríamos o restante do app v5.
# Nesta versão demo, mostramos apenas a Visão Executiva incrementada.
