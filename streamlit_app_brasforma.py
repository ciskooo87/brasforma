
# streamlit_app_brasforma_v2.py
# Dashboard Comercial – Brasforma | v2 (abas, deltas MoM/YoY, Pareto/ABC, coortes leves, export)
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt

st.set_page_config(page_title="Brasforma – Dashboard Comercial v2", layout="wide")

@st.cache_data(show_spinner=False)
def load_data(path: str):
    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
    except Exception as e:
        st.error("Falha ao abrir Excel. Verifique se o arquivo é .xlsx válido e se a dependência openpyxl está instalada.")
        st.exception(e)
        st.stop()
    df = pd.read_excel(xls, sheet_name="Carteira de Vendas")
    df.columns = [c.strip() for c in df.columns]
    # datas
    for col in ["Data / Mês", "Data Final", "Data do Pedido", "Data da Entrega", "Data Inserção"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    # numéricos
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
    # derivadas
    if "Data / Mês" in df.columns:
        df["Ano"] = df["Data / Mês"].dt.year
        df["Mes"] = df["Data / Mês"].dt.month
        df["Ano-Mes"] = df["Data / Mês"].dt.to_period("M").astype(str)
    if "Data do Pedido" in df.columns and "Data da Entrega" in df.columns:
        df["LeadTime (dias)"] = (df["Data da Entrega"] - df["Data do Pedido"]).dt.days
    if "Atrasado / No prazo" in df.columns:
        df["AtrasadoFlag"] = df["Atrasado / No prazo"].astype(str).str.contains("Atras", case=False, na=False)
    # chave pedido item para contagens robustas
    if "Pedido" in df.columns and "ITEM" in df.columns:
        df["PedidoItemKey"] = df["Pedido"].astype(str) + "||" + df["ITEM"].astype(str)
    return df

DATA_PATH = "Dashboard - Comite Semanal - Brasforma (1).xlsx"
st.sidebar.title("Filtros")
uploaded = st.sidebar.file_uploader("Base: Dashboard - Comitê Semanal - Brasforma (1).xlsx", type=["xlsx"], accept_multiple_files=False)
if uploaded is not None:
    DATA_PATH = uploaded
df = load_data(DATA_PATH)

# filtros
if "Data / Mês" in df.columns:
    min_date = pd.to_datetime(df["Data / Mês"]).min()
    max_date = pd.to_datetime(df["Data / Mês"]).max()
    d_ini, d_fim = st.sidebar.date_input("Período (Data / Mês)", value=(min_date, max_date))
else:
    d_ini = d_fim = None

reg = st.sidebar.multiselect("Regional", sorted(df["Regional"].dropna().unique()) if "Regional" in df.columns else [])
rep = st.sidebar.multiselect("Representante", sorted(df["Representante"].dropna().unique()) if "Representante" in df.columns else [])
uf  = st.sidebar.multiselect("UF", sorted(df["UF"].dropna().unique()) if "UF" in df.columns else [])
stat = st.sidebar.multiselect("Status Produção/Faturamento", sorted(df["Status de Produção / Faturamento"].dropna().unique()) if "Status de Produção / Faturamento" in df.columns else [])
cliente = st.sidebar.text_input("Cliente (contém)")
item = st.sidebar.text_input("SKU/Item (contém)")

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
    return flt

flt = apply_filters(df)

def fmt_money(v): 
    if pd.isna(v) or v is None: return "R$ 0"
    return f"R$ {v:,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")

def calc_kpis(_df):
    fat = _df["Valor Pedido R$"].sum() if "Valor Pedido R$" in _df.columns else np.nan
    n_ped = _df["Pedido"].nunique() if "Pedido" in _df.columns else len(_df)
    n_cli = _df["Nome Cliente"].nunique() if "Nome Cliente" in _df.columns else np.nan
    ticket = (fat / n_ped) if (n_ped and n_ped>0) else np.nan
    n_sku = _df["ITEM"].nunique() if "ITEM" in _df.columns else np.nan
    pct_atraso = np.nan
    if "AtrasadoFlag" in _df.columns:
        base = _df["AtrasadoFlag"].notna().sum()
        if base>0:
            pct_atraso = 100*_df["AtrasadoFlag"].mean()
    return fat, n_ped, n_cli, ticket, n_sku, pct_atraso

def month_period(s):
    try:
        y, m = s.split("-")
        return int(y), int(m)
    except:
        return None

def kpi_with_deltas(_df):
    fat, n_ped, n_cli, ticket, n_sku, pct_atraso = calc_kpis(_df)
    mom = yoy = None
    if "Ano-Mes" in _df.columns and "Valor Pedido R$" in _df.columns:
        gr = _df.groupby("Ano-Mes", as_index=False)["Valor Pedido R$"].sum()
        gr = gr.sort_values("Ano-Mes")
        if len(gr)>=2:
            curr = gr.iloc[-1]["Valor Pedido R$"]
            prev = gr.iloc[-2]["Valor Pedido R$"]
            if prev and prev!=0:
                mom = 100*(curr-prev)/prev
        last_label = gr.iloc[-1]["Ano-Mes"]
        mp = month_period(last_label)
        if mp:
            y, m = mp
            last_yoy = f"{y-1}-{str(m).zfill(2)}"
            base = gr[gr["Ano-Mes"]==last_yoy]
            if len(base)==1:
                prev_y = base.iloc[0]["Valor Pedido R$"]
                if prev_y and prev_y!=0:
                    yoy = 100*(curr - prev_y)/prev_y
    return (fat, n_ped, n_cli, ticket, n_sku, pct_atraso, mom, yoy)

tab_exec, tab_cli, tab_sku, tab_rep, tab_geo, tab_ops, tab_pareto, tab_export = st.tabs([
    "Visão Executiva", "Clientes", "Produtos", "Representantes", "Geografia", "Operacional", "Pareto/ABC", "Exportar"
])

with tab_exec:
    st.subheader("KPIs Executivos")
    fat, n_ped, n_cli, ticket, n_sku, pct_atraso, mom, yoy = kpi_with_deltas(flt)
    c1, c2, c3 = st.columns(3)
    c1.metric("Faturamento", fmt_money(fat), None if mom is None else f"{mom:,.1f}% MoM".replace(",", "X").replace(".", ",").replace("X","."))
    c2.metric("Pedidos", f"{n_ped:,}".replace(",", "."))
    c3.metric("Ticket Médio", fmt_money(ticket) if pd.notna(ticket) else "-",
              None if yoy is None else f"{yoy:,.1f}% YoY".replace(",", "X").replace(".", ",").replace("X","."))
    c4, c5, c6 = st.columns(3)
    c4.metric("Clientes Ativos", f"{n_cli:,}".replace(",", ".") if pd.notna(n_cli) else "-")
    c5.metric("SKUs Vendidos", f"{n_sku:,}".replace(",", ".") if pd.notna(n_sku) else "-")
    c6.metric("% Pedidos Atrasados", f"{pct_atraso:,.1f}%".replace(",", "X").replace(".", ",").replace("X",".") if pd.notna(pct_atraso) else "-")

    st.markdown("---")
    if "Ano-Mes" in flt.columns and "Valor Pedido R$" in flt.columns:
        serie = flt.groupby("Ano-Mes", as_index=False)["Valor Pedido R$"].sum().sort_values("Ano-Mes")
        st.subheader("Faturamento por Mês")
        chart = alt.Chart(serie).mark_bar().encode(
            x=alt.X("Ano-Mes:N", sort=None, title="Ano-Mês"),
            y=alt.Y("Valor Pedido R$:Q", title="Faturamento (R$)"),
            tooltip=["Ano-Mes", alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
        ).properties(height=300)
        st.altair_chart(chart, use_container_width=True)

with tab_cli:
    st.subheader("Top Clientes")
    if {"Nome Cliente","Valor Pedido R$"}.issubset(flt.columns):
        top_cli = flt.groupby("Nome Cliente", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        st.dataframe(top_cli.head(50))
        ch = alt.Chart(top_cli.head(15)).mark_bar().encode(
            x=alt.X("Valor Pedido R$:Q", title="Faturamento (R$)"),
            y=alt.Y("Nome Cliente:N", sort="-x", title="Cliente"),
            tooltip=["Nome Cliente", alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
        ).properties(height=480)
        st.altair_chart(ch, use_container_width=True)
        if "Data do Pedido" in flt.columns:
            base = flt.dropna(subset=["Nome Cliente","Data do Pedido"])
            first = base.groupby("Nome Cliente")["Data do Pedido"].min().rename("PrimeiraCompra")
            tmp = base.merge(first, on="Nome Cliente", how="left")
            tmp["MesesDesde1a"] = ((tmp["Data do Pedido"].dt.to_period("M") - tmp["PrimeiraCompra"].dt.to_period("M")).apply(lambda x: x.n)).astype("float")
            cohort = tmp.groupby(["Nome Cliente"])["MesesDesde1a"].max().rename("MesesAtivo").reset_index()
            st.caption("Atividade de clientes (tempo de vida em meses desde a 1ª compra):")
            st.dataframe(cohort.sort_values("MesesAtivo", ascending=False).head(50))

with tab_sku:
    st.subheader("Top SKUs")
    if {"ITEM","Valor Pedido R$"}.issubset(flt.columns):
        top_sku = flt.groupby("ITEM", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        st.dataframe(top_sku.head(100))
        ch = alt.Chart(top_sku.head(15)).mark_bar().encode(
            x=alt.X("Valor Pedido R$:Q", title="Faturamento (R$)"),
            y=alt.Y("ITEM:N", sort="-x", title="SKU"),
            tooltip=["ITEM", alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
        ).properties(height=480)
        st.altair_chart(ch, use_container_width=True)

with tab_rep:
    st.subheader("Faturamento por Representante")
    if {"Representante","Valor Pedido R$"}.issubset(flt.columns):
        por_rep = flt.groupby("Representante", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        st.dataframe(por_rep.head(100))
        ch = alt.Chart(por_rep.head(20)).mark_bar().encode(
            x=alt.X("Valor Pedido R$:Q", title="Faturamento (R$)"),
            y=alt.Y("Representante:N", sort="-x"),
            tooltip=["Representante", alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
        ).properties(height=520)
        st.altair_chart(ch, use_container_width=True)

with tab_geo:
    st.subheader("Faturamento por UF")
    if {"UF","Valor Pedido R$"}.issubset(flt.columns):
        por_uf = flt.groupby("UF", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        st.dataframe(por_uf)
        ch = alt.Chart(por_uf).mark_bar().encode(
            x=alt.X("UF:N", sort="-y"),
            y=alt.Y("Valor Pedido R$:Q", title="Faturamento (R$)"),
            tooltip=["UF", alt.Tooltip("Valor Pedido R$:Q", format=",.0f")]
        ).properties(height=360)
        st.altair_chart(ch, use_container_width=True)

with tab_ops:
    st.subheader("Operacional – Lead Time & Atraso")
    c1, c2 = st.columns(2)
    if "LeadTime (dias)" in flt.columns:
        with c1:
            lt = flt["LeadTime (dias)"].dropna()
            if len(lt)>0:
                lt_desc = pd.Series(lt).describe()[["count","mean","50%","min","max"]]
                st.dataframe(lt_desc.to_frame("LeadTime (dias)").rename(index={"50%":"mediana"}))
    if "Atrasado / No prazo" in flt.columns and "Pedido" in flt.columns:
        with c2:
            atrasos = flt.groupby("Atrasado / No prazo", as_index=False)["Pedido"].nunique()
            atrasos = atrasos.rename(columns={"Pedido":"Qtde Pedidos"})
            st.dataframe(atrasos)
            ch = alt.Chart(atrasos).mark_bar().encode(
                x=alt.X("Atrasado / No prazo:N", title="Status SLA"),
                y=alt.Y("Qtde Pedidos:Q"),
                tooltip=["Atrasado / No prazo","Qtde Pedidos"]
            ).properties(height=300)
            st.altair_chart(ch, use_container_width=True)

with tab_pareto:
    st.subheader("Pareto 80/20 e Curva ABC")
    if "Valor Pedido R$" in flt.columns:
        if "Nome Cliente" in flt.columns:
            g = flt.groupby("Nome Cliente", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
            g["%Acum"] = 100 * g["Valor Pedido R$"].cumsum() / g["Valor Pedido R$"].sum()
            def classe(p):
                if p <= 80: return "A"
                if p <= 95: return "B"
                return "C"
            g["Classe"] = g["%Acum"].apply(classe)
            st.markdown("**Clientes – Curva ABC**")
            st.dataframe(g.head(200))
            ch = alt.Chart(g).mark_line(point=True).encode(
                x=alt.X("Nome Cliente:N", sort=None, title="Cliente"),
                y=alt.Y("%Acum:Q", title="% Acumulado"),
                tooltip=["Nome Cliente", alt.Tooltip("%Acum:Q", format=",.1f")]
            ).properties(height=320)
            st.altair_chart(ch, use_container_width=True)
        if "ITEM" in flt.columns:
            s = flt.groupby("ITEM", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
            s["%Acum"] = 100 * s["Valor Pedido R$"].cumsum() / s["Valor Pedido R$"].sum()
            s["Classe"] = s["%Acum"].apply(lambda p: "A" if p<=80 else ("B" if p<=95 else "C"))
            st.markdown("**SKUs – Curva ABC**")
            st.dataframe(s.head(300))

with tab_export:
    st.subheader("Exportar dados filtrados")
    st.download_button("Baixar CSV filtrado", data=flt.to_csv(index=False).encode("utf-8-sig"), file_name="brasforma_filtrado.csv", mime="text/csv")
    with st.expander("Prévia dos dados filtrados"):
        st.dataframe(flt)

st.caption("Fonte: Carteira de Vendas – Dashboard Comitê Semanal – Brasforma | v2")
