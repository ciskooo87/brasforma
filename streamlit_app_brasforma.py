# streamlit_app_brasforma_v4.py
# Dashboard Comercial – Brasforma | v4 (Rentabilidade com Custo)
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pathlib import Path

st.set_page_config(page_title="Brasforma – Dashboard Comercial v4", layout="wide")

def to_num_generic(x):
    if pd.isna(x):
        return np.nan
    if isinstance(x, (int, float, np.integer, np.floating)):
        return x
    s = str(x).strip().replace(".", "").replace(",", ".")
    try:
        return float(s)
    except Exception:
        return np.nan

@st.cache_data(show_spinner=False)
def load_data(path: str, sheet_name: str = "Carteira de Vendas"):
    try:
        xls = pd.ExcelFile(path, engine="openpyxl")
    except Exception as e:
        st.error("Falha ao abrir Excel. Verifique se o arquivo é .xlsx válido (engine openpyxl).")
        st.exception(e)
        st.stop()
    df = pd.read_excel(xls, sheet_name=sheet_name)
    df.columns = [c.strip() for c in df.columns]
    for col in ["Data / Mês", "Data Final", "Data do Pedido", "Data da Entrega", "Data Inserção"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")
    for col in ["Valor Pedido R$", "TICKET MÉDIO", "Quant. Pedidos", "Custo"]:
        if col in df.columns:
            if col == "Quant. Pedidos":
                df[col] = pd.to_numeric(df[col], errors="coerce")
            else:
                df[col] = df[col].apply(to_num_generic)
    if "Data / Mês" in df.columns:
        df["Ano"] = df["Data / Mês"].dt.year
        df["Mes"] = df["Data / Mês"].dt.month
        df["Ano-Mes"] = df["Data / Mês"].dt.to_period("M").astype(str)
    if "Data do Pedido" in df.columns and "Data da Entrega" in df.columns:
        df["LeadTime (dias)"] = (df["Data da Entrega"] - df["Data do Pedido"]).dt.days
    if "Atrasado / No prazo" in df.columns:
        df["AtrasadoFlag"] = df["Atrasado / No prazo"].astype(str).str.contains("Atras", case=False, na=False)
    if "Valor Pedido R$" in df.columns and "Custo" in df.columns:
        df["Lucro Bruto"] = df["Valor Pedido R$"] - df["Custo"]
        df["Margem %"] = np.where(df["Valor Pedido R$"]>0, 100.0 * df["Lucro Bruto"]/df["Valor Pedido R$"], np.nan)
    else:
        df["Lucro Bruto"] = np.nan
        df["Margem %"] = np.nan
    if "Pedido" in df.columns and "ITEM" in df.columns:
        df["PedidoItemKey"] = df["Pedido"].astype(str) + "||" + df["ITEM"].astype(str)
    return df

DEFAULT_DATA = "Dashboard - Comite Semanal - Brasforma IA (1).xlsx"
ALT_DATA = "Dashboard - Comite Semanal - Brasforma (1).xlsx"

st.sidebar.title("Fonte de dados")
uploaded = st.sidebar.file_uploader("Envie a base (.xlsx)", type=["xlsx"], accept_multiple_files=False)
data_path = uploaded if uploaded is not None else (DEFAULT_DATA if Path(DEFAULT_DATA).exists() else ALT_DATA)
st.sidebar.caption(f"Arquivo em uso: **{data_path}**")

df = load_data(data_path)

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
    lucro = _df["Lucro Bruto"].sum() if "Lucro Bruto" in _df.columns else np.nan
    margem_pct_weighted = np.nan
    if "Valor Pedido R$" in _df.columns and _df["Valor Pedido R$"].sum() > 0:
        margem_pct_weighted = 100.0 * (lucro / _df["Valor Pedido R$"].sum())
    ticket_margem = (lucro / n_ped) if (n_ped and n_ped>0) else np.nan
    pct_rentavel = np.nan
    if "Lucro Bruto" in _df.columns:
        base_itens = len(_df)
        if base_itens>0:
            pct_rentavel = 100.0 * (_df["Lucro Bruto"] > 0).sum() / base_itens
    return fat, n_ped, n_cli, ticket, n_sku, pct_atraso, lucro, margem_pct_weighted, ticket_margem, pct_rentavel

tabs = st.tabs(["Visão Executiva", "Rentabilidade", "Clientes", "Produtos", "Representantes", "Geografia", "Operacional", "Pareto/ABC", "Exportar"])
tab_exec, tab_profit, tab_cli, tab_sku, tab_rep, tab_geo, tab_ops, tab_pareto, tab_export = tabs

with tab_exec:
    st.subheader("KPIs Executivos")
    fat, n_ped, n_cli, ticket, n_sku, pct_atraso, lucro, margem_w, ticket_margem, pct_rentavel = calc_kpis(flt)
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Faturamento", fmt_money(fat))
    c2.metric("Pedidos", f"{n_ped:,}".replace(",", "."))
    c3.metric("Ticket Médio", fmt_money(ticket) if pd.notna(ticket) else "-")
    c4.metric("Lucro Bruto", fmt_money(lucro) if pd.notna(lucro) else "-")
    c5.metric("Margem Bruta (pond.)", f"{margem_w:,.1f}%".replace(",", "X").replace(".", ",").replace("X",".") if pd.notna(margem_w) else "-")
    c6, c7, c8 = st.columns(3)
    c6.metric("Clientes", f"{n_cli:,}".replace(",", ".") if pd.notna(n_cli) else "-")
    c7.metric("SKUs Vendidos", f"{n_sku:,}".replace(",", ".") if pd.notna(n_sku) else "-")
    c8.metric("% Itens Rentáveis", f"{pct_rentavel:,.1f}%".replace(",", "X").replace(".", ",").replace("X",".") if pd.notna(pct_rentavel) else "-")
    st.markdown("---")
    if "Ano-Mes" in flt.columns and "Valor Pedido R$" in flt.columns and "Lucro Bruto" in flt.columns:
        serie = flt.groupby("Ano-Mes", as_index=False).agg({"Valor Pedido R$":"sum","Lucro Bruto":"sum"}).sort_values("Ano-Mes")
        base = alt.Chart(serie).encode(x=alt.X("Ano-Mes:N", sort=None, title="Ano-Mês"))
        ch1 = base.mark_bar().encode(y=alt.Y("Valor Pedido R$:Q", title="Faturamento (R$)"), tooltip=[alt.Tooltip("Valor Pedido R$:Q", format=",.0f")])
        ch2 = base.mark_line(point=True).encode(y=alt.Y("Lucro Bruto:Q", title="Lucro Bruto (R$)"), tooltip=[alt.Tooltip("Lucro Bruto:Q", format=",.0f")])
        st.subheader("Faturamento e Lucro Bruto por Mês")
        st.altair_chart(ch1 + ch2, use_container_width=True)

with tab_profit:
    st.subheader("Rentabilidade – Lucro e Margem")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Lucro Bruto Total", fmt_money(flt["Lucro Bruto"].sum()) if "Lucro Bruto" in flt.columns else "-")
    if "Valor Pedido R$" in flt.columns and flt["Valor Pedido R$"].sum() > 0:
        margem_total = 100.0 * flt["Lucro Bruto"].sum()/flt["Valor Pedido R$"].sum()
        c2.metric("Margem Bruta Total", f"{margem_total:,.1f}%".replace(",", "X").replace(".", ",").replace("X","."))
    else:
        c2.metric("Margem Bruta Total", "-")
    c3.metric("Ticket de Margem", fmt_money(flt["Lucro Bruto"].sum()/flt["Pedido"].nunique()) if "Pedido" in flt.columns and flt["Pedido"].nunique()>0 else "-")
    c4.metric("% Linhas Negativas", f"{100.0*(flt['Lucro Bruto']<0).mean():.1%}".replace(".", ",") if "Lucro Bruto" in flt.columns and len(flt)>0 else "-")

    if {"Nome Cliente","Lucro Bruto"}.issubset(flt.columns):
        top_cli_lucro = flt.groupby("Nome Cliente", as_index=False)["Lucro Bruto"].sum().sort_values("Lucro Bruto", ascending=False).head(20)
        st.markdown("#### Top 20 – **Clientes** por Lucro Bruto")
        st.dataframe(top_cli_lucro)
        st.altair_chart(alt.Chart(top_cli_lucro).mark_bar().encode(x=alt.X("Lucro Bruto:Q", title="Lucro Bruto (R$)"), y=alt.Y("Nome Cliente:N", sort="-x")), use_container_width=True)

    if {"ITEM","Lucro Bruto"}.issubset(flt.columns):
        top_sku_lucro = flt.groupby("ITEM", as_index=False)["Lucro Bruto"].sum().sort_values("Lucro Bruto", ascending=False).head(20)
        st.markdown("#### Top 20 – **SKUs** por Lucro Bruto")
        st.dataframe(top_sku_lucro)
        st.altair_chart(alt.Chart(top_sku_lucro).mark_bar().encode(x=alt.X("Lucro Bruto:Q", title="Lucro Bruto (R$)"), y=alt.Y("ITEM:N", sort="-x")), use_container_width=True)

    if {"Representante","Lucro Bruto","Valor Pedido R$"}.issubset(flt.columns):
        st.markdown("#### Margem por Representante (Top 20 por Lucro)")
        por_rep = flt.groupby("Representante", as_index=False).agg({"Lucro Bruto":"sum","Valor Pedido R$":"sum"})
        por_rep["Margem %"] = np.where(por_rep["Valor Pedido R$"]>0, 100.0*por_rep["Lucro Bruto"]/por_rep["Valor Pedido R$"], np.nan)
        por_rep = por_rep.sort_values("Lucro Bruto", ascending=False).head(20)
        st.dataframe(por_rep)

    if {"Nome Cliente","Valor Pedido R$","Lucro Bruto"}.issubset(flt.columns):
        st.markdown("#### Dispersão – Valor vs Margem (%) por Cliente")
        disp = flt.groupby("Nome Cliente", as_index=False).agg({"Valor Pedido R$":"sum","Lucro Bruto":"sum"})
        disp["Margem %"] = np.where(disp["Valor Pedido R$"]>0, 100.0*disp["Lucro Bruto"]/disp["Valor Pedido R$"], np.nan)
        st.altair_chart(alt.Chart(disp).mark_circle(size=70).encode(x="Valor Pedido R$:Q", y="Margem %:Q"), use_container_width=True)

    if {"UF","Lucro Bruto","Valor Pedido R$"}.issubset(flt.columns):
        st.markdown("#### Heatmap – Margem por UF")
        por_uf = flt.groupby("UF", as_index=False).agg({"Lucro Bruto":"sum","Valor Pedido R$":"sum"})
        por_uf["Margem %"] = np.where(por_uf["Valor Pedido R$"]>0, 100.0*por_uf["Lucro Bruto"]/por_uf["Valor Pedido R$"], np.nan)
        st.dataframe(por_uf.sort_values("Margem %", ascending=False))

    if "Lucro Bruto" in flt.columns:
        st.markdown("#### Auditoria – Linhas com Margem Negativa")
        neg = flt[flt["Lucro Bruto"] < 0].copy()
        st.caption(f"{len(neg):,}".replace(",", ".") + " linhas com margem negativa no filtro atual.")
        cols_show = [c for c in ["Nome Cliente","Pedido","ITEM","Representante","UF","Valor Pedido R$","Custo","Lucro Bruto","Margem %","Data do Pedido","Data / Mês"] if c in neg.columns]
        st.dataframe(neg[cols_show].head(500))

with tab_cli:
    st.subheader("Clientes – Faturamento")
    if {"Nome Cliente","Valor Pedido R$"}.issubset(flt.columns):
        top_cli = flt.groupby("Nome Cliente", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        st.dataframe(top_cli.head(50))

with tab_sku:
    st.subheader("Produtos – Faturamento")
    if {"ITEM","Valor Pedido R$"}.issubset(flt.columns):
        top_sku = flt.groupby("ITEM", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        st.dataframe(top_sku.head(100))

with tab_rep:
    st.subheader("Representantes – Faturamento")
    if {"Representante","Valor Pedido R$"}.issubset(flt.columns):
        por_rep = flt.groupby("Representante", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        st.dataframe(por_rep.head(100))

with tab_geo:
    st.subheader("Geografia – Faturamento por UF")
    if {"UF","Valor Pedido R$"}.issubset(flt.columns):
        por_uf_fat = flt.groupby("UF", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        st.dataframe(por_uf_fat)

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
            atrasos = flt.groupby("Atrasado / No prazo", as_index=False)["Pedido"].nunique().rename(columns={"Pedido":"Qtde Pedidos"})
            st.dataframe(atrasos)

with tab_pareto:
    st.subheader("Pareto 80/20 e Curva ABC (Faturamento)")
    if "Valor Pedido R$" in flt.columns:
        if "Nome Cliente" in flt.columns:
            g = flt.groupby("Nome Cliente", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
            g["%Acum"] = 100 * g["Valor Pedido R$"].cumsum() / g["Valor Pedido R$"].sum()
            g["Classe"] = g["%Acum"].apply(lambda p: "A" if p<=80 else ("B" if p<=95 else "C"))
            st.markdown("**Clientes – Curva ABC**")
            st.dataframe(g.head(200))
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

st.caption("Fonte: Carteira de Vendas – Dashboard Comitê Semanal – Brasforma | v4 (Rentabilidade com Custo)")
