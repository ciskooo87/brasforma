
# streamlit_app_brasforma_v5.py
import streamlit as st
import pandas as pd
import numpy as np
import altair as alt
from pathlib import Path

st.set_page_config(page_title="Brasforma – Dashboard Comercial v5", layout="wide")

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

def compute_rfm(_df, ref_date=None):
    base = _df.dropna(subset=["Nome Cliente"])
    if ref_date is None:
        if "Data do Pedido" in base.columns and base["Data do Pedido"].notna().any():
            ref_date = pd.to_datetime(base["Data do Pedido"]).max()
        elif "Data / Mês" in base.columns and base["Data / Mês"].notna().any():
            ref_date = pd.to_datetime(base["Data / Mês"]).max()
        else:
            ref_date = pd.Timestamp.today().normalize()
    if "Data do Pedido" in base.columns and base["Data do Pedido"].notna().any():
        last_buy = base.groupby("Nome Cliente")["Data do Pedido"].max().rename("UltimaCompra")
    else:
        last_buy = base.groupby("Nome Cliente")["Data / Mês"].max().rename("UltimaCompra")
    freq = base.groupby("Nome Cliente")["Pedido"].nunique().rename("Frequencia") if "Pedido" in base.columns else base.groupby("Nome Cliente").size().rename("Frequencia")
    val = base.groupby("Nome Cliente")["Valor Pedido R$"].sum().rename("Valor") if "Valor Pedido R$" in base.columns else None
    rfm = pd.concat([last_buy, freq, val], axis=1)
    rfm["RecenciaDias"] = (pd.to_datetime(ref_date) - pd.to_datetime(rfm["UltimaCompra"])).dt.days
    def safe_qcut(s, labels):
        try:
            return pd.qcut(s.rank(method="first"), q=len(labels), labels=labels)
        except Exception:
            return pd.Series([labels[len(labels)//2]]*len(s), index=s.index)
    rfm["R_Score"] = safe_qcut(-rfm["RecenciaDias"].fillna(rfm["RecenciaDias"].max()), labels=[1,2,3])
    rfm["F_Score"] = safe_qcut(rfm["Frequencia"].fillna(0), labels=[1,2,3])
    rfm["M_Score"] = safe_qcut(rfm["Valor"].fillna(0), labels=[1,2,3])
    rfm["Score"] = rfm[["R_Score","F_Score","M_Score"]].astype(int).sum(axis=1)
    def seg(row):
        r,f,m = int(row["R_Score"]), int(row["F_Score"]), int(row["M_Score"])
        if r>=3 and f>=3 and m>=3: return "Campeões"
        if f>=3 and r>=2: return "Leais"
        if r==1 and m>=2: return "Em risco"
        if r==1 and f==1: return "Perdidos"
        return "Oportunidades"
    rfm["Segmento"] = rfm.apply(seg, axis=1)
    rfm = rfm.sort_values(["Score","Valor","Frequencia"], ascending=[False,False,False]).reset_index()
    rfm.rename(columns={"index":"Nome Cliente"}, inplace=True)
    return rfm

tabs = st.tabs(["Visão Executiva","Clientes – RFM","Rentabilidade","Clientes","Produtos","Representantes","Geografia","Operacional","Pareto/ABC","Exportar"])
tab_exec, tab_rfm, tab_profit, tab_cli, tab_sku, tab_rep, tab_geo, tab_ops, tab_pareto, tab_export = tabs

with tab_exec:
    st.subheader("KPIs Executivos")
    fat, n_ped, n_cli, n_sku, ticket, lucro, margem_w, pct_rentavel = calc_kpis(flt)
    c1, c2, c3 = st.columns(3)
    c1.metric("Faturamento", fmt_money(fat))
    c2.metric("Pedidos", fmt_int(n_ped))
    c3.metric("Ticket Médio", fmt_money(ticket) if pd.notna(ticket) else "-")
    c4, c5, c6 = st.columns(3)
    c4.metric("Lucro Bruto", fmt_money(lucro))
    c5.metric("Margem Bruta (pond.)", fmt_pct(margem_w) if pd.notna(margem_w) else "-")
    c6.metric("% Itens Rentáveis", fmt_pct(pct_rentavel) if pd.notna(pct_rentavel) else "-")

with tab_rfm:
    st.subheader("Clientes – RFM (Recência, Frequência, Valor)")
    rfm = compute_rfm(flt, ref_date=None)
    segs = sorted(rfm["Segmento"].unique())
    pick = st.multiselect("Segmentos", segs, default=segs)
    view = rfm[rfm["Segmento"].isin(pick)]
    st.metric("Clientes avaliados", fmt_int(len(view)))
    cols = ["Nome Cliente","RecenciaDias","Frequencia","Valor","R_Score","F_Score","M_Score","Score","Segmento"]
    display_table(view[cols], money_cols=["Valor"], int_cols=["RecenciaDias","Frequencia","Score"])

with tab_profit:
    st.subheader("Rentabilidade – Lucro e Margem")
    if {"Nome Cliente","Lucro Bruto"}.issubset(flt.columns):
        top_cli = flt.groupby("Nome Cliente", as_index=False)["Lucro Bruto"].sum().sort_values("Lucro Bruto", ascending=False).head(20)
        display_table(top_cli, money_cols=["Lucro Bruto"])
    if {"ITEM","Lucro Bruto"}.issubset(flt.columns):
        top_sku = flt.groupby("ITEM", as_index=False)["Lucro Bruto"].sum().sort_values("Lucro Bruto", ascending=False).head(20)
        display_table(top_sku, money_cols=["Lucro Bruto"])
    if {"Representante","Lucro Bruto","Valor Pedido R$"}.issubset(flt.columns):
        por_rep = flt.groupby("Representante", as_index=False).agg({"Lucro Bruto":"sum","Valor Pedido R$":"sum"})
        por_rep["Margem %"] = np.where(por_rep["Valor Pedido R$"]>0, 100.0*por_rep["Lucro Bruto"]/por_rep["Valor Pedido R$"], np.nan)
        display_table(por_rep.sort_values("Lucro Bruto", ascending=False), money_cols=["Lucro Bruto","Valor Pedido R$"], pct_cols=["Margem %"])
    if "Lucro Bruto" in flt.columns:
        neg = flt[flt["Lucro Bruto"] < 0].copy()
        st.markdown("#### Auditoria – Linhas com Margem Negativa")
        cols_show = [c for c in ["Nome Cliente","Pedido","ITEM","Representante","UF","Valor Pedido R$","Custo","Custo Total","Lucro Bruto","Margem %","Data do Pedido","Data / Mês"] if c in neg.columns]
        display_table(neg[cols_show], money_cols=["Valor Pedido R$","Custo","Custo Total","Lucro Bruto"], pct_cols=["Margem %"])

with tab_cli:
    st.subheader("Clientes – Faturamento")
    if {"Nome Cliente","Valor Pedido R$"}.issubset(flt.columns):
        top_cli = flt.groupby("Nome Cliente", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        display_table(top_cli.head(50), money_cols=["Valor Pedido R$"])

with tab_sku:
    st.subheader("Produtos – Faturamento")
    if {"ITEM","Valor Pedido R$"}.issubset(flt.columns):
        top_sku = flt.groupby("ITEM", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        display_table(top_sku.head(100), money_cols=["Valor Pedido R$"])

with tab_rep:
    st.subheader("Representantes – Faturamento")
    if {"Representante","Valor Pedido R$"}.issubset(flt.columns):
        por_rep = flt.groupby("Representante", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        display_table(por_rep.head(100), money_cols=["Valor Pedido R$"])

with tab_geo:
    st.subheader("Geografia – Faturamento por UF")
    if {"UF","Valor Pedido R$"}.issubset(flt.columns):
        por_uf_fat = flt.groupby("UF", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
        display_table(por_uf_fat, money_cols=["Valor Pedido R$"])

with tab_ops:
    st.subheader("Operacional – Lead Time & Atraso")
    if "LeadTime (dias)" in flt.columns:
        lt = flt["LeadTime (dias)"].dropna()
        if len(lt)>0:
            desc = pd.Series(lt).describe()[["count","mean","50%","min","max"]].rename({"50%":"mediana"})
            display_table(desc.to_frame("LeadTime (dias)").T, int_cols=["count","min","max"])
    if "Atrasado / No prazo" in flt.columns and "Pedido" in flt.columns:
        atrasos = flt.groupby("Atrasado / No prazo", as_index=False)["Pedido"].nunique().rename(columns={"Pedido":"Qtde Pedidos"})
        display_table(atrasos, int_cols=["Qtde Pedidos"])

with tab_pareto:
    st.subheader("Pareto 80/20 e Curva ABC (Faturamento)")
    if "Valor Pedido R$" in flt.columns:
        if "Nome Cliente" in flt.columns:
            g = flt.groupby("Nome Cliente", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
            g["%Acum"] = 100 * g["Valor Pedido R$"].cumsum() / g["Valor Pedido R$"].sum()
            g["Classe"] = g["%Acum"].apply(lambda p: "A" if p<=80 else ("B" if p<=95 else "C"))
            display_table(g.head(200), money_cols=["Valor Pedido R$"], pct_cols=["%Acum"])
        if "ITEM" in flt.columns:
            s = flt.groupby("ITEM", as_index=False)["Valor Pedido R$"].sum().sort_values("Valor Pedido R$", ascending=False)
            s["%Acum"] = 100 * s["Valor Pedido R$"].cumsum() / s["Valor Pedido R$"].sum()
            s["Classe"] = s["%Acum"].apply(lambda p: "A" if p<=80 else ("B" if p<=95 else "C"))
            display_table(s.head(300), money_cols=["Valor Pedido R$"], pct_cols=["%Acum"])

with tab_export:
    st.subheader("Exportar")
    st.download_button("Baixar CSV filtrado", data=flt.to_csv(index=False).encode("utf-8-sig"), file_name="brasforma_filtrado.csv", mime="text/csv")
    with st.expander("Prévia dos dados filtrados"):
        st.dataframe(flt)

if qty_col:
    st.caption(f"✓ Custo calculado como **unitário × quantidade**. Coluna de quantidade: **{qty_col}**.")
else:
    st.caption("! Atenção: coluna de quantidade não identificada — usando Custo como total.")
