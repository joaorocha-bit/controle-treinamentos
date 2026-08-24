# -*- coding: utf-8 -*-
"""
Dashboard de Controle de Treinamentos - Equipe de Transporte
Layout Fixo: Linhas 10 e 11 (Cabeçalhos) | Linhas 12 a 72 (Dados)
"""

import io
from datetime import datetime
import pandas as pd
import plotly.express as px
import requests
import streamlit as st

# --------------------------------------------------------------------------
# Configuração da Página
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Controle de Treinamentos - Transporte",
    page_icon="🚑",
    layout="wide",
)

VALIDADE_DIAS = 365

STATUS_CONCLUIDO = "Concluído"
STATUS_REFORCO = "Necessita de reforço"
STATUS_PENDENTE = "Pendente"

STATUS_COLORS = {
    STATUS_CONCLUIDO: "#2ecc71",
    STATUS_REFORCO: "#f1c40f",
    STATUS_PENDENTE: "#e74c3c",
    "Sem dado": "#bdc3c7",
}

GOOGLE_SHEET_ID = "1HQ9GRicZfVP_rUR51AxNpDaz_5GcZgAcBiP9qGkKZDk"
GOOGLE_SHEET_GID = "1064660493"
GOOGLE_SHEET_EXPORT_URL = (
    f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export"
    f"?format=xlsx&gid={GOOGLE_SHEET_GID}"
)


@st.cache_data(show_spinner="Baixando dados atualizados...", ttl=300)
def baixar_planilha(url: str) -> bytes:
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"
    }
    resposta = requests.get(url, timeout=30, headers=headers)
    resposta.raise_for_status()
    return resposta.content


def extrair_data(val):
    if pd.isna(val) or val is None:
        return None
    try:
        dt = pd.to_datetime(val, dayfirst=True, errors="coerce")
        if pd.notna(dt):
            return dt
    except Exception:
        pass
    return None


def limpar_status(val, tem_data_valida=False):
    if pd.isna(val) or val is None:
        return STATUS_CONCLUIDO if tem_data_valida else "Sem dado"
    val_str = str(val).strip()
    if not val_str or val_str.lower() in ["nan", "none", "", "nat", "0"]:
        return STATUS_CONCLUIDO if tem_data_valida else "Sem dado"

    v_low = val_str.lower()
    if any(k in v_low for k in ["conclu", "ok", "realizad", "sim", "100%", "100", "c", "feito"]):
        return STATUS_CONCLUIDO
    if any(k in v_low for k in ["refor", "atenc", "atenç", "alerta", "r"]):
        return STATUS_REFORCO
    if any(k in v_low for k in ["pend", "não", "nao", "falta", "a fazer", "p"]):
        return STATUS_PENDENTE

    return val_str


@st.cache_data(show_spinner="Processando planilha (linhas 10-11 e 12-72)...")
def processar_planilha(file_bytes: bytes):
    excel = pd.ExcelFile(io.BytesIO(file_bytes))
    df_raw = pd.read_excel(excel, sheet_name=0, header=None, nrows=72)

    # Linha 10 do Excel (Índice 9 no Python): Módulos/Grupos
    group_header = df_raw.iloc[9].fillna("").astype(str).str.strip()
    group_header_ffill = group_header.replace("", None).ffill().fillna("")

    # Linha 11 do Excel (Índice 10 no Python): Data / Status / Nome / Matrícula
    sub_header = df_raw.iloc[10].fillna("").astype(str).str.strip()

    col_matricula = 0
    col_nome = 1

    for idx, val in sub_header.items():
        v_low = str(val).lower()
        if any(k in v_low for k in ["matr", "re", "código", "codigo", "id"]):
            col_matricula = idx
        if any(k in v_low for k in ["nome", "colaborador", "funciona"]):
            col_nome = idx

    modulos = []
    cols_usadas = {col_matricula, col_nome}
    n_cols = len(df_raw.columns)
    i = 0

    while i < n_cols:
        if i in cols_usadas:
            i += 1
            continue

        gh_curr = group_header_ffill.iloc[i] if i < len(group_header_ffill) else ""
        sh_curr = sub_header.iloc[i] if i < len(sub_header) else ""

        gh_next = group_header_ffill.iloc[i + 1] if (i + 1 < len(group_header_ffill)) else ""
        sh_next = sub_header.iloc[i + 1] if (i + 1 < len(sub_header)) else ""

        sh_curr_low = sh_curr.lower()
        sh_next_low = sh_next.lower()

        is_data_1 = any(k in sh_curr_low for k in ["data", "dt", "realiza"])
        is_status_1 = any(k in sh_curr_low for k in ["status", "situa", "conceito"])
        is_data_2 = any(k in sh_next_low for k in ["data", "dt", "realiza"])
        is_status_2 = any(k in sh_next_low for k in ["status", "situa", "conceito"])

        tem_pareamento = (is_data_1 and is_status_2) or (is_status_1 and is_data_2)
        mesmo_grupo = (gh_curr != "" and gh_curr == gh_next)

        if (mesmo_grupo or tem_pareamento) and (i + 1 < n_cols) and (i + 1 not in cols_usadas):
            if is_status_1:
                c_st, c_dt = i, i + 1
            else:
                c_dt, c_st = i, i + 1

            label = gh_curr if gh_curr else (sh_curr if sh_curr not in ["Data", "Status", ""] else f"Módulo {len(modulos)+1}")
            modulos.append({"label": label, "col_data": c_dt, "col_status": c_st})
            cols_usadas.update([i, i + 1])
            i += 2
            continue

        label = gh_curr if gh_curr else sh_curr
        if label and label.lower() not in ["nan", "none", "unnamed", ""]:
            modulos.append({"label": label, "col_data": None, "col_status": i})
            cols_usadas.add(i)

        i += 1

    # Linhas 12 a 72 do Excel (Índices 11 a 71 no Python)
    df_dados = df_raw.iloc[11:72].reset_index(drop=True)
    registros = []

    for _, row in df_dados.iterrows():
        nome_raw = row[col_nome] if col_nome in row else None
        matr_raw = row[col_matricula] if col_matricula in row else None

        nome = str(nome_raw).strip() if pd.notna(nome_raw) else ""
        matr = str(matr_raw).strip() if pd.notna(matr_raw) else ""

        if not nome or nome.lower() in ["nan", "none", "0", "total", "subtotal", "nome"]:
            continue

        for m in modulos:
            c_dt = m["col_data"]
            c_st = m["col_status"]

            dt_raw = row[c_dt] if c_dt is not None and c_dt in row else None
            st_raw = row[c_st] if c_st is not None and c_st in row else None

            dt_parsed = extrair_data(dt_raw)
            tem_data = (dt_parsed is not None)

            status_limpo = limpar_status(st_raw, tem_data_valida=tem_data)

            if status_limpo == "Sem dado" and pd.notna(dt_raw) and not tem_data:
                status_limpo = limpar_status(dt_raw, tem_data_valida=False)

            registros.append({
                "Matrícula": matr,
                "Nome": nome,
                "Treinamento": m["label"],
                "Data": dt_parsed,
                "Status": status_limpo,
            })

    return pd.DataFrame(registros)


# --------------------------------------------------------------------------
# Carga de Dados
# --------------------------------------------------------------------------
st.sidebar.title("🔍 Filtros de Busca")

if st.sidebar.button("🔄 Recarregar Planilha", use_container_width=True):
    baixar_planilha.clear()
    processar_planilha.clear()
    st.rerun()

try:
    file_bytes = baixar_planilha(GOOGLE_SHEET_EXPORT_URL)
    df_long = processar_planilha(file_bytes)
except Exception as e:
    st.title("🚑 Controle de Treinamentos - Equipe de Transporte")
    st.error(f"Erro ao acessar a planilha: {e}")
    st.stop()

if df_long.empty:
    st.title("🚑 Controle de Treinamentos - Equipe de Transporte")
    st.warning("Nenhum registro de colaborador encontrado entre as linhas 12 e 72.")
    st.stop()

# Filtros simples na barra lateral
filtro_nomes = st.sidebar.multiselect("Filtrar por Colaborador", sorted(df_long["Nome"].unique()))
filtro_treinamentos = st.sidebar.multiselect("Filtrar por Módulo/Treinamento", sorted(df_long["Treinamento"].unique()))

df = df_long.copy()
if filtro_nomes:
    df = df[df["Nome"].isin(filtro_nomes)]
if filtro_treinamentos:
    df = df[df["Treinamento"].isin(filtro_treinamentos)]

# --------------------------------------------------------------------------
# Cálculos Globais
# --------------------------------------------------------------------------
hoje = pd.Timestamp(datetime.now().date())

# Filtro de Validade (1 Ano / 365 dias)
df_concluidos_com_data = df[(df["Status"] == STATUS_CONCLUIDO) & (df["Data"].notna())].copy()
df_concluidos_com_data["Vencimento"] = df_concluidos_com_data["Data"] + pd.Timedelta(days=VALIDADE_DIAS)
df_concluidos_com_data["Dias Restantes"] = (df_concluidos_com_data["Vencimento"] - hoje).dt.days

df_concluidos_com_data["Situação Vencimento"] = df_concluidos_com_data["Dias Restantes"].apply(
    lambda d: "Vencido" if d < 0 else ("A vencer (em até 60 dias)" if d <= 60 else "Dentro do prazo")
)

df_vencimento_alerta = df_concluidos_com_data[
    df_concluidos_com_data["Situação Vencimento"].isin(["Vencido", "A vencer (em até 60 dias)"])
].sort_values("Dias Restantes")

# Resumo por Módulo (% Capacitação)
resumo_modulo = df.groupby("Treinamento")["Status"].value_counts().unstack(fill_value=0)
for c in [STATUS_CONCLUIDO, STATUS_REFORCO, STATUS_PENDENTE, "Sem dado"]:
    if c not in resumo_modulo.columns:
        resumo_modulo[c] = 0

resumo_modulo["Total Colaboradores"] = resumo_modulo[[STATUS_CONCLUIDO, STATUS_REFORCO, STATUS_PENDENTE, "Sem dado"]].sum(axis=1)
resumo_modulo["% Capacitado"] = (resumo_modulo[STATUS_CONCLUIDO] / resumo_modulo["Total Colaboradores"] * 100).round(1)
tabela_capacitacao = resumo_modulo.reset_index().sort_values("% Capacitado", ascending=False)

# --------------------------------------------------------------------------
# PAINEL PRINCIPAL (LAYOUT FIXO)
# --------------------------------------------------------------------------
st.title("🚑 Controle de Treinamentos - Equipe de Transporte")
st.caption(f"Dados extraídos estritamente das linhas 12 a 72 da planilha • Validade configurada: **1 ano (365 dias)**")

# Resumo em métricas
k1, k2, k3, k4 = st.columns(4)
k1.metric("Total de Colaboradores", df["Nome"].nunique())
k2.metric("Total de Módulos", df["Treinamento"].nunique())
k3.metric("Treinamentos Concluídos", int((df["Status"] == STATUS_CONCLUIDO).sum()))
k4.metric("Treinamentos a Vencer / Vencidos", len(df_vencimento_alerta))

st.divider()

# ==========================================================================
# VISÃO 1: TREINAMENTOS A VENCER (1 ANO DE VALIDADE)
# ==========================================================================
st.header("1. ⏰ Treinamentos a Vencer (Validade: 1 Ano)")
st.write("Abaixo estão listados os treinamentos já realizados que estão **vencidos** ou **a vencer nos próximos 60 dias** (considerando o ciclo de 365 dias).")

if df_vencimento_alerta.empty:
    st.success("Nenhum treinamento vencido ou próximo do vencimento encontrado.")
else:
    df_exibicao_venc = df_vencimento_alerta[
        ["Matrícula", "Nome", "Treinamento", "Data", "Vencimento", "Dias Restantes", "Situação Vencimento"]
    ].copy()
    df_exibicao_venc["Data"] = df_exibicao_venc["Data"].dt.strftime("%d/%m/%Y")
    df_exibicao_venc["Vencimento"] = df_exibicao_venc["Vencimento"].dt.strftime("%d/%m/%Y")

    col_tabela, col_grafico = st.columns([0.6, 0.4])
    with col_tabela:
        st.dataframe(df_exibicao_venc, use_container_width=True, hide_index=True)
    with col_grafico:
        fig_venc = px.bar(
            df_vencimento_alerta.groupby(["Treinamento", "Situação Vencimento"]).size().reset_index(name="Quantidade"),
            x="Treinamento", y="Quantidade", color="Situação Vencimento",
            title="Alertas de Vencimento por Módulo",
            color_discrete_map={"Vencido": "#e74c3c", "A vencer (em até 60 dias)": "#f39c12"}
        )
        st.plotly_chart(fig_venc, use_container_width=True)

st.divider()

# ==========================================================================
# VISÃO 2: TREINAMENTOS REALIZADOS
# ==========================================================================
st.header("2. ✅ Treinamentos Realizados")
st.write("Visão geral dos treinamentos concluídos e comparativo da situação atual da equipe.")

col_realiz_1, col_realiz_2 = st.columns(2)

with col_realiz_1:
    df_concluidos_qtd = (
        df[df["Status"] == STATUS_CONCLUIDO]
        .groupby("Treinamento").size().reset_index(name="Concluídos")
        .sort_values("Concluídos", ascending=False)
    )
    fig_realiz = px.bar(
        df_concluidos_qtd, x="Treinamento", y="Concluídos",
        title="Total de Conclusões por Módulo",
        text_auto=True, color_discrete_sequence=["#2ecc71"]
    )
    st.plotly_chart(fig_realiz, use_container_width=True)

with col_realiz_2:
    df_distribuicao = df["Status"].value_counts().reset_index()
    df_distribuicao.columns = ["Status", "Quantidade"]
    fig_pie = px.pie(
        df_distribuicao, names="Status", values="Quantidade",
        title="Distribuição Geral dos Status de Treinamento",
        color="Status", color_discrete_map=STATUS_COLORS
    )
    st.plotly_chart(fig_pie, use_container_width=True)

st.divider()

# ==========================================================================
# VISÃO 3: % DA EQUIPE CAPACITADA EM CADA MÓDULO
# ==========================================================================
st.header("3. 📈 % da Equipe Capacitada por Módulo")
st.write("Percentual de colaboradores que concluíram com sucesso cada um dos módulos.")

col_cap_1, col_cap_2 = st.columns([0.5, 0.5])

with col_cap_1:
    fig_cap = px.bar(
        tabela_capacitacao, x="Treinamento", y="% Capacitado",
        title="Percentual de Capacitação (%)",
        color="% Capacitado", color_continuous_scale="RdYlGn", range_color=[0, 100],
        text_auto=".1f"
    )
    st.plotly_chart(fig_cap, use_container_width=True)

with col_cap_2:
    st.dataframe(
        tabela_capacitacao[["Treinamento", STATUS_CONCLUIDO, STATUS_PENDENTE, STATUS_REFORCO, "Total Colaboradores", "% Capacitado"]],
        use_container_width=True, hide_index=True
    )

st.divider()

# ==========================================================================
# VISÃO 4: STATUS DE CADA COLABORADOR POR TREINAMENTO
# ==========================================================================
st.header("4. 👤 Status de Cada Colaborador por Treinamento")
st.write("Matriz de acompanhamento individual. Cada linha representa um colaborador e cada coluna um módulo de treinamento.")

matriz_status = df.pivot_table(
    index=["Matrícula", "Nome"], columns="Treinamento", values="Status", aggfunc="first"
).reset_index()

def aplicar_cor_status(val):
    cor = STATUS_COLORS.get(val, "")
    return f"background-color: {cor}; color: white; font-weight: bold;" if cor else ""

cols_modulos = [c for c in matriz_status.columns if c not in ("Matrícula", "Nome")]
st.dataframe(
    matriz_status.style.map(aplicar_cor_status, subset=cols_modulos),
    use_container_width=True, hide_index=True
)
