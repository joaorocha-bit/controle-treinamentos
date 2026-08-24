# -*- coding: utf-8 -*-
"""
Dashboard de Controle de Treinamentos - Equipe de Transporte
--------------------------------------------------------------
Lê a "Planilha Geral de Controle" (mesmo layout do Google Sheets original,
com módulos em pares Data/Status e colunas de status único ao final),
e apresenta:
  - Treinamentos a vencer (validade de 1 ano)
  - Treinamentos realizados
  - % da equipe capacitada por módulo
  - Status de cada colaborador por treinamento

Layout, proporções e filtros são configuráveis pela barra lateral.
"""

import io
import json
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import plotly.express as px
import requests
import streamlit as st

from data_parser import (
    ler_planilha as _ler_planilha,
    parse_estrutura,
    identificar_colunas,
    montar_dataframe_longo,
    limpar_status,
)

# --------------------------------------------------------------------------
# Configuração geral da página
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Controle de Treinamentos - Transporte",
    page_icon="🚑",
    layout="wide",
)

CONFIG_PATH = Path("data/layout_config.json")
VALIDADE_DIAS = 365  # validade de 1 ano

# --------------------------------------------------------------------------
# Fonte de dados: Google Sheets (planilha pública, buscada automaticamente)
# --------------------------------------------------------------------------
GOOGLE_SHEET_ID = "1HQ9GRicZfVP_rUR51AxNpDaz_5GcZgAcBiP9qGkKZDk"
GOOGLE_SHEET_GID = "1064660493"
GOOGLE_SHEET_EXPORT_URL = (
    f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export"
    f"?format=xlsx&gid={GOOGLE_SHEET_GID}"
)


@st.cache_data(show_spinner="Buscando planilha atualizada do Google Sheets...", ttl=300)
def baixar_planilha_google(url: str) -> bytes:
    """Baixa a planilha do Google Sheets como .xlsx (a planilha precisa estar
    compartilhada como 'Qualquer pessoa com o link pode visualizar')."""
    resposta = requests.get(url, timeout=30)
    resposta.raise_for_status()
    return resposta.content

STATUS_CONCLUIDO = "Concluído"
STATUS_REFORCO = "Necessita de reforço"
STATUS_PENDENTE = "Pendente"

STATUS_COLORS = {
    STATUS_CONCLUIDO: "#2ecc71",
    STATUS_REFORCO: "#f1c40f",
    STATUS_PENDENTE: "#e74c3c",
    "Sem dado": "#bdc3c7",
}

SECOES_PADRAO = [
    "Indicadores gerais",
    "Treinamentos a vencer",
    "Treinamentos realizados",
    "% capacitação por módulo",
    "Status por colaborador",
]


# --------------------------------------------------------------------------
# Config de layout (persistida em disco quando possível; sempre em sessão)
# --------------------------------------------------------------------------
def carregar_config():
    if CONFIG_PATH.exists():
        try:
            return json.loads(CONFIG_PATH.read_text(encoding="utf-8"))
        except Exception:
            pass
    return {}


def salvar_config(cfg: dict):
    try:
        CONFIG_PATH.parent.mkdir(parents=True, exist_ok=True)
        CONFIG_PATH.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        return True
    except Exception:
        return False


if "config" not in st.session_state:
    st.session_state.config = carregar_config()


# --------------------------------------------------------------------------
# Leitura e parsing dinâmico da planilha
# --------------------------------------------------------------------------
@st.cache_data(show_spinner=False)
def ler_planilha(file_bytes: bytes, nome_aba: str | None, linha_max: int):
    """Lê a planilha e devolve (df_bruto, nome_da_aba_usada, lista_de_abas). Cacheado pelo Streamlit."""
    return _ler_planilha(file_bytes, nome_aba, linha_max)


# --------------------------------------------------------------------------
# Sidebar: fonte de dados
# --------------------------------------------------------------------------
st.sidebar.title("⚙️ Configurações")

st.sidebar.subheader("📄 Fonte de dados")
st.sidebar.caption("Planilha lida automaticamente do Google Sheets.")
linha_max = st.sidebar.number_input(
    "Última linha com dados (planilha)", min_value=2, value=72, step=1,
    help="Número da linha do Excel (contando do topo) onde os dados terminam.",
)
if st.sidebar.button("🔄 Atualizar dados agora", use_container_width=True):
    baixar_planilha_google.clear()
    st.rerun()

file_bytes = None
erro_download = None
try:
    file_bytes = baixar_planilha_google(GOOGLE_SHEET_EXPORT_URL)
except Exception as e:
    erro_download = str(e)

if file_bytes is None:
    st.title("🚑 Controle de Treinamentos - Equipe de Transporte")
    st.warning("Não foi possível carregar a planilha do Google Sheets.")
    st.info(
        "Verifique se a planilha está compartilhada como **'Qualquer pessoa com "
        "o link pode visualizar'** (Compartilhar → Acesso geral) e se o link/ID "
        "e o GID configurados no código estão corretos."
    )
    if erro_download:
        st.caption(f"Detalhe técnico: {erro_download}")
    st.stop()

df_bruto, aba_usada, abas_disponiveis = ler_planilha(file_bytes, st.session_state.config.get("aba"), linha_max)

aba_escolhida = st.sidebar.selectbox(
    "Aba da planilha", abas_disponiveis, index=abas_disponiveis.index(aba_usada)
)
if aba_escolhida != aba_usada:
    df_bruto, aba_usada, abas_disponiveis = ler_planilha(file_bytes, aba_escolhida, linha_max)

grupo, sub, dados, dados_idx = parse_estrutura(df_bruto)
col_matricula, col_nome, modulos, unicos = identificar_colunas(grupo, sub)

if not modulos and not unicos:
    st.error(
        "Não foi possível identificar as colunas de módulos/status automaticamente. "
        "Verifique se a aba selecionada segue o layout esperado."
    )
    st.stop()

df_long = montar_dataframe_longo(dados, col_matricula, col_nome, modulos, unicos)
df_long = df_long[df_long["Nome"].notna()]

# --------------------------------------------------------------------------
# Sidebar: nomes editáveis dos módulos
# --------------------------------------------------------------------------
st.sidebar.subheader("✏️ Nomes exibidos dos treinamentos")
st.sidebar.caption("Alguns rótulos vêm duplicados na planilha original (ex.: dois blocos 'M2'/'M3'). Renomeie aqui se quiser deixá-los mais claros.")

nomes_editados = st.session_state.config.get("nomes_editados", {})
with st.sidebar.expander("Editar rótulos", expanded=False):
    todos_labels = [m["label"] for m in modulos] + [u["label"] for u in unicos]
    for label in todos_labels:
        novo = st.text_input(label, value=nomes_editados.get(label, label), key=f"label_{label}")
        nomes_editados[label] = novo

df_long["Treinamento"] = df_long["Treinamento"].map(lambda x: nomes_editados.get(x, x))

col1, col2 = st.sidebar.columns(2)
if col1.button("💾 Salvar layout", use_container_width=True):
    st.session_state.config["nomes_editados"] = nomes_editados
    ok = salvar_config(st.session_state.config)
    st.sidebar.success("Salvo!" if ok else "Salvo nesta sessão (disco não gravável neste ambiente).")
if col2.button("↩️ Restaurar padrão", use_container_width=True):
    st.session_state.config["nomes_editados"] = {}
    salvar_config(st.session_state.config)
    st.rerun()

# --------------------------------------------------------------------------
# Sidebar: filtros
# --------------------------------------------------------------------------
st.sidebar.subheader("🔎 Filtros")
lista_nomes = sorted(df_long["Nome"].dropna().unique().tolist())
lista_treinamentos = sorted(df_long["Treinamento"].dropna().unique().tolist())
lista_status = sorted(df_long["Status"].dropna().unique().tolist())

filtro_nomes = st.sidebar.multiselect("Colaborador(es)", lista_nomes)
filtro_treinamentos = st.sidebar.multiselect("Treinamento(s)/Módulo(s)", lista_treinamentos)
filtro_status = st.sidebar.multiselect("Status", lista_status)
dias_alerta = st.sidebar.slider(
    "Alertar vencimento em até (dias)", min_value=15, max_value=180, value=60, step=15
)

df_filtrado = df_long.copy()
if filtro_nomes:
    df_filtrado = df_filtrado[df_filtrado["Nome"].isin(filtro_nomes)]
if filtro_treinamentos:
    df_filtrado = df_filtrado[df_filtrado["Treinamento"].isin(filtro_treinamentos)]
if filtro_status:
    df_filtrado = df_filtrado[df_filtrado["Status"].isin(filtro_status)]

# --------------------------------------------------------------------------
# Sidebar: layout (seções exibidas, ordem e proporções)
# --------------------------------------------------------------------------
st.sidebar.subheader("🧩 Layout do dashboard")
secoes_ativas = st.sidebar.multiselect(
    "Seções exibidas (a ordem escolhida define a ordem na página)",
    SECOES_PADRAO,
    default=st.session_state.config.get("secoes_ativas", SECOES_PADRAO),
)
n_col_kpi = st.sidebar.slider("Nº de colunas nos indicadores", 2, 5, st.session_state.config.get("n_col_kpi", 4))
prop_grafico_esq = st.sidebar.slider(
    "Proporção coluna esquerda (gráficos lado a lado)", 0.2, 0.8,
    float(st.session_state.config.get("prop_grafico_esq", 0.5)), 0.05,
)

if st.sidebar.button("💾 Salvar preferências de layout", use_container_width=True):
    st.session_state.config.update(
        {"secoes_ativas": secoes_ativas, "n_col_kpi": n_col_kpi, "prop_grafico_esq": prop_grafico_esq, "aba": aba_escolhida}
    )
    ok = salvar_config(st.session_state.config)
    st.sidebar.success("Preferências salvas!" if ok else "Salvo nesta sessão.")

# --------------------------------------------------------------------------
# Cálculos
# --------------------------------------------------------------------------
df_modulos_apenas = df_filtrado[df_filtrado["Tipo"] == "Módulo"].copy()

# Vencimento = Data de conclusão + 365 dias, apenas para status Concluído com data válida
df_venc = df_modulos_apenas[
    (df_modulos_apenas["Status"] == STATUS_CONCLUIDO) & (df_modulos_apenas["Data"].notna())
].copy()
hoje = pd.Timestamp(datetime.now().date())
df_venc["Vencimento"] = df_venc["Data"] + pd.Timedelta(days=VALIDADE_DIAS)
df_venc["Dias restantes"] = (df_venc["Vencimento"] - hoje).dt.days
df_venc["Situação vencimento"] = df_venc["Dias restantes"].apply(
    lambda d: "Vencido" if d < 0 else ("A vencer" if d <= dias_alerta else "Dentro do prazo")
)
df_a_vencer = df_venc[df_venc["Situação vencimento"].isin(["Vencido", "A vencer"])].sort_values("Dias restantes")

total_colaboradores = df_long["Nome"].nunique()

# % capacitação por módulo/treinamento (sobre o total de colaboradores)
resumo_modulo = (
    df_filtrado.groupby("Treinamento")["Status"]
    .value_counts(normalize=False)
    .unstack(fill_value=0)
)
for c in [STATUS_CONCLUIDO, STATUS_REFORCO, STATUS_PENDENTE, "Sem dado"]:
    if c not in resumo_modulo.columns:
        resumo_modulo[c] = 0
resumo_modulo["Total"] = resumo_modulo[[STATUS_CONCLUIDO, STATUS_REFORCO, STATUS_PENDENTE, "Sem dado"]].sum(axis=1)
resumo_modulo["% Concluído"] = (resumo_modulo[STATUS_CONCLUIDO] / resumo_modulo["Total"] * 100).round(1)

# --------------------------------------------------------------------------
# Cabeçalho
# --------------------------------------------------------------------------
st.title("🚑 Controle de Treinamentos - Equipe de Transporte")
st.caption(f"Fonte: aba **{aba_usada}** • {total_colaboradores} colaboradores • dados até a linha {linha_max}")


def render_indicadores_gerais():
    st.subheader("📊 Indicadores gerais")
    total_treinos = len(df_modulos_apenas["Treinamento"].unique())
    concluidos = (df_modulos_apenas["Status"] == STATUS_CONCLUIDO).sum()
    pendentes = (df_modulos_apenas["Status"] == STATUS_PENDENTE).sum()
    pct_geral = (
        df_modulos_apenas["Status"].eq(STATUS_CONCLUIDO).sum() / len(df_modulos_apenas) * 100
        if len(df_modulos_apenas) else 0
    )
    cols = st.columns(n_col_kpi)
    valores = [
        ("Colaboradores", total_colaboradores),
        ("Módulos/treinamentos", total_treinos),
        ("Registros concluídos", int(concluidos)),
        ("Registros pendentes", int(pendentes)),
        ("% geral concluído", f"{pct_geral:.1f}%"),
        ("A vencer / vencidos", len(df_a_vencer)),
    ]
    for i, (label, valor) in enumerate(valores):
        cols[i % n_col_kpi].metric(label, valor)


def render_a_vencer():
    st.subheader(f"⏰ Treinamentos a vencer (validade de {VALIDADE_DIAS} dias)")
    if df_a_vencer.empty:
        st.success("Nenhum treinamento vencido ou próximo do vencimento no filtro atual.")
        return
    st.dataframe(
        df_a_vencer[["Matrícula", "Nome", "Treinamento", "Data", "Vencimento", "Dias restantes", "Situação vencimento"]]
        .rename(columns={"Data": "Data conclusão"})
        .sort_values("Dias restantes"),
        use_container_width=True,
        hide_index=True,
    )
    fig = px.bar(
        df_a_vencer.groupby("Treinamento").size().reset_index(name="Qtde"),
        x="Treinamento", y="Qtde", title="Registros a vencer/vencidos por treinamento",
        color_discrete_sequence=["#e67e22"],
    )
    st.plotly_chart(fig, use_container_width=True)


def render_realizados():
    st.subheader("✅ Treinamentos realizados")
    c1, c2 = st.columns([prop_grafico_esq, 1 - prop_grafico_esq])
    realizados_por_modulo = (
        df_modulos_apenas[df_modulos_apenas["Status"] == STATUS_CONCLUIDO]
        .groupby("Treinamento").size().reset_index(name="Concluídos")
        .sort_values("Concluídos", ascending=False)
    )
    with c1:
        fig = px.bar(realizados_por_modulo, x="Treinamento", y="Concluídos", title="Concluídos por treinamento")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        por_status = df_modulos_apenas["Status"].value_counts().reset_index()
        por_status.columns = ["Status", "Qtde"]
        fig2 = px.pie(
            por_status, names="Status", values="Qtde", title="Distribuição geral de status",
            color="Status", color_discrete_map=STATUS_COLORS,
        )
        st.plotly_chart(fig2, use_container_width=True)


def render_percentual_modulo():
    st.subheader("📈 % da equipe capacitada por módulo")
    tabela = resumo_modulo.reset_index()[
        ["Treinamento", STATUS_CONCLUIDO, STATUS_REFORCO, STATUS_PENDENTE, "Sem dado", "Total", "% Concluído"]
    ].sort_values("% Concluído", ascending=False)
    c1, c2 = st.columns([prop_grafico_esq, 1 - prop_grafico_esq])
    with c1:
        fig = px.bar(
            tabela, x="Treinamento", y="% Concluído", title="% concluído por treinamento",
            color="% Concluído", color_continuous_scale="RdYlGn", range_color=[0, 100],
        )
        fig.add_hline(y=100, line_dash="dot", line_color="gray")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        st.dataframe(tabela, use_container_width=True, hide_index=True)


def render_status_colaborador():
    st.subheader("👤 Status de cada colaborador por treinamento")
    matriz = df_filtrado.pivot_table(
        index=["Matrícula", "Nome"], columns="Treinamento", values="Status", aggfunc="first"
    ).reset_index()

    def cor_status(val):
        cor = STATUS_COLORS.get(val, "")
        return f"background-color: {cor}; color: white;" if cor else ""

    cols_status = [c for c in matriz.columns if c not in ("Matrícula", "Nome")]
    st.dataframe(
        matriz.style.applymap(cor_status, subset=cols_status),
        use_container_width=True,
        hide_index=True,
    )
    csv = matriz.to_csv(index=False).encode("utf-8-sig")
    st.download_button("⬇️ Baixar tabela (CSV)", csv, "status_por_colaborador.csv", "text/csv")


RENDERERS = {
    "Indicadores gerais": render_indicadores_gerais,
    "Treinamentos a vencer": render_a_vencer,
    "Treinamentos realizados": render_realizados,
    "% capacitação por módulo": render_percentual_modulo,
    "Status por colaborador": render_status_colaborador,
}

for secao in secoes_ativas:
    RENDERERS[secao]()
    st.divider()

if not secoes_ativas:
    st.info("Nenhuma seção selecionada. Escolha ao menos uma na barra lateral em **Layout do dashboard**.")
