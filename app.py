# -*- coding: utf-8 -*-
"""
Dashboard de Controle de Treinamentos - Equipe de Transporte
--------------------------------------------------------------
Lê a planilha exatamente das linhas 10/11 (cabeçalhos) e 12 a 72 (dados).
"""

import io
import json
from datetime import datetime
from pathlib import Path

import pandas as pd
import plotly.express as px
import requests
import streamlit as st

# --------------------------------------------------------------------------
# Configuração geral da página
# --------------------------------------------------------------------------
st.set_page_config(
    page_title="Controle de Treinamentos - Transporte",
    page_icon="🚑",
    layout="wide",
)

CONFIG_PATH = Path("data/layout_config.json")
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

SECOES_PADRAO = [
    "Indicadores gerais",
    "Treinamentos a vencer",
    "Treinamentos realizados",
    "% capacitação por módulo",
    "Status por colaborador",
]

GOOGLE_SHEET_ID = "1HQ9GRicZfVP_rUR51AxNpDaz_5GcZgAcBiP9qGkKZDk"
GOOGLE_SHEET_GID = "1064660493"
GOOGLE_SHEET_EXPORT_URL = (
    f"https://docs.google.com/spreadsheets/d/{GOOGLE_SHEET_ID}/export"
    f"?format=xlsx&gid={GOOGLE_SHEET_GID}"
)


@st.cache_data(show_spinner="Buscando planilha atualizada...", ttl=300)
def baixar_planilha_google(url: str) -> bytes:
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/115.0.0.0 Safari/537.36"
    }
    resposta = requests.get(url, timeout=30, headers=headers)
    resposta.raise_for_status()
    if resposta.content.startswith(b"<!DOCTYPE") or resposta.content.startswith(b"<html"):
        raise ValueError("O link retornou HTML. Verifique o compartilhamento da planilha.")
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
    # Se tiver data válida de conclusão e o status estiver em branco, assume Concluído
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


@st.cache_data(show_spinner="Lendo linhas da planilha...")
def processar_planilha_fixa(
    file_bytes: bytes,
    nome_aba_req: str | None,
    l_grupo_excel: int,
    l_sub_excel: int,
    l_ini_excel: int,
    l_fim_excel: int,
):
    excel = pd.ExcelFile(io.BytesIO(file_bytes))
    abas = excel.sheet_names
    aba_usada = nome_aba_req if (nome_aba_req and nome_aba_req in abas) else abas[0]

    # Lê até a linha limite necessária
    df_raw = pd.read_excel(excel, sheet_name=aba_usada, header=None, nrows=l_fim_excel)

    # Conversão para índices 0-based do Python
    idx_grupo = max(0, l_grupo_excel - 1)
    idx_sub = max(0, l_sub_excel - 1)
    idx_ini = max(0, l_ini_excel - 1)
    idx_fim = l_fim_excel

    # Linha 10 (Módulos/Grupos superior)
    group_header = df_raw.iloc[idx_grupo].fillna("").astype(str).str.strip()
    group_header_ffill = group_header.replace("", None).ffill().fillna("")

    # Linha 11 (Subcabeçalho Data/Status/Nome/Matrícula)
    sub_header = df_raw.iloc[idx_sub].fillna("").astype(str).str.strip()

    # Identifica colunas Matrícula e Nome
    col_matricula = None
    col_nome = None

    for idx, val in sub_header.items():
        v_low = str(val).lower()
        if col_matricula is None and any(k in v_low for k in ["matr", "re", "código", "codigo", "id"]):
            col_matricula = idx
        if col_nome is None and any(k in v_low for k in ["nome", "colaborador", "funciona", "funcioná"]):
            col_nome = idx

    if col_nome is None:
        for idx, val in group_header.items():
            v_low = str(val).lower()
            if any(k in v_low for k in ["nome", "colaborador"]):
                col_nome = idx
                break

    if col_nome is None:
        col_nome = 1 if len(sub_header) > 1 else 0
    if col_matricula is None and col_nome != 0:
        col_matricula = 0

    # Pareamento de Módulos (Data + Status)
    modulos = []
    unicos = []
    cols_usadas = set()
    if col_matricula is not None:
        cols_usadas.add(col_matricula)
    if col_nome is not None:
        cols_usadas.add(col_nome)

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

        mesmo_grupo = (gh_curr != "" and gh_curr == gh_next)

        sh_curr_low = sh_curr.lower()
        sh_next_low = sh_next.lower()

        is_data_1 = any(k in sh_curr_low for k in ["data", "dt", "realiza"])
        is_status_1 = any(k in sh_curr_low for k in ["status", "situa", "conceito", "resultado"])
        is_data_2 = any(k in sh_next_low for k in ["data", "dt", "realiza"])
        is_status_2 = any(k in sh_next_low for k in ["status", "situa", "conceito", "resultado"])

        tem_pareamento = (is_data_1 and is_status_2) or (is_status_1 and is_data_2)

        if (mesmo_grupo or tem_pareamento) and (i + 1 < n_cols) and (i + 1 not in cols_usadas):
            if is_status_1 or "status" in sh_curr_low:
                c_st, c_dt = i, i + 1
            else:
                c_dt, c_st = i, i + 1

            label = gh_curr if gh_curr else (sh_curr if sh_curr not in ["Data", "Status", ""] else f"Módulo {len(modulos)+1}")
            modulos.append({"label": label, "col_data": c_dt, "col_status": c_st})
            cols_usadas.add(i)
            cols_usadas.add(i + 1)
            i += 2
            continue

        label = gh_curr if gh_curr else sh_curr
        if label and label.lower() not in ["nan", "none", "unnamed", ""]:
            unicos.append({"label": label, "col_status": i})
            cols_usadas.add(i)

        i += 1

    # Leitura estrita da linha 12 à 72 do Excel
    df_dados = df_raw.iloc[idx_ini:idx_fim].reset_index(drop=True)
    registros = []

    for _, row in df_dados.iterrows():
        nome_raw = row[col_nome] if col_nome is not None and col_nome in row else None
        matr_raw = row[col_matricula] if col_matricula is not None and col_matricula in row else None

        nome = str(nome_raw).strip() if pd.notna(nome_raw) else ""
        matr = str(matr_raw).strip() if pd.notna(matr_raw) else ""

        if not nome or nome.lower() in ["nan", "none", "0", "total", "subtotal", "nome"]:
            continue

        for m in modulos:
            dt_raw = row[m["col_data"]] if m["col_data"] in row else None
            st_raw = row[m["col_status"]] if m["col_status"] in row else None

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
                "Tipo": "Módulo",
            })

        for u in unicos:
            st_raw = row[u["col_status"]] if u["col_status"] in row else None
            registros.append({
                "Matrícula": matr,
                "Nome": nome,
                "Treinamento": u["label"],
                "Data": None,
                "Status": limpar_status(st_raw, tem_data_valida=False),
                "Tipo": "Único",
            })

    df_long = pd.DataFrame(registros)
    return df_long, aba_usada, abas


# --------------------------------------------------------------------------
# Sessão / Configuração
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
# Sidebar & Processamento
# --------------------------------------------------------------------------
st.sidebar.title("⚙️ Mapeamento de Linhas")

l_grupo = st.sidebar.number_input("Linha do Módulo (Excel)", min_value=1, value=10)
l_sub = st.sidebar.number_input("Linha do Subcabeçalho (Excel)", min_value=1, value=11)
l_ini = st.sidebar.number_input("Linha Inicial dos Dados (Excel)", min_value=1, value=12)
l_fim = st.sidebar.number_input("Linha Final dos Dados (Excel)", min_value=1, value=72)

if st.sidebar.button("🔄 Recarregar Planilha", use_container_width=True):
    baixar_planilha_google.clear()
    processar_planilha_fixa.clear()
    st.rerun()

try:
    file_bytes = baixar_planilha_google(GOOGLE_SHEET_EXPORT_URL)
    df_long, aba_usada, abas_disponiveis = processar_planilha_fixa(
        file_bytes, st.session_state.config.get("aba"), l_grupo, l_sub, l_ini, l_fim
    )
except Exception as e:
    st.title("🚑 Controle de Treinamentos - Equipe de Transporte")
    st.error(f"Erro ao carregar a planilha: {e}")
    st.stop()

aba_escolhida = st.sidebar.selectbox(
    "Aba da planilha", abas_disponiveis, index=abas_disponiveis.index(aba_usada) if aba_usada in abas_disponiveis else 0
)
if aba_escolhida != aba_usada:
    st.session_state.config["aba"] = aba_escolhida
    df_long, aba_usada, abas_disponiveis = processar_planilha_fixa(
        file_bytes, aba_escolhida, l_grupo, l_sub, l_ini, l_fim
    )

if df_long.empty:
    st.title("🚑 Controle de Treinamentos - Equipe de Transporte")
    st.warning(f"Nenhum registro encontrado entre as linhas {l_ini} e {l_fim} da aba '{aba_usada}'. Verifique a seleção das linhas na barra lateral.")
    st.stop()

# --------------------------------------------------------------------------
# Edição de Nomes e Filtros
# --------------------------------------------------------------------------
st.sidebar.subheader("✏️ Nomes exibidos dos treinamentos")
nomes_editados = st.session_state.config.get("nomes_editados", {})
todos_treinos = sorted(df_long["Treinamento"].unique().tolist())

with st.sidebar.expander("Editar rótulos", expanded=False):
    for idx, label in enumerate(todos_treinos):
        novo = st.text_input(label, value=nomes_editados.get(label, label), key=f"lbl_{idx}")
        nomes_editados[label] = novo

df_long["Treinamento"] = df_long["Treinamento"].map(lambda x: nomes_editados.get(x, x))

col1, col2 = st.sidebar.columns(2)
if col1.button("💾 Salvar", use_container_width=True):
    st.session_state.config["nomes_editados"] = nomes_editados
    salvar_config(st.session_state.config)
    st.sidebar.success("Salvo!")
if col2.button("↩️ Restaurar", use_container_width=True):
    st.session_state.config["nomes_editados"] = {}
    salvar_config(st.session_state.config)
    st.rerun()

st.sidebar.subheader("🔎 Filtros")
filtro_nomes = st.sidebar.multiselect("Colaborador(es)", sorted(df_long["Nome"].unique()))
filtro_treinamentos = st.sidebar.multiselect("Treinamento(s)", sorted(df_long["Treinamento"].unique()))
filtro_status = st.sidebar.multiselect("Status", sorted(df_long["Status"].unique()))
dias_alerta = st.sidebar.slider("Alertar vencimento em até (dias)", 15, 180, 60, 15)

df_filtrado = df_long.copy()
if filtro_nomes:
    df_filtrado = df_filtrado[df_filtrado["Nome"].isin(filtro_nomes)]
if filtro_treinamentos:
    df_filtrado = df_filtrado[df_filtrado["Treinamento"].isin(filtro_treinamentos)]
if filtro_status:
    df_filtrado = df_filtrado[df_filtrado["Status"].isin(filtro_status)]

st.sidebar.subheader("🧩 Layout")
secoes_ativas = st.sidebar.multiselect("Seções exibidas", SECOES_PADRAO, default=SECOES_PADRAO)
n_col_kpi = st.sidebar.slider("Nº de colunas nos indicadores", 2, 5, 4)
prop_grafico_esq = st.sidebar.slider("Proporção coluna esquerda", 0.2, 0.8, 0.5, 0.05)

# --------------------------------------------------------------------------
# Cálculos
# --------------------------------------------------------------------------
df_modulos_apenas = df_filtrado[df_filtrado["Tipo"] == "Módulo"].copy()
df_venc = df_modulos_apenas[(df_modulos_apenas["Status"] == STATUS_CONCLUIDO) & (df_modulos_apenas["Data"].notna())].copy()

hoje = pd.Timestamp(datetime.now().date())
df_venc["Vencimento"] = df_venc["Data"] + pd.Timedelta(days=VALIDADE_DIAS)
df_venc["Dias restantes"] = (df_venc["Vencimento"] - hoje).dt.days
df_venc["Situação vencimento"] = df_venc["Dias restantes"].apply(
    lambda d: "Vencido" if d < 0 else ("A vencer" if d <= dias_alerta else "Dentro do prazo")
)
df_a_vencer = df_venc[df_venc["Situação vencimento"].isin(["Vencido", "A vencer"])].sort_values("Dias restantes")

total_colaboradores = df_long["Nome"].nunique()

resumo_modulo = df_filtrado.groupby("Treinamento")["Status"].value_counts().unstack(fill_value=0)
for c in [STATUS_CONCLUIDO, STATUS_REFORCO, STATUS_PENDENTE, "Sem dado"]:
    if c not in resumo_modulo.columns:
        resumo_modulo[c] = 0
resumo_modulo["Total"] = resumo_modulo[[STATUS_CONCLUIDO, STATUS_REFORCO, STATUS_PENDENTE, "Sem dado"]].sum(axis=1)
resumo_modulo["% Concluído"] = (resumo_modulo[STATUS_CONCLUIDO] / resumo_modulo["Total"] * 100).round(1)

# --------------------------------------------------------------------------
# Painel Principal
# --------------------------------------------------------------------------
st.title("🚑 Controle de Treinamentos - Equipe de Transporte")
st.caption(f"Aba: **{aba_usada}** • Linhas {l_ini} a {l_fim} • **{total_colaboradores} colaboradores carregados**")


def render_indicadores_gerais():
    st.subheader("📊 Indicadores gerais")
    total_treinos = df_modulos_apenas["Treinamento"].nunique()
    concluidos = (df_modulos_apenas["Status"] == STATUS_CONCLUIDO).sum()
    pendentes = (df_modulos_apenas["Status"] == STATUS_PENDENTE).sum()
    pct_geral = (concluidos / len(df_modulos_apenas) * 100) if len(df_modulos_apenas) else 0

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
        df_a_vencer[["Matrícula", "Nome", "Treinamento", "Data", "Vencimento", "Dias restantes", "Situação vencimento"]],
        use_container_width=True, hide_index=True,
    )
    fig = px.bar(
        df_a_vencer.groupby("Treinamento").size().reset_index(name="Qtde"),
        x="Treinamento", y="Qtde", title="A vencer/vencidos por treinamento",
        color_discrete_sequence=["#e67e22"],
    )
    st.plotly_chart(fig, use_container_width=True)


def render_realizados():
    st.subheader("✅ Treinamentos realizados")
    c1, c2 = st.columns([prop_grafico_esq, 1 - prop_grafico_esq])
    realizados = (
        df_modulos_apenas[df_modulos_apenas["Status"] == STATUS_CONCLUIDO]
        .groupby("Treinamento").size().reset_index(name="Concluídos")
        .sort_values("Concluídos", ascending=False)
    )
    with c1:
        fig = px.bar(realizados, x="Treinamento", y="Concluídos", title="Concluídos por treinamento")
        st.plotly_chart(fig, use_container_width=True)
    with c2:
        por_status = df_modulos_apenas["Status"].value_counts().reset_index()
        por_status.columns = ["Status", "Qtde"]
        fig2 = px.pie(por_status, names="Status", values="Qtde", title="Distribuição geral", color="Status", color_discrete_map=STATUS_COLORS)
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
            color="% Concluído", color_continuous_scale="RdYlGn", range_color=[0, 100]
        )
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
    st.dataframe(matriz.style.map(cor_status, subset=cols_status), use_container_width=True, hide_index=True)


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
