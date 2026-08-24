# -*- coding: utf-8 -*-
"""
Dashboard de Controle de Treinamentos - Equipe de Transporte
Layout Analisável por Abas: Módulo a Módulo (Quem Fez x Quem Não Fez)
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
    "Sem dado": "#95a5a6",
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

    # Linha 10 do Excel (Índice 9): Módulos
    group_header = df_raw.iloc[9].fillna("").astype(str).str.strip()
    group_header_ffill = group_header.replace("", None).ffill().fillna("")

    # Linha 11 do Excel (Índice 10): Subcabeçalho
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
# Carga e Filtros
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
    st.error(f"Erro ao carregar a planilha: {e}")
    st.stop()

if df_long.empty:
    st.title("🚑 Controle de Treinamentos - Equipe de Transporte")
    st.warning("Nenhum colaborador encontrado entre as linhas 12 e 72.")
    st.stop()

filtro_nomes = st.sidebar.multiselect("Filtrar por Colaborador", sorted(df_long["Nome"].unique()))

df = df_long.copy()
if filtro_nomes:
    df = df[df["Nome"].isin(filtro_nomes)]

hoje = pd.Timestamp(datetime.now().date())

# --------------------------------------------------------------------------
# PAINEL PRINCIPAL COM ABAS ANALISÁVEIS
# --------------------------------------------------------------------------
st.title("🚑 Controle de Treinamentos - Equipe de Transporte")
st.caption("Extração estrita das Linhas 12 a 72 da planilha • Cabeçalhos das Linhas 10 e 11")

tab_por_treino, tab_geral, tab_vencimentos, tab_matriz = st.tabs([
    "📌 Análise por Treinamento (Quem Fez x Quem NÃO Fez)",
    "📊 Visão Geral & % Capacitação",
    "⏰ Vencimentos (Validade 1 Ano)",
    "👤 Matriz Geral por Colaborador"
])

# ==========================================================================
# ABA 1: QUEM FEZ X QUEM NÃO FEZ (DETALHADO POR TREINAMENTO)
# ==========================================================================
with tab_por_treino:
    st.subheader("Análise Detalhada Módulo a Módulo")
    lista_treinamentos = sorted(df["Treinamento"].unique().tolist())

    if not lista_treinamentos:
        st.info("Nenhum treinamento encontrado.")
    else:
        treino_selecionado = st.selectbox(
            "👉 Selecione o Treinamento para analisar:",
            lista_treinamentos
        )

        df_mod = df[df["Treinamento"] == treino_selecionado].copy()
        
        # Filtra Quem Fez x Quem Não Fez
        df_fez = df_mod[df_mod["Status"] == STATUS_CONCLUIDO].copy()
        df_nao_fez = df_mod[df_mod["Status"] != STATUS_CONCLUIDO].copy()

        total_equipe = len(df_mod)
        qtd_fez = len(df_fez)
        qtd_nao_fez = len(df_nao_fez)
        pct_fez = (qtd_fez / total_equipe * 100) if total_equipe > 0 else 0

        # Métrica do Módulo
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Total de Colaboradores", total_equipe)
        c2.metric("Concluíram (Quem Fez)", qtd_fez, f"{pct_fez:.1f}% da equipe")
        c3.metric("Pendentes/Faltantes (Quem NÃO Fez)", qtd_nao_fez)
        c4.metric("Status do Módulo", "Crítico" if pct_fez < 70 else ("Atenção" if pct_fez < 90 else "Ideal"))

        st.divider()

        col_left, col_right = st.columns(2)

        # QUEM FEZ
        with col_left:
            st.success(f"🟢 **QUEM FEZ ({qtd_fez} colaboradores)**")
            if df_fez.empty:
                st.warning("Nenhum colaborador concluiu este treinamento ainda.")
            else:
                df_fez["Data Realização"] = df_fez["Data"].apply(lambda d: d.strftime("%d/%m/%Y") if pd.notna(d) else "Data N/A")
                df_fez["Vencimento"] = df_fez["Data"].apply(lambda d: (d + pd.Timedelta(days=VALIDADE_DIAS)).strftime("%d/%m/%Y") if pd.notna(d) else "N/A")
                df_fez["Dias Restantes"] = df_fez["Data"].apply(lambda d: (d + pd.Timedelta(days=VALIDADE_DIAS) - hoje).days if pd.notna(d) else "N/A")

                st.dataframe(
                    df_fez[["Matrícula", "Nome", "Data Realização", "Vencimento", "Dias Restantes"]],
                    use_container_width=True, hide_index=True
                )

        # QUEM NÃO FEZ
        with col_right:
            st.error(f"🔴 **QUEM NÃO FEZ / PENDENTE ({qtd_nao_fez} colaboradores)**")
            if df_nao_fez.empty:
                st.balloons()
                st.success("Toda a equipe concluiu este treinamento!")
            else:
                st.dataframe(
                    df_nao_fez[["Matrícula", "Nome", "Status"]],
                    use_container_width=True, hide_index=True
                )

# ==========================================================================
# ABA 2: VISÃO GERAL & % CAPACITAÇÃO
# ==========================================================================
with tab_geral:
    st.subheader("Evolução Global da Capacitação")

    resumo_modulo = df.groupby("Treinamento")["Status"].value_counts().unstack(fill_value=0)
    for c in [STATUS_CONCLUIDO, STATUS_REFORCO, STATUS_PENDENTE, "Sem dado"]:
        if c not in resumo_modulo.columns:
            resumo_modulo[c] = 0

    resumo_modulo["Total"] = resumo_modulo[[STATUS_CONCLUIDO, STATUS_REFORCO, STATUS_PENDENTE, "Sem dado"]].sum(axis=1)
    resumo_modulo["% Capacitado"] = (resumo_modulo[STATUS_CONCLUIDO] / resumo_modulo["Total"] * 100).round(1)
    tabela_capacitacao = resumo_modulo.reset_index().sort_values("% Capacitado", ascending=False)

    col_g1, col_g2 = st.columns([0.55, 0.45])

    with col_g1:
        fig_cap = px.bar(
            tabela_capacitacao, x="Treinamento", y="% Capacitado",
            title="% da Equipe Capacitada por Módulo",
            color="% Capacitado", color_continuous_scale="RdYlGn", range_color=[0, 100],
            text_auto=".1f"
        )
        st.plotly_chart(fig_cap, use_container_width=True)

    with col_g2:
        st.markdown("**Resumo Numérico por Módulo**")
        st.dataframe(
            tabela_capacitacao[["Treinamento", STATUS_CONCLUIDO, STATUS_PENDENTE, STATUS_REFORCO, "Total", "% Capacitado"]],
            use_container_width=True, hide_index=True
        )

# ==========================================================================
# ABA 3: VENCIMENTOS (1 ANO)
# ==========================================================================
with tab_vencimentos:
    st.subheader("Acompanhamento de Validade dos Treinamentos (365 Dias)")
    
    df_concluidos_com_data = df[(df["Status"] == STATUS_CONCLUIDO) & (df["Data"].notna())].copy()
    df_concluidos_com_data["Vencimento"] = df_concluidos_com_data["Data"] + pd.Timedelta(days=VALIDADE_DIAS)
    df_concluidos_com_data["Dias Restantes"] = (df_concluidos_com_data["Vencimento"] - hoje).dt.days

    df_concluidos_com_data["Situação Vencimento"] = df_concluidos_com_data["Dias Restantes"].apply(
        lambda d: "Vencido" if d < 0 else ("A vencer (em até 60 dias)" if d <= 60 else "Dentro do prazo")
    )

    df_vencidos_alertas = df_concluidos_com_data[
        df_concluidos_com_data["Situação Vencimento"].isin(["Vencido", "A vencer (em até 60 dias)"])
    ].sort_values("Dias Restantes")

    if df_vencidos_alertas.empty:
        st.success("Excelente! Nenhum treinamento realizado está vencido ou próximo do vencimento (60 dias).")
    else:
        df_exib_venc = df_vencidos_alertas[
            ["Matrícula", "Nome", "Treinamento", "Data", "Vencimento", "Dias Restantes", "Situação Vencimento"]
        ].copy()
        df_exib_venc["Data"] = df_exib_venc["Data"].dt.strftime("%d/%m/%Y")
        df_exib_venc["Vencimento"] = df_exib_venc["Vencimento"].dt.strftime("%d/%m/%Y")

        st.warning(f"Existem **{len(df_vencidos_alertas)} registros** que necessitam de atenção para reciclagem.")
        st.dataframe(df_exib_venc, use_container_width=True, hide_index=True)

# ==========================================================================
# ABA 4: MATRIZ GERAL POR COLABORADOR
# ==========================================================================
with tab_matriz:
    st.subheader("Matriz de Status de Todos os Colaboradores")
    
    busca = st.text_input("🔎 Pesquisar colaborador por nome:")
    
    matriz_status = df.pivot_table(
        index=["Matrícula", "Nome"], columns="Treinamento", values="Status", aggfunc="first"
    ).reset_index()

    if busca:
        matriz_status = matriz_status[matriz_status["Nome"].str.contains(busca, case=False, na=False)]

    def aplicar_cor_status(val):
        cor = STATUS_COLORS.get(val, "")
        return f"background-color: {cor}; color: white; font-weight: bold;" if cor else ""

    cols_modulos = [c for c in matriz_status.columns if c not in ("Matrícula", "Nome")]
    st.dataframe(
        matriz_status.style.map(aplicar_cor_status, subset=cols_modulos),
        use_container_width=True, hide_index=True
    )
