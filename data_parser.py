# -*- coding: utf-8 -*-
"""
data_parser.py
--------------------------------------------------------------
Funções de leitura e parsing da "Planilha Geral de Controle" de
treinamentos da equipe de transporte.

Layout esperado da planilha (linhas do Excel, 1-indexadas):
  Linha 1 (grupo) : nome do módulo/treinamento, mesclado sobre o par
                    de colunas Data/Status (ex.: "M1", "M1", "M2", "M2", ...)
                    e, ao final, o próprio rótulo das colunas de status
                    único (ex.: "Certificação").
  Linha 2 (sub)   : "Data" / "Status" para cada módulo; vazio para as
                    colunas de status único.
  Linha 3 em diante: dados dos colaboradores. As duas primeiras colunas
                    são Matrícula e Nome.

Este módulo é tolerante a pequenas variações (linha de grupo mesclada
representada apenas na primeira célula do par, acentuação, maiúsculas/
minúsculas etc.).
"""

from __future__ import annotations

import io
import re
import unicodedata

import pandas as pd

STATUS_CONCLUIDO = "Concluído"
STATUS_REFORCO = "Necessita de reforço"
STATUS_PENDENTE = "Pendente"
STATUS_SEM_DADO = "Sem dado"


# --------------------------------------------------------------------------
# Utilidades
# --------------------------------------------------------------------------
def _normalizar(texto) -> str:
    """Remove acentos, baixa a caixa e tira espaços extras para comparação."""
    if texto is None:
        return ""
    texto = str(texto).strip()
    if not texto or texto.lower() == "nan":
        return ""
    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(c for c in texto if not unicodedata.combining(c))
    return texto.lower().strip()


def limpar_status(valor) -> str:
    """Normaliza o texto de status de uma célula para um dos status padrão."""
    v = _normalizar(valor)
    if not v:
        return STATUS_SEM_DADO

    if v in {"concluido", "concluida", "ok", "feito", "realizado", "sim", "c"}:
        return STATUS_CONCLUIDO
    if "reforc" in v:
        return STATUS_REFORCO
    if v in {"pendente", "nao realizado", "n realizado", "nao", "pendencia", "p"}:
        return STATUS_PENDENTE
    if "conclu" in v:
        return STATUS_CONCLUIDO

    # não reconhecido -> mantém como pendente-ish "Sem dado" para não mascarar dado estranho
    return STATUS_SEM_DADO


_RE_ISO = re.compile(r"^\d{4}-\d{2}-\d{2}")


def _parse_data(valor):
    """Converte célula de data para Timestamp (ou NaT se vazia/ inválida).

    Aceita tanto datas já convertidas pelo pandas/openpyxl (objetos
    Timestamp/datetime) quanto texto em formato ISO (AAAA-MM-DD) ou
    brasileiro (DD/MM/AAAA)."""
    if valor is None or (isinstance(valor, float) and pd.isna(valor)):
        return pd.NaT
    if isinstance(valor, pd.Timestamp):
        return valor

    texto = str(valor).strip()
    if not texto or texto.lower() == "nan":
        return pd.NaT

    if _RE_ISO.match(texto):
        # Já está em AAAA-MM-DD: não há ambiguidade dia/mês.
        return pd.to_datetime(texto, format="%Y-%m-%d", errors="coerce")

    # Formatos como DD/MM/AAAA (padrão brasileiro).
    return pd.to_datetime(texto, dayfirst=True, errors="coerce")


# --------------------------------------------------------------------------
# Leitura do arquivo
# --------------------------------------------------------------------------
def ler_planilha(file_bytes: bytes, nome_aba: str | None, linha_max: int):
    """
    Lê a planilha (bytes de um .xlsx) e devolve:
      (df_bruto, nome_da_aba_usada, lista_de_abas)

    df_bruto é lido sem cabeçalho (header=None), preservando as duas
    primeiras linhas como linhas de estrutura (grupo/sub), limitado a
    `linha_max` linhas do Excel.
    """
    xls = pd.ExcelFile(io.BytesIO(file_bytes))
    abas_disponiveis = xls.sheet_names

    if not abas_disponiveis:
        raise ValueError("A planilha não contém nenhuma aba.")

    if nome_aba not in abas_disponiveis:
        nome_aba = abas_disponiveis[0]

    df_bruto = pd.read_excel(
        xls, sheet_name=nome_aba, header=None, nrows=int(linha_max)
    )

    return df_bruto, nome_aba, abas_disponiveis


# --------------------------------------------------------------------------
# Identificação da estrutura (linhas de cabeçalho x dados)
# --------------------------------------------------------------------------
def parse_estrutura(df_bruto: pd.DataFrame):
    """
    Separa o dataframe bruto em:
      grupo     -> Series com o rótulo do módulo/coluna para cada coluna
                   (linha 1 do Excel, com merge simulado via ffill)
      sub       -> Series com "Data"/"Status" (linha 2 do Excel)
      dados     -> DataFrame só com as linhas de dados (a partir da linha 3)
      dados_idx -> índice original (linha do Excel) de cada linha de dados,
                   útil para mensagens de erro/depuração
    """
    if len(df_bruto) < 3:
        raise ValueError(
            "A planilha precisa ter ao menos 3 linhas: grupo, sub-cabeçalho e dados."
        )

    linha_grupo = df_bruto.iloc[0].copy()
    linha_sub = df_bruto.iloc[1].copy()

    # Simula célula mesclada: propaga o rótulo do grupo para a direita
    # quando a célula seguinte está vazia (comum em export de merge do Excel/Sheets).
    linha_grupo = linha_grupo.ffill()

    dados = df_bruto.iloc[2:].reset_index(drop=True)
    dados_idx = df_bruto.index[2:] + 1  # +1 porque df_bruto é 0-indexado e representa linha 1 do Excel

    return linha_grupo, linha_sub, dados, dados_idx


# --------------------------------------------------------------------------
# Identificação das colunas (Matrícula, Nome, módulos em par Data/Status,
# colunas de status único)
# --------------------------------------------------------------------------
def identificar_colunas(grupo: pd.Series, sub: pd.Series):
    """
    Varre as linhas de estrutura e identifica:
      col_matricula -> índice da coluna de Matrícula
      col_nome      -> índice da coluna de Nome
      modulos       -> lista de dicts {"label", "col_data", "col_status"}
      unicos        -> lista de dicts {"label", "col_status"}
    """
    n_cols = len(grupo)
    col_matricula = None
    col_nome = None
    modulos = []
    unicos = []

    i = 0
    while i < n_cols:
        rotulo_grupo = _normalizar(grupo.iloc[i])
        rotulo_sub = _normalizar(sub.iloc[i])

        if col_matricula is None and ("matricul" in rotulo_grupo or "matricul" in rotulo_sub):
            col_matricula = i
            i += 1
            continue

        if col_nome is None and (rotulo_grupo == "nome" or rotulo_sub == "nome"):
            col_nome = i
            i += 1
            continue

        if rotulo_sub == "data":
            # procura a coluna de status correspondente logo em seguida
            label = str(grupo.iloc[i]).strip() if pd.notna(grupo.iloc[i]) else f"Coluna {i + 1}"
            if i + 1 < n_cols and _normalizar(sub.iloc[i + 1]) == "status":
                modulos.append({"label": label, "col_data": i, "col_status": i + 1})
                i += 2
                continue
            else:
                # coluna de data isolada, sem status ao lado -> trata como único (status ausente)
                modulos.append({"label": label, "col_data": i, "col_status": None})
                i += 1
                continue

        if rotulo_sub == "status" and not rotulo_grupo:
            # status "solto" sem par de data explícito, associado ao grupo anterior
            i += 1
            continue

        # coluna sem "Data"/"Status" no sub-cabeçalho: status único (ex.: Certificação, Situação)
        if rotulo_grupo:
            label = str(grupo.iloc[i]).strip()
            unicos.append({"label": label, "col_status": i})

        i += 1

    _deduplicar_labels(modulos, unicos)

    return col_matricula, col_nome, modulos, unicos


def _deduplicar_labels(modulos, unicos):
    """Garante rótulos únicos entre módulos e status únicos.

    A planilha original pode ter blocos com o mesmo rótulo (ex.: dois
    módulos chamados "M2"). Sem isso, eles se misturariam como se fossem
    o mesmo treinamento e quebrariam widgets do Streamlit que usam o
    rótulo como key. Ocorrências repetidas viram "M2", "M2 (2)", "M2 (3)"...
    """
    contagem = {}
    for item in modulos + unicos:
        original = item["label"]
        contagem[original] = contagem.get(original, 0) + 1
        if contagem[original] > 1:
            item["label"] = f"{original} ({contagem[original]})"


# --------------------------------------------------------------------------
# Montagem do dataframe "longo" (uma linha por colaborador x treinamento)
# --------------------------------------------------------------------------
def montar_dataframe_longo(dados: pd.DataFrame, col_matricula, col_nome, modulos, unicos):
    """
    Constrói um DataFrame longo com colunas:
      Matrícula, Nome, Treinamento, Tipo ("Módulo"/"Único"), Data, Status
    """
    linhas = []

    for _, linha in dados.iterrows():
        matricula = linha.iloc[col_matricula] if col_matricula is not None else None
        nome = linha.iloc[col_nome] if col_nome is not None else None

        if pd.isna(nome) and pd.isna(matricula):
            continue

        for m in modulos:
            data_val = _parse_data(linha.iloc[m["col_data"]]) if m["col_data"] is not None else pd.NaT
            status_val = (
                limpar_status(linha.iloc[m["col_status"]])
                if m["col_status"] is not None
                else limpar_status(None)
            )
            linhas.append(
                {
                    "Matrícula": matricula,
                    "Nome": nome,
                    "Treinamento": m["label"],
                    "Tipo": "Módulo",
                    "Data": data_val,
                    "Status": status_val,
                }
            )

        for u in unicos:
            status_val = limpar_status(linha.iloc[u["col_status"]])
            linhas.append(
                {
                    "Matrícula": matricula,
                    "Nome": nome,
                    "Treinamento": u["label"],
                    "Tipo": "Único",
                    "Data": pd.NaT,
                    "Status": status_val,
                }
            )

    df_long = pd.DataFrame(
        linhas, columns=["Matrícula", "Nome", "Treinamento", "Tipo", "Data", "Status"]
    )
    return df_long
