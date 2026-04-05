"""
Conferência SIPAC x SIAFI - Ferramenta de Reconciliação Orçamentária
UFPB - Universidade Federal da Paraíba
"""

import streamlit as st
import pandas as pd
from io import BytesIO
from difflib import SequenceMatcher
import re

# ============================================================
# CONFIGURAÇÃO
# ============================================================
st.set_page_config(
    page_title="Conferência SIPAC x SIAFI",
    page_icon="🔍",
    layout="wide",
    initial_sidebar_state="expanded",
)

SIAFI_COLUMNS = [
    "natureza", "natureza_desc", "ptres", "fonte", "fonte_desc",
    "esfera", "esfera_desc", "pi", "pi_desc", "ug", "ug_desc",
    "ugr", "ugr_desc", "credito", "empenhado",
]

# ============================================================
# TABELA DE MAPEAMENTO UG + UGR → UNIDADE SIPAC (fixa)
# ============================================================
MAPPING_DATA = [
    ("153065", "150646", "110046"),
    ("153065", "150647", "110047"),
    ("153065", "150648", "110048"),
    ("153065", "150649", "110049"),
    ("153065", "150650", "110050"),
    ("153065", "150651", "110051"),
    ("153065", "150652", "110052"),
    ("153065", "150653", "110053"),
    ("153065", "150654", "110054"),
    ("153065", "150655", "110055"),
    ("153065", "150656", "110056"),
    ("153065", "150657", "110057"),
    ("153065", "150658", "110058"),
    ("153065", "150659", "110059"),
    ("153065", "150660", "110060"),
    ("153065", "150899", "11004638"),
    ("153065", "150900", "110040"),
    ("153065", "150901", "110041"),
    ("153065", "150902", "110042"),
    ("153065", "150903", "HULW"),
    ("153065", "150904", "110144"),
    ("153065", "150905", "110044"),
    ("153065", "150906", "110045"),
    ("153065", "151616", "110061"),
    ("153065", "151831", "110062"),
    ("153065", "152306", "110063"),
    ("153065", "152643", "110064"),
    ("153065", "152644", "110065"),
    ("153065", "152645", "110066"),
    ("153065", "155504", "110070"),
    ("153065", "156042", "110079"),
    ("153065", "157046", "11000012"),
    ("153066", "150646", "11004638"),
    ("153066", "150647", "11004638"),
    ("153066", "150648", "11004638"),
    ("153066", "150649", "11004638"),
    ("153066", "150650", "11004638"),
    ("153066", "150651", "11004638"),
    ("153066", "150652", "11004638"),
    ("153066", "150653", "11004638"),
    ("153066", "150654", "11004638"),
    ("153066", "150655", "11004638"),
    ("153066", "150656", "11004638"),
    ("153066", "150657", "11004638"),
    ("153066", "150658", "11004638"),
    ("153066", "150659", "11004638"),
    ("153066", "150660", "11004638"),
    ("153066", "150899", "11004638"),
    ("153066", "150900", "11004638"),
    ("153066", "150901", "11004638"),
    ("153066", "150902", "11004638"),
    ("153066", "150903", "11004638"),
    ("153066", "150904", "11004638"),
    ("153066", "150905", "11004638"),
    ("153066", "150906", "11004638"),
    ("153066", "151616", "11004638"),
    ("153066", "151831", "11004638"),
    ("153066", "152306", "11004638"),
    ("153066", "152643", "11004638"),
    ("153066", "152644", "11004638"),
    ("153066", "152645", "11004638"),
    ("153066", "155504", "11004638"),
    ("153066", "156042", "11004638"),
    ("153066", "157046", "11004638"),
    ("153073", "150646", "110044"),
    ("153073", "150647", "110044"),
    ("153073", "150648", "110044"),
    ("153073", "150649", "110044"),
    ("153073", "150650", "110044"),
    ("153073", "150651", "110044"),
    ("153073", "150652", "110044"),
    ("153073", "150653", "110044"),
    ("153073", "150654", "110044"),
    ("153073", "150655", "110044"),
    ("153073", "150656", "110044"),
    ("153073", "150657", "110044"),
    ("153073", "150658", "110044"),
    ("153073", "150659", "110044"),
    ("153073", "150660", "110044"),
    ("153073", "150899", "110044"),
    ("153073", "150900", "110044"),
    ("153073", "150901", "110044"),
    ("153073", "150902", "110044"),
    ("153073", "150903", "110044"),
    ("153073", "150904", "110044"),
    ("153073", "150905", "110044"),
    ("153073", "150906", "110044"),
    ("153073", "151616", "110044"),
    ("153073", "151831", "110044"),
    ("153073", "152306", "110044"),
    ("153073", "152643", "110044"),
    ("153073", "152644", "110044"),
    ("153073", "152645", "110044"),
    ("153073", "155504", "110044"),
    ("153073", "156042", "110044"),
    ("153073", "157046", "110044"),
    ("153074", "150646", "110045"),
    ("153074", "150647", "110045"),
    ("153074", "150648", "110045"),
    ("153074", "150649", "110045"),
    ("153074", "150650", "110045"),
    ("153074", "150651", "110045"),
    ("153074", "150652", "110045"),
    ("153074", "150653", "110045"),
    ("153074", "150654", "110045"),
    ("153074", "150655", "110045"),
    ("153074", "150656", "110045"),
    ("153074", "150657", "110045"),
    ("153074", "150658", "110045"),
    ("153074", "150659", "110045"),
    ("153074", "150660", "110045"),
    ("153074", "150899", "110045"),
    ("153074", "150900", "110045"),
    ("153074", "150901", "110045"),
    ("153074", "150902", "110045"),
    ("153074", "150903", "110045"),
    ("153074", "150904", "110045"),
    ("153074", "150905", "110045"),
    ("153074", "150906", "110045"),
    ("153074", "151616", "110045"),
    ("153074", "151831", "110045"),
    ("153074", "152306", "110045"),
    ("153074", "152643", "110045"),
    ("153074", "152644", "110045"),
    ("153074", "152645", "110045"),
    ("153074", "155504", "110045"),
    ("153074", "156042", "110045"),
    ("153074", "157046", "110045"),
]

# Dicionário de mapeamento para busca rápida
MAPPING_DICT = {(ug, ugr): unidade for ug, ugr, unidade in MAPPING_DATA}


# ============================================================
# FUNÇÕES DE PARSING
# ============================================================

def parse_siafi(file):
    """Parse arquivo SIAFI (estrutura conferência sipac)."""
    df = pd.read_excel(file, header=None)

    # Detectar primeira linha de dados (valor numérico na col 13 = crédito)
    data_start = 0
    for i in range(len(df)):
        val = df.iloc[i, 13]
        if isinstance(val, (int, float)) and not pd.isna(val):
            data_start = i
            break

    df = df.iloc[data_start:].reset_index(drop=True)

    if df.shape[1] < 15:
        st.error("Arquivo SIAFI não possui 15 colunas esperadas.")
        return None

    df = df.iloc[:, :15]
    df.columns = SIAFI_COLUMNS

    # Forward-fill TODAS as colunas hierárquicas (arquivo tem células mescladas)
    for col in ["natureza", "natureza_desc", "ptres", "fonte", "fonte_desc",
                 "esfera", "esfera_desc", "pi", "pi_desc", "ug", "ug_desc"]:
        df[col] = df[col].ffill()

    # Converter valores monetários
    df["credito"] = pd.to_numeric(df["credito"], errors="coerce").fillna(0)
    df["empenhado"] = pd.to_numeric(df["empenhado"], errors="coerce").fillna(0)

    # Crédito SIAFI = crédito + empenhado
    df["credito_total_siafi"] = df["credito"] + df["empenhado"]

    # Limpar códigos (remover aspas, espaços)
    for col in ["natureza", "ptres", "fonte", "esfera", "pi", "ug", "ugr"]:
        df[col] = df[col].astype(str).str.strip().str.replace("'", "", regex=False)

    # Remover linhas sem fonte válida
    df = df[df["fonte"].notna() & ~df["fonte"].isin(["nan", ""])].reset_index(drop=True)

    # Excluir naturezas terminadas em "00" (ex: 339000, 319000, 449000)
    df = df[~df["natureza"].str.endswith("00")].reset_index(drop=True)

    return df


def parse_sipac(file):
    """Parse arquivo SIPAC (saldo orçamentário). Filtra apenas ano 2026.
    Retorna (DataFrame, dict_nomes_unidades).
    """
    df = pd.read_excel(file)

    # Mapeamento flexível de colunas
    col_map = {}
    for col in df.columns:
        cl = col.lower().strip()
        if "nat" in cl and "desp" in cl:
            col_map[col] = "nat_despesa"
        elif cl == "ptres":
            col_map[col] = "ptres"
        elif "fonte" in cl and "recurso" in cl:
            col_map[col] = "fonte_recurso"
        elif cl == "esfera":
            col_map[col] = "esfera"
        elif cl == "pi":
            col_map[col] = "pi"
        elif cl == "ano":
            col_map[col] = "ano"
        elif "distribu" in cl:
            col_map[col] = "distribuido"
        elif "recebido" in cl:
            col_map[col] = "recebido"
        elif "entrada" in cl and "remanej" in cl:
            col_map[col] = "entrada_remanej"
        elif ("sa" in cl or "saída" in cl or "saida" in cl) and "remanej" in cl:
            col_map[col] = "saida_remanej"
        elif "anulado" in cl:
            col_map[col] = "anulado"
        elif "transferido" in cl:
            col_map[col] = "transferido"
        elif "contido" in cl:
            col_map[col] = "contido"
        elif "empenho" in cl:
            col_map[col] = "empenhos"
        elif "saldo" in cl:
            col_map[col] = "saldo"
        elif cl == "unidade":
            col_map[col] = "unidade_nome"
        elif "unidade" in cl and "gestora" not in cl and "gest" not in cl and "siorg" not in cl and cl != "unidade":
            col_map[col] = "cod_unidade"
        elif "gestora" in cl and "siafi" in cl:
            col_map[col] = "ug_siafi"
        elif "gestao" in cl or "gestão" in cl:
            col_map[col] = "cod_gestao"
        elif "siorg" in cl:
            col_map[col] = "cod_siorg"
        else:
            col_map[col] = col

    df = df.rename(columns=col_map)

    # Filtrar apenas ano 2026 (aceita formatos: 2026, 2.026, 2026.0)
    if "ano" in df.columns:
        df["ano"] = df["ano"].astype(str).str.replace(".", "", regex=False).str.strip()
        df["ano"] = pd.to_numeric(df["ano"], errors="coerce")
        df = df[df["ano"] == 2026].copy()

    # Converter códigos para string
    for col in ["nat_despesa", "ptres", "fonte_recurso", "esfera", "pi", "cod_unidade"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()

    # Crédito SIPAC líquido = (positivas) - (negativas)
    pos_cols = [c for c in ["distribuido", "recebido", "entrada_remanej"] if c in df.columns]
    neg_cols = [c for c in ["saida_remanej", "anulado", "transferido"] if c in df.columns]
    df["credito_sipac"] = df[pos_cols].fillna(0).sum(axis=1) - df[neg_cols].fillna(0).sum(axis=1)

    # Construir mapeamento código → nome da unidade
    nomes_unidades = {}
    if "cod_unidade" in df.columns and "unidade_nome" in df.columns:
        for _, r in df[["cod_unidade", "unidade_nome"]].drop_duplicates().iterrows():
            nomes_unidades[str(r["cod_unidade"]).strip()] = str(r["unidade_nome"]).strip()

    return df, nomes_unidades


# ============================================================
# NORMALIZAÇÃO DE PI
# ============================================================

PI_SEM_INFO = {"-8", "ND", "nan", "", "SEM INFORMACAO"}


def normalize_pi(pi_str):
    """Normaliza PI para comparação. SIAFI '-8' <-> SIPAC 'ND'."""
    pi = str(pi_str).strip().upper()
    if pi in PI_SEM_INFO:
        return "__SEM_INFO__"
    return pi


# ============================================================
# FUNÇÕES DE MATCHING DE FONTE
# ============================================================

def normalize_fonte(fonte_str):
    """Remove letras, mantém apenas dígitos."""
    return re.sub(r"[^0-9]", "", str(fonte_str))


def fonte_similarity(f1, f2):
    """Similaridade entre duas fontes (normalizado)."""
    n1, n2 = normalize_fonte(f1), normalize_fonte(f2)
    if n1 == n2:
        return 1.0
    seq_ratio = SequenceMatcher(None, n1, n2).ratio()
    prefix = 0
    for a, b in zip(n1, n2):
        if a == b:
            prefix += 1
        else:
            break
    prefix_ratio = prefix / max(len(n1), len(n2), 1)
    return 0.6 * seq_ratio + 0.4 * prefix_ratio


def find_best_fonte(siafi_fonte, sipac_fontes):
    """Retorna (melhor_fonte_sipac, score)."""
    if not sipac_fontes:
        return None, 0.0
    best, best_score = None, 0.0
    for sf in sipac_fontes:
        score = fonte_similarity(siafi_fonte, sf)
        if score > best_score:
            best_score = score
            best = sf
    return best, best_score


# ============================================================
# CONFERÊNCIA PRINCIPAL
# ============================================================

def run_conference(siafi_df, sipac_df):
    """Executa a conferência e retorna resultados categorizados."""

    sipac = sipac_df.copy()
    siafi = siafi_df.copy()

    # Normalizar PI em ambos os sistemas
    siafi["pi_norm"] = siafi["pi"].apply(normalize_pi)
    sipac["pi_norm"] = sipac["pi"].apply(normalize_pi)

    # Mapear SIAFI -> UNIDADE SIPAC
    siafi["unidade_sipac"] = siafi.apply(
        lambda r: MAPPING_DICT.get((str(r["ug"]).strip(), str(r["ugr"]).strip())), axis=1
    )

    # Separar linhas com/sem mapeamento
    sem_info = ["-8", "nan", "", "SEM INFORMACAO"]
    mask_no_map = siafi["unidade_sipac"].isna()
    mask_ugr_sem_info = siafi["ugr"].isin(sem_info)

    falha_map = siafi[mask_no_map & ~mask_ugr_sem_info].copy()
    siafi_sem_ugr = siafi[mask_no_map & mask_ugr_sem_info].copy()
    siafi_ok = siafi[~mask_no_map].copy()

    conferidos, sem_sipac, fonte_fuzzy = [], [], []
    sipac_matched_idx = set()

    # ----- Conferir linhas SIAFI mapeadas -----
    for idx, row in siafi_ok.iterrows():
        nat = str(row["natureza"]).strip()
        ptres = str(row["ptres"]).strip()
        esfera = str(row["esfera"]).strip()
        pi_norm = str(row["pi_norm"]).strip()
        unidade = str(row["unidade_sipac"]).strip()
        fonte_siafi = str(row["fonte"]).strip()

        mask = (
            (sipac["nat_despesa"] == nat)
            & (sipac["ptres"] == ptres)
            & (sipac["esfera"] == esfera)
            & (sipac["pi_norm"] == pi_norm)
        )
        if "cod_unidade" in sipac.columns:
            mask = mask & (sipac["cod_unidade"].str.startswith(unidade))

        candidates = sipac[mask]

        if candidates.empty:
            sem_sipac.append(_build_sem_sipac_row(row, unidade))
            continue

        # Marcar TODOS os candidatos SIPAC como encontrados (mesma nat/ptres/esf/pi/unidade)
        for si in candidates.index:
            sipac_matched_idx.add(si)

        cand_fontes = candidates["fonte_recurso"].unique().tolist()
        best_fonte, fscore = find_best_fonte(fonte_siafi, cand_fontes)

        if best_fonte is None:
            sem_sipac.append(_build_sem_sipac_row(row, unidade))
            continue

        sipac_match = candidates[candidates["fonte_recurso"] == best_fonte]

        credito_sipac = sipac_match["credito_sipac"].sum()
        empenhos_sipac = sipac_match["empenhos"].sum() if "empenhos" in sipac_match.columns else 0

        rec = _build_match_record(row, unidade, fonte_siafi, best_fonte, fscore, credito_sipac, empenhos_sipac)

        if fscore < 1.0:
            fonte_fuzzy.append(rec)

        conferidos.append(rec)

    # ----- Conferir linhas SIAFI sem UGR (usando UG direto) -----
    for idx, row in siafi_sem_ugr.iterrows():
        nat = str(row["natureza"]).strip()
        ptres = str(row["ptres"]).strip()
        esfera = str(row["esfera"]).strip()
        pi_norm = str(row["pi_norm"]).strip()
        fonte_siafi = str(row["fonte"]).strip()
        ug = str(row["ug"]).strip()

        mask = (
            (sipac["nat_despesa"] == nat)
            & (sipac["ptres"] == ptres)
            & (sipac["esfera"] == esfera)
            & (sipac["pi_norm"] == pi_norm)
        )
        if "ug_siafi" in sipac.columns:
            mask = mask & (sipac["ug_siafi"].astype(str).str.strip().str.replace(".0", "", regex=False) == ug)

        candidates = sipac[mask]
        if candidates.empty:
            sem_sipac.append(_build_sem_sipac_row(row, f"UG={ug} (sem UGR)"))
            continue

        # Marcar TODOS os candidatos SIPAC como encontrados
        for si in candidates.index:
            sipac_matched_idx.add(si)

        cand_fontes = candidates["fonte_recurso"].unique().tolist()
        best_fonte, fscore = find_best_fonte(fonte_siafi, cand_fontes)
        if best_fonte is None:
            sem_sipac.append(_build_sem_sipac_row(row, f"UG={ug} (sem UGR)"))
            continue

        sipac_match = candidates[candidates["fonte_recurso"] == best_fonte]

        credito_sipac = sipac_match["credito_sipac"].sum()
        empenhos_sipac = sipac_match["empenhos"].sum() if "empenhos" in sipac_match.columns else 0

        rec = _build_match_record(row, f"UG={ug} (sem UGR)", fonte_siafi, best_fonte, fscore, credito_sipac, empenhos_sipac)

        if fscore < 1.0:
            fonte_fuzzy.append(rec)

        conferidos.append(rec)

    # ----- Conferência reversa: SIPAC → SIAFI via mapeamento reverso -----
    reverse_map = {}
    for ug_code, ugr_code, unidade_code in MAPPING_DATA:
        if unidade_code not in reverse_map:
            reverse_map[unidade_code] = set()
        reverse_map[unidade_code].add(ug_code)

    sipac_remaining = sipac[~sipac.index.isin(sipac_matched_idx)].copy()
    for idx, row in sipac_remaining.iterrows():
        cod = str(row.get("cod_unidade", "")).strip()
        nat = str(row.get("nat_despesa", "")).strip()
        ptres = str(row.get("ptres", "")).strip()
        esfera = str(row.get("esfera", "")).strip()
        pi_n = str(row.get("pi_norm", "")).strip()

        # Encontrar unidade do mapeamento que é prefixo do cod_unidade SIPAC
        possible_ugs = set()
        for unidade_code, ugs in reverse_map.items():
            if cod.startswith(unidade_code) or unidade_code.startswith(cod):
                possible_ugs.update(ugs)

        if not possible_ugs:
            continue

        # Buscar SIAFI com mesma chave + UG do mapeamento reverso
        mask_siafi = (
            (siafi["natureza"] == nat)
            & (siafi["ptres"] == ptres)
            & (siafi["esfera"] == esfera)
            & (siafi["pi_norm"] == pi_n)
            & (siafi["ug"].isin(possible_ugs))
        )
        if mask_siafi.any():
            sipac_matched_idx.add(idx)

    # ----- SIPAC sem correspondência no SIAFI -----
    sipac_unmatched = sipac[~sipac.index.isin(sipac_matched_idx)].copy()
    has_value = (sipac_unmatched.get("credito_sipac", pd.Series(dtype=float)).abs() > 0.01) | (
        sipac_unmatched.get("empenhos", pd.Series(dtype=float)).abs() > 0.01
    )
    sipac_unmatched = sipac_unmatched[has_value]

    sipac_sem_siafi = []
    for _, r in sipac_unmatched.iterrows():
        sipac_sem_siafi.append({
            "nat_despesa": r.get("nat_despesa", ""),
            "ptres": r.get("ptres", ""),
            "fonte_recurso": r.get("fonte_recurso", ""),
            "esfera": r.get("esfera", ""),
            "pi": r.get("pi", ""),
            "unidade": r.get("unidade_nome", ""),
            "cod_unidade": r.get("cod_unidade", ""),
            "credito_sipac": r.get("credito_sipac", 0),
            "empenhos": r.get("empenhos", 0),
            "saldo": r.get("saldo", 0),
        })

    # ----- Falhas de mapeamento -----
    falha_list = []
    for _, r in falha_map.iterrows():
        falha_list.append({
            "natureza": r["natureza"],
            "ptres": r["ptres"],
            "fonte": r["fonte"],
            "esfera": r["esfera"],
            "pi": r["pi"],
            "ug": r["ug"],
            "ugr": r["ugr"],
            "credito_total_siafi": r["credito_total_siafi"],
            "motivo": f"UG={r['ug']} + UGR={r['ugr']} não encontrado na tabela de mapeamento",
        })

    df_conferidos = pd.DataFrame(conferidos) if conferidos else pd.DataFrame()

    return {
        "conferidos": df_conferidos,
        "sem_sipac": pd.DataFrame(sem_sipac) if sem_sipac else pd.DataFrame(),
        "fonte_fuzzy": pd.DataFrame(fonte_fuzzy) if fonte_fuzzy else pd.DataFrame(),
        "sipac_sem_siafi": pd.DataFrame(sipac_sem_siafi) if sipac_sem_siafi else pd.DataFrame(),
        "falha_map": pd.DataFrame(falha_list) if falha_list else pd.DataFrame(),
        "total_siafi": len(siafi),
        "total_sipac": len(sipac),
    }


def _build_match_record(row, unidade, fonte_siafi, best_fonte, fscore, credito_sipac, empenhos_sipac):
    return {
        "natureza": str(row["natureza"]).strip(),
        "natureza_desc": row.get("natureza_desc", ""),
        "ptres": str(row["ptres"]).strip(),
        "fonte_siafi": fonte_siafi,
        "fonte_sipac": best_fonte,
        "fonte_score": round(fscore, 4),
        "esfera": str(row["esfera"]).strip(),
        "pi": str(row["pi"]).strip(),
        "pi_desc": row.get("pi_desc", ""),
        "ug": row["ug"],
        "ugr": row["ugr"],
        "unidade_sipac": unidade,
        "credito_total_siafi": row["credito_total_siafi"],
        "credito_sipac": credito_sipac,
        "diff_credito": round(row["credito_total_siafi"] - credito_sipac, 2),
        "empenhado_siafi": row["empenhado"],
        "empenhos_sipac": empenhos_sipac,
        "diff_empenhado": round(row["empenhado"] - empenhos_sipac, 2),
    }


def _build_sem_sipac_row(row, unidade_info):
    return {
        "natureza": row["natureza"],
        "natureza_desc": row.get("natureza_desc", ""),
        "ptres": row["ptres"],
        "fonte_siafi": row["fonte"],
        "fonte_desc": row.get("fonte_desc", ""),
        "esfera": row["esfera"],
        "pi": row["pi"],
        "pi_desc": row.get("pi_desc", ""),
        "ug": row["ug"],
        "ugr": row["ugr"],
        "unidade_sipac_esperada": unidade_info,
        "credito_total_siafi": row["credito_total_siafi"],
        "empenhado_siafi": row["empenhado"],
    }


# ============================================================
# UTILITÁRIOS
# ============================================================

def to_excel_bytes(df):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dados")
    return output.getvalue()


def download_button(df, filename, label):
    if df is not None and not df.empty:
        st.download_button(
            label=label,
            data=to_excel_bytes(df),
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )


def apply_filters(df, filters):
    """Aplica múltiplos filtros ao DataFrame."""
    if df.empty:
        return df
    result = df.copy()
    for col_candidates, value in filters.items():
        if not value or value == "Todas":
            continue
        # col_candidates é uma tupla de possíveis nomes de coluna
        col = None
        for c in col_candidates:
            if c in result.columns:
                col = c
                break
        if col is None:
            continue
        result = result[result[col].astype(str).str.upper().str.contains(value.upper(), na=False)]
    return result


def get_all_unidades(res, nomes_unidades):
    """Extrai todas as unidades SIPAC únicas dos resultados, com nomes."""
    unidades = set()
    for key in ["conferidos", "sem_sipac", "fonte_fuzzy"]:
        df = res.get(key, pd.DataFrame())
        if not df.empty:
            col = "unidade_sipac" if "unidade_sipac" in df.columns else "unidade_sipac_esperada"
            if col in df.columns:
                unidades.update(df[col].astype(str).unique())
    for key in ["sipac_sem_siafi"]:
        df = res.get(key, pd.DataFrame())
        if not df.empty and "cod_unidade" in df.columns:
            unidades.update(df["cod_unidade"].astype(str).unique())
    # Retornar lista de (código, "código - nome") para o selectbox
    result = []
    for code in sorted(unidades):
        nome = nomes_unidades.get(code, "")
        label = f"{code} - {nome}" if nome else code
        result.append((code, label))
    return result


# ============================================================
# INTERFACE PRINCIPAL
# ============================================================

def main():
    st.title("🔍 Conferência SIPAC x SIAFI")
    st.markdown(
        "Ferramenta de reconciliação orçamentária — "
        "identifica divergências entre lançamentos SIAFI e SIPAC."
    )

    # ----- SIDEBAR: Upload de Arquivos -----
    with st.sidebar:
        st.header("📁 Importação de Dados")

        st.subheader("1. Arquivo SIAFI")
        siafi_file = st.file_uploader(
            "Dados SIAFI do dia (estrutura 'conferência sipac')",
            type=["xlsx", "xls"],
            key="siafi",
        )

        st.subheader("2. Arquivo SIPAC")
        sipac_file = st.file_uploader(
            "Saldo orçamentário SIPAC do dia",
            type=["xlsx", "xls"],
            key="sipac",
        )

        st.divider()
        st.caption("Mapeamento UG/UGR → Unidade SIPAC embutido na aplicação.")

    # ----- Validar uploads -----
    if not all([siafi_file, sipac_file]):
        st.info(
            "Importe os dois arquivos na barra lateral para iniciar a conferência:\n"
            "1. **SIAFI** — arquivo de conferência com créditos/empenhos do dia\n"
            "2. **SIPAC** — saldo orçamentário exportado do SIPAC\n\n"
            "O mapeamento UG/UGR → Unidade SIPAC já está embutido na aplicação."
        )
        return

    # ----- Parse dos arquivos -----
    with st.spinner("Processando arquivos..."):
        siafi_df = parse_siafi(siafi_file)
        sipac_result = parse_sipac(sipac_file)
        sipac_df, nomes_unidades = sipac_result

    if siafi_df is None or sipac_df is None:
        st.error("Erro ao processar os arquivos. Verifique a estrutura.")
        return

    # ----- Preview dos dados -----
    with st.expander("Pré-visualização dos dados importados", expanded=False):
        tab_s, tab_p, tab_m = st.tabs(["SIAFI", "SIPAC", "Mapeamento"])
        with tab_s:
            st.caption(f"{len(siafi_df)} linhas importadas (naturezas com final '00' excluídas)")
            st.dataframe(siafi_df.head(50), use_container_width=True, height=300)
        with tab_p:
            st.caption(f"{len(sipac_df)} linhas importadas (filtrado: ano 2026)")
            st.dataframe(sipac_df.head(50), use_container_width=True, height=300)
        with tab_m:
            mapping_df = pd.DataFrame(MAPPING_DATA, columns=["UG", "UGR", "Unidade SIPAC"])
            st.caption(f"{len(mapping_df)} mapeamentos (embutido)")
            st.dataframe(mapping_df, use_container_width=True, height=300)

    # ----- Executar conferência -----
    st.divider()
    if st.button("Executar Conferência", type="primary", use_container_width=True):
        with st.spinner("Executando conferência SIAFI x SIPAC..."):
            results = run_conference(siafi_df, sipac_df)
        st.session_state["results"] = results
        st.session_state["nomes_unidades"] = nomes_unidades

    # ----- Exibir resultados -----
    if "results" not in st.session_state:
        return

    res = st.session_state["results"]
    nomes = st.session_state.get("nomes_unidades", {})
    st.divider()
    show_results(res, nomes)


def show_results(res, nomes_unidades=None):
    """Exibe os resultados da conferência."""
    if nomes_unidades is None:
        nomes_unidades = {}

    conferidos = res["conferidos"]
    total_siafi = res["total_siafi"]
    n_conferidos = len(conferidos)
    n_sem_sipac = len(res["sem_sipac"])
    n_fonte_fuzzy = len(res["fonte_fuzzy"])
    n_sipac_sem = len(res["sipac_sem_siafi"])
    n_falha = len(res["falha_map"])

    # Calcular divergências a partir dos conferidos
    n_div_credito = 0
    n_div_empenho = 0
    if not conferidos.empty:
        n_div_credito = len(conferidos[conferidos["diff_credito"].abs() > 0.01])
        n_div_empenho = len(conferidos[conferidos["diff_empenhado"].abs() > 0.01])

    # ----- Métricas resumo -----
    st.subheader("Resumo da Conferência")
    c1, c2, c3, c4, c5, c6 = st.columns(6)
    c1.metric("Conferidos", n_conferidos)
    c2.metric("SIAFI sem SIPAC", n_sem_sipac, delta=f"-{n_sem_sipac}" if n_sem_sipac else "0", delta_color="inverse")
    c3.metric("Erro Crédito", n_div_credito, delta=f"-{n_div_credito}" if n_div_credito else "0", delta_color="inverse")
    c4.metric("Erro Empenho", n_div_empenho, delta=f"-{n_div_empenho}" if n_div_empenho else "0", delta_color="inverse")
    c5.metric("SIPAC sem SIAFI", n_sipac_sem, delta=f"-{n_sipac_sem}" if n_sipac_sem else "0", delta_color="inverse")
    c6.metric("Fonte Aprox.", n_fonte_fuzzy)

    st.info(
        f"**Cobertura:** {n_conferidos} de {total_siafi} linhas SIAFI encontraram correspondência no SIPAC. "
        f"{n_sem_sipac} sem correspondência. {n_falha} falha de mapeamento. "
        f"SIPAC analisadas: {res['total_sipac']}."
    )

    # ----- Filtros -----
    st.subheader("Filtros")
    unidades_list = get_all_unidades(res, nomes_unidades)
    labels = ["Todas"] + [label for _, label in unidades_list]
    codes = ["Todas"] + [code for code, _ in unidades_list]
    selected_label = st.selectbox("Unidade SIPAC", options=labels, index=0)
    f_unidade = codes[labels.index(selected_label)]

    fc1, fc2, fc3, fc4, fc5 = st.columns(5)
    f_natureza = fc1.text_input("Natureza", placeholder="ex: 339049")
    f_ptres = fc2.text_input("PTRES", placeholder="ex: 230106")
    f_pi = fc3.text_input("Plano Interno", placeholder="ex: V0000N01OXN")
    f_ug = fc4.text_input("UG", placeholder="ex: 153065")
    f_ugr = fc5.text_input("UGR", placeholder="ex: 150658")

    # Montar dicionário de filtros (chave = tupla de possíveis nomes de coluna)
    filters = {
        ("unidade_sipac", "unidade_sipac_esperada", "cod_unidade"): f_unidade,
        ("natureza", "nat_despesa"): f_natureza,
        ("ptres",): f_ptres,
        ("pi",): f_pi,
        ("ug",): f_ug,
        ("ugr",): f_ugr,
    }

    # ----- Definir colunas de crédito e empenho -----
    COLS_CREDITO = [
        "natureza", "natureza_desc", "ptres", "fonte_siafi", "fonte_sipac",
        "fonte_score", "esfera", "pi", "pi_desc", "ug", "ugr", "unidade_sipac",
        "credito_total_siafi", "credito_sipac", "diff_credito",
    ]
    COLS_EMPENHO = [
        "natureza", "natureza_desc", "ptres", "fonte_siafi", "fonte_sipac",
        "fonte_score", "esfera", "pi", "pi_desc", "ug", "ugr", "unidade_sipac",
        "empenhado_siafi", "empenhos_sipac", "diff_empenhado",
    ]
    CFG_CREDITO = {
        "credito_total_siafi": st.column_config.NumberColumn("Crédito SIAFI", format="R$ %.2f"),
        "credito_sipac": st.column_config.NumberColumn("Crédito SIPAC", format="R$ %.2f"),
        "diff_credito": st.column_config.NumberColumn("Diferença", format="R$ %.2f"),
        "fonte_score": st.column_config.ProgressColumn("Score Fonte", min_value=0, max_value=1),
    }
    CFG_EMPENHO = {
        "empenhado_siafi": st.column_config.NumberColumn("Empenhado SIAFI", format="R$ %.2f"),
        "empenhos_sipac": st.column_config.NumberColumn("Empenhos SIPAC", format="R$ %.2f"),
        "diff_empenhado": st.column_config.NumberColumn("Diferença", format="R$ %.2f"),
        "fonte_score": st.column_config.ProgressColumn("Score Fonte", min_value=0, max_value=1),
    }

    # ----- Abas -----
    tabs = st.tabs([
        f"Fontes Ajustadas ({n_fonte_fuzzy})",
        f"Créditos - Todos ({n_conferidos})",
        f"Créditos - Erros ({n_div_credito})",
        f"Empenhos - Todos ({n_conferidos})",
        f"Empenhos - Erros ({n_div_empenho})",
        f"SIAFI sem SIPAC ({n_sem_sipac})",
        f"SIPAC sem SIAFI ({n_sipac_sem})",
    ])

    # --- Tab 1: Fontes Ajustadas ---
    with tabs[0]:
        st.markdown(
            "**Lançamentos onde a fonte SIAFI foi associada à fonte SIPAC mais próxima** "
            "(o SIPAC não permite letras na fonte). Verifique se a associação está correta."
        )
        df_view = apply_filters(res["fonte_fuzzy"], filters)
        if not df_view.empty:
            st.dataframe(df_view, use_container_width=True, height=400,
                         column_config={"fonte_score": st.column_config.ProgressColumn("Score", min_value=0, max_value=1)})
            download_button(df_view, "fontes_ajustadas.xlsx", "Baixar Fontes Ajustadas")
        else:
            st.info("Todas as fontes tiveram correspondência exata.")

    # --- Tab 2: Créditos - Todos (certas + erradas) ---
    with tabs[1]:
        st.markdown(
            "**Todos os créditos analisados (SIAFI x SIPAC).** "
            "SIAFI: crédito + empenhado. "
            "SIPAC: (Distribuído + Recebido + Entrada Remanej.) - (Saída Remanej. + Anulado + Transferido)."
        )
        df_view = apply_filters(conferidos, filters)
        if not df_view.empty:
            cols = [c for c in COLS_CREDITO if c in df_view.columns]
            st.dataframe(df_view[cols], use_container_width=True, height=400, column_config=CFG_CREDITO)
            download_button(df_view[cols], "creditos_todos.xlsx", "Baixar Créditos - Todos")
        else:
            st.info("Nenhum registro.")

    # --- Tab 3: Créditos - Erros (só divergências) ---
    with tabs[2]:
        st.markdown("**Apenas linhas com divergência de CRÉDITO.**")
        df_all = apply_filters(conferidos, filters)
        df_view = df_all[df_all["diff_credito"].abs() > 0.01] if not df_all.empty else pd.DataFrame()
        if not df_view.empty:
            cols = [c for c in COLS_CREDITO if c in df_view.columns]
            st.dataframe(df_view[cols], use_container_width=True, height=400, column_config=CFG_CREDITO)
            download_button(df_view[cols], "creditos_erros.xlsx", "Baixar Créditos - Erros")
        else:
            st.success("Todos os créditos conferem" + ("." if f_unidade == "Todas" and not any([f_natureza, f_ptres, f_pi, f_ug, f_ugr]) else " para os filtros selecionados."))

    # --- Tab 4: Empenhos - Todos (certas + erradas) ---
    with tabs[3]:
        st.markdown(
            "**Todos os empenhos analisados (SIAFI x SIPAC).** "
            "Compara empenhado SIAFI com a coluna Empenhos do SIPAC."
        )
        df_view = apply_filters(conferidos, filters)
        if not df_view.empty:
            cols = [c for c in COLS_EMPENHO if c in df_view.columns]
            st.dataframe(df_view[cols], use_container_width=True, height=400, column_config=CFG_EMPENHO)
            download_button(df_view[cols], "empenhos_todos.xlsx", "Baixar Empenhos - Todos")
        else:
            st.info("Nenhum registro.")

    # --- Tab 5: Empenhos - Erros (só divergências) ---
    with tabs[4]:
        st.markdown("**Apenas linhas com divergência de EMPENHO.**")
        df_all = apply_filters(conferidos, filters)
        df_view = df_all[df_all["diff_empenhado"].abs() > 0.01] if not df_all.empty else pd.DataFrame()
        if not df_view.empty:
            cols = [c for c in COLS_EMPENHO if c in df_view.columns]
            st.dataframe(df_view[cols], use_container_width=True, height=400, column_config=CFG_EMPENHO)
            download_button(df_view[cols], "empenhos_erros.xlsx", "Baixar Empenhos - Erros")
        else:
            st.success("Todos os empenhos conferem" + ("." if f_unidade == "Todas" and not any([f_natureza, f_ptres, f_pi, f_ug, f_ugr]) else " para os filtros selecionados."))

    # --- Tab 6: SIAFI sem SIPAC ---
    with tabs[5]:
        st.markdown("**Lançamentos no SIAFI que NÃO foram encontrados no SIPAC.**")
        df_view = apply_filters(res["sem_sipac"], filters)
        if not df_view.empty:
            st.dataframe(df_view, use_container_width=True, height=400,
                         column_config={
                             "credito_total_siafi": st.column_config.NumberColumn("Crédito SIAFI", format="R$ %.2f"),
                             "empenhado_siafi": st.column_config.NumberColumn("Empenhado SIAFI", format="R$ %.2f"),
                         })
            download_button(df_view, "siafi_sem_sipac.xlsx", "Baixar SIAFI sem SIPAC")
        else:
            st.success("Todos os lançamentos SIAFI possuem correspondência no SIPAC.")

    # --- Tab 7: SIPAC sem SIAFI ---
    with tabs[6]:
        st.markdown("**Lançamentos no SIPAC que NÃO possuem correspondência no SIAFI.**")
        df_view = apply_filters(res["sipac_sem_siafi"], filters)
        if not df_view.empty:
            st.dataframe(df_view, use_container_width=True, height=400,
                         column_config={
                             "credito_sipac": st.column_config.NumberColumn("Crédito SIPAC", format="R$ %.2f"),
                             "empenhos": st.column_config.NumberColumn("Empenhos", format="R$ %.2f"),
                             "saldo": st.column_config.NumberColumn("Saldo", format="R$ %.2f"),
                         })
            download_button(df_view, "sipac_sem_siafi.xlsx", "Baixar SIPAC sem SIAFI")
        else:
            st.success("Todos os lançamentos SIPAC possuem correspondência no SIAFI.")

    # ----- Relatório completo para download -----
    st.divider()
    st.subheader("Relatório Completo")
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        if not conferidos.empty:
            cols_c = [c for c in COLS_CREDITO if c in conferidos.columns]
            cols_e = [c for c in COLS_EMPENHO if c in conferidos.columns]
            conferidos[cols_c].to_excel(writer, index=False, sheet_name="Créditos - Todos")
            div_c = conferidos[conferidos["diff_credito"].abs() > 0.01]
            if not div_c.empty:
                div_c[cols_c].to_excel(writer, index=False, sheet_name="Créditos - Erros")
            conferidos[cols_e].to_excel(writer, index=False, sheet_name="Empenhos - Todos")
            div_e = conferidos[conferidos["diff_empenhado"].abs() > 0.01]
            if not div_e.empty:
                div_e[cols_e].to_excel(writer, index=False, sheet_name="Empenhos - Erros")
        if not res["fonte_fuzzy"].empty:
            res["fonte_fuzzy"].to_excel(writer, index=False, sheet_name="Fontes Ajustadas")
        if not res["sem_sipac"].empty:
            res["sem_sipac"].to_excel(writer, index=False, sheet_name="SIAFI sem SIPAC")
        if not res["sipac_sem_siafi"].empty:
            res["sipac_sem_siafi"].to_excel(writer, index=False, sheet_name="SIPAC sem SIAFI")
        if not res["falha_map"].empty:
            res["falha_map"].to_excel(writer, index=False, sheet_name="Falha Mapeamento")
        resumo = pd.DataFrame([{
            "Total linhas SIAFI": total_siafi,
            "Total linhas SIPAC": res["total_sipac"],
            "Conferidos": n_conferidos,
            "SIAFI sem SIPAC": n_sem_sipac,
            "Erros Crédito": n_div_credito,
            "Erros Empenho": n_div_empenho,
            "Fontes Ajustadas": n_fonte_fuzzy,
            "SIPAC sem SIAFI": n_sipac_sem,
            "Falha Mapeamento": n_falha,
        }])
        resumo.to_excel(writer, index=False, sheet_name="Resumo")

    st.download_button(
        label="Baixar Relatório Completo (Excel)",
        data=output.getvalue(),
        file_name="relatorio_conferencia_sipac_siafi.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )


if __name__ == "__main__":
    main()
