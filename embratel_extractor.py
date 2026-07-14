# -*- coding: utf-8 -*-
"""
embratel_extractor.py

Módulo de extração de dados das faturas EMBRATEL/Claro para ArcelorMittal.

Cada "trio" de arquivos é composto por:
  - *_PDF.pdf  -> Boleto / Conta de Prestação de Serviços (1 página, dados
                  cadastrais e totais da fatura)
  - *_RPS.pdf  -> Recibos Provisórios de Serviço (1 página por filial/CNPJ,
                  imposto = ISS)
  - *_NF.pdf   -> Notas Fiscais de Telecomunicação / NFCom (1 página por
                  filial/UF, imposto = ICMS)

Regra de negócio (definida pelo usuário):
  Cada linha da planilha final representa a SOMA de todos os valores de
  base de cálculo (RPS: valor total do serviço: NF: base de cálculo do
  ICMS) que compartilham a MESMA alíquota (coluna "Tipo"), somando RPS e
  NF juntos, dentro do mesmo trio de arquivos.

O texto dos PDFs de exemplo é texto real (não é imagem escaneada), então a
extração usa pdfplumber. Caso apareça algum PDF escaneado, há um fallback
opcional de OCR (pytesseract) acionado automaticamente quando uma página
não devolve texto.
"""

from __future__ import annotations

import io
import re
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pdfplumber

logger = logging.getLogger("embratel_extractor")
logger.setLevel(logging.INFO)

# --------------------------------------------------------------------------
# Colunas finais, na ordem exata do modelo EMBRATEL_COSS_13071100.xlsx
# --------------------------------------------------------------------------
OUTPUT_COLUMNS = [
    "CNPJ Arcelor",
    "CNPJ Fornecedor",
    "Data Emissao",
    "Data VENCIMENTO",
    "Codigo Cliente",
    "Numero NF",
    "Numero Fatura",
    "Valor",
    "Tipo",
    "Texto Boleto",
    "Código da Fatura",
    "Código de Barras",
    "Numero NFCom",
    "Serie NFCom",
    "Chave de Acesso",
    "Protocolo de Autorização",
]


# --------------------------------------------------------------------------
# Utilidades de número / texto em formato BR
# --------------------------------------------------------------------------
def br_to_float(value: str) -> Optional[float]:
    """Converte '1.234,56' -> 1234.56 . Retorna None se não for número."""
    if value is None:
        return None
    v = value.strip()
    if not v:
        return None
    v = v.replace(".", "").replace(",", ".")
    try:
        return float(v)
    except ValueError:
        return None


def norm_aliquota(value: Optional[float]) -> Optional[float]:
    """Normaliza a alíquota para 2 casas decimais, usada como chave de
    agrupamento (evita que 5.00 e 5.001 virem grupos diferentes)."""
    if value is None:
        return None
    return round(value, 2)


def get_page_text(page) -> str:
    """Extrai texto da página; usa OCR como fallback se a página não tiver
    texto embutido (PDF escaneado / imagem)."""
    text = page.extract_text() or ""
    if text.strip():
        return text
    # Fallback OCR (opcional) -----------------------------------------
    try:
        import pytesseract
        from PIL import Image

        im = page.to_image(resolution=300).original
        if not isinstance(im, Image.Image):
            im = Image.open(io.BytesIO(im))
        text = pytesseract.image_to_string(im, lang="por")
        logger.info("Página sem texto embutido — usado OCR como fallback.")
    except Exception as exc:  # pragma: no cover
        logger.warning("Falha no fallback de OCR: %s", exc)
        text = ""
    return text


# --------------------------------------------------------------------------
# Identificação e agrupamento dos trios de arquivo
# --------------------------------------------------------------------------
SUFFIX_RE = re.compile(r"_(PDF|NF|RPS)\.pdf$", re.IGNORECASE)


def classify_file(filename: str) -> Optional[str]:
    """Retorna 'PDF', 'NF' ou 'RPS' a partir do nome do arquivo, ou None."""
    m = SUFFIX_RE.search(filename)
    if not m:
        return None
    return m.group(1).upper()


def group_trios(paths: List[Path]) -> Dict[str, Dict[str, Path]]:
    """Agrupa uma lista de caminhos de PDF em trios, pela parte do nome que
    antecede o sufixo _PDF/_NF/_RPS."""
    groups: Dict[str, Dict[str, Path]] = {}
    for p in paths:
        kind = classify_file(p.name)
        if kind is None:
            logger.warning("Arquivo ignorado (sufixo não reconhecido): %s", p.name)
            continue
        key = SUFFIX_RE.sub("", p.name)
        groups.setdefault(key, {})[kind] = p
    return groups


# --------------------------------------------------------------------------
# Extração de "Código da Fatura" a partir do nome do arquivo
# (ex.: ..._00007624776-0126-DADOS_2026-06-28_NF.pdf -> 00007624776-0126-DADOS)
# --------------------------------------------------------------------------
CODIGO_FATURA_RE = re.compile(r"(\d{6,}-\d{2,}-DADOS)")


def codigo_fatura_from_filename(name: str) -> str:
    m = CODIGO_FATURA_RE.search(name)
    return m.group(1) if m else ""


# --------------------------------------------------------------------------
# Parsing do BOLETO (_PDF)
# --------------------------------------------------------------------------
@dataclass
class BoletoData:
    cnpj_arcelor: str = ""
    cnpj_fornecedor: str = ""
    data_emissao: str = ""
    data_vencimento: str = ""
    codigo_cliente: str = ""
    numero_fatura: str = ""
    total_a_pagar: Optional[float] = None
    codigo_debito_automatico: str = ""
    codigo_barras: str = ""


def parse_boleto(path: Path) -> BoletoData:
    with pdfplumber.open(path) as pdf:
        text = get_page_text(pdf.pages[0])

    data = BoletoData()

    m = re.search(r"CNPJ:\s*([\d./-]+)\s*-\s*IE", text)
    if m:
        data.cnpj_fornecedor = m.group(1)

    m = re.search(r"CPF/CNPJ:\s*([\d./-]+)", text)
    if m:
        data.cnpj_arcelor = m.group(1)

    m = re.search(r"Data de Emiss[ãa]o:\s*(\d{2}/\d{2}/\d{4})", text)
    if m:
        data.data_emissao = m.group(1)

    m = re.search(r"N[uú]mero\.?\s*da Fatura:\s*([\w/\-]+)", text)
    if m:
        data.numero_fatura = m.group(1)

    m = re.search(r"(\d{9,12}\s*-\s*\d{3,5})\s+(\d{2}/\d{2}/\d{4})\s+([\d.,]+)", text)
    if m:
        data.codigo_cliente = re.sub(r"\s*-\s*", " - ", m.group(1)).strip()
        data.data_vencimento = m.group(2)
        data.total_a_pagar = br_to_float(m.group(3))

    m = re.search(r"C[óo]d\.?D[ée]bito Autom[áa]tico:\s*([\w\-]+)", text)
    if m:
        data.codigo_debito_automatico = m.group(1)

    # última linha de dígitos = código de barras (4 blocos)
    m = re.search(
        r"(\d{11,14})\s+(\d{11,14})\s+(\d{11,14})\s+(\d{11,14})\s*$", text.strip(), re.M
    )
    if m:
        data.codigo_barras = "".join(m.groups())

    return data


# --------------------------------------------------------------------------
# Parsing do RPS (_RPS) -> lista de (aliquota, valor) por item de serviço
# --------------------------------------------------------------------------
RPS_HEADER_RE = re.compile(
    r"DISCRIMINA[ÇC][ÃA]O DOS SERVI[ÇC]OS.*?VALOR DO ISS", re.S
)
RPS_ITEM_RE = re.compile(
    r"^(?P<desc>.+?)\s+"
    r"(?P<qtde>[\d.,]+)\s+"
    r"(?P<v_unit>[\d.,]+)\s+"
    r"(?P<v_total>[\d.,]+)\s+"
    r"(?P<aliquota>[\d.,]+)\s+"
    r"(?P<v_iss>[\d.,]+)\s*$"
)


def parse_rps(path: Path) -> List[Tuple[float, float]]:
    """Retorna lista de (aliquota, valor_total) — um item por linha de
    serviço encontrada em cada página do RPS."""
    results: List[Tuple[float, float]] = []
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = get_page_text(page)
            block = RPS_HEADER_RE.search(text)
            if not block:
                logger.warning("RPS pág. %s: cabeçalho de serviços não encontrado.", i + 1)
                continue
            # pega as linhas entre o cabeçalho e "TOTAL DA NOTA"
            after = text[block.end():]
            stop = after.find("TOTAL DA NOTA")
            body = after[:stop] if stop != -1 else after
            found_any = False
            for line in body.splitlines():
                line = line.strip()
                if not line:
                    continue
                m = RPS_ITEM_RE.match(line)
                if m:
                    aliquota = br_to_float(m.group("aliquota"))
                    valor = br_to_float(m.group("v_total"))
                    if aliquota is not None and valor is not None:
                        results.append((aliquota, valor))
                        found_any = True
            if not found_any:
                logger.warning("RPS pág. %s: nenhum item de serviço reconhecido.", i + 1)
    return results


# --------------------------------------------------------------------------
# Parsing da NF (_NF) -> lista de (aliquota, base_calculo) por página/nota
# --------------------------------------------------------------------------
NF_TOTALS_RE = re.compile(
    r"VALOR TOTAL BASE DE C[ÁA]LCULO ICMS AL[ÍI]QUOTA VALOR DO ICMS "
    r"VALOR ISENTO VALOR OUTROS\s*\n"
    r"([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)"
)

# Campos opcionais de NFCom eletrônica (podem não existir neste layout)
NFCOM_NUMERO_RE = re.compile(r"N[úu]mero\s*NFCom[:\s]+([\w./-]+)", re.I)
NFCOM_SERIE_RE = re.compile(r"S[ée]rie\s*NFCom[:\s]+([\w./-]+)", re.I)
CHAVE_ACESSO_RE = re.compile(r"Chave de Acesso[:\s]+([\d\s]{20,})", re.I)
PROTOCOLO_RE = re.compile(r"Protocolo de Autoriza[çc][ãa]o[:\s]+([\w./-]+)", re.I)


def parse_nf(path: Path) -> Tuple[List[Tuple[float, float]], Dict[str, str]]:
    """Retorna:
      - lista de (aliquota, base_calculo_icms) — uma por página/nota
      - dict com campos de NFCom eletrônica, se encontrados em qualquer
        página (Numero NFCom, Serie NFCom, Chave de Acesso, Protocolo)
    """
    results: List[Tuple[float, float]] = []
    nfcom_extra = {
        "Numero NFCom": "",
        "Serie NFCom": "",
        "Chave de Acesso": "",
        "Protocolo de Autorização": "",
    }
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            text = get_page_text(page)

            m = NF_TOTALS_RE.search(text)
            if not m:
                logger.warning("NF pág. %s: linha de totais ICMS não encontrada.", i + 1)
                continue
            base_calculo = br_to_float(m.group(2))
            aliquota = br_to_float(m.group(3))
            if base_calculo is not None and aliquota is not None:
                results.append((aliquota, base_calculo))

            if not nfcom_extra["Numero NFCom"]:
                m2 = NFCOM_NUMERO_RE.search(text)
                if m2:
                    nfcom_extra["Numero NFCom"] = m2.group(1).strip()
            if not nfcom_extra["Serie NFCom"]:
                m2 = NFCOM_SERIE_RE.search(text)
                if m2:
                    nfcom_extra["Serie NFCom"] = m2.group(1).strip()
            if not nfcom_extra["Chave de Acesso"]:
                m2 = CHAVE_ACESSO_RE.search(text)
                if m2:
                    nfcom_extra["Chave de Acesso"] = re.sub(r"\s+", "", m2.group(1))
            if not nfcom_extra["Protocolo de Autorização"]:
                m2 = PROTOCOLO_RE.search(text)
                if m2:
                    nfcom_extra["Protocolo de Autorização"] = m2.group(1).strip()

    return results, nfcom_extra


# --------------------------------------------------------------------------
# Processamento de um trio completo -> linhas agregadas por alíquota
# --------------------------------------------------------------------------
@dataclass
class TrioResult:
    key: str
    rows: List[Dict] = field(default_factory=list)
    boleto: Optional[BoletoData] = None
    total_boleto: Optional[float] = None
    total_agregado: Optional[float] = None
    diferenca: Optional[float] = None
    avisos: List[str] = field(default_factory=list)


def process_trio(key: str, files: Dict[str, Path]) -> TrioResult:
    result = TrioResult(key=key)

    faltando = [k for k in ("PDF", "RPS", "NF") if k not in files]
    if faltando:
        result.avisos.append(
            f"Trio '{key}': arquivo(s) ausente(s) -> {', '.join(faltando)}. "
            "Linhas não puderam ser geradas para os tipos faltantes."
        )

    boleto = parse_boleto(files["PDF"]) if "PDF" in files else BoletoData()
    result.boleto = boleto
    result.total_boleto = boleto.total_a_pagar

    itens_aliquota: Dict[float, float] = {}

    if "RPS" in files:
        for aliquota, valor in parse_rps(files["RPS"]):
            a = norm_aliquota(aliquota)
            itens_aliquota[a] = itens_aliquota.get(a, 0.0) + valor

    nfcom_extra = {
        "Numero NFCom": "",
        "Serie NFCom": "",
        "Chave de Acesso": "",
        "Protocolo de Autorização": "",
    }
    if "NF" in files:
        nf_items, nfcom_extra = parse_nf(files["NF"])
        for aliquota, valor in nf_items:
            a = norm_aliquota(aliquota)
            itens_aliquota[a] = itens_aliquota.get(a, 0.0) + valor

    codigo_fatura = codigo_fatura_from_filename(
        (files.get("PDF") or files.get("NF") or files.get("RPS")).name
    )

    texto_boleto = (
        f"Valor Total do Boleto: {boleto.total_a_pagar:,.2f}".replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
        if boleto.total_a_pagar is not None
        else ""
    )

    for aliquota in sorted(itens_aliquota.keys()):
        valor = itens_aliquota[aliquota]
        row = {
            "CNPJ Arcelor": boleto.cnpj_arcelor,
            "CNPJ Fornecedor": boleto.cnpj_fornecedor,
            "Data Emissao": boleto.data_emissao,
            "Data VENCIMENTO": boleto.data_vencimento,
            "Codigo Cliente": boleto.codigo_cliente,
            "Numero NF": "",  # deliberadamente em branco (definido pelo usuário)
            "Numero Fatura": boleto.numero_fatura,
            "Valor": round(valor, 2),
            "Tipo": aliquota,
            "Texto Boleto": texto_boleto,
            "Código da Fatura": codigo_fatura,
            "Código de Barras": boleto.codigo_barras,
            "Numero NFCom": nfcom_extra["Numero NFCom"],
            "Serie NFCom": nfcom_extra["Serie NFCom"],
            "Chave de Acesso": nfcom_extra["Chave de Acesso"],
            "Protocolo de Autorização": nfcom_extra["Protocolo de Autorização"],
        }
        result.rows.append(row)

    result.total_agregado = round(sum(itens_aliquota.values()), 2)
    if result.total_boleto is not None:
        result.diferenca = round(result.total_agregado - result.total_boleto, 2)
        if abs(result.diferenca) > 0.05:
            result.avisos.append(
                f"Trio '{key}': soma agregada (R$ {result.total_agregado:,.2f}) "
                f"difere do Total a Pagar do boleto (R$ {result.total_boleto:,.2f}) "
                f"em R$ {result.diferenca:,.2f}. Vale conferir manualmente."
            )

    return result


# --------------------------------------------------------------------------
# API de alto nível
# --------------------------------------------------------------------------
def process_files(paths: List[Path]) -> Tuple[List[Dict], List[TrioResult]]:
    """Recebe uma lista de caminhos de PDF (múltiplos trios misturados),
    agrupa, processa cada trio e devolve:
      - lista de linhas (dicts) já no formato final, prontas para virar
        DataFrame / planilha
      - lista de TrioResult (um por trio) com metadados/avisos, útil para
        exibir logs/validações na interface (Streamlit)
    """
    trios = group_trios(paths)
    all_rows: List[Dict] = []
    trio_results: List[TrioResult] = []

    for key, files in sorted(trios.items()):
        try:
            tr = process_trio(key, files)
        except Exception as exc:  # nunca deixa um trio ruim derrubar o lote
            logger.error("Erro processando trio '%s': %s", key, exc)
            tr = TrioResult(key=key, avisos=[f"Trio '{key}': erro ao processar -> {exc}"])
        trio_results.append(tr)
        all_rows.extend(tr.rows)

    return all_rows, trio_results
