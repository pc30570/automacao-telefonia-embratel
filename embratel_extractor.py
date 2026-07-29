# -*- coding: utf-8 -*-
"""
embratel_extractor.py

Módulo de extração de dados das faturas EMBRATEL/Claro para ArcelorMittal.

Cada "trio" de arquivos é composto por:
  - *_PDF.pdf  -> Boleto / Conta de Prestação de Serviços (1a página, dados
                  cadastrais e totais da fatura — pode ter páginas extras
                  de mensagens)
  - *_RPS.pdf  -> Recibos Provisórios de Serviço (1 página por filial/CNPJ,
                  imposto = ISS)
  - *_NF.pdf   -> Notas Fiscais de Telecomunicação / NFCom (1 página por
                  filial/UF, imposto = ICMS)

Regra de negócio (definida pelo usuário):
  Cada linha da planilha final representa a SOMA de todos os valores de
  base de cálculo (RPS: valor total do serviço; NF: base de cálculo do
  ICMS) que compartilham a MESMA alíquota (coluna "Tipo"), somando RPS e
  NF juntos, dentro do mesmo trio de arquivos.

  Não há validação/comparação obrigatória contra o "Total a Pagar" do
  boleto: esse total pode ser maior que a soma de RPS+NF por incluir juros
  e multa, então divergência entre os dois NÃO é tratada como erro.

Texto dos PDFs:
  Alguns lotes têm texto real embutido (extração direta via pdfplumber);
  outros são escaneados (imagem, sem texto). Para páginas sem texto, o
  código cai automaticamente para OCR (pytesseract), tentando algumas
  combinações de idioma/modo de leitura até uma funcionar — sem nunca
  somar a mesma página duas vezes.
"""

from __future__ import annotations

import io
import re
import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, Iterator, List, Optional, Tuple

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

# CNPJ fixo do emissor (Claro S.A / Embratel) — usado como último recurso
# quando não é possível ler o CNPJ do fornecedor no documento (ex.: OCR
# ruim). Todos os exemplos vistos até agora usam este mesmo CNPJ.
CNPJ_FORNECEDOR_PADRAO = "40.432.544/0001-47"


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


# --------------------------------------------------------------------------
# Leitura de página com múltiplas tentativas (texto embutido -> OCR)
# --------------------------------------------------------------------------
_AVAILABLE_OCR_LANGS: Optional[List[str]] = None


def _get_ocr_langs() -> List[str]:
    """Detecta os idiomas do Tesseract disponíveis (uma vez só), preferindo
    português; usa inglês como alternativa (melhor que nada — os números e
    boa parte das palavras-chave continuam legíveis)."""
    global _AVAILABLE_OCR_LANGS
    if _AVAILABLE_OCR_LANGS is not None:
        return _AVAILABLE_OCR_LANGS
    langs = ["eng"]
    try:
        import pytesseract

        available = pytesseract.get_languages(config="")
        if "por" in available:
            langs = ["por", "eng"]
        else:
            logger.warning(
                "Pacote de idioma 'por' não encontrado no Tesseract — usando "
                "'eng' como OCR (menos preciso para acentos). Para melhor "
                "qualidade, instale: apt install tesseract-ocr-por"
            )
    except Exception as exc:  # pragma: no cover
        logger.warning("Não foi possível checar idiomas do Tesseract: %s", exc)
    _AVAILABLE_OCR_LANGS = langs
    return langs


def _ocr_variants(page) -> Iterator[str]:
    """Gera candidatos de texto via OCR, do mais para o menos confiável.
    Cada candidato é uma tentativa INDEPENDENTE (idioma x modo de leitura);
    quem consome deve usar o PRIMEIRO candidato que produzir um resultado
    válido e parar — nunca somar valores de mais de um candidato da mesma
    página, para não contar a mesma nota duas vezes."""
    try:
        import pytesseract
        from PIL import Image
    except ImportError:  # pragma: no cover
        logger.warning("pytesseract/Pillow não instalados — sem fallback de OCR.")
        return

    try:
        im = page.to_image(resolution=400).original
        if not isinstance(im, Image.Image):
            im = Image.open(io.BytesIO(im))
    except Exception as exc:  # pragma: no cover
        logger.warning("Falha ao rasterizar página para OCR: %s", exc)
        return

    # psm 6 (bloco uniforme de texto) lê tabelas com colunas bem melhor;
    # psm 3 (automático) às vezes recupera campos fora da tabela principal.
    for lang in _get_ocr_langs():
        for psm in (6, 3):
            try:
                text = pytesseract.image_to_string(im, lang=lang, config=f"--psm {psm}")
            except Exception as exc:  # pragma: no cover
                logger.warning("OCR falhou (lang=%s, psm=%s): %s", lang, psm, exc)
                continue
            if text.strip():
                yield text


def page_text_candidates(page) -> Iterator[str]:
    """Gera candidatos de texto para a página, do mais para o menos
    confiável: texto embutido primeiro (se existir, é sempre usado sozinho
    — não precisa de OCR); senão, tentativas de OCR."""
    embedded = page.extract_text() or ""
    if embedded.strip():
        yield embedded
        return
    yield from _ocr_variants(page)


def get_page_text(page) -> str:
    """Retorna o primeiro candidato de texto não vazio para a página
    (usado para campos onde não há risco de dupla contagem, ex.: dados
    cadastrais)."""
    for text in page_text_candidates(page):
        return text
    return ""


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
# Campos derivados diretamente do NOME DO ARQUIVO — muito mais confiáveis
# que OCR para estes dois campos específicos, pois seguem um padrão fixo:
#   IM_..._ARCELORMITTAL_<CNPJ 14 dígitos>_EMBRATEL_<codigo cliente>-DADOS_<data>_<TIPO>.pdf
# --------------------------------------------------------------------------
CODIGO_FATURA_RE = re.compile(r"(\d{6,}-\d{2,}-DADOS)")
CNPJ_FILENAME_RE = re.compile(r"ARCELORMITTAL_(\d{14})_")
CODIGO_CLIENTE_FILENAME_RE = re.compile(r"EMBRATEL_(\d{6,}-\d{2,})-DADOS")


def codigo_fatura_from_filename(name: str) -> str:
    m = CODIGO_FATURA_RE.search(name)
    return m.group(1) if m else ""


def format_cnpj(digits: str) -> str:
    """14 dígitos -> XX.XXX.XXX/XXXX-XX"""
    if len(digits) != 14 or not digits.isdigit():
        return digits
    return f"{digits[0:2]}.{digits[2:5]}.{digits[5:8]}/{digits[8:12]}-{digits[12:14]}"


def cnpj_arcelor_from_filename(name: str) -> str:
    m = CNPJ_FILENAME_RE.search(name)
    return format_cnpj(m.group(1)) if m else ""


def codigo_cliente_from_filename(name: str) -> str:
    m = CODIGO_CLIENTE_FILENAME_RE.search(name)
    if not m:
        return ""
    return re.sub(r"\s*-\s*", " - ", m.group(1))


# --------------------------------------------------------------------------
# Regex "tolerantes" para campos comuns — usadas em cascata sobre o texto
# do boleto, e depois (se não encontrado) sobre a 1a página do RPS e da NF,
# já que os três documentos repetem os mesmos dados cadastrais e o boleto
# costuma ser o mais difícil de ler via OCR (logotipo/código de barras
# atrapalham o layout).
# --------------------------------------------------------------------------
RE_CNPJ_FORNECEDOR = re.compile(r"CNPJ[:\s]*([\d./]+-?\d+)\s*-?\s*IE", re.I)
RE_CNPJ_ARCELOR = re.compile(r"CPF\s*/?\s*CNPJ\s*[:\.]?\s*([\d./-]{14,20})", re.I)
RE_DATA_EMISSAO = re.compile(r"Data\s*(?:de)?\s*Emiss[ãa]o\s*[:\.]?\s*(\d{2}/\d{2}/\d{4})", re.I)
RE_NUMERO_FATURA = re.compile(
    r"N[ºo°ª]?\s*\.?\s*da\s*Fatura\s*[:\.]?\s*([\w/\-]+)|"
    r"N[ºo°ª]?\s*FATURA\s*[:\.]?\s*([\w/\-]+)",
    re.I,
)
RE_CODIGO_CLIENTE_LINE = re.compile(
    r"(\d{9,14}\s*-\s*\d{3,6})\s+(\d{2}/\d{2}/\d{4})\s+([\d.,]+)"
)
RE_COD_CLIENTE_LABEL = re.compile(r"C[óo]D?\.?\s*CLIENTE\s*[:\.]?\s*(\d{6,}\s*-\s*\d{2,})", re.I)
RE_DATA_VENCIMENTO_LABEL = re.compile(r"Data\s*de\s*Vencimento\s*[:\.]?\s*(\d{2}/\d{2}/\d{4})", re.I)
RE_TOTAL_A_PAGAR_LABEL = re.compile(
    r"Total\s*a\s*Pagar\s*(?:\(R\$\)|\(RS\))?\s*[:\.]?\s*([\d.,]+)", re.I
)
RE_COD_DEBITO_AUTO = re.compile(
    r"C[óo]d\.?\s*D[ée]bito\s*Autom[áa]tico\s*[:\.]?\s*([\w\-]+\s*\-?\s*\d*)", re.I
)
# "Canhoto"/ficha de compensação no rodapé do boleto — layout tabular
# simples (sem logo/código de barras por perto), lido de forma muito mais
# confiável pelo OCR do que o cabeçalho. Ex.:
#   "CODIGO DA CONTA  NUMERO DA FATURA  DATA DE VENCIMENTO  VALOR DA CONTA"
#   "00007624776-0035  26/05/51500002-4  25/06/2026  6.873,01"
RE_CANHOTO = re.compile(
    r"(\d{6,}\s*-\s*\d{2,})\s+(\d{2}/\d{2}/\d+-?\d*)\s+(\d{2}/\d{2}/\d{4})\s+([\d.,]+)"
)


def _first_match(patterns_and_texts) -> Optional[re.Match]:
    """Tenta cada (regex, texto) na ordem dada e devolve o primeiro match."""
    for pattern, text in patterns_and_texts:
        if not text:
            continue
        m = pattern.search(text)
        if m:
            return m
    return None


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


def _extract_barcode(text: str) -> str:
    m = re.search(r"(\d{11,14})\s+(\d{11,14})\s+(\d{11,14})\s+(\d{11,14})\s*$", text.strip(), re.M)
    if m:
        return "".join(m.groups())
    return ""


def parse_boleto(path: Path, filename_hint: str = "") -> BoletoData:
    with pdfplumber.open(path) as pdf:
        boleto_text = get_page_text(pdf.pages[0])
        # páginas extras (mensagens) às vezes têm campos que a 1a não tem
        extra_text = ""
        if len(pdf.pages) > 1:
            extra_text = get_page_text(pdf.pages[1])

    data = BoletoData()
    texts = [boleto_text, extra_text]
    combined = "\n".join(t for t in texts if t)

    # --- campos preferencialmente vindos do NOME DO ARQUIVO (mais confiável
    #     que OCR para esses dois) ------------------------------------------------
    if filename_hint:
        data.cnpj_arcelor = cnpj_arcelor_from_filename(filename_hint)
        data.codigo_cliente = codigo_cliente_from_filename(filename_hint)

    # --- fallback via texto, caso o nome do arquivo não siga o padrão -----
    if not data.cnpj_arcelor:
        m = RE_CNPJ_ARCELOR.search(combined)
        if m:
            data.cnpj_arcelor = m.group(1)

    if not data.codigo_cliente:
        m = RE_CODIGO_CLIENTE_LINE.search(combined) or RE_COD_CLIENTE_LABEL.search(combined)
        if m:
            data.codigo_cliente = re.sub(r"\s*-\s*", " - ", m.group(1)).strip()

    # --- CNPJ do fornecedor (Claro/Embratel) -------------------------------
    m = RE_CNPJ_FORNECEDOR.search(combined)
    if m and len(re.sub(r"\D", "", m.group(1))) == 14:
        data.cnpj_fornecedor = m.group(1)
    else:
        data.cnpj_fornecedor = CNPJ_FORNECEDOR_PADRAO

    # --- demais campos ------------------------------------------------------
    m = RE_DATA_EMISSAO.search(combined)
    if m:
        data.data_emissao = m.group(1)

    # --- canhoto (rodapé) — fonte preferida para codigo cliente, numero
    #     fatura, data vencimento e total a pagar, mais confiável no OCR
    #     que o cabeçalho (que tem logo/código de barras atrapalhando) ---
    m_canhoto = RE_CANHOTO.search(combined)
    if m_canhoto:
        if not data.codigo_cliente:
            data.codigo_cliente = re.sub(r"\s*-\s*", " - ", m_canhoto.group(1)).strip()
        data.numero_fatura = m_canhoto.group(2)
        data.data_vencimento = m_canhoto.group(3)
        data.total_a_pagar = br_to_float(m_canhoto.group(4))

    if not data.numero_fatura:
        m = RE_NUMERO_FATURA.search(combined)
        if m:
            data.numero_fatura = m.group(1) or m.group(2)

    if data.total_a_pagar is None:
        m = RE_CODIGO_CLIENTE_LINE.search(combined)
        if m:
            data.data_vencimento = data.data_vencimento or m.group(2)
            data.total_a_pagar = br_to_float(m.group(3))
        else:
            m = RE_DATA_VENCIMENTO_LABEL.search(combined)
            if m:
                data.data_vencimento = data.data_vencimento or m.group(1)
            m = RE_TOTAL_A_PAGAR_LABEL.search(combined)
            if m:
                data.total_a_pagar = br_to_float(m.group(1))

    m = RE_COD_DEBITO_AUTO.search(combined)
    if m:
        data.codigo_debito_automatico = m.group(1).strip()

    data.codigo_barras = _extract_barcode(boleto_text) or _extract_barcode(combined)

    return data


def enrich_boleto_from_other_docs(boleto: BoletoData, rps_text0: str, nf_text0: str) -> None:
    """Preenche campos ainda vazios do boleto usando a 1a página do RPS e/ou
    da NF como fallback — esses documentos costumam ter layout mais simples
    e OCR mais limpo que o boleto."""
    fallback_texts = [rps_text0, nf_text0]

    if not boleto.numero_fatura:
        m = _first_match([(RE_NUMERO_FATURA, t) for t in fallback_texts])
        if m:
            boleto.numero_fatura = m.group(1) or m.group(2)

    if not boleto.data_emissao:
        m = _first_match([(RE_DATA_EMISSAO, t) for t in fallback_texts])
        if m:
            boleto.data_emissao = m.group(1)

    if not boleto.codigo_cliente:
        m = _first_match([(RE_COD_CLIENTE_LABEL, t) for t in fallback_texts])
        if m:
            boleto.codigo_cliente = re.sub(r"\s*-\s*", " - ", m.group(1)).strip()

    if not boleto.cnpj_arcelor:
        m = _first_match([(RE_CNPJ_ARCELOR, t) for t in fallback_texts])
        if m:
            boleto.cnpj_arcelor = m.group(1)


# --------------------------------------------------------------------------
# Parsing do RPS (_RPS) -> lista de (aliquota, valor) por item de serviço
# --------------------------------------------------------------------------
# Cabeçalho tolerante a OCR: para na palavra ALIQUOTA e ignora o resto da
# linha (ex.: "VALOR DO ISS" pode virar "VALOR DOISS" ou coisa pior).
RPS_HEADER_RE = re.compile(
    r"DISCRIMINA[ÇCG][ÃA]O\s+DOS\s+SERVI[ÇCG]OS.*?AL[ÍI]QUOTA", re.S | re.I
)
RPS_ITEM_RE = re.compile(
    r"^(?P<desc>.+?)\s+"
    r"(?P<qtde>\d[\d.,]*)\s+"
    r"(?P<v_unit>\d[\d.,]*)\s+"
    r"(?P<v_total>\d[\d.,]*)\s+"
    r"(?P<aliquota>\d[\d.,]*)\s+"
    r"(?P<v_iss>\d[\d.,]*)\s*$"
)


def _extract_rps_items(text: str) -> List[Tuple[float, float]]:
    """Extrai (aliquota, valor_total) de todas as linhas de item de serviço
    reconhecidas no texto de uma página de RPS."""
    if not RPS_HEADER_RE.search(text):
        return []
    items: List[Tuple[float, float]] = []
    for line in text.splitlines():
        line = line.strip()
        if not line:
            continue
        m = RPS_ITEM_RE.match(line)
        if m:
            aliquota = br_to_float(m.group("aliquota"))
            valor = br_to_float(m.group("v_total"))
            if aliquota is not None and valor is not None:
                items.append((aliquota, valor))
    return items


def parse_rps(path: Path) -> Tuple[List[Tuple[float, float]], str]:
    """Retorna:
      - lista de (aliquota, valor_total) — um item por linha de serviço
        encontrada em cada página do RPS
      - texto da 1a página (para reaproveitar em enrich_boleto_from_other_docs)
    """
    results: List[Tuple[float, float]] = []
    first_page_text = ""
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            page_ok = False
            for candidate_text in page_text_candidates(page):
                if i == 0 and not first_page_text:
                    first_page_text = candidate_text
                items = _extract_rps_items(candidate_text)
                if items:
                    results.extend(items)
                    page_ok = True
                    break  # não tenta mais candidatos desta página (evita duplicar)
            if not page_ok:
                logger.warning(
                    "RPS pág. %s: nenhum item de serviço reconhecido em nenhuma tentativa.",
                    i + 1,
                )
    return results, first_page_text


# --------------------------------------------------------------------------
# Parsing da NF (_NF) -> lista de (aliquota, base_calculo) por página/nota
# --------------------------------------------------------------------------
# Também tolerante a OCR: para em ALIQUOTA e ignora o resto da linha de
# cabeçalho ("VALOR DO ICMS VALOR ISENTO VALOR OUTROS" pode sair corrompido).
NF_TOTALS_RE = re.compile(
    r"VALOR\s+TOTAL\s+BASE\s+DE\s+C[ÁA]LCULO\s+ICMS\s+AL[ÍI]QUOTA[^\n]*\n\s*"
    r"([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)",
    re.I,
)

# Campos opcionais de NFCom eletrônica (podem não existir neste layout)
NFCOM_NUMERO_RE = re.compile(r"N[úu]mero\s*NFCom[:\s]+([\w./-]+)", re.I)
NFCOM_SERIE_RE = re.compile(r"S[ée]rie\s*NFCom[:\s]+([\w./-]+)", re.I)
CHAVE_ACESSO_RE = re.compile(r"Chave de Acesso[:\s]+([\d\s]{20,})", re.I)
PROTOCOLO_RE = re.compile(r"Protocolo de Autoriza[çc][ãa]o[:\s]+([\w./-]+)", re.I)


def _extract_nf_total(text: str) -> Optional[Tuple[float, float]]:
    m = NF_TOTALS_RE.search(text)
    if not m:
        return None
    base_calculo = br_to_float(m.group(2))
    aliquota = br_to_float(m.group(3))
    if base_calculo is None or aliquota is None:
        return None
    return aliquota, base_calculo


def _extract_nfcom_extra(text: str, current: Dict[str, str]) -> None:
    if not current["Numero NFCom"]:
        m = NFCOM_NUMERO_RE.search(text)
        if m:
            current["Numero NFCom"] = m.group(1).strip()
    if not current["Serie NFCom"]:
        m = NFCOM_SERIE_RE.search(text)
        if m:
            current["Serie NFCom"] = m.group(1).strip()
    if not current["Chave de Acesso"]:
        m = CHAVE_ACESSO_RE.search(text)
        if m:
            current["Chave de Acesso"] = re.sub(r"\s+", "", m.group(1))
    if not current["Protocolo de Autorização"]:
        m = PROTOCOLO_RE.search(text)
        if m:
            current["Protocolo de Autorização"] = m.group(1).strip()


def parse_nf(path: Path) -> Tuple[List[Tuple[float, float]], Dict[str, str], str]:
    """Retorna:
      - lista de (aliquota, base_calculo_icms) — uma por página/nota
      - dict com campos de NFCom eletrônica, se encontrados em qualquer
        página (Numero NFCom, Serie NFCom, Chave de Acesso, Protocolo)
      - texto da 1a página (para reaproveitar em enrich_boleto_from_other_docs)
    """
    results: List[Tuple[float, float]] = []
    nfcom_extra = {
        "Numero NFCom": "",
        "Serie NFCom": "",
        "Chave de Acesso": "",
        "Protocolo de Autorização": "",
    }
    first_page_text = ""
    with pdfplumber.open(path) as pdf:
        for i, page in enumerate(pdf.pages):
            page_ok = False
            for candidate_text in page_text_candidates(page):
                if i == 0 and not first_page_text:
                    first_page_text = candidate_text
                found = _extract_nf_total(candidate_text)
                if found:
                    results.append(found)
                    _extract_nfcom_extra(candidate_text, nfcom_extra)
                    page_ok = True
                    break  # não tenta mais candidatos desta página (evita duplicar)
            if not page_ok:
                logger.warning(
                    "NF pág. %s: linha de totais ICMS não encontrada em nenhuma tentativa.",
                    i + 1,
                )
    return results, nfcom_extra, first_page_text


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
    avisos: List[str] = field(default_factory=list)


def process_trio(key: str, files: Dict[str, Path]) -> TrioResult:
    result = TrioResult(key=key)

    faltando = [k for k in ("PDF", "RPS", "NF") if k not in files]
    if faltando:
        result.avisos.append(
            f"Trio '{key}': arquivo(s) ausente(s) -> {', '.join(faltando)}. "
            "Alguns campos/valores podem não ser gerados."
        )

    filename_hint = (files.get("PDF") or files.get("NF") or files.get("RPS")).name

    boleto = (
        parse_boleto(files["PDF"], filename_hint=filename_hint)
        if "PDF" in files
        else BoletoData(
            cnpj_arcelor=cnpj_arcelor_from_filename(filename_hint),
            codigo_cliente=codigo_cliente_from_filename(filename_hint),
            cnpj_fornecedor=CNPJ_FORNECEDOR_PADRAO,
        )
    )
    result.boleto = boleto
    result.total_boleto = boleto.total_a_pagar

    itens_aliquota: Dict[float, float] = {}
    rps_first_page_text = ""
    nf_first_page_text = ""

    if "RPS" in files:
        rps_items, rps_first_page_text = parse_rps(files["RPS"])
        for aliquota, valor in rps_items:
            a = norm_aliquota(aliquota)
            itens_aliquota[a] = itens_aliquota.get(a, 0.0) + valor

    nfcom_extra = {
        "Numero NFCom": "",
        "Serie NFCom": "",
        "Chave de Acesso": "",
        "Protocolo de Autorização": "",
    }
    if "NF" in files:
        nf_items, nfcom_extra, nf_first_page_text = parse_nf(files["NF"])
        for aliquota, valor in nf_items:
            a = norm_aliquota(aliquota)
            itens_aliquota[a] = itens_aliquota.get(a, 0.0) + valor

    # completa campos do boleto que não puderam ser lidos, usando RPS/NF
    enrich_boleto_from_other_docs(boleto, rps_first_page_text, nf_first_page_text)

    codigo_fatura = codigo_fatura_from_filename(filename_hint)

    texto_boleto = (
        f"Valor Total do Boleto: {boleto.total_a_pagar:,.2f}".replace(",", "X")
        .replace(".", ",")
        .replace("X", ".")
        if boleto.total_a_pagar is not None
        else ""
    )

    if not itens_aliquota:
        result.avisos.append(
            f"Trio '{key}': nenhum valor/alíquota foi extraído de RPS ou NF. "
            "Verifique se os PDFs estão corretos ou se o OCR conseguiu ler as páginas."
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

    result.total_agregado = round(sum(itens_aliquota.values()), 2) if itens_aliquota else 0.0

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
        exibir logs na interface (Streamlit)
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
