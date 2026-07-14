# Extrator de Faturas EMBRATEL/Claro — ArcelorMittal

Lê trios de PDF (boleto `_PDF`, recibos `_RPS`, notas fiscais `_NF`) e gera
uma planilha Excel consolidada, uma linha por alíquota de imposto (ICMS das
NFs + ISS dos RPS somados juntos, por trio de arquivos).

## Arquivos

- **`embratel_extractor.py`** — módulo central com toda a lógica de
  extração/agrupamento. É importado tanto pelo script de linha de comando
  quanto pelo app Streamlit.
- **`gerar_planilha.py`** — script de linha de comando, útil para testar
  localmente sem subir o Streamlit.
- **`app.py`** — aplicação Streamlit (upload em lote de vários trios,
  processamento e download da planilha final).
- **`requirements.txt`** — dependências.

## Como funciona a extração

1. **Identificação dos trios**: os arquivos são agrupados pelo nome,
   removendo o sufixo `_PDF.pdf` / `_RPS.pdf` / `_NF.pdf`. Todo o resto do
   nome precisa ser idêntico entre os três arquivos de um mesmo trio.
2. **Boleto (`_PDF`)**: 1 página. Extrai os dados cadastrais/constantes de
   cada linha da planilha (CNPJ Arcelor, CNPJ Fornecedor, datas, código do
   cliente, número da fatura, texto do total, código de barras, etc.).
3. **RPS (`_RPS`)**: 1 página por filial. Para cada item de serviço lido
   (ex.: `GRP - GERENCIA DE REDE`), captura `Valor Total` e `Alíquota (%)`
   do ISS.
4. **NF (`_NF`)**: 1 página por nota fiscal (filial/UF). Lê a linha de
   totais `VALOR TOTAL | BASE DE CÁLCULO ICMS | ALÍQUOTA | VALOR DO ICMS |
   VALOR ISENTO | VALOR OUTROS` e captura a **base de cálculo** e a
   **alíquota** do ICMS.
5. **Agregação**: os valores de RPS e NF do mesmo trio são somados juntos,
   agrupados pela alíquota (coluna `Tipo`). Cada alíquota distinta vira uma
   linha na planilha final, com `Valor` = soma de todas as bases de
   cálculo/valores de serviço daquela alíquota.
6. **Validação automática**: para cada trio, o total agregado é comparado
   com o "Total a Pagar" do boleto. Se a diferença for maior que R$ 0,05,
   um aviso é exibido (útil para pegar erro de extração num lote grande).

A coluna **`Numero NF`** fica sempre em branco, conforme definido.

## Uso via linha de comando

```bash
pip install -r requirements.txt
python gerar_planilha.py /caminho/para/pasta/com/pdfs saida.xlsx
```

A pasta pode conter **vários trios diferentes** misturados — todos são
processados e consolidados na mesma planilha de saída.

## Uso via Streamlit

```bash
pip install -r requirements.txt
streamlit run app.py
```

Na interface: envie todos os PDFs de uma vez (pode ser vários trios de
meses/clientes diferentes juntos), clique em **Processar e gerar
planilha** e baixe o Excel consolidado. A tela mostra, por trio, o total
do boleto, o total agregado e eventuais avisos de divergência antes de
você baixar o arquivo.

## Sobre OCR

Os PDFs testados têm texto real embutido (não são imagens escaneadas), então
a extração usa `pdfplumber` diretamente — mais rápido e confiável que OCR.
Caso apareça algum PDF escaneado num lote futuro, o código detecta
automaticamente a página sem texto e cai para OCR (`pytesseract`) como
fallback — só é necessário instalar o Tesseract no sistema
(`apt install tesseract-ocr tesseract-ocr-por`) para esse caso.

## Campos que podem ficar em branco

`Numero NFCom`, `Serie NFCom`, `Chave de Acesso` e `Protocolo de
Autorização` são específicos do modelo de Nota Fiscal de Consumidor
eletrônica (NFCom). No layout atual da Claro/Embratel (regime especial)
esses campos não aparecem, então ficam em branco — mas a extração já está
pronta para capturá-los automaticamente caso apareçam em notas de um
layout diferente.
