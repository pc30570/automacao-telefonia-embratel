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
2. **CNPJ Arcelor e Código do Cliente**: extraídos preferencialmente do
   **nome do arquivo** (padrão `ARCELORMITTAL_<CNPJ 14 dígitos>_EMBRATEL_
   <código cliente>-DADOS_`), já que esse padrão é mais confiável que OCR
   em PDFs escaneados. Se o nome não seguir o padrão, cai para leitura do
   texto do boleto.
3. **Boleto (`_PDF`)**: extrai os demais dados cadastrais/constantes de
   cada linha da planilha (datas, número da fatura, texto do total,
   código de barras, etc.). Prioriza a linha do **canhoto** no rodapé
   (`CÓDIGO DA CONTA | NÚMERO DA FATURA | DATA DE VENCIMENTO | VALOR DA
   CONTA`), que é lida de forma bem mais confiável pelo OCR do que o
   cabeçalho (que tem logotipo/código de barras atrapalhando).
4. **RPS (`_RPS`)**: 1 página por filial. Para cada item de serviço lido
   (ex.: `GRP - GERENCIA DE REDE`), captura `Valor Total` e `Alíquota (%)`
   do ISS.
5. **NF (`_NF`)**: 1 página por nota fiscal (filial/UF). Lê a linha de
   totais `VALOR TOTAL | BASE DE CÁLCULO ICMS | ALÍQUOTA | VALOR DO ICMS |
   VALOR ISENTO | VALOR OUTROS` e captura a **base de cálculo** e a
   **alíquota** do ICMS.
6. **Agregação**: os valores de RPS e NF do mesmo trio são somados juntos,
   agrupados pela alíquota (coluna `Tipo`). Cada alíquota distinta vira uma
   linha na planilha final, com `Valor` = soma de todas as bases de
   cálculo/valores de serviço daquela alíquota.
7. **Sem validação bloqueante de totais**: o total do boleto e o total
   agregado (RPS + NF) são mostrados lado a lado só como informação — não
   há aviso de erro se forem diferentes, já que o boleto pode incluir
   juros e multa que não aparecem em RPS/NF.

A coluna **`Numero NF`** fica sempre em branco, conforme definido.

## OCR (PDFs escaneados)

Nem todo lote vem com texto embutido no PDF — alguns são digitalizados
(imagem). Para esses casos, o código:

1. Detecta automaticamente que a página não tem texto extraível.
2. Tenta OCR com `pytesseract`, testando **várias combinações** de idioma
   (português, se disponível; senão inglês) e modo de leitura de página
   (`--psm 6`, melhor para tabelas; `--psm 3`, automático), na ordem mais
   confiável primeiro.
3. Usa o **primeiro resultado que conseguir reconhecer os dados
   esperados** (ex.: linha de item do RPS, linha de totais da NF) e para
   por aí — nunca soma valores de mais de uma tentativa de OCR da mesma
   página, para não contar a mesma nota em duplicidade.
4. Se nenhuma tentativa funcionar, a página é ignorada e um aviso é
   registrado (aparece no Streamlit e no log do CLI), mas o restante do
   lote continua sendo processado normalmente.

Os PDFs escaneados testados aqui usaram OCR em inglês porque o pacote de
idioma português (`tesseract-ocr-por`) não estava instalado no ambiente de
teste — os regex já foram ajustados para tolerar as trocas de caractere
mais comuns do OCR em português lido como inglês (ex.: `Ç` → `G`,
`Á/Í` → `A/I`). **Para melhor precisão em produção, instale
`tesseract-ocr-por`** (`apt install tesseract-ocr-por` no Ubuntu/Debian);
o código detecta e usa automaticamente se disponível.

OCR é mais lento que leitura de texto embutido — em lotes grandes com
muitas páginas escaneadas, o processamento pode demorar bastante mais.

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
do boleto e o total agregado lado a lado, além de eventuais avisos (ex.:
arquivo faltando no trio, ou nenhuma página reconhecida) antes de você
baixar o arquivo.

## Campos que podem ficar em branco

`Numero NFCom`, `Serie NFCom`, `Chave de Acesso` e `Protocolo de
Autorização` são específicos do modelo de Nota Fiscal de Consumidor
eletrônica (NFCom). No layout atual da Claro/Embratel (regime especial)
esses campos não aparecem, então ficam em branco — mas a extração já está
pronta para capturá-los automaticamente caso apareçam em notas de um
layout diferente.
