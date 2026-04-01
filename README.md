# Conversor REGATI / PEFIN — TXT → XLSX

Converte arquivos **texto delimitados** (formato típico **REGATI / PEFIN**, separador `;`) em planilhas **Excel (.xlsx)** com colunas padronizadas, formatação de moeda, datas e documentos (CPF/CNPJ), para uso em análises e integrações no projeto **CCOBI – SERASA**.

---

## O que o conversor faz

1. Lê a **primeira linha** do TXT como **cabeçalho** e localiza colunas **pelo nome** (a ordem das colunas no arquivo pode variar).  
2. Para cada linha de dados, extrai campos e grava uma planilha com **6 colunas** na ordem abaixo.  
3. Aplica formatação adequada no Excel (datas como data, valores em reais, CPF/CNPJ como texto com máscara quando os dígitos verificadores são válidos).

### Colunas geradas na planilha

| # | Coluna no XLSX | Origem no TXT (nome do cabeçalho) |
|---|----------------|-----------------------------------|
| 1 | **Número do auto** | `CONTRATO` |
| 2 | **Data de inscrição** | `DT INCLUSAO` (DDMMAAAA → data Excel) |
| 3 | **Data da ocorrência** | `DT OCORRENCIA` |
| 4 | **Data exclusão** | `DT EXCLUSAO` |
| 5 | **Valor inscrito** | `VALOR` (interpretado como **centavos** → convertido para **reais**) |
| 6 | **CPF/CNPJ** | `DOC PRINCIPAL 1` (formatação com validação de dígitos verificadores) |

Colunas adicionais detectadas no cabeçalho (ex.: `HR INCLUSAO`, `USER INCLUSAO`) são usadas internamente quando necessário; a planilha principal de saída segue as 6 colunas acima para alinhar com o fluxo de trabalho combinado.

### Cabeçalhos obrigatórios no TXT

Para a conversão padrão, o arquivo deve conter pelo menos:

- `DOC PRINCIPAL 1`  
- `VALOR`  
- `CONTRATO`  
- `DT INCLUSAO`  
- `DT OCORRENCIA`  
- `DT EXCLUSAO`  

Encoding recomendado: **UTF-8** (o leitor usa `utf-8` com substituição de caracteres inválidos).

---

## Requisitos

- **Python** 3.7+  
- **xlsxwriter** (gravação eficiente de XLSX, modo `constant_memory` para arquivos grandes)

```bash
pip install -r requirements.txt
```

---

## Como usar

### 1) Interface gráfica (recomendado)

```bash
python txt_para_xlsx.py
```

Ou dê duplo clique em **`Converter REGATI para XLSX.bat`**.

- Clique em **Selecionar arquivo TXT** e escolha o arquivo.  
- Clique em **Converter para XLSX**.  
- O arquivo gerado fica na **mesma pasta do TXT**, com nome:

  `{nome_do_txt}_{AAAAMMDD}_{HHMMSS}.xlsx`

Assim **não sobrescreve** conversões anteriores.

### 2) Arrastar e soltar (Windows)

Arraste o arquivo `.TXT` sobre o `.bat` — a conversão roda com o TXT indicado.

### 3) Linha de comando

```bash
# Abre a GUI
python txt_para_xlsx.py

# Converte um arquivo (nome de saída automático com data/hora)
python txt_para_xlsx.py "C:\pasta\arquivo.TXT"

# Converte e define o XLSX de saída
python txt_para_xlsx.py "C:\pasta\arquivo.TXT" "C:\pasta\saida.xlsx"
```

---

## Detalhes de formatação

### CPF e CNPJ

A função `formatar_cpf_cnpj` usa os **algoritmos oficiais de dígitos verificadores** (módulo 11) para decidir se o número é CPF (11 dígitos) ou CNPJ (14 dígitos) antes de aplicar máscara. Isso reduz erros de classificação quando há zeros à esquerda ou strings ambíguas.

### Valores

O campo `VALOR` é tratado como **inteiro em centavos** (somente dígitos); o Excel recebe o valor em **reais** com formato `R$ #.##0,00`.

### Datas

Datas no padrão **DDMMAAAA** (8 dígitos) são convertidas para tipo data no Excel. Se o parse falhar, o texto original é preservado.

### Horários

Funções auxiliares formatam `HHMM` / `HHMMSS` como texto `HH:MM` ou `HH:MM:SS` quando aplicável (útil para colunas de inclusão no cabeçalho).

---

## Desempenho e arquivos grandes

O uso de `constant_memory` no **XlsxWriter** e processamento linha a linha permite lidar com **dezenas de milhares de linhas** sem carregar todo o TXT na memória de uma só vez. A interface pode registrar progresso a cada 50.000 linhas.

---

## Solução de problemas

| Problema | Ação |
|----------|------|
| `Permission denied` ao salvar | Feche o `.xlsx` no Excel e tente de novo. |
| Coluna não encontrada | Confira se o TXT tem o cabeçalho com os nomes exatos (incluindo espaços). |
| Valores ou datas estranhos | Verifique se o arquivo não está corrompido ou com separador diferente de `;`. |

---

## Estrutura do projeto

| Arquivo | Função |
|---------|--------|
| `txt_para_xlsx.py` | Lógica + GUI Tkinter |
| `Converter REGATI para XLSX.bat` | Execução rápida no Windows |
| `requirements.txt` | Dependência `xlsxwriter` |

---

## Contexto

Ferramenta de apoio ao trabalho com bases **REGATI/PEFIN** no âmbito **CCOBI – SERASA**.
