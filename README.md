# Conversor REGATI/PEFIN – TXT para XLSX

Converte arquivos TXT do sistema REGATI/PEFIN em planilhas XLSX com 5 colunas formatadas:

- **Documento principal (CPF/CNPJ)** – máscara 000.000.000-00 / 00.000.000/0000-00  
- **Valor (R$)** – centavos convertidos para reais (R$ #.##0,00)  
- **Data inclusão** – DDMMAAAA → dd/mm/aaaa  
- **Hora inclusão** – HHMM/HHMMSS → texto HH:MM ou HH:MM:SS  
- **Usuário inclusão** – texto  

## Como usar

### Opção 1 – Interface gráfica (recomendado)
1. Dê duplo clique em **`Converter REGATI para XLSX.bat`** ou execute: `python txt_para_xlsx.py`
2. Na janela que abrir, clique em **"Selecionar arquivo TXT"** e escolha o arquivo que deseja converter.
3. Clique em **"Converter para XLSX"**.
4. A planilha será criada na **mesma pasta do TXT**, com nome: **nome_do_arquivo_AAAAMMDD_HHMMSS.xlsx** (não sobrescreve arquivos anteriores).

### Opção 2 – Arrastar e soltar
1. Arraste o arquivo `.TXT` sobre o arquivo **`Converter REGATI para XLSX.bat`**.
2. A planilha será criada na mesma pasta do TXT, com nome incluindo data e hora da conversão.

### Opção 3 – Linha de comando
```bash
# Abre a interface para escolher o TXT
python txt_para_xlsx.py

# Converte um arquivo específico (gera nome com data/hora)
python txt_para_xlsx.py "C:\pasta\arquivo.TXT"

# Converte e define nome da planilha de saída
python txt_para_xlsx.py "arquivo.TXT" "saida.xlsx"
```

## Nome da planilha gerada
Cada conversão gera uma **nova** planilha, sem sobrescrever as anteriores:
- **Nome:** `{nome_do_txt}_{AAAAMMDD}_{HHMMSS}.xlsx`  
- **Exemplo:** `R.008.M3287.PEFIN.REGATI.D260202.H014750_20260203_143052.xlsx`

## Requisitos
- Python 3.7+
- Pacote: `xlsxwriter` (`pip install xlsxwriter`)

**Dica:** Se der erro ao salvar (“Permission denied”), feche a planilha XLSX no Excel e execute a conversão novamente.

## Automação para outros arquivos
A rotina detecta as colunas **pelo nome do cabeçalho** do TXT (`DOC PRINCIPAL 1`, `VALOR`, `DT INCLUSAO`, `HR INCLUSAO`, `USER INCLUSAO`), então serve para qualquer arquivo TXT REGATI/PEFIN que tenha o mesmo cabeçalho, mesmo que a ordem das colunas seja diferente.
