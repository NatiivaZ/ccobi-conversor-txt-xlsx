# -*- coding: utf-8 -*-
"""
Automação: converte arquivos TXT REGATI/PEFIN em XLSX com colunas selecionadas e formatadas.
Planilha: 6 colunas — Número do auto, Data de inscrição, Data da ocorrência, Data exclusão, Valor inscrito, CPF/CNPJ.

Uso:
  python txt_para_xlsx.py                          -> abre a interface para escolher o TXT e converter
  python txt_para_xlsx.py "C:\pasta\arquivo.TXT"   -> converte o arquivo informado (linha de comando)
"""

import re
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import Tk, Label, Frame, filedialog, messagebox, DISABLED, NORMAL, X, BOTH, RIGHT
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk

import xlsxwriter

# Nomes das colunas no cabeçalho do TXT (para detectar índices)
NOME_DOC_PRINCIPAL = "DOC PRINCIPAL 1"
NOME_VALOR = "VALOR"
NOME_CONTRATO = "CONTRATO"
NOME_DT_INCLUSAO = "DT INCLUSAO"
NOME_DT_OCORRENCIA = "DT OCORRENCIA"
NOME_DT_EXCLUSAO = "DT EXCLUSAO"
NOME_HR_INCLUSAO = "HR INCLUSAO"
NOME_USER_INCLUSAO = "USER INCLUSAO"


def achar_indices_colunas(header: str, sep: str = ";"):
    """Obtém os índices das colunas pelo nome do cabeçalho (evita troca de colunas)."""
    colunas = [c.strip() for c in header.split(sep)]
    indices = {}
    for i, nome in enumerate(colunas):
        if nome == NOME_DOC_PRINCIPAL:
            indices["doc"] = i
        elif nome == NOME_VALOR:
            indices["valor"] = i
        elif nome == NOME_CONTRATO:
            indices["contrato"] = i
        elif nome == NOME_DT_INCLUSAO:
            indices["dt_inclusao"] = i
        elif nome == NOME_DT_OCORRENCIA:
            indices["dt_ocorrencia"] = i
        elif nome == NOME_DT_EXCLUSAO:
            indices["dt_exclusao"] = i
        elif nome == NOME_HR_INCLUSAO:
            indices["hr_inclusao"] = i
        elif nome == NOME_USER_INCLUSAO:
            indices["user_inclusao"] = i
    return indices, colunas


def _validar_cpf(digits: str) -> bool:
    """
    Valida CPF pelos dígitos verificadores (algoritmo módulo 11 - Receita Federal).
    digits: string com exatamente 11 dígitos.
    """
    if len(digits) != 11 or not digits.isdigit():
        return False
    if digits == digits[0] * 11:  # rejeita 111.111.111-11 etc.
        return False
    # Primeiro dígito verificador: pesos 10,9,8,7,6,5,4,3,2 nos 9 primeiros
    s = sum(int(digits[i]) * (10 - i) for i in range(9))
    d1 = 0 if (s % 11) < 2 else 11 - (s % 11)
    if int(digits[9]) != d1:
        return False
    # Segundo dígito: pesos 11,10,...,2 nos 10 primeiros
    s2 = sum(int(digits[i]) * (11 - i) for i in range(10))
    d2 = 0 if (s2 % 11) < 2 else 11 - (s2 % 11)
    return int(digits[10]) == d2


def _validar_cnpj(digits: str) -> bool:
    """
    Valida CNPJ pelos dígitos verificadores (algoritmo módulo 11 - Receita Federal).
    digits: string com exatamente 14 dígitos.
    """
    if len(digits) != 14 or not digits.isdigit():
        return False
    # Pesos para o primeiro dígito (12 primeiros): 5,4,3,2,9,8,7,6,5,4,3,2
    pesos1 = (5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    s1 = sum(int(digits[i]) * pesos1[i] for i in range(12))
    d1 = 0 if (s1 % 11) < 2 else 11 - (s1 % 11)
    if int(digits[12]) != d1:
        return False
    # Pesos para o segundo (13 primeiros): 6,5,4,3,2,9,8,7,6,5,4,3,2
    pesos2 = (6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    s2 = sum(int(digits[i]) * pesos2[i] for i in range(13))
    d2 = 0 if (s2 % 11) < 2 else 11 - (s2 % 11)
    return int(digits[13]) == d2


def formatar_cpf_cnpj(val: str) -> str:
    """
    Formata como CPF ou CNPJ conforme validação pelos dígitos verificadores (Receita Federal).
    Só formata como CPF se os 11 dígitos (ou os 11 últimos, quando há zeros à esquerda) forem
    válidos; só formata como CNPJ se os 14 dígitos forem válidos. Evita trocar CPF por CNPJ
    e vice-versa. Sem API e sem internet.
    """
    if not val:
        return ""
    digits = re.sub(r"\D", "", val.strip())
    if len(digits) == 11:
        if _validar_cpf(digits):
            return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"
        return val.strip()
    if len(digits) == 12:
        d11 = digits[-11:]
        if _validar_cpf(d11):
            return f"{d11[:3]}.{d11[3:6]}.{d11[6:9]}-{d11[9:]}"
        return val.strip()
    if len(digits) == 13:
        d11 = digits[-11:]
        if _validar_cpf(d11):
            return f"{d11[:3]}.{d11[3:6]}.{d11[6:9]}-{d11[9:]}"
        return val.strip()
    if len(digits) >= 14:
        d14 = digits[-14:] if len(digits) > 14 else digits
        if _validar_cnpj(d14):
            return f"{d14[:2]}.{d14[2:5]}.{d14[5:8]}/{d14[8:12]}-{d14[12:]}"
        # 14 dígitos mas CNPJ inválido: pode ser 00+CPF
        d11 = d14[-11:]
        if _validar_cpf(d11):
            return f"{d11[:3]}.{d11[3:6]}.{d11[6:9]}-{d11[9:]}"
        return val.strip()
    return val.strip()


def parse_data_ddmmaaaa(val: str):
    """Converte DDMMAAAA para objeto date. Retorna None se inválido."""
    if not val or len(val.strip()) < 8:
        return None
    s = re.sub(r"\D", "", val.strip())[:8]
    if len(s) != 8:
        return None
    try:
        d, m, a = int(s[:2]), int(s[2:4]), int(s[4:8])
        if 1 <= d <= 31 and 1 <= m <= 12 and 1900 <= a <= 2100:
            return datetime(a, m, d).date()
    except ValueError:
        pass
    return None


def formatar_hora_como_texto(val: str) -> str:
    """
    Converte HHMM ou HHMMSS do TXT em texto HH:MM ou HH:MM:SS para o Excel.
    Gravando como texto evita ######## e garante exibição correta.
    """
    if not val:
        return ""
    s = re.sub(r"\D", "", val.strip())
    if len(s) == 4:
        try:
            h, m = int(s[:2]), int(s[2:4])
            if 0 <= h <= 23 and 0 <= m <= 59:
                return f"{h:02d}:{m:02d}"
        except ValueError:
            pass
    if len(s) >= 6:
        try:
            h, m, sec = int(s[:2]), int(s[2:4]), int(s[4:6])
            if 0 <= h <= 23 and 0 <= m <= 59 and 0 <= sec <= 59:
                return f"{h:02d}:{m:02d}:{sec:02d}"
        except ValueError:
            pass
    return val.strip()


def valor_centavos_para_reais(val: str):
    """Converte string em centavos para float em reais. Retorna None se inválido."""
    if not val:
        return None
    s = re.sub(r"\D", "", val.strip())
    if not s:
        return None
    try:
        return int(s) / 100.0
    except ValueError:
        return None


def _nome_planilha_com_data_hora(arquivo_txt: Path) -> Path:
    """Gera nome da planilha: nome do TXT + data e hora da conversão (não sobrescreve)."""
    agora = datetime.now()
    nome = f"{arquivo_txt.stem}_{agora:%Y%m%d}_{agora:%H%M%S}.xlsx"
    return arquivo_txt.parent / nome


def converter_txt_para_xlsx(arquivo_txt: Path, arquivo_xlsx: Path = None, log_callback=None) -> Path:
    """
    Converte um TXT REGATI/PEFIN em XLSX com 6 colunas formatadas.
    Se arquivo_xlsx não for informado, gera novo arquivo: nome_do_txt_AAAAMMDD_HHMMSS.xlsx.
    log_callback(msg): opcional, chamado para exibir mensagens (ex.: na interface).
    """
    arquivo_txt = Path(arquivo_txt).resolve()
    if not arquivo_txt.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {arquivo_txt}")
    if arquivo_xlsx is None:
        arquivo_xlsx = _nome_planilha_com_data_hora(arquivo_txt)
    else:
        arquivo_xlsx = Path(arquivo_xlsx).resolve()

    def _log(msg: str):
        if log_callback:
            log_callback(msg)
        else:
            print(msg)

    _log(f"Arquivo TXT: {arquivo_txt}")
    with open(arquivo_txt, "r", encoding="utf-8", errors="replace") as f:
        header = f.readline()
    indices, _ = achar_indices_colunas(header)
    for k, nome in [
        ("doc", "DOC PRINCIPAL 1"),
        ("valor", "VALOR"),
        ("contrato", "CONTRATO"),
        ("dt_inclusao", "DT INCLUSAO"),
        ("dt_ocorrencia", "DT OCORRENCIA"),
        ("dt_exclusao", "DT EXCLUSAO"),
    ]:
        if k not in indices:
            raise ValueError(f"Coluna '{nome}' não encontrada no cabeçalho do TXT.")

    idx_doc = indices["doc"]
    idx_valor = indices["valor"]
    idx_contrato = indices["contrato"]
    idx_dt = indices["dt_inclusao"]
    idx_dt_ocorrencia = indices["dt_ocorrencia"]
    idx_dt_exclusao = indices["dt_exclusao"]

    _log("Criando planilha XLSX...")
    workbook = xlsxwriter.Workbook(str(arquivo_xlsx), {"constant_memory": True})
    sheet = workbook.add_worksheet("REGATI")

    fmt_header = workbook.add_format({"bold": True, "align": "center", "valign": "vcenter", "text_wrap": True})
    fmt_moeda = workbook.add_format({"num_format": '"R$" #,##0.00', "align": "right"})
    fmt_data = workbook.add_format({"num_format": "dd/mm/yyyy", "align": "center"})
    fmt_texto = workbook.add_format({"align": "left"})

    # Ordem e nomes na planilha: Número do auto → Data de inscrição → Data da ocorrência → Data exclusão → Valor inscrito → CPF/CNPJ
    headers_xlsx = [
        "Número do auto",
        "Data de inscrição",
        "Data da ocorrência",
        "Data exclusão",
        "Valor inscrito",
        "CPF/CNPJ",
    ]
    for col, title in enumerate(headers_xlsx):
        sheet.write(0, col, title, fmt_header)

    sheet.set_column(0, 0, 24)   # Número do auto (Contrato)
    sheet.set_column(1, 1, 14)   # Data de inscrição
    sheet.set_column(2, 2, 14)   # Data da ocorrência
    sheet.set_column(3, 3, 14)   # Data exclusão
    sheet.set_column(4, 4, 20)   # Valor inscrito (evitar ####)
    sheet.set_column(5, 5, 24)   # CPF/CNPJ (máscara)

    row_excel = 1
    with open(arquivo_txt, "r", encoding="utf-8", errors="replace") as f:
        next(f)
        for num_linha, line in enumerate(f, start=2):
            if not line.strip():
                continue
            parts = line.split(";")
            max_idx = max(idx_doc, idx_valor, idx_contrato, idx_dt, idx_dt_ocorrencia, idx_dt_exclusao)
            if len(parts) <= max_idx:
                continue
            doc = parts[idx_doc].strip() if idx_doc < len(parts) else ""
            val = parts[idx_valor].strip() if idx_valor < len(parts) else ""
            contrato = parts[idx_contrato].strip() if idx_contrato < len(parts) else ""
            dt_inc = parts[idx_dt].strip() if idx_dt < len(parts) else ""
            dt_ocor = parts[idx_dt_ocorrencia].strip() if idx_dt_ocorrencia < len(parts) else ""
            dt_excl = parts[idx_dt_exclusao].strip() if idx_dt_exclusao < len(parts) else ""

            # 1. Número do auto (Contrato)
            sheet.write_string(row_excel, 0, contrato, fmt_texto)

            # 2. Data de inscrição (DT INCLUSAO)
            dt = parse_data_ddmmaaaa(dt_inc)
            if dt:
                sheet.write_datetime(row_excel, 1, datetime.combine(dt, datetime.min.time()), fmt_data)
            else:
                sheet.write_string(row_excel, 1, dt_inc or "", fmt_texto)

            # 3. Data da ocorrência (DT OCORRENCIA)
            dt_oc = parse_data_ddmmaaaa(dt_ocor)
            if dt_oc:
                sheet.write_datetime(row_excel, 2, datetime.combine(dt_oc, datetime.min.time()), fmt_data)
            else:
                sheet.write_string(row_excel, 2, dt_ocor or "", fmt_texto)

            # 4. Data exclusão (DT EXCLUSAO) — formato data (dd/mm/aaaa)
            dt_ex = parse_data_ddmmaaaa(dt_excl)
            if dt_ex:
                sheet.write_datetime(row_excel, 3, datetime.combine(dt_ex, datetime.min.time()), fmt_data)
            else:
                sheet.write_string(row_excel, 3, dt_excl or "", fmt_texto)

            # 5. Valor inscrito (centavos -> reais)
            v = valor_centavos_para_reais(val)
            if v is not None:
                sheet.write_number(row_excel, 4, v, fmt_moeda)
            else:
                sheet.write_string(row_excel, 4, val or "", fmt_texto)

            # 6. CPF/CNPJ (Documento principal)
            sheet.write_string(row_excel, 5, formatar_cpf_cnpj(doc), fmt_texto)

            row_excel += 1
            if num_linha % 50000 == 0:
                _log(f"  Processadas {num_linha - 1} linhas...")

    try:
        workbook.close()
    except xlsxwriter.exceptions.FileCreateError as e:
        if "Permission denied" in str(e) or "13" in str(e):
            raise OSError(
                f"Não foi possível salvar em {arquivo_xlsx}. "
                "Feche a planilha no Excel (se estiver aberta) e execute a conversão novamente."
            ) from e
        raise
    _log(f"Planilha salva em: {arquivo_xlsx}")
    _log(f"Total de linhas de dados: {row_excel - 1}")
    return arquivo_xlsx


# Cores e estilo da interface (tema claro profissional)
_COR_FUNDO = "#f1f5f9"
_COR_CARD = "#ffffff"
_COR_TITULO = "#0f172a"
_COR_SUBTITULO = "#64748b"
_COR_TEXTO = "#334155"
_COR_BORDA = "#e2e8f0"
_COR_PRIMARIA = "#0369a1"
_COR_PRIMARIA_HOVER = "#0284c7"
_COR_SUCESSO = "#059669"


def _rodar_interface():
    """Abre a interface gráfica para escolher o TXT e converter."""
    try:
        root = Tk()
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível abrir a interface:\n{e}")
        return

    root.title("Conversor REGATI/PEFIN — TXT para XLSX")
    root.minsize(560, 520)
    root.resizable(True, True)
    root.configure(bg=_COR_FUNDO)

    # Estilos ttk
    style = ttk.Style()
    style.configure(
        "Primario.TButton",
        font=("Segoe UI", 10, "bold"),
        padding=(16, 8),
    )
    style.configure(
        "Secundario.TButton",
        font=("Segoe UI", 10),
        padding=(14, 6),
    )
    style.configure(
        "Card.TFrame",
        background=_COR_CARD,
    )

    pasta_inicial = Path(__file__).resolve().parent
    arquivo_selecionado = [None]

    # ---- Cabeçalho (altura fixa para o subtítulo não ser cortado) ----
    frm_header = Frame(root, bg=_COR_CARD, highlightthickness=0)
    frm_header.pack(fill=X, side="top")
    lbl_titulo = Label(
        frm_header,
        text="Conversor REGATI/PEFIN",
        font=("Segoe UI", 18, "bold"),
        fg=_COR_TITULO,
        bg=_COR_CARD,
    )
    lbl_titulo.pack(anchor="w", padx=24, pady=(20, 4))

    lbl_sub = Label(
        frm_header,
        text="TXT → XLSX com formatação (CPF/CNPJ, datas, valores)",
        font=("Segoe UI", 10),
        fg=_COR_SUBTITULO,
        bg=_COR_CARD,
    )
    lbl_sub.pack(anchor="w", padx=24, pady=(0, 20))

    # ---- Card: Arquivo ----
    frm_arquivo = Frame(root, bg=_COR_CARD, padx=20, pady=16, highlightbackground=_COR_BORDA, highlightthickness=1)
    frm_arquivo.pack(fill=X, padx=20, pady=(8, 10))

    Label(
        frm_arquivo,
        text="Arquivo TXT",
        font=("Segoe UI", 11, "bold"),
        fg=_COR_TITULO,
        bg=_COR_CARD,
    ).pack(anchor="w")

    lbl_arquivo = Label(
        frm_arquivo,
        text="Nenhum arquivo selecionado.",
        font=("Segoe UI", 9),
        fg=_COR_SUBTITULO,
        bg=_COR_CARD,
        anchor="w",
    )
    lbl_arquivo.pack(fill=X, pady=(4, 10))

    btn_sel = ttk.Button(
        frm_arquivo,
        text="Selecionar arquivo TXT",
        command=lambda: None,
        style="Secundario.TButton",
    )
    btn_sel.pack(anchor="w")

    # ---- Card: Log (área que expande ao redimensionar) ----
    frm_log = Frame(root, bg=_COR_CARD, padx=20, pady=14, highlightbackground=_COR_BORDA, highlightthickness=1)
    frm_log.pack(fill=BOTH, expand=True, padx=20, pady=(0, 8))

    Label(
        frm_log,
        text="Log / Status",
        font=("Segoe UI", 11, "bold"),
        fg=_COR_TITULO,
        bg=_COR_CARD,
    ).pack(anchor="w")

    txt_log = ScrolledText(
        frm_log,
        height=8,
        wrap="word",
        font=("Consolas", 9),
        bg="#f8fafc",
        fg=_COR_TEXTO,
        insertbackground=_COR_TEXTO,
        relief="flat",
        padx=10,
        pady=8,
    )
    txt_log.pack(fill=BOTH, expand=True, pady=(8, 0))

    # ---- Barra de botões (espaço fixo, não aperta ao redimensionar) ----
    frm_btns = Frame(root, bg=_COR_FUNDO, pady=16, padx=20)
    frm_btns.pack(fill=X, side="bottom")

    def append_log(msg: str):
        try:
            txt_log.insert("end", msg + "\n")
            txt_log.see("end")
        except Exception:
            pass

    def selecionar_arquivo():
        path = filedialog.askopenfilename(
            title="Selecionar arquivo TXT REGATI/PEFIN",
            initialdir=str(pasta_inicial),
            filetypes=[
                ("Arquivos TXT", "*.txt;*.TXT"),
                ("Todos os arquivos", "*.*"),
            ],
        )
        if path:
            arquivo_selecionado[0] = Path(path).resolve()
            if not arquivo_selecionado[0].exists():
                messagebox.showerror("Erro", "Arquivo não encontrado.")
                return
            nome = arquivo_selecionado[0].name
            lbl_arquivo.config(
                text=f"📄 {nome}" if len(nome) < 55 else f"📄 ...{nome[-50:]}",
                fg=_COR_TITULO,
            )
            txt_log.delete("1.0", "end")
            append_log(f"Arquivo selecionado: {arquivo_selecionado[0]}")

    btn_sel.configure(command=selecionar_arquivo)

    def converter():
        if not arquivo_selecionado[0] or not arquivo_selecionado[0].exists():
            messagebox.showwarning("Aviso", "Selecione um arquivo TXT antes de converter.")
            return
        btn_convert.config(state=DISABLED)
        btn_sel.config(state=DISABLED)
        btn_sel2.config(state=DISABLED)
        txt_log.delete("1.0", "end")
        append_log("Iniciando conversão... (aguarde)")

        def _concluido(ok, erro):
            btn_convert.config(state=NORMAL)
            btn_sel.config(state=NORMAL)
            btn_sel2.config(state=NORMAL)
            if ok:
                append_log("Conversão concluída com sucesso.")
                messagebox.showinfo("Concluído", "Planilha XLSX gerada com sucesso.\nO nome inclui a data e hora da conversão.")
            else:
                append_log(f"Erro: {erro}")
                messagebox.showerror("Erro na conversão", str(erro))

        def executar():
            try:
                converter_txt_para_xlsx(
                    arquivo_selecionado[0],
                    arquivo_xlsx=None,
                    log_callback=lambda m: root.after(0, lambda msg=m: append_log(msg)),
                )
                root.after(0, lambda: _concluido(True, None))
            except BaseException as e:
                # SystemExit (ex.: coluna não encontrada) não é Exception;
                # sem capturar, a GUI ficava travada com botões desabilitados
                err_msg = str(e)
                root.after(0, lambda: _concluido(False, err_msg))

        threading.Thread(target=executar, daemon=True).start()

    btn_convert = ttk.Button(
        frm_btns,
        text="Converter para XLSX",
        command=converter,
        style="Primario.TButton",
    )
    btn_convert.pack(side=RIGHT, padx=(8, 0))

    btn_sel2 = ttk.Button(
        frm_btns,
        text="Selecionar arquivo",
        command=selecionar_arquivo,
        style="Secundario.TButton",
    )
    btn_sel2.pack(side=RIGHT)

    try:
        root.mainloop()
    except Exception as e:
        messagebox.showerror("Erro", str(e))


def main():
    try:
        pasta = Path(__file__).resolve().parent
        if len(sys.argv) >= 2:
            # Linha de comando: converte o arquivo informado (nome com data/hora)
            arquivo_txt = Path(sys.argv[1]).resolve()
            if not arquivo_txt.is_absolute():
                arquivo_txt = (pasta / sys.argv[1]).resolve()
            if not arquivo_txt.exists():
                print(f"Erro: arquivo não encontrado: {arquivo_txt}")
                input("Pressione Enter para sair...")
                sys.exit(1)
            arquivo_xlsx = Path(sys.argv[2]).resolve() if len(sys.argv) >= 3 else None
            if arquivo_xlsx and not arquivo_xlsx.is_absolute():
                arquivo_xlsx = pasta / sys.argv[2]
            converter_txt_para_xlsx(arquivo_txt, arquivo_xlsx)
        else:
            # Sem argumentos: abre a interface para escolher o TXT
            _rodar_interface()
    except Exception as e:
        msg = str(e)
        print(f"Erro: {msg}")
        try:
            root = Tk()
            root.withdraw()
            messagebox.showerror("Erro", msg)
            root.destroy()
        except Exception:
            input("Pressione Enter para sair...")
        sys.exit(1)


if __name__ == "__main__":
    main()
