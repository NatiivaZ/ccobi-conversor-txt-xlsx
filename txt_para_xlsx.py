# -*- coding: utf-8 -*-
"""Conversor principal de TXT REGATI/PEFIN para XLSX.

Pode rodar pela interface ou direto pela linha de comando, mas a lógica de
conversão fica concentrada aqui.
"""

import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import Tk, Label, Frame, filedialog, messagebox, DISABLED, NORMAL, X, BOTH, RIGHT
from tkinter.scrolledtext import ScrolledText
from tkinter import ttk

import xlsxwriter
from txt_utils import (
    achar_indices_colunas,
    formatar_cpf_cnpj,
    formatar_hora_como_texto,
    nome_planilha_com_data_hora,
    parse_data_ddmmaaaa,
    valor_centavos_para_reais,
)


def converter_txt_para_xlsx(arquivo_txt: Path, arquivo_xlsx: Path = None, log_callback=None) -> Path:
    """Converte o TXT em XLSX.

    Se o destino não vier informado, cria um nome novo com data e hora para não
    sobrescrever uma conversão anterior.
    """
    arquivo_txt = Path(arquivo_txt).resolve()
    if not arquivo_txt.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {arquivo_txt}")
    if arquivo_xlsx is None:
        arquivo_xlsx = nome_planilha_com_data_hora(arquivo_txt)
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

    # Essa é a ordem final da planilha exportada.
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

            # Número do auto.
            sheet.write_string(row_excel, 0, contrato, fmt_texto)

            # Data de inscrição.
            dt = parse_data_ddmmaaaa(dt_inc)
            if dt:
                sheet.write_datetime(row_excel, 1, datetime.combine(dt, datetime.min.time()), fmt_data)
            else:
                sheet.write_string(row_excel, 1, dt_inc or "", fmt_texto)

            # Data da ocorrência.
            dt_oc = parse_data_ddmmaaaa(dt_ocor)
            if dt_oc:
                sheet.write_datetime(row_excel, 2, datetime.combine(dt_oc, datetime.min.time()), fmt_data)
            else:
                sheet.write_string(row_excel, 2, dt_ocor or "", fmt_texto)

            # Data de exclusão.
            dt_ex = parse_data_ddmmaaaa(dt_excl)
            if dt_ex:
                sheet.write_datetime(row_excel, 3, datetime.combine(dt_ex, datetime.min.time()), fmt_data)
            else:
                sheet.write_string(row_excel, 3, dt_excl or "", fmt_texto)

            # Valor inscrito.
            v = valor_centavos_para_reais(val)
            if v is not None:
                sheet.write_number(row_excel, 4, v, fmt_moeda)
            else:
                sheet.write_string(row_excel, 4, val or "", fmt_texto)

            # Documento principal.
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


# Paleta e estilo da interface.
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
    """Abre a interface gráfica do conversor."""
    try:
        root = Tk()
    except Exception as e:
        messagebox.showerror("Erro", f"Não foi possível abrir a interface:\n{e}")
        return

    root.title("Conversor REGATI/PEFIN — TXT para XLSX")
    root.minsize(560, 520)
    root.resizable(True, True)
    root.configure(bg=_COR_FUNDO)

    # Estilos dos botões principais.
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

    # Cabeçalho da janela.
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

    # Bloco do arquivo selecionado.
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

    # Área de log e andamento da conversão.
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

    # Barra de ações da parte de baixo.
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
                # Alguns erros passam como BaseException. Se não capturar aqui,
                # a interface pode ficar travada com os botões desabilitados.
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
            # Com argumento, roda direto pela linha de comando.
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
            # Sem argumento, abre a interface normal.
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
