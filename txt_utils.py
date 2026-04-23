"""Funções menores que sustentam a conversão do TXT para XLSX."""

import re
from datetime import datetime
from pathlib import Path


NOME_DOC_PRINCIPAL = "DOC PRINCIPAL 1"
NOME_VALOR = "VALOR"
NOME_CONTRATO = "CONTRATO"
NOME_DT_INCLUSAO = "DT INCLUSAO"
NOME_DT_OCORRENCIA = "DT OCORRENCIA"
NOME_DT_EXCLUSAO = "DT EXCLUSAO"
NOME_HR_INCLUSAO = "HR INCLUSAO"
NOME_USER_INCLUSAO = "USER INCLUSAO"


def achar_indices_colunas(header: str, sep: str = ";"):
    """Lê o cabeçalho e devolve onde está cada coluna que interessa."""
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
    """Valida CPF pelo cálculo padrão dos dígitos."""
    if len(digits) != 11 or not digits.isdigit():
        return False
    if digits == digits[0] * 11:
        return False
    soma = sum(int(digits[i]) * (10 - i) for i in range(9))
    digito_1 = 0 if (soma % 11) < 2 else 11 - (soma % 11)
    if int(digits[9]) != digito_1:
        return False
    soma_2 = sum(int(digits[i]) * (11 - i) for i in range(10))
    digito_2 = 0 if (soma_2 % 11) < 2 else 11 - (soma_2 % 11)
    return int(digits[10]) == digito_2


def _validar_cnpj(digits: str) -> bool:
    """Valida CNPJ pelo cálculo padrão dos dígitos."""
    if len(digits) != 14 or not digits.isdigit():
        return False
    pesos1 = (5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    soma1 = sum(int(digits[i]) * pesos1[i] for i in range(12))
    digito_1 = 0 if (soma1 % 11) < 2 else 11 - (soma1 % 11)
    if int(digits[12]) != digito_1:
        return False
    pesos2 = (6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2)
    soma2 = sum(int(digits[i]) * pesos2[i] for i in range(13))
    digito_2 = 0 if (soma2 % 11) < 2 else 11 - (soma2 % 11)
    return int(digits[13]) == digito_2


def formatar_cpf_cnpj(val: str) -> str:
    """Tenta formatar como CPF ou CNPJ sem forçar quando o número não fecha."""
    if not val:
        return ""
    digits = re.sub(r"\D", "", val.strip())
    if len(digits) == 11:
        if _validar_cpf(digits):
            return f"{digits[:3]}.{digits[3:6]}.{digits[6:9]}-{digits[9:]}"
        return val.strip()
    if len(digits) in {12, 13}:
        d11 = digits[-11:]
        if _validar_cpf(d11):
            return f"{d11[:3]}.{d11[3:6]}.{d11[6:9]}-{d11[9:]}"
        return val.strip()
    if len(digits) >= 14:
        d14 = digits[-14:] if len(digits) > 14 else digits
        if _validar_cnpj(d14):
            return f"{d14[:2]}.{d14[2:5]}.{d14[5:8]}/{d14[8:12]}-{d14[12:]}"
        d11 = d14[-11:]
        if _validar_cpf(d11):
            return f"{d11[:3]}.{d11[3:6]}.{d11[6:9]}-{d11[9:]}"
        return val.strip()
    return val.strip()


def parse_data_ddmmaaaa(val: str):
    """Converte texto no formato DDMMAAAA para data."""
    if not val or len(val.strip()) < 8:
        return None
    digits = re.sub(r"\D", "", val.strip())[:8]
    if len(digits) != 8:
        return None
    try:
        dia, mes, ano = int(digits[:2]), int(digits[2:4]), int(digits[4:8])
        if 1 <= dia <= 31 and 1 <= mes <= 12 and 1900 <= ano <= 2100:
            return datetime(ano, mes, dia).date()
    except ValueError:
        pass
    return None


def formatar_hora_como_texto(val: str) -> str:
    """Transforma hora do TXT em um texto mais amigável para o Excel."""
    if not val:
        return ""
    digits = re.sub(r"\D", "", val.strip())
    if len(digits) == 4:
        try:
            horas, minutos = int(digits[:2]), int(digits[2:4])
            if 0 <= horas <= 23 and 0 <= minutos <= 59:
                return f"{horas:02d}:{minutos:02d}"
        except ValueError:
            pass
    if len(digits) >= 6:
        try:
            horas, minutos, segundos = int(digits[:2]), int(digits[2:4]), int(digits[4:6])
            if 0 <= horas <= 23 and 0 <= minutos <= 59 and 0 <= segundos <= 59:
                return f"{horas:02d}:{minutos:02d}:{segundos:02d}"
        except ValueError:
            pass
    return val.strip()


def valor_centavos_para_reais(val: str):
    """Converte o valor em centavos para reais."""
    if not val:
        return None
    digits = re.sub(r"\D", "", val.strip())
    if not digits:
        return None
    try:
        return int(digits) / 100.0
    except ValueError:
        return None


def nome_planilha_com_data_hora(arquivo_txt: Path) -> Path:
    """Cria o nome do arquivo de saída com data e hora para não sobrescrever."""
    agora = datetime.now()
    return arquivo_txt.parent / f"{arquivo_txt.stem}_{agora:%Y%m%d}_{agora:%H%M%S}.xlsx"
