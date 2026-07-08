"""PORTFOLIO — conversao TXT REGATI/PEFIN omitida."""
from portfolio_omitted import omit


def converter_txt_para_xlsx(*args, **kwargs):
    omit("conversao TXT -> XLSX (mapeamento REGATI/PEFIN)")


def main():
    omit("execucao do conversor TXT")


def __getattr__(name):
    def _m(*a, **k):
        omit(f"txt_para_xlsx.{name}")

    return _m
