"""PORTFOLIO — parsers TXT omitidos."""
from portfolio_omitted import omit


def __getattr__(name):
    def _m(*a, **k):
        omit(f"txt_utils.{name}")

    return _m
