"""Provider answers for tests - shapes copied from the live probe of 2026-10-25; every figure is invented."""


def answer(matrix):
    """A res_ExecuteGrid like the provider's, for a dense matrix (float = data cell)."""
    h, w = len(matrix), len(matrix[0])
    cells = [v for row in matrix for v in row]
    vals = "|".join("" if v is None else (repr(v) if isinstance(v, float) else v) for v in cells)
    types = "|".join("2" if isinstance(v, float) else ("7" if v is None else "0") for v in cells)
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes" ?><res_ExecuteGrid><grid><slices>'
            f'<slice rows="{h}" cols="{w}"><data><generations/><range start="0" end="{h * w - 1}">'
            f'<vals enc="1">{vals}</vals><types enc="0">{types}</types></range></data></slice></slices></grid>'
            '<stats><Total_Provider_Time>9</Total_Provider_Time></stats></res_ExecuteGrid>')
