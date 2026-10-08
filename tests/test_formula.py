r"""Formula engine: evaluation, Excel rounding, reference shifting, translation.

Run from the project folder:  .venv\Scripts\python.exe tests\test_formula.py"""
import os, sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
from atlas_formula import (Calc, shift_refs, rename_sheet_refs, translate, format_general,
                           DIV0, CIRC, NAME, REF)

cells = {("S", 1, 1): 2.0, ("S", 2, 1): 3.0, ("S", 3, 1): "x", ("S", 1, 2): "=A1+A2",
         ("S", 2, 2): "=SUM(A1:A3)", ("S", 3, 2): "=ROUND(2.5,0)", ("S", 4, 2): "=ROUND(-2.5,0)",
         ("S", 5, 2): "=ROUND(SQRT((E5-9)^2+(F5-6)^2),0)", ("S", 5, 5): 11.0, ("S", 5, 6): 6.0,
         ("S", 6, 2): "=1/0", ("S", 7, 2): "=B7", ("S", 8, 2): "=FOO(1)", ("S", 9, 2): "=-2^2",
         ("S", 10, 2): "='Other Sheet'!A1*2", ("Other Sheet", 1, 1): 21.0,
         ("S", 11, 2): '="a"&A1&"b"', ("S", 12, 2): "=IF(A1>1,\"big\",\"small\")",
         ("S", 13, 2): "=AVERAGE(A1:A3)", ("S", 14, 2): "=50%", ("S", 15, 2): "=ROUND(G5*1.5,1)",
         ("S", 5, 7): 3.0, ("S", 16, 2): "=COUNT(A1:A3)", ("S", 17, 2): "=A1+A3"}
calc = Calc(lambda s, r, c: cells.get((s, r, c)), lambda: ["S", "Other Sheet"])
v = lambda r, c: calc.value("S", r, c)
assert v(1, 2) == 5 and v(2, 2) == 5, (v(1, 2), v(2, 2))
assert v(3, 2) == 3 and v(4, 2) == -3
assert v(5, 2) == 2
assert v(6, 2) == DIV0 and v(7, 2) == CIRC and v(8, 2) == NAME
assert v(9, 2) == 4, v(9, 2)
assert v(10, 2) == 42
assert v(11, 2) == "a2b" and v(12, 2) == "big"
assert v(13, 2) == 2.5 and abs(v(14, 2) - 0.5) < 1e-12 and v(15, 2) == 4.5
assert v(16, 2) == 2 and v(17, 2) == "#VALUE!"
assert format_general(4.5) == "4.5" and format_general(5.0) == "5" and format_general(0.1 + 0.2) == "0.3"

# shifting
assert shift_refs("=A1+A5", "S", "S", "row", 3, 2) == "=A1+A7"
assert shift_refs("=SUM(A1:A5)", "S", "S", "row", 3, 2) == "=SUM(A1:A7)"
assert shift_refs("=SUM(A1:A5)", "S", "S", "row", 2, -2) == "=SUM(A1:A3)"
assert shift_refs("=A3", "S", "S", "row", 2, -2) == "=#REF!"
assert shift_refs("=A3", "T", "S", "row", 1, 5) == "=A3"          # other sheet untouched
assert shift_refs("=S!A3", "T", "S", "row", 1, 5) == "=S!A8"
assert shift_refs("=ROUND(SQRT((E5-9)^2+(F5-6)^2),0)", "S", "S", "col", 5, 1) == "=ROUND(SQRT((F5-9)^2+(G5-6)^2),0)"
assert rename_sheet_refs("='Other Sheet'!A1*2", "Other Sheet", "New") == "=New!A1*2"
assert rename_sheet_refs("=New!A1", "New", "Two words") == "='Two words'!A1"
assert translate("=A1+$A$1+A$1", 2, 1) == "=B3+$A$1+B$1"
print("OK")
