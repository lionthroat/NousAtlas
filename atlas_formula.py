"""A small formula engine: arithmetic, cell references, ranges and a handful of
functions. It covers what text-heavy workbooks actually use (ROUND, SQRT,
SUM, ...) and nothing else.

Formulas are kept as the same text Excel stores ("=ROUND(G5*1.5,1)"), so the
workbook stays a normal .xlsx file. Besides evaluating, this module can move
references when rows or columns are inserted or deleted, rename a sheet
inside formulas, and shift relative references for copy and paste.
"""

import decimal
import math
import re

# ---------------------------------------------------------------- errors


class XlError(str):
    """An Excel error value such as #DIV/0!. A str so it displays as itself."""


DIV0 = XlError("#DIV/0!")
VALUE = XlError("#VALUE!")
REF = XlError("#REF!")
NAME = XlError("#NAME?")
NUM = XlError("#NUM!")
NA = XlError("#N/A")
CIRC = XlError("#CIRC!")      # Excel shows 0 and a warning; we say what happened
ERRORS = {e: e for e in (DIV0, VALUE, REF, NAME, NUM, NA, CIRC, XlError("#NULL!"))}


# ---------------------------------------------------------------- A1 helpers


def col_to_num(letters):
    n = 0
    for ch in letters.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def num_to_col(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s


# ---------------------------------------------------------------- tokens

_SHEET = r"(?:'(?:[^']|'')+'|[A-Za-z_][\w.]*)!"
_CELL = r"\$?[A-Za-z]{1,3}\$?\d+"
_TOKEN_RE = re.compile(
    r"(?P<ws>\s+)"
    r"|(?P<str>\"(?:[^\"]|\"\")*\")"
    r"|(?P<err>#(?:DIV/0!|VALUE!|REF!|NAME\?|NUM!|N/A|NULL!|CIRC!))"
    rf"|(?P<ref>(?:{_SHEET})?{_CELL}(?::{_CELL})?)(?![\w(])"
    r"|(?P<func>[A-Za-z_][A-Za-z0-9_.]*)(?=\s*\()"
    r"|(?P<num>(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)"
    r"|(?P<bool>TRUE|FALSE)(?![\w(])"
    r"|(?P<op><>|<=|>=|[-+*/^&=<>%(),:])"
    r"|(?P<name>[A-Za-z_][A-Za-z0-9_.]*)",
    re.IGNORECASE)

_CELL_PARTS = re.compile(r"(\$?)([A-Za-z]{1,3})(\$?)(\d+)")


def tokenize(text):
    """Yield (kind, text, start) for the formula body (without the '=')."""
    pos = 0
    out = []
    while pos < len(text):
        m = _TOKEN_RE.match(text, pos)
        if not m:
            raise SyntaxError(f"Unexpected '{text[pos]}'")
        kind = m.lastgroup
        if kind != "ws":
            out.append((kind, m.group(), pos))
        pos = m.end()
    return out


class Ref:
    """A parsed reference: optional sheet, and a (possibly one-cell) rectangle.
    Each corner keeps its own $ flags so it can be written back unchanged."""

    __slots__ = ("sheet", "c1", "r1", "c2", "r2", "abs", "is_range")

    def __init__(self, sheet, c1, r1, c2, r2, abs_flags, is_range):
        self.sheet, self.c1, self.r1, self.c2, self.r2 = sheet, c1, r1, c2, r2
        self.abs = abs_flags          # (c1abs, r1abs, c2abs, r2abs)
        self.is_range = is_range

    @classmethod
    def parse(cls, text):
        sheet = None
        if "!" in text:
            sheet, text = text.rsplit("!", 1)
            if sheet.startswith("'"):
                sheet = sheet[1:-1].replace("''", "'")
        parts = text.split(":")
        a = _CELL_PARTS.fullmatch(parts[0])
        b = _CELL_PARTS.fullmatch(parts[-1])
        c1, r1 = col_to_num(a.group(2)), int(a.group(4))
        c2, r2 = col_to_num(b.group(2)), int(b.group(4))
        flags = [bool(a.group(1)), bool(a.group(3)), bool(b.group(1)), bool(b.group(3))]
        # keep the rectangle normalised (top-left first)
        if c2 < c1:
            c1, c2 = c2, c1
            flags[0], flags[2] = flags[2], flags[0]
        if r2 < r1:
            r1, r2 = r2, r1
            flags[1], flags[3] = flags[3], flags[1]
        return cls(sheet, c1, r1, c2, r2, tuple(flags), len(parts) == 2)

    def text(self):
        def cell(c, r, ca, ra):
            return f"{'$' if ca else ''}{num_to_col(c)}{'$' if ra else ''}{r}"
        s = cell(self.c1, self.r1, self.abs[0], self.abs[1])
        if self.is_range:
            s += ":" + cell(self.c2, self.r2, self.abs[2], self.abs[3])
        if self.sheet is not None:
            s = quote_sheet(self.sheet) + "!" + s
        return s


def quote_sheet(name):
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_.]*", name) and not _CELL_PARTS.fullmatch(name):
        return name
    return "'" + name.replace("'", "''") + "'"


# ---------------------------------------------------------------- parser
# AST nodes are tuples: ("num", v) ("str", v) ("bool", v) ("err", e)
# ("ref", Ref) ("func", NAME, [args]) ("bin", op, a, b) ("neg", a) ("pct", a)


class _Parser:
    def __init__(self, tokens):
        self.t = tokens
        self.i = 0

    def peek(self):
        return self.t[self.i] if self.i < len(self.t) else (None, None, None)

    def take(self, text=None):
        tok = self.peek()
        if tok[0] is None or (text is not None and tok[1] != text):
            raise SyntaxError(f"Expected {text or 'more'}")
        self.i += 1
        return tok

    def parse(self):
        node = self.compare()
        if self.i != len(self.t):
            raise SyntaxError(f"Unexpected '{self.peek()[1]}'")
        return node

    def compare(self):
        node = self.concat()
        while self.peek()[1] in ("=", "<>", "<", ">", "<=", ">="):
            op = self.take()[1]
            node = ("bin", op, node, self.concat())
        return node

    def concat(self):
        node = self.additive()
        while self.peek()[1] == "&":
            self.take()
            node = ("bin", "&", node, self.additive())
        return node

    def additive(self):
        node = self.term()
        while self.peek()[1] in ("+", "-"):
            op = self.take()[1]
            node = ("bin", op, node, self.term())
        return node

    def term(self):
        node = self.power()
        while self.peek()[1] in ("*", "/"):
            op = self.take()[1]
            node = ("bin", op, node, self.power())
        return node

    def power(self):
        node = self.unary()
        while self.peek()[1] == "^":
            self.take()
            node = ("bin", "^", node, self.unary())
        return node

    def unary(self):
        if self.peek()[1] == "-":
            self.take()
            return ("neg", self.unary())
        if self.peek()[1] == "+":
            self.take()
            return self.unary()
        node = self.primary()
        while self.peek()[1] == "%":
            self.take()
            node = ("pct", node)
        return node

    def primary(self):
        kind, text, _ = self.take()
        if kind == "num":
            return ("num", float(text))
        if kind == "str":
            return ("str", text[1:-1].replace('""', '"'))
        if kind == "bool":
            return ("bool", text.upper() == "TRUE")
        if kind == "err":
            return ("err", ERRORS.get(text.upper(), XlError(text.upper())))
        if kind == "ref":
            return ("ref", Ref.parse(text))
        if kind == "func":
            name = text.upper()
            self.take("(")
            args = []
            if self.peek()[1] != ")":
                while True:
                    if self.peek()[1] in (",", ")"):
                        args.append(("missing",))
                    else:
                        args.append(self.compare())
                    if self.peek()[1] == ",":
                        self.take()
                        continue
                    break
            self.take(")")
            return ("func", name, args)
        if text == "(":
            node = self.compare()
            self.take(")")
            return node
        if kind == "name":
            return ("err", NAME)
        raise SyntaxError(f"Unexpected '{text}'")


_ast_cache = {}


def parse(formula):
    """Parse '=...' (or the body without '='). Raises SyntaxError."""
    body = formula[1:] if formula.startswith("=") else formula
    node = _ast_cache.get(body)
    if node is None:
        node = _Parser(tokenize(body)).parse()
        if len(_ast_cache) > 5000:
            _ast_cache.clear()
        _ast_cache[body] = node
    return node


def is_formula(value):
    return isinstance(value, str) and value.startswith("=") and len(value) > 1


# ---------------------------------------------------------------- evaluation


def excel_round(x, digits, mode="half"):
    """Excel rounding: half away from zero (not Python's banker's rounding)."""
    digits = int(digits)
    q = decimal.Decimal(1).scaleb(-digits)
    d = decimal.Decimal(repr(float(x)))
    rounding = {"half": decimal.ROUND_HALF_UP, "up": decimal.ROUND_UP,
                "down": decimal.ROUND_DOWN}[mode]
    return float(d.quantize(q, rounding=rounding))


def to_number(v):
    if isinstance(v, XlError):
        return v
    if v is None:
        return 0.0
    if isinstance(v, bool):
        return 1.0 if v else 0.0
    if isinstance(v, (int, float)):
        return float(v)
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            return 0.0
        try:
            return float(s.rstrip("%")) / (100 if s.endswith("%") else 1)
        except ValueError:
            return VALUE
    return VALUE


def to_text(v):
    if v is None:
        return ""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, float):
        return format_general(v)
    return str(v)


def to_bool(v):
    if isinstance(v, XlError):
        return v
    if isinstance(v, str):
        if v.upper() in ("TRUE", "FALSE"):
            return v.upper() == "TRUE"
        return VALUE
    n = to_number(v)
    return n if isinstance(n, XlError) else n != 0


def format_general(v):
    """Excel's 'General' number display: up to ~11 significant digits."""
    if isinstance(v, bool):
        return "TRUE" if v else "FALSE"
    if isinstance(v, int):
        return str(v)
    if math.isnan(v) or math.isinf(v):
        return NUM
    if v == int(v) and abs(v) < 1e11:
        return str(int(v))
    s = f"{v:.10g}"
    if "e" in s:
        mant, exp = s.split("e")
        return f"{mant}E{'+' if int(exp) >= 0 else '-'}{abs(int(exp)):02d}"
    return s


class RangeValue(list):
    """Values of a range, row by row, flattened. Marks 'came from a range' so
    SUM etc. can skip text and booleans the way Excel does."""


class Calc:
    """Evaluates formulas against a workbook. `get_raw(sheet, row, col)`
    returns what is typed in a cell (value or '=formula'); results are cached
    until invalidate() is called."""

    def __init__(self, get_raw, sheet_names):
        self.get_raw = get_raw
        self.sheet_names = sheet_names      # callable returning current titles
        self.cache = {}
        self.busy = set()

    def invalidate(self):
        self.cache.clear()

    def value(self, sheet, row, col):
        key = (sheet, row, col)
        if key in self.cache:
            return self.cache[key]
        raw = self.get_raw(sheet, row, col)
        if not is_formula(raw):
            return raw
        if key in self.busy:
            return CIRC
        self.busy.add(key)
        try:
            try:
                node = parse(raw)
            except SyntaxError:
                result = NAME
            else:
                result = self.scalar(self.ev(node, sheet), sheet)
        except RecursionError:
            result = CIRC
        finally:
            self.busy.discard(key)
        if isinstance(result, float) and (math.isnan(result) or math.isinf(result)):
            result = NUM
        self.cache[key] = result
        return result

    # -- helpers
    def _sheet_of(self, ref, here):
        if ref.sheet is None:
            return here
        for name in self.sheet_names():
            if name.lower() == ref.sheet.lower():
                return name
        return None

    def range_values(self, ref, here):
        sheet = self._sheet_of(ref, here)
        if sheet is None:
            return REF
        out = RangeValue()
        for r in range(ref.r1, ref.r2 + 1):
            for c in range(ref.c1, ref.c2 + 1):
                out.append(self.value(sheet, r, c))
        return out

    def scalar(self, v, here):
        if isinstance(v, RangeValue):
            return v[0] if len(v) == 1 else VALUE
        return v

    def ev(self, node, here):
        kind = node[0]
        if kind in ("num", "str", "bool", "err"):
            return node[1]
        if kind == "missing":
            return None
        if kind == "ref":
            ref = node[1]
            if ref.is_range:
                return self.range_values(ref, here)
            sheet = self._sheet_of(ref, here)
            if sheet is None:
                return REF
            v = self.value(sheet, ref.r1, ref.c1)
            return v
        if kind == "neg":
            n = to_number(self.scalar(self.ev(node[1], here), here))
            return n if isinstance(n, XlError) else -n
        if kind == "pct":
            n = to_number(self.scalar(self.ev(node[1], here), here))
            return n if isinstance(n, XlError) else n / 100
        if kind == "bin":
            return self.binary(node[1], self.scalar(self.ev(node[2], here), here),
                               self.scalar(self.ev(node[3], here), here))
        if kind == "func":
            fn = FUNCTIONS.get(node[1])
            if fn is None:
                return NAME
            return fn(self, node[2], here)
        return VALUE

    def binary(self, op, a, b):
        if isinstance(a, XlError):
            return a
        if isinstance(b, XlError):
            return b
        if op == "&":
            return to_text(a) + to_text(b)
        if op in ("=", "<>", "<", ">", "<=", ">="):
            return _compare(op, a, b)
        x, y = to_number(a), to_number(b)
        if isinstance(x, XlError):
            return x
        if isinstance(y, XlError):
            return y
        if op == "+":
            return x + y
        if op == "-":
            return x - y
        if op == "*":
            return x * y
        if op == "/":
            return DIV0 if y == 0 else x / y
        if op == "^":
            try:
                r = x ** y
            except (OverflowError, ZeroDivisionError):
                return NUM
            return NUM if isinstance(r, complex) else r
        return VALUE


def _compare(op, a, b):
    def key(v):
        # Excel orders numbers < text < booleans; empty acts like 0 or ""
        if isinstance(v, bool):
            return (2, v)
        if isinstance(v, (int, float)):
            return (0, v)
        if v is None:
            return None
        return (1, str(v).lower())
    ka, kb = key(a), key(b)
    if ka is None:
        ka = (0, 0) if kb is None or kb[0] == 0 else (1, "") if kb[0] == 1 else (2, False)
    if kb is None:
        kb = (0, 0) if ka[0] == 0 else (1, "") if ka[0] == 1 else (2, False)
    return {"=": ka == kb, "<>": ka != kb, "<": ka < kb, ">": ka > kb,
            "<=": ka <= kb, ">=": ka >= kb}[op]


# ---------------------------------------------------------------- functions


def _numbers(calc, args, here, count_text=False):
    """Numbers from the arguments, Excel style: inside ranges only real
    numbers count; typed-in arguments are coerced."""
    out = []
    for a in args:
        v = calc.ev(a, here)
        if isinstance(v, RangeValue):
            for x in v:
                if isinstance(x, XlError):
                    return x
                if isinstance(x, (int, float)) and not isinstance(x, bool):
                    out.append(float(x))
        else:
            if isinstance(v, XlError):
                return v
            if v is None:
                continue
            n = to_number(v)
            if isinstance(n, XlError):
                return n
            out.append(n)
    return out


def _fixed(n_args):
    """Decorator: evaluate exactly these scalar args (missing → None)."""
    def wrap(fn):
        def run(calc, args, here):
            lo, hi = n_args
            if not (lo <= len(args) <= hi):
                return VALUE
            vals = [calc.scalar(calc.ev(a, here), here) for a in args]
            for v in vals:
                if isinstance(v, XlError):
                    return v
            return fn(*vals)
        return run
    return wrap


def _num_args(fn):
    def run(*vals):
        nums = [to_number(v) for v in vals]
        for n in nums:
            if isinstance(n, XlError):
                return n
        return fn(*nums)
    return run


def f_sum(calc, args, here):
    nums = _numbers(calc, args, here)
    return nums if isinstance(nums, XlError) else float(math.fsum(nums))


def f_average(calc, args, here):
    nums = _numbers(calc, args, here)
    if isinstance(nums, XlError):
        return nums
    return DIV0 if not nums else math.fsum(nums) / len(nums)


def f_min(calc, args, here):
    nums = _numbers(calc, args, here)
    return nums if isinstance(nums, XlError) else (min(nums) if nums else 0.0)


def f_max(calc, args, here):
    nums = _numbers(calc, args, here)
    return nums if isinstance(nums, XlError) else (max(nums) if nums else 0.0)


def f_count(calc, args, here):
    n = 0
    for a in args:
        v = calc.ev(a, here)
        vals = v if isinstance(v, RangeValue) else [v]
        for x in vals:
            if isinstance(x, (int, float)) and not isinstance(x, bool):
                n += 1
            elif not isinstance(v, RangeValue) and not isinstance(to_number(x), XlError) and x is not None:
                n += 1
    return float(n)


def f_counta(calc, args, here):
    n = 0
    for a in args:
        v = calc.ev(a, here)
        vals = v if isinstance(v, RangeValue) else [v]
        n += sum(1 for x in vals if x is not None and x != "")
    return float(n)


def f_if(calc, args, here):
    if not 1 <= len(args) <= 3:
        return VALUE
    cond = to_bool(calc.scalar(calc.ev(args[0], here), here))
    if isinstance(cond, XlError):
        return cond
    if cond:
        return calc.scalar(calc.ev(args[1], here), here) if len(args) > 1 else True
    return calc.scalar(calc.ev(args[2], here), here) if len(args) > 2 else False


def f_iferror(calc, args, here):
    if len(args) != 2:
        return VALUE
    v = calc.scalar(calc.ev(args[0], here), here)
    return calc.scalar(calc.ev(args[1], here), here) if isinstance(v, XlError) else v


def _logic(combine):
    def run(calc, args, here):
        vals = []
        for a in args:
            v = calc.ev(a, here)
            for x in (v if isinstance(v, RangeValue) else [v]):
                if x is None or (isinstance(v, RangeValue) and isinstance(x, str)):
                    continue
                b = to_bool(x)
                if isinstance(b, XlError):
                    return b
                vals.append(b)
        return VALUE if not vals else combine(vals)
    return run


def f_concat(calc, args, here):
    parts = []
    for a in args:
        v = calc.ev(a, here)
        for x in (v if isinstance(v, RangeValue) else [v]):
            if isinstance(x, XlError):
                return x
            parts.append(to_text(x))
    return "".join(parts)


def _sqrt(x):
    return NUM if x < 0 else math.sqrt(x)


def _mod(a, b):
    return DIV0 if b == 0 else a - b * math.floor(a / b)


def _power(a, b):
    try:
        r = a ** b
    except (OverflowError, ZeroDivisionError):
        return NUM
    return NUM if isinstance(r, complex) else r


FUNCTIONS = {
    "SUM": f_sum, "AVERAGE": f_average, "MIN": f_min, "MAX": f_max,
    "COUNT": f_count, "COUNTA": f_counta,
    "ROUND": _fixed((1, 2))(_num_args(lambda x, d=0: excel_round(x, d))),
    "ROUNDUP": _fixed((1, 2))(_num_args(lambda x, d=0: excel_round(x, d, "up"))),
    "ROUNDDOWN": _fixed((1, 2))(_num_args(lambda x, d=0: excel_round(x, d, "down"))),
    "INT": _fixed((1, 1))(_num_args(lambda x: float(math.floor(x)))),
    "ABS": _fixed((1, 1))(_num_args(abs)),
    "SQRT": _fixed((1, 1))(_num_args(_sqrt)),
    "MOD": _fixed((2, 2))(_num_args(_mod)),
    "POWER": _fixed((2, 2))(_num_args(_power)),
    "PI": _fixed((0, 0))(lambda: math.pi),
    "IF": f_if, "IFERROR": f_iferror,
    "AND": _logic(all), "OR": _logic(any),
    "NOT": _fixed((1, 1))(lambda v: (lambda b: b if isinstance(b, XlError) else not b)(to_bool(v))),
    "CONCAT": f_concat, "CONCATENATE": f_concat,
    "LEN": _fixed((1, 1))(lambda v: float(len(to_text(v)))),
    "UPPER": _fixed((1, 1))(lambda v: to_text(v).upper()),
    "LOWER": _fixed((1, 1))(lambda v: to_text(v).lower()),
    "TRIM": _fixed((1, 1))(lambda v: " ".join(to_text(v).split())),
}


# ---------------------------------------------------------------- rewriting


def _rewrite_refs(formula, fn):
    """Return formula with each reference token passed through fn(Ref) ->
    Ref | XlError | None (None = unchanged)."""
    if not is_formula(formula):
        return formula
    body = formula[1:]
    try:
        tokens = tokenize(body)
    except SyntaxError:
        return formula
    out, last = [], 0
    changed = False
    for kind, text, start in tokens:
        if kind != "ref":
            continue
        new = fn(Ref.parse(text))
        if new is None:
            continue
        out.append(body[last:start])
        out.append(new if isinstance(new, XlError) else new.text())
        last = start + len(text)
        changed = True
    if not changed:
        return formula
    out.append(body[last:])
    return "=" + "".join(out)


def shift_refs(formula, formula_sheet, target_sheet, axis, at, delta):
    """Adjust references after inserting (delta > 0) or deleting (delta < 0)
    rows or columns on target_sheet. `at` is the first row/col inserted or
    deleted. References to deleted cells become #REF!."""
    def fn(ref):
        sheet = ref.sheet if ref.sheet is not None else formula_sheet
        if sheet.lower() != target_sheet.lower():
            return None
        lo, hi = (ref.r1, ref.r2) if axis == "row" else (ref.c1, ref.c2)
        if delta > 0:
            nlo = lo + delta if lo >= at else lo
            nhi = hi + delta if hi >= at else hi
        else:
            gone_lo, gone_hi = at, at - delta - 1
            def move(x):
                return x + delta if x > gone_hi else x
            if lo >= gone_lo and hi <= gone_hi:
                return REF
            nlo = gone_lo if gone_lo <= lo <= gone_hi else move(lo)
            nhi = gone_lo - 1 if gone_lo <= hi <= gone_hi else move(hi)
        if (nlo, nhi) == (lo, hi):
            return None
        new = Ref(ref.sheet, ref.c1, ref.r1, ref.c2, ref.r2, ref.abs, ref.is_range)
        if axis == "row":
            new.r1, new.r2 = nlo, nhi
        else:
            new.c1, new.c2 = nlo, nhi
        return new
    return _rewrite_refs(formula, fn)


def rename_sheet_refs(formula, old, new_name):
    def fn(ref):
        if ref.sheet is None or ref.sheet.lower() != old.lower():
            return None
        return Ref(new_name, ref.c1, ref.r1, ref.c2, ref.r2, ref.abs, ref.is_range)
    return _rewrite_refs(formula, fn)


def translate(formula, drow, dcol):
    """Shift relative references, as when a formula is copied elsewhere."""
    def fn(ref):
        a = ref.abs
        new = Ref(ref.sheet,
                  ref.c1 if a[0] else ref.c1 + dcol, ref.r1 if a[1] else ref.r1 + drow,
                  ref.c2 if a[2] else ref.c2 + dcol, ref.r2 if a[3] else ref.r2 + drow,
                  a, ref.is_range)
        if min(new.c1, new.c2, new.r1, new.r2) < 1:
            return REF
        return new
    return _rewrite_refs(formula, fn)


def references(formula):
    """All Ref objects in a formula (for highlighting)."""
    if not is_formula(formula):
        return []
    try:
        return [Ref.parse(t) for k, t, _ in tokenize(formula[1:]) if k == "ref"]
    except SyntaxError:
        return []
