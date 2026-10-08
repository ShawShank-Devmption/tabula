"""Formula evaluator with Excel semantics (design.md sections 4.2 and 5.4).

Evaluates ONE formula AST.  It never recurses into other cells: the engine
schedules cells in topological order and the context hands back already
computed values.  That is what removes Revision 1's recursion-depth limit.

Context protocol (implemented by engine.Engine):
    value(sheet, host, col, row)        -> value of one cell (REF error for unknown sheet)
    range_cells(sheet, host, rng)       -> [(col, row, value)] of non-blank cells, row-major,
                                           or an ErrorValue for an unknown sheet
    resolve_name(name_node, host)       -> AST the defined name stands for, or None
"""
from __future__ import annotations

import math
import re
from decimal import ROUND_HALF_UP, Decimal, InvalidOperation

from . import parser as P
from .refs import CellRef, RangeRef
from .values import (BLANK, DIV0, NAME_ERR, NUM_ERR, VALUE_ERR, ErrorValue, compare,
                     is_error, is_number, to_bool, to_number, to_text)


class RangeArg:
    """A reference argument (range or single cell): its non-blank cells, plus enough
    to address any cell in it."""

    def __init__(self, cells, origin, shape, sheet, host):
        self.cells = cells      # [(col, row, value)], blanks omitted
        self.origin = origin    # (c1, r1)
        self.shape = shape      # (rows, cols)
        self.sheet = sheet
        self.host = host

    @property
    def blanks(self) -> int:
        return self.shape[0] * self.shape[1] - len(self.cells)


class Evaluator:
    def __init__(self, ctx):
        self.ctx = ctx

    # ------------------------------------------------------------ expressions
    def evaluate(self, node, host: str):
        if isinstance(node, P.Number):
            return node.value
        if isinstance(node, P.Text):
            return node.value
        if isinstance(node, P.Bool):
            return node.value
        if isinstance(node, P.ErrorLit):
            return ErrorValue(node.code)
        if isinstance(node, P.Ref):
            return self.ctx.value(node.sheet, host, node.ref.col, node.ref.row)
        if isinstance(node, P.RangeNode):
            return VALUE_ERR  # a range where one value is needed
        if isinstance(node, P.Name):
            target = self.ctx.resolve_name(node, host)
            return NAME_ERR if target is None else self.evaluate(target, host)
        if isinstance(node, P.Unary):
            v = self.evaluate(node.operand, host)
            if node.op == "+":
                return v
            n = to_number(v)
            return n if is_error(n) else -n
        if isinstance(node, P.Percent):
            n = to_number(self.evaluate(node.operand, host))
            return n if is_error(n) else n / 100
        if isinstance(node, P.Binary):
            return self._binary(node, host)
        if isinstance(node, P.Call):
            return self._call(node, host)
        return VALUE_ERR

    def _binary(self, node, host):
        left = self.evaluate(node.left, host)
        if is_error(left):
            return left
        right = self.evaluate(node.right, host)
        if is_error(right):
            return right
        op = node.op
        if op == "&":
            return to_text(left) + to_text(right)
        if op in ("=", "<>", "<", "<=", ">", ">="):
            c = compare(left, right)
            return {"=": c == 0, "<>": c != 0, "<": c < 0, "<=": c <= 0,
                    ">": c > 0, ">=": c >= 0}[op]
        a, b = to_number(left), to_number(right)
        if is_error(a):
            return a
        if is_error(b):
            return b
        return _arith(op, a, b)

    # ------------------------------------------------------------ arguments
    def _range_of(self, node, host):
        """RangeArg if node is a reference -- a range or a single cell, directly or via
        a defined name -- else None.  Excel treats both alike in aggregates: text,
        logicals and blanks reached through a reference are ignored."""
        if isinstance(node, P.Name):
            target = self.ctx.resolve_name(node, host)
            if isinstance(target, (P.RangeNode, P.Ref)):
                node = target
        if isinstance(node, P.Ref):
            rng = RangeRef(node.ref.plain(), node.ref.plain())
        elif isinstance(node, P.RangeNode):
            rng = node.rng
        else:
            return None
        cells = self.ctx.range_cells(node.sheet, host, rng)
        if is_error(cells):
            return cells
        c1, r1, _, _ = rng.bounds()
        return RangeArg(cells, (c1, r1), rng.shape(), node.sheet, host)

    def _values(self, args, host):
        """[(value, from_range)] for aggregate-style arguments."""
        out = []
        for a in args:
            rng = self._range_of(a, host)
            if is_error(rng):
                out.append((rng, False))
            elif rng is not None:
                out.extend((v, True) for _, _, v in rng.cells)
            else:
                out.append((self.evaluate(a, host), False))
        return out

    def _numbers(self, args, host):
        """Numbers per aggregate rules: ranges keep only numbers, scalars are coerced."""
        nums = []
        for v, from_range in self._values(args, host):
            if is_error(v):
                return v
            if from_range:
                if is_number(v):
                    nums.append(float(v))
            else:
                n = to_number(v)
                if is_error(n):
                    return n
                nums.append(n)
        return nums

    # ------------------------------------------------------------ functions
    def _call(self, node, host):
        f, args = node.fname, node.args
        handler = _HANDLERS.get(f)
        if handler is None:
            return ErrorValue("#NAME?", f"function {f} is not simulated")
        try:
            return handler(self, args, host)
        except IndexError:  # wrong arity reached runtime (only from loaded files)
            return VALUE_ERR

    def _fn_sum(self, args, host):
        nums = self._numbers(args, host)
        return nums if is_error(nums) else float(sum(nums))

    def _fn_average(self, args, host):
        nums = self._numbers(args, host)
        if is_error(nums):
            return nums
        return DIV0 if not nums else sum(nums) / len(nums)

    def _fn_min(self, args, host):
        nums = self._numbers(args, host)
        return nums if is_error(nums) else (min(nums) if nums else 0.0)

    def _fn_max(self, args, host):
        nums = self._numbers(args, host)
        return nums if is_error(nums) else (max(nums) if nums else 0.0)

    def _fn_count(self, args, host):
        n = 0
        for v, from_range in self._values(args, host):
            if from_range:
                n += is_number(v)
            elif v is not BLANK and not is_error(v) and not is_error(to_number(v)):
                n += 1
        return float(n)

    def _fn_counta(self, args, host):
        return float(sum(1 for v, _ in self._values(args, host) if v is not BLANK))

    def _fn_countif(self, args, host):
        rng = self._range_of(args[0], host)
        if rng is None:
            return VALUE_ERR
        if is_error(rng):
            return rng
        crit = self.evaluate(args[1], host)
        if is_error(crit):
            return crit
        match = _criterion(crit)
        hits = sum(1 for _, _, v in rng.cells if match(v))
        return float(hits + (rng.blanks if match(BLANK) else 0))

    def _fn_sumif(self, args, host):
        rng = self._range_of(args[0], host)
        if rng is None:
            return VALUE_ERR
        if is_error(rng):
            return rng
        crit = self.evaluate(args[1], host)
        if is_error(crit):
            return crit
        sum_rng = rng
        if len(args) == 3:
            sum_rng = self._range_of(args[2], host)
            if sum_rng is None:
                return VALUE_ERR
            if is_error(sum_rng):
                return sum_rng
        match = _criterion(crit)
        total = 0.0
        for col, row, v in rng.cells:
            if not match(v):
                continue
            tc = sum_rng.origin[0] + (col - rng.origin[0])
            tr = sum_rng.origin[1] + (row - rng.origin[1])
            target = self.ctx.value(sum_rng.sheet, host, tc, tr)
            if is_error(target):
                return target
            if is_number(target):
                total += target
        if match(BLANK) and rng.blanks and sum_rng is not rng:
            # blank criterion cells can match (e.g. "<>x"); scan the sum area instead
            occupied = {(c, r) for c, r, _ in rng.cells}
            sx, sy = sum_rng.origin
            h, w = rng.shape
            area = RangeRef(CellRef(sx, sy), CellRef(sx + w - 1, sy + h - 1))
            cells = self.ctx.range_cells(sum_rng.sheet, host, area)
            if is_error(cells):
                return cells
            for c, r, target in cells:
                if (rng.origin[0] + c - sx, rng.origin[1] + r - sy) in occupied:
                    continue
                if is_error(target):
                    return target
                if is_number(target):
                    total += target
        return total

    def _scalars(self, args, host):
        vals = [self.evaluate(a, host) for a in args]
        for v in vals:
            if is_error(v):
                return v
        return vals

    def _fn_round(self, args, host):
        vals = self._scalars(args, host)
        if is_error(vals):
            return vals
        x, d = to_number(vals[0]), to_number(vals[1])
        if is_error(x):
            return x
        if is_error(d):
            return d
        return excel_round(x, int(d))

    def _fn_abs(self, args, host):
        n = to_number(self.evaluate(args[0], host))
        return n if is_error(n) else abs(n)

    def _fn_sqrt(self, args, host):
        n = to_number(self.evaluate(args[0], host))
        if is_error(n):
            return n
        return NUM_ERR if n < 0 else math.sqrt(n)

    def _fn_mod(self, args, host):
        vals = self._scalars(args, host)
        if is_error(vals):
            return vals
        a, b = to_number(vals[0]), to_number(vals[1])
        if is_error(a):
            return a
        if is_error(b):
            return b
        return DIV0 if b == 0 else a - b * math.floor(a / b)

    def _fn_power(self, args, host):
        vals = self._scalars(args, host)
        if is_error(vals):
            return vals
        a, b = to_number(vals[0]), to_number(vals[1])
        if is_error(a):
            return a
        if is_error(b):
            return b
        return _arith("^", a, b)

    def _fn_if(self, args, host):
        cond = to_bool(self.evaluate(args[0], host))
        if is_error(cond):
            return cond
        if cond:
            return self.evaluate(args[1], host)
        return self.evaluate(args[2], host) if len(args) == 3 else False

    def _fn_iferror(self, args, host):
        v = self.evaluate(args[0], host)
        return self.evaluate(args[1], host) if is_error(v) else v

    def _logicals(self, args, host):
        out = []
        for v, from_range in self._values(args, host):
            if is_error(v):
                return v
            if from_range:
                if isinstance(v, bool) or is_number(v):
                    out.append(bool(v))
            elif v is not BLANK:
                b = to_bool(v)
                if is_error(b):
                    return b
                out.append(b)
        return out if out else VALUE_ERR

    def _fn_and(self, args, host):
        vals = self._logicals(args, host)
        return vals if is_error(vals) else all(vals)

    def _fn_or(self, args, host):
        vals = self._logicals(args, host)
        return vals if is_error(vals) else any(vals)

    def _fn_not(self, args, host):
        b = to_bool(self.evaluate(args[0], host))
        return b if is_error(b) else not b

    def _text1(self, args, host):
        return to_text(self.evaluate(args[0], host))

    def _fn_len(self, args, host):
        t = self._text1(args, host)
        return t if is_error(t) else float(len(t))

    def _fn_upper(self, args, host):
        t = self._text1(args, host)
        return t if is_error(t) else t.upper()

    def _fn_lower(self, args, host):
        t = self._text1(args, host)
        return t if is_error(t) else t.lower()

    def _fn_trim(self, args, host):
        t = self._text1(args, host)
        return t if is_error(t) else re.sub(" +", " ", t).strip(" ")

    def _fn_concat(self, args, host):
        parts = []
        for v, _ in self._values(args, host):
            if is_error(v):
                return v
            parts.append(to_text(v))
        return "".join(parts)

    def _fn_isnumber(self, args, host):
        return is_number(self.evaluate(args[0], host))

    def _fn_isblank(self, args, host):
        return self.evaluate(args[0], host) is BLANK

    def _fn_iserror(self, args, host):
        return is_error(self.evaluate(args[0], host))


_HANDLERS = {
    "SUM": Evaluator._fn_sum, "AVERAGE": Evaluator._fn_average,
    "MIN": Evaluator._fn_min, "MAX": Evaluator._fn_max,
    "COUNT": Evaluator._fn_count, "COUNTA": Evaluator._fn_counta,
    "COUNTIF": Evaluator._fn_countif, "SUMIF": Evaluator._fn_sumif,
    "ROUND": Evaluator._fn_round, "ABS": Evaluator._fn_abs, "SQRT": Evaluator._fn_sqrt,
    "MOD": Evaluator._fn_mod, "POWER": Evaluator._fn_power,
    "IF": Evaluator._fn_if, "IFERROR": Evaluator._fn_iferror,
    "AND": Evaluator._fn_and, "OR": Evaluator._fn_or, "NOT": Evaluator._fn_not,
    "LEN": Evaluator._fn_len, "UPPER": Evaluator._fn_upper, "LOWER": Evaluator._fn_lower,
    "TRIM": Evaluator._fn_trim, "CONCAT": Evaluator._fn_concat,
    "CONCATENATE": Evaluator._fn_concat,
    "ISNUMBER": Evaluator._fn_isnumber, "ISBLANK": Evaluator._fn_isblank,
    "ISERROR": Evaluator._fn_iserror,
}


def _arith(op: str, a: float, b: float):
    try:
        if op == "+":
            r = a + b
        elif op == "-":
            r = a - b
        elif op == "*":
            r = a * b
        elif op == "/":
            if b == 0:
                return DIV0
            r = a / b
        elif op == "^":
            if a == 0 and b < 0:
                return DIV0
            if (a == 0 and b == 0) or (a < 0 and b != int(b)):
                return NUM_ERR
            r = float(a ** b)
        else:
            return VALUE_ERR
    except OverflowError:
        return NUM_ERR
    return r if math.isfinite(r) else NUM_ERR  # Excel has no infinities


def excel_round(x: float, digits: int) -> float:
    """ROUND with ties away from zero, on the decimal representation (as Excel does)."""
    try:
        q = Decimal(1).scaleb(-digits)
        return float(Decimal(repr(x)).quantize(q, rounding=ROUND_HALF_UP))
    except InvalidOperation:
        return x


_CRIT_OPS = (">=", "<=", "<>", ">", "<", "=")


def _criterion(crit):
    """Matcher for COUNTIF/SUMIF criteria (number, text with wildcards, or 'op value')."""
    if isinstance(crit, bool):
        return lambda v: isinstance(v, bool) and v == crit
    if is_number(crit):
        return lambda v: is_number(v) and v == crit
    text = to_text(crit)  # a blank criterion behaves like ""
    op = "="
    for candidate in _CRIT_OPS:
        if text.startswith(candidate):
            op, text = candidate, text[len(candidate):]
            break
    num = to_number(text) if text.strip() else VALUE_ERR
    if not is_error(num):
        def numeric(v):
            if not is_number(v):
                return op == "<>"
            c = (v > num) - (v < num)
            return {"=": c == 0, "<>": c != 0, ">": c > 0, ">=": c >= 0,
                    "<": c < 0, "<=": c <= 0}[op]
        return numeric
    if op in ("=", "<>"):
        if text == "":  # "" / "=" match empty cells; "<>" matches non-empty ones
            hit = lambda v: v is BLANK or v == ""
        else:
            pattern = re.compile(_wildcard_regex(text), re.IGNORECASE | re.DOTALL)
            hit = lambda v: isinstance(v, str) and pattern.fullmatch(v) is not None
        return hit if op == "=" else (lambda v: not hit(v))

    def textual(v):
        if not isinstance(v, str):
            return False
        c = compare(v, text)
        return {">": c > 0, ">=": c >= 0, "<": c < 0, "<=": c <= 0}[op]
    return textual


def _wildcard_regex(text: str) -> str:
    out, i = [], 0
    while i < len(text):
        ch = text[i]
        if ch == "~" and i + 1 < len(text):
            out.append(re.escape(text[i + 1]))
            i += 2
            continue
        out.append(".*" if ch == "*" else "." if ch == "?" else re.escape(ch))
        i += 1
    return "".join(out)
