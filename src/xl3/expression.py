"""Tiny expression language for `{{ ... }}` template bodies.

Grammar (XTL 0.1):

    expr      := or_expr
    or_expr   := and_expr (("&") and_expr)*           ; concatenation
    and_expr  := cmp_expr
    cmp_expr  := add_expr ((cmp_op) add_expr)?        ; =, !=, >, <, >=, <=
    add_expr  := mul_expr (("+"|"-") mul_expr)*
    mul_expr  := unary    (("*"|"/") unary)*
    unary     := ("-")? primary
    primary   := number | string | bracket_ref | source_ref | "TRUE"/"FALSE" | call | "(" expr ")"
    bracket_ref := "[" identifier "]"                 ; active source column
    source_ref  := identifier "[" identifier "]"      ; named source column (incl. __sheet__[key])
    call        := identifier "(" args? ")"           ; function call (case-insensitive)

Function names are case-insensitive; the language elsewhere is case-sensitive.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Any, Union


# ---------------------------------------------------------------------------
# AST nodes
# ---------------------------------------------------------------------------


@dataclass
class NumberLit:
    value: float


@dataclass
class StringLit:
    value: str


@dataclass
class BoolLit:
    value: bool


@dataclass
class BracketRef:
    """`[Column]` — the active source's current row's column."""

    column: str


@dataclass
class SourceRef:
    """`SourceName[Column]` — named source / reserved-sheet reference.

    Reserved-sheet refs use `__sheet__[key]` form: source="__inputs__",
    column="month" for `__inputs__[month]`.
    """

    source: str
    column: str


@dataclass
class FuncCall:
    """A function call. `name` is upper-cased on parse for case-insensitivity."""

    name: str
    args: list["Expr"]


@dataclass
class BinOp:
    op: str  # one of: & + - * / = != > < >= <=
    left: "Expr"
    right: "Expr"


@dataclass
class UnaryNeg:
    operand: "Expr"


Expr = Union[
    NumberLit, StringLit, BoolLit, BracketRef, SourceRef, FuncCall, BinOp, UnaryNeg
]


# ---------------------------------------------------------------------------
# Tokenizer
# ---------------------------------------------------------------------------


@dataclass
class Token:
    kind: str  # NUM STR IDENT LBRK RBRK LPAREN RPAREN COMMA OP
    value: Any
    pos: int


_NUM_RE = re.compile(r"\d+(?:\.\d+)?(?:[eE][+-]?\d+)?")
_IDENT_RE = re.compile(r"[A-Za-z_][A-Za-z0-9_]*")
_OPS_2CH = {"!=", ">=", "<="}
_OPS_1CH = set("&+-*/=<>")


class ExpressionParseError(Exception):
    """Raised when an expression body cannot be parsed."""


def tokenize(s: str) -> list[Token]:
    tokens: list[Token] = []
    i = 0
    n = len(s)
    while i < n:
        c = s[i]
        if c.isspace():
            i += 1
            continue
        if c == '"':
            # String literal; XTL does not specify escapes — keep simple.
            j = i + 1
            buf: list[str] = []
            while j < n and s[j] != '"':
                buf.append(s[j])
                j += 1
            if j >= n:
                raise ExpressionParseError(f"unterminated string starting at {i}")
            tokens.append(Token("STR", "".join(buf), i))
            i = j + 1
            continue
        if c == "[":
            tokens.append(Token("LBRK", "[", i))
            i += 1
            continue
        if c == "]":
            tokens.append(Token("RBRK", "]", i))
            i += 1
            continue
        if c == "(":
            tokens.append(Token("LPAREN", "(", i))
            i += 1
            continue
        if c == ")":
            tokens.append(Token("RPAREN", ")", i))
            i += 1
            continue
        if c == ",":
            tokens.append(Token("COMMA", ",", i))
            i += 1
            continue
        m = _NUM_RE.match(s, i)
        if m:
            tokens.append(Token("NUM", float(m.group(0)), i))
            i = m.end()
            continue
        m = _IDENT_RE.match(s, i)
        if m:
            tokens.append(Token("IDENT", m.group(0), i))
            i = m.end()
            continue
        # 2-char op
        if i + 1 < n and s[i : i + 2] in _OPS_2CH:
            tokens.append(Token("OP", s[i : i + 2], i))
            i += 2
            continue
        if c in _OPS_1CH:
            tokens.append(Token("OP", c, i))
            i += 1
            continue
        raise ExpressionParseError(f"unexpected character {c!r} at {i}")
    return tokens


# ---------------------------------------------------------------------------
# Parser (recursive descent)
# ---------------------------------------------------------------------------


@dataclass
class _Parser:
    tokens: list[Token]
    pos: int = 0

    def peek(self, k: int = 0) -> Token | None:
        if self.pos + k < len(self.tokens):
            return self.tokens[self.pos + k]
        return None

    def eat(self) -> Token:
        t = self.tokens[self.pos]
        self.pos += 1
        return t

    def expect(self, kind: str, value: str | None = None) -> Token:
        t = self.peek()
        if t is None or t.kind != kind or (value is not None and t.value != value):
            raise ExpressionParseError(f"expected {kind} {value or ''} at {t}")
        return self.eat()

    # entry
    def parse_expr(self) -> Expr:
        e = self._parse_concat()
        if self.peek() is not None:
            raise ExpressionParseError(f"trailing tokens at {self.peek()}")
        return e

    def _parse_concat(self) -> Expr:
        left = self._parse_compare()
        while True:
            t = self.peek()
            if t and t.kind == "OP" and t.value == "&":
                self.eat()
                right = self._parse_compare()
                left = BinOp("&", left, right)
            else:
                return left

    def _parse_compare(self) -> Expr:
        left = self._parse_add()
        t = self.peek()
        if t and t.kind == "OP" and t.value in {"=", "!=", ">", "<", ">=", "<="}:
            op = self.eat().value
            right = self._parse_add()
            return BinOp(op, left, right)
        return left

    def _parse_add(self) -> Expr:
        left = self._parse_mul()
        while True:
            t = self.peek()
            if t and t.kind == "OP" and t.value in {"+", "-"}:
                op = self.eat().value
                right = self._parse_mul()
                left = BinOp(op, left, right)
            else:
                return left

    def _parse_mul(self) -> Expr:
        left = self._parse_unary()
        while True:
            t = self.peek()
            if t and t.kind == "OP" and t.value in {"*", "/"}:
                op = self.eat().value
                right = self._parse_unary()
                left = BinOp(op, left, right)
            else:
                return left

    def _parse_unary(self) -> Expr:
        t = self.peek()
        if t and t.kind == "OP" and t.value == "-":
            self.eat()
            return UnaryNeg(self._parse_unary())
        return self._parse_primary()

    def _parse_primary(self) -> Expr:
        t = self.peek()
        if t is None:
            raise ExpressionParseError("unexpected end of expression")
        if t.kind == "NUM":
            self.eat()
            return NumberLit(t.value)
        if t.kind == "STR":
            self.eat()
            return StringLit(t.value)
        if t.kind == "LBRK":
            self.eat()
            ident = self.expect("IDENT").value
            self.expect("RBRK")
            return BracketRef(ident.strip())
        if t.kind == "LPAREN":
            self.eat()
            e = self._parse_concat()
            self.expect("RPAREN")
            return e
        if t.kind == "IDENT":
            ident = self.eat().value
            nxt = self.peek()
            # Source-prefixed bracket: `Name[Column]`
            if nxt and nxt.kind == "LBRK":
                self.eat()
                col = self.expect("IDENT").value
                self.expect("RBRK")
                return SourceRef(ident, col.strip())
            # Function call
            if nxt and nxt.kind == "LPAREN":
                self.eat()
                args: list[Expr] = []
                if self.peek() and self.peek().kind != "RPAREN":  # type: ignore[union-attr]
                    args.append(self._parse_concat())
                    while self.peek() and self.peek().kind == "COMMA":  # type: ignore[union-attr]
                        self.eat()
                        args.append(self._parse_concat())
                self.expect("RPAREN")
                return FuncCall(ident.upper(), args)
            # Bare identifiers — TRUE/FALSE first.
            up = ident.upper()
            if up == "TRUE":
                return BoolLit(True)
            if up == "FALSE":
                return BoolLit(False)
            # Per language.md "Group Keys", a bare identifier in a sheet
            # name or file pattern resolves to the active source row's
            # column. Several conformance fixtures (006, 007, 049, 085)
            # rely on the same form working in cell bodies, so we accept
            # it uniformly. (See PORTING_NOTES.md #6.)
            return BracketRef(ident)
        raise ExpressionParseError(f"unexpected token {t}")


def parse_expression(body: str) -> Expr:
    """Parse an expression body (the contents inside `{{ ... }}`)."""
    return _Parser(tokenize(body.strip())).parse_expr()


def parse_filename_or_sheet_expression(body: str) -> Expr:
    """Like `parse_expression`, but allows bare identifiers as group-key
    references (per `language.md` §"Group Keys"). A bare `Customer` is
    rewritten to `BracketRef("Customer")`.
    """
    body = body.strip()
    # If the body is exactly a single identifier, treat as group-key reference.
    if re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", body):
        # Allow TRUE/FALSE to remain as Booleans.
        if body.upper() in ("TRUE", "FALSE"):
            return BoolLit(body.upper() == "TRUE")
        return BracketRef(body)
    return parse_expression(body)


# ---------------------------------------------------------------------------
# Template-cell parse: `Hello {{ x }} world {{ y }}` → list of segments
# ---------------------------------------------------------------------------


@dataclass
class TextSegment:
    text: str


@dataclass
class ExprSegment:
    expr: Expr
    body: str  # the original `{{ ... }}` body, for error messages


@dataclass
class DirectiveSegment:
    """A `{{ @... }}` body — directives are parsed lazily by the directive
    module. We keep the raw body here so the parser can dispatch later
    without forcing a circular import."""

    body: str


CellSegment = Union[TextSegment, ExprSegment, DirectiveSegment]


_EXPR_RE = re.compile(r"\{\{(.*?)\}\}", re.DOTALL)


@dataclass
class CellTemplate:
    segments: list[CellSegment] = field(default_factory=list)

    @property
    def is_pure_text(self) -> bool:
        return all(isinstance(s, TextSegment) for s in self.segments)

    @property
    def is_single_expression(self) -> bool:
        """A cell whose ENTIRE content is exactly one `{{ ... }}` expression."""
        return len(self.segments) == 1 and isinstance(self.segments[0], ExprSegment)

    @property
    def is_directive_cell(self) -> bool:
        """A cell whose ENTIRE content is exactly one `{{ @... }}` directive."""
        return len(self.segments) == 1 and isinstance(self.segments[0], DirectiveSegment)

    def expressions(self) -> list[Expr]:
        return [s.expr for s in self.segments if isinstance(s, ExprSegment)]


def parse_cell_template(text: str) -> CellTemplate:
    """Split a cell's text into literal/expression/directive segments.

    `text` is the cell's effective text (rich-text concatenation handled by
    the caller). Returns a CellTemplate; cells without `{{ }}` produce a
    single TextSegment. A `{{ @... }}` body becomes a DirectiveSegment.
    """
    segs: list[CellSegment] = []
    i = 0
    for m in _EXPR_RE.finditer(text):
        if m.start() > i:
            segs.append(TextSegment(text[i : m.start()]))
        body = m.group(1)
        if body.lstrip().startswith("@"):
            segs.append(DirectiveSegment(body=body))
        else:
            try:
                expr = parse_expression(body)
            except ExpressionParseError as e:
                raise ExpressionParseError(
                    f"failed to parse {{{{ {body} }}}}: {e}"
                ) from e
            segs.append(ExprSegment(expr=expr, body=body))
        i = m.end()
    if i < len(text):
        segs.append(TextSegment(text[i:]))
    if not segs:
        # Empty cell text is also pure text.
        segs.append(TextSegment(""))
    return CellTemplate(segments=segs)


def collect_referenced_columns(expr: Expr) -> set[str]:
    """All `[Column]` (active-source) references inside `expr`."""
    out: set[str] = set()

    def walk(e: Expr) -> None:
        if isinstance(e, BracketRef):
            out.add(e.column)
        elif isinstance(e, FuncCall):
            for a in e.args:
                walk(a)
        elif isinstance(e, BinOp):
            walk(e.left)
            walk(e.right)
        elif isinstance(e, UnaryNeg):
            walk(e.operand)

    walk(expr)
    return out


_AGGREGATE_NAMES = frozenset({"SUM", "COUNT", "AVERAGE", "AVG", "MIN", "MAX"})
# Source-named refs that are NOT per-row data references.
_RESERVED_NAMESPACE_SOURCES = frozenset({"__inputs__", "__config__"})


def expression_has_per_row_ref(expr: Expr) -> bool:
    """True iff `expr` references the *current* data row (or its joined row).

    Used to classify a template row as a data row vs static row. A SourceRef
    INSIDE an aggregate (`SUM(Renewals[Amount])`) does NOT count — it
    aggregates over a row SET, not the active row. A SourceRef into a
    reserved namespace (`__inputs__[name]`) is constant per render.
    """

    def walk(e: Expr) -> bool:
        if isinstance(e, BracketRef):
            return True
        if isinstance(e, SourceRef):
            return e.source not in _RESERVED_NAMESPACE_SOURCES
        if isinstance(e, FuncCall):
            if e.name in _AGGREGATE_NAMES:
                # arguments are row-set scoped, not per-row
                return False
            if e.name == "XLOOKUP":
                # ADR-0013: args 1 and 2 are array refs (typed SourceRef);
                # they DON'T count as per-row references. Args 0 and 3
                # (lookup_value, fallback) are scalar — recurse normally.
                if e.args and walk(e.args[0]):
                    return True
                if len(e.args) >= 4 and walk(e.args[3]):
                    return True
                return False
            return any(walk(a) for a in e.args)
        if isinstance(e, BinOp):
            return walk(e.left) or walk(e.right)
        if isinstance(e, UnaryNeg):
            return walk(e.operand)
        return False

    return walk(expr)
