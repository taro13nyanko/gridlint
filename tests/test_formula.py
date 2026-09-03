"""Unit tests for the tokenizer, parser and evaluator, independent of any workbook."""
from __future__ import annotations

import pytest

from gridlint.formula.ast_nodes import BinOp, FuncCall, Literal, RangeRef, Ref, UnaryOp
from gridlint.formula.evaluator import evaluate
from gridlint.formula.functions import Matrix
from gridlint.formula.parser import col_to_index, index_to_col, parse_formula
from gridlint.formula.tokenizer import FormulaSyntaxError, TokType, tokenize
from gridlint.formula.values import BLANK, ExcelError, compare, excel_round, general_format, to_number


class FakeResolver:
    """A tiny grid so evaluator tests do not need a real file."""

    def __init__(self, cells=None, sheets=("Sheet1", "Data")):
        self.cells = cells or {}
        self.sheets = set(sheets)

    def cell(self, sheet, col, row):
        return self.cells.get((sheet, col, row), BLANK)

    def used_bounds(self, sheet):
        cols = [k[1] for k in self.cells if k[0] == sheet] or [1]
        rows = [k[2] for k in self.cells if k[0] == sheet] or [1]
        return max(cols), max(rows)

    def sheet_exists(self, sheet):
        return sheet in self.sheets

    def defined_name(self, name):
        return None


def ev(formula, cells=None):
    return evaluate(parse_formula(formula), "Sheet1", FakeResolver(cells))


# --------------------------------------------------------------------------- tokenizer
def test_tokenizer_classifies_refs_ranges_and_functions():
    toks = tokenize("=SUM(A1:B2, Sheet1!C3, 'My Sheet'!D4)*2")
    kinds = [t.type for t in toks]
    assert TokType.FUNC in kinds and TokType.RANGE in kinds and TokType.REF in kinds
    assert [t.value for t in toks if t.type is TokType.REF] == ["Sheet1!C3", "My Sheet!D4"]


def test_tokenizer_handles_escaped_quotes_and_errors():
    toks = tokenize('="he said ""hi"""&#REF!')
    assert toks[0].value == 'he said "hi"'
    assert toks[2].type is TokType.ERROR


@pytest.mark.parametrize("bad", ["=SUM(", '="unterminated', "=@@@", "={1,2}"])
def test_tokenizer_or_parser_rejects_malformed(bad):
    with pytest.raises(FormulaSyntaxError):
        parse_formula(bad)


def test_column_letters_round_trip():
    for i in (1, 26, 27, 52, 703, 16384):
        assert col_to_index(index_to_col(i)) == i


# --------------------------------------------------------------------------- parser
def test_unary_minus_binds_tighter_than_power():
    # Excel: -2^2 = 4. This is the single most commonly mis-implemented rule.
    assert ev("=-2^2") == 4.0


def test_power_is_left_associative():
    assert ev("=2^3^2") == 64.0        # (2^3)^2, not 2^(3^2)


def test_precedence_and_parens():
    assert ev("=2+3*4-6/3") == 12.0
    assert ev("=((2+3)*(4-1))/5") == 3.0


def test_percent_is_postfix():
    assert ev("=50%*40") == 20.0
    assert ev("=(10+10)%") == pytest.approx(0.2)


def test_absolute_and_relative_refs_parse_the_same_cell():
    for raw in ("A1", "$A1", "A$1", "$A$1"):
        node = parse_formula("=" + raw)
        assert isinstance(node, Ref) and (node.col, node.row) == (1, 1)


def test_empty_argument_is_allowed():
    node = parse_formula("=IF(A1,,2)")
    assert isinstance(node, FuncCall) and len(node.args) == 3


def test_whole_column_range_parses():
    node = parse_formula("=SUM(B:D)")
    rng = node.args[0]
    assert isinstance(rng, RangeRef) and (rng.col1, rng.col2) == (2, 4)


# --------------------------------------------------------------------------- values
def test_excel_rounds_half_away_from_zero_not_bankers():
    assert excel_round(0.5) == 1.0
    assert excel_round(-0.5) == -1.0
    assert excel_round(2.5) == 3.0            # Python's round(2.5) is 2
    assert excel_round(2.675, 2) == 2.68      # binary representation trap


def test_general_format_matches_excel_text_conversion():
    assert general_format(1234567.0) == "1234567"
    assert general_format(0.25) == "0.25"
    assert general_format(-3.0) == "-3"


def test_cross_type_ordering_number_then_text_then_boolean():
    assert compare(1, "a") < 0
    assert compare("z", True) < 0
    assert compare("ABC", "abc") == 0         # text comparison is case-insensitive


def test_blank_coerces_to_zero_and_empty_string():
    assert to_number(BLANK) == 0.0
    assert ev("=Z9+5") == 5.0
    assert ev('=Z9&"x"') == "x"


# --------------------------------------------------------------------------- evaluator
def test_division_by_zero_returns_error_value():
    r = ev("=1/0")
    assert isinstance(r, ExcelError) and r.code == "#DIV/0!"


def test_errors_propagate_through_arithmetic_and_sum():
    assert isinstance(ev("=1+1/0"), ExcelError)
    assert isinstance(ev("=SUM(1,1/0)"), ExcelError)


def test_if_does_not_evaluate_the_branch_it_does_not_take():
    assert ev("=IF(TRUE,1,1/0)") == 1.0
    assert ev("=IF(FALSE,1/0,2)") == 2.0


def test_iferror_catches_and_passes_through():
    assert ev('=IFERROR(1/0,"caught")') == "caught"
    assert ev("=IFERROR(6/2,0)") == 3.0


def test_sum_ignores_text_in_a_range_but_coerces_a_direct_argument():
    cells = {("Sheet1", 1, r): v for r, v in enumerate([1.0, "text", 3.0], start=1)}
    assert evaluate(parse_formula("=SUM(A1:A3)"), "Sheet1", FakeResolver(cells)) == 4.0
    assert ev('=SUM("3",4)') == 7.0


def test_count_counts_numbers_counta_counts_anything():
    cells = {("Sheet1", 1, 1): 1.0, ("Sheet1", 1, 2): "x", ("Sheet1", 1, 3): 3.0}
    r = FakeResolver(cells)
    assert evaluate(parse_formula("=COUNT(A1:A3)"), "Sheet1", r) == 2.0
    assert evaluate(parse_formula("=COUNTA(A1:A3)"), "Sheet1", r) == 3.0


def test_countif_and_sumif_criteria_forms():
    cells = {}
    for i, (region, amount) in enumerate([("North", 10.0), ("South", 20.0), ("North", 30.0)], start=1):
        cells[("Sheet1", 1, i)] = region
        cells[("Sheet1", 2, i)] = amount
    r = FakeResolver(cells)
    assert evaluate(parse_formula('=COUNTIF(A1:A3,"North")'), "Sheet1", r) == 2.0
    assert evaluate(parse_formula('=COUNTIF(B1:B3,">15")'), "Sheet1", r) == 2.0
    assert evaluate(parse_formula('=COUNTIF(A1:A3,"N*")'), "Sheet1", r) == 2.0
    assert evaluate(parse_formula('=SUMIF(A1:A3,"North",B1:B3)'), "Sheet1", r) == 40.0


def test_vlookup_exact_and_approximate():
    cells = {}
    for i, (k, v) in enumerate([(0.0, "tiny"), (10.0, "small"), (30.0, "mid")], start=1):
        cells[("Sheet1", 1, i)] = k
        cells[("Sheet1", 2, i)] = v
    r = FakeResolver(cells)
    assert evaluate(parse_formula("=VLOOKUP(20,A1:B3,2,TRUE)"), "Sheet1", r) == "small"
    assert evaluate(parse_formula("=VLOOKUP(10,A1:B3,2,FALSE)"), "Sheet1", r) == "small"
    assert isinstance(evaluate(parse_formula("=VLOOKUP(99,A1:B3,2,FALSE)"), "Sheet1", r), ExcelError)


def test_reference_to_missing_sheet_is_a_ref_error():
    r = ev("=Nope!A1+1")
    assert isinstance(r, ExcelError) and r.code == "#REF!"


def test_unknown_function_is_a_name_error():
    r = ev("=XLOOKUP(1,A1:A2,B1:B2)")
    assert isinstance(r, ExcelError) and r.code == "#NAME?"


def test_text_functions():
    assert ev('=LEFT("hello",3)&RIGHT("hello",2)') == "hello"
    assert ev('=TRIM("  a   b  ")') == "a b"
    assert ev('=TEXT(0.1234,"0.0%")') == "12.3%"


def test_concatenation_uses_general_number_format():
    assert ev("=1/4&\"\"") == "0.25"
