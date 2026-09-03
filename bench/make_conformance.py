"""Generate the conformance workbook with Excel itself, so its cached values are ground truth.

Run on a Windows machine with Excel installed:

    python bench/make_conformance.py

It writes samples/conformance.xlsx. `tests/test_conformance.py` then asserts that
Gridlint's engine reproduces every cached value. Excel is only needed to *create*
the fixture; the committed .xlsx makes the test runnable anywhere.
"""
from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
OUT = ROOT / "samples" / "conformance.xlsx"

# Each case is a formula placed in column B; column A holds a label.
# Anything Excel can compute, the engine must reproduce exactly.
CASES: list[tuple[str, str]] = [
    # --- operator precedence and Excel's two famous surprises -----------------
    ("unary minus binds tighter than ^", "=-2^2"),
    ("caret is left associative", "=2^3^2"),
    ("percent postfix", "=50%*40"),
    ("percent on parenthesis", "=(10+10)%"),
    ("mixed precedence", "=2+3*4-6/3"),
    ("comparison yields boolean", "=3>2"),
    ("comparison chain via AND", "=AND(3>2,2>1)"),
    ("negative exponent", "=2^-2"),
    ("nested parens", "=((2+3)*(4-1))/5"),
    ("unary minus twice", "=--5"),
    # --- numeric coercion ----------------------------------------------------
    ("text number adds", '="7"+1'),
    ("bool adds as one", "=TRUE+1"),
    ("blank cell adds as zero", "=Z100+5"),
    ("blank cell concatenates as empty", '=Z100&"x"'),
    ("number to text general format", "=1/4&\"\""),
    ("big number to text", "=1234567&\"\""),
    ("text compare is case insensitive", '="abc"="ABC"'),
    ("text sorts after numbers", '=("a">1)'),
    ("true sorts after text", '=(TRUE>"z")'),
    # --- errors --------------------------------------------------------------
    ("divide by zero", "=1/0"),
    ("iferror catches div0", '=IFERROR(1/0,"caught")'),
    ("iferror passes through", "=IFERROR(6/2,0)"),
    ("error propagates through sum", "=SUM(1,1/0)"),
    ("iserror true", "=ISERROR(1/0)"),
    ("iserror false", "=ISERROR(1)"),
    ("value error from text", '=1+"abc"'),
    ("if short circuits error branch", '=IF(TRUE,1,1/0)'),
    ("na literal", "=IFERROR(#N/A,\"na\")"),
    # --- rounding ------------------------------------------------------------
    ("round half away from zero", "=ROUND(0.5,0)"),
    ("round half away negative", "=ROUND(-0.5,0)"),
    ("round 2.5 is 3 not 2", "=ROUND(2.5,0)"),
    ("round to two places", "=ROUND(2.675,2)"),
    ("round negative digits", "=ROUND(12345,-2)"),
    ("roundup", "=ROUNDUP(2.001,2)"),
    ("rounddown", "=ROUNDDOWN(2.999,2)"),
    ("int floors negatives", "=INT(-2.1)"),
    ("abs", "=ABS(-7.5)"),
    ("sqrt", "=SQRT(16)"),
    ("power", "=POWER(3,4)"),
    # --- aggregates over ranges ---------------------------------------------
    ("sum range", "=SUM($E$2:$E$11)"),
    ("sum ignores text in range", "=SUM($F$2:$F$11)"),
    ("sum coerces direct text arg", '=SUM("3",4)'),
    ("average range", "=AVERAGE($E$2:$E$11)"),
    ("average ignores text", "=AVERAGE($F$2:$F$11)"),
    ("min range", "=MIN($E$2:$E$11)"),
    ("max range", "=MAX($E$2:$E$11)"),
    ("count numbers only", "=COUNT($F$2:$F$11)"),
    ("counta non empty", "=COUNTA($F$2:$F$11)"),
    ("countblank", "=COUNTBLANK($F$2:$F$11)"),
    ("product", "=PRODUCT(2,3,4)"),
    ("sum of empty range is zero", "=SUM($Y$50:$Y$60)"),
    ("nested aggregate", "=SUM(MAX($E$2:$E$11),MIN($E$2:$E$11))"),
    ("sum with mixed args", "=SUM($E$2:$E$5,100,$E$6)"),
    # --- conditional aggregates ---------------------------------------------
    ("countif greater", '=COUNTIF($E$2:$E$11,">50")'),
    ("countif equal number", "=COUNTIF($E$2:$E$11,40)"),
    ("countif text", '=COUNTIF($D$2:$D$11,"north")'),
    ("countif not equal", '=COUNTIF($D$2:$D$11,"<>north")'),
    ("countif wildcard", '=COUNTIF($D$2:$D$11,"n*")'),
    ("sumif text criteria", '=SUMIF($D$2:$D$11,"north",$E$2:$E$11)'),
    ("sumif numeric criteria", '=SUMIF($E$2:$E$11,">50")'),
    ("sumif separate range", '=SUMIF($D$2:$D$11,"south",$E$2:$E$11)'),
    ("averageif", '=AVERAGEIF($D$2:$D$11,"north",$E$2:$E$11)'),
    # --- logic ---------------------------------------------------------------
    ("if true branch", '=IF(1=1,"yes","no")'),
    ("if false branch", '=IF(1=2,"yes","no")'),
    ("nested if", '=IF($E$2>100,"big",IF($E$2>10,"mid","small"))'),
    ("and or not", "=AND(OR(FALSE,TRUE),NOT(FALSE))"),
    ("if returns number", "=IF(TRUE,42,0)"),
    ("isblank on empty", "=ISBLANK($Z$100)"),
    ("isnumber", "=ISNUMBER($E$2)"),
    ("istext", "=ISTEXT($D$2)"),
    # --- text ----------------------------------------------------------------
    ("concat operator", '="a"&"b"&"c"'),
    ("concatenate function", '=CONCATENATE("x",1,"y")'),
    ("len", '=LEN("hello")'),
    ("left", '=LEFT("hello",3)'),
    ("right", '=RIGHT("hello",2)'),
    ("mid", '=MID("hello",2,3)'),
    ("trim", '=TRIM("  a   b  ")'),
    ("upper lower", '=UPPER("ab")&LOWER("CD")'),
    ("value of text", '=VALUE("12.5")+1'),
    ("text two decimals", '=TEXT(3.14159,"0.00")'),
    ("text percent", '=TEXT(0.1234,"0.0%")'),
    ("number joined to text", '=ROUND(1/3,4)&" done"'),
    # --- lookup --------------------------------------------------------------
    ("vlookup exact", '=VLOOKUP("c",$H$2:$I$6,2,FALSE)'),
    ("vlookup approximate", "=VLOOKUP(35,$K$2:$L$6,2,TRUE)"),
    ("match exact", '=MATCH("c",$H$2:$H$6,0)'),
    ("index row", "=INDEX($E$2:$E$11,3)"),
    ("index match", '=INDEX($I$2:$I$6,MATCH("d",$H$2:$H$6,0))'),
    # --- references ----------------------------------------------------------
    ("absolute ref", "=$E$2*2"),
    ("mixed ref", "=E$2+$E3"),
    ("cross sheet ref", "=Data!B2*10"),
    ("cross sheet range", "=SUM(Data!B2:B5)"),
    ("quoted sheet name", "='Odd Sheet'!B2+1"),
    ("chain of refs", "=B2+B3+B4"),
    # --- multi-criteria aggregates, which real models are built from ----------
    ("sumifs one condition", '=SUMIFS($E$2:$E$11,$D$2:$D$11,"North")'),
    ("sumifs two conditions", '=SUMIFS($E$2:$E$11,$D$2:$D$11,"North",$E$2:$E$11,">50")'),
    ("sumifs no match", '=SUMIFS($E$2:$E$11,$D$2:$D$11,"Nowhere")'),
    ("countifs one condition", '=COUNTIFS($D$2:$D$11,"South")'),
    ("countifs two conditions", '=COUNTIFS($D$2:$D$11,"North",$E$2:$E$11,"<80")'),
    ("countifs numeric range", '=COUNTIFS($E$2:$E$11,">=40",$E$2:$E$11,"<=90")'),
    ("averageifs", '=AVERAGEIFS($E$2:$E$11,$D$2:$D$11,"North")'),
    ("maxifs", '=MAXIFS($E$2:$E$11,$D$2:$D$11,"South")'),
    ("minifs", '=MINIFS($E$2:$E$11,$D$2:$D$11,"North")'),
    ("sumproduct two arrays", "=SUMPRODUCT($E$2:$E$6,$I$2:$I$6)"),
    ("sumproduct as a condition", '=SUMPRODUCT(($D$2:$D$11="North")*$E$2:$E$11)'),
    ("sumproduct single array", "=SUMPRODUCT($E$2:$E$11)"),
    # --- logic and lookup ----------------------------------------------------
    ("ifs first match", '=IFS($E$2>100,"big",$E$2>10,"mid",TRUE,"small")'),
    ("ifs later match", '=IFS($E$3>100,"big",$E$3>10,"mid",TRUE,"small")'),
    ("switch matched", '=SWITCH(2,1,"one",2,"two","other")'),
    ("switch default", '=SWITCH(9,1,"one",2,"two","other")'),
    ("choose", '=CHOOSE(2,"a","b","c")'),
    ("xlookup exact", '=XLOOKUP("c",$H$2:$H$6,$I$2:$I$6)'),
    ("xlookup not found", '=XLOOKUP("zz",$H$2:$H$6,$I$2:$I$6,"none")'),
    ("xlookup next larger", "=XLOOKUP(35,$K$2:$K$6,$L$2:$L$6,\"none\",1)"),
    ("hlookup exact", '=HLOOKUP("q2",$N$1:$Q$2,2,FALSE)'),
    ("isna on lookup", '=ISNA(VLOOKUP("zz",$H$2:$I$6,2,FALSE))'),
    ("na literal caught", "=IFNA(NA(),0)"),
    # --- dates ---------------------------------------------------------------
    ("date builds a serial", "=DATE(2026,9,3)*1"),
    ("year month day", "=YEAR($M$2)*10000+MONTH($M$2)*100+DAY($M$2)"),
    ("date arithmetic", "=$M$3-$M$2"),
    ("eomonth this month", "=EOMONTH($M$2,0)-$M$2"),
    ("eomonth previous", "=EOMONTH($M$2,-1)*1"),
    ("edate plus three", "=EDATE($M$2,3)*1"),
    ("days between", "=DAYS($M$3,$M$2)"),
    ("weekday default", "=WEEKDAY($M$2)"),
    ("weekday monday one", "=WEEKDAY($M$2,2)"),
    ("datedif months", '=DATEDIF($M$2,$M$3,"m")'),
    ("date rolls month overflow", "=DATE(2026,14,1)*1"),
    # --- finance -------------------------------------------------------------
    ("npv", "=ROUND(NPV(0.1,$E$2:$E$6),6)"),
    ("npv plus initial", "=ROUND(NPV(0.08,$E$2:$E$6)-1000,6)"),
    ("irr", "=ROUND(IRR($J$2:$J$7),6)"),
    ("pmt", "=ROUND(PMT(0.05/12,360,-300000),6)"),
    ("pmt zero rate", "=ROUND(PMT(0,120,-12000),6)"),
    ("pv", "=ROUND(PV(0.06/12,120,-500),6)"),
    ("fv", "=ROUND(FV(0.06/12,120,-500),6)"),
    # --- maths ---------------------------------------------------------------
    ("mod positive", "=MOD(17,5)"),
    ("mod negative dividend", "=MOD(-17,5)"),
    ("mod negative divisor", "=MOD(17,-5)"),
    ("ceiling", "=CEILING(23,10)"),
    ("ceiling exact", "=CEILING(20,10)"),
    ("floor", "=FLOOR(27,10)"),
    ("sign", "=SIGN(-4)+SIGN(0)+SIGN(9)"),
    ("median odd", "=MEDIAN($E$2:$E$10)"),
    ("median even", "=MEDIAN($E$2:$E$11)"),
    ("stdev sample", "=ROUND(STDEV($E$2:$E$11),6)"),
    ("large second", "=LARGE($E$2:$E$11,2)"),
    ("small second", "=SMALL($E$2:$E$11,2)"),
    # --- text ----------------------------------------------------------------
    ("textjoin skipping blanks", '=TEXTJOIN("-",TRUE,$H$2:$H$6)'),
    ("textjoin keeping blanks", '=TEXTJOIN("-",FALSE,$D$2:$D$4)'),
    ("substitute all", '=SUBSTITUTE("a-b-a","a","X")'),
    ("substitute nth", '=SUBSTITUTE("a-b-a","a","X",2)'),
    ("replace", '=REPLACE("abcdef",2,3,"ZZ")'),
    ("find is case sensitive", '=FIND("B","aBcB")'),
    ("search is not", '=SEARCH("b","aBcB")'),
    ("proper", '=PROPER("hello world")'),
    ("rept", '=REPT("ab",3)'),
    ("len of joined", '=LEN(TEXTJOIN(",",TRUE,$H$2:$H$6))'),
]

# Support data placed on the Calc sheet.
E_VALUES = [120, 40, 75, 10, 55, 90, 5, 65, 30, 20]
F_VALUES = [1, "text", 3, None, 5, "x", 7, None, 9, 10]
D_VALUES = ["North", "South", "North", "East", "South", "North", "West", "South", "North", "East"]
H_VALUES = ["a", "b", "c", "d", "e"]
I_VALUES = [10, 20, 30, 40, 50]
K_VALUES = [0, 10, 30, 60, 100]
L_VALUES = ["tiny", "small", "mid", "large", "huge"]
#: A cash-flow series with a sign change, so IRR has a root to find.
J_VALUES = [-1000, 250, 300, 350, 400, 450]
#: Two dates, so the date functions have something real to work on.
M_DATES = ["2026-03-15", "2027-01-04"]
#: A horizontal table for HLOOKUP.
HLOOKUP_HEADERS = ["q1", "q2", "q3", "q4"]
HLOOKUP_VALUES = [11, 22, 33, 44]


def main() -> int:
    import win32com.client as win32

    app = win32.Dispatch("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        wb = app.Workbooks.Add()
        while wb.Worksheets.Count < 3:
            wb.Worksheets.Add()
        calc = wb.Worksheets(1)
        calc.Name = "Calc"
        data = wb.Worksheets(2)
        data.Name = "Data"
        odd = wb.Worksheets(3)
        odd.Name = "Odd Sheet"

        for i, v in enumerate(E_VALUES):
            calc.Cells(2 + i, 5).Value = v
        for i, v in enumerate(F_VALUES):
            if v is not None:
                calc.Cells(2 + i, 6).Value = v
        for i, v in enumerate(D_VALUES):
            calc.Cells(2 + i, 4).Value = v
        for i, v in enumerate(H_VALUES):
            calc.Cells(2 + i, 8).Value = v
        for i, v in enumerate(I_VALUES):
            calc.Cells(2 + i, 9).Value = v
        for i, v in enumerate(K_VALUES):
            calc.Cells(2 + i, 11).Value = v
        for i, v in enumerate(L_VALUES):
            calc.Cells(2 + i, 12).Value = v
        for i, v in enumerate(J_VALUES):
            calc.Cells(2 + i, 10).Value = v
        import datetime as _dt
        for i, iso in enumerate(M_DATES):
            c = calc.Cells(2 + i, 13)
            c.Value = _dt.datetime.strptime(iso, "%Y-%m-%d")
            c.NumberFormat = "yyyy-mm-dd"
        for i, (h, v) in enumerate(zip(HLOOKUP_HEADERS, HLOOKUP_VALUES)):
            calc.Cells(1, 14 + i).Value = h
            calc.Cells(2, 14 + i).Value = v

        for i, v in enumerate([7, 11, 13, 17]):
            data.Cells(2 + i, 2).Value = v
        odd.Cells(2, 2).Value = 99

        row = 2
        calc.Cells(1, 1).Value = "case"
        calc.Cells(1, 2).Value = "formula"
        for label, formula in CASES:
            calc.Cells(row, 1).Value = label
            calc.Cells(row, 2).Formula = formula
            row += 1

        app.Calculate()
        OUT.parent.mkdir(parents=True, exist_ok=True)
        if OUT.exists():
            OUT.unlink()
        wb.SaveAs(str(OUT), FileFormat=51)
        wb.Close(SaveChanges=False)
        print(f"wrote {OUT} with {len(CASES)} conformance cases")
        return 0
    finally:
        app.Quit()


if __name__ == "__main__":
    sys.exit(main())
