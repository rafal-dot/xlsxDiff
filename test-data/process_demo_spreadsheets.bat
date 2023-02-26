REM Just compare two spreadsheets
python ..\xlsxDiff.py in1.xlsx in2.xlsx out_demo.xlsx
REM Just compare two spreadsheets, but use formulas rather than data
python ..\xlsxDiff.py in1.xlsx in2.xlsx out_demo_formulas.xlsx -f
REM Compare two spreadsheets and find added/deleted rows/columns
python ..\xlsxDiff.py in1.xlsx in2.xlsx out_demo_cBr1.xlsx -c books!B -r books!1  -c rmdata!A -r rmdata!1 -c staff!B -r staff!1
REM Compare two spreadsheets and find added/deleted rows/columns, use extended index
python ..\xlsxDiff.py in1.xlsx in2.xlsx out_demo_cBCr1.xlsx -c books!B -r books!1  -c rmdata!A -r rmdata!1 -c staff!B,C -r staff!1
REM Compare two spreadsheets and find added/deleted rows/columns. No special highlighting
python ..\xlsxDiff.py in1.xlsx in2.xlsx out_demo_cBCr1X.xlsx -c books!B -r books!1  -c rmdata!A -r rmdata!1 -c staff!B,C -r staff!1 -X
