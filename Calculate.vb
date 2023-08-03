Sub Total()

Dim MIN_RANGE As Integer
Dim MAX_RANGE As Integer
Dim t_range As Integer
Dim mrange As Integer
Dim SERIAL As Integer
Dim DIFF As Integer
Dim M_COL As Integer
Dim skill As Integer
Dim kettles As Integer
Dim steam As Integer
Dim HOLD As Integer
Dim limit As Integer
Dim Result As String
Dim Total As Integer

Application.ScreenUpdating = False

SERIAL = 2
MIN_RANGE = 2
MAX_RANGE = 700
DIFF = 28
M_COL = 4
skill = 4
kettles = 4
steam = 4
HOLD = 15
limit = 3
Total = 0


Sheets("Totals").Cells(1, 1) = "Steamers"
Sheets("Totals").Cells(2, 1) = "Heavy Body"
Sheets("Totals").Cells(2, 2) = Application.WorksheetFunction.CountIf(Sheets("Steamers").range("C4:C700"), "Heavy Body")
Sheets("Totals").Cells(3, 1) = "Classic"
Sheets("Totals").Cells(3, 2) = Application.WorksheetFunction.CountIf(Sheets("Steamers").range("C4:C700"), "Classic")
Sheets("Totals").Cells(4, 1) = "STM10"
Sheets("Totals").Cells(4, 2) = Application.WorksheetFunction.CountIf(Sheets("Steamers").range("C4:C700"), "STM10")
Sheets("Totals").Cells(5, 1) = "STMSC"
Sheets("Totals").Cells(5, 2) = Application.WorksheetFunction.CountIf(Sheets("Steamers").range("C4:C700"), "STMSC")
Sheets("Totals").Cells(7, 1) = "Kettles"
Sheets("Totals").Cells(8, 1) = "KGT"
Sheets("Totals").Cells(8, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "KGT")
Sheets("Totals").Cells(9, 1) = "KGL25"
Sheets("Totals").Cells(9, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "KGL25")
Sheets("Totals").Cells(10, 1) = "MFS/Other"
Sheets("Totals").Cells(10, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "MFS/OTHER")
Sheets("Totals").Cells(11, 1) = "Stand"
Sheets("Totals").Cells(11, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "Stand")
Sheets("Totals").Cells(12, 1) = "Food-Pump"
Sheets("Totals").Cells(12, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "Food-Pump")
Sheets("Totals").Cells(13, 1) = "Cook Chill"
Sheets("Totals").Cells(13, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "Cook Chill")
Sheets("Totals").Cells(14, 1) = "KGL"
Sheets("Totals").Cells(14, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "KGL")
Sheets("Totals").Cells(15, 1) = "KEL"
Sheets("Totals").Cells(15, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "KEL")
Sheets("Totals").Cells(16, 1) = "KDT"
Sheets("Totals").Cells(16, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "KDT")
Sheets("Totals").Cells(17, 1) = "KDL"
Sheets("Totals").Cells(17, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "KDL")
Sheets("Totals").Cells(18, 1) = "KET"
Sheets("Totals").Cells(18, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "KET")
Sheets("Totals").Cells(19, 1) = "HAMKGL"
Sheets("Totals").Cells(19, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "HAMKGL")
Sheets("Totals").Cells(20, 1) = "HAMKDL"
Sheets("Totals").Cells(20, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "HAMKDL")
Sheets("Totals").Cells(21, 1) = "MKGL"
Sheets("Totals").Cells(21, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "MKGL")
Sheets("Totals").Cells(22, 1) = "MKDL"
Sheets("Totals").Cells(22, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "MKDL")
Sheets("Totals").Cells(23, 1) = "MKEL"
Sheets("Totals").Cells(23, 2) = Application.WorksheetFunction.CountIf(Sheets("Kettles").range("C4:C700"), "MKEL")
Sheets("Totals").Cells(25, 1) = "Skillets"
Sheets("Totals").Cells(26, 1) = "SET"
Sheets("Totals").Cells(26, 2) = Application.WorksheetFunction.CountIf(Sheets("Skillets").range("C4:C700"), "SET")
Sheets("Totals").Cells(27, 1) = "SELT1"
Sheets("Totals").Cells(27, 2) = Application.WorksheetFunction.CountIf(Sheets("Skillets").range("C4:C700"), "SELT1")
Sheets("Totals").Cells(28, 1) = "SGLT1"
Sheets("Totals").Cells(28, 2) = Application.WorksheetFunction.CountIf(Sheets("Skillets").range("C4:C700"), "SGLT1")
Sheets("Totals").Cells(29, 1) = "SELTR"
Sheets("Totals").Cells(29, 2) = Application.WorksheetFunction.CountIf(Sheets("Skillets").range("C4:C700"), "SELTR")
Sheets("Totals").Cells(30, 1) = "SGLTR"
Sheets("Totals").Cells(30, 2) = Application.WorksheetFunction.CountIf(Sheets("Skillets").range("C4:C700"), "SGLTR")


Sheets("Totals").Select
 range("A1:B1").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Font.Bold = True
 range("A7:B7").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Font.Bold = True
 range("A25:B25").Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    Selection.Merge
    Selection.Font.Bold = True


End Sub
