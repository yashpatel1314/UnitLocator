Sub Change()
'
' Change Macro
'
Dim MIN_RANGE As Integer
Dim MAX_RANGE As Integer
Dim mrange As Integer
Dim SERIAL As Integer
Dim DIFF As Integer
Dim M_COL As Integer
Dim skill As Integer
Dim kettles As Integer
Dim steam As Integer
Dim HOLD As Integer
Dim Result As String

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

For mrange = MIN_RANGE To MAX_RANGE

    Sheets("PLT").Select
    Cells(mrange, SERIAL).Select
    
    If Selection.Style = "Neutral" Then
    
   
        
        If IsEmpty(Cells(mrange, HOLD)) = True Then
        
        
            For type_range = MIN_RANGE To MAX_RANGE
            
        
                If Sheets("PLT").Cells(mrange, M_COL) = Sheets("Month Quota").Cells(type_range, M_COL) Then
                
                    Sheets("Month Quota").Select
                    
                    Result = StrComp(Cells(type_range, DIFF), "Skillets", vbTextCompare)
                   
                    If Result = 0 Then
                    
                        Sheets("PLT").Select
                        Sheets("Skillets").Cells(skill, SERIAL) = Cells(mrange, SERIAL).Value
                        skill = skill + 1
                        type_range = MAX_RANGE
            
                    End If
                    
                    
                    Result = StrComp(Cells(type_range, DIFF), "Kettles", vbTextCompare)
                   
                    If Result = 0 Then
                    
                         Sheets("PLT").Select
                       Sheets("Kettles").Cells(kettles, SERIAL) = Cells(mrange, SERIAL).Value
                      kettles = kettles + 1
                      type_range = MAX_RANGE
            
                    End If
                    
                     Result = StrComp(Cells(type_range, DIFF), "Cook Chill", vbTextCompare)
                   
                    If Result = 0 Then
                    
                         Sheets("PLT").Select
                       Sheets("Kettles").Cells(kettles, SERIAL) = Cells(mrange, SERIAL).Value
                      kettles = kettles + 1
                      type_range = MAX_RANGE
            
                    End If
                     
        
                    Result = StrComp(Cells(type_range, DIFF), "Steam", vbTextCompare)
                   
                    If Result = 0 Then
                    
                         Sheets("PLT").Select
                       Sheets("Steamers").Cells(steam, SERIAL) = Cells(mrange, SERIAL).Value
                      steam = steam + 1
                      type_range = MAX_RANGE
            
                    End If
                    
             
                End If
        
            Next type_range
            
        End If
        
    End If
    
    If IsEmpty(Cells(mrange + 1, SERIAL)) = True Then

        mrange = MAX_RANGE

    End If

Next mrange

       ActiveWorkbook.Worksheets("Steamers").ListObjects("Table1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Steamers").ListObjects("Table1").Sort.SortFields. _
        Add2 Key:=range("Table1[Type]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, CustomOrder:="Heavy Body,Classic,STM10,STMSC", DataOption:= _
        xlSortNormal
    With ActiveWorkbook.Worksheets("Steamers").ListObjects("Table1").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
                
     ActiveWorkbook.Worksheets("Kettles").ListObjects("Table515").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Kettles").ListObjects("Table515").Sort.SortFields. _
        Add2 Key:=range("Table515[Type]"), SortOn:=xlSortOnValues, Order:= _
        xlAscending, CustomOrder:= _
        "KGT,KGL25,MFS/Other,Stand,Food-Pump,Cook Chill,KGL,KEL,KDL,KDT,KET", _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Kettles").ListObjects("Table515").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
                    
Sheets("Input").Select

End Sub
