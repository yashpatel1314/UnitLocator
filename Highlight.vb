Sub Highlight()
'
' Operation_Real Macro
'


Dim notvariant As Integer
Dim notqty As Integer
Dim s_range As Integer
Dim m_range As Integer
Dim qty As Integer
Dim qty_range As Integer
Dim count As Integer
Dim MAX_RANGE As Integer
Dim HOLD As Integer
Dim MIN_RANGE As Integer
Dim M_COL As Integer
Dim S_COL As Integer
Dim N_COL As Integer
Dim NOTQTY_LOCATION As Integer
Dim QTY_LOCATION As Integer
Dim SERIAL As Integer
Dim search As Integer

Application.ScreenUpdating = False


NOTQTY_LOCATION = 2
N_COL = 1
nonvariant = 2
notqty = 2
s_range = 2
MAX_RANGE = 700
HOLD = 14
MIN_RANGE = 1
M_COL = 4
S_COL = 4
SERIAL = 2
QTY_LOCATION = 30

For m_range = MIN_RANGE To MAX_RANGE

    qty = 0
    
    If Sheets("Month Quota").Cells(s_range, S_COL) = Sheets("PLT").Cells(m_range, M_COL) Then
    
        Sheets("PLT").Select
        
        If IsEmpty(Cells(m_range, HOLD)) = True Then
        
            Sheets("Month Quota").Select
            qty = Cells(s_range, QTY_LOCATION).Value
            Sheets("PLT").Select
            Cells(m_range, M_COL).Select
            Selection.Style = "Neutral"
            Cells(m_range, SERIAL).Select
            Selection.Style = "Neutral"
            Cells(m_range, HOLD).Value = 1
            Sheets("Month Quota").Select
            
            If Cells(s_range, QTY_LOCATION) > 1 Then
            
                qty = Cells(s_range, QTY_LOCATION).Value
                Sheets("PLT").Select
                count = 1
                        
                    For qty_range = MIN_RANGE To MAX_RANGE
                            
                        If Sheets("Month Quota").Cells(s_range, S_COL) = Sheets("PLT").Cells(qty_range, M_COL) Then
                            
                            If IsEmpty(Cells(qty_range, HOLD)) = True Then
         
                                Sheets("PLT").Select
                                Cells(qty_range, M_COL).Select
                                Selection.Style = "Neutral"
                                Cells(qty_range, SERIAL).Select
                                Selection.Style = "Neutral"
                                Cells(qty_range, HOLD).Value = 1
                                count = count + 1
                                              
                                If count = qty Then
                                            
                                    qty_range = MAX_RANGE
                                           
                                End If
                                        
                            End If
                                    
                        End If
                                
                    Next qty_range
                    
                    If count < qty Then
                    
                        Sheets("Extra").Select
                        Sheets("Extra").Cells(nonvariant, N_COL) = Sheets("Month Quota").Cells(s_range, S_COL)
                        Cells(notqty, NOTQTY_LOCATION).Value = qty - count
                        nonvariant = nonvariant + 1
                        notqty = notqty + 1
                        
                    End If
                        
          
            End If
            
            s_range = s_range + 1
            m_range = MIN_RANGE - 1
            Sheets("Month Quota").Select
            
            If IsEmpty(Cells(s_range, S_COL)) = True Then

                m_range = MAX_RANGE

            End If
            
        End If
            
    End If
        
    If qty = 0 And m_range = MAX_RANGE Then
        
        Sheets("Extra").Select
        Sheets("Extra").Cells(nonvariant, N_COL) = Sheets("Month Quota").Cells(s_range, S_COL)
        Sheets("Extra").Cells(notqty, NOTQTY_LOCATION) = Sheets("Month Quota").Cells(s_range, QTY_LOCATION)
        nonvariant = nonvariant + 1
        notqty = notqty + 1
        s_range = s_range + 1
        Sheets("Month Quota").Select
            
        If IsEmpty(Cells(s_range, S_COL)) = False Then
                  
                m_range = MIN_RANGE - 1
                    
        End If
            
    End If
                
Next m_range

Sheets("Input").Select
End Sub
