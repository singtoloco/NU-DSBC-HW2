Attribute VB_Name = "Module1"
Sub sumvolume()

    Dim lrow As Long 'last row of a worksheet
    Dim volume As Double
    Dim ticker As String
        
    Dim t_count As Long 'ticker counts
    Dim srow As Long 'starting row of a ticker
    
    'sum data
'    Dim lrow_k As Long 'last row of sum data
    
    Dim max_pct As Double
    Dim max_pct_index As Long
    Dim min_pct As Double
    Dim min_pct_index As Long
    Dim max_volume As Double
    Dim max_volume_index As Long
        
    Dim WS_Count As Integer
    Dim w As Integer
    
    ' Set WS_Count equal to the number of worksheets in the active
         ' workbook.
    WS_Count = ActiveWorkbook.Worksheets.Count
        
    'LoopWorksheets
    For w = 1 To WS_Count
    
        
        '***********************
        
        'Worksheets("2015").Activate
        Worksheets(w).Activate
        
        lrow = Cells(Rows.Count, 1).End(xlUp).Row
        'MsgBox lrow 'just for debugging
        
        volume = 0
        t_count = 1
        srow = 2
            
        For I = 2 To lrow
        
            If Cells(I, 3).Value = 0 Then
            
                srow = srow + 1
                
            Else
            
                volume = volume + Cells(I, 7).Value
                
                If Cells(I, 1).Value <> Cells(I + 1, 1).Value Then
                    
                    t_count = t_count + 1
                    Cells(t_count, 9).Value = Cells(I, 1).Value
                    Cells(t_count, 12).Value = volume
                    volume = 0 'reset volume for next ticker
                    
                    'Calc Yearly Chg
                    Cells(t_count, 10).Value = Cells(I, 6).Value - Cells(srow, 3).Value
                    'Conditional formatting (if Yearly Chg = 0 then no color)
                    If Cells(t_count, 10).Value < 0 Then
                        Cells(t_count, 10).Interior.ColorIndex = 3
                    ElseIf Cells(t_count, 10).Value > 0 Then
                        Cells(t_count, 10).Interior.ColorIndex = 4
                    End If
                    
                    
                    'Calc Yearly Percent Chg
                    Cells(t_count, 11).Value = Cells(t_count, 10).Value / Cells(srow, 3).Value
                    'Formatting the cell to percentage
                    Cells(t_count, 11).NumberFormat = "0.00%"
                    
                    srow = I + 1 'reset srow for next ticker
                    
                End If
                
            End If
            
        Next I
        
        'Dealing with sum data
        lrow_k = Cells(Rows.Count, 11).End(xlUp).Row
        
        'Set Rng_pct = Range("K2:K" & lrow_k)
        
        max_pct = Application.WorksheetFunction.Max(Range("K1:K" & lrow_k))
        max_pct_index = Application.WorksheetFunction.Match(max_pct, Range("K1:K" & lrow_k), 0)
        Range("P2").Value = Cells(max_pct_index, 9).Value
        Range("Q2").Value = max_pct
        Range("Q2").NumberFormat = "0.00%"
    
        min_pct = Application.WorksheetFunction.Min(Range("K1:K" & lrow_k))
        min_pct_index = Application.WorksheetFunction.Match(min_pct, Range("K1:K" & lrow_k), 0)
        Range("P3").Value = Cells(min_pct_index, 9).Value
        Range("Q3").Value = min_pct
        Range("Q3").NumberFormat = "0.00%"
        
        'Set Rng_volume = Range("L2:L" & lrow_k)
    
        max_volume = Application.WorksheetFunction.Max(Range("L1:L" & lrow_k))
        max_volume_index = Application.WorksheetFunction.Match(max_volume, Range("L1:L" & lrow_k), 0)
        Range("P4").Value = Cells(max_volume_index, 9).Value
        Range("Q4").Value = max_volume
        Range("Q4").NumberFormat = "0"
        
        
        '***********************
    Next w
    
        
    
End Sub
