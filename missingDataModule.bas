Attribute VB_Name = "Module2"
Sub Button7_Click()
    Sheets("Grass Cut Summary").Select
    
    lastRow = Cells(1, 6).Value
    
    For i = 6 To lastRow Step 1
        For j = 1 To 9 Step 1
            Cells(i, j).Interior.ColorIndex = xlNone
            If (Cells(i, j).Value = "" Or Cells(i, j).Value = " " Or Cells(i, j).Value = "  " Or Cells(i, j).Value = "MISSING DATA" Or Cells(i, j).Value = "URGENT MISSING DATA") Then
                If (j = 7 Or j = 8 Or j = 9) Then
                    Cells(i, j).Interior.Color = RGB(255, 0, 0)
                    Cells(i, j).Value = "URGENT MISSING DATA"
                
                ElseIf (j <> 6) Then
                    Cells(i, j).Interior.Color = RGB(255, 165, 0)
                    Cells(i, j).Value = "MISSING DATA"
                End If
            End If
        Next
    Next
    
    
End Sub
