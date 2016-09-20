Attribute VB_Name = "Module1"
'Isaac Chang
'Version 5.0
'2016-06-27

Sub CalculateSheet_Click()

    'create new daily task list
    Dim dailyList As Worksheet
    Set dailyList = ThisWorkbook.Sheets.Add()
    calcDate = Worksheets("Grass Cut Summary").DTPicker1.Value
    calcDateString = Format(calcDate, "yyyy/mm/dd")
    dailyList.Name = calcDateString
    
    'Populate Headers in Daily Sheet
    copyRange = "A" & 3 & ":" & "W" & 3
    Worksheets("Grass Cut Summary").Range(copyRange).Copy
    Worksheets(calcDateString).Select
    Worksheets(calcDateString).Range("D3").Select
    Worksheets(calcDateString).Paste
    
    copyRange = "A" & 4 & ":" & "W" & 4
    Worksheets("Grass Cut Summary").Range(copyRange).Copy
    Worksheets(calcDateString).Select
    Worksheets(calcDateString).Range("D4").Select
    Worksheets(calcDateString).Paste
    
    copyRange = "A" & 5 & ":" & "W" & 5
    Worksheets("Grass Cut Summary").Range(copyRange).Copy
    Worksheets(calcDateString).Select
    Worksheets(calcDateString).Range("D5").Select
    Worksheets(calcDateString).Paste
 
    'Variable Declaration---------------------------------------------------
    Dim monthNames
    monthNames = Array("", "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December")
    
    'cell count starts at first total cost
    
    'output location on daily task list
    'outputRow = 3
    'outputCol = 1
    
    'where the first of the data will start being pasted into the new sheet
    RowpasteCount = 6
    
    'where calculation starts
    RowCount = 6
    ColumnCount = 7
    
    'getting values for paste sheet
    Sheets("Grass Cut Summary").Select
    lastRow = Cells(1, 6).Value
    
    monthToCalc = Worksheets("Grass Cut Summary").DTPicker1.Month
    dayToCalc = Worksheets("Grass Cut Summary").DTPicker1.Day
    '----------------------------------------------------------------------------
    
    'getting col of month to calc--------------------------------------------------------
    monthToCalcName = monthNames(monthToCalc)
    colNum = 1
    rowNum = 5
    
    While Trim(Cells(rowNum, colNum).Value) <> Trim(monthToCalcName) And colNum < 23
        colNum = colNum + 1
        'MsgBox (monthToCalcName)
        'MsgBox (Cells(rowNum, colNum).Value)
    Wend
    
    '---------------------------------------------------------------------------------------
    
    'colDiff is the difference from Total cost to month to calc
    colDiff = colNum - ColumnCount
    
    'variable for frequency for month
    Dim dayGap As Integer
    
    While RowCount <= lastRow
    If (LCase(Trim(Cells(RowCount, ColumnCount + 17).Value)) = "y") Then
        Sheets("Grass Cut Summary").Select
        
        'variable for whether monthly, daily or seasonal
        Dim payPlan As String
        
        payPlan = LCase(Trim(Cells(RowCount, ColumnCount - 2).Value))
        
        
        'dayGap based on frequency of service and early buffer region as specified by freq
        If (Cells(RowCount, ColumnCount + 2).Value = 2) Then
            dayGap = Int((30 / Cells(RowCount, ColumnCount + 2).Value) - 2)
        ElseIf (Cells(RowCount, ColumnCount + 2).Value = 4) Then
            dayGap = Int((30 / Cells(RowCount, ColumnCount + 2).Value))
        ElseIf (Cells(RowCount, ColumnCount + 2).Value = 3) Then
            dayGap = 9
        Else
            MsgBox ("Invalid Frequency")
        End If
        '------------------------------------------------------------------------------------
        
        'reset values
        monthToCalc = Worksheets("Grass Cut Summary").DTPicker1.Month
        dayToCalc = Worksheets("Grass Cut Summary").DTPicker1.Day
        colDiff = colNum - ColumnCount
        
        '-----------------------------------------------------------------------------------------------------------------------------
        'calculate cost per service
        Cells(RowCount, ColumnCount + 3).Interior.ColorIndex = xlNone
        If (Cells(RowCount, ColumnCount).Value <> "" And Cells(RowCount, ColumnCount + 2).Value <> "") Then
            If (payPlan = "month") Then
                Cells(RowCount, ColumnCount + 3).Value = Round(Cells(RowCount, ColumnCount).Value / Cells(RowCount, ColumnCount + 2).Value, 2)
            ElseIf (payPlan = "day") Then
                Cells(RowCount, ColumnCount + 3).Value = Round(Cells(RowCount, ColumnCount).Value, 2)
            ElseIf (payPlan = "seasonal") Then
                Cells(RowCount, ColumnCount + 3).Value = Round(Cells(RowCount, ColumnCount).Value / (6 * Cells(RowCount, ColumnCount + 2).Value), 2)
            End If
        Else
            Cells(RowCount, ColumnCount + 3).Interior.Color = RGB(255, 0, 0)
            Cells(RowCount, ColumnCount + 3).Value = "MISSING DATA"
        End If
        '-----------------------------------------------------------------------------------------------------------------------------
        
        '-----------------------------------------------------------------------------------------------------------------------------
        'calculate owed amount
        Cells(RowCount, ColumnCount - 1).Interior.ColorIndex = xlNone
        If (Cells(RowCount, ColumnCount).Value <> "" And Cells(RowCount, ColumnCount + 1).Value <> "") Then
            If (payPlan = "seasonal") Then
                Cells(RowCount, ColumnCount - 1).Value = Cells(RowCount, ColumnCount).Value - Cells(RowCount, ColumnCount + 1).Value
            End If
        Else
            Cells(RowCount, ColumnCount + 3).Interior.Color = RGB(255, 0, 0)
            Cells(RowCount, ColumnCount + 3).Value = "MISSING DATA"
        End If
        '-----------------------------------------------------------------------------------------------------------------------------
        
        '-----------------------------------------------------------------------------------------------------------------------------
        'getting last serviced date from comma expression
        
        fromPrevMonth = False 'last serviced date was in last month colDiff is set at prev month
        numMonthsBefore = 0
        
        If (Cells(RowCount, ColumnCount + colDiff).Value = "") Then
            fromPrevMonth = True
            While (Cells(RowCount, ColumnCount + colDiff).Value = "")
                colDiff = colDiff - 2
                numMonthsBefore = numMonthsBefore + 1
            Wend
            'now colDiff is now the difference from total cost column number to the month of last service
        End If
        
        'getting last service date from comma experssion
        daysServiced = Split(Cells(RowCount, ColumnCount + colDiff).Value, ",")
        Dim lastServicedDate As Integer
        If IsNumeric(daysServiced(UBound(daysServiced))) Then
            lastServicedDate = Int(daysServiced(UBound(daysServiced)))
        End If
        '-----------------------------------------------------------------------------------------------------------------------------
        
        
        '-----------------------------------------------------------------------------------------------------------------------------
        'getting next service date based on number of days in month
        
        inNextMonth = False 'next service date is in next month bool check
        
        'monthToCalc is now the last serviced month
        monthToCalc = monthToCalc - numMonthsBefore 'dont worry monthToCalc is reset for each loop
        
        Dim minNextServicDate As Integer
        'months with 31 days
        If (monthToCalc = 1 Or monthToCalc = 3 Or monthToCalc = 5 Or monthToCalc = 7 Or monthToCalc = 8 Or monthToCalc = 10 Or monthToCalc = 12) Then
            If (lastServicedDate + dayGap > 31) Then
                minNextServiceDate = (lastServicedDate + dayGap) - 31
                inNextMonth = True
            Else
                minNextServiceDate = lastServicedDate + dayGap
            End If
        'february has 28 days (not counting leap years here)
        ElseIf (monthToCalc = 2) Then
            If (lastServicedDate + dayGap > 28) Then
                minNextServiceDate = (lastServicedDate + dayGap) - 28
                inNextMonth = True
            Else
                minNextServiceDate = lastServicedDate + dayGap
            End If
        'months with 30 days
        Else
            If (lastServicedDate + dayGap > 30) Then
                minNextServiceDate = (lastServicedDate + dayGap) - 30
                inNextMonth = True
            Else
                minNextServiceDate = lastServicedDate + dayGap
            End If
        End If
        '-----------------------------------------------------------------------------------------------------------------------------
        
        'Copy paste Row if conditions for required work is met-------------------------------------------------------
        
        'fromPrevMonth is lower limit of month and inNextMonth is upper limit of month
        'Case where way overdue, might not even be touched but whatevs
        
        CurrentMonth = Worksheets("Grass Cut Summary").DTPicker1.Month
        
        '
        diffInDay = dayToCalc - minNextServiceDate
        copyRange = "A" & RowCount & ":" & "W" & RowCount
        
        'this condition will also take into account if more than 1 month before
        'so that wont have to be checked from now on
        If (numMonthsBefore > 1) Then
        
            Worksheets("Grass Cut Summary").Range(copyRange).Copy
            Worksheets(calcDateString).Select
            
            Cells(RowpasteCount, 2).Value = "OVERDUE FROM PREVIOUS MONTHS"
            
            pasteRange = "D" & RowpasteCount
            
            Worksheets(calcDateString).Range(pasteRange).Select
            Worksheets(calcDateString).Paste
            RowpasteCount = RowpasteCount + 1
        
        
        ElseIf (fromPrevMonth = True And inNextMonth = False) Then
            Worksheets("Grass Cut Summary").Range(copyRange).Copy
            Worksheets(calcDateString).Select
            
            Cells(RowpasteCount, 2).Value = "OVERDUE FROM LAST MONTH"
            
            pasteRange = "D" & RowpasteCount
            
            Worksheets(calcDateString).Range(pasteRange).Select
            Worksheets(calcDateString).Paste
            
            RowpasteCount = RowpasteCount + 1
              
        ElseIf (fromPrevMonth = True And inNextMonth = True) Then
            If (diffInDay >= -2) Then
                Worksheets("Grass Cut Summary").Range(copyRange).Copy
                Worksheets(calcDateString).Select
                
                'because infront of dayToCalc is actually negative
                'while before (so late) is positive since Im subtracting from dayToCalc
                outputDiffInDay = -diffInDay
                Cells(RowpasteCount, 2).Value = outputDiffInDay
            
                pasteRange = "D" & RowpasteCount
            
                Worksheets(calcDateString).Range(pasteRange).Select
                Worksheets(calcDateString).Paste
                
                RowpasteCount = RowpasteCount + 1
            End If
        
        ElseIf (fromPrevMonth = False And inNextMonth = False) Then
            If (diffInDay >= -2) Then
                Worksheets("Grass Cut Summary").Range(copyRange).Copy
                Worksheets(calcDateString).Select
                
                outputDiffInDay = -diffInDay
                Cells(RowpasteCount, 2).Value = outputDiffInDay
            
                pasteRange = "D" & RowpasteCount
            
                Worksheets(calcDateString).Range(pasteRange).Select
                Worksheets(calcDateString).Paste
                
                RowpasteCount = RowpasteCount + 1
            End If
        ElseIf (fromPrevMonth = False And inNextMonth = True) Then
            '31 day month
            If (CurrentMonth = 1 Or CurrentMonth = 3 Or CurrentMonth = 5 Or CurrentMonth = 7 Or CurrentMonth = 8 Or CurrentMonth = 10 Or CurrentMonth = 12) Then
                If ((31 - dayToCalc + minNextServiceDate) <= 2) Then
                    Worksheets("Grass Cut Summary").Range(copyRange).Copy
                    Worksheets(calcDateString).Select
            
                    Cells(RowpasteCount, 2).Value = Abs(31 - dayToCalc + minNextServiceDate)
            
                    pasteRange = "D" & RowpasteCount
            
                    Worksheets(calcDateString).Range(pasteRange).Select
                    Worksheets(calcDateString).Paste
                    
                    RowpasteCount = RowpasteCount + 1
                End If
            '28 day month
            ElseIf (CurrentMonth = 2) Then
                If ((28 - dayToCalc + minNextServiceDate) <= 2) Then
                    Worksheets("Grass Cut Summary").Range(copyRange).Copy
                    Worksheets(calcDateString).Select
                    
                    Cells(RowpasteCount, 2).Value = Abs(28 - dayToCalc + minNextServiceDate)
            
                    pasteRange = "D" & RowpasteCount
            
                    Worksheets(calcDateString).Range(pasteRange).Select
                    Worksheets(calcDateString).Paste
                    
                    RowpasteCount = RowpasteCount + 1
                End If
            '30 day month
            Else
                If ((30 - dayToCalc + minNextServiceDate) <= 2) Then
                    Worksheets("Grass Cut Summary").Range(copyRange).Copy
                    Worksheets(calcDateString).Select
                    
                    Cells(RowpasteCount, 2).Value = Abs(30 - dayToCalc + minNextServiceDate)
                    
                    pasteRange = "D" & RowpasteCount
            
                    Worksheets(calcDateString).Range(pasteRange).Select
                    Worksheets(calcDateString).Paste
                    
                    RowpasteCount = RowpasteCount + 1
                End If
            End If
        End If
        '------------------------------------------------------------------------------------------
    End If
        'Increase RowCount for  next row
        RowCount = RowCount + 1
        Sheets("Grass Cut Summary").Select
    
    Wend
    
    
    Sheets(calcDateString).Select
    
    'populate extra headers for daily task list
    dailyList.Cells(1, 1).Font.Size = 20
    dailyList.Cells(1, 1).Value = "Daily Task List"
    dailyList.Cells(5, 1).Font.FontStyle = "Bold"
    dailyList.Cells(5, 1).Value = "Order"
     Worksheets(calcDateString).Cells(5, 2).Font.FontStyle = "Bold"
    Worksheets(calcDateString).Cells(5, 2).Value = "Days to Service"
    Worksheets(calcDateString).Cells(5, 3).Font.FontStyle = "Bold"
    Worksheets(calcDateString).Cells(5, 3).Value = "Time Scheduled"
    
    'finding column num of cost per service------------------------------------
    colNum = 1
    rowNum = 5
    
    'colNum will be column number of cost per service
    dailyList.Select
    While Trim(Cells(rowNum, colNum).Value) <> "Cost per Service" And colNum < 23
        colNum = colNum + 1
    Wend
    '-----------------------------------------------------------------------------------
    
    'reusing variable
    rowNum = 6
    
    Dim TotalDailyRevenue As Double
    TotalDailyRevenue = 0
    
    While (Cells(rowNum, colNum).Value <> "")
        TotalDailyRevenue = TotalDailyRevenue + Cells(rowNum, colNum).Value
        rowNum = rowNum + 1
    Wend
    rowNum = rowNum + 1
    
    Cells(rowNum, 1).Font.Size = 16
    Cells(rowNum, 1).Value = "Total Daily Revenue: " & "$" & TotalDailyRevenue
    'MsgBox (Cells(1, 1))
End Sub

'column calcaulations for daily task list is bsed on names colNums not hard coded
'except for "order" title and "Days to Service" title
