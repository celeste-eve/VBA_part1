Attribute VB_Name = "Module1"

Sub GenerateVisitorCalendar()
    Dim ws As Worksheet, wsData As Worksheet
    Dim startDate As Date, endDate As Date
    Dim cell As Range, dataRange As Range
    Dim rowNum As Integer, colNum As Integer
    Dim calendarStart As Date
    Dim monthCounter As Integer
    Dim currentMonth As Integer, currentYear As Integer
    Dim visitorInfo As String
    Dim maxVisitors As Integer
    Dim visitorRow As Integer
    Dim dayVisitorCount As Integer
    Dim weekMaxVisitors As Integer
    Dim currentWeekStart As Date, dayOfWeek As Integer
    Dim dateRow As Integer
    Dim thisDate As Date
    Dim visRow As Integer
    Dim checkDate As Date
    Dim lastRow As Long, lastCol As Long
    
    
    
    ' Define the worksheets
    Set wsData = ThisWorkbook.Sheets("PlannedVisitorsSheet") ' Data sheet
    Set ws = ThisWorkbook.Sheets("PlannedVisitorsSheet") ' Calendar sheet
    
    ' Define start column
    Dim startColumn As Integer
    startColumn = 12 ' Start calendar in column L
    
    ' Get current month and year
    currentMonth = Month(Date)
    currentYear = Year(Date)

' Clear previous calendar data only in the intended range
If Not ws Is Nothing Then
    ' Get absolute last row of the sheet to ensure I clear everything
    lastRow = ws.Rows.Count
    lastCol = ws.Columns.Count
    
    ' Clear the entire calendar area (from column L to end of sheet)
    ws.Range(ws.Cells(1, startColumn), ws.Cells(lastRow, lastCol)).Clear
Else
    MsgBox "Error: Worksheet not found!", vbExclamation, "Calendar Error"
    Exit Sub
End If

    ' Initialize collections to store visitor counts per day
    Dim visitorArray() As Integer
    ReDim visitorArray(31) ' Max days in a month
    
       ' Define a consistent structure
    Const TITLE_ROW_HEIGHT As Integer = 1     ' Height of month title
    Const HEADER_ROW_HEIGHT As Integer = 1    ' Height of day headers
    Const SPACING_AFTER_MONTH As Integer = 2  ' Space between months
    
    
    ' Loop through the next 3 months
    rowNum = 1 ' Start from the first row
    For monthCounter = 0 To 2
        ' Get first and last day of the target month
        calendarStart = DateSerial(currentYear, currentMonth + monthCounter, 1)
        endDate = DateSerial(currentYear, currentMonth + monthCounter + 1, 0)


        ' Store the starting row for this month
        Dim monthStartRow As Integer
        monthStartRow = rowNum

        ' Insert Month Header
        ws.Cells(rowNum, startColumn).Value = Format(calendarStart, "MMMM YYYY")
        ws.Cells(rowNum, startColumn).Font.Bold = True
        ws.Cells(rowNum, startColumn).Font.Color = RGB(200, 0, 200)
        ws.Cells(rowNum, startColumn).Font.Size = 14
        rowNum = monthStartRow + TITLE_ROW_HEIGHT

        ' Create Day Headers
        ws.Cells(rowNum, startColumn).Value = "Monday"
        ws.Cells(rowNum, startColumn + 1).Value = "Tuesday"
        ws.Cells(rowNum, startColumn + 2).Value = "Wednesday"
        ws.Cells(rowNum, startColumn + 3).Value = "Thursday"
        ws.Cells(rowNum, startColumn + 4).Value = "Friday"
        ws.Cells(rowNum, startColumn + 5).Value = "Saturday"
        ws.Cells(rowNum, startColumn + 6).Value = "Sunday"
        
        ws.Range(ws.Cells(rowNum, startColumn), ws.Cells(rowNum, startColumn + 6)).Font.Bold = True
        ws.Range(ws.Cells(rowNum, startColumn), ws.Cells(rowNum, startColumn + 6)).Font.Color = RGB(200, 0, 200)
        rowNum = rowNum + 1

        ' Count visitors for each date in this month
        Set dataRange = wsData.Range("A2:A" & wsData.Cells(Rows.Count, 1).End(xlUp).Row)
        
        ' Reset the visitor array
        Dim i As Integer
        For i = 1 To 31
            visitorArray(i) = 0
        Next i
        
        ' Count visitors for each day
        For Each cell In dataRange
            If IsDate(cell.Value) Then ' Ensure we only process actual dates
                If cell.Value >= calendarStart And cell.Value <= endDate Then
                    Dim dayNum As Integer
                    dayNum = Day(cell.Value)
                    visitorArray(dayNum) = visitorArray(dayNum) + 1
                End If
            End If
        Next cell

        ' Find the first day of the week for this month
        Dim firstDayCol As Integer
        firstDayCol = Weekday(calendarStart, vbMonday) ' 1=Monday, 7=Sunday
        
        ' Process each week until we've passed the end of the month
        Dim currentDay As Integer
        currentDay = 1
        colNum = firstDayCol ' Start with the correct weekday column

        ' Continue until we've processed all days in the month
        Do While currentDay <= Day(endDate)
            ' Find max visitors for this week to determine row height
            weekMaxVisitors = 0
            Dim tempDay As Integer
            For i = 0 To 6
                tempDay = currentDay + i
                If tempDay <= Day(endDate) Then
                    If visitorArray(tempDay) > weekMaxVisitors Then
                        weekMaxVisitors = visitorArray(tempDay)
                    End If
                End If
            Next i
            
            ' Start row for this week's dates
            dateRow = rowNum
            
          
            ' Process each day of the week
            Dim dayCol As Integer
            For dayCol = 1 To 7
                ' Only process days that are part of this month
                If colNum <= 7 Then
                    ' We're in the first week and might need to skip some days
                    If dayCol >= firstDayCol Or currentDay > 1 Then
                        ' don't exceed the month
                        If currentDay <= Day(endDate) Then
                            thisDate = DateSerial(currentYear, currentMonth + monthCounter, currentDay)
                            
                            ' Write the date
                            ws.Cells(dateRow, startColumn + dayCol - 1).Value = thisDate
                            ws.Cells(dateRow, startColumn + dayCol - 1).NumberFormat = "DD-MMM"
                            
                            ' Add visitors for this date
                            visRow = dateRow + 1
                            
                            Dim insertedVisitors As Object
                            Set insertedVisitors = CreateObject("Scripting.Dictionary")
                            
                            For Each cell In dataRange
                                If Not IsEmpty(cell.Value) Then
                                    If cell.Value = thisDate Then
                                        visitorInfo = cell.Offset(0, 1).Value & " - " & cell.Offset(0, 2).Value
                                        
                                        ' Check if this visitor has already been added
                                        If Not insertedVisitors.exists(visitorInfo) Then
                                            ws.Cells(visRow, startColumn + dayCol - 1).Value = visitorInfo
                                            ws.Cells(visRow, startColumn + dayCol - 1).Font.Color = RGB(200, 0, 200)
                                            insertedVisitors.Add visitorInfo, True
                                            visRow = visRow + 1
                                        End If
                                    End If
                                End If
                            Next cell
            
                            currentDay = currentDay + 1
                        End If
                    End If
                End If
                colNum = colNum + 1
            Next dayCol
            
            ' Reset column for next week
            colNum = 1
            ' Move to the next row
            rowNum = rowNum + weekMaxVisitors + 2
            firstDayCol = 1
        Loop
        

        ' add consistent spacing
        rowNum = rowNum + SPACING_AFTER_MONTH
    Next monthCounter


    ' Auto-adjust column width
    ws.Cells.Columns.AutoFit
    
   
    ' Add borders only to the calendar area
    Dim calendarRange As Range
    Set calendarRange = ws.Range(ws.Cells(1, startColumn), ws.Cells(rowNum - 1, startColumn + 6))
    calendarRange.Borders.LineStyle = xlContinuous
    
    ' formatting to the calendar area
    calendarRange.Interior.ColorIndex = xlNone ' Clear any background colors
End Sub

