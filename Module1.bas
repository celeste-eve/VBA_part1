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

    ' Initialize collections to store visitor information for each day
    Dim visitorDict As Object
    Set visitorDict = CreateObject("Scripting.Dictionary")
    
    ' Populate visitor dictionary
    Set dataRange = wsData.Range("A2:A" & wsData.Cells(Rows.Count, 1).End(xlUp).Row)
    For Each cell In dataRange
        If IsDate(cell.Value) Then
            Dim visitorDate As Date
            visitorDate = cell.Value
            
            ' Create or add to the list of visitors for this date
            If Not visitorDict.Exists(visitorDate) Then
                visitorDict.Add visitorDate, cell.Offset(0, 1).Value & " - " & cell.Offset(0, 2).Value
            Else
                visitorDict(visitorDate) = visitorDict(visitorDate) & vbNewLine & cell.Offset(0, 1).Value & " - " & cell.Offset(0, 2).Value
            End If
        End If
    Next cell

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
        ' adjust start point for each month
        firstWeekDay = Weekday(DateSerial(currentYear, currentMonth, 1)) ' 1 = Sunday, 7 = Saturday
        firstWeekDay = firstWeekDay - 1
        If firstWeekDay = 0 Then firstWeekDay = 7


    
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

       ' Calculate first day of the month position (0 = Monday, 6 = Sunday)
        Dim firstDayCol As Integer
        ' Get the weekday of 1st of month (1=Sunday, 2=Monday, ..., 7=Saturday in vbSunday)
        Dim firstOfMonth As Date
        firstOfMonth = DateSerial(currentYear, currentMonth + monthCounter, 1)
        
        ' Convert to 0=Monday to 6=Sunday
        firstDayCol = (Weekday(firstOfMonth, vbSunday) + 5) Mod 7
                
        ' Process each week until we've passed the end of the month
        Dim currentDay As Integer
        currentDay = 1
        
        ' Continue until we've processed all days in the month
        Do While currentDay <= Day(endDate)
            ' Find max visitors for this week to determine row height
            weekMaxVisitors = 0
            Dim tempDay As Integer, tempCol As Integer
            tempCol = firstDayCol
            
            For i = 0 To 6
                If tempCol < 7 Then  ' Only check days within this week
                    tempDay = currentDay + i
                    If tempDay <= Day(endDate) Then
                        If visitorArray(tempDay) > weekMaxVisitors Then
                            weekMaxVisitors = visitorArray(tempDay)
                        End If
                    End If
                    tempCol = tempCol + 1
                End If
            Next i
            
            ' Ensure at least 1 row for visitors
            If weekMaxVisitors = 0 Then weekMaxVisitors = 1
            
            ' Start row for this week's dates
            dateRow = rowNum
            
            ' Process each day of the week
            Dim dayCol As Integer
            For dayCol = 0 To 6  ' 0=Monday, 6=Sunday
                ' Calculate the actual column to place this day
                Dim actualCol As Integer
                actualCol = startColumn + dayCol
                
                ' Check if we should place a day in this position
                If currentDay <= Day(endDate) Then
                    ' In the first week, only place dates from the firstDayCol onwards
                    If dayCol >= firstDayCol Or rowNum > (monthStartRow + TITLE_ROW_HEIGHT + HEADER_ROW_HEIGHT) Then
                        thisDate = DateSerial(currentYear, currentMonth + monthCounter, currentDay)
                        
                        ' Write the date
                        ws.Cells(dateRow, actualCol).Value = thisDate
                        ws.Cells(dateRow, actualCol).NumberFormat = "DD-MMM"
                        
                        ' Add visitors for this date
                        visRow = dateRow + 1
                        
                        Dim insertedVisitors As Object
                        Set insertedVisitors = CreateObject("Scripting.Dictionary")
                        
                        For Each cell In dataRange
                            If Not IsEmpty(cell.Value) Then
                                If cell.Value = thisDate Then
                                    visitorInfo = cell.Offset(0, 1).Value & " - " & cell.Offset(0, 2).Value
                                    
                                    ' Check if this visitor has already been added
                                    If Not insertedVisitors.Exists(visitorInfo) Then
                                        ws.Cells(visRow, actualCol).Value = visitorInfo
                                        ws.Cells(visRow, actualCol).Font.Color = RGB(200, 0, 200)
                                        insertedVisitors.Add visitorInfo, True
                                        visRow = visRow + 1
                                    End If
                                End If
                            End If
                        Next cell
        
                        currentDay = currentDay + 1
                    End If
                End If
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

