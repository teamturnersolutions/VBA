Sub ErrorReportWeeklySummary()
    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRowSource As Long
    Dim lastRowDest As Long
    Dim i As Long, j As Long, k As Long
    Dim uniqueData As Collection
    Dim key As Variant
    Dim tbl As ListObject
    Dim errorDates() As Date
    Dim dateCount As Integer
    Dim flagged As Boolean
    Dim currentDate As Date
    Dim mostRecentDate As Date
    Dim flagColor As Long
    Dim wsName As String
    
    ' Set the worksheet to your specific sheet
    Set wsSource = ThisWorkbook.Sheets("ErrorLogNew") ' Change to your actual sheet name
    
    ' Check if "TM Data Extraction" sheet already exists, and delete it if so
    wsName = "TM Data Extraction"
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets(wsName).Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    ' Create a new worksheet for the extracted data
    Set wsDest = ThisWorkbook.Sheets.Add
    wsDest.Name = wsName
    
    ' Cleanse the data in the source worksheet
    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row
    
    For i = lastRowSource To 2 Step -1
        If Trim(UCase(wsSource.Cells(i, 10).Value)) = "RE-ASSIGNED" Or _
           Trim(UCase(wsSource.Cells(i, 10).Value)) = "VOID" Or _
           Trim(UCase(wsSource.Cells(i, 10).Value)) = "COMBINED" Or _
           Trim(UCase(wsSource.Cells(i, 12).Value)) = "SEPARATION" Or _
           Trim(UCase(wsSource.Cells(i, 12).Value)) = "RESIGNED" Then
            wsSource.Rows(i).Delete
        End If
    Next i
    
    ' Initialize the collection to store unique TM data
    Set uniqueData = New Collection
    
    ' Get the current date for 60-day comparison
    currentDate = Date
    
    ' Loop through the cleansed data to populate the unique TM data collection
    On Error Resume Next ' Handle error if item already exists in the collection
    For i = 2 To lastRowSource
        key = wsSource.Cells(i, 9).Value ' TM Sign On (remove SHIFT | DEPT)
        uniqueData.Add key, key
    Next i
    On Error GoTo 0 ' Reset error handling
    
    ' Set headers in the destination worksheet (keeping only TM Sign On and Error Dates)
    wsDest.Name = "TM Data Extraction"
    wsDest.Cells(1, 1).Value = "TM SIGN ON"
    wsDest.Cells(1, 2).Value = "ERROR DATE(S)"
    wsDest.Cells(1, 3).Value = "FLAG"
    
    ' Populate the destination worksheet with the unique TM data
    lastRowDest = 2 ' Start from row 2
    For Each key In uniqueData
        wsDest.Cells(lastRowDest, 1).Value = key ' TM SIGN ON
        
        ' Extract error dates for this TM
        dateCount = 0
        mostRecentDate = DateSerial(1900, 1, 1) ' Initialize with a very old date
        For k = 2 To lastRowSource
            If wsSource.Cells(k, 9).Value = wsDest.Cells(lastRowDest, 1).Value Then
                ReDim Preserve errorDates(1 To dateCount + 1)
                errorDates(dateCount + 1) = wsSource.Cells(k, 2).Value
                If errorDates(dateCount + 1) > mostRecentDate Then
                    mostRecentDate = errorDates(dateCount + 1)
                End If
                dateCount = dateCount + 1
            End If
        Next k
        
        ' Concatenate error dates into a single cell
        For j = LBound(errorDates) To UBound(errorDates)
            wsDest.Cells(lastRowDest, 2).Value = wsDest.Cells(lastRowDest, 2).Value & _
                                                  IIf(wsDest.Cells(lastRowDest, 2).Value <> "", ", ", "") & _
                                                  errorDates(j)
        Next j
        
        ' Check if the most recent date is older than 60 days
        flagged = False
        If DateDiff("d", mostRecentDate, currentDate) <= 60 Then
            ' Check for rolling 60-day windows with 3 or more errors
            If dateCount >= 3 Then
                For k = 1 To UBound(errorDates)
                    If k + 2 <= UBound(errorDates) Then
                        If DateDiff("d", errorDates(k), errorDates(k + 2)) <= 60 Then
                            flagged = True
                            Exit For
                        End If
                    End If
                Next k
            End If
        End If
        
        ' Apply the flag if necessary
        If flagged Then
            wsDest.Cells(lastRowDest, 3).Value = "X"
        Else
            wsDest.Cells(lastRowDest, 3).Value = ""
        End If
        
        lastRowDest = lastRowDest + 1
    Next key
    
    ' Convert the extracted data into a table format
    Set tbl = wsDest.ListObjects.Add(xlSrcRange, wsDest.Range("A1").CurrentRegion, , xlYes)
    tbl.Name = "TM_Data_Table"
    
    ' Autofit the columns to match the content
    wsDest.Columns.AutoFit
    
    ' Align text to the left
    wsDest.Cells.HorizontalAlignment = xlLeft
    
    ' Apply conditional formatting for flagged rows
    flagColor = RGB(255, 200, 201) ' #ffc8c9 in RGB
    
    With wsDest.Range("A2:C" & lastRowDest - 1).FormatConditions.Add(Type:=xlExpression, Formula1:="=$C2=""X""")
        .Interior.Color = flagColor
    End With
    
    MsgBox "Data cleansing, extraction, flagging, and formatting completed.", vbInformation
End Sub


