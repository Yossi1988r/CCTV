' Constants for column references
Const COL_CABINT_C As String = "C"
Const COL_LOST_AND_BROKEN_F As String = "F"
Const COL_LOST_AND_BROKEN_G As String = "G"
Const COL_DATABASE_F As String = "F"
Const COL_DATABASE_G As String = "G"
Const COL_ENVELOPE_KEYS_F As String = "F"
Const COL_ENVELOPE_KEYS_G As String = "G"

Sub CompareKeysAndRecordSerialNumbers()
    On Error GoTo ErrorHandler

    ' Compare keys and record serial numbers from the LOST AND BROKEN KEYS sheet
    CompareKeysAndRecordSerialNumbersLostAndBroken

    ' Compare keys and record serial numbers from the DATABASE and ENVELOPE KEYS sheets
    CompareKeysAndRecordSerialNumbersDB

    ' Update Columns D, G, and H with missing values and write the results to the "Issues" sheet
    'UpdateColumnsDGH

    ' Find duplicate keys
    FindDuplicateKeys

    ' Find missing values in series
    FindMissingValuesInSeries

ExitSub:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbExclamation, "Error"
    Resume ExitSub
End Sub

Sub CompareKeysAndRecordSerialNumbersLostAndBroken()
    On Error GoTo ErrorHandler

    Dim wsCabint As Worksheet
    Dim wsLostAndBroken As Worksheet
    Dim cabintRow As Long
    Dim lostAndBrokenRow As Long
    Dim lastCabintRow As Long
    Dim lastLostAndBrokenRow As Long
    Dim serialNumbers As String

    Set wsCabint = ThisWorkbook.Worksheets("CABINT")
    Set wsLostAndBroken = ThisWorkbook.Worksheets("LOST AND BROKEN KEYS")

    lastCabintRow = wsCabint.Cells(wsCabint.Rows.count, COL_CABINT_C).End(xlUp).row
    lastLostAndBrokenRow = wsLostAndBroken.Cells(wsLostAndBroken.Rows.count, COL_LOST_AND_BROKEN_F).End(xlUp).row

    For cabintRow = 2 To lastCabintRow
        serialNumbers = ""

        For lostAndBrokenRow = 2 To lastLostAndBrokenRow
            If wsCabint.Cells(cabintRow, COL_CABINT_C).value = wsLostAndBroken.Cells(lostAndBrokenRow, COL_LOST_AND_BROKEN_F).value Then
                If serialNumbers = "" Then
                    serialNumbers = wsLostAndBroken.Cells(lostAndBrokenRow, COL_LOST_AND_BROKEN_G).value
                Else
                    serialNumbers = serialNumbers & "," & wsLostAndBroken.Cells(lostAndBrokenRow, COL_LOST_AND_BROKEN_G).value
                End If
            End If
        Next lostAndBrokenRow

        wsCabint.Cells(cabintRow, "H").value = serialNumbers
    Next cabintRow

ExitSub:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbExclamation, "Error"
    Resume ExitSub
End Sub

Sub CompareKeysAndRecordSerialNumbersDB()
    On Error GoTo ErrorHandler

    Dim wsCabint As Worksheet
    Dim wsDatabase As Worksheet
    Dim wsEnvelopeKeys As Worksheet
    Dim cabintRow As Long
    Dim databaseRow As Long
    Dim envelopeRow As Long
    Dim lastCabintRow As Long
    Dim lastDatabaseRow As Long
    Dim lastEnvelopeRow As Long
    Dim serialNumbers As String

    Set wsCabint = ThisWorkbook.Worksheets("CABINT")
      Set wsDatabase = ThisWorkbook.Worksheets("DATABASE")
    Set wsEnvelopeKeys = ThisWorkbook.Worksheets("ENVELOPE KEYS")

    lastCabintRow = wsCabint.Cells(wsCabint.Rows.count, COL_CABINT_C).End(xlUp).row
    lastDatabaseRow = wsDatabase.Cells(wsDatabase.Rows.count, COL_DATABASE_F).End(xlUp).row
    lastEnvelopeRow = wsEnvelopeKeys.Cells(wsEnvelopeKeys.Rows.count, COL_ENVELOPE_KEYS_F).End(xlUp).row

    For cabintRow = 2 To lastCabintRow
        serialNumbers = ""

        For databaseRow = 2 To lastDatabaseRow
            If wsCabint.Cells(cabintRow, COL_CABINT_C).value = wsDatabase.Cells(databaseRow, COL_DATABASE_F).value Then
                If serialNumbers = "" Then
                    serialNumbers = wsDatabase.Cells(databaseRow, COL_DATABASE_G).value
                Else
                    serialNumbers = serialNumbers & "," & wsDatabase.Cells(databaseRow, COL_DATABASE_G).value
                End If
            End If
        Next databaseRow

        For envelopeRow = 2 To lastEnvelopeRow
            If wsCabint.Cells(cabintRow, COL_CABINT_C).value = wsEnvelopeKeys.Cells(envelopeRow, COL_ENVELOPE_KEYS_F).value Then
                If serialNumbers = "" Then
                    serialNumbers = wsEnvelopeKeys.Cells(envelopeRow, COL_ENVELOPE_KEYS_G).value
                Else
                    serialNumbers = serialNumbers & "," & wsEnvelopeKeys.Cells(envelopeRow, COL_ENVELOPE_KEYS_G).value
                End If
            End If
        Next envelopeRow

        wsCabint.Cells(cabintRow, "G").value = serialNumbers
    Next cabintRow

ExitSub:
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred: " & Err.Description, vbExclamation, "Error"
    Resume ExitSub
End Sub


Sub FindDuplicateKeys()

    Dim wsCabint As Worksheet, wsIssues As Worksheet
    Dim cabintRow As Long, issuesRow As Long
    Dim snCABINT As Variant, snOUT As Variant, snLost As Variant
    Dim duplicateFound As Boolean
    Dim duplicateCount As Long

    Set wsCabint = ThisWorkbook.Worksheets("CABINT")
    Set wsIssues = ThisWorkbook.Worksheets("Issues")
    duplicateCount = 0

    issuesRow = 2

    ' Clear the Issues sheet
    wsIssues.Cells.ClearContents
    wsIssues.Cells(1, 1).value = "HOOK NO"
    wsIssues.Cells(1, 2).value = "KEY NO"
    wsIssues.Cells(1, 3).value = "S/N COMMENTS"
    wsIssues.Cells(1, 4).value = "S/N OUT"
    wsIssues.Cells(1, 5).value = "S/N LOST/BROKEN"

    ' Iterate through each row in CABINT
    For cabintRow = 2 To wsCabint.Cells(wsCabint.Rows.count, 1).End(xlUp).row
        snCABINT = Split(wsCabint.Cells(cabintRow, 4).value, ",")
        snOUT = Split(wsCabint.Cells(cabintRow, 7).value, ",")
        snLost = Split(wsCabint.Cells(cabintRow, 8).value, ",")
        
        duplicateFound = False
        
        For Each sn1 In snCABINT
            For Each sn2 In snOUT
                If sn1 = sn2 Then
                    duplicateFound = True
                    Exit For
                End If
            Next sn2
            If duplicateFound Then Exit For
            For Each sn2 In snLost
                If sn1 = sn2 Then
                    duplicateFound = True
                    Exit For
                End If
            Next sn2
            If duplicateFound Then Exit For
        Next sn1
        
        If duplicateFound Then
            wsIssues.Cells(issuesRow, 1).value = wsCabint.Cells(cabintRow, 1).value
            wsIssues.Cells(issuesRow, 2).value = wsCabint.Cells(cabintRow, 3).value
            wsIssues.Cells(issuesRow, 3).value = wsCabint.Cells(cabintRow, 4).value
            wsIssues.Cells(issuesRow, 4).value = wsCabint.Cells(cabintRow, 7).value
            wsIssues.Cells(issuesRow, 5).value = wsCabint.Cells(cabintRow, 8).value
            issuesRow = issuesRow + 1
             duplicateCount = duplicateCount + 1 ' Increment the duplicate count
        End If
    Next cabintRow
ThisWorkbook.Worksheets("Dashboard").Cells(8, 6).value = duplicateCount
End Sub

Sub FindMissingValuesInSeries()
    Dim missingValuesCount As Long
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim row As Long
    Dim values As Collection
    Dim missingValues As String
    
    Set ws = ThisWorkbook.Worksheets("CABINT")
    Dim wsIssues As Worksheet
    Set wsIssues = ThisWorkbook.Worksheets("Issues")
    missingValuesCount = 0

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).row
    
    For row = 1 To lastRow
        ' Combine the values in columns D, G, and H for each row
        Set values = New Collection
        Call AddValuesToCollection(values, ws.Cells(row, "D").value)
        Call AddValuesToCollection(values, ws.Cells(row, "G").value)
        Call AddValuesToCollection(values, ws.Cells(row, "H").value)
        
        ' Find the missing values in the series
        missingValues = GetMissingValues(values)
        
        ' Record the missing values in the last available row in the "Issues" column
     If missingValues <> "" Then
    Dim message As String

    message = "Missing " & CStr(ws.Cells(row, "A").value) & " - " & CStr(ws.Cells(row, "B").value) & " - " & CStr(ws.Cells(row, "C").value)
    wsIssues.Cells(wsIssues.Rows.count, "A").End(xlUp).Offset(1, 0).value = message
    wsIssues.Cells(wsIssues.Rows.count, "B").End(xlUp).Offset(1, 0).value = missingValues
    missingValuesCount = missingValuesCount + 1 ' Increment the missing values count
End If


    Next row
    'ThisWorkbook.Worksheets("Summary").Cells(2, 2).value = missingValuesCount
End Sub

Sub AddValuesToCollection(col As Collection, inputStr As String)
    Dim splitValues() As String
    Dim i As Long
    
    splitValues = Split(inputStr, ",")
    For i = LBound(splitValues) To UBound(splitValues)
        If IsNumeric(Trim(splitValues(i))) Then
            col.Add CLng(Trim(splitValues(i)))
        End If
    Next i
End Sub


Function GetMissingValues(values As Collection) As String
    Dim i As Long
    Dim maxVal As Long
    Dim missingVals As String
    Dim dict As Object
    
    ' Find the maximum value
    maxVal = 0
    For i = 1 To values.count
        maxVal = Application.Max(maxVal, values(i))
    Next i
    
    ' Find the missing values
    Set dict = CreateObject("Scripting.Dictionary")
    For i = 1 To values.count
        dict(values(i)) = True
    Next i
    
    missingVals = ""
    For i = 1 To maxVal
        If Not dict.Exists(i) Then
            missingVals = missingVals & IIf(missingVals = "", "", ", ") & i
        End If
    Next i
    
    GetMissingValues = missingVals
End Function

Sub UpdateColumnsDGH()
    Dim wsCabint As Worksheet
    Dim wsIssues As Worksheet
    Dim cabintRow As Long
    Dim issuesRow As Long
    Dim lastCabintRow As Long
    Dim DValues As String, GValues As String, HValues As String
    Dim updatedDValues As String
    Dim hasIssues As Boolean

    Set wsCabint = ThisWorkbook.Worksheets("CABINT")
    Set wsIssues = ThisWorkbook.Worksheets("Issues")
    lastCabintRow = wsCabint.Cells(wsCabint.Rows.count, "C").End(xlUp).row

    issuesRow = wsIssues.Cells(wsIssues.Rows.count, "A").End(xlUp).row + 1

    For cabintRow = 2 To lastCabintRow
        DValues = wsCabint.Cells(cabintRow, "D").value
        GValues = wsCabint.Cells(cabintRow, "G").value
        HValues = wsCabint.Cells(cabintRow, "H").value

        updatedDValues = CheckMissingValues(DValues, GValues, HValues)
        hasIssues = (updatedDValues <> DValues)

        If hasIssues Then
            wsIssues.Cells(issuesRow, "A").value = "Row " & cabintRow & " in CABINT"
            wsIssues.Cells(issuesRow, "B").value = "D: " & DValues
            wsIssues.Cells(issuesRow, "C").value = "G: " & GValues
            wsIssues.Cells(issuesRow, "D").value = "H: " & HValues
            wsIssues.Cells(issuesRow, "E").value = "Updated D: " & updatedDValues
            issuesRow = issuesRow + 1
        End If
    Next cabintRow
End Sub


Function CheckMissingValues(DValues As String, GValues As String, HValues As String) As String
    Dim values As Collection
    Dim missingValues As String

    Set values = New Collection
    Call AddValuesToCollection(values, DValues)
    Call AddValuesToCollection(values, GValues)
    Call AddValuesToCollection(values, HValues)

    missingValues = GetMissingValues(values)
    CheckMissingValues = IIf(missingValues = "", DValues, DValues & "," & missingValues)
End Function

