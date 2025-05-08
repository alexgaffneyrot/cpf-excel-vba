Sub ResetSheetsFromMaster()
    Dim wsMaster As Worksheet, wsTarget As Worksheet, sht As Worksheet
    Dim rowIndex As Long, targetRow As Long
    Dim mrnCol As Long, nameCol As Long, consentCol As Long, hrcpCol As Long, cpCol As Long
    Dim cell As Range
    Dim sheetDict As Object, processedDict As Object
    Dim consentValue As String, hrcpValue As String, cpValue As String
    Dim key As String
    Dim allSheets As Collection
    Dim lastCol As Long, r As Long, c As Long
    Dim headerRow As Range

    On Error GoTo Cleanup

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsMaster = ThisWorkbook.Sheets("Master")
    Set sheetDict = CreateObject("Scripting.Dictionary")
    Set processedDict = CreateObject("Scripting.Dictionary")
    Set allSheets = New Collection

    ' Sheet mapping for Consent
    sheetDict.Add "Yes", "Consented"
    sheetDict.Add "Declined", "Declined"
    sheetDict.Add "Has Forms", "Has Forms"
    sheetDict.Add "Outborn", "Outborn"
    sheetDict.Add "Not Approached", "Not Approached"

    ' Add all target sheets
    For Each sht In ThisWorkbook.Sheets
        If sht.Name <> "Master" Then allSheets.Add sht
    Next sht

    ' Identify column positions
    For Each cell In wsMaster.Rows(1).Cells
        Select Case LCase(Trim(cell.Value))
            Case "mrn": mrnCol = cell.Column
            Case "name": nameCol = cell.Column
            Case "consent": consentCol = cell.Column
            Case "hrcp diagnosis": hrcpCol = cell.Column
            Case "cp diagnosis": cpCol = cell.Column
        End Select
    Next cell

    If mrnCol = 0 Or nameCol = 0 Or consentCol = 0 Or hrcpCol = 0 Or cpCol = 0 Then GoTo Cleanup

    ' Get last used column and header
    lastCol = wsMaster.Cells(1, wsMaster.Columns.Count).End(xlToLeft).Column
    Set headerRow = wsMaster.Range(wsMaster.Cells(1, 1), wsMaster.Cells(1, lastCol))

    ' Clear and reformat each sheet
    For Each sht In allSheets
        With sht
            .Rows("2:" & .Rows.Count).Clear
            headerRow.Copy Destination:=.Range("A1")
            ' Copy formats (not just values)
            headerRow.Copy
            .Range("A1").PasteSpecial Paste:=xlPasteFormats
        End With
    Next sht

    ' Loop through Master and copy rows
    For rowIndex = 2 To wsMaster.Cells(wsMaster.Rows.Count, nameCol).End(xlUp).Row
        If Trim(wsMaster.Cells(rowIndex, nameCol).Value) = "" Then GoTo NextRow

        key = Trim(wsMaster.Cells(rowIndex, mrnCol).Value) & "|" & Trim(wsMaster.Cells(rowIndex, nameCol).Value)
        If processedDict.exists(key) Then GoTo NextRow
        processedDict.Add key, True

        ' Copy to Consent sheet
        consentValue = Trim(wsMaster.Cells(rowIndex, consentCol).Value)
        If sheetDict.exists(consentValue) Then
            Set wsTarget = ThisWorkbook.Sheets(sheetDict(consentValue))
            targetRow = wsTarget.Cells(wsTarget.Rows.Count, nameCol).End(xlUp).Row + 1
            wsMaster.Rows(rowIndex).Copy Destination:=wsTarget.Rows(targetRow)
        End If

        ' Copy to HRCPCP if HRCP or CP = Yes
        hrcpValue = LCase(Trim(wsMaster.Cells(rowIndex, hrcpCol).Value))
        cpValue = LCase(Trim(wsMaster.Cells(rowIndex, cpCol).Value))
        If hrcpValue = "yes" Or cpValue = "yes" Then
            Set wsTarget = ThisWorkbook.Sheets("HRCPCP")
            targetRow = wsTarget.Cells(wsTarget.Rows.Count, nameCol).End(xlUp).Row + 1
            wsMaster.Rows(rowIndex).Copy Destination:=wsTarget.Rows(targetRow)
        End If

NextRow:
    Next rowIndex

    ' Reapply dropdowns and formatting from Master to each sheet
    For Each sht In allSheets
        With sht
            For c = 1 To lastCol
                If wsMaster.Cells(2, c).Validation.Type <> -1 Then
                    .Range(.Cells(2, c), .Cells(.Rows.Count, c)).Validation.Delete
                    wsMaster.Cells(2, c).Copy
                    .Cells(2, c).PasteSpecial Paste:=xlPasteValidation
                End If
            Next c
        End With
    Next sht

Cleanup:
    Application.CutCopyMode = False
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    MsgBox "Reset complete!"
End Sub

