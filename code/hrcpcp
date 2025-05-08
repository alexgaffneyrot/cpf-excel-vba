Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wsMaster As Worksheet, wsThis As Worksheet, wsTarget As Worksheet
    Dim mrnCol As Long, consentCol As Long, hrcpCol As Long, cpCol As Long
    Dim lastRow As Long, masterRow As Variant
    Dim sheetDict As Object
    Dim consentValue As String, hrcpValue As String, cpValue As String
    Dim cell As Range

    On Error GoTo Cleanup

    If Target.Cells.CountLarge > 1 Then GoTo Cleanup
    If Target.Row = 1 Or IsEmpty(Target.Value) Then GoTo Cleanup

    Application.EnableEvents = False

    Set wsThis = Me
    Set wsMaster = ThisWorkbook.Sheets("Master")

    Set sheetDict = CreateObject("Scripting.Dictionary")
    sheetDict.Add "Yes", "Consented"
    sheetDict.Add "Declined", "Declined"
    sheetDict.Add "Has Forms", "Has Forms"
    sheetDict.Add "Outborn", "Outborn"
    sheetDict.Add "Not Approached", "Not Approached"

    ' Identify columns dynamically
    For Each cell In wsThis.Rows(1).Cells
        Select Case LCase(Trim(cell.Value))
            Case "mrn": mrnCol = cell.Column
            Case "consent": consentCol = cell.Column
            Case "hrcp diagnosis": hrcpCol = cell.Column
            Case "cp diagnosis": cpCol = cell.Column
        End Select
    Next cell

    If mrnCol = 0 Or consentCol = 0 Or hrcpCol = 0 Or cpCol = 0 Then GoTo Cleanup

    Dim mrnVal As String: mrnVal = Trim(wsThis.Cells(Target.Row, mrnCol).Value)
    If mrnVal = "" Then GoTo Cleanup

    masterRow = Application.Match(mrnVal, wsMaster.Columns(mrnCol), 0)
    If IsError(masterRow) Then masterRow = 0

    ' === Consent column changed ===
    If Target.Column = consentCol Then
        consentValue = Trim(Target.Value)
        If masterRow > 0 Then wsMaster.Cells(masterRow, consentCol).Value = consentValue

        If sheetDict.exists(consentValue) Then
            Set wsTarget = ThisWorkbook.Sheets(sheetDict(consentValue))
            lastRow = wsTarget.Cells(wsTarget.Rows.Count, mrnCol).End(xlUp).Row + 1

            If masterRow > 0 Then
                wsTarget.Rows(lastRow).Value = wsMaster.Rows(masterRow).Value
                wsMaster.Rows(masterRow).Copy
                wsTarget.Rows(lastRow).PasteSpecial Paste:=xlPasteFormats
            Else
                wsTarget.Rows(lastRow).Value = wsThis.Rows(Target.Row).Value
                wsThis.Rows(Target.Row).Copy
                wsTarget.Rows(lastRow).PasteSpecial Paste:=xlPasteFormats
            End If
        End If
    End If

    ' === HRCP/CP column changed ===
    If Target.Column = hrcpCol Or Target.Column = cpCol Then
        hrcpValue = LCase(Trim(wsThis.Cells(Target.Row, hrcpCol).Value))
        cpValue = LCase(Trim(wsThis.Cells(Target.Row, cpCol).Value))

        If masterRow > 0 Then
            wsMaster.Cells(masterRow, hrcpCol).Value = wsThis.Cells(Target.Row, hrcpCol).Value
            wsMaster.Cells(masterRow, cpCol).Value = wsThis.Cells(Target.Row, cpCol).Value
        End If

        ' Only keep row in HRCPCP if either value is "yes"
        If Not (hrcpValue = "yes" Or cpValue = "yes") Then
            wsThis.Rows(Target.Row).Delete
        End If
    End If

Cleanup:
    Application.CutCopyMode = False
    Application.EnableEvents = True
End Sub

