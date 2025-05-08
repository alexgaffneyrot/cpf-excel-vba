Private Sub Worksheet_Change(ByVal Target As Range)
    Dim wsMaster As Worksheet: Set wsMaster = Me
    Dim consentCol As Long, hrcpCol As Long, cpCol As Long, mrnCol As Long, nameCol As Long
    Dim changedCol As Long: changedCol = Target.Column
    Dim rowNum As Long: rowNum = Target.Row
    Dim val As Variant, key As String
    Dim col As Range, sheetName As Variant
    Dim sheetDict As Object, wsTarget As Worksheet, wsHrcp As Worksheet
    Dim findRow As Range, destRow As Long
    Dim mrnVal As String, nameVal As String
    Dim hrcpVal As String, cpVal As String
    Dim inHrcp As Boolean: inHrcp = False

    If Target.Cells.Count > 1 Or rowNum = 1 Then Exit Sub
    Application.EnableEvents = False

    ' === Column lookups ===
    For Each col In wsMaster.Rows(1).Cells
        Select Case LCase(Trim(col.Value))
            Case "mrn": mrnCol = col.Column
            Case "name": nameCol = col.Column
            Case "consent": consentCol = col.Column
            Case "hrcp diagnosis": hrcpCol = col.Column
            Case "cp diagnosis": cpCol = col.Column
        End Select
    Next col

    If mrnCol = 0 Or nameCol = 0 Or consentCol = 0 Or hrcpCol = 0 Or cpCol = 0 Then GoTo Cleanup

    mrnVal = Trim(wsMaster.Cells(rowNum, mrnCol).Value)
    nameVal = Trim(wsMaster.Cells(rowNum, nameCol).Value)
    If mrnVal = "" And nameVal = "" Then GoTo Cleanup

    ' Create composite key using MRN + Name (for fallback if MRN is missing)
    key = LCase(mrnVal & "|" & nameVal)

    Set sheetDict = CreateObject("Scripting.Dictionary")
    sheetDict.Add "yes", "Consented"
    sheetDict.Add "declined", "Declined"
    sheetDict.Add "has forms", "Has Forms"
    sheetDict.Add "outborn", "Outborn"
    sheetDict.Add "not approached", "Not Approached"

    ' === Handle Consent change ===
    If changedCol = consentCol Then
        val = LCase(Trim(Target.Value))
        If sheetDict.exists(val) Then
            Set wsTarget = ThisWorkbook.Sheets(sheetDict(val))

            ' Remove from all other consent sheets
            For Each sheetName In sheetDict.Items
                If sheetName <> sheetDict(val) Then
                    With ThisWorkbook.Sheets(sheetName)
                        Set findRow = .UsedRange.Columns(mrnCol).Find(mrnVal, LookIn:=xlValues, LookAt:=xlWhole)
                        If findRow Is Nothing And mrnVal = "" Then
                            Set findRow = .UsedRange.Columns(nameCol).Find(nameVal, LookIn:=xlValues, LookAt:=xlWhole)
                        End If
                        If Not findRow Is Nothing Then .Rows(findRow.Row).Delete
                    End With
                End If
            Next sheetName

            ' Add or update in target sheet
            With wsTarget
                Set findRow = .UsedRange.Columns(mrnCol).Find(mrnVal, LookIn:=xlValues, LookAt:=xlWhole)
                If findRow Is Nothing And mrnVal = "" Then
                    Set findRow = .UsedRange.Columns(nameCol).Find(nameVal, LookIn:=xlValues, LookAt:=xlWhole)
                End If
                If Not findRow Is Nothing Then
                    destRow = findRow.Row
                Else
                    destRow = .Cells(.Rows.Count, mrnCol).End(xlUp).Row + 1
                End If
                .Rows(destRow).Value = wsMaster.Rows(rowNum).Value
                wsMaster.Rows(rowNum).Copy
                .Rows(destRow).PasteSpecial Paste:=xlPasteFormats
            End With
        End If
    End If

    ' === HRCPCP Logic ===
    hrcpVal = LCase(Trim(wsMaster.Cells(rowNum, hrcpCol).Value))
    cpVal = LCase(Trim(wsMaster.Cells(rowNum, cpCol).Value))

    If hrcpVal = "yes" Or cpVal = "yes" Then inHrcp = True

    Set wsHrcp = ThisWorkbook.Sheets("HRCPCP")
    Set findRow = wsHrcp.UsedRange.Columns(mrnCol).Find(mrnVal, LookIn:=xlValues, LookAt:=xlWhole)
    If findRow Is Nothing And mrnVal = "" Then
        Set findRow = wsHrcp.UsedRange.Columns(nameCol).Find(nameVal, LookIn:=xlValues, LookAt:=xlWhole)
    End If

    If inHrcp Then
        If findRow Is Nothing Then
            destRow = wsHrcp.Cells(wsHrcp.Rows.Count, mrnCol).End(xlUp).Row + 1
            wsHrcp.Rows(destRow).Value = wsMaster.Rows(rowNum).Value
            wsMaster.Rows(rowNum).Copy
            wsHrcp.Rows(destRow).PasteSpecial Paste:=xlPasteFormats
        Else
            wsHrcp.Rows(findRow.Row).Value = wsMaster.Rows(rowNum).Value
        End If
    Else
        If Not findRow Is Nothing Then wsHrcp.Rows(findRow.Row).Delete
    End If

Cleanup:
    Application.CutCopyMode = False
    Application.EnableEvents = True
End Sub

