Private Sub Worksheet_Change(ByVal Target As Range)
    On Error GoTo Cleanup

    If Target.Cells.CountLarge > 1 Then GoTo Cleanup ' <-- Use CountLarge to avoid overflow
    If Target.Row = 1 Then GoTo Cleanup

    Application.EnableEvents = False

    Dim wsThis As Worksheet, wsMaster As Worksheet, wsTarget As Worksheet
    Dim sheetDict As Object
    Dim consentCol As Long, mrnCol As Long, masterRow As Long, lastRow As Long
    Dim consentValue As String, mrn As String
    Dim cell As Range

    Set wsThis = Me
    Set wsMaster = ThisWorkbook.Sheets("Master")

    Set sheetDict = CreateObject("Scripting.Dictionary")
    sheetDict.Add "Yes", "Consented"
    sheetDict.Add "Declined", "Declined"
    sheetDict.Add "Has Forms", "Has Forms"
    sheetDict.Add "Outborn", "Outborn"
    sheetDict.Add "Not Approached", "Not Approached"

    ' Identify MRN and Consent columns
    For Each cell In wsMaster.Rows(1).Cells
        If Trim(LCase(cell.Value)) = "mrn" Then mrnCol = cell.Column
        If Trim(LCase(cell.Value)) = "consent" Then consentCol = cell.Column
    Next cell

    If Target.Column = consentCol Then
        consentValue = Target.Value
        If sheetDict.exists(consentValue) Then
            mrn = wsThis.Cells(Target.Row, mrnCol).Value

            ' Find row in Master
            masterRow = 0
            For Each cell In wsMaster.Columns(mrnCol).Cells
                If cell.Value = mrn Then
                    masterRow = cell.Row
                    Exit For
                End If
            Next cell

            If masterRow > 0 Then
                Set wsTarget = ThisWorkbook.Sheets(sheetDict(consentValue))
                lastRow = wsTarget.Cells(wsTarget.Rows.Count, mrnCol).End(xlUp).Row + 1

                ' Update master sheet
                wsMaster.Cells(masterRow, consentCol).Value = consentValue

                ' Copy row to correct sheet
                wsTarget.Rows(lastRow).Value = wsMaster.Rows(masterRow).Value

                ' Delete from current sheet
                wsThis.Rows(Target.Row).Delete
            End If
        End If
    End If

Cleanup:
    Application.EnableEvents = True
End Sub

