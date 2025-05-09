Sub ResetSheetsFromMaster_Optimized()
    Dim wsMaster As Worksheet, wsTarget As Worksheet
    Dim sheetDict As Object, sheetName As Variant
    Dim lastRow As Long, destRow As Long
    Dim mrnCol As Long, nameCol As Long, consentCol As Long, hrcpCol As Long, cpCol As Long
    Dim row As Long
    Dim key As String
    Dim dataDict As Object
    Dim cell As Range

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    Set wsMaster = ThisWorkbook.Sheets("Master")
    Set sheetDict = CreateObject("Scripting.Dictionary")
    sheetDict.Add "Consented", "yes"
    sheetDict.Add "Declined", "declined"
    sheetDict.Add "Not Approached", "not approached"
    sheetDict.Add "Has Forms", "has forms"
    sheetDict.Add "Outborn", "outborn"
    sheetDict.Add "Lost to FU", "lost to f/u"
    sheetDict.Add "RIP", "rip"

    ' Identify column indexes
    Debug.Print "-- Column Identification --"
    For Each cell In wsMaster.Rows(1).Cells
        Debug.Print "Found Header: " & cell.Value
        Select Case LCase(Trim(cell.Value))
            Case "mrn": mrnCol = cell.Column
            Case "name": nameCol = cell.Column
            Case "consent": consentCol = cell.Column
            Case "hrcp diagnosis": hrcpCol = cell.Column
            Case "cp diagnosis": cpCol = cell.Column
        End Select
    Next cell

    Debug.Print "MRN Column: " & mrnCol & " | Name Column: " & nameCol & " | Consent Column: " & consentCol
    Debug.Print "HRCP Column: " & hrcpCol & " | CP Column: " & cpCol

    If mrnCol = 0 Or nameCol = 0 Or consentCol = 0 Then
        MsgBox "Columns missing or header names do not match. Please check the headers.", vbCritical
        GoTo Cleanup
    End If

    ' Find last row in the Master sheet dynamically
    lastRow = wsMaster.Cells(wsMaster.Rows.Count, mrnCol).End(xlUp).Row
    Debug.Print "Last Row Identified: " & lastRow

    If lastRow < 2 Then
        MsgBox "No data found in the Master sheet.", vbExclamation
        GoTo Cleanup
    End If

    ' Clear all target sheets and restore headers
    For Each sheetName In sheetDict.Keys
        On Error Resume Next
        Set wsTarget = ThisWorkbook.Sheets(sheetName)
        On Error GoTo 0
        If Not wsTarget Is Nothing Then
            wsTarget.Cells.ClearContents
            wsMaster.Rows(1).Copy Destination:=wsTarget.Rows(1)
            Debug.Print "Cleared and reset sheet: " & sheetName
        Else
            Debug.Print "Sheet not found: " & sheetName
        End If
    Next sheetName

    ' Clear HRCPCP sheet
    Set wsTarget = ThisWorkbook.Sheets("HRCPCP")
    wsTarget.Cells.ClearContents
    wsMaster.Rows(1).Copy Destination:=wsTarget.Rows(1)
    Debug.Print "Cleared and reset sheet: HRCPCP"

    ' Initialize Dictionary for duplicate tracking
    Set dataDict = CreateObject("Scripting.Dictionary")

    ' Loop through master sheet rows
    Debug.Print "-- Row Loop Start --"
    For row = 2 To lastRow
        Dim mrnVal As String: mrnVal = LCase(Trim(wsMaster.Cells(row, mrnCol).Value))
        Dim nameVal As String: nameVal = LCase(Trim(wsMaster.Cells(row, nameCol).Value))
        Dim consentVal As String: consentVal = LCase(Trim(wsMaster.Cells(row, consentCol).Value))
        Dim hrcpVal As String: hrcpVal = LCase(Trim(wsMaster.Cells(row, hrcpCol).Value))
        Dim cpVal As String: cpVal = LCase(Trim(wsMaster.Cells(row, cpCol).Value))

        Debug.Print "Row " & row & " | MRN: " & mrnVal & " | Name: " & nameVal & " | Consent: " & consentVal
        Debug.Print "HRCP: " & hrcpVal & " | CP: " & cpVal

        If mrnVal <> "" And nameVal <> "" Then
            key = nameVal & "|" & mrnVal

            If Not dataDict.Exists(key) Then
                dataDict.Add key, True

                ' Consent-based sheets
                If sheetDict.Exists(consentVal) Then
                    Set wsTarget = ThisWorkbook.Sheets(sheetDict(consentVal))
                    destRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row + 1
                    wsMaster.Rows(row).Copy Destination:=wsTarget.Rows(destRow)
                    Debug.Print "Copied to sheet: " & sheetDict(consentVal)
                End If

                ' HRCPCP sheet logic
                If hrcpVal = "yes" Or cpVal = "yes" Then
                    Set wsTarget = ThisWorkbook.Sheets("HRCPCP")
                    destRow = wsTarget.Cells(wsTarget.Rows.Count, 1).End(xlUp).Row + 1
                    wsMaster.Rows(row).Copy Destination:=wsTarget.Rows(destRow)
                    Debug.Print "Copied to HRCPCP"
                End If
            End If
        End If
    Next row

Cleanup:
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Sheets reset successfully from Master.", vbInformation
End Sub
