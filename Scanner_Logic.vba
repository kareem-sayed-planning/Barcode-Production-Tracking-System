' =========================================================================
' Snippet: Real-time Barcode Capture & Shift Calculation
' Purpose: Validates scanned VIN, calculates production shift, and saves record.
' =========================================================================

Private Sub btsave_Click()
    Dim wsList As Worksheet, wsDatabase As Worksheet
    Dim chassisNo As String
    Dim workBittern As Integer
    Dim workTime As Double
    Dim shiftValue As String
    
    Set wsList = ThisWorkbook.Sheets("List")
    Set wsDatabase = ThisWorkbook.Sheets("Database")
    chassisNo = Trim(Me.chs.Value)
    
    ' 1. Validation Logic
    If Len(chassisNo) <> 17 Then
        MsgBox "Please check the chassis number! Must be 17 characters.", vbExclamation, "Invalid Input"
        Exit Sub
    End If
    
    ' Prevent Duplicates
    If Not IsError(Application.Match(chassisNo, wsList.Range("B:B"), 0)) Then
        MsgBox "Chassis number already exists!", vbCritical, "Duplicate Entry"
        Exit Sub
    End If
    
    ' 2. Shift Calculation Logic
    workBittern = wsList.Cells(2, 15).Value
    workTime = TimeValue(Format(Now, "hh:mm AM/PM"))
    If Not IsNumeric(workBittern) Then workBittern = 0
    
    Select Case workBittern
        Case 0
            If workTime >= TimeValue("07:00") And workTime < TimeValue("16:00") Then
                shiftValue = "Shift1"
            Else
                shiftValue = "Shift2"
            End If
        Case 1
            If workTime >= TimeValue("07:00") And workTime < TimeValue("17:00") Then
                shiftValue = "Shift1"
            Else
                shiftValue = "Shift2"
            End If
        Case Else
            shiftValue = "Unknown"
    End Select
    
    ' 3. Save Record (Snippet truncated for portfolio)
    ' ... (Code to write to next empty row in Database) ...
    
    MsgBox "Data saved successfully!", vbInformation, "Success"
    Call UpdateListBox
End Sub
