' =========================================================================
' Snippet: Data Consolidation Pipeline
' Purpose: Uses Arrays and Dictionaries to rapidly merge data from multiple shop floors.
' =========================================================================

Sub Get_Data_From_Barcode()
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim wsSSAR As Worksheet
    Dim existingIDs As Object
    Dim tempArray As Variant
    Dim r As Long
    Dim helperID As String
    
    Set wsSSAR = ThisWorkbook.Sheets("SSAR")
    Set existingIDs = CreateObject("Scripting.Dictionary")
    existingIDs.CompareMode = vbTextCompare
    
    ' 1. Load existing IDs into Dictionary for O(1) fast lookup
    Dim LastRowSSAR As Long
    LastRowSSAR = wsSSAR.Cells(wsSSAR.Rows.Count, "A").End(xlUp).Row
    If LastRowSSAR >= 2 Then
        Dim existingRangeValues As Variant
        existingRangeValues = wsSSAR.Range("A2:A" & LastRowSSAR).Value2
        For r = 1 To UBound(existingRangeValues, 1)
            helperID = CStr(existingRangeValues(r, 1))
            If helperID <> "" And Not existingIDs.Exists(helperID) Then
                existingIDs.Add helperID, Null
            End If
        Next r
    End If
    
    ' 2. Batch process new files from shared drives (Paths anonymized)
    Dim fileNames As Variant
    fileNames = Array("[SHARED_DRIVE]\Body_OUT.xlsm", "[SHARED_DRIVE]\Paint_IN.xlsm")
    
    ' ... (Code logic to loop through workbooks, load data into Memory Arrays, 
    ' check against Dictionary, and write back in one single batch operation) ...

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
End Sub
