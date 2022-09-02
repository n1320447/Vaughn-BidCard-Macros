Sub Create_Summary()
    'Func1()
    '   Copy/Pastework categories from Sheet Creator
    '
    Sheets("SHEET CREATOR").Select
    Dim x As Integer
    
    'Application.ScreenUpdating = False
    
    NumRows = Range("A" & Rows.Count).End(xlUp).Row
    Dim arr() As String
    For x = 1 To NumRows
        Sheets("SHEET CREATOR").Select
        y = "A" + CStr(x)
        Range(y).Select
        Dim var As Variant
        var = Range(y).Value
        ReDim Preserve arr(x)
        arr(x) = var
    Next
    
  
        'Debug.Print j
        Sheets("SHEET CREATOR").Select
        Range("A1", "A" + CStr(NumRows - 1)).Copy
        Sheets("SUMMARY").Select
        Range("A2").Select
        ActiveSheet.Paste

    
    
    'Func2()
    '   Assign 'Subcontractor selected' column to respective sheet on excel spreadSheet
    '
    'Func3()
    '   Assign 'Selected Sub $ Amount' column to respective cell in correct sheet on excel spreadSheet
    '
    '
    '
    '
    '
    
    
    

End Sub
