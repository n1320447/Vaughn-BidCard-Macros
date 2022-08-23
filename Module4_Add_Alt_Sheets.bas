Attribute VB_Name = "Module4"
Sub Add_Alt_Sheets()
'Updateby Extendoffice
    Dim xRg As Excel.Range
    Dim wSh As Excel.Worksheet
    Dim wBk As Excel.Workbook
    Set wSh = ActiveSheet
    Set wBk = ActiveWorkbook
    Application.ScreenUpdating = False
    For Each xRg In wSh.Range("M1:M75")
        With wBk
            .Sheets.Add After:=.Sheets(.Sheets.Count)
            On Error Resume Next
            ActiveSheet.Name = "ALT " + xRg.Value
            If Err.Number = 1004 Then
              Debug.Print xRg.Value & " already used as a sheet name"
            End If
            On Error GoTo 0
        End With
        
    If IsEmpty(xRg) = True Then
    Exit For
    End If

    Next xRg
    Application.ScreenUpdating = True
    
    Application.DisplayAlerts = False
With ActiveWorkbook
    .Worksheets(.Worksheets.Count).Delete
End With
Application.DisplayAlerts = True

Application.Goto Reference:=Sheets("SHEET CREATOR").Range("A1")

End Sub



