
Sub Copy_Paste_WBSCards_to_Sheets()


'LOOP FOR CYCLING THROUGH SHEET NAMES
    Sheets("SHEET CREATOR").Select
    Dim x As Integer
    Application.ScreenUpdating = False
    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    'Debug.Print NumRows
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
        
'START OF RANGE LOOP
    Sheets("CARD DUMP").Select
    Dim SrchRng As Range, cel As Range
    Set SrchRng = Range("A1:P1000000")
    Start = 0
    end1 = 0
    wallet = 0
    Start2 = 0
    end2 = 0
    end3 = 0
    Banana = 0
    phone = 0
    monkey = 0
    mouse = 0
    Dim arr1() As String
    For Each cel In SrchRng
        If InStr(1, cel.Value, "CARD HOLDER") > 0 Then

            Start = "A" + CStr(cel.Row - 3)
            Start2 = CStr(cel.Row + 10)
            Start3 = CStr(cel.Row + 2)
            ReDim Preserve arr1(wallet * 2)
            arr1(wallet * 2) = Start

        End If

        If InStr(1, cel.Value, "Grand Total") > 0 Then

            end1 = "Z" + CStr(cel.Row)
            end2 = Trim(Replace(end1, "Z", ""))
            end2 = CStr(cel.Row)
            end3 = CStr(cel.Row + 1)

            ReDim Preserve arr1((wallet * 2) + 1)
            arr1((wallet * 2) + 1) = end1
            wallet = wallet + 1

        End If
      
    
    Next cel



    For j = 0 To (NumRows - 1)
        Sheets("CARD DUMP").Select
        Range(arr1(j * 2), arr1((j * 2) + 1)).Copy
        Sheets(arr(j + 1)).Select
        Range("A1").Select
        ActiveSheet.Paste

        Dim rng3 As Range
        Set rng3 = Range("S1:S100") ' Identify your range
        v = 0
            For Each w In rng3.Cells
                If w.Value <> "" And w.Value = "Contact:" Then '<--- Will search if the cell is not empty and not equal to phrase. If you want to check empty cells too remove c.value <> ""
                    'MsgBox (c.Address & "found") '<---- Your code goes here
                    v = w.Row
                    Rows(v).EntireRow.Delete
                End If
            Next w
       ' Debug.Print "Just outside CardHolder find loop"
        Dim rng2 As Range
        Set rng2 = Range("C2:C100000")
        x = 0
            For Each d In rng2.Cells
                If d.Value <> "" And d.Value = "CARD HOLDER:" Then
                'Debug.Print "Found Holder"
                x = d.Row 'Card Holder row
                x2 = x - 1 'Bid Date row
                x3 = x - 2 'Job row
                x4 = x + 7 'CATEGORY SCOPE ROW
                x5 = x + 1 'ADDENDA ROW
                x6 = x + 2 'BONDRATE ROW
                x7 = x + 3 'INSURANCE ROW
                x8 = x + 4 'HUB ROW
                x9 = x + 5 'WAGE RATE ROW
                x10 = x + 6 'INCLUDES TAXES ROW
                x11 = x + 7 'QTY ROW
                x19 = x + 9 'FIRST SCOPE ITEM
                Application.DisplayAlerts = False
                Range("A" + (CStr(x3)) + ":D" + (CStr(x3))).Merge
                Range("A" + (CStr(x2)) + ":D" + (CStr(x2))).Merge
                Range("A" + (CStr(x)) + ":D" + (CStr(x))).Merge
                Range("E" + (CStr(x3)) + ":R" + (CStr(x3))).Merge
                Range("E" + (CStr(x2)) + ":R" + (CStr(x2))).Merge
                Range("E" + (CStr(x)) + ":R" + (CStr(x))).Merge
                Range("S" + (CStr(x))) = "Contact:"
                Range("S" + (CStr(x)) + ":U" + (CStr(x))).Merge
                Range("T" + (CStr(x4))).MergeCells = False
                Range("A" + (CStr(x5)) + ":D" + (CStr(x5))).Merge 'ADDENDA ROW
                Range("A" + (CStr(x6)) + ":R" + (CStr(x6))).Merge 'BOND ROW
                Range("A" + (CStr(x7)) + ":R" + (CStr(x7))).Merge 'INSURANCE ROW
                Range("A" + (CStr(x8)) + ":R" + (CStr(x8))).Merge 'HUB ROW
                Range("A" + (CStr(x9)) + ":R" + (CStr(x9))).Merge 'WAGE RATE ROW
                Range("A" + (CStr(x10)) + ":R" + (CStr(x10))).Merge 'INCLUDES TAXES ROW
                Range("E" + (CStr(x5)) + ":R" + (CStr(x5))).Merge

                Range("H" + (CStr(x4)) + ":K" + (CStr(x4))).Merge 'QTY MERGE
                Range("L" + (CStr(x4)) + ":N" + (CStr(x4))).Merge 'UNIT MERGE

                Range("S" + (CStr(x4)) + ":U" + (CStr(x4))).Merge 'TOTAL MERGE
                Range("P" + (CStr(x4)) + ":R" + (CStr(x4))).MergeCells = False 'RATE UNMERGE
                Range("P" + (CStr(x4)) + ":Q" + (CStr(x4))).Merge 'RATE MERGE
                Range("A:C").ColumnWidth = 0.1 'MINIMIZE A-C

                Range("D:D").ColumnWidth = 20 'ADJUST COLUMN
                Range("F:F").ColumnWidth = 50 'ADJUST COLUMN

                Range("H:J").ColumnWidth = 2
                Range("K:L").ColumnWidth = 3

                Range("P:Q").ColumnWidth = 5
                Range("R:R").ColumnWidth = 0.1

                Range("N:N").Delete

                Range("M:M").ColumnWidth = 7
                Range("11:11").RowHeight = 15

                Range("N:N").Delete

                Range("A11") = "CATEGORY/SCOPE"
                Range("A" + (CStr(x4)) + ":G" + (CStr(x4))).Merge 'CATEGORY/SCOPE MERGE
                
                Range("K:K").ColumnWidth = 7
                
                Columns("U:Y").Insert Shift:=xlToRight
                Range("S:S").Delete
                Range("K:K").Delete
                Range("G:G").Delete
                Range("M:M").Delete
                Range("P:P").Delete
                
                Range("L:L").ColumnWidth = 7.5
                Range("H:H").ColumnWidth = 4
                
                
                    Application.ScreenUpdating = True

                Set rng = Range("A12:P100000") ' Identify your range finds row to put sum formula in
                d = 0
                    For Each cel In rng.Cells
                        If InStr(1, cel.Value, "Grand Total") > 0 Then
                            'MsgBox (c.Address & "found") '<---- Your code goes here
                            d = cel.Row 'Row to put SUM formula in
                            d2 = d - 1 'lower bound of SUM formula
                            d3 = x + 7 'upper bound of SUM formula
                            'Debug.Print d2
                            Range("Q" + (CStr(d))).Value = "=Sum(Q" + (CStr(d3)) + ":Q" + (CStr(d2)) + ")"
                            Range("R" + (CStr(d))).Value = "=Sum(R" + (CStr(d3)) + ":R" + (CStr(d2)) + ")"
                            Range("S" + (CStr(d))).Value = "=Sum(S" + (CStr(d3)) + ":S" + (CStr(d2)) + ")"
                            Range("T" + (CStr(d))).Value = "=Sum(T" + (CStr(d3)) + ":T" + (CStr(d2)) + ")"
                            Range("U" + (CStr(d))).Value = "=Sum(U" + (CStr(d3)) + ":U" + (CStr(d2)) + ")"
                            Range("V" + (CStr(d))).Value = "=Sum(V" + (CStr(d3)) + ":V" + (CStr(d2)) + ")"
                            Range("W" + (CStr(d))).Value = "=Sum(W" + (CStr(d3)) + ":W" + (CStr(d2)) + ")"
                            Range("X" + (CStr(d))).Value = "=Sum(X" + (CStr(d3)) + ":X" + (CStr(d2)) + ")"
                        End If
                    Next cel
                    
                    
                    
                    'Debug.Print B
                    For Each cel In rng.Cells
                        If InStr(1, cel.Value, "Grand Total") > 0 Then
                        d = cel.Row 'Gives bottom of card aka GRAND TOTAL ROW
                        B = x - 2 'Gives Top Left of Card aka JOB: ROW
                        Range("A" + (CStr(B)) + ":X" + (CStr(d))).Borders.LineStyle = xlContinuous ' Gives entire card lines
                         Range("A" + (CStr(B)) + ":D" + (CStr(d))).BorderAround xlContinuous, xlThin 'Gives Border to Job to Card Holder area
                        Range("A" + (CStr(B)) + ":X" + (CStr(d))).BorderAround xlContinuous, xlThick 'Gives Border to Card
                        Range("Q" + (CStr(B)) + ":X" + (CStr(d))).BorderAround xlContinuous, xlThick 'Gives Border to Q-X
                        Range("A" + (CStr(x5)) + ":X" + (CStr(x10))).BorderAround xlContinuous, xlThick 'Gives Border to Addend-Incl Tax cells
                        End If
                    Next

                    Range("P:P").Delete
                    Range("K:K").ColumnWidth = 3
                Dim rngx As Range
                Set rngx = Range("V12:V100000") ' Identify your range
                d = 0
                    For Each c In rngx.Cells
                        If c.Value = "Page 2 of " Then '<--- Will search if the cell is not empty and not equal to phrase. If you want to check empty cells too remove c.value <> ""
                            d = c.Row
                            Rows(d).EntireRow.Delete
                            'Debug.Print d; "test"
                        End If
                    Next c
                    For Each c In rngx.Cells
                        If c.Value = "Page 3 of " Then '<--- Will search if the cell is not empty and not equal to phrase. If you want to check empty cells too remove c.value <> ""
                            'MsgBox (c.Address & "found") '<---- Your code goes here
                            d = c.Row
                            Rows(d).EntireRow.Delete
                        End If
                    Next c

                Range("P:W").ColumnWidth = 18
                Range("A:W").Font.Name = "Calibri"
                Range("N" + (CStr(x)) + ":O" + (CStr(x))).HorizontalAlignment = xlRight
                Range("10:10").Cells.Font.Size = "11"
                'adds orange highlight to cards
                Range("P" + "11" + ":W" + "11").Interior.ColorIndex = 44
                Range("P:W").NumberFormat = "$#,##0"
                Range("P:W").HorizontalAlignment = xlCenter
                Range("A" + (CStr(x3)) + ":X" + (CStr(x10))).NumberFormat = "General"
                
                 For Each cel In rng.Cells
                        If InStr(1, cel.Value, "Grand Total") > 0 Then
                        d = cel.Row 'Gives bottom of card aka GRAND TOTAL ROW
                        d2 = d + 4
                        d3 = d + 5
                        Range("P" + (CStr(d2))).HorizontalAlignment = xlLeft
                        Range("P" + (CStr(d3))).HorizontalAlignment = xlLeft
                        Range("G" + (CStr(d2))).Value = "Subcontractor in Add/Cut is:"
                        Range("G" + (CStr(d3))).Value = "Bid Amount in Add/Cut is:"
                        Range("P" + (CStr(d2))).Value = "(Only Bid Captain fills in, let them know if this does not match bid card.)"
                        Range("P" + (CStr(d3))).Value = "(Only Bid Captain fills in, let them know if this does not match bid card.)"
                        Range("P" + (CStr(d2))).Font.Size = "10"
                        Range("P" + (CStr(d3))).Font.Size = "10"
                        Range("G" + (CStr(d2))).Font.Size = "10"
                        Range("G" + (CStr(d3))).Font.Size = "10"
                        Range("N" + (CStr(d2)) + ":O" + (CStr(d2))).Merge
                        Range("N" + (CStr(d3)) + ":O" + (CStr(d3))).Merge
                        Range("N" + (CStr(d3)) + ":O" + (CStr(d3))).NumberFormat = "$#,##0"
                        Range("N" + (CStr(d2)) + ":O" + (CStr(d2))).BorderAround xlContinuous, xlThick
                        Range("N" + (CStr(d3)) + ":O" + (CStr(d3))).BorderAround xlContinuous, xlThick
                        End If
                    Next
                
                    End If
            Next d

    Next


        Application.Goto Reference:=Sheets("SHEET CREATOR").Range("A1")
        MsgBox ("Base Bid Cards have now been copied to each sheet.")
End Sub







