Attribute VB_Name = "Module2"
Sub Copy_Paste_Cards_to_Sheets()


'LOOP FOR CYCLING THROUGH SHEET NAMES
    Sheets("SHEET CREATOR").Select
    Dim x As Integer
    Application.ScreenUpdating = False
    ' Set numrows = number of rows of data.
    NumRows = Range("A1", Range("A1").End(xlDown)).Rows.Count
    'Debug.Print NumRows
    ' Select cell a1.
    'Range("A1").Select
    ' Establish "For" loop to loop "numrows" number of times.
    'var = ""
    Dim arr() As String
    For x = 1 To NumRows
        Sheets("SHEET CREATOR").Select
        y = "A" + CStr(x)
        Range(y).Select
        ' Insert your code here.
        ' Selects cell down 1 row from active cell.
        ' ActiveCell.Offset(1, 0).Select
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
'    Dim arr2() As String
'    Dim arr3() As String
'    Dim arr4() As String
'    Dim arr5() As String
    For Each cel In SrchRng
        If InStr(1, cel.Value, "CARD HOLDER") > 0 Then

            Start = "A" + CStr(cel.Row - 3)
            'Start2 = Trim(Replace(Start, "A", ""))
            Start2 = CStr(cel.Row + 10)
            Start3 = CStr(cel.Row + 2)
            ReDim Preserve arr1(wallet * 2)
            arr1(wallet * 2) = Start
            
'            ReDim Preserve arr4(monkey + 10)
'            arr4(monkey + 10) = Start2
'            monkey = monkey + 1

'            ReDim Preserve arr5(mouse + 2) 'Gives top left corner for thick vertical border
'            arr5(mouse + 2) = Start3
'            mouse = mouse + 1
            'Debug.Print Start; "START"

        End If

        If InStr(1, cel.Value, "Grand Total") > 0 Then


            end1 = "Z" + CStr(cel.Row)
            end2 = Trim(Replace(end1, "Z", ""))
            end2 = CStr(cel.Row)
            end3 = CStr(cel.Row + 1)

            ReDim Preserve arr1((wallet * 2) + 1)
            arr1((wallet * 2) + 1) = end1
            wallet = wallet + 1
            Debug.Print end1
            
'            ReDim Preserve arr2(Banana + 1)
'            arr2(Banana + 1) = end2
'            Banana = Banana + 1
            
'            ReDim Preserve arr3(phone + 1)
'            arr3(phone + 1) = end3
'            phone = phone + 1
'            'Debug.Print end1; "END"
        End If
      
    
    Next cel



        


'    Next
    For j = 0 To (NumRows - 1)
        Sheets("CARD DUMP").Select
        Range(arr1(j * 2), arr1((j * 2) + 1)).Copy
        Sheets(arr(j + 1)).Select
        'Call AddOutsideBorders(ActiveWorkbook.Worksheets(arr(j + 1)).Range("A3:S10"))
        Range("A1").Select
        'Range("W:Z").ColumnWidth = 14
        ActiveSheet.Paste
        'Range("A3:T5").BorderAround xlContinuous, xlMedium
        'ActiveSheet.Range("A3:T10").BorderAround xlContinuous, xlThick
        'ActiveSheet.Range("A3:T11").BorderAround xlContinuous, xlThick

        Dim rng3 As Range
        Set rng3 = Range("S1:S100") ' Identify your range
        v = 0
            For Each w In rng3.Cells
                If w.Value <> "" And w.Value = "Contact:" Then '<--- Will search if the cell is not empty and not equal to phrase. If you want to check empty cells too remove c.value <> ""
                    'MsgBox (c.Address & "found") '<---- Your code goes here
                    v = w.Row
                    Rows(v).EntireRow.Delete
                    'Debug.Print v
                End If
            Next w

        Dim rng2 As Range
        Set rng2 = Range("B2:B100000")
        x = 0
            For Each d In rng2.Cells
                If d.Value <> "" And d.Value = "CARD HOLDER:" Then
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

                Range("A" + (CStr(x3)) + ":D" + (CStr(x3))).Merge
                Range("A" + (CStr(x2)) + ":D" + (CStr(x2))).Merge
                Range("A" + (CStr(x)) + ":D" + (CStr(x))).Merge
                Range("E" + (CStr(x3)) + ":R" + (CStr(x3))).Merge
                Range("E" + (CStr(x2)) + ":R" + (CStr(x2))).Merge
                Range("E" + (CStr(x)) + ":R" + (CStr(x))).Merge
                Range("S" + (CStr(x))) = "Contact:"
                Range("S" + (CStr(x)) + ":U" + (CStr(x))).Merge
                Range("T" + (CStr(x4))).MergeCells = False
                'Range("A" + (CStr(x4)) + ":D" + (CStr(x4))).Merge
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
                'Range("A" + (CStr(x19)) + ":G" + (CStr(x19))).Merge 'FIRST SCOPE ITEM



                'Range("E:R").ColumnWidth = 5 'ADJUST COLUMN
                Range("D:D").ColumnWidth = 20 'ADJUST COLUMN
                Range("F:F").ColumnWidth = 50 'ADJUST COLUMN


'                Range("O:O").Delete
'                Range("N:N").Delete
'                Range("K:K").Delete

                Range("H:J").ColumnWidth = 2
                Range("K:L").ColumnWidth = 3
                'Range("M:N").ColumnWidth = 3

                Range("P:Q").ColumnWidth = 5
                Range("R:R").ColumnWidth = 0.1

                Range("N:N").Delete

                Range("M:M").ColumnWidth = 7
                Range("11:11").RowHeight = 15

                Range("N:N").Delete

                Range("A11") = "CATEGORY/SCOPE"
                Range("A" + (CStr(x4)) + ":G" + (CStr(x4))).Merge 'CATEGORY/SCOPE MERGE
                





                

                    
'                Range("P:P").Delete
'                Range("V:V").Delete
'                Range("U:U").Delete
'                Range("T:T").Delete
'                Range("O:O").Delete
'                Range("K:K").Delete
'                Range("G:G").Delete
'                Range("P:P").Delete
'                Range("Q:Q").Delete
                
                
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
'
'                                   Dim rng As Range
                Set rng = Range("A12:P100000") ' Identify your range finds row to put sum formula in
                d = 0
                    For Each cel In rng.Cells
                        If InStr(1, cel.Value, "Grand Total") > 0 Then
                            'MsgBox (c.Address & "found") '<---- Your code goes here
                            d = cel.Row 'Row to put SUM formula in
                            d2 = d - 1 'lower bound of SUM formula
                            d3 = x + 7 'upper bound of SUM formula
                            'Debug.Print d2
                            'Debug.Print d3
                            Range("Q" + (CStr(d))).Value = "=Sum(Q" + (CStr(d3)) + ":Q" + (CStr(d2)) + ")"
                            Range("R" + (CStr(d))).Value = "=Sum(R" + (CStr(d3)) + ":R" + (CStr(d2)) + ")"
                            Range("S" + (CStr(d))).Value = "=Sum(S" + (CStr(d3)) + ":S" + (CStr(d2)) + ")"
                            Range("T" + (CStr(d))).Value = "=Sum(T" + (CStr(d3)) + ":T" + (CStr(d2)) + ")"
                            Range("U" + (CStr(d))).Value = "=Sum(U" + (CStr(d3)) + ":U" + (CStr(d2)) + ")"
                            Range("V" + (CStr(d))).Value = "=Sum(V" + (CStr(d3)) + ":V" + (CStr(d2)) + ")"
                            Range("W" + (CStr(d))).Value = "=Sum(W" + (CStr(d3)) + ":W" + (CStr(d2)) + ")"
                            Range("X" + (CStr(d))).Value = "=Sum(X" + (CStr(d3)) + ":X" + (CStr(d2)) + ")"
                            'Debug.Print d; "test2"
                        End If
                    Next cel
                    
                    
                    
                    Debug.Print B
                    For Each cel In rng.Cells
                        If InStr(1, cel.Value, "Grand Total") > 0 Then
                        d = cel.Row 'Gives bottom of card aka GRAND TOTAL ROW
                        B = x - 2 'Gives Top Left of Card aka JOB: ROW
                        Range("A" + (CStr(B)) + ":X" + (CStr(d))).Borders.LineStyle = xlContinuous ' Gives entire card lines
                         Range("A" + (CStr(B)) + ":D" + (CStr(d))).BorderAround xlContinuous, xlThin 'Gives Border to Job to Card Holder area
                        Range("A" + (CStr(B)) + ":X" + (CStr(d))).BorderAround xlContinuous, xlThick 'Gives Border to Card
                        Range("Q" + (CStr(B)) + ":X" + (CStr(d))).BorderAround xlContinuous, xlThick 'Gives Border to Q-X
                        'Range("A" + (CStr(B)) + ":O" + (CStr(d))).BorderAround xlContinuous, xlThin 'Gives Border to non-editable area from job to card total
                        Range("A" + (CStr(x5)) + ":X" + (CStr(x10))).BorderAround xlContinuous, xlThick 'Gives Border to Addend-Incl Tax cells
                        'Range("A" + (CStr(B)) + ":M" + (CStr(d))).BorderAround xlContinuous, xlThin 'Gives Border to Job to Card Holder area
                        End If
                    Next
                    '    For T = 0 To (NumRows - 1)
                '        Sheets("CARD DUMP").Select
                '        r5 = CStr(arr5(T + 2)) 'Gives A3 on each card
                '        r = CStr(arr3(T + 1)) 'Cell to put total from sum formula
                '        r2 = CStr(arr4(T + 10)) 'Upper bound of range to sum
                '        r3 = CStr((arr4(T + 10) + 1)) 'Upper bound of range sum + 1
                '        r6 = CStr((arr5(T + 2)) + 2) 'Gives A5 on each card
                '        Range("G" + r2 + ":H" + r).BorderAround xlContinuous, xlMedium
                '        Range("G" + r2 + ":I" + r).BorderAround xlContinuous, xlMedium
                '        Range("A" + r3 + ":T" + r3).BorderAround xlContinuous, xlMedium
                '        Range("A" + r5 + ":A" + r6).BorderAround xlContinuous, xlMedium
                '        Range("A" + r5 + ":L" + r).BorderAround xlContinuous, xlThick
                '        Range("A" + r5 + ":T" + r).BorderAround xlContinuous, xlThick
                '        Range("A" + r5 + ":J" + r).BorderAround xlContinuous, xlThick
                '        Range("K" + r + ":T" + r).BorderAround xlContinuous, xlThick
                '
                '        'fixes fonts below
                '        Range("M" + r2 + ":T" + r).NumberFormat = "$#,##0"
                '        Range("A" + ":T").Font.Name = "Calibri"
                '
                '        'adds orange highlight to cards
                '        Range("M" + r2 + ":T" + r2).Interior.ColorIndex = 44
                '
                '        'center all cells where bidders will type
                '        Range("M" + ":T").HorizontalAlignment = xlCenter
                '
                '        'autofit all rows on cards
                '        'Range("A:A").Columns.AutoFit
                '

                    Range("P:P").Delete
                    Range("K:K").ColumnWidth = 3
                Dim rngx As Range
                Set rngx = Range("V12:V100000") ' Identify your range
                d = 0
                    For Each c In rngx.Cells
                        If c.Value = "Page 2 of " Then '<--- Will search if the cell is not empty and not equal to phrase. If you want to check empty cells too remove c.value <> ""
                            'MsgBox (c.Address & "found") '<---- Your code goes here
                            d = c.Row
                            Rows(d).EntireRow.Delete
                            Debug.Print d; "test"
                        End If
                    Next c
                    For Each c In rngx.Cells
                        If c.Value = "Page 3 of " Then '<--- Will search if the cell is not empty and not equal to phrase. If you want to check empty cells too remove c.value <> ""
                            'MsgBox (c.Address & "found") '<---- Your code goes here
                            d = c.Row
                            Rows(d).EntireRow.Delete
                            'Debug.Print d; "test"
                        End If
                    Next c
'                    For Each c In rng.Cells
'                        If c.Value = "Page 4 of " Then '<--- Will search if the cell is not empty and not equal to phrase. If you want to check empty cells too remove c.value <> ""
'                            'MsgBox (c.Address & "found") '<---- Your code goes here
'                            d = c.Row
'                            Rows(d).EntireRow.Delete
'                            Debug.Print d; "test"
'                        End If
'                    Next c
'                    For Each c In rng.Cells
'                        If c.Value = "Page 5 of " Then '<--- Will search if the cell is not empty and not equal to phrase. If you want to check empty cells too remove c.value <> ""
'                            'MsgBox (c.Address & "found") '<---- Your code goes here
'                            d = c.Row
'                            Rows(d).EntireRow.Delete
'                            Debug.Print d; "test"
'                        End If
'                    Next c
'
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

     
            
'        Dim rng1 As Range
'        Set rng1 = Range("A12:A100000") ' Identify your range
'        n = 0
'            For Each k In rng1.Cells
'                If k.Value <> "" And k.Value = "CARD TOTAL MC2:" Then '<--- Will search if the cell is not empty and not equal to phrase. If you want to check empty cells too remove c.value <> ""
'                    'MsgBox (c.Address & "found") '<---- Your code goes here
'                    n = k.Row
'                    n2 = n + 8
'                    n3 = n + 9
'                    'Debug.Print "B" + n2
'                    Range("G" + (CStr(n2))).Value = "Subcontractor in Add/Cut is:"
'                    Range("G" + (CStr(n3))).Value = "Bid Amount in Add/Cut is:"
'                    Range("M" + (CStr(n2))).Value = "(Only Bid Captain fills in, let them know if this does not match bid card.)"
'                    Range("M" + (CStr(n3))).Value = "(Only Bid Captain fills in, let them know if this does not match bid card.)"
'                    Range("G" + (CStr(n2))).Font.Size = "14"
'                    Range("G" + (CStr(n3))).Font.Size = "14"
'                    Range("K" + (CStr(n2)) + ":L" + (CStr(n2))).Merge
'                    Range("K" + (CStr(n3)) + ":L" + (CStr(n3))).Merge
'                    Range("K" + (CStr(n3)) + ":L" + (CStr(n3))).NumberFormat = "$#,##0"
'                    Range("K" + (CStr(n2)) + ":L" + (CStr(n2))).BorderAround xlContinuous, xlThick
'                    Range("K" + (CStr(n3)) + ":L" + (CStr(n3))).BorderAround xlContinuous, xlThick
'
'                End If
'            Next k
        'Range ("K" + (CStr(n3)) + ":L" + (CStr(n3))).value = "=FormatConditions.Add(xlvalue,xlNotEqual,M79:T79)"
'
    Next




        Application.Goto Reference:=Sheets("SHEET CREATOR").Range("A1")
        MsgBox ("Base Bid Cards have now been copied to each sheet.")
End Sub






