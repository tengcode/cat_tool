Sub summrize()
    'For all study by Bruce Ma, feel free to contact me if any suggestion
    Dim Path, Name, WbName
    Dim Wb As Workbook, newwb As Workbook
    Dim G As Long, m As Long, n As Long, i As Long
                    
    Application.ScreenUpdating = False
    Path = ActiveWorkbook.Path
    Name = Dir(Path & "\" & "*.xls")
    WbName = ActiveWorkbook.Name
    Set newwb = ActiveWorkbook
    With newwb.Sheets(1)
        .Range("A1:K1").Interior.Color = RGB(146, 208, 80)
        .Columns("A").ColumnWidth = 43
        .Columns("B:K").ColumnWidth = 15
        .Columns("B:K").HorizontalAlignment = xlCenter
        .Cells(1, 1) = "Listing"
        .Cells(1, 2) = "Sheet"
        .Cells(1, 3) = "Domain"
        .Cells(1, 4) = "Listing ID"
        .Cells(1, 5) = "Total Record"
        .Cells(1, 6) = "Red"
        .Cells(1, 6).Interior.ColorIndex = 22
        .Cells(1, 7) = "Purple"
        .Cells(1, 7).Interior.ColorIndex = 17
        .Cells(1, 8) = "White"
        .Cells(1, 8).Interior.ColorIndex = 15
        .Cells(1, 9) = "Query"
        .Cells(1, 10) = "Pending next round"
        .Cells(1, 11) = "User1's Comments"
        .UsedRange.WrapText = True
        .UsedRange.Font.Name = "Calibri"
    
    End With
            
    m = 2
    red_cou = 0
    pur_cou = 0
    whi_cou = 0
    que_cou = 0
    pen_cou = 0
    com_cou = 0
        
    Do While Name <> ""
        If Name <> WbName Then
        Set Wb = Workbooks.Open(Path & "\" & Name)
        With newwb.Sheets(1)
        For G = 1 To Sheets.Count
        .Cells(m, 1) = Name
        .Cells(m, 2) = G
        .Cells(m, 3) = Mid(Name, 1, InStr(1, Name, "_") - 1)
        .Cells(m, 4) = Mid(Name, InStr(1, Name, "_") + 1, InStr(6, Name, "_") - InStr(1, Name, "_") - 1)
            For i = 1 To 10
            s = Wb.Sheets(G).Cells(i, 1).Value
            If InStr(s, "in this report") > 0 Then
            Lcount = Mid(s, 1, Len(s) - 20)
            .Cells(m, 5) = Lcount
            Exit For
            End If
            Next
                For n = i + 2 To Lcount + i + 1
                If Wb.Sheets(G).Cells(n, 1).Interior.ColorIndex = 22 Then
                red_cou = red_cou + 1
                ElseIf Wb.Sheets(G).Cells(n, 1).Interior.ColorIndex = 17 Then
                pur_cou = pur_cou + 1
                Else: whi_cou = whi_cou + 1
                End If
                .Cells(m, 6) = red_cou
                .Cells(m, 7) = pur_cou
                .Cells(m, 8) = whi_cou
                Next
    
                For o = 1 To Wb.Sheets(G).UsedRange.Columns.Count
                If Wb.Sheets(G).Cells(i + 1, o) = "Record tag" Then
                    For p = i + 2 To Lcount + i + 1
                    If Wb.Sheets(G).Cells(p, o) = "Query" Then
                    que_cou = que_cou + 1
                    ElseIf Wb.Sheets(G).Cells(p, o) = "Pending next round" Then
                    pen_cou = pen_cou + 1
                    End If
                    .Cells(m, 9) = que_cou
                    .Cells(m, 10) = pen_cou
                    Next
                End If
                
                If Wb.Sheets(G).Cells(i + 1, o) = "The user1's comments" Then
                    For q = i + 2 To Lcount + i + 1
                    If Wb.Sheets(G).Cells(q, o) <> "" Then
                    com_cou = com_cou + 1
                    End If
                    .Cells(m, 11) = com_cou
                    Next
                End If
                Next
                red_cou = 0
                pur_cou = 0
                whi_cou = 0
                que_cou = 0
                pen_cou = 0
                com_cou = 0
        m = m + 1
        
        Next
        Wb.Close False
        End With
        End If
        Name = Dir
    Loop
    
    Application.ScreenUpdating = True
End Sub