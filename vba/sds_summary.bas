Sub SDS_summary()
    'for all study by Bruce Ma
    '2021-10-11 适用于赛诺菲的sds文件
    Dim i, m, n, o, a, b, c, d As Long, form1(), form2(), folder1()
    'find MatrixMTXCRF
    Sheets.Add: ActiveSheet.Name = "SDS summary"
    a = Sheets("Fields").UsedRange.Rows.Count
    c = Sheets("Forms").UsedRange.Rows.Count
    d = Sheets("Folders").UsedRange.Rows.Count
    
    With Sheets("SDS summary"):
        For i = 1 To a
        'find form OID and field OID
        .Cells(i, "A") = Sheets("Fields").Cells(i, "B")
        .Cells(i, "B") = Sheets("Fields").Cells(i, "O")
        .Cells(i, "C") = Sheets("Fields").Cells(i, "AA")
        .Cells(i, "D") = Sheets("Fields").Cells(i, "F")
        .Cells(i, "E") = Sheets("Fields").Cells(i, "L")
        .Cells(i, "F") = Sheets("Fields").Cells(i, "A")
        .Cells(1, "G") = "Form Name"
        Next
        End With
        
        'find form name
        ReDim form1(1 To c)
        For i = 1 To c
        form1(i) = Sheets("Forms").Range("A" & i)
        Next
        
        For i = 1 To a
            For o = 1 To UBound(form1)
            If form1(o) = Sheets("SDS summary").Cells(i, "F") Then
            Sheets("SDS summary").Cells(i, "G") = Sheets("Forms").Cells(o, "C")
            End If
            Next
        Next
    
        'find the relationship between form and folder
        For n = 1 To Sheets.Count
        If InStr(Sheets(n).Name, "MTXCRF") > 0 Then
        b = Sheets(n).UsedRange.Rows.Count
        
            ReDim form2(1 To b)
            For i = 1 To b
            form2(i) = Sheets(n).Range("A" & i)
            Next
            
                For o = 2 To UBound(form2)
                Sheets("SDS summary").Cells(1, o + 6) = Sheets(n).Cells(1, o)
                    For i = 2 To a
                    If form2(o) = Sheets("SDS summary").Cells(i, "F") Then
                    For m = 1 To d
                    Sheets("SDS summary").Cells(i, m + 7) = Sheets(n).Cells(o, m + 1)
                    Next
                    End If
                Next
            Next
        
        End If
    Next
    
    'find folder name
    Sheets("SDS summary").Rows(1).Insert
    ReDim folder1(1 To d)
        For i = 1 To d
        folder1(i) = Sheets("Folders").Range("A" & i)
        Next
        
        For i = 1 To b - 1
            For o = 1 To UBound(folder1)
            If folder1(o) = Sheets("SDS summary").Cells(2, i + 8) Then
            Sheets("SDS summary").Cells(1, i + 8) = Sheets("Folders").Cells(o, 3)
            End If
            Next
        Next
        
        'improve format
        With Sheets("SDS summary"):
        .Range("A1:E2").Interior.ColorIndex = 22
        .Range("F1:G2").Interior.ColorIndex = 44
        For i = 1 To d
        .Cells(1, i + 7).Interior.ColorIndex = 43
        .Cells(2, i + 7).Interior.ColorIndex = 43
        Next
        .Columns("A:G").ColumnWidth = 15
        .UsedRange.WrapText = True
        .UsedRange.Font.Name = "Calibri"
        End With
End Sub