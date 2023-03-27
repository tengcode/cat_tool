Sub SDS_summary()
    Dim i, m, a, b As Long
    Sheets.Add: ActiveSheet.Name = "SDS summary"
    a = Sheets("Fields").UsedRange.Rows.Count
    b = Sheets("Matrix21#MTXCRF").UsedRange.Columns.Count
    with Sheets("SDS summary")
        For i = 1 To a
        'find form OID and field OID
            .Cells(i, "A") = Sheets("Fields").Cells(i, "B").Value
            .Cells(i, "B") = Sheets("Fields").Cells(i, "O").Value
            .Cells(i, "C") = Sheets("Fields").Cells(i, "Y").Value
            .Cells(i, "D") = Sheets("Fields").Cells(i, "AA").Value
            .Cells(i, "F") = Sheets("Fields").Cells(i, "A").Value
        Next

        'find form name
        Dim form1, form2 As Range
        For i = 1 To a
            Set form1 = Sheets("Forms").Range("A:A").Find(.Cells(i, "F").Value)
            If TypeName(form1) <> "Nothing" Then
            .Cells(i, "G") = Sheets("Forms").Cells(form1.Row, "C").Value
            End If
        Next
        'find the relationship between form and folder
        For i = 2 To a
            Set form2 = Sheets("Matrix21#MTXCRF").Range("A:A").Find(.Cells(i, "F").Value)
                For m = 8 To b + 6
                .Cells(1, m) = Sheets("Matrix21#MTXCRF").Cells(1, m - 6).Value
                If TypeName(form2) <> "Nothing" Then
                .Cells(i, m) = Sheets("Matrix21#MTXCRF").Cells(form2.Row, m - 6).Value
                End If
            Next
        Next
        .Range("A1:D1").Interior.ColorIndex = 22
        .Range("E1:G1").Interior.ColorIndex = 44
        For i = 8 To b + 6
        .Cells(1, i).Interior.ColorIndex = 43
        Next
        .Range("G1") = "Form Name"
    end with

End Sub