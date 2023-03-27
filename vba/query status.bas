Sub query_summary()
'Source data: 360 query management
'Only for TED14856 by Bruce Ma, feel free to contact me if any suggestion
 
    Dim i As Integer
    For i = 1 To 2
    Columns(3).Insert
    Next
    Cells(1, "C") = "Country"
    Cells(1, "D") = "Part"
    Cells(1, "AA") = "Pending"
    Cells(1, "AB") = "Aging"

    For i = 2 To UsedRange.Rows.Count

    'Pending to
    If Cells(i, "L") = "CRA from System" Then
    Cells(i, "AA") = "To CRA"
    End If
    If Cells(i, "S") <> "" Then
            If Cells(i, "L") = "Site from CRA" Or Cells(i, "L") = "Site from System" Then
            Cells(i, "AA") = "To CRA"
            End If
    End If
    If Cells(i, "S") = "" Then
            If Cells(i, "L") = "Site from DM" Or Cells(i, "L") = "Site from CRA" Or Cells(i, "L") = "Site from Coder" Or Cells(i, "L") = "Site from System" Then
            Cells(i, "AA") = "To INV"
            End If
    End If

    'Aging
    If Cells(i, "M").Value <= 15 Then
        Cells(i, "AB") = "<= 15 days"
        ElseIf Cells(i, "M").Value >= 16 And Cells(i, "M").Value <= 28 Then
        Cells(i, "AB") = "16 - 28 days"
        Else: Cells(i, "AB") = "> 28 days"
    End If

    'Part
    Dim a, s
    s = Cells(i, "F").Value
    a = Mid(s, Len(s) - 3, 1)
    part_ID = Array("1", "2", "3", "4")
    part = Array("part A", "part B", "part C", "part D")
    For m = 0 To UBound(part_ID)
        If a = part_ID(m) Then
        Cells(i, "D") = part(m)
        End If
    Next

    'Country
    Dim b, d
    b = Cells(i, "E").Value
    d = Mid(b, 1, 3)
    country_ID = Array("056", "124", "203", "250", "620", "724", "826", "840")
    country = Array("BELGIUM", "CANADA", "CZECH REPUBLIC", "FRANCE", "PORTUGAL", "SPAIN", "UNITED KINGDOM", "UNITED STATES")
    For m = 0 To UBound(country_ID)
        If d = country_ID(m) Then
        Cells(i, "C") = country(m)
        End If
    Next

    Next
End sub