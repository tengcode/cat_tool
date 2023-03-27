Sub summarize()
    Dim MyPath, MyName, AWbName
    Dim Wb As Workbook, WbN As String
    Dim G As Long, Num As Long, BOX As String
    flag = 0
    
    Application.ScreenUpdating = False
    MyPath = ActiveWorkbook.Path
    MyName = Dir(MyPath & "\" & "*.xls")
    AWbName = ActiveWorkbook.Name
    Num = 0
    
    Do While MyName <> ""
        If MyName <> AWbName Then
            Set Wb = Workbooks.Open(MyPath & "\" & MyName)
            Num = Num + 1
            With Workbooks(1).ActiveSheet
                For G = 1 To Sheets.Count
                    If flag = 0 Then
                        Wb.Sheets(G).UsedRange.Copy .Cells(.Range("A65536").End(xlUp).Row, 1)
                        flag = 1
                    Else
                        Wb.Sheets(G).Range("a1", Wb.Sheets(G).Cells.SpecialCells(xlCellTypeLastCell)).Copy .Cells(.Range("A65536").End(xlUp).Row + 1, 1)

                    End If
                Next
                WbN = WbN & Chr(13) & Wb.Name
                Wb.Close False
            End With
        End If
        MyName = Dir
    Loop
        Range("A1").Select
        
    Application.ScreenUpdating = True
End sub