Sub Pivot table()
    'Only for TED14856 by Bruce Ma, feel free to contact me if any suggestion
    
    Dim sht As Worksheet
    Dim pvtCache As PivotCache
    Dim pvt As PivotTable
    Dim StartPvt As String
    Dim SrcData As String
        SrcData = ActiveSheet.Name & "!" & UsedRange.Address(ReferenceStyle:=xlR1C1)
        Set sht = Sheets.Add
        StartPvt = sht.Name & "!" & sht.Range("A1").Address(ReferenceStyle:=xlR1C1)
    
    'Create Pivot Cache from Source Data
        Set pvtCache = ActiveWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, _
            SourceData:=SrcData)
        
        'Create Pivot table from Pivot Cache
        Set pvt = pvtCache.CreatePivotTable( _
            TableDestination:=StartPvt, _
            TableName:="PivotTable1")
        
        Set pvt = ActiveSheet.PivotTables("PivotTable1")
            pvt.PivotFields("Country").Orientation = xlRowField
            pvt.PivotFields("Country").Position = 1
            pvt.PivotFields("Site").Orientation = xlRowField
            pvt.PivotFields("Site").Position = 2
            pvt.PivotFields("Aging").Orientation = xlColumnField
            
            pvt.AddDataField pvt.PivotFields("Page"), "Count of Page", xlCount
    
End Sub