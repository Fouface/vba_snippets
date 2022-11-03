Attribute VB_Name = "data_analysis_tut"
Sub ProdFlag_v3()
'https://simpleprogrammer.com/vba-data-analysis-automation/
'data analysis procedure

'Run this on a ProductReport to find records without any specs and copy all of the records _
without specs onto a separate worksheet (only copy columns A-G).  Analysis tab added with pivot tables.
Application.ScreenUpdating = False
Dim PRtable As Range, rngX As Range, SpecHeader As Variant
Dim PRTableRows As Long, PRTableColumns As Long
Dim PRsht As Worksheet, nsPRsht As Worksheet
Dim FinalRow As Long, i As Long
Dim IstatActive As Variant, IstatInactive As Variant

Istat = "Item Status"

Set PRsht = Worksheets("Export Worksheet")

PRTableRows = PRsht.Cells.Find("*", searchorder:=xlByRows, SearchDirection:=xlPrevious).Row
PRTableColumns = PRsht.Cells.Find("*", searchorder:=xlByColumns, SearchDirection:=xlPrevious).Column

Set PRtable = PRsht.Range("A1", Cells(PRTableRows, PRTableColumns))

Sheets.Add After:=ActiveSheet
Sheets("Sheet1").Name = "NoSpecs"

Set nsPRsht = Worksheets("NoSpecs")
FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
PRsht.Range("1:1").AutoFilter

    For Each SpecHeader In PRsht.Range("H1:BM1").Cells
        Range(SpecHeader.Offset(1), SpecHeader.Offset(FinalRow)) _
        .AutoFilter Field:=SpecHeader.Column, Criteria1:="="
    Next SpecHeader

PRtable.Resize(PRTableRows, 7).Copy _
nsPRsht.Range("A1")
Application.CutCopyMode = False
PRsht.ShowAllData
nsPRsht.Range("A1").Select

FinalRow = Cells(Rows.Count, 2).End(xlUp).Row
For i = 2 To FinalRow
    If Cells(i, 1) = "Active" Then
        Cells(i, 1).Resize(, 5).Font.ColorIndex = 25
    ElseIf Cells(i, 1) = "Inactive" Then
        Cells(i, 1).Resize(, 5).Font.ColorIndex = 3
    Else
        With Cells(i, 1).Resize(, 5).Font
            .Name = "TimesNewRoman"
            .Bold = True
        End With
    End If
Next i

ActiveSheet.Range("1:1").AutoFilter
Set rngX = ActiveSheet.Range("1:1").Find(Istat, LookAt:=xlPart)
        If Not rngX Is Nothing Then
        End If

IstatActive = Application.CountIf(Range(rngX.Offset(1), rngX.Offset(FinalRow)), "Active")
IstatInactive = Application.CountIf(Range(rngX.Offset(1), rngX.Offset(FinalRow)), "Inactive")

FinalRow = Cells(Rows.Count, 2).End(xlUp).Row

'Create the NoSpecs_CategoryAnalysis tab with pivot tables
Sheets.Add
ActiveSheet.Name = "NoSpecs_Analysis"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "NoSpecs!R1C1:R23405C7", Version:=6).CreatePivotTable TableDestination:= _
        "NoSpecs_Analysis!R3C1", TableName:="PivotTable1", DefaultVersion:=6
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Item Catalog Category" _
        )
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Item"), "Item Count", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Item Status")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Item Catalog Category"). _
        AutoSort xlDescending, "Item Count"
        
        ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Item"), "Percent", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Percent")
        .Calculation = xlPercentOfTotal
        .NumberFormat = "0.00%"
    End With

ActiveWorkbook.Worksheets("NoSpecs_Analysis").PivotTables("PivotTable1"). _
        PivotCache.CreatePivotTable TableDestination:="NoSpecs_Analysis!R13C1", _
        TableName:="PivotTable2", DefaultVersion:=6
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Item Status")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Item"), "Item Count", xlCount
    ActiveSheet.PivotTables("PivotTable2").AddDataField ActiveSheet.PivotTables( _
        "PivotTable2").PivotFields("Item"), "Percent", xlCount
    Range("C13").Select
    With ActiveSheet.PivotTables("PivotTable2").PivotFields("Percent")
        .Calculation = xlPercentOfTotal
        .NumberFormat = "0.00%"
    End With

Sheets.Add
ActiveSheet.Name = "SpecAnalysis"
    ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:= _
        "Export Worksheet!R1C1:R26363C65", Version:=6).CreatePivotTable TableDestination:= _
        "SpecAnalysis!R3C1", TableName:="PivotTable3", DefaultVersion:=6
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Item Catalog Category" _
        )
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Item"), "Item Count", xlCount
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Item Status")
        .Orientation = xlColumnField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable3").PivotFields("Item Catalog Category"). _
        AutoSort xlDescending, "Item Count"
        
        ActiveSheet.PivotTables("PivotTable3").AddDataField ActiveSheet.PivotTables( _
        "PivotTable3").PivotFields("Item"), "Percent", xlCount
    With ActiveSheet.PivotTables("PivotTable3").PivotFields("Percent")
        .Calculation = xlPercentOfTotal
        .NumberFormat = "0.00%"
    End With

Sheets.Add
    ActiveSheet.Name = "Label_Material"
ActiveWorkbook.Worksheets("SpecAnalysis").PivotTables("PivotTable3").PivotCache _
        .CreatePivotTable TableDestination:="Label_Material!R3C1", TableName:="PivotTable1" _
        , DefaultVersion:=6
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Item"), "Item Count", xlCount
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Item"), "Percent", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Item Catalog Category" _
        )
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Label Material")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Item Catalog Category"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Item Catalog Category")
        .PivotItems("Bar Wrapper").Visible = False
        .PivotItems("IFC & Inner Tray").Visible = False
        .PivotItems("Printed Pouches & Packets").Visible = False
        .PivotItems("Shrink Sleeve").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Percent")
        .Calculation = xlPercentOfTotal
        .NumberFormat = "0.00%"
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Label Material").AutoSort _
        xlDescending, "Percent"
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Label Material")
        .PivotItems("(blank)").Visible = False
    End With
    
Sheets.Add
    ActiveSheet.Name = "COATING"
ActiveWorkbook.Worksheets("SpecAnalysis").PivotTables("PivotTable3").PivotCache _
        .CreatePivotTable TableDestination:="COATING!R3C1", TableName:="PivotTable1" _
        , DefaultVersion:=6
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Item"), "Item Count", xlCount
    ActiveSheet.PivotTables("PivotTable1").AddDataField ActiveSheet.PivotTables( _
        "PivotTable1").PivotFields("Item"), "Percent", xlCount
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Item Catalog Category" _
        )
        .Orientation = xlPageField
        .Position = 1
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("COATING")
        .Orientation = xlRowField
        .Position = 1
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("Item Catalog Category"). _
        CurrentPage = "(All)"
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Item Catalog Category")
        .PivotItems("Bar Wrapper").Visible = False
        .PivotItems("IFC & Inner Tray").Visible = False
        .PivotItems("Printed Pouches & Packets").Visible = False
        .PivotItems("Shrink Sleeve").Visible = False
    End With
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("Percent")
        .Calculation = xlPercentOfTotal
        .NumberFormat = "0.00%"
    End With
    ActiveSheet.PivotTables("PivotTable1").PivotFields("COATING").AutoSort _
        xlDescending, "Percent"
    With ActiveSheet.PivotTables("PivotTable1").PivotFields("COATING")
        .PivotItems("(blank)").Visible = False
    End With
    
    MsgBox ("There are " & FinalRow & " records with no specifications" & vbNewLine & _
vbNewLine & "Number of Active records without specs: " & IstatActive & vbNewLine _
& "Number of Inactive records without specs: " & IstatInactive)
    
    Application.ScreenUpdating = True
    
End Sub

