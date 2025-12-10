Attribute VB_Name = "Module1"
Sub SlicerConnection()
Attribute SlicerConnection.VB_ProcData.VB_Invoke_Func = " \n14"

'ActiveSheet.Shapes.Range(Array("Region")).Select(When we recorded the macro,
'Excel captured your action of clicking the slicer, so it wrote this line of code and i deleted then.

If Sheet1.Range("A1").Value = True Then

        ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable1"))
Else
        
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable1"))
End If

If Sheet1.Range("D1").Value = True Then

        ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable5"))
Else
        
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable5"))
End If

If Sheet1.Range("G1").Value = True Then

        ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable6"))
Else
        
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable6"))
End If

If Sheet1.Range("J1").Value = True Then

        ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.AddPivotTable ( _
        ActiveSheet.PivotTables("PivotTable7"))
Else
        
    ActiveWorkbook.SlicerCaches("Slicer_Region").PivotTables.RemovePivotTable ( _
        ActiveSheet.PivotTables("PivotTable7"))
End If
        
End Sub
