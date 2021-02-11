Attribute VB_Name = "Module1"
Sub CreateTerritoriesRRTable()
Attribute CreateTerritoriesRRTable.VB_ProcData.VB_Invoke_Func = " \n14"
'
' CreateTerritoriesRRTable Macro
'

'
    ActiveWorkbook.SlicerCaches("Slicer_Territory").VisibleSlicerItemsList = Array _
        ( _
        "[Team].[Territory].&[T01]")
    Range("B3:F41").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    Range("B5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Chart").Select
    ActiveWorkbook.SlicerCaches("Slicer_Territory").VisibleSlicerItemsList = Array _
        ( _
        "[Team].[Territory].&[T02]")
    Range("B4:F41").Select
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    Range("B5").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=15
    Range("B44").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Chart").Select
    ActiveWorkbook.SlicerCaches("Slicer_Territory").VisibleSlicerItemsList = Array _
        ( _
        "[Team].[Territory].&[T03]")
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    Range("B44").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=15
    Range("B82").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Chart").Select
    ActiveWorkbook.SlicerCaches("Slicer_Territory").VisibleSlicerItemsList = Array _
        ( _
        "[Team].[Territory].&[T04]")
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    Selection.End(xlDown).Select
    Range("B120").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Chart").Select
    ActiveWorkbook.SlicerCaches("Slicer_Territory").VisibleSlicerItemsList = Array _
        ( _
        "[Team].[Territory].&[T05]")
    Application.CutCopyMode = False
    Selection.Copy
    Sheets("Sheet1").Select
    Selection.End(xlDown).Select
    Range("B158").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Sheets("Chart").Select
    ActiveWorkbook.SlicerCaches("Slicer_Territory").VisibleSlicerItemsList = Array _
        ( _
        "[Team].[Territory].&[T06]")
End Sub
Sub CreateTerritoriesRRTableJda()
'
' CreateTerritoriesRRTable Changed
'

'
    Dim i As Integer
    Debug.Print "Init at " & Now
    
    Sheets("TerritoryTb").Delete

    ActiveWorkbook.SlicerCaches("Slicer_Territory").ClearManualFilter
    
    Range("B3:F41").Select
    Selection.Copy
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "TerritoryTb"
    Range("B5").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Sheets("Chart").Select
    
    For i = 1 To 30
        'ActiveWorkbook.RefreshAll
        
        ActiveWorkbook.SlicerCaches("Slicer_Territory").VisibleSlicerItemsList = Array("[Team].[Territory].&[" & strTerr(i) & "]")
        
        While Not (Worksheets("Chart").Range("b4") = strTerr(i))
            DoEvents
            If Application.Wait(Now + TimeValue("0:00:10")) Then
                Debug.Print "10 secs passed"
            End If
        Wend
        
        Application.CutCopyMode = False
        Range("B4:F41").Select
        Selection.Copy
        Sheets("TerritoryTb").Select
        Selection.End(xlDown).Select
        ActiveCell.Offset(1, 0).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
            :=False, Transpose:=False
        Selection.PasteSpecial Paste:=xlPasteFormats, Operation:=xlNone, _
            SkipBlanks:=False, Transpose:=False
        Sheets("Chart").Select
        Debug.Print "Territory " & i & " at " & Now
        'If i = 3 Then Stop
    Next
    
    Sheets("TerritoryTb").Select
    Sheets("TerritoryTb").Copy
    ChDir "C:\My Documents\C - Project Local RetentionRate\PerTerritory"
    ActiveWorkbook.SaveAs Filename:= _
        "C:\My Documents\C - Project Local RetentionRate\PerTerritory\2021-02 - SFDC Retention" & "_" & Format(Now, "yyyymmdd_hhmm") & ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Windows("2021_02 - SFDC data - account retention project M2.xlsx").Activate
    Sheets("Chart").Select
    
    Debug.Print "End at " & Now
    
End Sub

Function strTerr(i As Integer) As String

    If i < 10 Then
        strTerr = "T0" & i
    Else
        strTerr = "T" & i
    End If
    
End Function
Sub initSteps()
Attribute initSteps.VB_ProcData.VB_Invoke_Func = " \n14"
'
' initSteps Macro
'

'
'    ActiveWorkbook.SlicerCaches("Slicer_Territory").ClearManualFilter
    
    Sheets("TerritoryTb").Select
    Sheets("TerritoryTb").Copy
    ChDir "C:\My Documents\C - Project Local RetentionRate\PerTerritory"
    ActiveWorkbook.SaveAs Filename:= _
        "C:\My Documents\C - Project Local RetentionRate\PerTerritory\2021-02 - SFDC Retention" & "_" & Format(Now, "yyyymmdd_hhmm") & ".xlsx" _
        , FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Windows("2021_02 - SFDC data - account retention project M2.xlsx").Activate
    Sheets("Chart").Select
End Sub

Public Sub refreshPivotTbls()

    Dim ws  As Worksheet
    Dim pt  As PivotTable

    
    For Each ws In ActiveWorkbook.Worksheets
        For Each pt In ws.PivotTables
            pt.RefreshDataSourceValues
            pt.RefreshTable
            pt.RepeatAllLabels
        Next pt
    Next ws

End Sub

Sub waiting10secs()
If Application.Wait(Now + TimeValue("0:00:10")) Then
 MsgBox "Time expired"
End If
End Sub
