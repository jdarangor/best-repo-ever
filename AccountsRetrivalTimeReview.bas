Attribute VB_Name = "Module1"
Option Explicit

Public Sub CY_AccountRetrival()
    Dim wbData As Excel.Workbook
    Dim vrtSetList As Variant, vrtSet As Variant
    Dim intRows As Long
    Dim intColSet As Long
    Dim vrtSetCol As Variant
    Dim x As Variant
    Dim strStartAt As String, strAppStartAt As String
    Dim iCol As Integer
    Dim vrtMonths As Variant, vrtMnth As Variant
    Dim vrtScenar As Variant, vrtScnr As Variant
    Dim vrtGroups As Variant, vrtGrp As Variant
    
    intRows = ThisWorkbook.Sheets("MetaData").Range("I4")
    intColSet = WorksheetFunction.RoundUp(ThisWorkbook.Sheets("MetaData").Range("I5") / 2, 0)
    vrtSetCol = Array(0, 1)
    vrtScenar = Array("Actual Without Integration")
    vrtMonths = ListingMonths()
    ThisWorkbook.Sheets("MetaData").Range("H11").Value = "'FY-2020"
    
    Set wbData = Workbooks.Add '.Open(ThisWorkbook.Path & "\TestingPullsRetrival.xlsx")
    iCol = 0
    
    strAppStartAt = Now()
    Debug.Print "AccountRetrival started at: " & strAppStartAt
    Stop
    
    'Scenario
    For Each vrtScnr In vrtScenar
        ThisWorkbook.Sheets("MetaData").Range("D11") = vrtScnr
        
        'Months
        For Each vrtMnth In vrtMonths
            ThisWorkbook.Sheets("MetaData").Range("I11") = vrtMnth
            
            'Cols set
            For Each vrtSet In vrtSetCol
                Debug.Print vbCr & "Cols: " & intColSet & ", Rows: " & intRows & ". " & intColSet * intRows & " Cells pull."
                
                'Columns
                'select other than Cntry, Psdo, Rgn (CPR)
                ThisWorkbook.Sheets("MetaData").Range("B11:I11").Copy
                'paste
                wbData.Sheets("sheet1").Range("E2:" & Cells(2, 4 + intColSet).Address).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=True
                
                'select CPR
                Dim strCPR As String
                strCPR = "E" & 12 + vrtSet * intColSet & ":" & Cells(11 + intColSet * (vrtSet + 1), 7).Address
                
                ThisWorkbook.Sheets("MetaData").Range(strCPR).Copy
                wbData.Sheets("sheet1").Range("E5").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=True
                
                'Rows
                'select
                'ThisWorkbook.Sheets("Metadata").Range("J12:K12").Copy
                ThisWorkbook.Sheets("Metadata").Range("J12:" & Cells(11 + intRows, 11).Address).Copy
                
                wbData.Sheets("sheet1").Range("C10").PasteSpecial
                
                'Dim x As Long
                'x = HypSetActiveConnection("")
                
                'Select Pull range
                
        '        MsgBox "Please review pull and continue", , ThisWorkbook.Name
                strStartAt = CStr(Now)
                Debug.Print "Pull started at: " & strStartAt
                
                wbData.Activate
                wbData.Sheets("sheet1").Range("C2:" & Cells(9 + intRows, 4 + intColSet).Address).Select
                wbData.Sheets("sheet1").Range("E10:G13").ClearContents
                Debug.Print "Pull selection range: " & Selection.Address
                
                x = HypMenuVRefresh()
                ' cdbl(now()) 'Time conversion to double
                
                Do While Range("E10").Value = ""
                    Application.Wait (Now + TimeValue("0:00:05"))
                    Debug.Print "Pull did not succeded. Returned : " & x & " at " & Now()
                    Stop
                    x = HypMenuVRefresh()
                Loop
                
                Debug.Print "Pull ended at: " & Now
                Debug.Print "Pulling time: " & TimeDifferenceToNow(strStartAt) & ".  " & vrtScnr & vrtMnth & "_" & vrtSet + 1
                
                'select and copy
                Selection.Copy
                ThisWorkbook.Worksheets.Add
                
                Select Case vrtScnr
                    Case "Actual Without Integration"
                        ThisWorkbook.ActiveSheet.Name = "Actl_" & vrtMnth & "_" & vrtSet + 1
                    Case "Budget"
                        ThisWorkbook.ActiveSheet.Name = "Bdgt_" & vrtMnth & "_" & vrtSet + 1
                End Select
                
                ThisWorkbook.ActiveSheet.Range("c2").PasteSpecial
                wbData.Sheets("sheet1").UsedRange.ClearContents
                
                Application.Wait (Now + TimeValue("0:00:25"))
                iCol = iCol + 1
                DoEvents
            Next
            
        Next
        
    Next
    
    Call copyValues
    
    Debug.Print "Time expent during this extraction is: " & TimeDifferenceToNow(strAppStartAt) & vbCr & " for " & ThisWorkbook.Name
    wbData.Close False
    Set wbData = Nothing
    
    Application.SendKeys "~"
    ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\AccountsRetrivalTimeReview-" & Format(Now(), "yyyymmdd-hhmm") & ".xlsx", _
         FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    MsgBox "Please copy log to word.", , ActiveWorkbook.Name
    Stop

End Sub

Function TimeDifferenceToNow(oldTimeIn As String) As String
    Dim dblDifference As Double
    Dim hrs As Double, min As Double, sec As Double
    
    dblDifference = CDbl(TimeValue(Now)) - CDbl(TimeValue(oldTimeIn))
    
    sec = WorksheetFunction.RoundDown(dblDifference * 1440 * 60, 0)
    hrs = (sec \ 3600)
    min = (sec \ 60)
    
    min = min - (hrs * 60)
    sec = sec - ((hrs * 60 * 60) + min * 60)
    
    TimeDifferenceToNow = hrs & "h " & min & "m " & sec & "s"

End Function

Sub copyValues()
'
' copyValues Macro
'

'
    Dim wbSum As Excel.Workbook
    Dim iCol As Integer
    Dim vrtScenar As Variant, vrtMonths As Variant
    Dim vrtScnr As Variant, vrtMnth As Variant
    Dim strAppStartAt As String
    Dim strStartTab As String
    
    Set wbSum = Workbooks.Add
    ActiveSheet.Name = "Accounts-Countries LatAmeri_1"
    wbSum.Worksheets.Add
    ActiveSheet.Name = "Accounts-Countries LatAmeri_2"
    
    'Select data 1
    
    strStartTab = "Actl_" & ListingMonths(0) & "_1"
    ThisWorkbook.Sheets(strStartTab).Activate '.Select
    Selection.Copy
    wbSum.Activate
    Worksheets("Accounts-Countries LatAmeri_1").Select
    Range("C2").PasteSpecial
    'clear data
    Range("E10").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    'Select data 2
    strStartTab = "Actl_" & ListingMonths(0) & "_2"
    ThisWorkbook.Sheets(strStartTab).Activate '.Select
    Selection.Copy
    wbSum.Activate
    Worksheets("Accounts-Countries LatAmeri_2").Select
    Range("C2").PasteSpecial
    'clear data
    Range("E10").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    vrtScenar = Array("Actl")
    vrtMonths = ListingMonths()
    
    iCol = 1
    
    strAppStartAt = Now()
    Debug.Print "App started at: " & strAppStartAt
    
    For iCol = 1 To 2
        'Scenario
        For Each vrtScnr In vrtScenar
            'Months
            For Each vrtMnth In vrtMonths
        
                ThisWorkbook.Activate
                Sheets(vrtScnr & "_" & vrtMnth & "_" & iCol).Select
                Application.CutCopyMode = False
                Range("E10").Select
                Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Copy
                wbSum.Activate
                
                Select Case iCol
                    Case 1
                        Worksheets("Accounts-Countries LatAmeri_1").Activate
                    Case 2
                        Worksheets("Accounts-Countries LatAmeri_2").Activate
                End Select
                Range("E10").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks _
                    :=False, Transpose:=False
                
                
                Application.CutCopyMode = False
                Debug.Print "Added " & vrtScnr & "_" & vrtMnth & "_" & iCol
        
            Next
                
            Debug.Print "Set " & iCol & " Added at " & Now()
            
        Next
        'iCol = iCol + 1
    Next
    
    wbSum.Activate
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\AccountsRetrivalFigures-" & Format(Now(), "yyyymmdd-hhmm") & ".xlsx", _
         FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    ActiveWindow.Close
    
End Sub

Public Function ListingMonths() As Variant
    ListingMonths = Array("Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan")
    'ListingMonths = Array("Jan")
End Function



Public Sub PY_AccountRetrival()
    Dim wbData As Excel.Workbook
    Dim vrtSetList As Variant, vrtSet As Variant
    Dim intRows As Long
    Dim intColSet As Long
    Dim vrtSetCol As Variant
    Dim x As Variant
    Dim strStartAt As String, strAppStartAt As String
    Dim iCol As Integer
    Dim vrtMonths As Variant, vrtMnth As Variant
    Dim vrtScenar As Variant, vrtScnr As Variant
    Dim vrtGroups As Variant, vrtGrp As Variant
    Dim wbNameSplit As Variant
    
    intRows = ThisWorkbook.Sheets("MetaData").Range("I4")
    intColSet = WorksheetFunction.RoundUp(ThisWorkbook.Sheets("MetaData").Range("I5") / 2, 0)
    vrtSetCol = Array(0, 1)
    vrtScenar = Array("Actual Without Integration")
    vrtMonths = ALL_ListingMonths()
    ThisWorkbook.Sheets("MetaData").Range("H11").Value = "'FY-2019"
    
    Set wbData = Workbooks.Add '.Open(ThisWorkbook.Path & "\TestingPullsRetrival.xlsx")
    iCol = 0
    
    strAppStartAt = Now()
    Debug.Print "PY AccountRetrival started at: " & strAppStartAt
    Stop
    
    'Scenario
    For Each vrtScnr In vrtScenar
        ThisWorkbook.Sheets("MetaData").Range("D11") = vrtScnr
        
        'Months
        For Each vrtMnth In vrtMonths
            ThisWorkbook.Sheets("MetaData").Range("I11") = vrtMnth
            
            'Cols set
            For Each vrtSet In vrtSetCol
                Debug.Print vbCr & "Cols: " & intColSet & ", Rows: " & intRows & ". " & intColSet * intRows & " Cells pull."
                
                'Columns
                'select other than Cntry, Psdo, Rgn (CPR)
                ThisWorkbook.Sheets("MetaData").Range("B11:I11").Copy
                'paste
                wbData.Sheets("sheet1").Range("E2:" & Cells(2, 4 + intColSet).Address).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=True
                
                'select CPR
                Dim strCPR As String
                strCPR = "E" & 12 + vrtSet * intColSet & ":" & Cells(11 + intColSet * (vrtSet + 1), 7).Address
                
                ThisWorkbook.Sheets("MetaData").Range(strCPR).Copy
                wbData.Sheets("sheet1").Range("E5").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=True
                
                'Rows
                'select
                'ThisWorkbook.Sheets("Metadata").Range("J12:K12").Copy
                ThisWorkbook.Sheets("Metadata").Range("J12:" & Cells(11 + intRows, 11).Address).Copy
                
                wbData.Sheets("sheet1").Range("C10").PasteSpecial
                
                'Dim x As Long
                'x = HypSetActiveConnection("")
                
                'Select Pull range
        '        MsgBox "Please review and continue", , ThisWorkbook.Name
                strStartAt = CStr(Now)
                Debug.Print "Pull started at: " & strStartAt
                
                wbData.Activate
                wbData.Sheets("sheet1").Range("C2:" & Cells(9 + intRows, 4 + intColSet).Address).Select
                wbData.Sheets("sheet1").Range("E10:G13").ClearContents
                Debug.Print "Pull selection range: " & Selection.Address
                
                x = HypMenuVRefresh()
                ' cdbl(now()) 'Time conversion to double
                Do While Range("E10").Value = ""
                    Application.Wait (Now + TimeValue("0:00:05"))
                    Debug.Print "Pull did not succeded. Returned : " & x & " at " & Now()
                    Stop
                    x = HypMenuVRefresh()
                Loop
                
                Debug.Print "Pull ended at: " & Now
                Debug.Print "Pulling time: " & TimeDifferenceToNow(strStartAt) & ".  " & vrtScnr & vrtMnth & "_" & vrtSet + 1
                
                'select and copy
                Selection.Copy
                ThisWorkbook.Worksheets.Add
                
                Select Case vrtScnr
                    Case "Actual Without Integration"
                        ThisWorkbook.ActiveSheet.Name = "Actl_" & vrtMnth & "_" & vrtSet + 1
                    Case "Budget"
                        ThisWorkbook.ActiveSheet.Name = "Bdgt_" & vrtMnth & "_" & vrtSet + 1
                End Select
                
                ThisWorkbook.ActiveSheet.Range("c2").PasteSpecial
                wbData.Sheets("sheet1").UsedRange.ClearContents
                
                Application.Wait (Now + TimeValue("0:00:25"))
                iCol = iCol + 1
                DoEvents
            Next
            
        Next
        
    Next
    
    Debug.Print "Time expent during this extraction is: " & TimeDifferenceToNow(strAppStartAt) & vbCr & " for " & ThisWorkbook.Name
    wbData.Close False
    
    Call PY_CopyValues
    
    wbNameSplit = Split(ThisWorkbook.Name, ".")
    
    Application.SendKeys "~"
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\" & wbNameSplit(0) & "-PY_" & Format(Now(), "yyyymmdd-hhmm"), xlWorkbookDefault
    
    MsgBox "Please copy log to word.", , ActiveWorkbook.Name
    Stop

    Set wbData = Nothing
    
End Sub


Public Sub BgtCY_AccountRetrival()
    Dim wbData As Excel.Workbook
    Dim vrtSetList As Variant, vrtSet As Variant
    Dim intRows As Long
    Dim intColSet As Long
    Dim vrtSetCol As Variant
    Dim x As Variant
    Dim strStartAt As String, strAppStartAt As String
    Dim iCol As Integer
    Dim vrtMonths As Variant, vrtMnth As Variant
    Dim vrtScenar As Variant, vrtScnr As Variant
    Dim vrtGroups As Variant, vrtGrp As Variant
    Dim wbNameSplit As Variant
    
    intRows = ThisWorkbook.Sheets("MetaData").Range("I4")
    intColSet = WorksheetFunction.RoundUp(ThisWorkbook.Sheets("MetaData").Range("I5") / 2, 0)
    vrtSetCol = Array(0, 1)
    vrtScenar = Array("Budget")
    vrtMonths = ALL_ListingMonths()
    ThisWorkbook.Sheets("MetaData").Range("H11").Value = "'FY-2020"
    
    Set wbData = Workbooks.Add '.Open(ThisWorkbook.Path & "\TestingPullsRetrival.xlsx")
    iCol = 0
    
    strAppStartAt = Now()
    Debug.Print "BgtCY AccountRetrival started at: " & strAppStartAt
    Stop
    
    'Scenario
    For Each vrtScnr In vrtScenar
        ThisWorkbook.Sheets("MetaData").Range("D11") = vrtScnr
        
        'Months
        For Each vrtMnth In vrtMonths
            ThisWorkbook.Sheets("MetaData").Range("I11") = vrtMnth
            
            'Cols set
            For Each vrtSet In vrtSetCol
                Debug.Print vbCr & "Cols: " & intColSet & ", Rows: " & intRows & ". " & intColSet * intRows & " Cells pull."
                
                'Columns
                'select other than Cntry, Psdo, Rgn (CPR)
                ThisWorkbook.Sheets("MetaData").Range("B11:I11").Copy
                'paste
                wbData.Sheets("sheet1").Range("E2:" & Cells(2, 4 + intColSet).Address).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=True
                
                'select CPR
                Dim strCPR As String
                strCPR = "E" & 12 + vrtSet * intColSet & ":" & Cells(11 + intColSet * (vrtSet + 1), 7).Address
                
                ThisWorkbook.Sheets("MetaData").Range(strCPR).Copy
                wbData.Sheets("sheet1").Range("E5").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=True
                
                'Rows
                'select
                'ThisWorkbook.Sheets("Metadata").Range("J12:K12").Copy
                ThisWorkbook.Sheets("Metadata").Range("J12:" & Cells(11 + intRows, 11).Address).Copy
                
                wbData.Sheets("sheet1").Range("C10").PasteSpecial
                
                'Dim x As Long
                'x = HypSetActiveConnection("")
                
                'Select Pull range
        '        MsgBox "Please review and continue", , ThisWorkbook.Name
                strStartAt = CStr(Now)
                Debug.Print "Pull started at: " & strStartAt
                
                wbData.Activate
                wbData.Sheets("sheet1").Range("C2:" & Cells(9 + intRows, 4 + intColSet).Address).Select
                wbData.Sheets("sheet1").Range("E10:G13").ClearContents
                Debug.Print "Pull selection range: " & Selection.Address
                
                x = HypMenuVRefresh()
                ' cdbl(now()) 'Time conversion to double
                Do While Range("E10").Value = ""
                    Application.Wait (Now + TimeValue("0:00:05"))
                    Debug.Print "Pull did not succeded. Returned : " & x & " at " & Now()
                    Stop
                    x = HypMenuVRefresh()
                Loop
                
                Debug.Print "Pull ended at: " & Now
                Debug.Print "Pulling time: " & TimeDifferenceToNow(strStartAt) & ".  " & vrtScnr & vrtMnth & "_" & vrtSet + 1
                
                'select and copy
                Selection.Copy
                ThisWorkbook.Worksheets.Add
                
                Select Case vrtScnr
                    Case "Actual Without Integration"
                        ThisWorkbook.ActiveSheet.Name = "Actl_" & vrtMnth & "_" & vrtSet + 1
                    Case "Budget"
                        ThisWorkbook.ActiveSheet.Name = "Bdgt_" & vrtMnth & "_" & vrtSet + 1
                End Select
                
                ThisWorkbook.ActiveSheet.Range("c2").PasteSpecial
                wbData.Sheets("sheet1").UsedRange.ClearContents
                
                Application.Wait (Now + TimeValue("0:00:25"))
                iCol = iCol + 1
                DoEvents
            Next
            
        Next
        
    Next
    
    Debug.Print "Time expent during this extraction is: " & TimeDifferenceToNow(strAppStartAt) & vbCr & " for " & ThisWorkbook.Name
    wbData.Close False
    
    Call BgtCY_CopyValues
    
    wbNameSplit = Split(ThisWorkbook.Name, ".")
    
    Application.SendKeys "~"
    ThisWorkbook.SaveAs ThisWorkbook.Path & "\" & wbNameSplit(0) & "-BgtCY_" & Format(Now(), "yyyymmdd-hhmm"), xlWorkbookDefault
    
    MsgBox "Please copy log to word.", , ActiveWorkbook.Name
    Stop

    Set wbData = Nothing
    
End Sub



Sub PY_CopyValues()
'
' copyValues Macro
'

'
    Dim wbSum As Excel.Workbook
    Dim iCol As Integer
    Dim vrtScenar As Variant, vrtMonths As Variant
    Dim vrtScnr As Variant, vrtMnth As Variant
    Dim strAppStartAt As String
    Dim strStartTab As String
    
    Set wbSum = Workbooks.Add
    ActiveSheet.Name = "Accounts-Countries LatAmeri_1"
    wbSum.Worksheets.Add
    ActiveSheet.Name = "Accounts-Countries LatAmeri_2"
    
    'Select data 1
    strStartTab = "Actl_" & ListingMonths(0) & "_1"
    ThisWorkbook.Sheets(strStartTab).Activate '.Select
    Selection.Copy
    wbSum.Activate
    Worksheets("Accounts-Countries LatAmeri_1").Select
    Range("C2").PasteSpecial
    'clear data
    Range("E10").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    'Select data 2
    strStartTab = "Actl_" & ListingMonths(0) & "_2"
    ThisWorkbook.Sheets(strStartTab).Activate '.Select
    Selection.Copy
    wbSum.Activate
    Worksheets("Accounts-Countries LatAmeri_2").Select
    Range("C2").PasteSpecial
    'clear data
    Range("E10").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    vrtScenar = Array("Actl")
    vrtMonths = ALL_ListingMonths()
    
    iCol = 1
    
    strAppStartAt = Now()
    Debug.Print vbCr & "Figures started at: " & strAppStartAt
    
    For iCol = 1 To 2
        'Scenario
        For Each vrtScnr In vrtScenar
            'Months
            For Each vrtMnth In vrtMonths
        
                ThisWorkbook.Activate
                Sheets(vrtScnr & "_" & vrtMnth & "_" & iCol).Select
                Application.CutCopyMode = False
                Range("E10").Select
                Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Copy
                wbSum.Activate
                
                Select Case iCol
                    Case 1
                        Worksheets("Accounts-Countries LatAmeri_1").Activate
                    Case 2
                        Worksheets("Accounts-Countries LatAmeri_2").Activate
                End Select
                Range("E10").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks _
                    :=False, Transpose:=False
                
                
                Application.CutCopyMode = False
                Debug.Print "Added " & vrtScnr & "_" & vrtMnth & "_" & iCol
        
            Next
                
        Next
        'iCol = iCol + 1
        Debug.Print "Set " & iCol & " Added at " & Now()
            
    Next
    
    wbSum.Activate
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\AccountsRetrivalFigures-PY_" & Format(Now(), "yyyymmdd-hhmm") & ".xlsx", _
         FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    ActiveWindow.Close
    
End Sub

Sub BgtCY_CopyValues()
'
' copyValues Macro
'

'
    Dim wbSum As Excel.Workbook
    Dim iCol As Integer
    Dim vrtScenar As Variant, vrtMonths As Variant
    Dim vrtScnr As Variant, vrtMnth As Variant
    Dim strAppStartAt As String
    Dim strStartTab As String
    
    Set wbSum = Workbooks.Add
    ActiveSheet.Name = "Accounts-Countries LatAmeri_1"
    wbSum.Worksheets.Add
    ActiveSheet.Name = "Accounts-Countries LatAmeri_2"
    
    'Select data 1
    
    strStartTab = "Bdgt_" & ListingMonths(0) & "_1"
    ThisWorkbook.Sheets(strStartTab).Activate '.Select
    Selection.Copy
    wbSum.Activate
    Worksheets("Accounts-Countries LatAmeri_1").Select
    Range("C2").PasteSpecial
    'clear data
    Range("E10").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    'Select data 2
    strStartTab = "Bdgt_" & ListingMonths(0) & "_2"
    ThisWorkbook.Sheets(strStartTab).Activate '.Select
    Selection.Copy
    wbSum.Activate
    Worksheets("Accounts-Countries LatAmeri_2").Select
    Range("C2").PasteSpecial
    'clear data
    Range("E10").Select
    Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
    Selection.ClearContents
    
    vrtScenar = Array("Bdgt")
    vrtMonths = ALL_ListingMonths()
    
    iCol = 1
    
    strAppStartAt = Now()
    Debug.Print vbCr & "Figures started at: " & strAppStartAt
    
    For iCol = 1 To 2
        'Scenario
        For Each vrtScnr In vrtScenar
            'Months
            For Each vrtMnth In vrtMonths
        
                ThisWorkbook.Activate
                Sheets(vrtScnr & "_" & vrtMnth & "_" & iCol).Select
                Application.CutCopyMode = False
                Range("E10").Select
                Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Copy
                wbSum.Activate
                
                Select Case iCol
                    Case 1
                        Worksheets("Accounts-Countries LatAmeri_1").Activate
                    Case 2
                        Worksheets("Accounts-Countries LatAmeri_2").Activate
                End Select
                Range("E10").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks _
                    :=False, Transpose:=False
                
                
                Application.CutCopyMode = False
                Debug.Print "Added " & vrtScnr & "_" & vrtMnth & "_" & iCol
        
            Next
                
        Next
        'iCol = iCol + 1
        Debug.Print "Set " & iCol & " Added at " & Now()
            
    Next
    
    wbSum.Activate
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\AccountsRetrivalFigures-BgtCY_" & Format(Now(), "yyyymmdd-hhmm") & ".xlsx", _
         FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    ActiveWindow.Close
    
End Sub

Public Function Scenarios() As Variant
    Dim arrScenar As Variant
    Dim s1 As Variant, s2 As Variant, s3 As Variant
        s1 = Array("Actual Without Integration", "'FY-2020")
        s2 = Array("Actual Without Integration", "'FY-2019")
        s3 = Array("Budget", "'FY-2020")
        
      arrScenar = Array(s1, s2, s3)
    
    Scenarios = arrScenar
'    Dim i As Integer, j As Integer
'
'        For i = 0 To 2
'            For j = 0 To 1
'                Debug.Print i + 1; arrScenar(i)(j)
'            Next
'        Next
'        For i = 0 To 2
'                Debug.Print i + 1; arrScenar(i)(0); "  "; arrScenar(i)(1)
'        Next
'
End Function

Public Sub ALL_AccountRetrival()
    Dim wbData As Excel.Workbook
    Dim vrtSetList As Variant, vrtSet As Variant
    Dim intRows As Long
    Dim intColSet As Long
    Dim vrtSetCol As Variant
    Dim x As Variant
    Dim strStartAt As String, strAppStartAt As String
    Dim iCol As Integer, i As Integer, j As Integer
    Dim vrtMonths As Variant, vrtMnth As Variant
    Dim vrtScenar As Variant, vrtScnr As Variant
    Dim vrtGroups As Variant, vrtGrp As Variant
    Dim vrtScenarios As Variant
    
    intRows = ThisWorkbook.Sheets("MetaData").Range("I4")
    intColSet = WorksheetFunction.RoundUp(ThisWorkbook.Sheets("MetaData").Range("I5") / 2, 0)
    vrtSetCol = Array(0, 1)
    vrtScenar = Scenarios() 'Array("Actual Without Integration")
    vrtMonths = ALL_ListingMonths()
    
    Set wbData = Workbooks.Add '.Open(ThisWorkbook.Path & "\TestingPullsRetrival.xlsx")
    iCol = 0
    
    strAppStartAt = Now()
    Debug.Print "AccountRetrival started at: " & strAppStartAt
    Stop
    
    'Scenario
    For Each vrtScnr In vrtScenar
       ThisWorkbook.Sheets("MetaData").Range("D11") = vrtScnr(0)
       ThisWorkbook.Sheets("MetaData").Range("H11").Value = vrtScnr(1) '"'FY-2020"
        
        'Months
        For Each vrtMnth In vrtMonths
            ThisWorkbook.Sheets("MetaData").Range("I11") = vrtMnth
                If ReviewCY(vrtScnr(0), vrtScnr(1), vrtMnth) Then
                    GoTo jumpMonth
                End If
            'Cols set
            For Each vrtSet In vrtSetCol
                Debug.Print vbCr & "Cols: " & intColSet & ", Rows: " & intRows & ". " & intColSet * intRows & " Cells pull."
                
                'Columns
                'select other than Cntry, Psdo, Rgn (CPR)
                ThisWorkbook.Sheets("MetaData").Range("B11:I11").Copy
                'paste
                wbData.Sheets("sheet1").Range("E2:" & Cells(2, 4 + intColSet).Address).PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=True
                
                'select CPR
                Dim strCPR As String
                strCPR = "E" & 12 + vrtSet * intColSet & ":" & Cells(11 + intColSet * (vrtSet + 1), 7).Address
                
                ThisWorkbook.Sheets("MetaData").Range(strCPR).Copy
                wbData.Sheets("sheet1").Range("E5").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
                    False, Transpose:=True
                
                'Rows
                'select
                'ThisWorkbook.Sheets("Metadata").Range("J12:K12").Copy
                ThisWorkbook.Sheets("Metadata").Range("J12:" & Cells(11 + intRows, 11).Address).Copy
                
                wbData.Sheets("sheet1").Range("C10").PasteSpecial
                
                'Dim x As Long
                'x = HypSetActiveConnection("")
                
                'Select Pull range
                
        '        MsgBox "Please review pull and continue", , ThisWorkbook.Name
                strStartAt = CStr(Now)
                Debug.Print "Pull started at: " & strStartAt
                
                wbData.Activate
                wbData.Sheets("sheet1").Range("C2:" & Cells(9 + intRows, 4 + intColSet).Address).Select
                wbData.Sheets("sheet1").Range("E10:G13").ClearContents
                Debug.Print "Pull selection range: " & Selection.Address
                
                x = HypMenuVRefresh()
                'Range("e10") = 1 'debug to count additions
                'GoTo jdaDebug1
                ' cdbl(now()) 'Time conversion to double
                
                Do While Range("E10").Value = ""
                    Application.Wait (Now + TimeValue("0:00:05"))
                    Debug.Print "Pull did not succeded. Returned : " & x & " at " & Now()
                    Stop
                    x = HypMenuVRefresh()
                Loop
jdaDebug1:
                Debug.Print "Pull ended at: " & Now
                Debug.Print "Pulling time: " & TimeDifferenceToNow(strStartAt) & ".  " & vrtScnr(0) & vrtMnth & "_" & vrtSet + 1
                
                'select and copy
                Selection.Copy
                ThisWorkbook.Worksheets.Add
                
                Select Case vrtScnr(0)
                    Case "Actual Without Integration"
                        If vrtScnr(1) = "'FY-2020" Then ThisWorkbook.ActiveSheet.Name = "Actl_" & vrtMnth & "_" & vrtSet + 1
                        If vrtScnr(1) = "'FY-2019" Then ThisWorkbook.ActiveSheet.Name = "prAc_" & vrtMnth & "_" & vrtSet + 1
                    Case "Budget"
                        ThisWorkbook.ActiveSheet.Name = "Bdgt_" & vrtMnth & "_" & vrtSet + 1
                End Select
                
                ThisWorkbook.ActiveSheet.Range("c2").PasteSpecial
                wbData.Sheets("sheet1").UsedRange.ClearContents
                
                Application.Wait (Now + TimeValue("0:00:25"))
                iCol = iCol + 1
                DoEvents
            Next
jumpMonth:
        Next
        
    Next
    
    Call ALL_copyValues
    
    Debug.Print "Time expent during this extraction is: " & TimeDifferenceToNow(strAppStartAt) & vbCr & " for " & ThisWorkbook.Name
    wbData.Close False
    Set wbData = Nothing
    
    Application.SendKeys "~"
    ThisWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\AccountsRetrivalTimeReview-" & Format(Now(), "yyyymmdd-hhmm") & ".xlsx", _
         FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    MsgBox "Please copy log to word.", , ActiveWorkbook.Name
    Stop

End Sub

Sub ALL_copyValues()
'
' copyValues Macro
'

'
    Dim wbSum As Excel.Workbook
    Dim iCol As Integer, i As Integer
    Dim vrtScenar As Variant, vrtMonths As Variant
    Dim vrtScnr As Variant, vrtMnth As Variant
    Dim strAppStartAt As String
    Dim strStartTab As String
    
    Set wbSum = Workbooks.Add
    ActiveSheet.Name = "Accounts-Countries LatAmeri_1"
    wbSum.Worksheets.Add
    ActiveSheet.Name = "Accounts-Countries LatAmeri_2"
    
    'copy headers
    For i = 1 To 2
        'Select data 1
        strStartTab = "Actl_" & ALL_ListingMonths(0) & "_" & i
        ThisWorkbook.Sheets(strStartTab).Activate '.Select
        Selection.Copy
        wbSum.Activate
        Worksheets("Accounts-Countries LatAmeri_" & i).Select
        Range("C2").PasteSpecial
        'clear data
        Range("E10").Select
        Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Select
        Selection.ClearContents
    Next
    
    vrtScenar = Array("Actl", "prAc", "Bdgt")
    vrtMonths = ALL_ListingMonths()
    
    iCol = 1
    
    strAppStartAt = Now()
    Debug.Print "App started at: " & strAppStartAt
    
    For iCol = 1 To 2
        'Scenario
        For Each vrtScnr In vrtScenar
            'Months
            For Each vrtMnth In vrtMonths
        
                ThisWorkbook.Activate
                If ReviewCopyMnth(vrtScnr, vrtMnth) Then
                    GoTo jumpMonthC
                End If
                Sheets(vrtScnr & "_" & vrtMnth & "_" & iCol).Select
                Application.CutCopyMode = False
                Range("E10").Select
                Range(Selection, ActiveCell.SpecialCells(xlLastCell)).Copy
                wbSum.Activate
                
                Select Case iCol
                    Case 1
                        Worksheets("Accounts-Countries LatAmeri_1").Activate
                    Case 2
                        Worksheets("Accounts-Countries LatAmeri_2").Activate
                End Select
                Range("E10").Select
                Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlAdd, SkipBlanks _
                    :=False, Transpose:=False
                
                
                Application.CutCopyMode = False
                Debug.Print "Added " & vrtScnr & "_" & vrtMnth & "_" & iCol
jumpMonthC:
            Next
                
            Debug.Print "Set " & iCol & " Added at " & Now()
            
        Next
        'iCol = iCol + 1
    Next
    
    wbSum.Activate
    ActiveWorkbook.SaveAs Filename:=ThisWorkbook.Path & "\AccountsRetrivalAllFgrs-" & Format(Now(), "yyyymmdd-hhmm") & ".xlsx", _
         FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    ActiveWindow.Close
    
End Sub

Public Function ALL_ListingMonths() As Variant
    ALL_ListingMonths = Array("Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Jan", "Feb", "Mar", "Apr", "May")
    'ALL_ListingMonths = Array("Jan")

End Function

Function ReviewCY(Scenario, year, month) As Boolean
    Dim arrHide As Variant
    Dim m As Variant
    
    arrHide = Array("Feb", "Mar", "Apr", "May")
    
    For Each m In arrHide
        If Scenario = "Actual Without Integration" And year = "'FY-2020" And month = m Then
            ReviewCY = True
            Exit For
        End If
        ReviewCY = False
    Next
End Function

Function ReviewCopyMnth(Scenario, month) As Boolean
    Dim arrHide As Variant
    Dim m As Variant
    'review those month not having tabs for actual year
    arrHide = Array("Feb", "Mar", "Apr", "May")
    
    For Each m In arrHide
        If Scenario = "Actl" And month = m Then
            ReviewCopyMnth = True
            Exit For
        End If
        ReviewCopyMnth = False
    Next
End Function
