Attribute VB_Name = "Aux"
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Declare PtrSafe Function HypCreateConnection Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtProvider As Variant, ByVal vtProviderURL As Variant, ByVal vtServerName As Variant, ByVal vtApplicationName As Variant, ByVal vtDatabaseName As Variant, ByVal vtFriendlyName As Variant, ByVal vtDescription As Variant) As Long
Declare PtrSafe Function HypConnect Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtFriendlyName As Variant) As Long
Declare PtrSafe Function HypMenuVRefresh Lib "HsAddin.dll" () As Long

Sub Fedex_Data_0001()
Attribute Fedex_Data_0001.VB_ProcData.VB_Invoke_Func = " \n14"

'
' Fedex_Data Code by Jda
'

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    Dim appAccess As Object
    Dim vrtCntry As Variant
    Dim arrTables As Variant
    
    Dim X
    Dim y
    
    X = HypCreateConnection("EssBase", "3811756", "FdXjn679", "Analytic Services Provider", "http://prh01612.prod.fedex.com:19000/aps/SmartView", "PRH01617", "FinICE", "FinICE", "FinICE", "Connection Process")
    
    y = HypConnect("EssBase", "3811756", "FdXjn679", "FinICE")
    
'    arrTables = Array("01 - 01 Pseudo_Verify", _
                    "01 - 02 Region_Verify", _
                    "01 - 03 Entity_Verify", _
                    "01 - 04 Revenue _Verify", _
                    "01 - 05 Expenses_Verify", _
                    "01 - 06 FTEs_Verify", _
                    "01 - 07Volume_Verify", _
                    "01 - 08 Weight_Verify")
                    
    arrTables = Array("01 - 01 Pseudo_Verify", _
                    "01 - 02 Region_Verify", _
                    "01 - 03 Entity_Verify", _
                    "01 - 04 Revenue _Verify", _
                    "01 - 05 Expenses_Verify", _
                    "01 - 06 FTEs_Verify", _
                    "01 - 07Volume_Verify", _
                    "01 - 08 Weight_Verify", _
                    "02 - 01 Master_Pull_Revenue", _
                    "02 - 02 Master_Pull_Expenses", _
                    "02 - 03 Master_Pull_FTEs", _
                    "02 - 04 Master_Pull_Volume", _
                    "02 - 05 Master_Pull_Weight")
    
    Workbooks.Open Filename:=CurrentPath & "Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx"
    
    Sheets(arrTables).Select Replace:=False
        
    'Prepare file 0001-0003 to receive data
    'Jda Reduce range to clear
    'Range("A1:AO210000").Select
    'Selection.Clear
    
    Columns("A:AZ").Delete
    
    Sheets(arrTables(0)).Select
    Range("A1").Select
'    Debug.Print "This workbook name = " & ThisWorkbook.Name, " Active Window Range = " & ActiveWindow.RangeSelection.Address

    'ActiveWindow.Close
    'jda Workbooks("Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx").Close
'    MsgBox "Review that File Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx is opened."
    Application.DisplayAlerts = False
    Application.ScreenUpdating = True
    
    'Run for each tab
'    Stop
    For Each vrtCntry In arrTables
        Application.Run "'" & CurrentPath & "Jda 0001-0002-Complete Data File-Expenses.xlsm'!Fedex_Data.Fedex_Data_01", vrtCntry
        
        Range("A1").Select
'        ActiveWorkbook.Save
'        ActiveWorkbook.Close
        
    Next
    
    Workbooks("Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx").Close True
    
End Sub

Public Function CurrentPath() As String
    CurrentPath = ThisWorkbook.Path & "\"
End Function

Public Function PrepIntegers(i As Integer) As String
    If i > 9 Then
        PrepIntegers = CStr(i)
    Else
        PrepIntegers = "0" & CStr(i)
    End If
End Function

Sub CreateSheets()
    Dim arrTables As Variant
    
    arrTables = Array("01 - 01 Pseudo_Verify", _
                    "01 - 02 Region_Verify", _
                    "01 - 03 Entity_Verify", _
                    "01 - 04 Revenue _Verify", _
                    "01 - 05 Expenses_Verify", _
                    "01 - 06 FTEs_Verify", _
                    "01 - 07Volume_Verify", _
                    "01 - 08 Weight_Verify", _
                    "02 - 01 Master_Pull_Revenue", _
                    "02 - 02 Master_Pull_Expenses", _
                    "02 - 03 Master_Pull_FTEs", _
                    "02 - 04 Master_Pull_Volume", _
                    "02 - 05 Master_Pull_Weight")
    For Each vrtCntry In arrTables
        'ThisWorkbook.Sheets.Add
        Workbooks("Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx").Sheets.Add
        ActiveSheet.Name = vrtCntry
    Next

        Workbooks("Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx").Save
End Sub

Public Sub ReadSheetNames(intSheets As Integer)
    Dim i As Integer
    
    For i = 1 To intSheets
       Debug.Print ActiveWorkbook.Sheets(i).Name
    Next
End Sub

Public Sub arraySheetNames() '(intSheets As Integer)
    Dim i As Integer, intSheets As Integer, intProcGroups As Integer, intGSize As Integer
    Dim strSheets As String
    
    intGSize = 5
    intProcGroups = 0
    intSheets = ActiveWorkbook.Sheets.Count
    strSheets = "( "
    Debug.Print "Ran"
    
    For i = 1 To intSheets
        strSheets = strSheets + """" & ActiveWorkbook.Sheets(i).Name & """, "
        If (i Mod intGSize = 0 And intSheets > intGSize) Then
            If intSheets - i < 1 And intSheets Mod intGSize = 0 Then
                Debug.Print Left(strSheets, Len(strSheets) - 2) & " )"
            Else
                Debug.Print strSheets & " _"
                strSheets = ""
                intProcGroups = intProcGroups + 1
                'Debug.Print """" & dbCrrntDB.tabledefs(i).Name & ""","  ' for array 2
                'Debug.Print """" & dbCrrntDB.tabledefs(i).Name & """, _", Len(dbCrrntDB.tabledefs(i).Name)   'for array 1
            End If
        Else
            
            If (intSheets - i < 1 And ((intSheets - intProcGroups * intGSize) Mod intGSize = intSheets Mod intGSize)) Then
                Debug.Print Left(strSheets, Len(strSheets) - 2) & " )"
                strSheets = ""
            End If
        End If
    
    Next

End Sub

Public Function ParentDir(strCrrDir As String) As String
    Dim strLen As Integer
    
    strLen = InStrRev(strCrrDir, "\")
    ParentDir = Left(strCrrDir, strLen)
    
End Function

Public Function ConvertTimeToDecimal(timeIn As Double) As Double
    Dim lngNatural As Double
    
    lngNatural = WorksheetFunction.RoundDown(timeIn, 0)
    ConvertTimeToDecimal = lngNatural + ((timeIn - lngNatural) / 6 * 10)
End Function

Sub replaceInPulls()
    Dim i As Integer
    Dim intSets As Integer
    Dim lngSet As Long
    Dim varToReplace As Variant
    Dim arrReplaceList As Variant
    
    arrReplaceList = Array("#Missing", "#Invalid")
    
    intSets = 1
    lngSet = 20000
    
    Debug.Print "Started " & Now()
    
    For Each varToReplace In arrReplaceList
        For i = 0 To intSets
            Range("F" & i * lngSet + 1 & ":AQ" & (1 + i) * lngSet).Replace varToReplace, 0#
            Debug.Print varToReplace & " - Set " & i & " at " & Now() & " From " & i * lngSet + 1 & " To " & (1 + i) * lngSet
        Next
    Next

End Sub
Sub AddApostrofe()
'
' apostrofe Macro
'

'
    Dim i As Integer
    
    For i = 1 To 50
        ActiveCell.Select
        ActiveCell.FormulaR1C1 = "'" & ActiveCell.Value
        ActiveCell.Offset(0, 1).Select
    Next
    
End Sub

Function TimeDifferenceToNow(oldTimeIn As String) As String
    Dim dblDifference As Double
    Dim hrs As Double, min As Double, sec As Double
    
    dblDifference = CDbl(TimeValue(Now)) - CDbl(TimeValue(oldTimeIn))
    Debug.Print dblDifference; " Diff"
    
    sec = WorksheetFunction.RoundDown(dblDifference * 1440 * 60, 0)
    hrs = (sec \ 3600)
    min = (sec \ 60)
    
    Debug.Print "sec= " & sec
    Debug.Print "min= " & min
    Debug.Print "hrs= " & hrs
    
    min = min - (hrs * 60)
    sec = sec - ((hrs * 60 * 60) + min * 60)
    
    Debug.Print "sec= " & sec
    Debug.Print "min= " & min
    Debug.Print "hrs= " & hrs
    
    TimeDifferenceToNow = hrs & "h " & min & "m " & sec & "s"

End Function
