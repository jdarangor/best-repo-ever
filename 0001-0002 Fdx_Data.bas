Attribute VB_Name = "Fedex_Data"
Option Explicit
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Declare PtrSafe Function HypCreateConnection Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtProvider As Variant, ByVal vtProviderURL As Variant, ByVal vtServerName As Variant, ByVal vtApplicationName As Variant, ByVal vtDatabaseName As Variant, ByVal vtFriendlyName As Variant, ByVal vtDescription As Variant) As Long
Declare PtrSafe Function HypConnect Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtFriendlyName As Variant) As Long
Declare PtrSafe Function HypMenuVRefresh Lib "HsAddin.dll" () As Long

Dim blnGraterThan20k As Boolean
Dim maxRowsToPull As Long, FedexLastRow As Long
Dim RowsFromDataDetails As Long, lngInitCell As Long

Sub Fedex_Data_01(ByVal vrtCntry As Variant)
Attribute Fedex_Data_01.VB_ProcData.VB_Invoke_Func = " \n14"

'
' Fedex_Data Macro
' jda review variables in procedures
    Dim y As Variant, x As Variant
    Dim repExtract As Integer
    Dim strDatabase As String
    
    blnGraterThan20k = False

    ThisWorkbook.Activate
    Sheets("Complete Data File").Select
    
    'Add DATA DETAILS Tab again
    Sheets.Add After:=ActiveSheet
    ActiveSheet.Name = "DATA DETAILS"
    ActiveWindow.DisplayGridlines = False
    
    Range("A1").Select
    
    strDatabase = Workbooks("Jda Main Console File - Data Information.xlsm").Sheets("Main Console").Range("G29")
    
'    Stop
'==========================================        'Calls Coutry  =========================================================================================================
        With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
        "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=" & CurrentPath & "" _
        , _
        strDatabase & ".accdb;Mode=Share Deny Wri" _
        , _
        "te;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Registry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engin" _
        , _
        "e Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bulk Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:Ne" _
        , _
        "w Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB:Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Co" _
        , _
        "mpact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SFP=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:By" _
        , _
        "pass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet OLEDB:Bypass ChoiceField Validation=False" _
        ), Destination:=Range("$A$4")).QueryTable
        .CommandType = xlCmdTable
        .CommandText = Array(vrtCntry)
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .SourceDataFile = CurrentPath & strDatabase & ".accdb"
'        .ListObject.DisplayName = "Table_" & strDatabase '& ".accdb"
        .Refresh BackgroundQuery:=False
    End With
    
    '?activesheet.range("a4").end(xldown).row
    With ActiveSheet
        RowsFromDataDetails = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    If RowsFromDataDetails > 20000 Then
        blnGraterThan20k = True
    End If
'    Stop
    'is more than 20k?
    'clean data file
    Sheets("Complete Data File").Select
    Columns("A:AZ").Delete
    
    maxRowsToPull = Workbooks("Jda Main Console File - Data Information.xlsm").Worksheets("Main Console").Range("g31")

    'make selection to pull (20k or current data)
    Sheets("DATA DETAILS").Activate
    If Not blnGraterThan20k Then
        Columns("A:E").Select
    Else
        Range("A1:E" & maxRowsToPull).Select
    End If
    
    Selection.Copy
    Debug.Print "FedEx_Data_01 New copy selection: " & Selection.Address
        
    Sheets("Complete Data File").Select
    Range("A1").Select 'Make this selection as per table set
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    
    Selection.Columns.AutoFit
        
    'Sheets("Complete Data File").Select
    Range("A4:E4").Select
    Selection.ClearContents
    
    If Left(vrtCntry, 2) = "01" Then
        Range("A4:E6").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    Else
        Range("A4:E4").Select
        Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
    End If
       
    If Left(vrtCntry, 2) = "01" Then
        Workbooks("Jda 0001-0001-Complete Data File-Program File.xlsm").Sheets("Verification").Range("C1:AN7").Copy
    Else
        Workbooks("Jda 0001-0001-Complete Data File-Program File.xlsm").Sheets("Extraction").Range("F1:AQ5").Copy
    End If
    
    Sheets("Complete Data File").Select
    
    If Left(vrtCntry, 2) = "01" Then
        Range("C1").Select
    Else
        Range("F1").Select
    End If
    
    ActiveSheet.Paste
    
    With ActiveSheet
        FedexLastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
    End With
    'Formating Pulling range
'    Range("C6:AO" & FedexLastRow).Select
    If Left(vrtCntry, 2) = "01" Then
        Range("C8:AN" & FedexLastRow).Select
    Else
        Range("F6:AQ" & FedexLastRow).Select
    End If
    
    Selection.NumberFormat = "#,##0.00"
         
    Range("A1").Select
    
    Application.DisplayAlerts = False
    If Not blnGraterThan20k Then
        Worksheets("DATA DETAILS").Delete
    End If
    'ActiveWorkbook.Save
    
    If Left(vrtCntry, 2) = "01" Then
        Range("A1:AN" & FedexLastRow).Select
    Else
        Range("A1:AQ" & FedexLastRow).Select
    End If
    
    'Review there is no data retrieved

    If InStr(vrtCntry, "Entities") = 16 Then
        Range("c4:an4").Value = "Account"
    End If
    MsgBox "Make the connection and run the pull." & vbCr & vbCr & "REVIEW ZERO REPLACE", , ThisWorkbook.Name
NoPullData:
'    Stop
    
    If Left(vrtCntry, 2) = "01" Then
        x = HypMenuVRefresh()
        Do While Range("c9").Value = ""
            'MsgBox "Please Refresh pull before continue!", , ThisWorkbook.Name
            'GoTo NoPullData
            Sleep (3000)
            Stop
            Debug.Print "Refresh again at " & Now
            x = HypMenuVRefresh()
        Loop
    Else
        x = HypMenuVRefresh()
        Do While WorksheetFunction.CountIf(Range("f7:g3000"), "#Invalid") > 0
            'MsgBox "Please Refresh pull before continue!", , ThisWorkbook.Name
            'GoTo NoPullData
            Sleep (3000)
            Stop
            Debug.Print "Refresh again at " & Now
            x = HypMenuVRefresh()
        Loop
        Do While Range("F7").Value = ""
            'MsgBox "Please Refresh pull before continue!", , ThisWorkbook.Name
            'GoTo NoPullData
            Sleep (3000)
            Stop
            Debug.Print "Again at " & Now
            x = HypMenuVRefresh()
        Loop
    End If
    
    
'auto        y = HypConnect("Complete Data File", "3811756", "", "FinICE Jda")
    Debug.Print "Connection result " & y
    'Debug.Assert y <> 0
    Sleep (55000)

        Application.StatusBar = "Pull called " & vrtCntry & " - at " & Now()
        Debug.Print "Pull called " & vrtCntry & " - at " & Now()
'    Stop
        
    Windows("Jda 0001-0002-Complete Data File-Expenses.xlsm").Activate
    
    Selection.Copy
    Debug.Print "Cell to paste to: " & Selection.Address
    'copy data to All countries Excel
    Windows("Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx").Activate
    
    Sheets(vrtCntry).Select
    Range("A1").Select
    ActiveSheet.Paste
    Range("A1").Select
    
    'jda ActiveWorkbook.Save
    'jda ActiveWindow.Close
    
    Windows("Jda 0001-0002-Complete Data File-Expenses.xlsm").Activate
    
    Range("A1").Select
    Application.CutCopyMode = False
    
    Calculate
    'Sleep (3000)

    If blnGraterThan20k Then
        Call RunGT20k(vrtCntry)
        ThisWorkbook.Worksheets("DATA DETAILS").Delete
    End If
           
End Sub
Sub RunGT20k(ByVal strCrrntExtractionTab As String)
    MsgBox "RunGT20k was called", , ThisWorkbook.Name
'    Stop
    Dim intRepStract As Integer
    Dim lngInitCell As Long
    Dim maindatalastcell As Long
    'assure init paste is less than 20001 rows
'    Stop
    If FedexLastRow > 20000 Then
        Workbooks("Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx").Sheets(strCrrntExtractionTab).Rows("1:" & (FedexLastRow - 20000)).Delete
    End If

    
    For intRepStract = 1 To (RowsFromDataDetails \ maxRowsToPull)
'       Stop
        lngInitCell = maxRowsToPull * intRepStract
        Workbooks("Jda 0001-0002-Complete Data File-Expenses.xlsm").Sheets("DATA DETAILS").Activate
        Range("A" & lngInitCell + 1 & ":E" & maxRowsToPull * (intRepStract + 1)).Select
                        
        Selection.Copy
       
'       Stop
       Debug.Print "RunGT20k New copy selection: " & Selection.Address
        'continue building pulls from
            'clean data file
        Sheets("Complete Data File").Select
        'Sheets("Complete Data File").Range("A1:AO210000").Clear
        Columns("A:AZ").Delete
        
        'make selection to pull (20k or current data)
        Sheets("DATA DETAILS").Activate
        
        Selection.Copy
            
        Sheets("Complete Data File").Select
'        Range("A1").Select
        If Left(strCrrntExtractionTab, 2) = "01" Then
            Range("A8").Select
        Else
            Range("A6").Select
        End If
        Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        
        Selection.Columns.AutoFit
            
        If Left(strCrrntExtractionTab, 2) = "01" Then
            Workbooks("Jda 0001-0001-Complete Data File-Program File.xlsm").Sheets("Verification").Range("C1:AN7").Copy
        Else
            Workbooks("Jda 0001-0001-Complete Data File-Program File.xlsm").Sheets("Extraction").Range("F1:AQ5").Copy
        End If
        
        Sheets("Complete Data File").Select
        
        If Left(strCrrntExtractionTab, 2) = "01" Then
            Range("C1").Select
        Else
            Range("F1").Select
        End If
        
        ActiveSheet.Paste
        
        With ActiveSheet
            FedexLastRow = .Cells(.Rows.Count, "A").End(xlUp).Row
        End With
        'Formating Pulling range
    '    Range("C6:AO" & FedexLastRow).Select
        If Left(strCrrntExtractionTab, 2) = "01" Then
            Range("C8:AN" & FedexLastRow).Select
        Else
            Range("F6:AQ" & FedexLastRow).Select
        End If
        
        Selection.NumberFormat = "#,##0.00"
             
        Range("A1").Select
        
        Application.DisplayAlerts = False
        If Not blnGraterThan20k Then
            Worksheets("DATA DETAILS").Delete
        End If
        'ActiveWorkbook.Save
        
        'Range("A1:AO" & FedexLastRow).Select
        If Left(strCrrntExtractionTab, 2) = "01" Then
            Range("A1:AN" & FedexLastRow).Select
        Else
            Range("A1:AQ" & FedexLastRow).Select
        End If
        'Review there is no data retrieved
    
        If InStr(strCrrntExtractionTab, "Entities") = 16 Then
            Range("c4:an4").Value = "Account"
        End If
'        MsgBox "Make the connection and run the pull.", , ThisWorkbook.Name
'        Stop
        Dim x As Variant
        If Left(strCrrntExtractionTab, 2) = "01" Then
            x = HypMenuVRefresh()
            Do While Range("c9").Value = ""
                'MsgBox "Please Refresh pull before continue!", , ThisWorkbook.Name
                'GoTo NoPullData
                Sleep (55000)
                Stop
            Loop
        Else
            x = HypMenuVRefresh()
            Do While WorksheetFunction.CountIf(Range("f7:g3000"), "#Invalid") > 0
                'MsgBox "Please Refresh pull before continue!", , ThisWorkbook.Name
                'GoTo NoPullData
                Sleep (55000)
                Stop
                Debug.Print "In loop " & Now()
                x = HypMenuVRefresh()
            Loop
            Do While Range("F7").Value = ""
                'MsgBox "Please Refresh pull before continue!", , ThisWorkbook.Name
                'GoTo NoPullData
                Sleep (55000)
                Stop
                Debug.Print "In loop " & Now()
                x = HypMenuVRefresh()
            Loop
        End If
    
    'auto        y = HypConnect("Complete Data File", "3811756", "", "FinICE Jda")
        Debug.Print "Connection result " & "review 'y' Declaration and update code next line Jda"
        ' Debug.Assert y <> 0
        
    'auto        x = HypMenuVRefresh()
            Sleep (55000)
            Debug.Print "Pull " & intRepStract & " called " & strCrrntExtractionTab & " - at " & Now()
            
        Windows("Jda 0001-0002-Complete Data File-Expenses.xlsm").Activate
        
        'copy data to All countries Excel

        If Left(strCrrntExtractionTab, 2) = "01" Then
            Range("A8:AN" & FedexLastRow).Select
        Else
            Range("A6:AQ" & FedexLastRow).Select
        End If
        
        Selection.Copy
        Windows("Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx").Activate
        
        Sheets(strCrrntExtractionTab).Select
        
        If Left(strCrrntExtractionTab, 2) = "01" Then
            Range("A" & (lngInitCell + 1)).Select
        Else
            Range("A" & (lngInitCell + 1)).Select
        End If
        Debug.Print "Range to paste to: " & Selection.Address
        ActiveSheet.Paste

'        ?range("A1").End(xlDown).Row
        
        Windows("Jda 0001-0002-Complete Data File-Expenses.xlsm").Activate
        
        Range("A1").Select
        Application.CutCopyMode = False
        
    Next
        
        'Dim maindatalastcell As Long
        With ActiveSheet
            maindatalastcell = .Cells(.Rows.Count, "A").End(xlUp).Row
        End With
        
'        Stop
        If Left(strCrrntExtractionTab, 2) = "01" Then
            'Range("A" & (lngInitCell + 1)).Select
            Debug.Print "Review Heading for " & strCrrntExtractionTab & " Tab."
        Else
            Workbooks("Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx").Sheets(strCrrntExtractionTab).Rows("2:4").Delete
            Workbooks("Jda 0001-0001-Complete Data File-Program File.xlsm").Sheets("MainData Header").Range("A1:Au1").Copy
            Workbooks("Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx").Sheets(strCrrntExtractionTab).Range("A1").PasteSpecial
            Workbooks("Jda 0001-0001-Complete Data File-Program File.xlsm").Sheets("MainData Header").Range("ar2:au2").Copy
            
            Workbooks("Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx").Sheets(strCrrntExtractionTab).Activate
            With ActiveSheet
                maindatalastcell = .Cells(.Rows.Count, "A").End(xlUp).Row
            End With
        
            Workbooks("Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx").Sheets(strCrrntExtractionTab).Range("Ar2:AR" & maindatalastcell).PasteSpecial
        
        End If
        
        Calculate


End Sub
Sub Process_Upload_File()

'
' Fedex_Data Macro
'
Workbooks.Open Filename:="C:\Users\3811756\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-03-Complete Data Information Essbase\0001-0003-Complete Data File-All Countries-Expenses.xlsx"

Sheets(Array("Anguilla", "Antigua", "Argentina", "Aruba", "Bahamas", "Barbados", "Belize", "Bermuda", "Bolivia", "Brazil", "BVI", "Cayman Islands", _
              "Chile", "Colombia", "Costa Rica", "Curacao", "Dominica", "Dominican Republic", "Ecuador", "El Salvador", "French Guiana", "Grenada", _
              "Guadeloupe", "Guatemala", "Guyana", "Haiti", "Honduras", "Jamaica", "Martinique", "Mexico", "Montserrat", "Nicaragua", "Panama", _
              "Paraguay", "Peru", "Puerto Rico", "St. Kitts-Nevis", "St. Lucia", "St. Maarten", "Saint Martin - French", "St. Maarten - Dutch", "St. Vincent", "Suriname", _
              "Trinidad And Tobago", "TCI", "U.S. Virgin Islands", "Uruguay", "Venezuela", "United States")).Select Replace:=False
 Range("F5:AO20000").Select
 Selection.NumberFormat = "#,##0.00"

Sheets("Anguilla").Select
Range("A1").Select
ActiveWorkbook.Save
ActiveWindow.Close
          
End Sub

Public Function CurrentPath() As String
    CurrentPath = ThisWorkbook.Path & "\"
End Function

