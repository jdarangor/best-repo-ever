Attribute VB_Name = "Fedex_Data"
Option Explicit
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
    
    arrTables = BuildArray()
    
    Workbooks.Open Filename:=CurrentPath & "Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx"
'    Stop
    ThisWorkbook.Worksheets("Verification").Range("C6:AN6").Value = Workbooks("Jda Main Console File - Data Information.xlsm").Worksheets("Main Console").Range("g32")
    ThisWorkbook.Worksheets("Extraction").Range("F4:AQ4").Value = Workbooks("Jda Main Console File - Data Information.xlsm").Worksheets("Main Console").Range("g32")
    
    'Prepare file 0001-0003 to receive data
    Select Case Workbooks("Jda Main Console File - Data Information.xlsm").Worksheets("Main Console").Range("g30")
        Case "Verify"
            For Each vrtCntry In arrTables
                Sheets(vrtCntry).Select
                Columns("A:AN").Clear
            Next
        Case "Final Data"
            For Each vrtCntry In arrTables
                Sheets(vrtCntry).Select
                Columns("A:AQ").ClearContents
            Next
    End Select
    
    Sheets(arrTables(0)).Select
    Range("A1").Select
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = True
    
    'Run for each tab
    For Each vrtCntry In arrTables
        If (vrtCntry = "01 - CountryXPseudoXRegionTb") Then
            Call RunAccessCPR_Qry
            GoTo NoRtrival
        End If
        Application.Run "'" & CurrentPath & "Jda 0001-0002-Complete Data File-Expenses.xlsm'!Fedex_Data.Fedex_Data_01", vrtCntry
NoRtrival:

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
    arrTables = BuildArray()
    
    For Each vrtCntry In arrTables
        'ThisWorkbook.Sheets.Add
        Workbooks("Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx").Sheets.Add
        ActiveSheet.Name = Left(vrtCntry, 31)
    Next

        Workbooks("Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx").Save
End Sub

Public Function BuildArray() As Variant
    Dim strSelection As String
    
    strSelection = Workbooks("Jda Main Console File - Data Information.xlsm").Sheets("Main Console").Range("G30")
    
    Select Case strSelection
        Case "Verify"
            BuildArray = Array("01 - CountriesXPseudosTb", "01 - CountriesXRegionsTb", "01 - CountryXPseudoXRegionTb")
        Case "Final Data"
            BuildArray = Array("02 Main DataTb")
            'BuildArray = Array("02 Expns_Colombia", "02 Expns_Venezuela")
    End Select
        
End Function

Public Sub RunAccessCPR_Qry()
    Dim dbAcc As Access.Application
    
    Set dbAcc = New Access.Application
    dbAcc.OpenCurrentDatabase (ThisWorkbook.Path & "\BaseLevel Pulls Jda.accdb")
    
    dbAcc.DoCmd.OpenQuery "01 - CountriesXPseudosXRegions", acViewNormal, acEdit
    
    dbAcc.CloseCurrentDatabase
    Set dbAcc = Nothing
    
End Sub

