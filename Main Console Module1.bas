Attribute VB_Name = "Module1"
Option Explicit

Sub RetrieveAccessQuery()
Attribute RetrieveAccessQuery.VB_ProcData.VB_Invoke_Func = " \n14"
'
' RetrieveAccessQuery Macro
'

'
    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:=Array( _
        "OLEDB;Provider=Microsoft.ACE.OLEDB.12.0;Password="""";User ID=Admin;Data Source=C:\Users\3811756\Documents\JuanArango\2019 3rd Level\3rd" _
        , _
        " Level\Pull Review Jda 20190725.accdb;Mode=Share Deny Write;Extended Properties="""";Jet OLEDB:System database="""";Jet OLEDB:Regist" _
        , _
        "ry Path="""";Jet OLEDB:Database Password="""";Jet OLEDB:Engine Type=6;Jet OLEDB:Database Locking Mode=0;Jet OLEDB:Global Partial Bul" _
        , _
        "k Ops=2;Jet OLEDB:Global Bulk Transactions=1;Jet OLEDB:New Database Password="""";Jet OLEDB:Create System Database=False;Jet OLEDB" _
        , _
        ":Encrypt Database=False;Jet OLEDB:Don't Copy Locale on Compact=False;Jet OLEDB:Compact Without Replica Repair=False;Jet OLEDB:SF" _
        , _
        "P=False;Jet OLEDB:Support Complex Data=False;Jet OLEDB:Bypass UserInfo Validation=False;Jet OLEDB:Limited DB Caching=False;Jet O" _
        , "LEDB:Bypass ChoiceField Validation=False"), Destination:=Range("$A$5")). _
        QueryTable
        .CommandType = xlCmdTable
        .CommandText = Array("02 Main Data")
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
        .SourceDataFile = _
        "C:\Users\3811756\Documents\JuanArango\2019 3rd Level\3rd Level\Pull Review Jda 20190725.accdb"
        .ListObject.DisplayName = "Table_Pull_Review_Jda_20190725.accdb"
        .Refresh BackgroundQuery:=False
    End With
End Sub
'=========================
'*** This workbook code.

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    Dim intResponse As Integer
    
    intResponse = MsgBox("Have you copy the LOG to Word?", vbYesNo, ThisWorkbook.Name)
    Debug.Print "Response " & intResponse
'    Stop
    
    If intResponse <> 7 Then
        'Stop
    Else
        MsgBox "Please save log and close again", , ThisWorkbook.Name
        Stop
    End If
End Sub


'=========================