Attribute VB_Name = "Multiple"
Option Compare Database
Option Explicit

'------------------------------------------------------------
' Create_Master_Pull_Expenses_by_Country
'
'------------------------------------------------------------
Function Create_Master_Pull_ExpensesTb_by_Country()
On Error GoTo Create_Master_Pull_ExpensesTb_by_Country_Err

    Dim vrtCntry As Variant
    Dim arrCntries As Variant
    Dim strQry As String
    
    arrCntries = Array("Colombia", "Venezuela")
    
'    arrCntries = Array("Anguilla", "Antigua", "Argentina", "Aruba", "Bahamas", "Barbados", "Belize", "Bermuda", "Bolivia", "Brazil", "BVI", "Cayman Islands", _
              "Chile", "Colombia", "Costa Rica", "Curacao", "Dominica", "Dominican Republic", "Ecuador", "El Salvador", "French Guiana", "Grenada", _
              "Guadeloupe", "Guatemala", "Guyana", "Haiti", "Honduras", "Jamaica", "Martinique", "Mexico", "Montserrat", "Nicaragua", "Panama", _
              "Paraguay", "Peru", "Puerto Rico", "St. Kitts-Nevis", "St. Lucia", "St. Maarten", "Saint Martin - French", "St. Maarten - Dutch", "St. Vincent", "Suriname", _
              "Trinidad And Tobago", "TCI", "U.S. Virgin Islands", "Uruguay", "Venezuela", "United States")

    For Each vrtCntry In arrCntries
        strQry = "SELECT [02 Main DataTb].Pseudo, [02 Main DataTb].Entity, [02 Main DataTb].Region, [02 Main DataTb].[GL-Acnt], [02 Main DataTb].Country " & _
             "INTO [02 Expns_" & vrtCntry & "] " & _
             "FROM [02 Main DataTb] " & _
             "WHERE ((([02 Main DataTb].Country)=""" & vrtCntry & """));"

        DoCmd.RunSQL (strQry)
    Next

'    strQry = "SELECT [02 - 02 Master_Pull_Expenses].Pesuso, [02 - 02 Master_Pull_Expenses].Entity, [02 - 02 Master_Pull_Expenses].Region, [02 - 02 Master_Pull_Expenses].GL, [02 - 02 Master_Pull_Expenses].Country INTO [02 Expns_US Virgin Islands] FROM [02 - 02 Master_Pull_Expenses] WHERE ((([02 - 02 Master_Pull_Expenses].Country)=""U.S. Virgin Islands""));"
'    strQry = "SELECT [02 - 02 Master_Pull_Expenses].Pesuso, [02 - 02 Master_Pull_Expenses].Entity, [02 - 02 Master_Pull_Expenses].Region, [02 - 02 Master_Pull_Expenses].GL, [02 - 02 Master_Pull_Expenses].Country INTO [02 Expns_St Vincent] FROM [02 - 02 Master_Pull_Expenses] WHERE ((([02 - 02 Master_Pull_Expenses].Country)=""St. Vincent""));"
'    strQry = "SELECT [02 - 02 Master_Pull_Expenses].Pesuso, [02 - 02 Master_Pull_Expenses].Entity, [02 - 02 Master_Pull_Expenses].Region, [02 - 02 Master_Pull_Expenses].GL, [02 - 02 Master_Pull_Expenses].Country INTO [02 Expns_St Maarten - Dutch] FROM [02 - 02 Master_Pull_Expenses] WHERE ((([02 - 02 Master_Pull_Expenses].Country)=""St. Maarten - Dutch""));"
'    strQry = "SELECT [02 - 02 Master_Pull_Expenses].Pesuso, [02 - 02 Master_Pull_Expenses].Entity, [02 - 02 Master_Pull_Expenses].Region, [02 - 02 Master_Pull_Expenses].GL, [02 - 02 Master_Pull_Expenses].Country INTO [02 Expns_St Maarten] FROM [02 - 02 Master_Pull_Expenses] WHERE ((([02 - 02 Master_Pull_Expenses].Country)=""St. Maarten""));"
'    strQry = "SELECT [02 - 02 Master_Pull_Expenses].Pesuso, [02 - 02 Master_Pull_Expenses].Entity, [02 - 02 Master_Pull_Expenses].Region, [02 - 02 Master_Pull_Expenses].GL, [02 - 02 Master_Pull_Expenses].Country INTO [02 Expns_St Lucia] FROM [02 - 02 Master_Pull_Expenses] WHERE ((([02 - 02 Master_Pull_Expenses].Country)=""St. Lucia""));"
'    strQry = "SELECT [02 - 02 Master_Pull_Expenses].Pesuso, [02 - 02 Master_Pull_Expenses].Entity, [02 - 02 Master_Pull_Expenses].Region, [02 - 02 Master_Pull_Expenses].GL, [02 - 02 Master_Pull_Expenses].Country INTO [02 Expns_St Kitts-Nevis] FROM [02 - 02 Master_Pull_Expenses] WHERE ((([02 - 02 Master_Pull_Expenses].Country)=""St. Kitts-Nevis""));"


Create_Master_Pull_Expenses_by_Country_Exit:
    Exit Function

Create_Master_Pull_ExpensesTb_by_Country_Err:
    MsgBox Error$
    Resume Create_Master_Pull_Expenses_by_Country_Exit

End Function

'------------------------------------------------------------
' Create_Tables_By_Query
'
'------------------------------------------------------------
Function Create_Final_Tables_By_Query()
On Error GoTo Create_Tables_By_Query_Err

    Dim vrtTbl As Variant, vrtTbsArray As Variant
    Dim vrtTblTo As Variant, vrtTbsArrayTo As Variant
    Dim vrtTbsTo00  As Variant, vrtTbsTo01 As Variant
    Dim vrtTbsGrp As Variant
    Dim vrtComplementaryString As Variant
    Dim strQry As String
    Dim i As Integer
    Dim strDateTime As String

    strDateTime = DateTime.Date$ & "_" & DateTime.Time$

    
    vrtTbsArray = Array("Accounts G level per Member 5 - L20", "Accounts G Level per Member 3 - L10")
    vrtComplementaryString = Array(" Final Accounts 5 - 20", " Final Accounts 3 - 10")
    vrtTbsTo00 = Array("Expense", "Revenue")
    vrtTbsTo01 = Array("FTEs", "DAYS", "Volume", "Weight")
    vrtTbsGrp = Array(vrtTbsTo00, vrtTbsTo01)
    i = 0
    
    For Each vrtTbl In vrtTbsArray
        vrtTbsArrayTo = vrtTbsGrp(i)
        
        For Each vrtTblTo In vrtTbsArrayTo
            strQry = "SELECT [Excel 02 Complete Data From Main].Pseudo, [Excel 02 Complete Data From Main].Entity, [Excel 02 Complete Data From Main].Region, [Excel 02 Complete Data From Main].Account, [Excel 02 Complete Data From Main].Country, [Excel 02 Complete Data From Main].[Actual-CY-Jun], [Excel 02 Complete Data From Main].[Actual-CY-Jul], [Excel 02 Complete Data From Main].[Actual-CY-Aug], [Excel 02 Complete Data From Main].[Actual-CY-Sep], [Excel 02 Complete Data From Main].[Actual-CY-Oct], [Excel 02 Complete Data From Main].[Actual-CY-Nov], [Excel 02 Complete Data From Main].[Actual-CY-Dec], [Excel 02 Complete Data From Main].[Actual-CY-Jan], [Excel 02 Complete Data From Main].[Actual-CY-Feb], [Excel 02 Complete Data From Main].[Actual-CY-Mar], [Excel 02 Complete Data From Main].[Actual-CY-Apr], [Excel 02 Complete Data From Main].[Actual-CY-May], [Excel 02 Complete Data From Main].[Actual-PY-Jun], [Excel 02 Complete Data From Main].[Actual-PY-Jul], " & _
            "[Excel 02 Complete Data From Main].[Actual-PY-Aug], [Excel 02 Complete Data From Main].[Actual-PY-Sep], [Excel 02 Complete Data From Main].[Actual-PY-Oct], [Excel 02 Complete Data From Main].[Actual-PY-Nov], " & _
            "[Excel 02 Complete Data From Main].[Actual-PY-Dec], [Excel 02 Complete Data From Main].[Actual-PY-Jan], [Excel 02 Complete Data From Main].[Actual-PY-Feb], [Excel 02 Complete Data From Main].[Actual-PY-Mar], [Excel 02 Complete Data From Main].[Actual-PY-Apr], [Excel 02 Complete Data From Main].[Actual-PY-May], [Excel 02 Complete Data From Main].[Budget-CY-Jun], [Excel 02 Complete Data From Main].[Budget-CY-Jul], [Excel 02 Complete Data From Main].[Budget-CY-Aug], [Excel 02 Complete Data From Main].[Budget-CY-Sep], [Excel 02 Complete Data From Main].[Budget-CY-Oct], [Excel 02 Complete Data From Main].[Budget-CY-Nov], [Excel 02 Complete Data From Main].[Budget-CY-Dec], [Excel 02 Complete Data From Main].[Budget-CY-Jan], [Excel 02 Complete Data From Main].[Budget-CY-Feb], [Excel 02 Complete Data From Main].[Budget-CY-Mar], [Excel 02 Complete Data From Main].[Budget-CY-Apr], [Excel 02 Complete Data From Main].[Budget-CY-May] " & _
            " , """ & strDateTime & """ as Stamp " & _
            " INTO [" & vrtTblTo & vrtComplementaryString(i) & "]" & _
            " FROM [Excel 02 Complete Data From Main] INNER JOIN [" & vrtTbl & "] ON [Excel 02 Complete Data From Main].Trim = [" & vrtTbl & "].[Account G Level]" & _
            " WHERE ((([" & vrtTbl & "].Member)=""" & vrtTblTo & """) AND (([Excel 02 Complete Data From Main].Verification)<>0));"
            'Debug.Print strQry
            DoCmd.RunSQL (strQry)
        Next
        i = i + 1
    Next

Create_Tables_By_Query_Exit:
    Exit Function

Create_Tables_By_Query_Err:
    MsgBox Error$
    Resume Create_Tables_By_Query_Exit

End Function

Sub ListTables()
    Dim i As Integer, ii As Integer
    Dim dbCrrntDB As Object
    Dim strTbls As String
    
    Set dbCrrntDB = Application.CurrentDb
    
    For i = 0 To dbCrrntDB.TableDefs.Count - 1
    
'        Debug.Print dbCrrntDB.TableDefs(i).Name             'simple list
        'No MSys tables
        If Left(dbCrrntDB.TableDefs(i).Name, 4) = "MSys" Then GoTo MsysTable
        
        strTbls = strTbls + """" & dbCrrntDB.TableDefs(i).Name & """, "           'trios
        If (i / 3 - i \ 3 = 0) Then                                               'trios
            Debug.Print strTbls & " _"                                            'trios
            strTbls = ""                                                          'trios
            ' Debug.Print """" & dbCrrntDB.tabledefs(i).Name & ""","  ' for array 2
            'Debug.Print """" & dbCrrntDB.tabledefs(i).Name & """, _", Len(dbCrrntDB.tabledefs(i).Name)   'for array 1
        End If                                                                    'trios

MsysTable:
    
    Next

End Sub



'------------------------------------------------------------
' Proc_001_0002_Update_Database_Links
'
'------------------------------------------------------------
Function Proc_001_0002_Update_Database_Links()
On Error GoTo Proc_001_0002_Update_Database_Links_Err

    Dim strPath As String
    Dim strExcelFile As String
    
    strPath = Application.CodeProject.Path & "\"
    strExcelFile = "Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx"

    DoCmd.SetWarnings False
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - 01 Pseudo_Verify", strPath & strExcelFile, False, "01 - 01 Pseudo_Verify!A1:ao10000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - 02 Region_Verify", strPath & strExcelFile, False, "01 - 02 Region_Verify!A1:ao10000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - 03 Entity_Verify", strPath & strExcelFile, False, "01 - 03 Entity_Verify!A1:ao10000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - 04 Revenue _Verify", strPath & strExcelFile, False, "01 - 04 Revenue _Verify!A1:ao10000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - 05 Expenses_Verify", strPath & strExcelFile, False, "01 - 05 Expenses_Verify!A1:ao10000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - 06 FTEs_Verify", strPath & strExcelFile, False, "01 - 06 FTEs_Verify!A1:ao10000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - 07Volume_Verify", strPath & strExcelFile, False, "01 - 07Volume_Verify!A1:ao10000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - 08 Weight_Verify", strPath & strExcelFile, False, "01 - 08 Weight_Verify!A1:ao10000"


Proc_001_0002_Update_Database_Links_Exit:
    Exit Function

Proc_001_0002_Update_Database_Links_Err:
    MsgBox Error$
    Resume Proc_001_0002_Update_Database_Links_Exit

End Function

Function Excel_Vefify_Links()
On Error GoTo Proc_001_0002_Update_Database_Links_Err

    Dim strPath As String
    Dim strExcelFile As String
    
    strPath = Application.CodeProject.Path & "\"
    strExcelFile = "Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx"

    'DoCmd.SetWarnings False
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - Countries-Regions_Verify", strPath & strExcelFile, False, "01 - Countries-RegionsTb!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - Countries-Pseudos_Verify", strPath & strExcelFile, False, "01 - Countries-PseudosTb!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - CountriesXAccounts G_Lvl_Verify", strPath & strExcelFile, False, "01 - CountriesXAccounts G_LvlTb!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - CountriesXEntities MD_LvTb_Verify", strPath & strExcelFile, False, "01 - CountriesXEntities MD_LvTb!A1:ao21000"


Proc_001_0002_Update_Database_Links_Exit:
    Exit Function

Proc_001_0002_Update_Database_Links_Err:
    MsgBox Error$
    Resume Proc_001_0002_Update_Database_Links_Exit

End Function

Function Excel_Hierarchy_Links()
On Error GoTo Proc_001_0002_Update_Database_Links_Err

    Dim strPath As String
    Dim strExcelFile As String
    
    strPath = Application.CodeProject.Path & "\"
    strExcelFile = "Hierarchies Final Product As 20190730.xlsx"

    'DoCmd.SetWarnings False
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - Countries-Regions_Verify", strPath & strExcelFile, False, "01 - Countries-RegionsTb!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - Countries-Pseudos_Verify", strPath & strExcelFile, False, "01 - Countries-PseudosTb!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - CountriesXAccounts G_Lvl_Verify", strPath & strExcelFile, False, "01 - CountriesXAccounts G_LvlTb!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - CountriesXEntities MD_LvTb_Verify", strPath & strExcelFile, False, "01 - CountriesXEntities MD_LvTb!A1:ao21000"


Proc_001_0002_Update_Database_Links_Exit:
    Exit Function

Proc_001_0002_Update_Database_Links_Err:
    MsgBox Error$
    Resume Proc_001_0002_Update_Database_Links_Exit

End Function

Function Excel_Refresh_Links()
On Error GoTo Excel_Refresh_Links_Err

    Dim strPath As String
    Dim strExcelFile As String
    Dim vrtExcelLinkedTbls As Variant, vrtLinkedTbl As Variant
    
    On Error Resume Next 'Jda Item not found - Did not go through
    vrtExcelLinkedTbls = Array("00 Accounts", "00 Pseudos", _
        "00 Countries", "00 Entities", "00 Regions", _
        "Excel 01 - Countries-Pseudos_Verify", "Excel 01 - Countries-Regions_Verify", _
        "Excel 01 - CountriesXAccounts G_Lvl_Verify", "Excel 01 - CountriesXEntities MD_LvTb_Verify", "Excel 02 Complete Data From Main", _
        "QC Excel Expense", "QC Excel FTEs", "QC Excel Revenue", _
        "QC Excel Volume", "QC Excel Weight")

    For Each vrtLinkedTbl In vrtExcelLinkedTbls
        'Access.CurrentDb.TableDefs.Delete ("00 Pseudos")
        Access.CurrentDb.TableDefs.Delete (vrtLinkedTbl)
    Next
    
    
    Err = 0
    
    strPath = Application.CodeProject.Path & "\"
    
    'Hierarchies
    'strExcelFile = "Hierarchies Final Product As 20190801.xlsx"
    strExcelFile = InputBox("Enter the date that you want to link to: " & vbCr & "For August 01 2019 enter 20190801." _
                    & vbCr & vbCr & "Verify that exist one file like 'Hierarchies Final Product As 20190801.xlsx' ", CurrentDb.Name)
    strExcelFile = "Hierarchies Final Product As " & strExcelFile & ".xlsx"
    'DoCmd.SetWarnings False
    DoCmd.TransferSpreadsheet acLink, 10, "00 Accounts", strPath & strExcelFile, True, "Accounts!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "00 Countries", strPath & strExcelFile, True, "Countries!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "00 Entities", strPath & strExcelFile, True, "Entities!A1:ao40000"
    DoCmd.TransferSpreadsheet acLink, 10, "00 Pseudos", strPath & strExcelFile, True, "Pseudos!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "00 Regions", strPath & strExcelFile, True, "Regions!A1:ao21000"
    
    'Verify and Final Data
    strExcelFile = "Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx"

    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - Countries-Regions_Verify", strPath & strExcelFile, False, "01 - Countries-RegionsTb!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - Countries-Pseudos_Verify", strPath & strExcelFile, False, "01 - Countries-PseudosTb!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - CountriesXAccounts G_Lvl_Verify", strPath & strExcelFile, False, "01 - CountriesXAccounts G_LvlTb!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - CountriesXEntities MD_LvTb_Verify", strPath & strExcelFile, False, "01 - CountriesXEntities MD_LvTb!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 02 Complete Data From Main", strPath & strExcelFile, True, "02 Main DataTb!A:at"

    'QC
    'strExcelFile = "QC Accounts X Countries 20190801.xlsx"
    strExcelFile = InputBox("Enter the date that you want to link to: " & vbCr & "For August 01 2019 enter 20190801." _
                    & vbCr & vbCr & "Verify that exist one file like 'QC Accounts X Countries 20190801.xlsx' ", CurrentDb.Name)
    strExcelFile = "QC Accounts X Countries " & strExcelFile & ".xlsx"

    DoCmd.TransferSpreadsheet acLink, 10, "QC Excel Expense", strPath & strExcelFile, False, "Expense!A1:ao100"
    DoCmd.TransferSpreadsheet acLink, 10, "QC Excel FTEs", strPath & strExcelFile, False, "FTEs!A1:ao100"
    DoCmd.TransferSpreadsheet acLink, 10, "QC Excel Revenue", strPath & strExcelFile, False, "Revenue!A1:ao100"
    DoCmd.TransferSpreadsheet acLink, 10, "QC Excel Volume", strPath & strExcelFile, False, "Volume!A1:ao100"
    DoCmd.TransferSpreadsheet acLink, 10, "QC Excel Weight", strPath & strExcelFile, False, "Weight!A1:ao100"
    DoCmd.TransferSpreadsheet acLink, 10, "QC Excel DAYS", strPath & strExcelFile, False, "DAYS!A1:ao100"

    Access.CurrentDb.TableDefs.Refresh

Excel_Refresh_Links_Exit:
    Exit Function

Excel_Refresh_Links_Err:
    MsgBox Error$
    Resume Excel_Refresh_Links_Exit

End Function

'------------------------------------------------------------
' testing_OpenQry
'
'------------------------------------------------------------
Function Run_VerifyTblsCreation()
On Error GoTo testing_OpenQry_Err

    DoCmd.OpenQuery "01 - Countries X Accounts G_LvlQry", acViewNormal, acEdit
    DoCmd.OpenQuery "01 - Countries X Entities MD_LvlQry", acViewNormal, acEdit
    DoCmd.OpenQuery "01 - Countries X PseudosQry", acViewNormal, acEdit
    DoCmd.OpenQuery "01 - Countries X RegionsQry", acViewNormal, acEdit

testing_OpenQry_Exit:
    Exit Function

testing_OpenQry_Err:
    MsgBox Error$
    Resume testing_OpenQry_Exit

End Function



Public Sub QC_Queries(CrrntMnth As String)
    Dim strQry As String, strQry2 As String
    
    strQry = "SELECT [Accounts 3 - L10 All CountriesQry].Member AS Account, [Accounts 3 - L10 All CountriesQry].Country AS Country, IIf(IsNumeric([F6]),[F6],0) AS [Direct From Essbase CY " & CrrntMnth & "], IIf(IsNumeric([Actual-CY-" & CrrntMnth & "]),[Actual-CY-" & CrrntMnth & "],0) AS [Through Access CY " & CrrntMnth & "], Round([Direct From Essbase CY " & CrrntMnth & "]-[Through Access CY " & CrrntMnth & "],2) AS [Difference Essbase-Access CY " & CrrntMnth & "], [Difference Essbase-Access CY " & CrrntMnth & "]/[Direct From Essbase CY " & CrrntMnth & "] AS " & CrrntMnth & "_Diff_Percentage" & _
             " FROM [Accounts 3 - L10 All CountriesQry] INNER JOIN [QC Excel FTEs-Weight-VolumeQry] ON ([Accounts 3 - L10 All CountriesQry].Country = [QC Excel FTEs-Weight-VolumeQry].F3) AND ([Accounts 3 - L10 All CountriesQry].[Member] = [QC Excel FTEs-Weight-VolumeQry].F2);"

    'DoCmd.RunSQL (strQry)
    'AcCommand.acCmdQueryTypeSQLDataDefinition strQry
    Debug.Print strQry
            
    strQry2 = "SELECT [Accounts 5 - L20 All CountriesQry].Country, [Accounts 5 - L20 All CountriesQry].Member, Round(CDbl([F6]),2) AS " & CrrntMnth & "Essbase, Round(CDbl([Actual-CY-" & CrrntMnth & "]),2) AS " & CrrntMnth & "Access, Round(CDbl([QC Excel Expense-RevenueQry]![F6]-[Accounts 5 - L20 All CountriesQry]![Actual-CY-" & CrrntMnth & "]),2) AS " & CrrntMnth & "Difference, [" & CrrntMnth & "Difference]/[" & CrrntMnth & "Essbase] AS " & CrrntMnth & "Diff_Percentage " & _
              " FROM [QC Excel Expense-RevenueQry] INNER JOIN [Accounts 5 - L20 All CountriesQry] ON ([QC Excel Expense-RevenueQry].F3 = [Accounts 5 - L20 All CountriesQry].Country) AND ([QC Excel Expense-RevenueQry].F2 = [Accounts 5 - L20 All CountriesQry].[Member]);"
    Debug.Print strQry2

End Sub
'=====================
' Form commands

Option Compare Database
Option Explicit

Private Sub Command11_Click()
    Call OpenExcelConsole_Click
End Sub

Private Sub OpenExcelConsole_Click()
    Dim wb As Excel.Workbook
    Dim xl As Excel.Application
    
    Set xl = New Excel.Application
    'Set wb = xl.Workbooks(CurrentProject.Path & "\Jda Main Console File - Data Information.xlsm")
'    Excel.Workbooks (CurrentProject.Path & "\Jda Main Console File - Data Information.xlsm")
'    MsgBox "Open " & CurrentProject.Path & "\Jda Main Console File - Data Information.xlsm"
    xl.Workbooks.Open (CurrentProject.Path & "\Jda Main Console File - Data Information.xlsm")
    xl.Visible = True
    
    Set xl = Nothing
'    Application.CompactRepair CurrentDb.Name, CurrentDb.Name
    Application.CloseCurrentDatabase

End Sub