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
    vrtComplementaryString = Array(" AWI Final Accounts 5 - 20", " AWI Final Accounts 3 - 10")
    vrtTbsTo00 = Array("Expense", "Revenue")
    vrtTbsTo01 = Array("FTEs", "DAYS", "Volume", "Weight")
    vrtTbsGrp = Array(vrtTbsTo00, vrtTbsTo01)
    i = 0
    
    For Each vrtTbl In vrtTbsArray
        vrtTbsArrayTo = vrtTbsGrp(i)
        
        For Each vrtTblTo In vrtTbsArrayTo
            strQry = "SELECT [Excel Data Extracted Main Data].Pseudo, [Excel Data Extracted Main Data].Entity, [Excel Data Extracted Main Data].Region, [Excel Data Extracted Main Data].Account, [Excel Data Extracted Main Data].Country, [Excel Data Extracted Main Data].[Actual-CY-Jun], [Excel Data Extracted Main Data].[Actual-CY-Jul], [Excel Data Extracted Main Data].[Actual-CY-Aug], [Excel Data Extracted Main Data].[Actual-CY-Sep], [Excel Data Extracted Main Data].[Actual-CY-Oct], [Excel Data Extracted Main Data].[Actual-CY-Nov], [Excel Data Extracted Main Data].[Actual-CY-Dec], [Excel Data Extracted Main Data].[Actual-CY-Jan], [Excel Data Extracted Main Data].[Actual-CY-Feb], [Excel Data Extracted Main Data].[Actual-CY-Mar], [Excel Data Extracted Main Data].[Actual-CY-Apr], [Excel Data Extracted Main Data].[Actual-CY-May], [Excel Data Extracted Main Data].[Actual-PY-Jun], [Excel Data Extracted Main Data].[Actual-PY-Jul], " & _
            "[Excel Data Extracted Main Data].[Actual-PY-Aug], [Excel Data Extracted Main Data].[Actual-PY-Sep], [Excel Data Extracted Main Data].[Actual-PY-Oct], [Excel Data Extracted Main Data].[Actual-PY-Nov], " & _
            "[Excel Data Extracted Main Data].[Actual-PY-Dec], [Excel Data Extracted Main Data].[Actual-PY-Jan], [Excel Data Extracted Main Data].[Actual-PY-Feb], [Excel Data Extracted Main Data].[Actual-PY-Mar], [Excel Data Extracted Main Data].[Actual-PY-Apr], [Excel Data Extracted Main Data].[Actual-PY-May], [Excel Data Extracted Main Data].[Budget-CY-Jun], [Excel Data Extracted Main Data].[Budget-CY-Jul], [Excel Data Extracted Main Data].[Budget-CY-Aug], [Excel Data Extracted Main Data].[Budget-CY-Sep], [Excel Data Extracted Main Data].[Budget-CY-Oct], [Excel Data Extracted Main Data].[Budget-CY-Nov], [Excel Data Extracted Main Data].[Budget-CY-Dec], [Excel Data Extracted Main Data].[Budget-CY-Jan], [Excel Data Extracted Main Data].[Budget-CY-Feb], [Excel Data Extracted Main Data].[Budget-CY-Mar], [Excel Data Extracted Main Data].[Budget-CY-Apr], [Excel Data Extracted Main Data].[Budget-CY-May] " & _
            " , """ & strDateTime & """ as Stamp " & _
            " INTO [" & vrtTblTo & vrtComplementaryString(i) & "]" & _
            " FROM [Excel Data Extracted Main Data] INNER JOIN [" & vrtTbl & "] ON [Excel Data Extracted Main Data].Trim = [" & vrtTbl & "].[Account G Level]" & _
            " WHERE ((([" & vrtTbl & "].Member)=""" & vrtTblTo & """) AND (([Excel Data Extracted Main Data].Verification)<>0));"
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
        "Excel 01 - CountriesXAccounts G_Lvl_Verify", "Excel 01 - CountriesXEntities MD_LvTb_Verify", "Excel Data Extracted Main Data", _
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
    DoCmd.TransferSpreadsheet acLink, 10, "Hrchy Accounts", strPath & strExcelFile, True, "Accounts!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Hrchy Countries", strPath & strExcelFile, True, "Countries!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Hrchy Entities", strPath & strExcelFile, True, "Entities!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Hrchy Pseudos", strPath & strExcelFile, True, "Pseudos!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Hrchy Regions", strPath & strExcelFile, True, "Regions!A1:ao21000"
    
    'Verify and Final Data
    strExcelFile = "Jda 0001-0003-Complete Data File-All Countries-Expenses.xlsx"

    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - Countries-Regions_Verify", strPath & strExcelFile, False, "01 - Countries-RegionsTb!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - Countries-Pseudos_Verify", strPath & strExcelFile, False, "01 - Countries-PseudosTb!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - CountriesXAccounts G_Lvl_Verify", strPath & strExcelFile, False, "01 - CountriesXAccounts G_LvlTb!A1:ao21000"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel 01 - CountriesXEntities MD_LvTb_Verify", strPath & strExcelFile, False, "01 - CountriesXEntities MD_LvTb!A1:ao21000"
    'DoCmd.TransferSpreadsheet acLink, 10, "Excel 02 Complete Data From Main", strPath & strExcelFile, True, "02 Main DataTb!A:au"
    DoCmd.TransferSpreadsheet acLink, 10, "Excel Data Extracted Main Data", strPath & strExcelFile, True, "02 Main DataTb!A:au"

    'QC
    'strExcelFile = "QC Accounts X Countries 20190801.xlsx"
    strExcelFile = InputBox("Enter the date that you want to link to: " & vbCr & "For August 01 2019 enter 20190801." _
                    & vbCr & vbCr & "Verify that exist one file like 'QC Accounts X Countries 20190801.xlsx' ", CurrentDb.Name)
    strExcelFile = "QC Accounts X Countries " & strExcelFile & ".xlsx"

    DoCmd.TransferSpreadsheet acLink, 10, "QC Excel Expense RANGE", strPath & strExcelFile, False, "Expense!A1:ao100"
    DoCmd.TransferSpreadsheet acLink, 10, "QC Excel FTEs RANGE", strPath & strExcelFile, False, "FTEs!A1:ao100"
    DoCmd.TransferSpreadsheet acLink, 10, "QC Excel Revenue RANGE", strPath & strExcelFile, False, "Revenue!A1:ao100"
    DoCmd.TransferSpreadsheet acLink, 10, "QC Excel Volume RANGE", strPath & strExcelFile, False, "Volume!A1:ao100"
    DoCmd.TransferSpreadsheet acLink, 10, "QC Excel Weight RANGE", strPath & strExcelFile, False, "Weight!A1:ao100"
    DoCmd.TransferSpreadsheet acLink, 10, "QC Excel DAYS RANGE", strPath & strExcelFile, False, "DAYS!A1:ao100"

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

    'DoCmd.OpenQuery "01 - Countries X Accounts G_LvlQry", acViewNormal, acEdit
    'DoCmd.OpenQuery "01 - Countries X Entities MD_LvlQry", acViewNormal, acEdit
    DoCmd.OpenQuery "01 - Countries X PseudosQry", acViewNormal, acEdit
    DoCmd.OpenQuery "01 - Countries X RegionsQry", acViewNormal, acEdit
    'DoCmd.OpenQuery "01 - CountriesXPseudosXRegions", acViewNormal, acEdit

testing_OpenQry_Exit:
    Exit Function

testing_OpenQry_Err:
    MsgBox Error$
    Resume testing_OpenQry_Exit

End Function

Public Function Create_MainDataTb() As Boolean
    'Read accounts
    Dim intCol As Integer, intTabs As Integer
    Dim lngAccts As Long, lngEntts As Long
    Dim strSql As String, strPseudo As String, strRegion As String, strCntry As String, strAcnnt As String, strEtty As String
    Dim rsAccnt As ADOR.Recordset, rsEtty As ADOR.Recordset, rsMainData As ADOR.Recordset
    Dim lngRcrdCnt As Long
    Dim intColInAccnt As Integer
    
    Debug.Print "Create Main Data Table Started at: " & Now()
    DoCmd.SetWarnings False
    
    
    
    Set rsMainData = New ADOR.Recordset
    rsMainData.ActiveConnection = CurrentProject.Connection
    'create main data Tb (Extraction)
    strSql = "SELECT [z Structure Of 02 Main DataTb].Pseudo, [z Structure Of 02 Main DataTb].Entity, [z Structure Of 02 Main DataTb].Region, [z Structure Of 02 Main DataTb].Account, [z Structure Of 02 Main DataTb].Country  " & _
                "INTO [02 Main DataTb] " & _
                "FROM [z Structure Of 02 Main DataTb];"
    DoCmd.RunSQL (strSql)
    
    rsMainData.Open "[02 Main DataTb]", CurrentProject.Connection, adOpenDynamic, adLockOptimistic
    
    intColInAccnt = CurrentDb.TableDefs("Verification Entities-Countries LatAmeri_1").Fields.Count
    lngRcrdCnt = 0
    For intTabs = 1 To 2
        Debug.Print vbCr & "Country-Region-Pseudo Group #" & intTabs
        'For intCol = 3 To 86
        For intCol = 3 To intColInAccnt
            strSql = " SELECT [Verification Accounts-Countries LatAmeri_" & intTabs & "].[F2] As Account, ([Verification Accounts-Countries LatAmeri_" & intTabs & "].[F" & intCol & "]) AS AccntValue " & _
                " FROM [Verification Accounts-Countries LatAmeri_" & intTabs & "] " & _
                " WHERE ((([Verification Accounts-Countries LatAmeri_" & intTabs & "].[F" & intCol & "])<>""#Missing"" And ([Verification Accounts-Countries LatAmeri_" & intTabs & "].[F" & intCol & "])<>"""" And ([Verification Accounts-Countries LatAmeri_" & intTabs & "].[F" & intCol & "])<>""#Invalid""))  OR ((([Verification Accounts-Countries LatAmeri_" & intTabs & "].[F" & intCol & "])=""0"")); "
            
            'Debug.Print strSql
            
            Set rsAccnt = New ADOR.Recordset
            rsAccnt.ActiveConnection = CurrentProject.Connection
            rsAccnt.CursorType = adOpenStatic
                    
                    
            rsAccnt.Open strSql
            'GoTo NoRecordsInRecordset
            If rsAccnt.BOF And rsAccnt.EOF Then
                Debug.Print intCol, " No Records"
            Else
                rsAccnt.Move (3)
                strCntry = Trim$(rsAccnt(1).Value)
                rsAccnt.MoveNext
                strPseudo = Trim$(rsAccnt(1).Value)
                rsAccnt.MoveNext
                strRegion = Trim$(rsAccnt(1).Value)
            End If
            'Read entities
                strSql = " SELECT [Verification Entities-Countries LatAmeri_" & intTabs & "].[F2] As Entity , ([Verification Entities-Countries LatAmeri_" & intTabs & "].[F" & intCol & "]) AS EttyValue " & _
                    " FROM [Verification Entities-Countries LatAmeri_" & intTabs & "] " & _
                    " WHERE ((([Verification Entities-Countries LatAmeri_" & intTabs & "].[F" & intCol & "])<>""#Missing"" And ([Verification Entities-Countries LatAmeri_" & intTabs & "].[F" & intCol & "])<>"""" And ([Verification Entities-Countries LatAmeri_" & intTabs & "].[F" & intCol & "])<>""#Invalid""  And ([Verification Entities-Countries LatAmeri_" & intTabs & "].[F" & intCol & "])<>""-""  )) OR ((([Verification Entities-Countries LatAmeri_" & intTabs & "].[F" & intCol & "])=""0"")); "
                
                Set rsEtty = New ADOR.Recordset
                rsEtty.ActiveConnection = CurrentProject.Connection
                rsEtty.CursorType = adOpenStatic
                rsEtty.Open strSql
            
            'Debug.Assert intCol <> 10
            Debug.Print "Column #" & intCol & " Number of accounts: " & rsAccnt.RecordCount, "Number of entities: " & rsEtty.RecordCount
            
            'Create country-pseudo-region combination and add account-entity
            '
            'build record
            'Are there records in rsAccnt?
            If rsAccnt.EOF And rsAccnt.BOF Then
                Debug.Print intCol, " No Records"
                GoTo NoRecordsInRecordset
            End If
            
            rsAccnt.Move (3)
            'rsEtty.Move (8)
            
            Do While Not rsAccnt.EOF
                strAcnnt = Trim$(Nz(rsAccnt(0).Value, "No Value"))
                'Debug.Print "Account " & Trim$(Nz(rsAccnt(0).Value, "No Value"))
                rsAccnt.MoveNext
            
                'GoTo NoRecordsInRecordset
                rsEtty.MoveFirst
                rsEtty.Move (8)
                
                Do While Not rsEtty.EOF
                    strEtty = Trim$(Nz(rsEtty(0).Value, "No Value"))
                    'Debug.Print "Entity " & Trim$(Nz(rsEtty(0).Value, "No Value"))
                    rsEtty.MoveNext
    'Build Tb
                    With rsMainData
                        .AddNew
                        .Fields("Pseudo").Value = strPseudo
                        .Fields("Region").Value = strRegion
                        .Fields("Country").Value = strCntry
                        .Fields("Account").Value = strAcnnt
                        .Fields("Entity").Value = strEtty
                        
                        lngRcrdCnt = lngRcrdCnt + 1
                    End With
                Loop
                
            Loop
            Debug.Print "Records passed: " & lngRcrdCnt
        Next
NoRecordsInRecordset:
        MsgBox "Please review debug log SET " & intTabs & " and copy into LOG", , CurrentDb.Name
        Stop
    Next
        
    
    DoCmd.SetWarnings True
    Debug.Print "Create Main Data Table Ended at: " & Now()
    'MsgBox "Data Pull Table was created.", , Access.CurrentDb.Name
    Create_MainDataTb = True
End Function
