Attribute VB_Name = "FedexData"
Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As LongPtr)
Declare PtrSafe Function HypCreateConnection Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtProvider As Variant, ByVal vtProviderURL As Variant, ByVal vtServerName As Variant, ByVal vtApplicationName As Variant, ByVal vtDatabaseName As Variant, ByVal vtFriendlyName As Variant, ByVal vtDescription As Variant) As Long
Declare PtrSafe Function HypConnect Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtFriendlyName As Variant) As Long
Declare PtrSafe Function HypMenuVRefresh Lib "HsAddin.dll" () As Long

Public Sub FEDEXDATA01_ACCOUNTS()

Application.DisplayAlerts = False
Application.ScreenUpdating = True

ChDir "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses"
Workbooks.Open Filename:="C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0004-0002-Accounts Verification File-All Countries-Expenses.xlsx"

Sheets(Array("Anguilla", "Antigua", "Argentina", "Aruba", "Bahamas", "Barbados", "Belize", "Bermuda", "Bolivia", "Brazil", "BVI", "Cayman Islands", _
              "Chile", "Colombia", "Costa Rica", "Curacao", "Dominica", "Dominican Republic", "Ecuador", "El Salvador", "French Guiana", "Grenada", _
              "Guadeloupe", "Guatemala", "Guyana", "Haiti", "Honduras", "Jamaica", "Martinique", "Mexico", "Montserrat", "Nicaragua", "Panama", _
              "Paraguay", "Peru", "Puerto Rico", "St. Kitts-Nevis", "St. Lucia", "St. Maarten", "Saint Martin - French", "St. Maarten - Dutch", "St. Vincent", "Suriname", _
              "Trinidad And Tobago", "TCI", "U.S. Virgin Islands", "Uruguay", "Venezuela", "United States")).Select Replace:=False
        
    Cells.Select
    Selection.ClearContents
    
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    Selection.Borders(xlEdgeLeft).LineStyle = xlNone
    Selection.Borders(xlEdgeTop).LineStyle = xlNone
    Selection.Borders(xlEdgeBottom).LineStyle = xlNone
    Selection.Borders(xlEdgeRight).LineStyle = xlNone
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
Sheets("Anguilla").Select
Range("A1").Select

ActiveWorkbook.Save
ActiveWindow.Close

Dim X
Dim y

X = HypCreateConnection("EssBase", "3720187", "Poweraa1", "Analytic Services Provider", "http://prh01612.prod.fedex.com:19000/aps/SmartView", "PRH01617", "Finrpt", "Finrpt", "Finrpt", "Connection Process")

y = HypConnect("EssBase", "3720187", "Poweraa1", "Finrpt")

Columns("A:A").Select
Selection.ColumnWidth = 80
Columns("B:B").Select
Selection.ColumnWidth = 60
Columns("C:XFD").Select
Selection.ColumnWidth = 10
Cells.Select
Selection.RowHeight = 15
Range("A1").Select

Dim COUNTRY As String
Dim COUNTRY_LIST As Variant
Dim COUNTRY_NAME As Variant

Dim CA01 As String
Dim CA02 As String
Dim CA03 As String
Dim CA04 As String
Dim CA05 As String
Dim CA06 As String
Dim CA07 As String
Dim CA08 As String
Dim CA09 As String
Dim CA10 As String
Dim CA11 As String
Dim CA12 As String
Dim CA13 As String
Dim CA14 As String
Dim CA15 As String
Dim CA16 As String
Dim CA17 As String
Dim CA18 As String
Dim CA19 As String
Dim CA20 As String
Dim CA21 As String
Dim CA22 As String
Dim CA23 As String
Dim CA24 As String
Dim CA25 As String
Dim CA26 As String
Dim CA27 As String
Dim CA28 As String
Dim CA29 As String
Dim CA30 As String
Dim CA31 As String
Dim CA32 As String
Dim CA33 As String
Dim CA34 As String
Dim CA35 As String
Dim CA36 As String
Dim CA37 As String
Dim CA38 As String
Dim CA39 As String
Dim CA40 As String
Dim CA41 As String
Dim CA42 As String
Dim CA43 As String
Dim CA44 As String
Dim CA45 As String
Dim CA46 As String
Dim CA47 As String
Dim CA48 As String

CA01 = "Anguilla"
CA02 = "Antigua"
CA03 = "Argentina"
CA04 = "Aruba"
CA05 = "Bahamas"
CA06 = "Barbados"
CA07 = "Belize"
CA08 = "Bermuda"
CA09 = "Bolivia"
CA10 = "Brazil"
CA11 = "BVI"
CA12 = "Cayman Islands"
CA13 = "Chile"
CA14 = "Colombia"
CA15 = "Costa Rica"
CA16 = "Curacao"
CA17 = "Dominica"
CA18 = "Dominican Republic"
CA19 = "Ecuador"
CA20 = "El Salvador"
CA21 = "French Guiana"
CA22 = "Grenada"
CA23 = "Guadeloupe"
CA24 = "Guatemala"
CA25 = "Guyana"
CA26 = "Haiti"
CA27 = "Honduras"
CA28 = "Jamaica"
CA29 = "Martinique"
CA30 = "Mexico"
CA31 = "Montserrat"
CA32 = "Nicaragua"
CA33 = "Panama"
CA34 = "Paraguay"
CA35 = "Peru"
CA36 = "Puerto Rico"
CA37 = "Saint Martin - French"
CA38 = "St. Kitts-Nevis"
CA39 = "St. Lucia"
CA40 = "St. Maarten - Dutch"
CA41 = "St. Vincent"
CA42 = "Suriname"
CA43 = "Trinidad And Tobago"
CA44 = "Turks & Caicos Islands"
CA45 = "U.S. Virgin Islands"
CA46 = "United States"
CA47 = "Uruguay"
CA48 = "Venezuela"

COUNTRY_LIST = Array(CA01, CA02, CA03, CA04, CA05, CA06, CA07, CA08, CA09, CA10, CA11, CA12, CA13, CA14, CA15, CA16, CA17, CA18, CA19, CA20, CA21, CA22, CA23, CA24, CA25, CA26, CA27, CA28, CA29, CA30, CA31, CA32, CA33, CA34, CA35, CA36, CA37, CA38, CA39, CA40, CA41, CA42, CA43, CA44, CA45, CA46, CA47, CA48)

ChDir "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses"
Workbooks.Open Filename:="C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0004-0002-Accounts Verification File-All Countries-Expenses.xlsx"

Windows("0004-0001-Accounts Verification File-Program File.xlsm").Activate

For Each COUNTRY_NAME In COUNTRY_LIST
    
Range("B7:AN3774").Select
Selection.ClearContents
Range("A1").Select

    Sheets("Anguilla").Select
    Sheets("MASTER").Visible = True
    Cells.Select
    Selection.Copy
    Sheets("Anguilla").Select
    Cells.Select
    ActiveSheet.Paste
    Range("A1").Select
    Application.CutCopyMode = False
    Sheets("MASTER").Select
    Range("A1").Select
    Sheets("MASTER").Select
    ActiveWindow.SelectedSheets.Visible = False
    Sheets("Anguilla").Select
    Range("A1").Select

Application.Goto Reference:="COMPLETE_VIEW"

If COUNTRY_NAME = "BVI" Then COUNTRY_NAME = "British Virgin Islands"
If COUNTRY_NAME = "St. Kitts-Nevis" Then COUNTRY_NAME = "St. Kitts-Nevis"

Range("B2").Select
ActiveCell.FormulaR1C1 = COUNTRY_NAME

If COUNTRY_NAME = "British Virgin Islands" Then COUNTRY_NAME = "BVI"
If COUNTRY_NAME = "St. Kitts-Nevis" Then COUNTRY_NAME = "St. Kitts-Nevis"
If COUNTRY_NAME = "St. Lucia" Then COUNTRY_NAME = "St. Lucia"
If COUNTRY_NAME = "St. Maarten - Dutch" Then COUNTRY_NAME = "St. Maarten - Dutch"
If COUNTRY_NAME = "St. Vincent" Then COUNTRY_NAME = "St. Vincent"
If COUNTRY_NAME = "Turks & Caicos Islands" Then COUNTRY_NAME = "TCI"

Range("B2").Select
Selection.Copy
Range("B2:AM2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False

ActiveWorkbook.Save

X = HypMenuVRefresh()

Range("A1").Select

    Columns("B:AM").Select
    Selection.ColumnWidth = 60
    Columns("AN:AN").Select
    Selection.ColumnWidth = 20
    
    Range("AN7:AN3774").Select
    Selection.NumberFormat = "#,##0"
        
    Range("AN7").Select
    ActiveCell.Formula = "=IF(B7<>0,1,IF(C7<>0,1,IF(D7<>0,1,IF(E7<>0,1,IF(F7<>0,1,IF(G7<>0,1,IF(H7<>0,1,IF(I7<>0,1,IF(J7<>0,1,IF(K7<>0,1,IF(L7<>0,1,IF(M7<>0,1,IF(N7<>0,1,IF(O7<>0,1,IF(P7<>0,1,IF(Q7<>0,1,IF(R7<>0,1,IF(S7<>0,1,IF(T7<>0,1,IF(U7<>0,1,IF(V7<>0,1,IF(W7<>0,1,IF(X7<>0,1,IF(Y7<>0,1,IF(Z7<>0,1,IF(AA7<>0,1,IF(AB7<>0,1,IF(AC7<>0,1,IF(AD7<>0,1,IF(AE7<>0,1,IF(AF7<>0,1,IF(AG7<>0,1,IF(AH7<>0,1,IF(AI7<>0,1,IF(AJ7<>0,1,IF(AK7<>0,1,IF(AL7<>0,1,IF(AM7<>0,1,0))))))))))))))))))))))))))))))))))))))"
    Range("AN7").Select
    Selection.Copy
    Range("AN7:AN3774").Select
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
    SkipBlanks:=False, Transpose:=False
    
    Columns("AN:AN").Select
    Selection.ColumnWidth = 20
    Application.CutCopyMode = False
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .IndentLevel = 0
        .ShrinkToFit = False
        .ReadingOrder = xlContext
        .MergeCells = False
    End With
    
    Range("AN1:AN6").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    
    Range("AN7:AN3774").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("A1").Select
       
ActiveWorkbook.Save
Range("A1").Select
    
Windows("0004-0001-Accounts Verification File-Program File.xlsm").Activate
Cells.Select
Selection.Copy
Windows("0004-0002-Accounts Verification File-All Countries-Expenses.xlsx").Activate

If COUNTRY_NAME = "BVI" Then Sheets("BVI").Select
If COUNTRY_NAME = "St.Kitts - Nevis" Then Sheets("St. Kitts-Nevis").Select
If COUNTRY_NAME = "St.Lucia" Then Sheets("St. Lucia").Select
If COUNTRY_NAME = "St. Maarten - Dutch" Then Sheets("St. Maarten - Dutch").Select
If COUNTRY_NAME = "Saint Martin - French" Then Sheets("Saint Martin - French").Select
If COUNTRY_NAME = "St.Vincent" Then Sheets("St. Vincent").Select
If COUNTRY_NAME = "Turks & Caicos Islands" Then Sheets("TCI").Select
If COUNTRY_NAME = "U.S. Virgin Islands" Then Sheets("U.S. Virgin Islands").Select

Sheets(COUNTRY_NAME).Select

Cells.Select
ActiveSheet.Paste
Range("A1").Select
ActiveWorkbook.Save

Windows("0004-0001-Accounts Verification File-Program File.xlsm").Activate

Range("A1").Select
Application.CutCopyMode = False

Calculate
Sleep (3000)

Next

Application.DisplayAlerts = False
 
Windows("0004-0002-Accounts Verification File-All Countries-Expenses.xlsx").Activate
Sheets("Anguilla").Select
Range("A1").Select
ActiveWorkbook.Save

ActiveWorkbook.SaveAs "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0004-0002-Accounts Verification File-All Countries-Expenses Details.xlsx"
ActiveWorkbook.SaveAs "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0004-0002-Accounts Verification File-All Countries-Expenses.xlsx"

Sheets(Array("Anguilla", "Antigua", "Argentina", "Aruba", "Bahamas", "Barbados", "Belize", "Bermuda", "Bolivia", "Brazil", "BVI", "Cayman Islands", "Chile", "Colombia" _
       , "Costa Rica", "Curacao", "Dominica", "Dominican Republic", "Ecuador", "El Salvador", "French Guiana", "Grenada", "Guadeloupe", "Guatemala", "Guyana", "Haiti", _
       "Honduras", "Jamaica", "Martinique", "Mexico", "Montserrat", "Nicaragua", "Panama", "Paraguay", "Peru", "Puerto Rico", "St. Kitts-Nevis", "St. Lucia", _
       "St. Maarten", "Saint Martin - French", "St. Maarten - Dutch", "St. Vincent", "Suriname", "Trinidad And Tobago", "TCI", "U.S. Virgin Islands", _
       "Uruguay", "Venezuela", "United States")).Select
  
Cells.Select
Selection.Copy
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False
    
Columns("B:AK").Select
Selection.Delete Shift:=xlToLeft
Range("A1").Select
   
Sheets("Anguilla").Select
Range("A1").Select
ActiveWorkbook.Save
ActiveWindow.Close

Windows("0004-0001-Accounts Verification File-Program File.xlsm").Activate

Range("B2").Select
ActiveCell.FormulaR1C1 = "'Anguilla"

Range("B2").Select
Selection.Copy
Range("B2:AM2").Select
Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
:=False, Transpose:=False

Range("B7:AM3774").Select
Selection.ClearContents

Columns("AN:AN").Select
Selection.Delete Shift:=xlToLeft

    Range("A1:AM3774").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

Range("A1").Select

ActiveWorkbook.Save

Application.CutCopyMode = False

End Sub

Public Sub FEDEXDATA01_DATABASES()

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Dim appAccess As Object

Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Anguilla.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Anguilla"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Antigua.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Antigua"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Argentina.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Argentina"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Aruba.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Aruba"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Bahamas.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Bahamas"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Barbados.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Barbados"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Belize.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Belize"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Bermuda.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Bermuda"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Bolivia.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Bolivia"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Brazil.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Brazil"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-BVI.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-BVI"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Cayman Islands.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Cayman Islands"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Chile.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Chile"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Colombia.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Colombia"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Costa Rica.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Costa Rica"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Curacao.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Curacao"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Dominica.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Dominica"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Dominican Republic.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Dominican Republic"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Ecuador.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Ecuador"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-El Salvador.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-El Salvador"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-French Guiana.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-French Guiana"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Grenada.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Grenada"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Guadeloupe.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Guadeloupe"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Guatemala.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Guatemala"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Guyana.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Guyana"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Haiti.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Haiti"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Honduras.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Honduras"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Jamaica.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Jamaica"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Martinique.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Martinique"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Mexico - 1A.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Mexico"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Mexico - 2A.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Mexico"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Montserrat.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Montserrat"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Nicaragua.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Nicaragua"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Panama.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Panama"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Paraguay.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Paraguay"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Peru.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Peru"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Puerto Rico.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Puerto Rico"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Saint Martin - French.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Saint Martin - French"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-St. Kitts-Nevis.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-St Kitts-Nevis"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-St. Lucia.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-St Lucia"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-St. Maarten - Dutch.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-St Maarten - Dutch"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-St. Vincent.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-St Vincent"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Suriname.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Suriname"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-TCI.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-TCI"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Trinidad And Tobago.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Trinidad And Tobago"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-U.S. Virgin Islands.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-US Virgin Islands"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Uruguay.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Uruguay"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-Venezuela.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-Venezuela"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing


Set appAccess = Nothing

Set appAccess = CreateObject("Access.Application")
appAccess.OpenCurrentDatabase "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\0001-0001-Data Information-United States.accdb"
appAccess.DoCmd.RunMacro "0001-0001-Process Country-United States"

appAccess.CloseCurrentDatabase

Set appAccess = Nothing

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

MsgBox "All Databases Have Been Processed."

Dim FedexDRMPath As String

FedexDRMPath = "C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\"

Call Shell("explorer.exe C:\Users\3720187\Documents\2018 Fedex Files\2018 Fedex Data Information System-Expenses\0001-02-Data Information Databases\", vbMaximizedFocus)
  
End Sub
