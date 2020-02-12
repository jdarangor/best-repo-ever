Attribute VB_Name = "Fedex"
Declare PtrSafe Function HypCreateConnection Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtProvider As Variant, ByVal vtProviderURL As Variant, ByVal vtServerName As Variant, ByVal vtApplicationName As Variant, ByVal vtDatabaseName As Variant, ByVal vtFriendlyName As Variant, ByVal vtDescription As Variant) As Long
Declare PtrSafe Function HypConnect Lib "HsAddin" (ByVal vtSheetName As Variant, ByVal vtUserName As Variant, ByVal vtPassword As Variant, ByVal vtFriendlyName As Variant) As Long
Declare PtrSafe Function HypMenuVRefresh Lib "HsAddin.dll" () As Long
Sub Fedex_A01_Main_Console_Program()

If Worksheets("Main Console").Range("I18") = "Initiate Data Information Process" Then

Application.Run "'Main Console File - Data Information.xlsm'!Fedex_A02_Process"

End If


    If Worksheets("Main Console").Range("I18") = "Initiate Essbase Data Process" Then
    
        Call Fedex_A03_Process
    
    End If
    
        
If Worksheets("Main Console").Range("I18") = "Reset Databases" Then

Application.Run "'Main Console File - Data Information.xlsm'!Fedex_A04_Process"

End If


If Worksheets("Main Console").Range("I18") = "Process Databases" Then

Application.Run "'Main Console File - Data Information.xlsm'!Fedex_A08_Process"

End If


If Worksheets("Main Console").Range("I18") = "Initiate Complete Data Process" Then

Application.Run "'Main Console File - Data Information.xlsm'!Fedex_A05_Process"

End If


If Worksheets("Main Console").Range("I18") = "Process All Essbase Files Into Main Database" Then

Application.Run "'Main Console File - Data Information.xlsm'!Fedex_A07_Process"

End If
        
      
    
End Sub


Sub Fedex_A03_Process()

    Application.DisplayAlerts = False
    Application.ScreenUpdating = True
    
    Debug.Print Range("g30").Value & " Process started: " & Now
    'Workbooks.Open Filename:=CurrentPath() & "0001-0001-Complete Data File-Program File.xlsm"
    
    Range("A1").Select
    Stop
    
    ActiveWorkbook.Save
    Application.Run "'" & CurrentPath() & "Jda 0001-0001-Complete Data File-Program File.xlsm'!Fedex_Data_0001"
    
    Windows("Jda 0001-0001-Complete Data File-Program File.xlsm").Activate
    ActiveWorkbook.Close True
    
    Windows("Jda 0001-0002-Complete Data File-Expenses.xlsm").Activate
    ActiveWorkbook.Close True
    
    Windows("Jda Main Console File - Data Information.xlsm").Activate
    Range("I18:T18").Select

    Debug.Print Range("g30").Value & " Process ended: " & Now

End Sub

Public Function CurrentPath() As String
    CurrentPath = ThisWorkbook.Path & "\"
End Function
