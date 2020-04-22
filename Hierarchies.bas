Attribute VB_Name = "Hierarchies"
Option Explicit

Public essBaseConn As Boolean
Public essPass As String

Sub Example_HypDisconnectAll()
    Dim sts As Long
    
    sts = HypDisconnectAll()

End Sub

Sub Example_HypConnect()
    Dim x As Long
    
    essBaseConn = False
    
    UserForm1.Show
    Application.Cursor = xlWait
    
    x = HypConnect(Empty, "3811xyz", essPass, "xyz#####_XyzXYZ_Xyz_DB-JDA")
'    Debug.Print "Connection result: " & x & " " & Now()
    Application.Cursor = xlDefault
    If x <> 0 Then
        MsgBox "Please review you have your EssBase AddIn available", vbCritical, "Connection Result"
        'essBaseConn = False
    Else
        essBaseConn = True
    End If
    
End Sub

Sub SpreadLevels()
    Dim i  As Long
    Dim strCol As String
    Dim tmStarted As Date: tmStarted = Now()
    
    Application.StatusBar = "Hierarchy creation started at " & tmStarted
    Debug.Print "Hierarchy creation started at " & tmStarted
    
    'Range(Range(Range("F4")), Cells(Range(ActiveCell.Address).Row + Range("D4"), Range(ActiveCell.Address).Column + Range("E4"))).Select
    Range("D8:R9000").Select
    Selection.ClearContents
    ' Place Generation 0 to column
    strCol = Range("G4").Value
    
    Range(Cells(4, 6)).Select  'Where to start
    
    For i = 1 To Cells(4, 4).Value  'Rows to run for
        
        Select Case ActiveCell.Offset(0, -2).Value
        
            Case 0
                ActiveCell.Offset(0, 0).Value = ActiveCell.Offset(0, -3).Value
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 5
                ActiveCell.Offset(0, 1).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 10
                ActiveCell.Offset(0, 2).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 15
                ActiveCell.Offset(0, 3).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            Case 20
                ActiveCell.Offset(0, 4).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 25
                ActiveCell.Offset(0, 5).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 30
                ActiveCell.Offset(0, 6).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 35
                ActiveCell.Offset(0, 7).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 40
                ActiveCell.Offset(0, 8).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 45
                ActiveCell.Offset(0, 9).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 50
                ActiveCell.Offset(0, 10).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 55
                ActiveCell.Offset(0, 11).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 60
                ActiveCell.Offset(0, 12).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 65
                ActiveCell.Offset(0, 13).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case Else
                ActiveCell.Value = "Pending"
                
        End Select
      
        ActiveCell.Offset(1, 0).Select
    
    Next
    
    ActiveWorkbook.Save
    Debug.Print "Spread Data to Generations at " & Now()
    Call CopyHierarchy
    Application.CutCopyMode = False
    
    MsgBox "Hierarchy created.", , ActiveWorkbook.Name

End Sub

Sub SpreadAccountLevels()
    Dim i  As Long
    Dim strCol As String
    Dim tmStarted As Date: tmStarted = Now()
    Dim intLevels As Integer
    
    intLevels = Range("E4").Value
    Debug.Print "Hierarchy creation started at " & tmStarted
    
    'Range(Range(Range("F4")), Cells(Range(ActiveCell.Address).Row + Range("D4"), Range(ActiveCell.Address).Column + Range("E4"))).Select
    Range("D8:R9000").Select
    Selection.ClearContents
    ' Place Generation 0 to column
    strCol = Range("G4").Value
    
    Range(Cells(4, 6)).Select  'Where to start
    
    For i = 1 To Cells(4, 4).Value  'Rows to run for
        
        
        Select Case ActiveCell.Offset(0, -2).Value
        
            Case 0
                ActiveCell.Offset(0, intLevels - 0).Value = ActiveCell.Offset(0, -3).Value
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 5
                ActiveCell.Offset(0, intLevels - 1).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 10
                ActiveCell.Offset(0, intLevels - 2).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 15
                ActiveCell.Offset(0, intLevels - 3).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            Case 20
                ActiveCell.Offset(0, intLevels - 4).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 25
                ActiveCell.Offset(0, intLevels - 5).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 30
                ActiveCell.Offset(0, intLevels - 6).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 35
                ActiveCell.Offset(0, intLevels - 7).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 40
                ActiveCell.Offset(0, intLevels - 8).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 45
                ActiveCell.Offset(0, intLevels - 9).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 50
                ActiveCell.Offset(0, intLevels - 10).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 55
                ActiveCell.Offset(0, intLevels - 11).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 60
                ActiveCell.Offset(0, intLevels - 12).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case 65
                ActiveCell.Offset(0, intLevels - 13).Value = Trim(ActiveCell.Offset(0, -3).Value)
                ActiveCell.Offset(0, (Range(strCol & ":" & strCol).Column) - (ActiveCell.Column)).Value = Trim(ActiveCell.Offset(0, -3).Value)
            
            Case Else
                ActiveCell.Value = "Pending"
                
        End Select
    
       ' Debug.Print "Hierarchy Generation " & ii & " at " & Now
        ActiveCell.Offset(1, 0).Select
        DoEvents
    Next
    
    ActiveWorkbook.Save
    Debug.Print "Spread Data to Generations at " & Now()
    Call CopyHierarchy
    Application.CutCopyMode = False

End Sub



Sub CopyHierarchy()
    Dim i As Integer
    Dim ii As Integer
    
    Range(Cells(4, 6)).Select 'Where to start
    ActiveCell.Copy
    
    For ii = 1 To Cells(4, 5).Value + 1     'Generations to run for
        
        For i = 1 To Cells(4, 4).Value      'Rows to run for
            
            If ActiveCell.Offset(1, 0) = "" Then
                ActiveCell.Offset(1, 0).Select
'                Debug.Assert ii <> 9
                If ActiveCell.Offset(0, -(ii)).Value > ii Then
                    ActiveSheet.Paste
                End If
            Else
                ActiveCell.Offset(1, 0).Copy
                ActiveCell.Offset(1, 0).Select
            End If
                        
        Next
        Application.StatusBar = "Already copy hierarchy Generation " & ii & " at " & Now
        Debug.Print "Already copy hierarchy Generation " & ii & " at " & Now
        Application.CutCopyMode = False
        Range(Cells(4, 6)).Select
        ActiveCell.Offset(0, ii).Select
        ActiveCell.Copy
        
        DoEvents
    Next

    Debug.Print "Hierarchy creation ended at " & Now()

End Sub

Public Function ConvertTimeToDecimal(timeIn As Double) As Double
    Dim lngNatural As Double
    
    lngNatural = WorksheetFunction.RoundDown(timeIn, 0)
    ConvertTimeToDecimal = lngNatural + ((timeIn - lngNatural) / 6 * 10)
End Function

Sub ReviewDimMembers()
    Dim cbItems As Long
    Dim i As Long
    Dim sts As Variant
    Dim vt As Variant
    Dim vtDims As Variant, vtDimNames As Variant
    Dim vtMbrs As Variant, vtMbr As Variant
    Dim strPreviousMemb As String
    Dim wsCur As Worksheet
    
'    Call Example_HypConnect
'    If Not essBaseConn Then
'        Exit Sub
'    End If
' HypQueryMembers (vtSheetName, vtMemberName, vtPredicate, vtOption, vtDimensionName, vtInput1, vtInput2, vtMemberArray)
' sts = HypQueryMembers(Empty, "Profit", HYP_CHILDREN, Empty, Empty, Empty, Empty,vArray)
' sts = HypQueryMembers(Empty, "Profit", HYP_DESCENDANTS, Empty, Empty, Empty,Empty, vArray)
' sts = HypQueryMembers(Empty, "Profit", HYP_BOTTOMLEVEL, Empty, Empty, Empty,Empty, vArray)
' sts = HypQueryMembers(Empty, "Sales", HYP_SIBLINGS, Empty, Empty, Empty, Empty,vArray)
' sts = HypQueryMembers(Empty, "Sales", HYP_SAMELEVEL, Empty, Empty, Empty, Empty,vArray)
' sts = HypQueryMembers(Empty, "Sales", HYP_SAMEGENERATION, Empty, Empty, Empty,Empty, vArray)
' sts = HypQueryMembers(Empty, "Sales", HYP_PARENT, Empty, Empty, Empty, Empty,vArray)
' sts = HypQueryMembers(Empty, "Sales", HYP_DIMENSION, Empty, Empty, Empty, Empty,vArray)
' sts = HypQueryMembers(Empty, "Year", HYP_NAMEDGENERATION, Empty, "Year", "Quarter", Empty, vArray)
' sts = HypQueryMembers(Empty, "Product", HYP_NAMEDLEVEL, Empty, "Product", "SKU",Empty, vArray)
' sts = HypQueryMembers(Empty, "Product", HYP_SEARCH, HYP_ALIASESONLY, "Product","Cola", Empty, vArray)
' sts = HypQueryMembers(Empty, "Year", HYP_WILDSEARCH, HYP_MEMBERSONLY, "Year","J*", Empty, vArray)
' sts = HypQueryMembers(Empty, "Market", HYP_USERATTRIBUTE, Empty, "Market", "Major Market ", Empty, vArray)"
' sts = HypQueryMembers(Empty, "Sales", HYP_ANCESTORS, Empty, Empty, Empty, Empty,vArray)
' sts = HypQueryMembers(Empty, "Jan", HYP_DTSMEMBER, Empty, Empty, Empty, Empty,vArray)
' sts = HypQueryMembers(Empty, "Product", HYP_DIMUSERATTRIBUTES, Empty, Empty,Empty, Empty, vArray)

' HypQueryMembers (vtSheetName, vtMemberName, vtPredicate, vtOption, vtDimensionName, vtInput1, vtInput2, vtMemberArray)
    
    Application.Cursor = xlWait
    Set wsCur = Worksheets("Hrchy Chng Review")
    
    wsCur.Range("A8:F10000").Clear
    ' worked sts = HypQueryMembers("Pull Members", "Fiscal Year", 2, Empty, "Fiscal Year", Empty, Empty, vt)
    'retuns region 1.. descendants sts = HypQueryMembers("Pull Members", "1..", 2, Empty, "Region", Empty, Empty, vt)
    'retuns region 2.. descendants
    vtDims = Array(Array("1..", "2..", "R.."), _
        Array("LATIN AM", "US DIV"), _
        Array("LA Region", "US Region", "XXX"), _
        Array("PeopleSoft LEE"), _
        Array("L........."))
    
    vtDimNames = Array("Region", _
         "Country", _
         "Pseudo", _
         "Account", _
         "Entity")
    
    Dim o As Integer
    o = 0
    For Each vtMbrs In vtDims
        
        wsCur.Cells(8, 1 + o).Select
        
        For Each vtMbr In vtMbrs
            sts = HypQueryMembers("Hrchy Chng Review", vtMbr, 2, Empty, vtDimNames(o), Empty, Empty, vt)
            
            If IsArray(vt) Then
                cbItems = UBound(vt) + 1
                'msgbox ("Number of elements for member '" & vtMbr & "' = " + Str(cbItems))
                
                'Debug.Print "Member = " & vtMbr
                ActiveCell.Value = "'" & vtMbr
                ActiveCell.Offset(1, 0).Activate
                
            If strPreviousMemb <> vt(0) Then
                For i = 0 To UBound(vt)
                    'MsgBox ("Member = " + vt(i))
                    'Debug.Print "Member = " + vt(i)
                    ActiveCell.Value = "'" & vt(i)
                    ActiveCell.Offset(1, 0).Activate
                    strPreviousMemb = vt(0)
                Next
            End If
            Else
            
                MsgBox ("Return Value = " + Str(vt))
            
            End If
        Next
        
        o = o + 1
    Next
    
'    Call Example_HypDisconnectAll
    
    wsCur.Range("A8").Select
    Application.Cursor = xlDefault
    
    If Range("G5") <> 0 Or Range("H5") <> 0 Or Range("I5") <> 0 Or Range("J5") <> 0 Or Range("K5") <> 0 Then
        MsgBox "Please recreate the hierarchies, there is one or more changes.", vbCritical, "FedEx Hierarchies Review."
    Else
        MsgBox "Review has been completed", vbInformation, "FedEx Hierarchies Review."
    End If
    
    Set wsCur = Nothing
    
    'test=array(1,2,3)
    'Range("T1:T3") = WorksheetFunction.Transpose(test)

End Sub



Sub ReviewAccountMembers()
    Dim cbItems As Long
    Dim i As Long
    Dim sts As Variant
    Dim vt As Variant
    Dim vtDims As Variant, vtDimNames As Variant
    Dim vtMbrs As Variant, vtMbr As Variant
    Dim strPreviousMemb As String
    Dim wsCur As Worksheet
    
' HypQueryMembers (vtSheetName, vtMemberName, vtPredicate, vtOption, vtDimensionName, vtInput1, vtInput2, vtMemberArray)
    
    Application.Cursor = xlWait
    Set wsCur = Worksheets("Hrchy Chng Review")
    
    wsCur.Range("D8:D10000").Clear
    ' worked sts = HypQueryMembers("Pull Members", "Fiscal Year", 2, Empty, "Fiscal Year", Empty, Empty, vt)
    'retuns region 1.. descendants sts = HypQueryMembers("Pull Members", "1..", 2, Empty, "Region", Empty, Empty, vt)
    'retuns region 2.. descendants
    vtDims = Array( _
        Array("PeopleSoft LEE"))
    
    vtDimNames = Array( _
         "Account")
    
    Dim o As Integer
    o = 0
    For Each vtMbrs In vtDims
        
        wsCur.Cells(8, 4 + o).Select ' 8 row 4 col (4 + 0)
        
        For Each vtMbr In vtMbrs
            sts = HypQueryMembers("Hrchy Chng Review", vtMbr, 2, Empty, vtDimNames(o), Empty, Empty, vt)
            
            If IsArray(vt) Then
                cbItems = UBound(vt) + 1
                'msgbox ("Number of elements for member '" & vtMbr & "' = " + Str(cbItems))
                
                'Debug.Print "Member = " & vtMbr
                ActiveCell.Value = "'" & vtMbr
                ActiveCell.Offset(1, 0).Activate
                
            If strPreviousMemb <> vt(0) Then
                For i = 0 To UBound(vt)
                    'MsgBox ("Member = " + vt(i))
                    'Debug.Print "Member = " + vt(i)
                    ActiveCell.Value = "'" & vt(i)
                    ActiveCell.Offset(1, 0).Activate
                    strPreviousMemb = vt(0)
                Next
            End If
            Else
            
                MsgBox ("Return Value = " + Str(vt))
            
            End If
        Next
        
        o = o + 1
    Next
    
'    Call Example_HypDisconnectAll
    
    wsCur.Range("A8").Select
    Application.Cursor = xlDefault
    
    If Range("G5") <> 0 Or Range("H5") <> 0 Or Range("I5") <> 0 Or Range("J5") <> 0 Or Range("K5") <> 0 Then
        MsgBox "Please recreate the hierarchies, there is one or more changes.", vbCritical, "FedEx Hierarchies Review."
    Else
        MsgBox "Review has been completed", vbInformation, "FedEx Hierarchies Review."
    End If
    
    Set wsCur = Nothing
    
    'test=array(1,2,3)
    'Range("T1:T3") = WorksheetFunction.Transpose(test)

End Sub

Sub apostrophe()
Dim i As Integer: i = 0

Range("a8").Select

Do While ActiveCell.Value <> ""
    ActiveCell.Value = "'" & ActiveCell.Value
    ActiveCell.Offset(1, 0).Select
    i = i + 1
    If i > 10000 Then GoTo jdaSalida
Loop
jdaSalida:
End Sub
