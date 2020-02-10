Attribute VB_Name = "Hierarchies"
Option Explicit

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

