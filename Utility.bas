Attribute VB_Name = "Utility"
Option Explicit


Public Function BtwnColonAndDot()
    Dim intColon As Integer
    Dim intDot As Integer
    
    intColon = WorksheetFunction.Find(":", ActiveCell.Text) + 1
    'Debug.Print "Colon: " & intColon
    intDot = WorksheetFunction.Find(".", ActiveCell.Text)
    'Debug.Print "Dot: " & intDot
    
    BtwnColonAndDot = Trim(Mid(ActiveCell.Text, intColon, intDot - intColon))

End Function

Public Function GetPSV(strPuntuation) As Variant 'Get Puntuation Separted Value
    Dim vrtItem As Variant, vrtItems As Variant
    Dim i As Integer
    
    vrtItems = Split(BtwnColonAndDot, strPuntuation)
        
    For i = 0 To UBound(vrtItems)
        vrtItems(i) = Trim(vrtItems(i))
    Next
    
    GetPSV = vrtItems

End Function

Public Sub PlaceItems() 'ItemsSet As Variant)
    Dim vrtItem As Variant, vrtItems As Variant
    Dim lngItemRow As Long, lngSetRow As Long
    
        'initial row
        Range("A1").Select
        lngSetRow = ActiveCell.Row
        lngItemRow = 1
        
    Do While (ActiveCell.Text <> "")
        'read ingridients
        vrtItems = GetPSV(",")
        
        'place ingridients
        For Each vrtItem In vrtItems
            Cells(lngItemRow, 5) = vrtItem
            lngItemRow = lngItemRow + 1
        Next
        
        'next ingridients
        ActiveCell.Offset(1, 0).Activate
        
    Loop
End Sub

