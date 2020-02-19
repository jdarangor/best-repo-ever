Sub RenameColumns(jdaTb as string)
    Dim i As Integer
    Dim f As Integer
    Dim n As String
    Dim t As String
    Dim fld As Field
    
    
    f = CurrentDb.TableDefs(jdaTb).Fields.Count
    
    For i = 0 To f - 1
        'Debug.Print i, Application.CurrentDb.TableDefs("LAC Data Local Cur Tb").Fields(i).Name
        n = Replace(Application.CurrentDb.TableDefs("LAC Data Local Cur Tb").Fields(i).Name, "-", "_")
        Debug.Print n
    Next


End Sub
