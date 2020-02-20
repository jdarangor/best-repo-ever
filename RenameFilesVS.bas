Attribute VB_Name = "Module1"
Option Explicit

Sub rename()

Dim d As String
Dim f As Variant
 Dim i As Integer
         d = Dir("C:\Users\3811756\Documents\JuanArango\INTERLINE\FY2020 Interline\Older\code\3rd Level\")

    For i = 1 To 10
        Debug.Print d
        d = Dir()
    Next
    
End Sub

Sub ShowFolderList(folderspec)
    'C:\Users\3811756\Documents\JuanArango\PowerBI\D3\tutorialpoint
    Dim fs, f, f1, fc, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each f1 In fc
        s = s & f1.Name
        s = s & vbCrLf
    Next
    MsgBox s
End Sub

Sub ShowFileAccessInfo(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = f.Name & " on Drive " & UCase(f.Drive) & vbCrLf
    s = s & "Created: " & f.DateCreated & vbCrLf
    s = s & "Last Accessed: " & f.DateLastAccessed & vbCrLf
    s = s & "Last Modified: " & f.DateLastModified
    MsgBox s, 0, "File Access Info"
End Sub

Sub RenameScript(filespec)
    Dim fs, f, s
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFile(filespec)
    s = Split(f.Name, ".")
    f.Name = s(0) & ".txt"
End Sub

Sub RenameFoldersFileScript(folderspec, Optional frm As String, Optional t As String)
    Dim fs, f, fl, fc, s, sa As Variant
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set f = fs.GetFolder(folderspec)
    Set fc = f.Files
    For Each fl In fc
        s = Split(fl.Name, ".")
        If s(1) = frm Then
    '        sa = sa & s(0) & "." & t & vbCrLf
            fl.Name = s(0) & "." & t
        End If
    Next
    '   MsgBox sa
End Sub
Sub callRename()
    Call RenameFoldersFileScript("C:\Users\3811756\Documents\JuanArango\PowerBI\D3\tutorialpoint\", "txt", "html")
End Sub
