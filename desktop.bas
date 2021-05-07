Attribute VB_Name = "Module4"
Option Explicit

Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type

Public FState As FormState              ' Array of user-defined types
Public gFindString As String            ' Holds the search text.
Public gFindCase As Integer             ' Key for case sensitive search
Public gFindDirection As Integer        ' Key for search direction.
Public gCurPos As Integer               ' Holds the cursor location.
Public gFirstTime As Integer            ' Key for start position.
Public Const ThisApp = "MDINote"        ' Registry App constant.
Public Const ThisKey = "Recent Files"   ' Registry Key constant.

Sub FileNew()
    Dim intResponse As Integer
    
    ' If the file has changed, save it
    If FState.Dirty = True Then
        intResponse = FileSave
        If intResponse = False Then Exit Sub
    End If
    ' Clear the textbox and update the caption.
    Form1.Text1.Text = ""
    Form1.Text2.Text = "None"
End Sub
Function FileSave() As Integer
    Dim strFilename As String

    If Form1.Text2.Text = "none" Then
        ' The file hasn't been saved yet.
        ' Get the filename, and then call the save procedure, GetFileName.
        strFilename = GetFileName(strFilename)
    Else
        ' The form's Caption contains the name of the open file.
        strFilename = Form1.Text2.Text
    End If
    ' Call the save procedure. If Filename = Empty, then
    ' the user chose Cancel in the Save As dialog box; otherwise,
    ' save the file.
    If strFilename <> "" Then
        SaveFileAs strFilename
        FileSave = True
    Else
        FileSave = False
    End If
End Function
Sub FileOpenProc()
    Dim intRetVal
    Dim intResponse As Integer
    Dim strOpenFileName As String
    
    ' If the file has changed, save it
    If FState.Dirty = True Then
        intResponse = FileSave
        If intResponse = False Then Exit Sub
    End If
    On Error Resume Next
    Form1.CommonDialog1.Filter = "Text Files" & _
    "(*.txt)|*.txt|Batch Files (*.bat)|*.bat"

    Form1.CommonDialog1.Filename = ""
    Form1.CommonDialog1.ShowOpen

    If err <> 32755 Then    ' User chose Cancel.
        strOpenFileName = Form1.CommonDialog1.Filename
        ' If the file is larger than 65K, it can't
        ' be opened, so cancel the operation.
        If FileLen(strOpenFileName) > 65000 Then
            Exit Sub
        End If
        
        OpenFile (strOpenFileName)
        End If
End Sub

Function GetFileName(Filename As Variant)
    ' Display a Save As dialog box and return a filename.
    ' If the user chooses Cancel, return an empty string.
    On Error Resume Next
    Form1.CommonDialog1.Filename = Filename
    Form1.CommonDialog1.ShowSave
    If err <> 32755 Then    ' User chose Cancel.
        GetFileName = Form1.CommonDialog1.Filename
    Else
        GetFileName = ""
    End If
End Function


Sub OpenFile(Filename)
    Dim fIndex As Integer
    
    On Error Resume Next
    ' Open the selected file.
    Open Filename For Input As #1
    If err Then
        MsgBox "Can't open file: " + Filename
        Exit Sub
    End If
    ' Change the mouse pointer to an hourglass.
    Screen.MousePointer = 11
    
    ' Change the form's caption and display the new text.
    Form1.Text1.Text = Input(LOF(1), 1)
    FState.Dirty = False
    Close #1
    ' Reset the mouse pointer.
    Screen.MousePointer = 0
End Sub

Sub SaveFileAs(Filename)
    On Error Resume Next
    Dim strContents As String

    ' Open the file.
    Open Filename For Output As #1
    ' Place the contents of the Q - Pad into a variable.
    strContents = Form1.Text1.Text
    ' Display the hourglass mouse pointer.
    Screen.MousePointer = 11
    ' Write the variable contents to a saved file.
    Print #1, strContents
    Close #1
    ' Reset the mouse pointer.
    Screen.MousePointer = 0
    ' Set the form's caption.
    If err Then
        MsgBox Error, 48, App.Title
    Else
        ' Reset the dirty flag.
        FState.Dirty = False
    End If
End Sub


