VERSION 5.00
Begin VB.UserControl ctlAutoType 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ClipControls    =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.ListBox lstHistory 
      Height          =   1035
      Left            =   930
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1200
      Visible         =   0   'False
      Width           =   3045
   End
   Begin VB.TextBox txtTypeIn 
      Height          =   315
      Left            =   15
      TabIndex        =   1
      Top             =   15
      Visible         =   0   'False
      Width           =   2070
   End
   Begin VB.ComboBox cmbTypeIn 
      Height          =   315
      Left            =   15
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   15
      Width           =   3705
   End
End
Attribute VB_Name = "ctlAutoType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'************************************************************************
'AutoType Control v1.0 Copyright 1999 By NeoText
'
'
'Support:
'   flash@quakeclan.net
'   neotext@quakeclan.net
'   neotext@email.com
'
'   http://www.quakeclan.net/neotext/
'
'
'Terms of Agreement:
' By using this source code, you agree to the following terms...
'  1) You may use this source code in personal projects and may compile
'     it into an .exe/.dll/.ocx and distribute it in binary format
'     freely and with no charge.
'  2) You MAY NOT redistribute this source code (for example to a
'     web site) without written permission from the original author.
'     Failure to do so is a violation of copyright laws.
'  3) You may link to this code from another website, provided it
'     is not wrapped in a frame.
'  4) The author of this code may have retained certain additional
'     copyright rights.If so, this is indicated in the author's
'     description.
'************************************************************************


Enum Styles
    sList = 1
    sText = 2
End Enum

Private Const Default_Text = "AutoType Control by NeoText (http://www.quakeclan.net/neotext/)"

Private cancelTracking As Boolean


Private isEnabled As Boolean
Private isCaseSensitive As Boolean
Private myStyle As Integer
Private myTracking As Integer
Private maxHistory As Integer
Private myToolTipText As String


Public Event Change()
Public Event Click()
Public Event DblClick()
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Public Event KeyPress(KeyAscii As Integer)
Public Event KeyUp(KeyCode As Integer, Shift As Integer)


Private Function IsOnListEx(tList, Item) As Integer
Dim cnt As Long
Dim ItemFound As Integer
Dim ItemLen As Integer
ItemLen = Len(LCase(Trim(Item)))
cnt = 0
ItemFound = -1
Do Until cnt = tList.ListCount Or ItemFound > -1
    If Left(LCase(Trim(tList.List(cnt))), ItemLen) = LCase(Trim(Item)) Then ItemFound = cnt
    cnt = cnt + 1
    Loop
IsOnListEx = ItemFound
End Function
Private Function IsOnList(tList, Item) As Integer
Dim cnt As Long
Dim ItemFound As Integer
Dim ItemLen As Integer
ItemLen = Len(Item)
cnt = 0
ItemFound = -1
Do Until cnt = tList.ListCount Or ItemFound > -1
    If Left(tList.List(cnt), ItemLen) = Item Then ItemFound = cnt
    cnt = cnt + 1
    Loop
IsOnList = ItemFound
End Function
Private Sub TrackText()
On Error GoTo catch
If Not cancelTracking Then
    cancelTracking = True
    Dim srhText As String
    Dim srhLen As Integer
    Dim lstIndex As Integer
    Select Case myStyle
        Case sList
            srhText = cmbTypeIn.Text
        Case sText
            srhText = txtTypeIn.Text
    End Select
    srhLen = Len(srhText)
    If srhText <> "" Then
        If isCaseSensitive Then
            lstIndex = IsOnList(lstHistory, srhText)
        Else
            lstIndex = IsOnListEx(lstHistory, srhText)
            End If
        If lstIndex > -1 Then
            Select Case myStyle
                Case sText
                    txtTypeIn.Text = lstHistory.List(lstIndex)
                    txtTypeIn.SelStart = srhLen
                    txtTypeIn.SelLength = Len(txtTypeIn.Text) - srhLen
                Case sList
                    cmbTypeIn.Text = lstHistory.List(lstIndex)
                    cmbTypeIn.SelStart = srhLen
                    cmbTypeIn.SelLength = Len(cmbTypeIn.Text) - srhLen
            End Select
            End If
        End If
    cancelTracking = False
    End If
Exit Sub
catch:
    Err.Clear
End Sub


Public Sub ListDown()
On Error Resume Next
If Not myStyle = sText Then
    cmbTypeIn.SetFocus
    SendKeys "%{DOWN}", True
    End If
Err.Clear
End Sub
Public Sub ListUp()
On Error Resume Next
If Not myStyle = sText Then
    cmbTypeIn.SetFocus
    SendKeys "%{UP}", True
    End If
Err.Clear
End Sub


Public Sub SetFocus()
On Error Resume Next
Select Case myStyle
    Case sList
        cmbTypeIn.SetFocus
    Case sText
        txtTypeIn.SetFocus
End Select
Err.Clear
End Sub


Public Property Get SelStart() As Integer
Select Case myStyle
    Case sList
        SelStart = cmbTypeIn.SelStart
    Case sText
        SelStart = txtTypeIn.SelStart
End Select
End Property
Public Property Let SelStart(ByVal newValue As Integer)
Select Case myStyle
    Case sList
        cmbTypeIn.SelStart = newValue
    Case sText
        txtTypeIn.SelStart = newValue
End Select
End Property


Public Property Get SelLength() As Integer
Select Case myStyle
    Case sList
        SelLength = cmbTypeIn.SelLength
    Case sText
        SelLength = txtTypeIn.SelLength
End Select
End Property
Public Property Let SelLength(ByVal newValue As Integer)
Select Case myStyle
    Case sList
        cmbTypeIn.SelLength = newValue
    Case sText
        txtTypeIn.SelLength = newValue
End Select
End Property


Public Sub SetToList(ByVal lstBox As Variant)
Me.ClearHistory
Dim cnt As Integer
If lstBox.ListCount > 0 Then
    For cnt = 0 To lstBox.ListCount - 1
        Me.AddItem lstBox.List(cnt)
        Next
    End If
End Sub


Public Sub AddItem(ByVal lstText As String)
If lstHistory.ListCount < maxHistory - 1 Then
    Dim lstIndex As Integer
    If isCaseSensitive Then
        lstIndex = IsOnList(lstHistory, lstText)
    Else
        lstIndex = IsOnListEx(lstHistory, lstText)
        End If
    If lstIndex = -1 Then
        lstHistory.AddItem lstText
        cmbTypeIn.AddItem lstText
        End If
    End If
End Sub
Public Sub RemoveItem(ByVal lstIndex As Integer)
lstHistory.RemoveItem lstIndex
cmbTypeIn.RemoveItem lstIndex
End Sub
Public Sub ClearHistory()
Do Until lstHistory.ListCount <= 0
    lstHistory.RemoveItem 0
    Loop
Do Until cmbTypeIn.ListCount <= 0
    cmbTypeIn.RemoveItem 0
    Loop
End Sub


Public Property Get ListCount() As Integer
ListCount = lstHistory.ListCount
End Property


Public Property Get HistorySize() As Integer
HistorySize = maxHistory
End Property
Public Property Let HistorySize(ByVal newValue As Integer)
maxHistory = newValue
End Property


Public Property Get CaseSensitive() As Boolean
CaseSensitive = isCaseSensitive
End Property
Public Property Let CaseSensitive(ByVal newValue As Boolean)
isCaseSensitive = newValue
End Property


Public Property Get ToolTipText() As String
ToolTipText = myToolTipText
End Property
Public Property Let ToolTipText(ByVal newValue As String)
myToolTipText = newValue
cmbTypeIn.ToolTipText = myToolTipText
txtTypeIn.ToolTipText = myToolTipText
End Property


Public Property Get Enabled() As Boolean
Enabled = isEnabled
End Property
Public Property Let Enabled(ByVal newValue As Boolean)
isEnabled = newValue
cmbTypeIn.Enabled = isEnabled
txtTypeIn.Enabled = isEnabled
End Property


Public Property Get Text() As String
Select Case myStyle
    Case sList
        Text = cmbTypeIn.Text
    Case sText
        Text = txtTypeIn.Text
End Select
End Property
Public Property Let Text(ByVal newText As String)
cmbTypeIn.Text = newText
txtTypeIn.Text = newText
End Property


Public Property Get Style() As Styles
Style = myStyle
End Property
Public Property Let Style(ByVal newValue As Styles)
myStyle = newValue
Select Case myStyle
    Case sList
        cmbTypeIn.Visible = True
        txtTypeIn.Visible = False
    Case sText
        cmbTypeIn.Visible = False
        txtTypeIn.Visible = True
End Select
End Property


Private Sub cmbTypeIn_Change()
If Not cancelTracking Then TrackText
txtTypeIn.Text = cmbTypeIn.Text
If myStyle = sList Then RaiseEvent Change
End Sub
Private Sub cmbTypeIn_Click()
If myStyle = sList Then RaiseEvent Click
End Sub
Private Sub cmbTypeIn_DblClick()
If myStyle = sList Then RaiseEvent DblClick
End Sub
Private Sub cmbTypeIn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Or KeyCode = 8 Then
    cancelTracking = True
Else
    cancelTracking = False
    End If
If myStyle = sList Then RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub cmbTypeIn_KeyPress(KeyAscii As Integer)
If myStyle = sList Then RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub cmbTypeIn_KeyUp(KeyCode As Integer, Shift As Integer)
If myStyle = sList Then RaiseEvent KeyUp(KeyCode, Shift)
End Sub


Private Sub txtTypeIn_Change()
If Not cancelTracking Then TrackText
cmbTypeIn.Text = txtTypeIn.Text
If myStyle = sText Then RaiseEvent Change
End Sub
Private Sub txtTypeIn_Click()
If myStyle = sText Then RaiseEvent Click
End Sub
Private Sub txtTypeIn_DblClick()
If myStyle = sText Then RaiseEvent DblClick
End Sub
Private Sub txtTypeIn_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 46 Or KeyCode = 8 Then
    cancelTracking = True
Else
    cancelTracking = False
    End If
If myStyle = sText Then RaiseEvent KeyDown(KeyCode, Shift)
End Sub
Private Sub txtTypeIn_KeyPress(KeyAscii As Integer)
If myStyle = sText Then RaiseEvent KeyPress(KeyAscii)
End Sub
Private Sub txtTypeIn_KeyUp(KeyCode As Integer, Shift As Integer)
If myStyle = sText Then RaiseEvent KeyUp(KeyCode, Shift)
End Sub


Private Sub UserControl_Initialize()
cancelTracking = False
End Sub
Private Sub UserControl_InitProperties()
Me.Style = sList
Me.HistorySize = 20
Me.CaseSensitive = True
Me.Enabled = True
Me.ToolTipText = ""
Me.Text = Default_Text
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
    myStyle = .ReadProperty("Style", sList)
    Me.Style = myStyle
    maxHistory = .ReadProperty("HistorySize", 20)
    Me.HistorySize = maxHistory
    isCaseSensitive = .ReadProperty("CaseSensitive", True)
    Me.CaseSensitive = isCaseSensitive
    isEnabled = .ReadProperty("Enabled", True)
    Me.Enabled = isEnabled
    myToolTipText = .ReadProperty("ToolTipText", "")
    Me.ToolTipText = myToolTipText
    cmbTypeIn.Text = .ReadProperty("Text", Default_Text)
    txtTypeIn.Text = .ReadProperty("Text", Default_Text)
    
End With
End Sub
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "Style", myStyle, sList
    .WriteProperty "HistorySize", maxHistory, 20
    .WriteProperty "CaseSensitive", isCaseSensitive, True
    .WriteProperty "Enabled", isEnabled, True
    .WriteProperty "ToolTipText", myToolTipText, ""
    Select Case myStyle
        Case sList
            .WriteProperty "Text", cmbTypeIn.Text, Default_Text
        Case sText
            .WriteProperty "Text", txtTypeIn.Text, Default_Text
    End Select
End With
End Sub
Private Sub UserControl_Resize()
cmbTypeIn.Top = 15
cmbTypeIn.Left = 15
txtTypeIn.Top = 15
txtTypeIn.Left = 15
If Height <> 345 Then Height = 345
cmbTypeIn.Width = UserControl.Width - 30
txtTypeIn.Width = UserControl.Width - 30
End Sub

