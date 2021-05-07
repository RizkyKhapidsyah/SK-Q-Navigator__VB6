VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H80000000&
   Caption         =   "Links"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8670
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4110
   ScaleWidth      =   8670
   StartUpPosition =   3  'Windows Default
   Begin QNavigator.GoldButton Command2 
      Height          =   375
      Left            =   7320
      TabIndex        =   3
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Close"
      Alignment       =   2
      ForeColor       =   -2147483630
      SkinDisabledText=   -2147483632
      SkinHighlight   =   -2147483628
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnHover         =   5
   End
   Begin QNavigator.GoldButton Command1 
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   120
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Refresh"
      Alignment       =   2
      ForeColor       =   -2147483630
      SkinDisabledText=   -2147483632
      SkinHighlight   =   -2147483628
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      OnHover         =   5
   End
   Begin VB.ListBox List1 
      Height          =   3375
      ItemData        =   "form5.frx":0000
      Left            =   0
      List            =   "form5.frx":0002
      TabIndex        =   1
      Top             =   720
      Width           =   8655
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000007&
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5655
   End
   Begin VB.Menu Popup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu copy 
         Caption         =   "Copy"
      End
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Sub tag_click()
On Error Resume Next
If Form1.TabStrip1.SelectedItem.Key = "sud11" Then
Label1.Caption = "That page contains " & Form1.WebBrowser1.Document.links.length & " links"
For i = 0 To Form1.WebBrowser1.Document.links.length - 1
    If Left$(LCase(Form1.WebBrowser1.Document.links.Item(i)), 4) = "http" Then List1.AddItem (Form1.WebBrowser1.Document.links.Item(i))
Next i
Else: GoTo 45
End If: Exit Sub

45: If Form1.TabStrip1.SelectedItem.Key = "sud22" Then
   Label1.Caption = "That page contains " & Form1.WebBrowser2.Document.links.length & " links"
   For i = 0 To Form1.WebBrowser2.Document.links.length - 1
    If Left$(LCase(Form1.WebBrowser2.Document.links.Item(i)), 4) = "http" Then List1.AddItem (Form1.WebBrowser2.Document.links.Item(i))
Next i
Else: GoTo 55
End If: Exit Sub

55: If Form1.TabStrip1.SelectedItem.Key = "sud33" Then
Label1.Caption = "That page contains " & Form1.WebBrowser3.Document.links.length & " links"
For i = 0 To Form1.WebBrowser3.Document.links.length - 1
    If Left$(LCase(Form1.WebBrowser3.Document.links.Item(i)), 4) = "http" Then List1.AddItem (Form1.WebBrowser3.Document.links.Item(i))
Next i
Else: GoTo 65
End If: Exit Sub

65: If Form1.TabStrip1.SelectedItem.Key = "sud44" Then
Label1.Caption = "That page contains " & Form1.WebBrowser4.Document.links.length & " links"
For i = 0 To Form1.WebBrowser4.Document.links.length - 1
    If Left$(LCase(Form1.WebBrowser4.Document.links.Item(i)), 4) = "http" Then List1.AddItem (Form1.WebBrowser4.Document.links.Item(i))
Next i
Else: GoTo 75
End If: Exit Sub

75 If Form1.TabStrip1.SelectedItem.Key = "sud55" Then
Label1.Caption = "That page contains " & Form1.WebBrowser5.Document.links.length & " links"
For i = 0 To Form1.WebBrowser5.Document.links.length - 1
    If Left$(LCase(Form1.WebBrowser5.Document.links.Item(i)), 4) = "http" Then List1.AddItem (Form1.WebBrowser5.Document.links.Item(i))
Next i
Else: GoTo 85
End If: Exit Sub

85 If Form1.TabStrip1.SelectedItem.Key = "sud66" Then
Label1.Caption = "That page contains " & Form1.WebBrowser6.Document.links.length & " links"
For i = 0 To Form1.WebBrowser6.Document.links.length - 1
    If Left$(LCase(Form1.WebBrowser6.Document.links.Item(i)), 4) = "http" Then List1.AddItem (Form1.WebBrowser6.Document.links.Item(i))
Next i
Else: GoTo 95
End If: Exit Sub

95 If Form1.TabStrip1.SelectedItem.Key = "sud77" Then
Label1.Caption = "That page contains " & Form1.WebBrowser7.Document.links.length & " links"
For i = 0 To Form1.WebBrowser7.Document.links.length - 1
    If Left$(LCase(Form1.WebBrowser7.Document.links.Item(i)), 4) = "http" Then List1.AddItem (Form1.WebBrowser7.Document.links.Item(i))
Next i
Else: GoTo 99
End If: Exit Sub

99 If Form1.TabStrip1.SelectedItem.Key = "sud88" Then
Label1.Caption = "That page contains " & Form1.WebBrowser8.Document.links.length & " links"
For i = 0 To Form1.WebBrowser8.Document.links.length - 1
    If Left$(LCase(Form1.WebBrowser8.Document.links.Item(i)), 4) = "http" Then List1.AddItem (Form1.WebBrowser8.Document.links.Item(i))
Next i
Else: GoTo 100
End If: Exit Sub

100 If Form1.TabStrip1.SelectedItem.Caption = "Browser 9" Then
Label1.Caption = "That page contains " & Form1.WebBrowser9.Document.links.length & " links"
For i = 0 To Form1.WebBrowser9.Document.links.length - 1
    If Left$(LCase(Form1.WebBrowser9.Document.links.Item(i)), 4) = "http" Then List1.AddItem (Form1.WebBrowser9.Document.links.Item(i))
Next i
Else: GoTo 105
End If: Exit Sub

105 If Form1.TabStrip1.SelectedItem.Key = "sud10" Then
Label1.Caption = "That page contains " & Form1.WebBrowser10.Document.links.length & " links"
For i = 0 To Form1.WebBrowser10.Document.links.length - 1
    If Left$(LCase(Form1.WebBrowser10.Document.links.Item(i)), 4) = "http" Then List1.AddItem (Form1.WebBrowser10.Document.links.Item(i))
Next i
End If: Exit Sub
'end
End Sub
Private Sub Command1_Click()
List1.Clear

tag_click
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Form1.Text4.Text = List1.Text
End Sub

Private Sub copy_Click()
Clipboard.SetText (List1.Text)

End Sub

Private Sub Form_Load()
Me.Caption = "Q Navigtor"
Me.Icon = Form1.Icon

tag_click
End Sub

Private Sub Form_Unload(Cancel As Integer)
List1.Clear
Unload Me
End Sub

Private Sub List1_DblClick()
If Form1.TabStrip1.SelectedItem.Key = "sud11" Then
Form1.WebBrowser1.Navigate (List1.Text)
End If

If Form1.TabStrip1.SelectedItem.Key = "sud22" Then
Form1.WebBrowser2.Navigate (List1.Text)
End If

If Form1.TabStrip1.SelectedItem.Key = "sud33" Then
Form1.WebBrowser3.Navigate (List1.Text)
End If

If Form1.TabStrip1.SelectedItem.Key = "sud44" Then
Form1.WebBrowser4.Navigate (List1.Text)
End If

If Form1.TabStrip1.SelectedItem.Key = "sud55" Then
Form1.WebBrowser5.Navigate (List1.Text)
End If

If Form1.TabStrip1.SelectedItem.Key = "sud66" Then
Form1.WebBrowser6.Navigate (List1.Text)
End If

If Form1.TabStrip1.SelectedItem.Key = "sud77" Then
Form1.WebBrowser7.Navigate (List1.Text)
End If

If Form1.TabStrip1.SelectedItem.Key = "sud88" Then
Form1.WebBrowser8.Navigate (List1.Text)
End If

If Form1.TabStrip1.SelectedItem.Caption = "Browser 9" Then
Form1.WebBrowser9.Navigate (List1.Text)
End If

If Form1.TabStrip1.SelectedItem.Key = "sud10" Then
Form1.WebBrowser10.Navigate (List1.Text)
End If
End Sub

Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.Popup

End Sub

