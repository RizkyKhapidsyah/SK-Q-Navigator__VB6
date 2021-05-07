VERSION 5.00
Begin VB.Form frmSearch 
   BackColor       =   &H8000000B&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Search"
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4890
   Icon            =   "frmSearch.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1545
   ScaleWidth      =   4890
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin QNavigator.ctlAutoType Combo1 
      Height          =   345
      Left            =   1440
      TabIndex        =   7
      Top             =   600
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   609
      CaseSensitive   =   0   'False
      Text            =   ""
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Command2"
      Height          =   195
      Left            =   0
      TabIndex        =   6
      Top             =   960
      Visible         =   0   'False
      Width           =   375
   End
   Begin QNavigator.GoldButton cmdnewsearch 
      Height          =   375
      Left            =   120
      TabIndex        =   5
      Top             =   1080
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "New Search"
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
   Begin QNavigator.GoldButton cmdSearch 
      Height          =   375
      Left            =   1440
      TabIndex        =   4
      Top             =   1080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      Caption         =   "SEARCH"
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
      Left            =   3720
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   661
      Caption         =   "Cancel"
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
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1440
      TabIndex        =   0
      Top             =   120
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000A&
      Caption         =   "       In"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000004&
      Caption         =   "Search for :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ar$

Private Sub cmdNewSearch_Click()
Text1.Text = ""
Combo1.Text = "Google"
End Sub

Private Sub cmdSearch_Click()

If Combo1.Text = "Select a Search Engine" Then
MsgBox "Please select a search engine to search", vbInformation, "Select Engine"
End If
If Text1.Text = "" Then
MsgBox "Please enter at least 1 word to search for", vbInformation, "Enter Word"
GoTo skip
End If
Select Case Combo1.Text
Case "Google"
ar$ = "http://www.google.com/search?q=" & Text1.Text & "&meta=lr%3D%26hl%3Den&btnG=Google+Search"
Case "Infoseek"
ar$ = "http://infoseek.go.com/Titles?col=WW&qt=%22" & Text1.Text & "%22&sv=IS&lk=noframes&svx=sbox_top&cc=WW&oq=" & Text1.Text
Case "Yahoo"
ar$ = "http://ink.yahoo.com/bin/query?p=" & Text1.Text & "&z=2&hc=0&hs=0"
Case "Altavista"
ar$ = "http://www.altavista.com/cgi-bin/query?pg=q&kl=XX&stype=stext&q=" & Text1.Text
Case "Lycos"
ar$ = "http://www.lycos.com/srch/?lpv=1&loc=searchhp&query=" & Text1.Text
Case "About.COM"
ar$ = "http://search.about.com/fullsearch.htm?terms=" & "&PM=59_0100_S&Action.x=9&Action.y=7 "
Case "MSN"
ar$ = "http://search.msn.com/results.asp?RS=CHECKED&FORM=MSNH&v=1&q=" & Text1.Text
Case "Excite"
ar$ = "http://search.excite.com/search.gw?search=" & Text1.Text
Case "Go To"
ar$ = "http://www.goto.com/d/search/;$sessionid$X5Y4C5AABJ3OTQFIEF1QPUQ?type=home&Keywords=" & Text1.Text & "Find+It%21.x=15&Find+It%21.y=26"
Case "Looksmart"
ar$ = "http://www.looksmart.com/r_search?look=&pin=000602x37c8f53a35211910451&key=" & Text1.Text
Case "Snap"
ar$ = "http://www.snap.com/search/directory/results/1,61,-0,00.html?tag=st.v2.fdsb.1&keyword=" & Text1.Text
End Select
If Form1.WebBrowser1.LocationURL = "about:blank" Or Form1.WebBrowser1.LocationURL = "" Then
On Error GoTo yel1
Form1.TabStrip1.Tabs.Add(2, "sud11", "Browser 1") = "Browser 1"
yel1:
Form1.WebBrowser1.Navigate (ar$)
GoTo 66
End If
If Form1.WebBrowser2.LocationURL = "about:blank" Or Form1.WebBrowser2.LocationURL = "" Then
On Error GoTo yel2
Form1.TabStrip1.Tabs.Add(3, "sud22", "Browser 2") = "Browser 2"
yel2:
Form1.WebBrowser2.Navigate (ar$)
GoTo 66
End If
If Form1.WebBrowser3.LocationURL = "about:blank" Or Form1.WebBrowser3.LocationURL = "" Then
On Error GoTo yel3
Form1.TabStrip1.Tabs.Add(4, "sud33", "Browser 3") = "Browser 3"
yel3:
Form1.WebBrowser3.Navigate (ar$)
GoTo 66
End If

If Form1.WebBrowser4.LocationURL = "about:blank" Or Form1.WebBrowser4.LocationURL = "" Then
On Error GoTo yel4
Form1.TabStrip1.Tabs.Add(5, "sud44", "Browser 44") = "Browser 4"
yel4:
Form1.WebBrowser4.Navigate (ar$)
GoTo 66
End If
If Form1.WebBrowser5.LocationURL = "about:blank" Or Form1.WebBrowser5.LocationURL = "" Then
On Error GoTo yel5
Form1.TabStrip1.Tabs.Add(6, "sud55", "Browser 55") = "Browser 5"
yel5:
Form1.WebBrowser5.Navigate (ar$)
GoTo 66
End If
If Form1.WebBrowser6.LocationURL = "about:blank" Or Form1.WebBrowser6.LocationURL = "" Then
On Error GoTo yel6
Form1.TabStrip1.Tabs.Add(7, "sud66", "Browser 6") = "Browser 6"
yel6:
Form1.WebBrowser6.Navigate (ar$)
GoTo 66
End If
If Form1.WebBrowser7.LocationURL = "about:blank" Or Form1.WebBrowser7.LocationURL = "" Then
On Error GoTo yel7
Form1.TabStrip1.Tabs.Add(8, "sud77", "Browser 7") = "Browser 7"
yel7:
Form1.WebBrowser7.Navigate (ar$)
GoTo 66
End If
If Form1.WebBrowser8.LocationURL = "about:blank" Or Form1.WebBrowser8.LocationURL = "" Then
On Error GoTo yel8
Form1.TabStrip1.Tabs.Add(9, "sud88", "Browser 8") = "Browser 8"
yel8:
Form1.WebBrowser8.Navigate (ar$)
GoTo 66
End If
If Form1.WebBrowser9.LocationURL = "about:blank" Or Form1.WebBrowser9.LocationURL = "" Then
On Error GoTo yel9
Form1.TabStrip1.Tabs.Add(10, "sud99", "Browser 9") = "Browser 9"
yel9:
Form1.WebBrowser9.Navigate (ar$)
GoTo 66
End If
If Form1.WebBrowser10.LocationURL = "about:blank" Or Form1.WebBrowser10.LocationURL = "" Then
On Error GoTo yel10
Form1.TabStrip1.Tabs.Add(11, "sud10", "Browser 10") = "Browser 10"
yel10:
Form1.WebBrowser10.Navigate (ar$)
GoTo 66
End If
MsgBox ("All Browsers are in use. To close an active unwanted page click 'CLEAR' above.")
66:
skip:
Me.Hide
End Sub

Private Sub Command1_Click()
Me.Hide
End Sub

Private Sub Command2_Click()
cmdSearch_Click
End Sub

Private Sub Form_Load()
frmSearch.Icon = Form1.Icon


Combo1.AddItem "Google"
Combo1.AddItem "Yahoo"
Combo1.AddItem "Infoseek"
Combo1.AddItem "Lycos"
Combo1.AddItem "Altavista"
Combo1.AddItem "MSN"
Combo1.AddItem "Go To"
Combo1.AddItem "Looksmart"
Combo1.AddItem "Snap"
Combo1.AddItem "Mamma"
Combo1.AddItem "Dogpile"
Combo1.AddItem "About.COM"
Combo1.AddItem "Web Crawler"
Combo1.AddItem "Netscape"
Combo1.AddItem "Excite"
End Sub

Private Sub Frame1_DragDrop(source As Control, X As Single, Y As Single)

End Sub

Private Sub Picture1_Click()
cmdSearch_Click
End Sub

Private Sub Picture2_Click()
Command1_Click
End Sub

Private Sub Search_Click()

End Sub
