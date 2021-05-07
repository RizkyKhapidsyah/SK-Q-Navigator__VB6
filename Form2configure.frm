VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H8000000B&
   Caption         =   "Form2"
   ClientHeight    =   3015
   ClientLeft      =   3660
   ClientTop       =   2520
   ClientWidth     =   5865
   ForeColor       =   &H8000000E&
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   5865
   ShowInTaskbar   =   0   'False
   Begin QNavigator.GoldButton GoldButton2 
      Height          =   375
      Left            =   4800
      TabIndex        =   11
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
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
   Begin QNavigator.GoldButton GoldButton1 
      Height          =   375
      Left            =   3720
      TabIndex        =   10
      Top             =   2520
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      Caption         =   "O K"
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
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   1320
      TabIndex        =   9
      Top             =   2040
      Width           =   4455
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   1320
      TabIndex        =   8
      Top             =   1560
      Width           =   4455
   End
   Begin VB.TextBox Text4 
      Height          =   315
      Left            =   1320
      TabIndex        =   7
      Top             =   1080
      Width           =   4455
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   1320
      TabIndex        =   6
      Top             =   600
      Width           =   4455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "WEB MAIL  2"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "WEB MAIL 1"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1560
      Width           =   975
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "HOMEPAGE"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   975
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "DOWNLOAD"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
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
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "NEWS:"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim s
Private Sub Command1_Click()
End Sub

Private Sub Label8_Click()

End Sub


Private Sub Form_Load()
Form2.Icon = Form1.Icon
s = 1
On Error GoTo r
Open "home.dat" For Input As #1
Input #1, home
Close #1
Text4.Text = home
r:
Open "sud.dat" For Input As #1
While Not EOF(1)
Input #1, fun, webadd
If fun = "NEWS" Then Text1.Text = webadd
If fun = "DOWNLOAD" Then Text3.Text = webadd
If fun = "HOMEPAGE" Then Text4.Text = webadd
If fun = "MAIL1" Then Text5.Text = webadd
If fun = "MAIL2" Then Text6.Text = webadd
Wend
Close #1
s = 0

End Sub

Private Sub Label3_Click()

End Sub

Private Sub Picture1_Click()
Unload Me
End Sub

Private Sub Picture2_Click()
Unload Me
End Sub

Private Sub GoldButton1_Click()
Unload Me
End Sub

Private Sub GoldButton2_Click()
Unload Me
End Sub

Private Sub Text1_Change()
If s = 0 Then
Open "sud.dat" For Input As #1
Open "temp.dat" For Output As #2
While Not EOF(1)
Input #1, fun, webadd
If fun = "NEWS" Then Write #2, "NEWS", Text1.Text
If fun <> "NEWS" Then Write #2, fun, webadd
Wend
Close #1, #2
Kill "sud.dat"
Name "temp.dat" As "sud.dat"
End If
End Sub


Private Sub Text3_Change()
If s = 0 Then
Open "sud.dat" For Input As #1
Open "temp.dat" For Output As #2
While Not EOF(1)
Input #1, fun, webadd
If fun = "DOWNLOAD" Then Write #2, "DOWNLOAD", Text3.Text
If fun <> "DOWNLOAD" Then Write #2, fun, webadd
Wend
Close #1, #2
Kill "sud.dat"
Name "temp.dat" As "sud.dat"
End If
End Sub

Private Sub Text4_Change()
If s = 0 Then
Open "home.dat" For Output As #1
Write #1, Text4.Text
Close #1
End If
End Sub

Private Sub Text5_Change()
If s = 0 Then
Open "sud.dat" For Input As #1
Open "temp.dat" For Output As #2
While Not EOF(1)
Input #1, fun, webadd
If fun = "MAIL1" Then Write #2, "MAIL1", Text5.Text
If fun <> "MAIL1" Then Write #2, fun, webadd
Wend
Close #1, #2
Kill "sud.dat"
Name "temp.dat" As "sud.dat"
End If
End Sub

Private Sub Text6_Change()
If s = 0 Then
Open "sud.dat" For Input As #1
Open "temp.dat" For Output As #2
While Not EOF(1)
Input #1, fun, webadd
If fun = "MAIL2" Then Write #2, "MAIL2", Text6.Text
If fun <> "MAIL2" Then Write #2, fun, webadd
Wend
Close #1, #2
Kill "sud.dat"
Name "temp.dat" As "sud.dat"
End If
End Sub
