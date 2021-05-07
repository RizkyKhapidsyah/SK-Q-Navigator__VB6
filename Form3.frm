VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H8000000B&
   Caption         =   "Form3"
   ClientHeight    =   1425
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form3"
   ScaleHeight     =   1425
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Only Recover the Browser"
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   720
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start Debug Program"
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   720
      Width           =   1095
   End
   Begin VB.PictureBox Picture9 
      Height          =   255
      Left            =   120
      ScaleHeight     =   195
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   120
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   4455
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim tenth As Long


#If Win32 Then
Private Declare Function BitBlt Lib "gdi32" _
(ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, _
ByVal nWidth As Long, ByVal nHeight As Long, _
ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, _
ByVal dwRop As Long) As Long
#Else
Private Declare Function BitBlt Lib "GDI" (ByVal hDestDC As _
Integer, ByVal X As Integer, ByVal Y As Integer, ByVal nWidth _
As Integer, ByVal nHeight As Integer, ByVal hSrcDC As Integer, _
ByVal xSrc As Integer, ByVal ySrc As Integer, ByVal dwRop As _
Long) As Integer
#End If

Sub UpdateStatus(FileBytes As Long)
'--------------------------------------------------------------------
' Update the Picture9 status bar
'--------------------------------------------------------------------
    Static progress As Long
    Dim r As Long
    Const SRCCOPY = &HCC0020
    Dim Txt$
    progress = progress + FileBytes
    If progress > Picture9.ScaleWidth Then
        progress = Picture9.ScaleWidth
    End If
    Txt$ = Format$(CLng((progress / Picture9.ScaleWidth) * 100)) + "%"
    Picture9.Cls
    Picture9.CurrentX = _
    (Picture9.ScaleWidth - Picture9.TextWidth(Txt$)) \ 2
    Picture9.CurrentY = _
    (Picture9.ScaleHeight - Picture9.TextHeight(Txt$)) \ 2
    Picture9.Print Txt$
    Picture9.Line (0, 0)-(progress, Picture9.ScaleHeight), _
    Picture9.ForeColor, BF
    r = BitBlt(Picture9.hdc, 0, 0, Picture9.ScaleWidth, _
        Picture9.ScaleHeight, Picture9.hdc, 0, 0, SRCCOPY)
End Sub

Private Sub Command4_Click()
End Sub

Private Sub Command1_Click()
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Label1.Caption = "Verifying files..."
Me.ScaleHeight = 830
Open "sud.dat" For Output As #1
Write #1, "NEWS", "www.news.com"
Write #1, "DOWNLOAD", "www.download.com"
Write #1, "MAIL1", "www.mail.com"
Write #1, "MAIL2", "mail.yahoo.com"
Close #1
Open "history.txt" For Output As #1
Close #1
Form1.Refresh
On Error GoTo err
Open "prev.dat" For Input As #1
Input #1, previnst
Input #1, previnst
Input #1, previnst
Input #1, previnst
Input #1, previnst
Input #1, previnst
Input #1, previnst
Input #1, previnst
Input #1, previnst
Input #1, previnst
Close #1
GoTo wer
err: Close
Open "prev.ini" For Output As #1
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Close #1

wer:
    Picture9.FontBold = True
    Picture9.AutoRedraw = True
    Picture9.BackColor = vbWhite
    Picture9.DrawMode = 10
    Picture9.FillStyle = 0
    Picture9.ForeColor = vbBlue
    Picture9.ScaleWidth = 109
    tenth = 10
    For i = 1 To 11
        Call UpdateStatus(tenth)
        X = Timer
        While Timer < X + 0.75
            DoEvents
        Wend
    Next
    MsgBox ("Debug Task Done. You may have to re-configure Q Navigator according to your needs by going to TOOLS >OPTIONS >")
    
    If MsgBox("Recover and Reset Browser To Previous Web Addresses ?", vbOKCancel) = vbOK Then
    Command3_Click
    End If
Unload Me

End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
On Error GoTo err
Open "prev.ini" For Input As #1
Input #1, previnst
Input #1, previnst
Input #1, previnst
Input #1, previnst
Input #1, previnst
Input #1, previnst
Input #1, previnst
Input #1, previnst
Input #1, previnst
Input #1, previnst
Close #1
GoTo wer
err: Close
Open "prev.ini" For Output As #1
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Write #1, "about:blank"
Close #1

wer:
    Open "prev.ini" For Input As #1
    Input #1, previnst
    If previnst = "about:blank" Then GoTo 21
    On Error GoTo 20
    Form1.TabStrip1.Tabs.Add(2, "sud11", "Browser 1") = "Browser 1"
20: Form1.WebBrowser1.Navigate (previnst)
21: Input #1, previnst
    If previnst = "about:blank" Then GoTo 31
    On Error GoTo 30
    Form1.TabStrip1.Tabs.Add(3, "sud22", "Browser 2") = "Browser 2"
30: Form1.WebBrowser2.Navigate (previnst)
31: Input #1, previnst
    If previnst = "about:blank" Then GoTo 41
    On Error GoTo 40
    Form1.TabStrip1.Tabs.Add(4, "sud33", "Browser 3") = "Browser 3"
40: Form1.WebBrowser3.Navigate (previnst)
41: Input #1, previnst
    If previnst = "about:blank" Then GoTo 51
    On Error GoTo 50
    Form1.TabStrip1.Tabs.Add(5, "sud44", "Browser 4") = "Browser 4"
50: Form1.WebBrowser4.Navigate (previnst)
51: Input #1, previnst
    If previnst = "about:blank" Then GoTo 61
    On Error GoTo 60
    Form1.TabStrip1.Tabs.Add(6, "sud55", "Browser 5") = "Browser 5"
60: Form1.WebBrowser5.Navigate (previnst)
61: Input #1, previnst
    If previnst = "about:blank" Then GoTo 71
    On Error GoTo 70
    Form1.TabStrip1.Tabs.Add(7, "sud66", "Browser 6") = "Browser 6"
70: Form1.WebBrowser6.Navigate (previnst)
71: Input #1, previnst
    If previnst = "about:blank" Then GoTo 81
    On Error GoTo 80
    Form1.TabStrip1.Tabs.Add(8, "sud77", "Browser 7") = "Browser 7"
80: Form1.WebBrowser7.Navigate (previnst)
81: Input #1, previnst
    If previnst = "about:blank" Then GoTo 91
    On Error GoTo 90
    Form1.TabStrip1.Tabs.Add(9, "sud88", "Browser 8") = "Browser 8"
90: Form1.WebBrowser8.Navigate (previnst)
91: Input #1, previnst
    If previnst = "about:blank" Then GoTo 101
    On Error GoTo 100
    Form1.TabStrip1.Tabs.Add(10, "sud99", "Browser 9") = "Browser 9"
100: Form1.WebBrowser9.Navigate (previnst)
101: Input #1, previnst
    If previnst = "about:blank" Then GoTo 111
    On Error GoTo 110
    Form1.TabStrip1.Tabs.Add(11, "sud10", "Browser 10") = "Browser 10"
110: Form1.WebBrowser10.Navigate (previnst)
111:
    Close #1
    
Unload Me

End Sub

Private Sub Form_Load()
Me.Icon = Form1.Icon
Me.Caption = "Self Debugger"
End Sub

