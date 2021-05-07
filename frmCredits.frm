VERSION 5.00
Begin VB.Form frmCredits 
   BackColor       =   &H00000000&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About - Q Navigator"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3180
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCredits.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   495
      Left            =   3720
      TabIndex        =   6
      Top             =   1560
      Width           =   615
   End
   Begin VB.TextBox Text1 
      Height          =   975
      Left            =   3480
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "frmCredits.frx":000C
      Top             =   600
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox txtScroll 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   0
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      MultiLine       =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   3255
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      TabIndex        =   2
      Top             =   4200
      Width           =   375
   End
   Begin VB.CommandButton cmdCredits 
      Caption         =   "Command1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   1
      Top             =   4200
      Width           =   375
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   0
      Top             =   0
   End
   Begin VB.PictureBox picScroll 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3855
      Left            =   0
      ScaleHeight     =   3855
      ScaleWidth      =   3255
      TabIndex        =   0
      Top             =   -360
      Width           =   3255
   End
   Begin VB.Label Label1 
      Height          =   735
      Left            =   3600
      TabIndex        =   5
      Top             =   2160
      Visible         =   0   'False
      Width           =   1215
   End
End
Attribute VB_Name = "frmCredits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    frmCredits.Icon = Form1.Icon
    Left = (Screen.Width - Width) \ 2
    Top = (Screen.Height - Height) \ 2
'Variable Declarations
    Dim iFileNum As Integer
    Dim lLineCount As Long
    Dim lLineHeight As Long
    
    On Error GoTo ErrHandler 'Goto to ErrHandler if an error occurs
    
    If cmdCredits.Caption = "Hide Credits" Then
        picScroll.Visible = False
        tmrScroll.Enabled = False
        cmdCredits.Caption = "&Roll Credits"
    Else
        iFileNum = FreeFile
        'open file and read text from it
        
        txtScroll = Text1.Text
        lLineCount = SendMessage(txtScroll.hwnd, EM_GETLINECOUNT, 0&, 0&)
        lLineHeight = TextHeight("TEST") 'Get the height of text in file
        txtScroll.Height = lLineHeight * lLineCount
        picScroll.Left = 0
        picScroll.Visible = True
        tmrScroll.Enabled = True
        cmdCredits.Caption = "Hide Credits"
    End If
    Exit Sub

ErrHandler:
    txtScroll.Text = "File Not Found !!!" & vbNewLine & "The Required file is missing"
    Resume Next
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbWhite
End Sub


Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = &HC0FFFF
End Sub

Private Sub tmrScroll_Timer()
    'scroll txtScroll
    If txtScroll.Top + txtScroll.Height < picScroll.Top Then 'picScroll.Top
        txtScroll.Top = picScroll.Height
    Else
        txtScroll.Top = txtScroll.Top - 25
    End If
End Sub

Private Sub txtScroll_GotFocus()
    Command1.SetFocus
    'Don't let the text box get focus, althought the text
    'box is locked it looks bad to see a cursor in the
    'text box as it scrolls up
    
    
End Sub

