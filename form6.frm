VERSION 5.00
Begin VB.Form Form6 
   Caption         =   "Source"
   ClientHeight    =   5400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8295
   LinkTopic       =   "Form6"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5400
   ScaleWidth      =   8295
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      BackColor       =   &H80000009&
      ForeColor       =   &H80000007&
      Height          =   5415
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   0
      Width           =   8295
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Icon = Form1.Icon
On Error Resume Next
    If Form1.TabStrip1.SelectedItem.Key = "sud11" Then Me.Caption = Form1.WebBrowser1.LocationURL: Text1.Text = Form1.WebBrowser1.Document.documentElement.innerHTML
    If Form1.TabStrip1.SelectedItem.Key = "sud22" Then Me.Caption = Form1.WebBrowser2.LocationURL: Text1.Text = Form1.WebBrowser2.Document.documentElement.innerHTML
    If Form1.TabStrip1.SelectedItem.Key = "sud33" Then Me.Caption = Form1.WebBrowser3.LocationURL: Text1.Text = Form1.WebBrowser3.Document.documentElement.innerHTML
    If Form1.TabStrip1.SelectedItem.Key = "sud44" Then Me.Caption = Form1.WebBrowser4.LocationURL: Text1.Text = Form1.WebBrowser4.Document.documentElement.innerHTML
    If Form1.TabStrip1.SelectedItem.Key = "sud55" Then Me.Caption = Form1.WebBrowser5.LocationURL: Text1.Text = Form1.WebBrowser5.Document.documentElement.innerHTML
    If Form1.TabStrip1.SelectedItem.Key = "sud66" Then Me.Caption = Form1.WebBrowser6.LocationURL: Text1.Text = Form1.WebBrowser6.Document.documentElement.innerHTML
    If Form1.TabStrip1.SelectedItem.Key = "sud77" Then Me.Caption = Form1.WebBrowser7.LocationURL: Text1.Text = Form1.WebBrowser7.Document.documentElement.innerHTML
    If Form1.TabStrip1.SelectedItem.Key = "sud88" Then Me.Caption = Form1.WebBrowser8.LocationURL: Text1.Text = Form1.WebBrowser8.Document.documentElement.innerHTML
    If Form1.TabStrip1.SelectedItem.Key = "sud99" Then Me.Caption = Form1.WebBrowser9.LocationURL: Text1.Text = Form1.WebBrowser9.Document.documentElement.innerHTML
    If Form1.TabStrip1.SelectedItem.Key = "sud10" Then Me.Caption = Form1.WebBrowser10.LocationURL: Text1.Text = Form1.WebBrowser10.Document.documentElement.innerHTML
     
End Sub

