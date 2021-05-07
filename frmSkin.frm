VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Q Navigator"
   ClientHeight    =   7965
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   11880
   Icon            =   "frmSkin.frx":0000
   LinkTopic       =   "Form1"
   MouseIcon       =   "frmSkin.frx":030A
   ScaleHeight     =   7965
   ScaleWidth      =   11880
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ProgressBar ProgressBar1 
      Height          =   255
      Left            =   1320
      TabIndex        =   45
      Top             =   7680
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComDlg.CommonDialog CommonDialog2 
      Left            =   11160
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      DialogTitle     =   "Open local html file ..."
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Download"
      Height          =   220
      Left            =   5760
      TabIndex        =   40
      Top             =   275
      Width           =   1095
   End
   Begin MSComctlLib.ProgressBar ProgressBar10 
      Height          =   255
      Left            =   6120
      TabIndex        =   24
      Top             =   7680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar9 
      Height          =   255
      Left            =   6720
      TabIndex        =   23
      Top             =   7560
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar8 
      Height          =   255
      Left            =   6960
      TabIndex        =   22
      Top             =   7320
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar7 
      Height          =   255
      Left            =   6360
      TabIndex        =   21
      Top             =   7080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar6 
      Height          =   255
      Left            =   4800
      TabIndex        =   20
      Top             =   7320
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar5 
      Height          =   255
      Left            =   5400
      TabIndex        =   19
      Top             =   7200
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar4 
      Height          =   255
      Left            =   5280
      TabIndex        =   18
      Top             =   7080
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox picScroll 
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000013&
      Height          =   855
      Left            =   9480
      ScaleHeight     =   855
      ScaleWidth      =   1680
      TabIndex        =   28
      Top             =   0
      Width           =   1680
      Begin VB.Timer Timer1 
         Interval        =   20
         Left            =   0
         Top             =   120
      End
      Begin VB.TextBox Text3 
         Alignment       =   2  'Center
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   0
         Locked          =   -1  'True
         MousePointer    =   1  'Arrow
         MultiLine       =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   240
         Width           =   1695
      End
   End
   Begin QNavigator.ctlAutoType cboURL 
      Height          =   345
      Left            =   960
      TabIndex        =   35
      Top             =   480
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   609
      HistorySize     =   100
      CaseSensitive   =   0   'False
      Text            =   ""
   End
   Begin MSComctlLib.TreeView treeFavorites 
      Height          =   6615
      Left            =   0
      TabIndex        =   34
      Top             =   1080
      Visible         =   0   'False
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   11668
      _Version        =   393217
      Indentation     =   212
      LabelEdit       =   1
      Style           =   7
      FullRowSelect   =   -1  'True
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "imgTreeFavorites"
      Appearance      =   1
      MouseIcon       =   "frmSkin.frx":0614
   End
   Begin MSComctlLib.TreeView treeHistory 
      Height          =   6615
      Left            =   0
      TabIndex        =   31
      Top             =   1080
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   11668
      _Version        =   393217
      Style           =   1
      HotTracking     =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Timer tmrScroll 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   8280
      Top             =   600
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   9600
      Top             =   720
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      Height          =   1425
      Left            =   3600
      TabIndex        =   1
      Top             =   8520
      Width           =   4335
   End
   Begin MSComctlLib.ProgressBar ProgressBar2 
      Height          =   255
      Left            =   5400
      TabIndex        =   16
      Top             =   7680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ProgressBar ProgressBar3 
      Height          =   255
      Left            =   5400
      TabIndex        =   17
      Top             =   7680
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   6855
      Left            =   0
      Picture         =   "frmSkin.frx":092E
      ScaleHeight     =   6855
      ScaleWidth      =   11880
      TabIndex        =   0
      Top             =   840
      Width           =   11880
      Begin VB.PictureBox Picture1 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   0
         ScaleHeight     =   225
         ScaleWidth      =   1950
         TabIndex        =   32
         Top             =   0
         Visible         =   0   'False
         Width           =   1980
         Begin VB.CommandButton Command4 
            Caption         =   "X"
            Height          =   255
            Left            =   1715
            TabIndex        =   33
            Top             =   0
            Visible         =   0   'False
            Width           =   255
         End
      End
      Begin SHDocVwCtl.WebBrowser upn 
         Height          =   255
         Left            =   4440
         TabIndex        =   27
         Top             =   480
         Visible         =   0   'False
         Width           =   1455
         ExtentX         =   2566
         ExtentY         =   450
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin SHDocVwCtl.WebBrowser update 
         Height          =   255
         Left            =   3240
         TabIndex        =   26
         Top             =   480
         Visible         =   0   'False
         Width           =   615
         ExtentX         =   1085
         ExtentY         =   450
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin VB.TextBox Text4 
         Height          =   285
         Left            =   3600
         TabIndex        =   25
         Text            =   "Text4"
         Top             =   4560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   6960
         TabIndex        =   14
         Text            =   "None"
         Top             =   5520
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.TextBox Text1 
         Height          =   615
         HideSelection   =   0   'False
         Left            =   8520
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   240
         Visible         =   0   'False
         Width           =   3255
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser10 
         Height          =   855
         Left            =   7800
         TabIndex        =   12
         Top             =   2400
         Width           =   975
         ExtentX         =   1720
         ExtentY         =   1508
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser9 
         Height          =   855
         Left            =   7440
         TabIndex        =   11
         Top             =   2400
         Width           =   975
         ExtentX         =   1720
         ExtentY         =   1508
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser8 
         Height          =   855
         Left            =   5400
         TabIndex        =   10
         Top             =   2400
         Width           =   615
         ExtentX         =   1085
         ExtentY         =   1508
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser7 
         Height          =   735
         Left            =   4080
         TabIndex        =   9
         Top             =   2400
         Width           =   975
         ExtentX         =   1720
         ExtentY         =   1296
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser6 
         Height          =   615
         Left            =   3000
         TabIndex        =   8
         Top             =   2400
         Width           =   855
         ExtentX         =   1508
         ExtentY         =   1085
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser5 
         Height          =   615
         Left            =   1800
         TabIndex        =   7
         Top             =   2400
         Width           =   975
         ExtentX         =   1720
         ExtentY         =   1085
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser1 
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   1800
         Width           =   11775
         ExtentX         =   20770
         ExtentY         =   661
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   1
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser2 
         Height          =   375
         Left            =   240
         TabIndex        =   3
         Top             =   1320
         Width           =   11775
         ExtentX         =   20770
         ExtentY         =   661
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser3 
         Height          =   375
         Left            =   240
         TabIndex        =   4
         Top             =   840
         Width           =   10095
         ExtentX         =   17806
         ExtentY         =   661
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin SHDocVwCtl.WebBrowser WebBrowser4 
         Height          =   615
         Left            =   240
         TabIndex        =   5
         Top             =   2400
         Width           =   1335
         ExtentX         =   2355
         ExtentY         =   1085
         ViewMode        =   0
         Offline         =   0
         Silent          =   0
         RegisterAsBrowser=   0
         RegisterAsDropTarget=   1
         AutoArrange     =   0   'False
         NoClientEdge    =   0   'False
         AlignLeft       =   0   'False
         NoWebView       =   0   'False
         HideFileNames   =   0   'False
         SingleClick     =   0   'False
         SingleSelection =   0   'False
         NoFolders       =   0   'False
         Transparent     =   0   'False
         ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
         Location        =   ""
      End
      Begin MSComctlLib.TabStrip TabStrip1 
         Height          =   6855
         Left            =   0
         TabIndex        =   6
         Top             =   0
         Width           =   12015
         _ExtentX        =   21193
         _ExtentY        =   12091
         Style           =   2
         HotTracking     =   -1  'True
         TabMinWidth     =   998
         _Version        =   393216
         BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
            NumTabs         =   2
            BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Q - Pad"
               Key             =   "nt"
               ImageVarType    =   2
            EndProperty
            BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
               Caption         =   "Browser 1"
               Key             =   "sud11"
               Object.Tag             =   "sud1"
               ImageVarType    =   2
            EndProperty
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Times New Roman"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.Label Label2 
         Height          =   255
         Left            =   4920
         TabIndex        =   38
         Top             =   6480
         Width           =   2535
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   9840
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   18
      ImageHeight     =   17
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":4212
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":5452
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":6526
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":693A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTreeFavorites 
      Left            =   9960
      Top             =   600
      _ExtentX        =   794
      _ExtentY        =   794
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":7A6E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":7EC0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":8312
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":88A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":8CF6
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":9448
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":9762
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":9EB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":AFE8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SHDocVwCtl.WebBrowser News 
      Height          =   255
      Left            =   10800
      TabIndex        =   30
      Top             =   7320
      Width           =   615
      ExtentX         =   1085
      ExtentY         =   450
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   ""
   End
   Begin MSComctlLib.ImageList Buttons 
      Left            =   4560
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   15
      ImageHeight     =   13
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":B436
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":B6FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":BB4A
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":BE0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":C2DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":C7DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":CCE2
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":D16E
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":D59E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":DAA2
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":DF6A
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":E08A
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSkin.frx":E13E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Height          =   495
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   9360
      _ExtentX        =   16510
      _ExtentY        =   873
      ButtonWidth     =   1349
      ButtonHeight    =   873
      Appearance      =   1
      Style           =   1
      ImageList       =   "Buttons"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   10
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Back"
            Key             =   "back"
            Description     =   "back"
            Object.ToolTipText     =   "Back"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Forward"
            Key             =   "forward"
            Description     =   "forward"
            Object.ToolTipText     =   "&Forward"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Stop"
            Key             =   "stop"
            Description     =   "stop"
            Object.ToolTipText     =   "&Stop"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Refresh"
            Key             =   "refresh"
            Description     =   "refresh"
            Object.ToolTipText     =   "&Refresh"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "&Home"
            Key             =   "home"
            Description     =   "home"
            Object.ToolTipText     =   "&Home"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Search"
            Key             =   "search"
            Description     =   "search"
            Object.ToolTipText     =   "&Search"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Links"
            Key             =   "links"
            Description     =   "links"
            Object.ToolTipText     =   "&Links"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
      MousePointer    =   1
      Begin VB.CommandButton Command9 
         Caption         =   "Mail 2"
         Height          =   220
         Left            =   8160
         TabIndex        =   44
         Top             =   245
         Width           =   1095
      End
      Begin VB.CommandButton GoldButton1 
         Caption         =   "Clear"
         Height          =   210
         Left            =   6960
         TabIndex        =   43
         Top             =   255
         Width           =   1095
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Mail 1"
         Height          =   210
         Left            =   8160
         TabIndex        =   42
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Search"
         Height          =   220
         Left            =   6960
         TabIndex        =   41
         Top             =   0
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "News"
         Height          =   220
         Left            =   5760
         TabIndex        =   39
         Top             =   0
         Width           =   1095
      End
   End
   Begin MSComctlLib.StatusBar StatusBar 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   37
      Top             =   7710
      Width           =   11880
      _ExtentX        =   20955
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address :"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   600
      Width           =   855
   End
   Begin VB.Menu fil 
      Caption         =   "&File"
      HelpContextID   =   3
      Index           =   1
      Begin VB.Menu n 
         Caption         =   "&New"
         Index           =   1
         Begin VB.Menu sdb 
            Caption         =   "Window"
            Shortcut        =   ^N
         End
         Begin VB.Menu winn 
            Caption         =   "Browser"
            Shortcut        =   ^M
         End
      End
      Begin VB.Menu mnuOpen 
         Caption         =   "&Open"
         Index           =   2
         Shortcut        =   ^O
      End
      Begin VB.Menu mnusave 
         Caption         =   "Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu spas 
         Caption         =   "-"
      End
      Begin VB.Menu pstp 
         Caption         =   "Page Setup"
      End
      Begin VB.Menu mnuprint 
         Caption         =   "Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu ptp 
         Caption         =   "-"
      End
      Begin VB.Menu pag 
         Caption         =   "&Page"
         Begin VB.Menu back 
            Caption         =   "&Back"
         End
         Begin VB.Menu ford 
            Caption         =   "&Forward"
         End
         Begin VB.Menu refr 
            Caption         =   "&Refresh"
            Shortcut        =   {F5}
         End
         Begin VB.Menu stp 
            Caption         =   "&Stop"
         End
         Begin VB.Menu home 
            Caption         =   "&Home"
         End
      End
      Begin VB.Menu woff 
         Caption         =   "Work Offline"
      End
      Begin VB.Menu ofg 
         Caption         =   "-"
      End
      Begin VB.Menu Close 
         Caption         =   "Close"
         Shortcut        =   ^W
      End
      Begin VB.Menu ext 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu ed 
      Caption         =   "&Edit"
      Begin VB.Menu cut 
         Caption         =   "Cu&t"
         Shortcut        =   ^X
      End
      Begin VB.Menu cp 
         Caption         =   "&Copy"
         Shortcut        =   ^C
      End
      Begin VB.Menu pt 
         Caption         =   "&Paste"
         Shortcut        =   ^V
      End
      Begin VB.Menu satl 
         Caption         =   "-"
      End
      Begin VB.Menu sall 
         Caption         =   "Select All"
         Shortcut        =   ^A
      End
      Begin VB.Menu fg 
         Caption         =   "-"
      End
      Begin VB.Menu fn 
         Caption         =   "Find on this page"
         Shortcut        =   {F3}
      End
   End
   Begin VB.Menu vw 
      Caption         =   "&View"
      Begin VB.Menu pretb 
         Caption         =   "Previous Tab"
         Shortcut        =   {F11}
      End
      Begin VB.Menu nextb 
         Caption         =   "Next Tab"
         Shortcut        =   {F12}
      End
      Begin VB.Menu sdr 
         Caption         =   "-"
      End
      Begin VB.Menu rab 
         Caption         =   "Refresh all Browsers"
      End
      Begin VB.Menu sab 
         Caption         =   "Stop all Browsers"
      End
      Begin VB.Menu tl 
         Caption         =   "-"
      End
      Begin VB.Menu c4 
         Caption         =   "Scrolling News"
      End
      Begin VB.Menu ch 
         Caption         =   "&Channels"
         Begin VB.Menu new 
            Caption         =   "&News"
         End
         Begin VB.Menu sch 
            Caption         =   "&Search"
         End
         Begin VB.Menu dwd 
            Caption         =   "&Download"
         End
         Begin VB.Menu wmo 
            Caption         =   "&Web Mail 1"
         End
         Begin VB.Menu wmt 
            Caption         =   "Web &Mail 2"
         End
      End
      Begin VB.Menu ts 
         Caption         =   "Te&xt Size"
         Begin VB.Menu ls 
            Caption         =   "Lar&gest"
         End
         Begin VB.Menu lg 
            Caption         =   "&Large"
         End
         Begin VB.Menu md 
            Caption         =   "&Medium"
         End
         Begin VB.Menu sll 
            Caption         =   "&Small"
         End
         Begin VB.Menu slt 
            Caption         =   "Sm&allest"
         End
      End
      Begin VB.Menu source 
         Caption         =   "Source"
      End
      Begin VB.Menu hsy 
         Caption         =   "History"
      End
   End
   Begin VB.Menu af 
      Caption         =   "Fa&vourites"
      Index           =   2
      Begin VB.Menu atf 
         Caption         =   "&Add To Favourites..."
      End
      Begin VB.Menu vf 
         Caption         =   "&View Favourites"
      End
      Begin VB.Menu of 
         Caption         =   "&Organize Favourites"
      End
   End
   Begin VB.Menu tls 
      Caption         =   "Tools"
      Begin VB.Menu mall 
         Caption         =   "Mail"
      End
      Begin VB.Menu upd 
         Caption         =   "Update"
      End
      Begin VB.Menu opns 
         Caption         =   "Options"
         HelpContextID   =   4
         Index           =   4
         Begin VB.Menu cnf 
            Caption         =   "Con&figure Channels"
         End
         Begin VB.Menu opt 
            Caption         =   "Internet Options"
         End
      End
   End
   Begin VB.Menu popi 
      Caption         =   "Popups"
      Begin VB.Menu dispop 
         Caption         =   "Disable Popup"
         Checked         =   -1  'True
         Shortcut        =   ^D
      End
   End
   Begin VB.Menu hlp 
      Caption         =   "Help"
      Begin VB.Menu honq 
         Caption         =   "Help On Q Navigator"
         Shortcut        =   {F1}
      End
      Begin VB.Menu hmbjh 
         Caption         =   "Help !!!"
      End
      Begin VB.Menu in 
         Caption         =   "Info..."
      End
   End
   Begin VB.Menu popup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu neb 
         Caption         =   "New Browser"
         Begin VB.Menu su 
            Caption         =   "Same URL"
         End
         Begin VB.Menu bu 
            Caption         =   "Blank URL"
         End
      End
      Begin VB.Menu delb 
         Caption         =   "Close Browser"
      End
      Begin VB.Menu dall 
         Caption         =   "Close All"
      End
      Begin VB.Menu atfl 
         Caption         =   "Add to Favourites"
      End
      Begin VB.Menu sa 
         Caption         =   "Save As..."
      End
      Begin VB.Menu pg 
         Caption         =   "View Source"
      End
      Begin VB.Menu vl 
         Caption         =   "View Links"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim dayAdded As Boolean
Dim NewLocation As String
Dim Today As String
Dim TodayInHistory As Integer
Dim ThisDayName As String
Dim SlashNumber
Dim Position
Dim OldLocation As String
Dim KeyNumber
Dim DayNumber As Integer
Dim nodCN As Node
Dim nodUrl As Node
Dim length
Dim tex

Dim cg
Dim Y

Dim version
Dim xadd1
Dim xadd2
Dim xadd3
Dim xadd4
Dim xadd5
Dim xadd6
Dim xadd7
Dim xadd8
Dim xadd9
Dim xadd10
Dim maxh
Dim xal
Dim X
Dim pos1
Dim pos2
Dim pos3
Dim pos4
Dim size

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Sub SaveHistory()
Close
Open App.Path & "\history.txt" For Output As #1
Dim currentNode As Node
For Each currentNode In treeHistory.Nodes
Select Case currentNode.Text
    Case "Sunday"
        Print #1, "Sunday"
    Case "Monday"
        Print #1, "Monday"
    Case "Tuesday"
        Print #1, "Tuesday"
    Case "Wednesday"
        Print #1, "Wednesday"
    Case "Thursday"
        Print #1, "Thursday"
    Case "Friday"
        Print #1, "Friday"
    Case "Saturday"
        Print #1, "Saturday"
    Case Else
        ' If currentNode.text is not a day, it might
        ' either a computer name or a complete URL
        If currentNode.Children > 0 Then
        ' currentNode.Children > 0 means currentNode.text
        ' is a computer name, then print one tab
            Print #1, vbTab; currentNode.Text
        Else
        ' currentNode.text is a complete URL, then print
        ' two tabs
            Print #1, vbTab; vbTab; currentNode.Text
        End If
End Select
Next currentNode
Close #1
End Sub
Private Sub LoadHistory()
On Error Resume Next
' This code will search a text file for tab to create a
' treeview depending on the tab number.
' I found this code on PSC, but I modified it to suit my
' needs
Dim tree_nodes() As Node
fnum = FreeFile
' Initialize KeyNumber and DayNumber
KeyNumber = 1
DayNumber = 1
' Open the history file
file_name = App.Path & "\history.txt"
    Open file_name For Input As fnum
    
    treeHistory.Nodes.Clear
    Do While Not EOF(fnum)
        ' Get a line.
        Line Input #fnum, text_line

        ' Find the level of indentation.
        Level = 1
        Do While Left$(text_line, 1) = vbTab
            Level = Level + 1
            text_line = Mid$(text_line, 2)
        Loop

        ' Make room for the new node.
        If Level > num_nodes Then
            num_nodes = Level
            ReDim Preserve tree_nodes(1 To num_nodes)
        End If

        ' Add the new node.
        Select Case Level
        Case 1
        ' If Level = 1, that means we have a day name
            Set tree_nodes(Level) = treeHistory.Nodes.Add(, , "day" & KeyNumber, text_line, 1)
                ' keyNumber will be used later in this
                ' sub and in the DeleteHistory sub
                KeyNumber = KeyNumber + 1
                If text_line = Today Then
                    ' Expand the day node
                    tree_nodes(Level).Expanded = True
                    ' TodayInHistory will be used later
                    ' in this sub and in the
                    ' DeleteHistory sub
                    TodayInHistory = 1
                    ' dayAdded will be used in the
                    ' AddToday sub
                    dayAdded = True
                    ' Today will be used in the AddToday
                    ' sub
                    Today = "day" & (KeyNumber - 1)
                    ' DayNumber will be used in the
                    ' DeleteHistory sub
                    DayNumber = KeyNumber
                End If
        Case 2
        ' If Level = 2, that means we have a computer
        ' name
            ' If TodayInHistory = 0, that means that
            ' today's name was not added to the history
            ' tree from the saved file yet. For example,
            ' if today is Wednesday and the loaded node
            ' is Tuesday (or Monday), this node will not
            ' be used to add a URL to history while
            ' using the browser, it will be used only
            ' to view the saved history, that's why there
            ' is no need to create a key for it
            If TodayInHistory = 0 Then
                Set tree_nodes(Level) = treeHistory.Nodes.Add(tree_nodes(Level - 1), tvwChild, , text_line, 2, 3)
            Else
            ' If TodayInHistory = 1, that means that today's
            ' name was added from the saved file, so a key is
            ' necessary to prevent adding a URL that is already
            ' in the history tree
                
            End If
        Case Else
            ' Same explanation as above
            If TodayInHistory = 0 Then
                Set tree_nodes(Level) = treeHistory.Nodes.Add(tree_nodes(Level - 1), tvwChild, , text_line, 4)
            Else
            End If
       End Select
    Loop

    Close fnum
End Sub
Public Sub TestToday()
' This sub is used to find out what day is today
Select Case Weekday(Now())
    Case 1
        Today = "Sunday"
    Case 2
        Today = "Monday"
    Case 3
        Today = "Tuesday"
    Case 4
        Today = "Wednesday"
    Case 5
        Today = "Thursday"
    Case 6
        Today = "Friday"
    Case 7
        Today = "Saturday"
End Select
' ThisDayName will be used in the AddToday sub
ThisDayName = Today
End Sub
Public Sub urlTest()
SlashNumber = 0
NewLocation = ""

length = Len(OldLocation)
' Count the slash number
For Position = 1 To length
    If Mid(OldLocation, Position, 1) = "/" Then
        SlashNumber = SlashNumber + 1
    End If
Next Position

Select Case SlashNumber
    Case 0
    ' Example : www.yahoo.com
        ' If there are not any slashes in the URL then
        ' there is no need to change it
        NewLocation = OldLocation
        ' If a slash is not added at the end of
        ' OldLocation, this will generate en error as
        ' NewLocation and OldLocation are used as keys
        ' in the TreeHistory
        OldLocation = OldLocation & "/"
        AddComputerNameToHistory
    Case 1
    ' Example : www.yahoo.com/r
        ' Call the OneSlashURL sub
        OneSlashURL
    Case 2
        If Left(OldLocation, 7) <> "http://" Then
        ' Example : www.yahoo.com/r/m1
            ' Call the OneSlashURL sub
            OneSlashURL
        Else
        ' Example : http://www.yahoo.com
            ' Call the TwoSlashURL sub
            TwoSlashURL
        End If
    Case Else
        If Left(OldLocation, 7) <> "http://" Then
        ' Example : www.yahoo.com/homer/?http://greetings.yahoo.com
            ' Call the OneSlashURL sub
            OneSlashURL
        Else
        ' Example : http://www.yahoo.com/r/m1
            ' Call the ThreeSlashURL sub
            ThreeSlashURL
        End If
End Select
End Sub
Public Sub OneSlashURL()
' This sub is used to retrieve the computer name from a
' URL if it looks like this : "www.yahoo.com/r/m1"
SlashNumber = 0
Position = 1
NewLocation = "" ' Null string

' The computer name in a URL is located before the first
' slash if there is no "http://" in it
While SlashNumber = 0
    If Mid(OldLocation, Position, 1) = "/" Then
        SlashNumber = SlashNumber + 1
    End If
    ' When the slash number is 1, the computer name is
    ' found
    If (SlashNumber = 0) And (Mid(OldLocation, Position, 1) <> "/") Then
        NewLocation = NewLocation & Mid(OldLocation, Position, 1)
    End If
    Position = Position + 1
Wend
' Call the AddComputerNameToHistory sub
AddComputerNameToHistory
End Sub

Public Sub TwoSlashURL()
' This sub is used to retrieve the computer name from a
' URL if it looks like this : "http://www.yahoo.com"
    ' If the slash number is 2, add "/" at the end of
    ' the URL so it can be used in the
    ' ThreeSlashURL sub because if the slash number
    ' is smaller than 3, we will have an infinite loop
    OldLocation = OldLocation & "/"
    ' Call the ThreeSlashURL sub
    ThreeSlashURL

End Sub

Public Sub ThreeSlashURL()
' This sub is used to retrieve the computer name from a
' URL if it looks like this : "http://www.yahoo.com/r/m1"
SlashNumber = 0
Position = 1
NewLocation = "" ' Null string

' The computer name in a URL is located between the
' "http://" and the next slash, which makes the slash
' number equals to 3
While SlashNumber < 3
    If Mid(OldLocation, Position, 1) = "/" Then
        SlashNumber = SlashNumber + 1
    End If
    ' When the slash number is 2, the computer name
    ' begins
    If (SlashNumber = 2) And (Mid(OldLocation, Position, 1) <> "/") Then
        NewLocation = NewLocation & Mid(OldLocation, Position, 1)
    End If
    Position = Position + 1
Wend
' Call the AddComputerNameToHistory sub
AddComputerNameToHistory
End Sub

Public Sub AddComputerNameToHistory()
' Error number 35602 is generated when the key is not
' unique. Since the NewLocation (Computer Name) is used
' as a key, the ErrHandler will work like the following:
' if the error number is not 35602, add the NewLocation
' to the HistoryTree. This is easier than assigning a
' different key to each node
On Error GoTo ErrHandler
' If you remove the WebBrowser1.GoBack from the Form_Load
' the NewLocation will be a null string and the
' OldLocation will be"http:///", that's why you will have
' to add "And OldLocation <> "http:///" in the following
' If statement
ErrHandler:
If err.Number <> 35602 Then
Set nodUrl = treeHistory.Nodes.Add(Today, tvwChild, NewLocation, NewLocation, 2, 3)
' Sort the nodes
nodUrl.Sorted = True
End If
' Call the AddUrlToHistory sub
AddUrlToHistory
End Sub


Public Sub AddUrlToHistory()
' Same explanation as AddComputerNameToHistory
On Error Resume Next
ErrHandler2:
If err.Number <> 35602 Then
treeHistory.Nodes.Add NewLocation, tvwChild, OldLocation, OldLocation, 4
End If
End Sub

Public Sub AddToday()
If dayAdded = False Then
    Set nodCN = treeHistory.Nodes.Add(, , Today, ThisDayName, 1)
    nodCN.Sorted = True
    nodCN.Expanded = True
    ' Change the value of dayAdded to True to prevent
    ' from adding the today's name to the TreeHistory
    ' again
    dayAdded = True
End If
End Sub

Public Sub DeleteHistory()
' In the LoadHistory sub the KeyNumber is increased by 1
' each time a name of a day is found. If today's name is
' found, the value in KeyNumber will be assigned to
' DayNumber, and the value 1 is assigned to
' TodayInHistory. If there are more days (after today) in
' the history file, the KeyNumber will increase and
' becomes greater than DayNumber.
' Here is an example of how the DeleteHistory sub works:
' if today is Wednesday, and Thursday was found in the
' history file, that means that this is the last week's
' history and it has to be cleared.
' But if today's name was not found in the history file,
' the value of KeyNumber will not be assigned to
' DayNumber (in the LoadHistory sub) which means that the
' value of KeyNumber will be greater than DayNumber and
' the history file will be cleared. To prevent that from
' happening, TodayInHistory is also used in the
' DeleteHistory sub like the following:

If (DayNumber < KeyNumber) And (TodayInHistory = 1) Then
    Open App.Path & "\history.txt" For Output As #4
    Close #4
    ' The TreeHistory must be cleared or else the old
    ' history will still be visible in it
    treeHistory.Nodes.Clear
    ' dayAdded will be used in the AddToday sub
    dayAdded = False
End If
End Sub

Private Function HyperJump(ByVal URL As String) As Long
    HyperJump = ShellExecute(0&, vbNullString, URL, vbNullString, vbNullString, vbNormalFocus)
End Function
Private Sub ae_Click()
cboURL.Text = "www.americanexpress.com"
cboURL_Click
End Sub

Private Sub atf_Click()
On Error Resume Next
Dim shellHelper As New ShellUIHelper
    Dim strLocationName, strLocationURL As String
    
If TabStrip1.SelectedItem.Key = "sud11" Then
    strLocationName = WebBrowser1.LocationName
    strLocationURL = WebBrowser1.LocationURL
    shellHelper.AddFavorite strLocationURL, strLocationName
End If
If TabStrip1.SelectedItem.Key = "sud22" Then
    strLocationName = WebBrowser2.LocationName
    strLocationURL = WebBrowser2.LocationURL
    shellHelper.AddFavorite strLocationURL, strLocationName
End If
If TabStrip1.SelectedItem.Key = "sud33" Then
    strLocationName = WebBrowser3.LocationName
    strLocationURL = WebBrowser3.LocationURL
    shellHelper.AddFavorite strLocationURL, strLocationName
End If
If TabStrip1.SelectedItem.Key = "sud44" Then
    strLocationName = WebBrowser4.LocationName
    strLocationURL = WebBrowser4.LocationURL
    shellHelper.AddFavorite strLocationURL, strLocationName
End If
If TabStrip1.SelectedItem.Key = "sud55" Then
    strLocationName = WebBrowser5.LocationName
    strLocationURL = WebBrowser5.LocationURL
    shellHelper.AddFavorite strLocationURL, strLocationName
End If
If TabStrip1.SelectedItem.Key = "sud66" Then
    strLocationName = WebBrowser6.LocationName
    strLocationURL = WebBrowser6.LocationURL
    shellHelper.AddFavorite strLocationURL, strLocationName
End If
If TabStrip1.SelectedItem.Key = "sud77" Then
    strLocationName = WebBrowser7.LocationName
    strLocationURL = WebBrowser7.LocationURL
    shellHelper.AddFavorite strLocationURL, strLocationName
End If
If TabStrip1.SelectedItem.Key = "sud88" Then
    strLocationName = WebBrowser8.LocationName
    strLocationURL = WebBrowser8.LocationURL
    shellHelper.AddFavorite strLocationURL, strLocationName
End If
If TabStrip1.SelectedItem.Key = "sud99" Then
    strLocationName = WebBrowser9.LocationName
    strLocationURL = WebBrowser9.LocationURL
    shellHelper.AddFavorite strLocationURL, strLocationName
End If
If TabStrip1.SelectedItem.Key = "sud10" Then
    strLocationName = WebBrowser10.LocationName
    strLocationURL = WebBrowser10.LocationURL
    shellHelper.AddFavorite strLocationURL, strLocationName
End If

 End Sub


Private Sub b1_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = b1.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub b10_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = b10.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub b2_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = b2.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub b3_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = b3.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub b4_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = b4.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub b5_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = b5.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub b6_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = b6.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub b7_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = b7.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub b8_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = b8.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub b9_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = b9.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub atfl_Click()
atf_Click
End Sub

Private Sub back_Click()
Picture4_Click
End Sub

Private Sub cad_Click()
Form4.Show
End Sub

Private Sub cac_Click()

End Sub

Private Sub bo_Click()
cboURL.Text = "http://www.thegroup.net/booktitl.htm"
cboURL_Click
End Sub

Private Sub byt_Click()
cboURL.Text = "www.byte.com"
cboURL_Click
End Sub


Private Sub cco_Click()
cboURL.Text = "www.comcentral.com"
cboURL_Click
End Sub

Private Sub cgr_Click()
cboURL.Text = "http://www.teleport.com/~ronjbeav/cybercard.shtml"
cboURL_Click
End Sub

Private Sub bu_Click()
su_Click
End Sub

Private Sub Close_Click()
If MsgBox("Are you sure you want to close this window ?", vbYesNo) = vbYes Then

Unload Me
End If
End Sub

Private Sub cnet_Click()
cboURL.Text = "www.cnet.com"
cboURL_Click
End Sub

Private Sub cnf_Click()
Form2.Show
End Sub

Private Sub cof_Click()
cboURL.Text = "www.free-n-cool.com"
cboURL_Click
End Sub

Private Sub Command1_Click()
Open "sud.dat" For Input As #1
While Not EOF(1)
Input #1, ar, s
If ar = "NEWS" Then tex = s
Wend
Close #1
tex_click
If Len(tex) > 0 Then
        cboURL.AddItem tex
        'try to navigate to the starting address
End If

End Sub


Private Sub Command10_Click()
    FileOpenProc
End Sub

Private Sub Command11_Click()
    CommonDialog1.Filter = "Text Files (*.txt)|*.txt|"

    CommonDialog1.ShowSave
    SaveFileAs (CommonDialog1.Filename)
    

End Sub

Private Sub Command12_Click()
    ' Call the new form procedure
    FileNew

End Sub

Private Sub Command2_Click()
frmSearch.Show
End Sub

Private Sub Command20_Click()

End Sub

Private Sub Command3_Click()
Open "sud.dat" For Input As #1
While Not EOF(1)
Input #1, ar, s
If ar = "DOWNLOAD" Then tex = s
Wend
Close #1
tex_click
If Len(tex) > 0 Then
        cboURL.AddItem tex
        'try to navigate to the starting address
End If

End Sub






Private Sub C4_Click()
If c4.Caption = "Scrolling News" Then
News.Navigate ("www.eth.net")
c4.Caption = "Disable News"
picScroll.Visible = True

Else
picScroll.Visible = False
c4.Caption = "Scrolling News"
End If
End Sub



Private Sub col_Click()
Picture2.Height = 1000
End Sub

Private Sub Command4_Click()
treeFavorites.Visible = False
treeHistory.Visible = False
Picture1.Visible = False
Command4.Visible = False
TabStrip1.Left = 0
TabStrip1.Width = 11895
Text1.Width = 11800
Text1.Left = 20
WebBrowser1.Left = 20
WebBrowser2.Left = 20
WebBrowser3.Left = 20
WebBrowser4.Left = 20
WebBrowser5.Left = 20
WebBrowser6.Left = 20
WebBrowser7.Left = 20
WebBrowser8.Left = 20
WebBrowser9.Left = 20
WebBrowser10.Left = 20
WebBrowser1.Width = 11800
WebBrowser2.Width = 11800
WebBrowser3.Width = 11800
WebBrowser4.Width = 11800
WebBrowser5.Width = 11800
WebBrowser6.Width = 11800
WebBrowser7.Width = 11800
WebBrowser8.Width = 11800
WebBrowser9.Width = 11800
WebBrowser10.Width = 11800

End Sub

Private Sub Command5_Click()
Open "sud.dat" For Input As #1
While Not EOF(1)
10 Input #1, ar, s
If ar = "MAIL1" Then tex = s
Wend
Close #1
tex_click
If Len(tex) > 0 Then
        cboURL.AddItem tex
        'try to navigate to the starting address
End If
    
End Sub

Private Sub Command6_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then WebBrowser1.Navigate ("about:blank"): oy = 1
If TabStrip1.SelectedItem.Key = "sud22" Then WebBrowser2.Navigate ("about:blank"): oy = 2
If TabStrip1.SelectedItem.Key = "sud33" Then WebBrowser3.Navigate ("about:blank"): oy = 3
If TabStrip1.SelectedItem.Key = "sud44" Then WebBrowser4.Navigate ("about:blank"): oy = 4
If TabStrip1.SelectedItem.Key = "sud55" Then WebBrowser5.Navigate ("about:blank"): oy = 5
If TabStrip1.SelectedItem.Key = "sud66" Then WebBrowser6.Navigate ("about:blank"): oy = 6
If TabStrip1.SelectedItem.Key = "sud77" Then WebBrowser7.Navigate ("about:blank"): oy = 7
If TabStrip1.SelectedItem.Key = "sud88" Then WebBrowser8.Navigate ("about:blank"): oy = 8
If TabStrip1.SelectedItem.Key = "sud99" Then WebBrowser9.Navigate ("about:blank"): oy = 9
If TabStrip1.SelectedItem.Key = "sud10" Then WebBrowser10.Navigate ("about:blank"): oy = 10
TabStrip1.Tabs.Remove (TabStrip1.SelectedItem.index)
End Sub

Private Sub Command7_Click()


End Sub
Private Sub Command8_Click()
Open "sud.dat" For Input As #1
While Not EOF(1)
10 Input #1, ar, s
If ar = "MAIL2" Then cboURL.Text = s
Wend
Close #1
cboURL_Click
If Len(cboURL.Text) > 0 Then
        cboURL.AddItem cboURL.Text
        'try to navigate to the starting address
End If

End Sub



Private Sub Command62_Click()

End Sub

Private Sub Command9_Click()
Open "sud.dat" For Input As #1
While Not EOF(1)
10 Input #1, ar, s
If ar = "MAIL2" Then tex = s
Wend
Close #1
tex_click
If Len(tex) > 0 Then
        cboURL.AddItem tex
        'try to navigate to the starting address
End If
    
End Sub

Private Sub cp_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then WebBrowser1.ExecWB OLECMDID_COPY, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud22" Then WebBrowser2.ExecWB OLECMDID_COPY, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud33" Then WebBrowser3.ExecWB OLECMDID_COPY, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud44" Then WebBrowser4.ExecWB OLECMDID_COPY, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud55" Then WebBrowser5.ExecWB OLECMDID_COPY, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud66" Then WebBrowser6.ExecWB OLECMDID_COPY, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud77" Then WebBrowser7.ExecWB OLECMDID_COPY, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud88" Then WebBrowser8.ExecWB OLECMDID_COPY, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud99" Then WebBrowser9.ExecWB OLECMDID_COPY, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud10" Then WebBrowser10.ExecWB OLECMDID_COPY, OLECMDEXECOPT_PROMPTUSER

End Sub

Private Sub csot_Click()
cboURL.Text = "cool.infi.net"
cboURL_Click
End Sub

Private Sub ctct_Click()
cboURL.Text = "www.chaitime.com"
cboURL_Click

End Sub

Private Sub cut_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then WebBrowser1.ExecWB OLECMDID_CUT, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud22" Then WebBrowser2.ExecWB OLECMDID_CUT, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud33" Then WebBrowser3.ExecWB OLECMDID_CUT, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud44" Then WebBrowser4.ExecWB OLECMDID_CUT, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud55" Then WebBrowser5.ExecWB OLECMDID_CUT, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud66" Then WebBrowser6.ExecWB OLECMDID_CUT, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud77" Then WebBrowser7.ExecWB OLECMDID_CUT, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud88" Then WebBrowser8.ExecWB OLECMDID_CUT, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud99" Then WebBrowser9.ExecWB OLECMDID_CUT, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud10" Then WebBrowser10.ExecWB OLECMDID_CUT, OLECMDEXECOPT_PROMPTUSER

End Sub

Private Sub d1_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = d1.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub d10_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = d10.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub d2_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = d2.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub d3_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = d3.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub d4_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = d4.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub d5_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = d5.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub d6_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = d6.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub d7_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = d7.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub d8_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = d8.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub d9_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = d9.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub dall_Click()

a = TabStrip1.Tabs.Count
If a < 3 Then Exit Sub
For t = 2 To a - 1
TabStrip1.Tabs(2).Selected = True
If WebBrowser1.Visible = True Then WebBrowser1.Navigate ("about:blank")
If WebBrowser2.Visible = True Then WebBrowser2.Navigate ("about:blank")
If WebBrowser3.Visible = True Then WebBrowser3.Navigate ("about:blank")
If WebBrowser4.Visible = True Then WebBrowser4.Navigate ("about:blank")
If WebBrowser5.Visible = True Then WebBrowser5.Navigate ("about:blank")
If WebBrowser6.Visible = True Then WebBrowser6.Navigate ("about:blank")
If WebBrowser7.Visible = True Then WebBrowser7.Navigate ("about:blank")
If WebBrowser8.Visible = True Then WebBrowser8.Navigate ("about:blank")
If WebBrowser9.Visible = True Then WebBrowser9.Navigate ("about:blank")
If WebBrowser10.Visible = True Then WebBrowser10.Navigate ("about:blank")


TabStrip1.Tabs.Remove (3)
Next t

End Sub

Private Sub delb_Click()
If WebBrowser1.Visible = True Then WebBrowser1.Navigate ("about:blank")
If WebBrowser2.Visible = True Then WebBrowser2.Navigate ("about:blank")
If WebBrowser3.Visible = True Then WebBrowser3.Navigate ("about:blank")
If WebBrowser4.Visible = True Then WebBrowser4.Navigate ("about:blank")
If WebBrowser5.Visible = True Then WebBrowser5.Navigate ("about:blank")
If WebBrowser6.Visible = True Then WebBrowser6.Navigate ("about:blank")
If WebBrowser7.Visible = True Then WebBrowser7.Navigate ("about:blank")
If WebBrowser8.Visible = True Then WebBrowser8.Navigate ("about:blank")
If WebBrowser9.Visible = True Then WebBrowser9.Navigate ("about:blank")
If WebBrowser10.Visible = True Then WebBrowser10.Navigate ("about:blank")

If TabStrip1.SelectedItem.index > 2 Then TabStrip1.Tabs.Remove (TabStrip1.SelectedItem.index)
End Sub

Private Sub dispop_Click()
If dispop.Checked = True Then dispop.Checked = False Else dispop.Checked = True
End Sub

Private Sub dwd_Click()
Command3_Click
End Sub

Private Sub fav_Click()
Form3.Show
End Sub

Private Sub dwndc_Click()
cboURL.Text = "www.download.com"
cboURL_Click
End Sub

Private Sub e1_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = e1.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub e10_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = e10.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub e2_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = e2.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub e3_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = e3.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub e4_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = e4.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub e5_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = e5.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub e6_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = e6.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub e7_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = e7.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub e8_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = e8.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub e9_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = e9.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub eco_Click()
cboURL.Text = "www.economist.com"
cboURL_Click

End Sub

Private Sub egr_Click()
cboURL.Text = "www.e-greetings.com"
cboURL_Click
End Sub

Private Sub ext_Click()
If MsgBox("Are you sure you want to quit all browser windows ?", vbYesNo) = vbYes Then
SaveHistory
Dim ret As Object

    For Each ret In Forms
        Unload ret
        Set ret = Nothing
    Next
End If
End Sub

Private Sub fh_Click()
cboURL.Text = "www.freewarehome.com"
cboURL_Click
End Sub

Private Sub fn_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then
On Error Resume Next
    SetFocusOnly = True
    TabStrip1.SetFocus
    WebBrowser1.SetFocus
    SendKeys "^f"
    Exit Sub
End If

If TabStrip1.SelectedItem.Key = "sud10" Then
On Error Resume Next
    SetFocusOnly = True
    TabStrip1.SetFocus
    WebBrowser10.SetFocus
    SendKeys "^f"
    Exit Sub
End If

If TabStrip1.SelectedItem.Key = "sud22" Then
On Error Resume Next
    SetFocusOnly = True
    TabStrip1.SetFocus
    WebBrowser2.SetFocus
    SendKeys "^f"
    Exit Sub
End If

If TabStrip1.SelectedItem.Key = "sud33" Then
On Error Resume Next
    SetFocusOnly = True
    TabStrip1.SetFocus
    WebBrowser3.SetFocus
    SendKeys "^f"
    Exit Sub
End If

If TabStrip1.SelectedItem.Key = "sud44" Then
On Error Resume Next
    SetFocusOnly = True
    TabStrip1.SetFocus
    WebBrowser4.SetFocus
    SendKeys "^f"
    Exit Sub
End If

If TabStrip1.SelectedItem.Key = "sud55" Then
On Error Resume Next
    SetFocusOnly = True
    TabStrip1.SetFocus
    WebBrowser5.SetFocus
    SendKeys "^f"
    Exit Sub
End If

If TabStrip1.SelectedItem.Key = "sud66" Then
On Error Resume Next
    SetFocusOnly = True
    TabStrip1.SetFocus
    WebBrowser6.SetFocus
    SendKeys "^f"
    Exit Sub
End If

If TabStrip1.SelectedItem.Key = "sud77" Then
On Error Resume Next
    SetFocusOnly = True
    TabStrip1.SetFocus
    WebBrowser7.SetFocus
    SendKeys "^f"
    Exit Sub
End If

If TabStrip1.SelectedItem.Key = "sud88" Then
On Error Resume Next
    SetFocusOnly = True
    TabStrip1.SetFocus
    WebBrowser8.SetFocus
    SendKeys "^f"
    Exit Sub
End If

If TabStrip1.SelectedItem.Key = "sud99" Then
On Error Resume Next
    SetFocusOnly = True
    TabStrip1.SetFocus
    WebBrowser9.SetFocus
    SendKeys "^f"
    Exit Sub
End If


End Sub

Private Sub ford_Click()
Picture5_Click
End Sub

Private Sub Form_Load()
dispop.Checked = False
Form_Resize
WebBrowser1.Navigate ("about:blank")

    StatusBar.Panels(1).Width = 2200
    StatusBar.Panels(2).Width = Me.ScaleWidth - (StatusBar.Panels(1).Width)
Call RepositionProgressBar

    
    
    
    
    If Left$(Command$, 5) = "http""" Or _
        Left$(Command$, 5) = "file""" Or _
        Right$(Command$, 4) = "htm""" Or _
        Right$(Command$, 5) = "html""" Then
        startingaddress = Left$(Command$, Len(Command$) - 1)
        startingaddress = Right$(startingaddress, Len(startingaddress) - 1)
    Else
        If Right$(Command$, 4) = "URL""" Then
            startingaddress = Left$(Command$, Len(Command$) - 1)
            startingaddress = Right$(startingaddress, Len(startingaddress) - 1)
        Else
            startingaddress = ""
        End If
    End If
If startingaddress <> "" Then WebBrowser1.Navigate (startingaddress)
' TodayInHistory will be used in the LoadHistory sub
TodayInHistory = 0
' Maximize the window
treeHistory.Visible = False
' dayAdded will be used in the AddToday sub
dayAdded = False
treeHistory.Nodes.Clear
' Call the TestToday sub
TestToday
' Call the LoadHistory sub
LoadHistory
' Call the DeleteHistory sub
DeleteHistory
' Call the AddToday sub
AddToday
Call GetFavorites
Timer1.Enabled = True
On Error GoTo yell
Open "address.ini" For Input As #1
While Not EOF(1)
Input #1, broadd
cboURL.AddItem (broadd)
Wend
Close #1
yell:
Dim l(30)
Dim b(10)
Dim c(10)
Dim d(10)
Dim g(10)
Dim e(10)
Dim h(10)
Dim m(10)


cg = WebBrowser1.Path + "notfound.html"
WebBrowser1.RegisterAsBrowser = True
WebBrowser2.RegisterAsBrowser = True
WebBrowser3.RegisterAsBrowser = True
WebBrowser4.RegisterAsBrowser = True
WebBrowser5.RegisterAsBrowser = True
WebBrowser6.RegisterAsBrowser = True
WebBrowser7.RegisterAsBrowser = True
WebBrowser8.RegisterAsBrowser = True
WebBrowser9.RegisterAsBrowser = True
WebBrowser10.RegisterAsBrowser = True
    
WebBrowser1.Top = 350
WebBrowser1.Left = 35
WebBrowser1.Height = 6250
WebBrowser1.Width = 11800

WebBrowser2.Top = 350
WebBrowser2.Left = 35
WebBrowser2.Height = 6250
WebBrowser2.Width = 11800

WebBrowser3.Top = 350
WebBrowser3.Left = 35
WebBrowser3.Height = 6250
WebBrowser3.Width = 11800

WebBrowser4.Top = 350
WebBrowser4.Left = 35
WebBrowser4.Height = 6250
WebBrowser4.Width = 11800

WebBrowser5.Top = 350
WebBrowser5.Left = 35
WebBrowser5.Height = 6250
WebBrowser5.Width = 11800

WebBrowser6.Top = 350
WebBrowser6.Left = 35
WebBrowser6.Height = 6250
WebBrowser6.Width = 11800

WebBrowser7.Top = 350
WebBrowser7.Left = 35
WebBrowser7.Height = 6250
WebBrowser7.Width = 11800

WebBrowser8.Top = 350
WebBrowser8.Left = 35
WebBrowser8.Height = 6250
WebBrowser8.Width = 11800

WebBrowser9.Top = 350
WebBrowser9.Left = 35
WebBrowser9.Height = 6250
WebBrowser9.Width = 11800

WebBrowser10.Top = 350
WebBrowser10.Left = 35
WebBrowser10.Height = 6250
WebBrowser10.Width = 11800

WebBrowser1.Visible = True

xc = 1
xadd1 = 1
xadd2 = 0
xadd3 = 0
xadd4 = 0
xadd5 = 0
xadd6 = 0
xadd7 = 0
xadd8 = 0
xadd9 = 0
xadd10 = 0

Form1.Caption = cboURL.Text + "- Q Navigator"
TabStrip1.Tabs(2).Selected = True
Form1.TabStrip1.Refresh
End Sub



Private Sub fort_Click()
cboURL.Text = "www.fortune.com"
cboURL_Click
End Sub

Private Sub ftm_Click()
cboURL.Text = "www.familytreemaker.com"
cboURL_Click

End Sub



Private Sub Form_Resize()
If treeFavorites.Visible = False And treeHistory.Visible = False Then
If Me.WindowState <> vbMinimized Then
On Error GoTo yu
Text3.Width = Form1.ScaleWidth - 9560
picScroll.Width = Form1.ScaleWidth - 9560
yu: Picture2.Height = Me.ScaleHeight - (Toolbar1.Height + cboURL.Height + 30 + StatusBar.Height)

WebBrowser1.Height = Me.ScaleHeight - (Toolbar1.Height + cboURL.Height + 300 + StatusBar.Height)
WebBrowser2.Height = Me.ScaleHeight - (Toolbar1.Height + cboURL.Height + 300 + StatusBar.Height)
WebBrowser3.Height = Me.ScaleHeight - (Toolbar1.Height + cboURL.Height + 300 + StatusBar.Height)
WebBrowser4.Height = Me.ScaleHeight - (Toolbar1.Height + cboURL.Height + 300 + StatusBar.Height)
WebBrowser5.Height = Me.ScaleHeight - (Toolbar1.Height + cboURL.Height + 300 + StatusBar.Height)
WebBrowser6.Height = Me.ScaleHeight - (Toolbar1.Height + cboURL.Height + 300 + StatusBar.Height)
WebBrowser7.Height = Me.ScaleHeight - (Toolbar1.Height + cboURL.Height + 300 + StatusBar.Height)
WebBrowser8.Height = Me.ScaleHeight - (Toolbar1.Height + cboURL.Height + 300 + StatusBar.Height)
WebBrowser9.Height = Me.ScaleHeight - (Toolbar1.Height + cboURL.Height + 300 + StatusBar.Height)
WebBrowser10.Height = Me.ScaleHeight - (Toolbar1.Height + cboURL.Height + 300 + StatusBar.Height)
Text1.Height = Me.ScaleHeight - (Toolbar1.Height + cboURL.Height + 300 + StatusBar.Height)
WebBrowser1.Width = Form1.ScaleWidth - 20
WebBrowser2.Width = Form1.ScaleWidth - 20
WebBrowser3.Width = Form1.ScaleWidth - 20
WebBrowser4.Width = Form1.ScaleWidth - 20
WebBrowser5.Width = Form1.ScaleWidth - 20
WebBrowser6.Width = Form1.ScaleWidth - 20
WebBrowser7.Width = Form1.ScaleWidth - 20
WebBrowser8.Width = Form1.ScaleWidth - 20
WebBrowser9.Width = Form1.ScaleWidth - 20
WebBrowser10.Width = Form1.ScaleWidth - 20
Text1.Width = Form1.ScaleWidth - 20
TabStrip1.Width = Form1.ScaleWidth - 20
StatusBar.Refresh
End If
If Me.WindowState <> vbMinimized And Me.WindowState = vbMaximized Then
On Error Resume Next
Dim lLeft, lTop, lWidth, lHeight As Long
    lLeft = Form1.TabStrip1.Left + 60
    lTop = Form1.TabStrip1.Top + 340
    lWidth = Form1.TabStrip1.Width - 215
    lHeight = Form1.TabStrip1.Height - 400
    Text1.Move lLeft, lTop, lWidth, lHeight
    WebBrowser1.Move lLeft, lTop, lWidth, lHeight
    WebBrowser2.Move lLeft, lTop, lWidth, lHeight
    WebBrowser3.Move lLeft, lTop, lWidth, lHeight
    WebBrowser4.Move lLeft, lTop, lWidth, lHeight
    WebBrowser5.Move lLeft, lTop, lWidth, lHeight
    WebBrowser6.Move lLeft, lTop, lWidth, lHeight
    WebBrowser7.Move lLeft, lTop, lWidth, lHeight
    WebBrowser8.Move lLeft, lTop, lWidth, lHeight
    WebBrowser9.Move lLeft, lTop, lWidth, lHeight
    WebBrowser10.Move lLeft, lTop, lWidth, lHeight
    StatusBar.Panels(1).Width = 2200
    StatusBar.Panels(2).Width = Me.ScaleWidth - (StatusBar.Panels(1).Width)
Call RepositionProgressBar
End If
End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
SaveHistory
Form2.Hide
Form3.Hide
Form4.Hide
Form6.Hide
frmCredits.Hide
frmSearch.Hide

End Sub

Private Sub hmbjh_Click()
Form3.Show
End Sub

Private Sub hsy_Click()
hist_Click
End Sub

Private Sub Label2_Change()
StatusBar.Panels(2).Text = Label2.Caption


End Sub

Private Sub pg_Click()
Form6.Show

End Sub

Private Sub qno_Click()

End Sub

Private Sub sa_Click()
mnusave_Click
End Sub

Private Sub sdb_Click()
su_Click
End Sub

Private Sub su_Click()
If WebBrowser1.LocationURL = "about:blank" Or WebBrowser1.LocationURL = "" Then
On Error GoTo yel1
TabStrip1.Tabs.Add(2, "sud11", "Browser 1") = "Browser 1"
yel1:
GoTo 66
End If
If WebBrowser2.LocationURL = "about:blank" Or WebBrowser2.LocationURL = "" Then
On Error GoTo yel2
TabStrip1.Tabs.Add(3, "sud22", "Browser 2") = "Browser 2"
yel2:
GoTo 66
End If
If WebBrowser3.LocationURL = "about:blank" Or WebBrowser3.LocationURL = "" Then
On Error GoTo yel3
TabStrip1.Tabs.Add(4, "sud33", "Browser 3") = "Browser 3"
yel3:
GoTo 66
End If

If WebBrowser4.LocationURL = "about:blank" Or WebBrowser4.LocationURL = "" Then
On Error GoTo yel4
TabStrip1.Tabs.Add(5, "sud44", "Browser 44") = "Browser 4"
yel4:
GoTo 66
End If
If WebBrowser5.LocationURL = "about:blank" Or WebBrowser5.LocationURL = "" Then
On Error GoTo yel5
TabStrip1.Tabs.Add(6, "sud55", "Browser 55") = "Browser 5"
yel5:
GoTo 66
End If
If WebBrowser6.LocationURL = "about:blank" Or WebBrowser6.LocationURL = "" Then
On Error GoTo yel6
TabStrip1.Tabs.Add(7, "sud66", "Browser 6") = "Browser 6"
yel6:
GoTo 66
End If
If WebBrowser7.LocationURL = "about:blank" Or WebBrowser7.LocationURL = "" Then
On Error GoTo yel7
TabStrip1.Tabs.Add(8, "sud77", "Browser 7") = "Browser 7"
yel7:
GoTo 66
End If
If WebBrowser8.LocationURL = "about:blank" Or WebBrowser8.LocationURL = "" Then
On Error GoTo yel8
TabStrip1.Tabs.Add(9, "sud88", "Browser 8") = "Browser 8"
yel8:
GoTo 66
End If
If WebBrowser9.LocationURL = "about:blank" Or WebBrowser9.LocationURL = "" Then
On Error GoTo yel9
TabStrip1.Tabs.Add(10, "sud99", "Browser 9") = "Browser 9"
yel9:
GoTo 66
End If
If WebBrowser10.LocationURL = "about:blank" Or WebBrowser10.LocationURL = "" Then
On Error GoTo yel10
TabStrip1.Tabs.Add(11, "sud10", "Browser 10") = "Browser 10"
yel10:
GoTo 66
End If
MsgBox ("All Browsers are in use. To close an active unwanted page click 'CLEAR' above.")
MsgBox ("The page will be opened in Internet Explorer")
66:
End Sub

Private Sub TabStrip1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 2 Then PopupMenu Me.Popup

End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim starting
On Error Resume Next
Select Case Button.Key
    Case "back"
     GoldButton2_Click
    Case "forward"
     GoldButton1_Click
    Case "stop"
     GoldButton3_Click
    Case "refresh"
     GoldButton4_Click
    Case "home"
     On Error GoTo Y
     Open "home.dat" For Input As #1
     Input #1, homet
     Close #1
     tex = homet
     tex_click
     Exit Sub
Y:   cboURL.Text = "about:blank"
     cboURL_Click
    Case "search"
     Command2_Click
    Case "links"
     GoldButton5_Click
    Case "news"
     Command1_Click
    Case "downloads"
     Command3_Click
    Case "mail"
     Command5_Click
    
End Select

End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

If ButtonMenu.Text = "Mail 1" Then Command5_Click
If ButtonMenu.Text = "Mail 2" Then Command9_Click

End Sub

Private Sub treeFavorites_NodeClick(ByVal Node As MSComctlLib.Node)

On Error GoTo treeFavorites_NodeClick_Error:
    'Navigate current Tab\Browser to the selected URL
    If Right(Node.Key, 4) = "_URL" Then
        Set Itm = Node
        cboURL.Text = Itm.Tag
        cboURL_Click
    End If
    Exit Sub
    
treeFavorites_NodeClick_Error:
    
End Sub

Private Sub TreeHistory_NodeClick(ByVal Node As MSComctlLib.Node)
' This code is used to expand, close and navigate the
' TreeHistory with a single click only

' If Node.Children = 0 that means that the clicked node
' is a URL, so the WebBrowser should navigate it
If Node.Children = 0 Then
    cboURL.Text = Node.Text

cboURL_Click
Else
' If Node.Children <> 0 that means that the clicked node
' is either a day or a computer name and the WebBrowser
' should not navigate it. If the node is expanded, it
' will be closed; if not, it will be expanded
    If Node.Expanded = True Then
        Node.Expanded = False
    Else
        Node.Expanded = True
    End If
End If

End Sub

Private Sub g1_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = g1.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub g10_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = g10.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub g2_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = g2.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub g3_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = g3.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub g4_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = g4.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub g5_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = g5.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub g6_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = g6.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub g7_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = g7.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub g8_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = g8.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub g9_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = g9.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub gc_Click()
cboURL.Text = "www.gamecenter.com"
cboURL_Click
End Sub

Private Sub gw_Click()
cboURL.Text = "www.gameworld.com"
cboURL_Click
End Sub

Private Sub h1_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = h1.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub h10_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = h10.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub h2_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = h2.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub h3_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = h3.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub h4_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = h4.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub h5_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = h5.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub h6_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = h6.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub h7_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = h7.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub h8_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = h8.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub h9_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = h9.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub ha_Click()
cboURL.Text = "www.cybercheeze.com"
cboURL_Click

End Sub

Private Sub GoldButton1_Click()
Picture5_Click
End Sub

Private Sub GoldButton2_Click()
Picture4_Click
End Sub

Private Sub GoldButton3_Click()
Picture6_Click
End Sub

Private Sub GoldButton4_Click()
Picture7_Click
End Sub

Private Sub GoldButton5_Click()
Form4.Show

End Sub

Private Sub GoldButton6_Click()
Picture8_Click

End Sub

Private Sub hist_Click()
TabStrip1.Width = 9900
TabStrip1.Left = 2000
WebBrowser1.Left = 2020
WebBrowser2.Left = 2020
WebBrowser3.Left = 2020
WebBrowser4.Left = 2020
WebBrowser5.Left = 2020
WebBrowser6.Left = 2020
WebBrowser7.Left = 2020
WebBrowser8.Left = 2020
WebBrowser9.Left = 2020
WebBrowser10.Left = 2020
Text1.Left = 2020
Text1.Width = 9900
WebBrowser1.Width = 9900
WebBrowser2.Width = 9900
WebBrowser3.Width = 9900
WebBrowser4.Width = 9900
WebBrowser5.Width = 9900
WebBrowser6.Width = 9900
WebBrowser7.Width = 9900
WebBrowser8.Width = 9900
WebBrowser9.Width = 9900
WebBrowser10.Width = 9900
treeHistory.Visible = True
Command4.Visible = True
Picture1.Visible = True

End Sub

Private Sub home_Click()
Picture8_Click
End Sub

Private Sub honq_Click()
On Error GoTo u
Dim HelpFile As String
          syspath = WindowsDirectory$
          HelpFile = syspath & "\HH.exe" & " " & App.Path & "\Help.chm"
          Shell (HelpFile), vbNormalFocus
          Exit Sub
u: MsgBox ("Help File Not Found")
End Sub

Private Sub hpg_Click()
Command6_Click
End Sub

Private Sub in_Click()
frmCredits.Show
End Sub

Private Sub jec_Click()
cboURL.Text = "www.jokes-everyday.com"
cboURL_Click

End Sub

Private Sub jok_Click()
cboURL.Text = "www.jokes.com"
cboURL_Click

End Sub


Private Sub mnq_Click()

Dim ret As Object

    For Each ret In Forms
        Unload ret
        Set ret = Nothing
    Next

End Sub

Private Sub l1_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l1.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1
End Sub

Private Sub l10_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l11.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l11_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l11.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l12_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l12.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l13_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l13.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l14_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l14.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l15_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l15.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l16_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l16.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l17_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l17.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l18_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l18.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l19_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l19.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l2_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l2.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l20_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l20.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l21_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l21.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l22_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l22.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l23_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l23.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l24_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l24.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l25_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l25.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l26_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l26.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l27_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l27.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l28_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l28.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l29_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l29.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l3_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l3.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l30_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l30.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l4_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l4.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l5_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l5.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l6_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l6.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l7_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l7.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l8_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l8.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub l9_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = l9.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub lg_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull
If TabStrip1.SelectedItem.Key = "sud22" Then WebBrowser2.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull
If TabStrip1.SelectedItem.Key = "sud33" Then WebBrowser3.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull
If TabStrip1.SelectedItem.Key = "sud44" Then WebBrowser4.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull
If TabStrip1.SelectedItem.Key = "sud55" Then WebBrowser5.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull
If TabStrip1.SelectedItem.Key = "sud66" Then WebBrowser6.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull
If TabStrip1.SelectedItem.Key = "sud77" Then WebBrowser7.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull
If TabStrip1.SelectedItem.Key = "sud88" Then WebBrowser8.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull
If TabStrip1.SelectedItem.Key = "sud99" Then WebBrowser9.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull
If TabStrip1.SelectedItem.Key = "sud10" Then WebBrowser10.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(3), vbNull

End Sub

Private Sub ls_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull
If TabStrip1.SelectedItem.Key = "sud22" Then WebBrowser2.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull
If TabStrip1.SelectedItem.Key = "sud33" Then WebBrowser3.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull
If TabStrip1.SelectedItem.Key = "sud44" Then WebBrowser4.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull
If TabStrip1.SelectedItem.Key = "sud55" Then WebBrowser5.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull
If TabStrip1.SelectedItem.Key = "sud66" Then WebBrowser6.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull
If TabStrip1.SelectedItem.Key = "sud77" Then WebBrowser7.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull
If TabStrip1.SelectedItem.Key = "sud88" Then WebBrowser8.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull
If TabStrip1.SelectedItem.Key = "sud99" Then WebBrowser9.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull
If TabStrip1.SelectedItem.Key = "sud10" Then WebBrowser10.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(4), vbNull
End Sub

Private Sub m1_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = m1.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub m10_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = m10.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub m2_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = m2.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub m3_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = m3.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub m4_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = m4.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub m5_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = m5.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub m6_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = m6.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub m7_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = m7.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub m8_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = m8.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub m9_Click()
Open App.Path + "\Favorites.ini" For Input As #1
While Not EOF(1)
Input #1, fold, URL, nam
If fold = "main" And nam = m9.Caption Then cboURL.Text = URL: cboURL_Click
Wend
Close #1

End Sub

Private Sub mall_Click()
Call Shell("start mailto:", vbHide)

End Sub

Private Sub md_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull
If TabStrip1.SelectedItem.Key = "sud22" Then WebBrowser2.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull
If TabStrip1.SelectedItem.Key = "sud33" Then WebBrowser3.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull
If TabStrip1.SelectedItem.Key = "sud44" Then WebBrowser4.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull
If TabStrip1.SelectedItem.Key = "sud55" Then WebBrowser5.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull
If TabStrip1.SelectedItem.Key = "sud66" Then WebBrowser6.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull
If TabStrip1.SelectedItem.Key = "sud77" Then WebBrowser7.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull
If TabStrip1.SelectedItem.Key = "sud88" Then WebBrowser8.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull
If TabStrip1.SelectedItem.Key = "sud99" Then WebBrowser9.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull
If TabStrip1.SelectedItem.Key = "sud10" Then WebBrowser10.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(2), vbNull

End Sub

Private Sub mnuprint_Click()
On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_PRINT, OLECMDEXECOPT_PROMPTUSER
End Sub

Private Sub mnusave_Click()
On Error Resume Next
    WebBrowser1.ExecWB OLECMDID_SAVEAS, OLECMDEXECOPT_PROMPTUSER

End Sub

Private Sub mon_Click()
cboURL.Text = "www.money.com"
cboURL_Click
End Sub

Private Sub mp_Click()
cboURL.Text = "www.mp3.com"
cboURL_Click
End Sub

Private Sub mtvo_Click()
cboURL.Text = "mtv.com/index.html"
cboURL_Click

End Sub

Private Sub new_Click()
Command1_Click
End Sub

Private Sub mnuOpen_Click(index As Integer)
On Error Resume Next
    Form1.CommonDialog2.Filter = "Internet Files" & _
    "(*.html,*.htm)|*.htm*|(*.xml)|*.xml|Picture Files (*.jpg)|*.jpg*|Picture Files (*.gif)|*.gif|Picture Files(*.bmp)|*.bmp"
    Form1.CommonDialog2.Filename = ""
    Form1.CommonDialog2.ShowOpen
tex = CommonDialog2.Filename

tex_click
End Sub


Private Sub newwk_Click()
cboURL.Text = "www.newsweek.com"
cboURL_Click

End Sub

Private Sub News_DocumentComplete(ByVal pDisp As Object, URL As Variant)
    Y = ""
    On Error GoTo ErrHandler 'Goto to ErrHandler if an error occurs
X = News.Document.documentElement.innerHTML
For a = 1 To Len(X) - 6
If Mid(X, a, 12) = "<PARAM NAME=" And Mid(X, a + 13, 12) = "message_file" Then
a = a + 34
While Mid(X, a + 1, 1) <> ">"
Y = Y + Mid(X, a, 1)
a = a + 1
Wend
End If
Next a
For a = 1 To Len(Y)
If Mid(Y, a, 1) = "(" Then
While Mid(Y, a, 1) <> ")"
a = a + 1
Wend
Else
yp = yp + Mid(Y, a, 1)
End If
Next a
Y = yp
If Y = "" Then GoTo abort
Y = "   NEWS  :  " + Y

Text3.Top = picScroll.Height
'Variable Declarations
    Dim iFileNum As Integer
    Dim lLineCount As Long
    Dim lLineHeight As Long
    
    On Error GoTo ErrHandler 'Goto to ErrHandler if an error occurs
    
        picScroll.Visible = True
        tmrScroll.Enabled = False
        'open file and read text from it
        Text3.Text = Y
        lLineCount = SendMessage(Text3.hwnd, EM_GETLINECOUNT, 0&, 0&)
        lLineHeight = TextHeight("TEST") 'Get the height of text in file
        Text3.Height = lLineHeight * lLineCount
        tmrScroll.Enabled = True
         picScroll.Visible = True
    Exit Sub

ErrHandler:
MsgBox ("The News Server Has Gone Down or is out of reach.")
    Resume Next
GoTo 1
abort: Text3.Text = "Connection Reset By Server"
1 End Sub

Private Sub News_DownloadBegin()
Text3.Text = "Contacting Server..."
End Sub

Private Sub News_DownloadComplete()
Text3.Text = "Loading News..."
End Sub

Private Sub nextb_Click()
Call sudnex
End Sub

Private Sub of_Click()
On Error GoTo mnu_OrganizeFavorites_Click_Error:
    Dim lpszRootFolder As String
    Dim success As Long
    Dim CSIDL As Long

    'open the organize folder at the path specified by the CSIDL
    CSIDL = CSIDL_FAVORITES
     
    lpszRootFolder = GetFolderPath(CSIDL)
    success = DoOrganizeFavDlg(hwnd, lpszRootFolder)
    
    Form1.Refresh
    Call GetFavorites  'To refresh the favorites tree
    Exit Sub
mnu_OrganizeFavorites_Click_Error:
    

End Sub

Private Sub ogr_Click()
cboURL.Text = "www.ogr.com"
cboURL_Click
End Sub


Private Sub pcg_Click()
cboURL.Text = "www.pcgamer.com"
cboURL_Click
End Sub

Private Sub pcw_Click()
cboURL.Text = "www.pcworld.com"
cboURL_Click
End Sub


Private Sub opt_Click()
    Dim RetVal
    RetVal = Shell("rundll32.exe shell32.dll,Control_RunDLL Inetcpl.cpl", vbNormalFocus)

End Sub

Private Sub Picture4_Click()

If TabStrip1.SelectedItem.Key = "sud11" Then
On Error Resume Next
WebBrowser1.GoBack
cboURL.Text = WebBrowser1.LocationURL
Form1.Caption = WebBrowser1.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud22" Then
On Error Resume Next
WebBrowser2.GoBack
cboURL.Text = WebBrowser2.LocationURL
Form1.Caption = WebBrowser2.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud33" Then
On Error Resume Next
WebBrowser3.GoBack
cboURL.Text = WebBrowser3.LocationURL
Form1.Caption = WebBrowser3.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud44" Then
On Error Resume Next
WebBrowser4.GoBack
cboURL.Text = WebBrowser4.LocationURL
Form1.Caption = WebBrowser4.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud55" Then
On Error Resume Next
WebBrowser5.GoBack
cboURL.Text = WebBrowser5.LocationURL
Form1.Caption = WebBrowser5.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud66" Then
On Error Resume Next
WebBrowser6.GoBack
cboURL.Text = WebBrowser6.LocationURL
Form1.Caption = WebBrowser6.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud77" Then
On Error Resume Next
WebBrowser7.GoBack
cboURL.Text = WebBrowser7.LocationURL
Form1.Caption = WebBrowser7.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud88" Then
On Error Resume Next
WebBrowser8.GoBack
cboURL.Text = WebBrowser8.LocationURL
Form1.Caption = WebBrowser8.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud99" Then
On Error Resume Next
WebBrowser9.GoBack
cboURL.Text = WebBrowser9.LocationURL
Form1.Caption = WebBrowser9.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud10" Then
On Error Resume Next
WebBrowser10.GoBack
cboURL.Text = WebBrowser10.LocationURL
Form1.Caption = WebBrowser10.LocationName
End If

End Sub

Private Sub Picture5_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then
On Error Resume Next
WebBrowser1.GoForward
cboURL.Text = WebBrowser1.LocationURL
Form1.Caption = WebBrowser1.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud22" Then
On Error Resume Next
WebBrowser2.GoForward
cboURL.Text = WebBrowser2.LocationURL
Form1.Caption = WebBrowser2.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud33" Then
On Error Resume Next
WebBrowser3.GoForward
cboURL.Text = WebBrowser3.LocationURL
Form1.Caption = WebBrowser3.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud44" Then
On Error Resume Next
WebBrowser4.GoForward
cboURL.Text = WebBrowser4.LocationURL
Form1.Caption = WebBrowser4.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud55" Then
On Error Resume Next
WebBrowser5.GoForward
cboURL.Text = WebBrowser5.LocationURL
Form1.Caption = WebBrowser5.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud66" Then
On Error Resume Next
WebBrowser6.GoForward
cboURL.Text = WebBrowser6.LocationURL
Form1.Caption = WebBrowser6.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud77" Then
On Error Resume Next
WebBrowser7.GoForward
cboURL.Text = WebBrowser7.LocationURL
Form1.Caption = WebBrowser7.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud88" Then
On Error Resume Next
WebBrowser8.GoForward
cboURL.Text = WebBrowser8.LocationURL
Form1.Caption = WebBrowser8.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud99" Then
On Error Resume Next
WebBrowser9.GoForward
cboURL.Text = WebBrowser9.LocationURL
Form1.Caption = WebBrowser9.LocationName
End If

If TabStrip1.SelectedItem.Key = "sud10" Then
On Error Resume Next
WebBrowser10.GoForward
cboURL.Text = WebBrowser10.LocationURL
Form1.Caption = WebBrowser10.LocationName
End If

End Sub

Private Sub Picture6_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then WebBrowser1.Stop
If TabStrip1.SelectedItem.Key = "sud22" Then WebBrowser2.Stop
If TabStrip1.SelectedItem.Key = "sud33" Then WebBrowser3.Stop
If TabStrip1.SelectedItem.Key = "sud44" Then WebBrowser4.Stop
If TabStrip1.SelectedItem.Key = "sud55" Then WebBrowser5.Stop
If TabStrip1.SelectedItem.Key = "sud66" Then WebBrowser6.Stop
If TabStrip1.SelectedItem.Key = "sud77" Then WebBrowser7.Stop
If TabStrip1.SelectedItem.Key = "sud88" Then WebBrowser8.Stop
If TabStrip1.SelectedItem.Key = "sud99" Then WebBrowser9.Stop
If TabStrip1.SelectedItem.Key = "sud10" Then WebBrowser10.Stop
End Sub

Private Sub Picture7_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then WebBrowser1.Refresh
If TabStrip1.SelectedItem.Key = "sud22" Then WebBrowser2.Refresh
If TabStrip1.SelectedItem.Key = "sud33" Then WebBrowser3.Refresh
If TabStrip1.SelectedItem.Key = "sud44" Then WebBrowser4.Refresh
If TabStrip1.SelectedItem.Key = "sud55" Then WebBrowser5.Refresh
If TabStrip1.SelectedItem.Key = "sud66" Then WebBrowser6.Refresh
If TabStrip1.SelectedItem.Key = "sud77" Then WebBrowser7.Refresh
If TabStrip1.SelectedItem.Key = "sud88" Then WebBrowser8.Refresh
If TabStrip1.SelectedItem.Key = "sud99" Then WebBrowser9.Refresh
If TabStrip1.SelectedItem.Key = "sud10" Then WebBrowser10.Refresh

End Sub

Private Sub Picture8_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then WebBrowser1.GoHome
If TabStrip1.SelectedItem.Key = "sud22" Then WebBrowser2.GoHome
If TabStrip1.SelectedItem.Key = "sud33" Then WebBrowser3.GoHome
If TabStrip1.SelectedItem.Key = "sud44" Then WebBrowser4.GoHome
If TabStrip1.SelectedItem.Key = "sud55" Then WebBrowser5.GoHome
If TabStrip1.SelectedItem.Key = "sud66" Then WebBrowser6.GoHome
If TabStrip1.SelectedItem.Key = "sud77" Then WebBrowser7.GoHome
If TabStrip1.SelectedItem.Key = "sud88" Then WebBrowser8.GoHome
If TabStrip1.SelectedItem.Key = "sud99" Then WebBrowser9.GoHome
If TabStrip1.SelectedItem.Key = "sud10" Then WebBrowser10.GoHome
End Sub

Private Sub cboURL_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then
If Len(cboURL.Text) > 0 Then
        cboURL.AddItem cboURL.Text
        'try to navigate to the starting address
End If
cboURL_Click
End If
End Sub
Private Sub cboURL_Click()
If cboURL.Text = "Monday" Or cboURL.Text = "Tuesday" Or cboURL.Text = "Wednesday" Or cboURL.Text = "Thursday" Or cboURL.Text = "Friday" Or cboURL.Text = "Saturday" Or cboURL.Text = "Sunday" Then cboURL.Text = "": TabStrip1_Click: Exit Sub

On Error GoTo u
Open "address.ini" For Input As #1
While Not EOF(1)
Input #1, broadd
If cboURL.Text = broadd Then Close #1: GoTo Y
Wend
Close #1
u: Open "address.ini" For Append As #1
Write #1, cboURL.Text
Close #1
Y:
If TabStrip1.SelectedItem.Key = "sud11" Then
WebBrowser1.Navigate (cboURL.Text)
OldLocation = WebBrowser1.LocationURL
om = 1
End If

If TabStrip1.SelectedItem.Key = "sud22" Then
WebBrowser2.Navigate (cboURL.Text)
OldLocation = WebBrowser2.LocationURL
om = 2
End If

If TabStrip1.SelectedItem.Key = "sud33" Then
WebBrowser3.Navigate (cboURL.Text)
OldLocation = WebBrowser3.LocationURL
om = 3
End If

If TabStrip1.SelectedItem.Key = "sud44" Then
WebBrowser4.Navigate (cboURL.Text)
OldLocation = WebBrowser4.LocationURL
om = 4
End If

If TabStrip1.SelectedItem.Key = "sud55" Then
WebBrowser5.Navigate (cboURL.Text)
OldLocation = WebBrowser5.LocationURL
om = 5
End If

If TabStrip1.SelectedItem.Key = "sud66" Then
WebBrowser6.Navigate (cboURL.Text)
OldLocation = WebBrowser6.LocationURL
om = 6
End If

If TabStrip1.SelectedItem.Key = "sud77" Then
WebBrowser7.Navigate (cboURL.Text)
OldLocation = WebBrowser7.LocationURL
om = 7
End If

If TabStrip1.SelectedItem.Key = "sud88" Then
WebBrowser8.Navigate (cboURL.Text)
OldLocation = WebBrowser8.LocationURL
om = 8
End If

If TabStrip1.SelectedItem.Key = "sud99" Then
WebBrowser9.Navigate (cboURL.Text)
OldLocation = WebBrowser9.LocationURL
om = 9
End If

If TabStrip1.SelectedItem.Key = "sud10" Then
WebBrowser10.Navigate (cboURL.Text)
OldLocation = WebBrowser10.LocationURL
om = 10
End If
If cboURL.Text <> "about:blank" Then
Open "prev.ini" For Input As #1
Open "temp.ini" For Output As #2
For a = 1 To 10
Input #1, previnst
If a = om Then
Write #2, cboURL.Text
Else
Write #2, previnst
End If
Next a
Close #1
Close #2
Kill "prev.ini"
Name "temp.ini" As "prev.ini"
End If
urlTest

End Sub

Private Sub Picture9_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then
Form4.Show
Else
MsgBox ("Not yet ready")
End If

If TabStrip1.SelectedItem.Key = "sud10" And WebBrowser10.Busy = False Then
Form4.Show
Else
MsgBox ("Not yet ready")
End If

If TabStrip1.SelectedItem.Key = "sud22" And WebBrowser2.Busy = False Then
Form4.Show
Else
MsgBox ("Not yet ready")
End If

If TabStrip1.SelectedItem.Key = "sud33" And WebBrowser3.Busy = False Then
Form4.Show
Else
MsgBox ("Not yet ready")
End If

If TabStrip1.SelectedItem.Key = "sud44" And WebBrowser4.Busy = False Then
Form4.Show
Else
MsgBox ("Not yet ready")
End If

If TabStrip1.SelectedItem.Key = "sud55" And WebBrowser5.Busy = False Then
Form4.Show
Else
MsgBox ("Not yet ready")
End If

If TabStrip1.SelectedItem.Key = "sud66" And WebBrowser6.Busy = False Then
Form4.Show
Else
MsgBox ("Not yet ready")
End If

If TabStrip1.SelectedItem.Key = "sud77" And WebBrowser7.Busy = False Then
Form4.Show
Else
MsgBox ("Not yet ready")
End If

If TabStrip1.SelectedItem.Key = "sud88" And WebBrowser8.Busy = False Then
Form4.Show
Else
MsgBox ("Not yet ready")
End If

If TabStrip1.SelectedItem.Key = "sud99" And WebBrowser9.Busy = False Then
Form4.Show
Else
MsgBox ("Not yet ready")
End If

End Sub

Private Sub piop_Click()
cboURL.Text = "www.programmersparadise.com"
cboURL_Click
End Sub

Private Sub pls_Click()
cboURL.Text = "www.playsite.com"
cboURL_Click
End Sub

Private Sub pretb_Click()
Call sudpre
End Sub

Private Sub pstp_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then WebBrowser1.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud22" Then WebBrowser2.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud33" Then WebBrowser3.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud44" Then WebBrowser4.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud55" Then WebBrowser5.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud66" Then WebBrowser6.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud77" Then WebBrowser7.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud88" Then WebBrowser8.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud99" Then WebBrowser9.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud10" Then WebBrowser10.ExecWB OLECMDID_PAGESETUP, OLECMDEXECOPT_PROMPTUSER

End Sub

Private Sub pt_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then WebBrowser1.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud22" Then WebBrowser2.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud33" Then WebBrowser3.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud44" Then WebBrowser4.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud55" Then WebBrowser5.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud66" Then WebBrowser6.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud77" Then WebBrowser7.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud88" Then WebBrowser8.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud99" Then WebBrowser9.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_PROMPTUSER
If TabStrip1.SelectedItem.Key = "sud10" Then WebBrowser10.ExecWB OLECMDID_PASTE, OLECMDEXECOPT_PROMPTUSER


End Sub

Private Sub ra_Click()
cboURL.Text = "www.realaudio.com"
cboURL_Click

End Sub

Private Sub rab_Click()
WebBrowser1.Refresh
WebBrowser2.Refresh
WebBrowser3.Refresh
WebBrowser4.Refresh
WebBrowser5.Refresh
WebBrowser6.Refresh
WebBrowser7.Refresh
WebBrowser8.Refresh
WebBrowser9.Refresh
WebBrowser10.Refresh

End Sub

Private Sub refr_Click()
Picture6_Click
End Sub

Private Sub sc_Click()
frmSearch.Show
End Sub

Private Sub sab_Click()
WebBrowser1.Stop
WebBrowser2.Stop
WebBrowser3.Stop
WebBrowser4.Stop
WebBrowser5.Stop
WebBrowser6.Stop
WebBrowser7.Stop
WebBrowser8.Stop
WebBrowser9.Stop
WebBrowser10.Stop

End Sub

Private Sub sall_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then
Clipboard.Clear
Me.WebBrowser1.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End If
If TabStrip1.SelectedItem.Key = "sud22" Then
Clipboard.Clear
Me.WebBrowser2.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End If
If TabStrip1.SelectedItem.Key = "sud33" Then
Clipboard.Clear
Me.WebBrowser3.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End If
If TabStrip1.SelectedItem.Key = "sud44" Then
Clipboard.Clear
Me.WebBrowser4.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End If
If TabStrip1.SelectedItem.Key = "sud55" Then
Clipboard.Clear
Me.WebBrowser5.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End If
If TabStrip1.SelectedItem.Key = "sud66" Then
Clipboard.Clear
Me.WebBrowser6.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End If
If TabStrip1.SelectedItem.Key = "sud77" Then
Clipboard.Clear
Me.WebBrowser7.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End If
If TabStrip1.SelectedItem.Key = "sud88" Then
Clipboard.Clear
Me.WebBrowser8.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End If
If TabStrip1.SelectedItem.Key = "sud99" Then
Clipboard.Clear
Me.WebBrowser9.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End If
If TabStrip1.SelectedItem.Key = "sud10" Then
Clipboard.Clear
Me.WebBrowser10.ExecWB OLECMDID_SELECTALL, OLECMDEXECOPT_DONTPROMPTUSER
End If






End Sub

Private Sub sch_Click()
Command2_Click
End Sub

Private Sub sit_Click()
frmSearch.Show
End Sub


Private Sub sn_Click()
Form1.SetFocus
End Sub


Private Sub sll_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull
If TabStrip1.SelectedItem.Key = "sud22" Then WebBrowser2.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull
If TabStrip1.SelectedItem.Key = "sud33" Then WebBrowser3.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull
If TabStrip1.SelectedItem.Key = "sud44" Then WebBrowser4.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull
If TabStrip1.SelectedItem.Key = "sud55" Then WebBrowser5.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull
If TabStrip1.SelectedItem.Key = "sud66" Then WebBrowser6.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull
If TabStrip1.SelectedItem.Key = "sud77" Then WebBrowser7.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull
If TabStrip1.SelectedItem.Key = "sud88" Then WebBrowser8.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull
If TabStrip1.SelectedItem.Key = "sud99" Then WebBrowser9.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull
If TabStrip1.SelectedItem.Key = "sud10" Then WebBrowser10.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(1), vbNull

End Sub

Private Sub slt_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then WebBrowser1.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull
If TabStrip1.SelectedItem.Key = "sud22" Then WebBrowser2.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull
If TabStrip1.SelectedItem.Key = "sud33" Then WebBrowser3.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull
If TabStrip1.SelectedItem.Key = "sud44" Then WebBrowser4.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull
If TabStrip1.SelectedItem.Key = "sud55" Then WebBrowser5.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull
If TabStrip1.SelectedItem.Key = "sud66" Then WebBrowser6.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull
If TabStrip1.SelectedItem.Key = "sud77" Then WebBrowser7.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull
If TabStrip1.SelectedItem.Key = "sud88" Then WebBrowser8.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull
If TabStrip1.SelectedItem.Key = "sud99" Then WebBrowser9.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull
If TabStrip1.SelectedItem.Key = "sud10" Then WebBrowser10.ExecWB OLECMDID_ZOOM, OLECMDEXECOPT_DONTPROMPTUSER, CLng(0), vbNull

End Sub

Private Sub ssk_Click()
cboURL.Text = "www.softseek.com"
cboURL_Click
End Sub

Private Sub source_Click()
Form6.Show
End Sub

Private Sub stp_Click()
Picture7_Click
End Sub
Private Sub text3_GotFocus()
    Form1.SetFocus
    
    'Don't let the text box get focus, althought the text
    'box is locked it looks bad to see a cursor in the
    'text box as it scrolls up
    
    
End Sub
Private Sub tex_click()
ar$ = tex
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

End Sub
Private Sub TabStrip1_Click()
If TabStrip1.SelectedItem.Key = "sud11" Then
Form1.Caption = WebBrowser1.LocationName + "-Q Navigator"
cboURL.Text = WebBrowser1.LocationURL
WebBrowser1.Visible = True
WebBrowser2.Visible = False
WebBrowser3.Visible = False
WebBrowser4.Visible = False
WebBrowser5.Visible = False
WebBrowser6.Visible = False
WebBrowser7.Visible = False
WebBrowser8.Visible = False
WebBrowser9.Visible = False
WebBrowser10.Visible = False

ProgressBar1.Visible = True
ProgressBar2.Visible = False
ProgressBar3.Visible = False
ProgressBar4.Visible = False
ProgressBar5.Visible = False
ProgressBar6.Visible = False
ProgressBar7.Visible = False
ProgressBar8.Visible = False
ProgressBar9.Visible = False
ProgressBar10.Visible = False


End If

If TabStrip1.SelectedItem.Key = "sud22" Then
 Form1.Caption = WebBrowser2.LocationName + "-Q Navigator"
cboURL.Text = WebBrowser2.LocationURL
WebBrowser1.Visible = False
WebBrowser2.Visible = True
WebBrowser3.Visible = False
WebBrowser4.Visible = False
WebBrowser5.Visible = False
WebBrowser6.Visible = False
WebBrowser7.Visible = False
WebBrowser8.Visible = False
WebBrowser9.Visible = False
WebBrowser10.Visible = False

ProgressBar1.Visible = False
ProgressBar2.Visible = True
ProgressBar3.Visible = False
ProgressBar4.Visible = False
ProgressBar5.Visible = False
ProgressBar6.Visible = False
ProgressBar7.Visible = False
ProgressBar8.Visible = False
ProgressBar9.Visible = False
ProgressBar10.Visible = False
End If

If TabStrip1.SelectedItem.Key = "sud33" Then
 Form1.Caption = WebBrowser3.LocationName + "-Q Navigator"
cboURL.Text = WebBrowser3.LocationURL
WebBrowser1.Visible = False
WebBrowser2.Visible = False
WebBrowser3.Visible = True
WebBrowser4.Visible = False
WebBrowser5.Visible = False
WebBrowser6.Visible = False
WebBrowser7.Visible = False
WebBrowser8.Visible = False
WebBrowser9.Visible = False
WebBrowser10.Visible = False

ProgressBar1.Visible = False
ProgressBar2.Visible = False
ProgressBar3.Visible = True
ProgressBar4.Visible = False
ProgressBar5.Visible = False
ProgressBar6.Visible = False
ProgressBar7.Visible = False
ProgressBar8.Visible = False
ProgressBar9.Visible = False
ProgressBar10.Visible = False

End If

If TabStrip1.SelectedItem.Key = "sud44" Then
 Form1.Caption = WebBrowser4.LocationName + "-Q Navigator"
cboURL.Text = WebBrowser4.LocationURL
WebBrowser1.Visible = False
WebBrowser2.Visible = False
WebBrowser3.Visible = False
WebBrowser4.Visible = True
WebBrowser5.Visible = False
WebBrowser6.Visible = False
WebBrowser7.Visible = False
WebBrowser8.Visible = False
WebBrowser9.Visible = False
WebBrowser10.Visible = False

ProgressBar1.Visible = False
ProgressBar2.Visible = False
ProgressBar3.Visible = False
ProgressBar4.Visible = True
ProgressBar5.Visible = False
ProgressBar6.Visible = False
ProgressBar7.Visible = False
ProgressBar8.Visible = False
ProgressBar9.Visible = False
ProgressBar10.Visible = False

End If

If TabStrip1.SelectedItem.Key = "sud55" Then
 Form1.Caption = WebBrowser5.LocationName + "-Q Navigator"
cboURL.Text = WebBrowser5.LocationURL
WebBrowser1.Visible = False
WebBrowser2.Visible = False
WebBrowser3.Visible = False
WebBrowser4.Visible = False
WebBrowser5.Visible = True
WebBrowser6.Visible = False
WebBrowser7.Visible = False
WebBrowser8.Visible = False
WebBrowser9.Visible = False
WebBrowser10.Visible = False

ProgressBar1.Visible = False
ProgressBar2.Visible = False
ProgressBar3.Visible = False
ProgressBar4.Visible = False
ProgressBar5.Visible = True
ProgressBar6.Visible = False
ProgressBar7.Visible = False
ProgressBar8.Visible = False
ProgressBar9.Visible = False
ProgressBar10.Visible = False
End If

If TabStrip1.SelectedItem.Key = "sud66" Then
 Form1.Caption = WebBrowser6.LocationName + "-Q Navigator"
cboURL.Text = WebBrowser6.LocationURL
WebBrowser1.Visible = False
WebBrowser2.Visible = False
WebBrowser3.Visible = False
WebBrowser4.Visible = False
WebBrowser5.Visible = False
WebBrowser6.Visible = True
WebBrowser7.Visible = False
WebBrowser8.Visible = False
WebBrowser9.Visible = False
WebBrowser10.Visible = False

ProgressBar1.Visible = False
ProgressBar2.Visible = False
ProgressBar3.Visible = False
ProgressBar4.Visible = False
ProgressBar5.Visible = False
ProgressBar6.Visible = True
ProgressBar7.Visible = False
ProgressBar8.Visible = False
ProgressBar9.Visible = False
ProgressBar10.Visible = False
End If

If TabStrip1.SelectedItem.Key = "sud77" Then
 Form1.Caption = WebBrowser7.LocationName + "-Q Navigator"
cboURL.Text = WebBrowser7.LocationURL
WebBrowser1.Visible = False
WebBrowser2.Visible = False
WebBrowser3.Visible = False
WebBrowser4.Visible = False
WebBrowser5.Visible = False
WebBrowser6.Visible = False
WebBrowser7.Visible = True
WebBrowser8.Visible = False
WebBrowser9.Visible = False
WebBrowser10.Visible = False

ProgressBar1.Visible = False
ProgressBar2.Visible = False
ProgressBar3.Visible = False
ProgressBar4.Visible = False
ProgressBar5.Visible = False
ProgressBar6.Visible = False
ProgressBar7.Visible = True
ProgressBar8.Visible = False
ProgressBar9.Visible = False
ProgressBar10.Visible = False

End If

If TabStrip1.SelectedItem.Key = "sud88" Then
 Form1.Caption = WebBrowser8.LocationName + "-Q Navigator"
cboURL.Text = WebBrowser8.LocationURL
WebBrowser1.Visible = False
WebBrowser2.Visible = False
WebBrowser3.Visible = False
WebBrowser4.Visible = False
WebBrowser5.Visible = False
WebBrowser6.Visible = False
WebBrowser7.Visible = False
WebBrowser8.Visible = True
WebBrowser9.Visible = False
WebBrowser10.Visible = False

ProgressBar1.Visible = False
ProgressBar2.Visible = False
ProgressBar3.Visible = False
ProgressBar4.Visible = False
ProgressBar5.Visible = False
ProgressBar6.Visible = False
ProgressBar7.Visible = False
ProgressBar8.Visible = True
ProgressBar9.Visible = False
ProgressBar10.Visible = False

End If

If TabStrip1.SelectedItem.Key = "sud99" Then
 Form1.Caption = WebBrowser9.LocationName + "-Q Navigator"
cboURL.Text = WebBrowser9.LocationURL
WebBrowser1.Visible = False
WebBrowser2.Visible = False
WebBrowser3.Visible = False
WebBrowser4.Visible = False
WebBrowser5.Visible = False
WebBrowser6.Visible = False
WebBrowser7.Visible = False
WebBrowser8.Visible = False
WebBrowser9.Visible = True
WebBrowser10.Visible = False

ProgressBar1.Visible = False
ProgressBar2.Visible = False
ProgressBar3.Visible = False
ProgressBar4.Visible = False
ProgressBar5.Visible = False
ProgressBar6.Visible = False
ProgressBar7.Visible = False
ProgressBar8.Visible = False
ProgressBar9.Visible = True
ProgressBar10.Visible = False

End If

If TabStrip1.SelectedItem.Key = "sud10" Then
 Form1.Caption = WebBrowser10.LocationName + "-Q Navigator"
cboURL.Text = WebBrowser10.LocationURL
WebBrowser1.Visible = False
WebBrowser2.Visible = False
WebBrowser3.Visible = False
WebBrowser4.Visible = False
WebBrowser5.Visible = False
WebBrowser6.Visible = False
WebBrowser7.Visible = False
WebBrowser8.Visible = False
WebBrowser9.Visible = False
WebBrowser10.Visible = True

ProgressBar1.Visible = False
ProgressBar2.Visible = False
ProgressBar3.Visible = False
ProgressBar4.Visible = False
ProgressBar5.Visible = False
ProgressBar6.Visible = False
ProgressBar7.Visible = False
ProgressBar8.Visible = False
ProgressBar9.Visible = False
ProgressBar10.Visible = True

End If

If TabStrip1.SelectedItem.Caption = "Q - Pad" Then

Text1.Top = 350
Text1.Left = 35
Text1.Width = 11800
Text1.Height = 6250
Text1.Visible = True
Form1.Caption = "Q - Pad -Q Navigator"

WebBrowser1.Visible = False
WebBrowser2.Visible = False
WebBrowser3.Visible = False
WebBrowser4.Visible = False
WebBrowser5.Visible = False
WebBrowser6.Visible = False
WebBrowser7.Visible = False
WebBrowser8.Visible = False
WebBrowser9.Visible = False
WebBrowser10.Visible = False

ProgressBar1.Visible = False
ProgressBar2.Visible = False
ProgressBar3.Visible = False
ProgressBar4.Visible = False
ProgressBar5.Visible = False
ProgressBar6.Visible = False
ProgressBar7.Visible = False
ProgressBar8.Visible = False
ProgressBar9.Visible = False
ProgressBar10.Visible = False



End If

If TabStrip1.SelectedItem.Caption <> "Q - Pad" Then
Text1.Visible = False

End If
If TabStrip1.SelectedItem.Caption = "" Then
TabStrip1.SelectedItem.Caption = "Blank"
End If
End Sub

Private Sub tun_Click()
cboURL.Text = "www.tunes.com"
cboURL_Click

End Sub

Private Sub tw_Click()
cboURL.Text = "www.tucows.com"
cboURL_Click
End Sub

Private Sub umc_Click()
cboURL.Text = "www.unitedmedia.com"
cboURL_Click

End Sub

Private Sub Text4_Change()
cboURL.Text = Text4.Text
If WebBrowser1.LocationURL = "about:blank" Then WebBrowser1.Navigate (Text4.Text): GoTo 669
If WebBrowser2.LocationURL = "about:blank" Then WebBrowser2.Navigate (Text4.Text): GoTo 669
If WebBrowser3.LocationURL = "about:blank" Then WebBrowser3.Navigate (Text4.Text): GoTo 669
If WebBrowser4.LocationURL = "about:blank" Then WebBrowser4.Navigate (Text4.Text): GoTo 669
If WebBrowser5.LocationURL = "about:blank" Then WebBrowser5.Navigate (Text4.Text): GoTo 669
If WebBrowser6.LocationURL = "about:blank" Then WebBrowser6.Navigate (Text4.Text): GoTo 669
If WebBrowser7.LocationURL = "about:blank" Then WebBrowser7.Navigate (Text4.Text): GoTo 669
If WebBrowser8.LocationURL = "about:blank" Then WebBrowser8.Navigate (Text4.Text): GoTo 669
If WebBrowser9.LocationURL = "about:blank" Then WebBrowser9.Navigate (Text4.Text): GoTo 669
If WebBrowser10.LocationURL = "about:blank" Then WebBrowser10.Navigate (Text4.Text): GoTo 669
MsgBox ("All Browsers are in use. To close an active unwanted page click 'CLEAR' above.")
MsgBox ("The page will be opened in Internet Explorer")
669: End Sub

Private Sub Timer1_Timer()
If IsConnected = True Then
News.Navigate ("www.eth.net")
Timer1.Enabled = False
c4.Caption = "Disable News"
End If

End Sub

Private Sub tmrScroll_Timer()
If Text3.Top + Text3.Height < picScroll.Top Then 'picScroll.Top
        Text3.Text = Y
        Text3.Top = picScroll.Height
    Else
        Text3.Top = Text3.Top - 10
    End If
    End Sub

Private Sub upd_Click()
If MsgBox("Search for Latest Updates ?", vbYesNo) = vbYes Then
If IsConnected = False Then MsgBox ("You need to be connected to the internet to use the update feature."): Exit Sub
'This function assume files "application.ver", "news.txt" and "application.zip"
'on server http://server.com/user (change "server.com/user" by your server name and path)
'Inspect contain of files "news.txt" and "application.ver" at examples
Dim version As String, News As String
    On Error GoTo ErrorMessage
    update.Navigate ("http://sudeepn.homepage.com/qn/update/application.html")
    'You can try this function on Your local disk, but You must change adresses:
    'for example: "file://c:\path\application.ver"
    
skip:
    Exit Sub
ErrorMessage:
    MsgBox "Could not contact server." & Chr(10) & "You must download new version of this application manually at http://sudeepn.homepage.com/qn", vbCritical

End If
End Sub

Private Sub update_DownloadComplete()
    version = update.Document.documentElement.innerHTML
If version = "" Then GoTo skip 'if file not found or file is empty then exit
    If version <= App.Major & "." & App.Minor Then
        MsgBox "No newer version was released.", vbInformation
        GoTo skip
    End If
    'now display MessageBox with news in newer version(s) of application and two buttons Yes(update), No(end)
    upn.Navigate ("http://sudeepn.homepage.com/qn/update/news.txt")
skip: Exit Sub
End Sub

Private Sub upn_DownloadComplete()
News = upn.Document.documentElement.innerHTML
    If MsgBox(Mid(News, 1, InStr(1, News, App.Major & "." & App.Minor) - 9), vbYesNo, "You can update from version " & App.Major & "." & App.Minor & " to version " & version) = vbYes Then
        HyperJump "http://sudeepn.homepage.com/update/application.exe" 'this will run default download manager (probable also open default browser)
    End If
End Sub

Private Sub vf_Click()
TabStrip1.Width = 9900
TabStrip1.Left = 2000
WebBrowser1.Left = 2020
WebBrowser2.Left = 2020
WebBrowser3.Left = 2020
WebBrowser4.Left = 2020
WebBrowser5.Left = 2020
WebBrowser6.Left = 2020
WebBrowser7.Left = 2020
WebBrowser8.Left = 2020
WebBrowser9.Left = 2020
WebBrowser10.Left = 2020
Text1.Left = 2020
Text1.Width = 9900
WebBrowser1.Width = 9900
WebBrowser2.Width = 9900
WebBrowser3.Width = 9900
WebBrowser4.Width = 9900
WebBrowser5.Width = 9900
WebBrowser6.Width = 9900
WebBrowser7.Width = 9900
WebBrowser8.Width = 9900
WebBrowser9.Width = 9900
WebBrowser10.Width = 9900
treeFavorites.Visible = True
Picture1.Visible = True
Command4.Visible = True


End Sub

Private Sub vl_Click()
GoldButton5_Click
End Sub

Private Sub WebBrowser1_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, FLAGS As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)




Dim strURL As String
strURL = URL
Dim bFound As Boolean
Dim i As Integer


If Not bFound Then
cboURL.AddItem strURL
End If

cboURL.Text = strURL


End Sub


Private Sub c_d()
If rty = 1 Then
While a <> Len(WebBrowser1.LocationURL)
a = a + 1
d = Mid(WebBrowser1.LocationURL, a, 1)
If d <> "." And d <> "/" And d <> "_" And d <> "?" Then
tot = tot + d
End If
Wend
valuet = (Int((99 * Rnd) + 11))
tot = "QN" + Right(tot, 3) + valuet
MsgBox (tot)
End If
End Sub

Private Sub WebBrowser1_DownloadBegin()
ProgressBar1.Visible = True
End Sub

Private Sub WebBrowser1_DownloadComplete()
ProgressBar1.Value = 100
ProgressBar1.Visible = False
End Sub

Private Sub WebBrowser1_NewWindow2(ppDisp As Object, Cancel As Boolean)
If dispop.Checked = True Then Cancel = True: Exit Sub
If WebBrowser1.LocationURL = "about:blank" Or WebBrowser1.LocationURL = "" Then
On Error GoTo yel1
TabStrip1.Tabs.Add(2, "sud11", "Browser 1") = "Browser 1"
yel1:
Set ppDisp = WebBrowser1.object
GoTo 66
End If
If WebBrowser2.LocationURL = "about:blank" Or WebBrowser2.LocationURL = "" Then
On Error GoTo yel2
TabStrip1.Tabs.Add(3, "sud22", "Browser 2") = "Browser 2"
yel2:
Set ppDisp = WebBrowser2.object
GoTo 66
End If
If WebBrowser3.LocationURL = "about:blank" Or WebBrowser3.LocationURL = "" Then
On Error GoTo yel3
TabStrip1.Tabs.Add(4, "sud33", "Browser 3") = "Browser 3"
yel3:
Set ppDisp = WebBrowser3.object
GoTo 66
End If

If WebBrowser4.LocationURL = "about:blank" Or WebBrowser4.LocationURL = "" Then
On Error GoTo yel4
TabStrip1.Tabs.Add(5, "sud44", "Browser 44") = "Browser 4"
yel4:
Set ppDisp = WebBrowser4.object
GoTo 66
End If
If WebBrowser5.LocationURL = "about:blank" Or WebBrowser5.LocationURL = "" Then
On Error GoTo yel5
TabStrip1.Tabs.Add(6, "sud55", "Browser 55") = "Browser 5"
yel5:
Set ppDisp = WebBrowser5.object
GoTo 66
End If
If WebBrowser6.LocationURL = "about:blank" Or WebBrowser6.LocationURL = "" Then
On Error GoTo yel6
TabStrip1.Tabs.Add(7, "sud66", "Browser 6") = "Browser 6"
yel6:
Set ppDisp = WebBrowser6.object
GoTo 66
End If
If WebBrowser7.LocationURL = "about:blank" Or WebBrowser7.LocationURL = "" Then
On Error GoTo yel7
TabStrip1.Tabs.Add(8, "sud77", "Browser 7") = "Browser 7"
yel7:
Set ppDisp = WebBrowser7.object
GoTo 66
End If
If WebBrowser8.LocationURL = "about:blank" Or WebBrowser8.LocationURL = "" Then
On Error GoTo yel8
TabStrip1.Tabs.Add(9, "sud88", "Browser 8") = "Browser 8"
yel8:
Set ppDisp = WebBrowser8.object
GoTo 66
End If
If WebBrowser9.LocationURL = "about:blank" Or WebBrowser9.LocationURL = "" Then
On Error GoTo yel9
TabStrip1.Tabs.Add(10, "sud99", "Browser 9") = "Browser 9"
yel9:
Set ppDisp = WebBrowser9.object
GoTo 66
End If
If WebBrowser10.LocationURL = "about:blank" Or WebBrowser10.LocationURL = "" Then
On Error GoTo yel10
TabStrip1.Tabs.Add(11, "sud10", "Browser 10") = "Browser 10"
yel10:
Set ppDisp = WebBrowser10.object
GoTo 66
End If
MsgBox ("All Browsers are in use. To close an active unwanted page click 'CLEAR' above.")
MsgBox ("The page will be opened in Internet Explorer")
66: TabStrip1.Refresh

End Sub

Private Sub WebBrowser1_ProgressChange(ByVal progress As Long, ByVal ProgressMax As Long)
On Error GoTo progressERR
If progress = -1 Then ProgressBar1.Value = 100
If progress > 0 And ProgressMax > 0 Then
    ProgressBar1.Value = progress * 100 / ProgressMax
    End If
    Exit Sub
progressERR:
End Sub

Private Sub WebBrowser1_StatusTextChange(ByVal Text As String)
If TabStrip1.SelectedItem.Key = "sud11" Then Label2.Caption = Text
End Sub

Private Sub WebBrowser1_TitleChange(ByVal Text As String)
For a = 1 To TabStrip1.Tabs.Count
If TabStrip1.Tabs(a).Key = "sud11" Then TabStrip1.Tabs(a).Caption = Left(Text, 15)
Next a

End Sub

Private Sub webBrowser10_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, FLAGS As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Dim strURL As String
strURL = URL
Dim bFound As Boolean
Dim i As Integer


If Not bFound Then
cboURL.AddItem strURL
End If

cboURL.Text = strURL
End Sub

Private Sub webBrowser10_DownloadBegin()
ProgressBar1.Visible = True
End Sub

Private Sub webBrowser10_DownloadComplete()
ProgressBar10.Value = 100
ProgressBar10.Visible = False

End Sub

Private Sub webBrowser10_NewWindow2(ppDisp As Object, Cancel As Boolean)
If dispop.Checked = True Then Cancel = True: Exit Sub
If WebBrowser1.LocationURL = "about:blank" Or WebBrowser1.LocationURL = "" Then
On Error GoTo yel1
TabStrip1.Tabs.Add(2, "sud11", "Browser 1") = "Browser 1"
yel1:
Set ppDisp = WebBrowser1.object
GoTo 667
End If
If WebBrowser2.LocationURL = "about:blank" Or WebBrowser2.LocationURL = "" Then
On Error GoTo yel2
TabStrip1.Tabs.Add(3, "sud22", "Browser 2") = "Browser 2"
yel2:
Set ppDisp = WebBrowser2.object
GoTo 667
End If
If WebBrowser3.LocationURL = "about:blank" Or WebBrowser3.LocationURL = "" Then
On Error GoTo yel3
TabStrip1.Tabs.Add(4, "sud33", "Browser 3") = "Browser 3"
yel3:
Set ppDisp = WebBrowser3.object
GoTo 667
End If

If WebBrowser4.LocationURL = "about:blank" Or WebBrowser4.LocationURL = "" Then
On Error GoTo yel4
TabStrip1.Tabs.Add(5, "sud44", "Browser 44") = "Browser 4"
yel4:
Set ppDisp = WebBrowser4.object
GoTo 667
End If
If WebBrowser5.LocationURL = "about:blank" Or WebBrowser5.LocationURL = "" Then
On Error GoTo yel5
TabStrip1.Tabs.Add(6, "sud55", "Browser 55") = "Browser 5"
yel5:
Set ppDisp = WebBrowser5.object
GoTo 667
End If
If WebBrowser6.LocationURL = "about:blank" Or WebBrowser6.LocationURL = "" Then
On Error GoTo yel6
TabStrip1.Tabs.Add(7, "sud66", "Browser 6") = "Browser 6"
yel6:
Set ppDisp = WebBrowser6.object
GoTo 667
End If
If WebBrowser7.LocationURL = "about:blank" Or WebBrowser7.LocationURL = "" Then
On Error GoTo yel7
TabStrip1.Tabs.Add(8, "sud77", "Browser 7") = "Browser 7"
yel7:
Set ppDisp = WebBrowser7.object
GoTo 667
End If
If WebBrowser8.LocationURL = "about:blank" Or WebBrowser8.LocationURL = "" Then
On Error GoTo yel8
TabStrip1.Tabs.Add(9, "sud88", "Browser 8") = "Browser 8"
yel8:
Set ppDisp = WebBrowser8.object
GoTo 667
End If
If WebBrowser9.LocationURL = "about:blank" Or WebBrowser9.LocationURL = "" Then
On Error GoTo yel9
TabStrip1.Tabs.Add(10, "sud99", "Browser 9") = "Browser 9"
yel9:
Set ppDisp = WebBrowser9.object
GoTo 667
End If
If WebBrowser10.LocationURL = "about:blank" Or WebBrowser10.LocationURL = "" Then
On Error GoTo yel10
TabStrip1.Tabs.Add(11, "sud10", "Browser 10") = "Browser 10"
yel10:
Set ppDisp = WebBrowser10.object
GoTo 667
End If
MsgBox ("All Browsers are in use. To close an active unwanted page click 'CLEAR' above.")
MsgBox ("The page will be opened in Internet Explorer")
667: TabStrip1.Refresh

End Sub

Private Sub webBrowser10_ProgressChange(ByVal progress As Long, ByVal ProgressMax As Long)
On Error GoTo progressERR
If progress = -1 Then ProgressBar10.Value = 100
If progress > 0 And ProgressMax > 0 Then
    ProgressBar10.Value = progress * 100 / ProgressMax
    End If
    Exit Sub
progressERR:
               
End Sub

Private Sub webBrowser10_StatusTextChange(ByVal Text As String)
If TabStrip1.SelectedItem.Key = "sud10" Then Label2.Caption = Text
End Sub

Private Sub WebBrowser10_TitleChange(ByVal Text As String)
For a = 1 To TabStrip1.Tabs.Count
If TabStrip1.Tabs(a).Key = "sud10" Then TabStrip1.Tabs(a).Caption = Left(Text, 15)
Next a

End Sub

Private Sub WebBrowser2_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, FLAGS As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)
Dim strURL As String
strURL = URL
Dim bFound As Boolean
Dim i As Integer


If Not bFound Then
cboURL.AddItem strURL
End If

cboURL.Text = strURL
End Sub

Private Sub WebBrowser2_DownloadBegin()
ProgressBar1.Visible = True
End Sub

Private Sub WebBrowser2_DownloadComplete()
ProgressBar2.Value = 100
ProgressBar2.Visible = False

End Sub

Private Sub WebBrowser2_NewWindow2(ppDisp As Object, Cancel As Boolean)
If dispop.Checked = True Then Cancel = True: Exit Sub
If WebBrowser1.LocationURL = "about:blank" Or WebBrowser1.LocationURL = "" Then
On Error GoTo yel1
TabStrip1.Tabs.Add(2, "sud11", "Browser 1") = "Browser 1"
yel1:
Set ppDisp = WebBrowser1.object
GoTo 668
End If
If WebBrowser2.LocationURL = "about:blank" Or WebBrowser2.LocationURL = "" Then
On Error GoTo yel2
TabStrip1.Tabs.Add(3, "sud22", "Browser 2") = "Browser 2"
yel2:
Set ppDisp = WebBrowser2.object
GoTo 668
End If
If WebBrowser3.LocationURL = "about:blank" Or WebBrowser3.LocationURL = "" Then
On Error GoTo yel3
TabStrip1.Tabs.Add(4, "sud33", "Browser 3") = "Browser 3"
yel3:
Set ppDisp = WebBrowser3.object
GoTo 668
End If

If WebBrowser4.LocationURL = "about:blank" Or WebBrowser4.LocationURL = "" Then
On Error GoTo yel4
TabStrip1.Tabs.Add(5, "sud44", "Browser 44") = "Browser 4"
yel4:
Set ppDisp = WebBrowser4.object
GoTo 668
End If
If WebBrowser5.LocationURL = "about:blank" Or WebBrowser5.LocationURL = "" Then
On Error GoTo yel5
TabStrip1.Tabs.Add(6, "sud55", "Browser 55") = "Browser 5"
yel5:
Set ppDisp = WebBrowser5.object
GoTo 668
End If
If WebBrowser6.LocationURL = "about:blank" Or WebBrowser6.LocationURL = "" Then
On Error GoTo yel6
TabStrip1.Tabs.Add(7, "sud66", "Browser 6") = "Browser 6"
yel6:
Set ppDisp = WebBrowser6.object
GoTo 668
End If
If WebBrowser7.LocationURL = "about:blank" Or WebBrowser7.LocationURL = "" Then
On Error GoTo yel7
TabStrip1.Tabs.Add(8, "sud77", "Browser 7") = "Browser 7"
yel7:
Set ppDisp = WebBrowser7.object
GoTo 668
End If
If WebBrowser8.LocationURL = "about:blank" Or WebBrowser8.LocationURL = "" Then
On Error GoTo yel8
TabStrip1.Tabs.Add(9, "sud88", "Browser 8") = "Browser 8"
yel8:
Set ppDisp = WebBrowser8.object
GoTo 668
End If
If WebBrowser9.LocationURL = "about:blank" Or WebBrowser9.LocationURL = "" Then
On Error GoTo yel9
TabStrip1.Tabs.Add(10, "sud99", "Browser 9") = "Browser 9"
yel9:
Set ppDisp = WebBrowser9.object
GoTo 668
End If
If WebBrowser10.LocationURL = "about:blank" Or WebBrowser10.LocationURL = "" Then
On Error GoTo yel10
TabStrip1.Tabs.Add(11, "sud10", "Browser 10") = "Browser 10"
yel10:
Set ppDisp = WebBrowser10.object
GoTo 668
End If
MsgBox ("All Browsers are in use. To close an active unwanted page click 'CLEAR' above.")
MsgBox ("The page will be opened in Internet Explorer")
668: TabStrip1.Refresh

End Sub

Private Sub WebBrowser2_ProgressChange(ByVal progress As Long, ByVal ProgressMax As Long)
On Error GoTo progressERR
If progress = -1 Then ProgressBar2.Value = 100
If progress > 0 And ProgressMax > 0 Then
    ProgressBar2.Value = progress * 100 / ProgressMax
    End If
    Exit Sub
progressERR:
               

End Sub



Private Sub WebBrowser2_StatusTextChange(ByVal Text As String)
If TabStrip1.SelectedItem.Key = "sud22" Then Label2.Caption = Text
End Sub

Private Sub WebBrowser2_TitleChange(ByVal Text As String)
For a = 1 To TabStrip1.Tabs.Count
If TabStrip1.Tabs(a).Key = "sud22" Then TabStrip1.Tabs(a).Caption = Left(Text, 15)
Next a

End Sub

Private Sub WebBrowser3_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, FLAGS As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

Dim strURL As String
strURL = URL
Dim bFound As Boolean
Dim i As Integer


If Not bFound Then
cboURL.AddItem strURL
End If

cboURL.Text = strURL



End Sub

Private Sub WebBrowser3_DownloadBegin()
ProgressBar1.Visible = True
End Sub

Private Sub WebBrowser3_DownloadComplete()
ProgressBar3.Value = 100
ProgressBar3.Visible = False

End Sub

Private Sub WebBrowser3_NewWindow2(ppDisp As Object, Cancel As Boolean)
If dispop.Checked = True Then Cancel = True: Exit Sub
If WebBrowser1.LocationURL = "about:blank" Or WebBrowser1.LocationURL = "" Then
On Error GoTo yel1
TabStrip1.Tabs.Add(2, "sud11", "Browser 1") = "Browser 1"
yel1:
Set ppDisp = WebBrowser1.object
GoTo 669
End If
If WebBrowser2.LocationURL = "about:blank" Or WebBrowser2.LocationURL = "" Then
On Error GoTo yel2
TabStrip1.Tabs.Add(3, "sud22", "Browser 2") = "Browser 2"
yel2:
Set ppDisp = WebBrowser2.object
GoTo 669
End If
If WebBrowser3.LocationURL = "about:blank" Or WebBrowser3.LocationURL = "" Then
On Error GoTo yel3
TabStrip1.Tabs.Add(4, "sud33", "Browser 3") = "Browser 3"
yel3:
Set ppDisp = WebBrowser3.object
GoTo 669
End If

If WebBrowser4.LocationURL = "about:blank" Or WebBrowser4.LocationURL = "" Then
On Error GoTo yel4
TabStrip1.Tabs.Add(5, "sud44", "Browser 44") = "Browser 4"
yel4:
Set ppDisp = WebBrowser4.object
GoTo 669
End If
If WebBrowser5.LocationURL = "about:blank" Or WebBrowser5.LocationURL = "" Then
On Error GoTo yel5
TabStrip1.Tabs.Add(6, "sud55", "Browser 55") = "Browser 5"
yel5:
Set ppDisp = WebBrowser5.object
GoTo 669
End If
If WebBrowser6.LocationURL = "about:blank" Or WebBrowser6.LocationURL = "" Then
On Error GoTo yel6
TabStrip1.Tabs.Add(7, "sud66", "Browser 6") = "Browser 6"
yel6:
Set ppDisp = WebBrowser6.object
GoTo 669
End If
If WebBrowser7.LocationURL = "about:blank" Or WebBrowser7.LocationURL = "" Then
On Error GoTo yel7
TabStrip1.Tabs.Add(8, "sud77", "Browser 7") = "Browser 7"
yel7:
Set ppDisp = WebBrowser7.object
GoTo 669
End If
If WebBrowser8.LocationURL = "about:blank" Or WebBrowser8.LocationURL = "" Then
On Error GoTo yel8
TabStrip1.Tabs.Add(9, "sud88", "Browser 8") = "Browser 8"
yel8:
Set ppDisp = WebBrowser8.object
GoTo 669
End If
If WebBrowser9.LocationURL = "about:blank" Or WebBrowser9.LocationURL = "" Then
On Error GoTo yel9
TabStrip1.Tabs.Add(10, "sud99", "Browser 9") = "Browser 9"
yel9:
Set ppDisp = WebBrowser9.object
GoTo 669
End If
If WebBrowser10.LocationURL = "about:blank" Or WebBrowser10.LocationURL = "" Then
On Error GoTo yel10
TabStrip1.Tabs.Add(11, "sud10", "Browser 10") = "Browser 10"
yel10:
Set ppDisp = WebBrowser10.object
GoTo 669
End If
MsgBox ("All Browsers are in use. To close an active unwanted page click 'CLEAR' above.")
MsgBox ("The page will be opened in Internet Explorer")
669: TabStrip1.Refresh
 
 End Sub

Private Sub WebBrowser3_ProgressChange(ByVal progress As Long, ByVal ProgressMax As Long)
On Error GoTo progressERR
If progress = -1 Then ProgressBar3.Value = 100
If progress > 0 And ProgressMax > 0 Then
    ProgressBar3.Value = progress * 100 / ProgressMax
    End If
    Exit Sub
progressERR:
               
End Sub

Private Sub WebBrowser3_StatusTextChange(ByVal Text As String)
If TabStrip1.SelectedItem.Key = "sud33" Then Label2.Caption = Text
End Sub

Private Sub WebBrowser3_TitleChange(ByVal Text As String)
For a = 1 To TabStrip1.Tabs.Count
If TabStrip1.Tabs(a).Key = "sud33" Then TabStrip1.Tabs(a).Caption = Left(Text, 15)
Next a

End Sub

Private Sub WebBrowser4_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, FLAGS As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

Dim strURL As String
strURL = URL
Dim bFound As Boolean
Dim i As Integer


If Not bFound Then
cboURL.AddItem strURL
End If

cboURL.Text = strURL



End Sub

Private Sub WebBrowser4_DownloadBegin()
ProgressBar1.Visible = True
End Sub

Private Sub WebBrowser4_DownloadComplete()
ProgressBar4.Value = 100
ProgressBar4.Visible = False
End Sub

Private Sub WebBrowser4_NewWindow2(ppDisp As Object, Cancel As Boolean)
If dispop.Checked = True Then Cancel = True: Exit Sub
If WebBrowser1.LocationURL = "about:blank" Or WebBrowser1.LocationURL = "" Then
On Error GoTo yel1
TabStrip1.Tabs.Add(2, "sud11", "Browser 1") = "Browser 1"
yel1:
Set ppDisp = WebBrowser1.object
GoTo 6691
End If
If WebBrowser2.LocationURL = "about:blank" Or WebBrowser2.LocationURL = "" Then
On Error GoTo yel2
TabStrip1.Tabs.Add(3, "sud22", "Browser 2") = "Browser 2"
yel2:
Set ppDisp = WebBrowser2.object
GoTo 6691
End If
If WebBrowser3.LocationURL = "about:blank" Or WebBrowser3.LocationURL = "" Then
On Error GoTo yel3
TabStrip1.Tabs.Add(4, "sud33", "Browser 3") = "Browser 3"
yel3:
Set ppDisp = WebBrowser3.object
GoTo 6691
End If

If WebBrowser4.LocationURL = "about:blank" Or WebBrowser4.LocationURL = "" Then
On Error GoTo yel4
TabStrip1.Tabs.Add(5, "sud44", "Browser 44") = "Browser 4"
yel4:
Set ppDisp = WebBrowser4.object
GoTo 6691
End If
If WebBrowser5.LocationURL = "about:blank" Or WebBrowser5.LocationURL = "" Then
On Error GoTo yel5
TabStrip1.Tabs.Add(6, "sud55", "Browser 55") = "Browser 5"
yel5:
Set ppDisp = WebBrowser5.object
GoTo 6691
End If
If WebBrowser6.LocationURL = "about:blank" Or WebBrowser6.LocationURL = "" Then
On Error GoTo yel6
TabStrip1.Tabs.Add(7, "sud66", "Browser 6") = "Browser 6"
yel6:
Set ppDisp = WebBrowser6.object
GoTo 6691
End If
If WebBrowser7.LocationURL = "about:blank" Or WebBrowser7.LocationURL = "" Then
On Error GoTo yel7
TabStrip1.Tabs.Add(8, "sud77", "Browser 7") = "Browser 7"
yel7:
Set ppDisp = WebBrowser7.object
GoTo 6691
End If
If WebBrowser8.LocationURL = "about:blank" Or WebBrowser8.LocationURL = "" Then
On Error GoTo yel8
TabStrip1.Tabs.Add(9, "sud88", "Browser 8") = "Browser 8"
yel8:
Set ppDisp = WebBrowser8.object
GoTo 6691
End If
If WebBrowser9.LocationURL = "about:blank" Or WebBrowser9.LocationURL = "" Then
On Error GoTo yel9
TabStrip1.Tabs.Add(10, "sud99", "Browser 9") = "Browser 9"
yel9:
Set ppDisp = WebBrowser9.object
GoTo 6691
End If
If WebBrowser10.LocationURL = "about:blank" Or WebBrowser10.LocationURL = "" Then
On Error GoTo yel10
TabStrip1.Tabs.Add(11, "sud10", "Browser 10") = "Browser 10"
yel10:
Set ppDisp = WebBrowser10.object
GoTo 6691
End If
MsgBox ("All Browsers are in use. To close an active unwanted page click 'CLEAR' above.")
MsgBox ("The page will be opened in Internet Explorer")
6691: TabStrip1.Refresh
 
 End Sub

Private Sub WebBrowser4_ProgressChange(ByVal progress As Long, ByVal ProgressMax As Long)
On Error GoTo progressERR
If progress = -1 Then ProgressBar4.Value = 100
If progress > 0 And ProgressMax > 0 Then
    ProgressBar4.Value = progress * 100 / ProgressMax
    End If
    Exit Sub
progressERR:
               
End Sub

Private Sub WebBrowser4_StatusTextChange(ByVal Text As String)
If TabStrip1.SelectedItem.Key = "sud44" Then Label2.Caption = Text
End Sub

Private Sub WebBrowser4_TitleChange(ByVal Text As String)
For a = 1 To TabStrip1.Tabs.Count
If TabStrip1.Tabs(a).Key = "sud44" Then TabStrip1.Tabs(a).Caption = Left(Text, 15)
Next a

End Sub

Private Sub WebBrowser5_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, FLAGS As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

Dim strURL As String
strURL = URL
Dim bFound As Boolean
Dim i As Integer


If Not bFound Then
cboURL.AddItem strURL
End If

cboURL.Text = strURL



End Sub

Private Sub WebBrowser5_DownloadBegin()
ProgressBar1.Visible = True
End Sub

Private Sub WebBrowser5_DownloadComplete()
ProgressBar5.Value = 100
ProgressBar5.Visible = False

End Sub

Private Sub WebBrowser5_NewWindow2(ppDisp As Object, Cancel As Boolean)
If dispop.Checked = True Then Cancel = True: Exit Sub
If WebBrowser1.LocationURL = "about:blank" Or WebBrowser1.LocationURL = "" Then
On Error GoTo yel1
TabStrip1.Tabs.Add(2, "sud11", "Browser 1") = "Browser 1"
yel1:
Set ppDisp = WebBrowser1.object
GoTo 6692
End If
If WebBrowser2.LocationURL = "about:blank" Or WebBrowser2.LocationURL = "" Then
On Error GoTo yel2
TabStrip1.Tabs.Add(3, "sud22", "Browser 2") = "Browser 2"
yel2:
Set ppDisp = WebBrowser2.object
GoTo 6692
End If
If WebBrowser3.LocationURL = "about:blank" Or WebBrowser3.LocationURL = "" Then
On Error GoTo yel3
TabStrip1.Tabs.Add(4, "sud33", "Browser 3") = "Browser 3"
yel3:
Set ppDisp = WebBrowser3.object
GoTo 6692
End If

If WebBrowser4.LocationURL = "about:blank" Or WebBrowser4.LocationURL = "" Then
On Error GoTo yel4
TabStrip1.Tabs.Add(5, "sud44", "Browser 44") = "Browser 4"
yel4:
Set ppDisp = WebBrowser4.object
GoTo 6692
End If
If WebBrowser5.LocationURL = "about:blank" Or WebBrowser5.LocationURL = "" Then
On Error GoTo yel5
TabStrip1.Tabs.Add(6, "sud55", "Browser 55") = "Browser 5"
yel5:
Set ppDisp = WebBrowser5.object
GoTo 6692
End If
If WebBrowser6.LocationURL = "about:blank" Or WebBrowser6.LocationURL = "" Then
On Error GoTo yel6
TabStrip1.Tabs.Add(7, "sud66", "Browser 6") = "Browser 6"
yel6:
Set ppDisp = WebBrowser6.object
GoTo 6692
End If
If WebBrowser7.LocationURL = "about:blank" Or WebBrowser7.LocationURL = "" Then
On Error GoTo yel7
TabStrip1.Tabs.Add(8, "sud77", "Browser 7") = "Browser 7"
yel7:
Set ppDisp = WebBrowser7.object
GoTo 6692
End If
If WebBrowser8.LocationURL = "about:blank" Or WebBrowser8.LocationURL = "" Then
On Error GoTo yel8
TabStrip1.Tabs.Add(9, "sud88", "Browser 8") = "Browser 8"
yel8:
Set ppDisp = WebBrowser8.object
GoTo 6692
End If
If WebBrowser9.LocationURL = "about:blank" Or WebBrowser9.LocationURL = "" Then
On Error GoTo yel9
TabStrip1.Tabs.Add(10, "sud99", "Browser 9") = "Browser 9"
yel9:
Set ppDisp = WebBrowser9.object
GoTo 6692
End If
If WebBrowser10.LocationURL = "about:blank" Or WebBrowser10.LocationURL = "" Then
On Error GoTo yel10
TabStrip1.Tabs.Add(11, "sud10", "Browser 10") = "Browser 10"
yel10:
Set ppDisp = WebBrowser10.object
GoTo 6692
End If
MsgBox ("All Browsers are in use. To close an active unwanted page click 'CLEAR' above.")
MsgBox ("The page will be opened in Internet Explorer")
6692: TabStrip1.Refresh
 
 End Sub

Private Sub WebBrowser5_ProgressChange(ByVal progress As Long, ByVal ProgressMax As Long)
On Error GoTo progressERR
If progress = -1 Then ProgressBar5.Value = 100
If progress > 0 And ProgressMax > 0 Then
    ProgressBar5.Value = progress * 100 / ProgressMax
    End If
    Exit Sub
progressERR:
               

End Sub

Private Sub WebBrowser5_StatusTextChange(ByVal Text As String)
If TabStrip1.SelectedItem.Key = "sud55" Then Label2.Caption = Text
End Sub

Private Sub WebBrowser5_TitleChange(ByVal Text As String)
For a = 1 To TabStrip1.Tabs.Count
If TabStrip1.Tabs(a).Key = "sud55" Then TabStrip1.Tabs(a).Caption = Left(Text, 15)
Next a

End Sub

Private Sub WebBrowser6_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, FLAGS As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

Dim strURL As String
strURL = URL
Dim bFound As Boolean
Dim i As Integer


If Not bFound Then
cboURL.AddItem strURL
End If

cboURL.Text = strURL



End Sub

Private Sub WebBrowser6_DownloadBegin()
ProgressBar1.Visible = True
End Sub

Private Sub WebBrowser6_DownloadComplete()
ProgressBar6.Value = 100
ProgressBar6.Visible = False
End Sub

Private Sub WebBrowser6_NewWindow2(ppDisp As Object, Cancel As Boolean)
If dispop.Checked = True Then Cancel = True: Exit Sub
If WebBrowser1.LocationURL = "about:blank" Or WebBrowser1.LocationURL = "" Then
On Error GoTo yel1
TabStrip1.Tabs.Add(2, "sud11", "Browser 1") = "Browser 1"
yel1:
Set ppDisp = WebBrowser1.object
GoTo 6693
End If
If WebBrowser2.LocationURL = "about:blank" Or WebBrowser2.LocationURL = "" Then
On Error GoTo yel2
TabStrip1.Tabs.Add(3, "sud22", "Browser 2") = "Browser 2"
yel2:
Set ppDisp = WebBrowser2.object
GoTo 6693
End If
If WebBrowser3.LocationURL = "about:blank" Or WebBrowser3.LocationURL = "" Then
On Error GoTo yel3
TabStrip1.Tabs.Add(4, "sud33", "Browser 3") = "Browser 3"
yel3:
Set ppDisp = WebBrowser3.object
GoTo 6693
End If

If WebBrowser4.LocationURL = "about:blank" Or WebBrowser4.LocationURL = "" Then
On Error GoTo yel4
TabStrip1.Tabs.Add(5, "sud44", "Browser 44") = "Browser 4"
yel4:
Set ppDisp = WebBrowser4.object
GoTo 6693
End If
If WebBrowser5.LocationURL = "about:blank" Or WebBrowser5.LocationURL = "" Then
On Error GoTo yel5
TabStrip1.Tabs.Add(6, "sud55", "Browser 55") = "Browser 5"
yel5:
Set ppDisp = WebBrowser5.object
GoTo 6693
End If
If WebBrowser6.LocationURL = "about:blank" Or WebBrowser6.LocationURL = "" Then
On Error GoTo yel6
TabStrip1.Tabs.Add(7, "sud66", "Browser 6") = "Browser 6"
yel6:
Set ppDisp = WebBrowser6.object
GoTo 6693
End If
If WebBrowser7.LocationURL = "about:blank" Or WebBrowser7.LocationURL = "" Then
On Error GoTo yel7
TabStrip1.Tabs.Add(8, "sud77", "Browser 7") = "Browser 7"
yel7:
Set ppDisp = WebBrowser7.object
GoTo 6693
End If
If WebBrowser8.LocationURL = "about:blank" Or WebBrowser8.LocationURL = "" Then
On Error GoTo yel8
TabStrip1.Tabs.Add(9, "sud88", "Browser 8") = "Browser 8"
yel8:
Set ppDisp = WebBrowser8.object
GoTo 6693
End If
If WebBrowser9.LocationURL = "about:blank" Or WebBrowser9.LocationURL = "" Then
On Error GoTo yel9
TabStrip1.Tabs.Add(10, "sud99", "Browser 9") = "Browser 9"
yel9:
Set ppDisp = WebBrowser9.object
GoTo 6693
End If
If WebBrowser10.LocationURL = "about:blank" Or WebBrowser10.LocationURL = "" Then
On Error GoTo yel10
TabStrip1.Tabs.Add(11, "sud10", "Browser 10") = "Browser 10"
yel10:
Set ppDisp = WebBrowser10.object
GoTo 6693
End If
MsgBox ("All Browsers are in use. To close an active unwanted page click 'CLEAR' above.")
MsgBox ("The page will be opened in Internet Explorer")
6693: TabStrip1.Refresh
 
 End Sub

Private Sub WebBrowser6_ProgressChange(ByVal progress As Long, ByVal ProgressMax As Long)
On Error GoTo progressERR
If progress = -1 Then ProgressBar6.Value = 100
If progress > 0 And ProgressMax > 0 Then
    ProgressBar6.Value = progress * 100 / ProgressMax
    End If
    Exit Sub
progressERR:
               
End Sub

Private Sub WebBrowser6_StatusTextChange(ByVal Text As String)
If TabStrip1.SelectedItem.Key = "sud66" Then Label2.Caption = Text
End Sub

Private Sub WebBrowser6_TitleChange(ByVal Text As String)
For a = 1 To TabStrip1.Tabs.Count
If TabStrip1.Tabs(a).Key = "sud66" Then TabStrip1.Tabs(a).Caption = Left(Text, 15)
Next a

End Sub

Private Sub WebBrowser7_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, FLAGS As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

Dim strURL As String
strURL = URL
Dim bFound As Boolean
Dim i As Integer


If Not bFound Then
cboURL.AddItem strURL
End If

cboURL.Text = strURL



End Sub

Private Sub WebBrowser7_DownloadBegin()
ProgressBar1.Visible = True
End Sub

Private Sub WebBrowser7_DownloadComplete()
ProgressBar7.Value = 100
ProgressBar7.Visible = False

End Sub

Private Sub WebBrowser7_NewWindow2(ppDisp As Object, Cancel As Boolean)
If dispop.Checked = True Then Cancel = True: Exit Sub
If WebBrowser1.LocationURL = "about:blank" Or WebBrowser1.LocationURL = "" Then
On Error GoTo yel1
TabStrip1.Tabs.Add(2, "sud11", "Browser 1") = "Browser 1"
yel1:
Set ppDisp = WebBrowser1.object
GoTo 6694
End If
If WebBrowser2.LocationURL = "about:blank" Or WebBrowser2.LocationURL = "" Then
On Error GoTo yel2
TabStrip1.Tabs.Add(3, "sud22", "Browser 2") = "Browser 2"
yel2:
Set ppDisp = WebBrowser2.object
GoTo 6694
End If
If WebBrowser3.LocationURL = "about:blank" Or WebBrowser3.LocationURL = "" Then
On Error GoTo yel3
TabStrip1.Tabs.Add(4, "sud33", "Browser 3") = "Browser 3"
yel3:
Set ppDisp = WebBrowser3.object
GoTo 6694
End If

If WebBrowser4.LocationURL = "about:blank" Or WebBrowser4.LocationURL = "" Then
On Error GoTo yel4
TabStrip1.Tabs.Add(5, "sud44", "Browser 44") = "Browser 4"
yel4:
Set ppDisp = WebBrowser4.object
GoTo 6694
End If
If WebBrowser5.LocationURL = "about:blank" Or WebBrowser5.LocationURL = "" Then
On Error GoTo yel5
TabStrip1.Tabs.Add(6, "sud55", "Browser 55") = "Browser 5"
yel5:
Set ppDisp = WebBrowser5.object
GoTo 6694
End If
If WebBrowser6.LocationURL = "about:blank" Or WebBrowser6.LocationURL = "" Then
On Error GoTo yel6
TabStrip1.Tabs.Add(7, "sud66", "Browser 6") = "Browser 6"
yel6:
Set ppDisp = WebBrowser6.object
GoTo 6694
End If
If WebBrowser7.LocationURL = "about:blank" Or WebBrowser7.LocationURL = "" Then
On Error GoTo yel7
TabStrip1.Tabs.Add(8, "sud77", "Browser 7") = "Browser 7"
yel7:
Set ppDisp = WebBrowser7.object
GoTo 6694
End If
If WebBrowser8.LocationURL = "about:blank" Or WebBrowser8.LocationURL = "" Then
On Error GoTo yel8
TabStrip1.Tabs.Add(9, "sud88", "Browser 8") = "Browser 8"
yel8:
Set ppDisp = WebBrowser8.object
GoTo 6694
End If
If WebBrowser9.LocationURL = "about:blank" Or WebBrowser9.LocationURL = "" Then
On Error GoTo yel9
TabStrip1.Tabs.Add(10, "sud99", "Browser 9") = "Browser 9"
yel9:
Set ppDisp = WebBrowser9.object
GoTo 6694
End If
If WebBrowser10.LocationURL = "about:blank" Or WebBrowser10.LocationURL = "" Then
On Error GoTo yel10
TabStrip1.Tabs.Add(11, "sud10", "Browser 10") = "Browser 10"
yel10:
Set ppDisp = WebBrowser10.object
GoTo 6694
End If
MsgBox ("All Browsers are in use. To close an active unwanted page click 'CLEAR' above.")
MsgBox ("The page will be opened in Internet Explorer")
6694: TabStrip1.Refresh
 
 End Sub

Private Sub WebBrowser7_ProgressChange(ByVal progress As Long, ByVal ProgressMax As Long)
On Error GoTo progressERR
If progress = -1 Then ProgressBar7.Value = 100
If progress > 0 And ProgressMax > 0 Then
    ProgressBar7.Value = progress * 100 / ProgressMax
    End If
    Exit Sub
progressERR:
               
End Sub

Private Sub WebBrowser7_StatusTextChange(ByVal Text As String)
If TabStrip1.SelectedItem.Key = "sud77" Then Label2.Caption = Text

End Sub

Private Sub WebBrowser7_TitleChange(ByVal Text As String)
For a = 1 To TabStrip1.Tabs.Count
If TabStrip1.Tabs(a).Key = "sud77" Then TabStrip1.Tabs(a).Caption = Left(Text, 15)
Next a

End Sub

Private Sub WebBrowser8_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, FLAGS As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

Dim strURL As String
strURL = URL
Dim bFound As Boolean
Dim i As Integer


If Not bFound Then
cboURL.AddItem strURL
End If

cboURL.Text = strURL



End Sub

Private Sub WebBrowser8_DownloadBegin()
ProgressBar1.Visible = True
End Sub

Private Sub WebBrowser8_DownloadComplete()
ProgressBar8.Value = 100
ProgressBar8.Visible = False

End Sub

Private Sub WebBrowser8_NewWindow2(ppDisp As Object, Cancel As Boolean)
If dispop.Checked = True Then Cancel = True: Exit Sub
If WebBrowser1.LocationURL = "about:blank" Or WebBrowser1.LocationURL = "" Then
On Error GoTo yel1
TabStrip1.Tabs.Add(2, "sud11", "Browser 1") = "Browser 1"
yel1:
Set ppDisp = WebBrowser1.object
GoTo 6695
End If
If WebBrowser2.LocationURL = "about:blank" Or WebBrowser2.LocationURL = "" Then
On Error GoTo yel2
TabStrip1.Tabs.Add(3, "sud22", "Browser 2") = "Browser 2"
yel2:
Set ppDisp = WebBrowser2.object
GoTo 6695
End If
If WebBrowser3.LocationURL = "about:blank" Or WebBrowser3.LocationURL = "" Then
On Error GoTo yel3
TabStrip1.Tabs.Add(4, "sud33", "Browser 3") = "Browser 3"
yel3:
Set ppDisp = WebBrowser3.object
GoTo 6695
End If

If WebBrowser4.LocationURL = "about:blank" Or WebBrowser4.LocationURL = "" Then
On Error GoTo yel4
TabStrip1.Tabs.Add(5, "sud44", "Browser 44") = "Browser 4"
yel4:
Set ppDisp = WebBrowser4.object
GoTo 6695
End If
If WebBrowser5.LocationURL = "about:blank" Or WebBrowser5.LocationURL = "" Then
On Error GoTo yel5
TabStrip1.Tabs.Add(6, "sud55", "Browser 55") = "Browser 5"
yel5:
Set ppDisp = WebBrowser5.object
GoTo 6695
End If
If WebBrowser6.LocationURL = "about:blank" Or WebBrowser6.LocationURL = "" Then
On Error GoTo yel6
TabStrip1.Tabs.Add(7, "sud66", "Browser 6") = "Browser 6"
yel6:
Set ppDisp = WebBrowser6.object
GoTo 6695
End If
If WebBrowser7.LocationURL = "about:blank" Or WebBrowser7.LocationURL = "" Then
On Error GoTo yel7
TabStrip1.Tabs.Add(8, "sud77", "Browser 7") = "Browser 7"
yel7:
Set ppDisp = WebBrowser7.object
GoTo 6695
End If
If WebBrowser8.LocationURL = "about:blank" Or WebBrowser8.LocationURL = "" Then
On Error GoTo yel8
TabStrip1.Tabs.Add(9, "sud88", "Browser 8") = "Browser 8"
yel8:
Set ppDisp = WebBrowser8.object
GoTo 6695
End If
If WebBrowser9.LocationURL = "about:blank" Or WebBrowser9.LocationURL = "" Then
On Error GoTo yel9
TabStrip1.Tabs.Add(10, "sud99", "Browser 9") = "Browser 9"
yel9:
Set ppDisp = WebBrowser9.object
GoTo 6695
End If
If WebBrowser10.LocationURL = "about:blank" Or WebBrowser10.LocationURL = "" Then
On Error GoTo yel10
TabStrip1.Tabs.Add(11, "sud10", "Browser 10") = "Browser 10"
yel10:
Set ppDisp = WebBrowser10.object
GoTo 6695
End If
MsgBox ("All Browsers are in use. To close an active unwanted page click 'CLEAR' above.")
MsgBox ("The page will be opened in Internet Explorer")
6695: TabStrip1.Refresh
 
 End Sub

Private Sub WebBrowser8_ProgressChange(ByVal progress As Long, ByVal ProgressMax As Long)
On Error GoTo progressERR
If progress = -1 Then ProgressBar8.Value = 100
If progress > 0 And ProgressMax > 0 Then
    ProgressBar8.Value = progress * 100 / ProgressMax
    End If
    Exit Sub
progressERR:
               
End Sub

Private Sub WebBrowser8_StatusTextChange(ByVal Text As String)
If TabStrip1.SelectedItem.Key = "sud88" Then Label2.Caption = Text
End Sub

Private Sub WebBrowser8_TitleChange(ByVal Text As String)
For a = 1 To TabStrip1.Tabs.Count
If TabStrip1.Tabs(a).Key = "sud88" Then TabStrip1.Tabs(a).Caption = Left(Text, 15)
Next a

End Sub

Private Sub webBrowser9_BeforeNavigate2(ByVal pDisp As Object, URL As Variant, FLAGS As Variant, TargetFrameName As Variant, PostData As Variant, Headers As Variant, Cancel As Boolean)

Dim strURL As String
strURL = URL
Dim bFound As Boolean
Dim i As Integer


If Not bFound Then
cboURL.AddItem strURL
End If

cboURL.Text = strURL



End Sub

Private Sub webBrowser9_DownloadBegin()
ProgressBar1.Visible = True
End Sub

Private Sub webBrowser9_DownloadComplete()
ProgressBar9.Value = 100
ProgressBar9.Visible = False
End Sub

Private Sub webBrowser9_NewWindow2(ppDisp As Object, Cancel As Boolean)
If dispop.Checked = True Then Cancel = True: Exit Sub
If WebBrowser1.LocationURL = "about:blank" Or WebBrowser1.LocationURL = "" Then
On Error GoTo yel1
TabStrip1.Tabs.Add(2, "sud11", "Browser 1") = "Browser 1"
yel1:
Set ppDisp = WebBrowser1.object
GoTo 6696
End If
If WebBrowser2.LocationURL = "about:blank" Or WebBrowser2.LocationURL = "" Then
On Error GoTo yel2
TabStrip1.Tabs.Add(3, "sud22", "Browser 2") = "Browser 2"
yel2:
Set ppDisp = WebBrowser2.object
GoTo 6696
End If
If WebBrowser3.LocationURL = "about:blank" Or WebBrowser3.LocationURL = "" Then
On Error GoTo yel3
TabStrip1.Tabs.Add(4, "sud33", "Browser 3") = "Browser 3"
yel3:
Set ppDisp = WebBrowser3.object
GoTo 6696
End If

If WebBrowser4.LocationURL = "about:blank" Or WebBrowser4.LocationURL = "" Then
On Error GoTo yel4
TabStrip1.Tabs.Add(5, "sud44", "Browser 44") = "Browser 4"
yel4:
Set ppDisp = WebBrowser4.object
GoTo 6696
End If
If WebBrowser5.LocationURL = "about:blank" Or WebBrowser5.LocationURL = "" Then
On Error GoTo yel5
TabStrip1.Tabs.Add(6, "sud55", "Browser 55") = "Browser 5"
yel5:
Set ppDisp = WebBrowser5.object
GoTo 6696
End If
If WebBrowser6.LocationURL = "about:blank" Or WebBrowser6.LocationURL = "" Then
On Error GoTo yel6
TabStrip1.Tabs.Add(7, "sud66", "Browser 6") = "Browser 6"
yel6:
Set ppDisp = WebBrowser6.object
GoTo 6696
End If
If WebBrowser7.LocationURL = "about:blank" Or WebBrowser7.LocationURL = "" Then
On Error GoTo yel7
TabStrip1.Tabs.Add(8, "sud77", "Browser 7") = "Browser 7"
yel7:
Set ppDisp = WebBrowser7.object
GoTo 6696
End If
If WebBrowser8.LocationURL = "about:blank" Or WebBrowser8.LocationURL = "" Then
On Error GoTo yel8
TabStrip1.Tabs.Add(9, "sud88", "Browser 8") = "Browser 8"
yel8:
Set ppDisp = WebBrowser8.object
GoTo 6696
End If
If WebBrowser9.LocationURL = "about:blank" Or WebBrowser9.LocationURL = "" Then
On Error GoTo yel9
TabStrip1.Tabs.Add(10, "sud99", "Browser 9") = "Browser 9"
yel9:
Set ppDisp = WebBrowser9.object
GoTo 6696
End If
If WebBrowser10.LocationURL = "about:blank" Or WebBrowser10.LocationURL = "" Then
On Error GoTo yel10
TabStrip1.Tabs.Add(11, "sud10", "Browser 10") = "Browser 10"
yel10:
Set ppDisp = WebBrowser10.object
GoTo 6696
End If
MsgBox ("All Browsers are in use. To close an active unwanted page click 'CLEAR' above.")
MsgBox ("The page will be opened in Internet Explorer")
6696: TabStrip1.Refresh
 
 End Sub

Private Sub webBrowser9_ProgressChange(ByVal progress As Long, ByVal ProgressMax As Long)
On Error GoTo progressERR
If progress = -1 Then ProgressBar9.Value = 100
If progress > 0 And ProgressMax > 0 Then
    ProgressBar9.Value = progress * 100 / ProgressMax
    End If
    Exit Sub
progressERR:
               
End Sub

Private Sub webBrowser9_StatusTextChange(ByVal Text As String)
If TabStrip1.SelectedItem.Key = "sud99" Then Label2.Caption = Text
End Sub

Private Sub WebBrowser9_TitleChange(ByVal Text As String)
For a = 1 To TabStrip1.Tabs.Count
If TabStrip1.Tabs(a).Key = "sud99" Then TabStrip1.Tabs(a).Caption = Left(Text, 15)
Next a

End Sub

Private Sub winn_Click()
'Dim newInstance As New Form1 'Create a new instance of our form
'newInstance.Show
Dim frmWB As Form1
Set frmWB = New Form1
frmWB.Visible = True

End Sub

Private Sub wmo_Click()
Command7_Click
End Sub

Private Sub wmt_Click()
Command8_Click
End Sub

Private Sub woff_Click()
If woff.Caption = "Work Offline" Then
WebBrowser1.Offline = True
WebBrowser2.Offline = True
WebBrowser3.Offline = True
WebBrowser4.Offline = True
WebBrowser5.Offline = True
WebBrowser6.Offline = True
WebBrowser7.Offline = True
WebBrowser8.Offline = True
WebBrowser9.Offline = True
WebBrowser10.Offline = True
woff.Caption = "Go Online"
ut = 1
End If
If ut = 1 Then GoTo 109

If woff.Caption = "Go Online" Then
WebBrowser1.Offline = False
WebBrowser2.Offline = False
WebBrowser3.Offline = False
WebBrowser4.Offline = False
WebBrowser5.Offline = False
WebBrowser6.Offline = False
WebBrowser7.Offline = False
WebBrowser8.Offline = False
WebBrowser9.Offline = False
WebBrowser10.Offline = False
woff.Caption = "Work Offline"
End If
109 End Sub

Private Sub ys_Click()
cboURL.Text = "www.yahoo.com/science/"
cboURL_Click

End Sub
