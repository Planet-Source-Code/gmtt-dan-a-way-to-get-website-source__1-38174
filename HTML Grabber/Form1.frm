VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "SHDOCVW.DLL"
Begin VB.Form Form1 
   Caption         =   "Source Code Stealer"
   ClientHeight    =   3075
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11670
   LinkTopic       =   "Form1"
   ScaleHeight     =   3075
   ScaleWidth      =   11670
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command6 
      Caption         =   "Hide Website"
      Height          =   495
      Left            =   0
      TabIndex        =   8
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Show Website"
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   2160
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Save As .txt"
      Height          =   495
      Left            =   0
      TabIndex        =   6
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Save As Word Document"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin SHDocVwCtl.WebBrowser WebBrowser1 
      Height          =   5415
      Left            =   0
      TabIndex        =   4
      Top             =   3120
      Width           =   11655
      ExtentX         =   20558
      ExtentY         =   9551
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
   Begin VB.CommandButton Command2 
      Caption         =   "Save As HTML"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      ToolTipText     =   "HTML Saved as text file on the desktop"
      Top             =   720
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3480
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin InetCtlsObjects.Inet Inet1 
      Left            =   4080
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1200
      TabIndex        =   2
      Text            =   "Url"
      Top             =   0
      Width           =   10455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Grab HTML"
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   1215
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2775
      Left            =   1200
      TabIndex        =   0
      Top             =   360
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   4895
      _Version        =   393217
      Enabled         =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"Form1.frx":0000
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim strHTML As String
strHTML = Inet1.OpenURL(Text1.Text)

RichTextBox1.Text = strHTML
WebBrowser1.Navigate2 (Text1)

End Sub

Private Sub Command2_Click()
Open CurDir & "/html.html" For Output As #1
Write #1, RichTextBox1.Text
Close #1
End Sub

Private Sub Command3_Click()
Open CurDir & "/html.doc" For Output As #1
Write #1, RichTextBox1.Text
Close #1
End Sub

Private Sub Command4_Click()
Open CurDir & "/html.txt" For Output As #1
Write #1, RichTextBox1.Text
Close #1
End Sub

Private Sub Command5_Click()
WebBrowser1.Visible = True
WindowState = vbMaximized
End Sub

Private Sub Command6_Click()
WebBrowser1.Visible = False
WindowState = vbNormal
End Sub

Private Sub Form_Load()
WebBrowser1.Visible = False
End Sub

Private Sub Timer1_Timer()

End Sub

