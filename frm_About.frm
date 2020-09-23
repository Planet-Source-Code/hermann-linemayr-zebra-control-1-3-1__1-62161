VERSION 5.00
Begin VB.Form frm_About 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Info"
   ClientHeight    =   5280
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   3030
   ClipControls    =   0   'False
   Icon            =   "frm_About.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3644.35
   ScaleMode       =   0  'Benutzerdefiniert
   ScaleWidth      =   2845.327
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Line Line2 
      X1              =   112.686
      X2              =   2704.469
      Y1              =   3147.393
      Y2              =   3147.393
   End
   Begin VB.Label Label13 
      Caption         =   "support@bsr.at"
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   4200
      Width           =   2295
   End
   Begin VB.Label Label12 
      Caption         =   "+43 (0)662 456323 122"
      Height          =   255
      Left            =   600
      TabIndex        =   13
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label11 
      Caption         =   "Tel.:"
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3960
      Width           =   495
   End
   Begin VB.Line Line1 
      X1              =   112.686
      X2              =   2704.469
      Y1              =   2484.784
      Y2              =   2484.784
   End
   Begin VB.Label Label10 
      Caption         =   "office@bsr.at"
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3000
      Width           =   2055
   End
   Begin VB.Label Label9 
      Caption         =   "+43 (0)662 455937 99"
      Height          =   255
      Left            =   600
      TabIndex        =   10
      Top             =   2760
      Width           =   1815
   End
   Begin VB.Label Label8 
      Caption         =   "Fax:"
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2760
      Width           =   495
   End
   Begin VB.Label Label7 
      Caption         =   "Tel.:"
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2520
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "+43 (0)662 456323 0"
      Height          =   255
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   1815
   End
   Begin VB.Label Label5 
      Caption         =   "Helpdesk:"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   3720
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "A-5020 Salzburg"
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   2160
      Width           =   2175
   End
   Begin VB.Label Label3 
      Caption         =   "Jakob-Haringer-Strasse 3"
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "BSR idware GmbH"
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   2175
   End
   Begin VB.Label lbl_link 
      Caption         =   "http://www.bsr.at"
      Height          =   255
      Left            =   120
      MouseIcon       =   "frm_About.frx":0442
      MousePointer    =   99  'Benutzerdefiniert
      TabIndex        =   2
      Top             =   3240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "©2003 - 2006 by MrLine"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   4920
      Width           =   1935
   End
   Begin VB.Label lbl_app 
      Caption         =   "Label1"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   4680
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1500
      Left            =   120
      Picture         =   "frm_About.frx":074C
      ToolTipText     =   "BSR Homepage"
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frm_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub

Private Sub Form_Load()
lbl_app.Caption = App.ProductName & " " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With lbl_link
    .FontBold = False
    .FontUnderline = False
    .ForeColor = vbBlack
End With
    Me.MousePointer = 0
End Sub

Private Sub Image1_Click()
ShellExecute hWnd, "open", "http://www.bsr.at", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub lbl_link_Click()
ShellExecute hWnd, "open", "http://www.bsr.at", vbNullString, vbNullString, conSwNormal
End Sub

Private Sub lbl_link_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
With lbl_link
    .FontBold = True
    .FontUnderline = True
    .ForeColor = vbBlue
End With
    Me.MousePointer = 99
End Sub
