VERSION 5.00
Begin VB.Form frm_Pass 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Set Pass"
   ClientHeight    =   750
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   1965
   Icon            =   "frm_Pass.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   750
   ScaleWidth      =   1965
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CommandButton cmd_Set 
      Caption         =   "&Set"
      Default         =   -1  'True
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.TextBox txt_Pass 
      Height          =   285
      Left            =   120
      MaxLength       =   4
      TabIndex        =   0
      Text            =   "1234"
      Top             =   360
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Passwort: (0000-9999)"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "frm_Pass"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Set_Click()

'Passwort übernehmen und ausgeben
Pass = txt_Pass.Text
ZPass = "^XA" & Chr$(13) & "^KP" & Pass & Chr$(13) & "^JUS" & Chr$(13) & "^XZ"
frm_Main.Rtx.Text = ZPass
Unload Me

End Sub

Private Sub Form_Load()
txt_Pass.SelLength = 4
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = 27 Then Unload Me
End Sub
Private Sub txt_Pass_KeyPress(KeyAscii As Integer)

Select Case KeyAscii
        Case 48 To 57
        Case 27: Unload Me
        Case Else: KeyAscii = 0
End Select
End Sub
