VERSION 5.00
Begin VB.Form frm_Prn_Set 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Communication Settings"
   ClientHeight    =   3255
   ClientLeft      =   6150
   ClientTop       =   5355
   ClientWidth     =   4950
   Icon            =   "frm_Prn_Set.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3255
   ScaleWidth      =   4950
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmd_Cancel 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Left            =   3000
      TabIndex        =   18
      Top             =   2160
      Width           =   1575
   End
   Begin VB.CommandButton cmd_OK 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   3000
      TabIndex        =   17
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Frame Frame_TCPIP 
      Caption         =   "TCP/IP"
      Enabled         =   0   'False
      Height          =   1095
      Left            =   120
      TabIndex        =   14
      Top             =   2040
      Width           =   2535
      Begin VB.TextBox txt_Port 
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   1680
         MaxLength       =   5
         TabIndex        =   16
         Top             =   600
         Width           =   615
      End
      Begin VB.TextBox txt_IP 
         ForeColor       =   &H00C0C0C0&
         Height          =   285
         Left            =   120
         MaxLength       =   15
         TabIndex        =   15
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "Port"
         Height          =   255
         Left            =   1680
         TabIndex        =   24
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "Adresse"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame_COM_Set 
      Caption         =   "COM Settings"
      Enabled         =   0   'False
      Height          =   1815
      Left            =   2400
      TabIndex        =   9
      Top             =   120
      Width           =   2415
      Begin VB.ComboBox cmb_StopBits 
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frm_Prn_Set.frx":0442
         Left            =   1320
         List            =   "frm_Prn_Set.frx":0444
         TabIndex        =   13
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox cmb_Par 
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frm_Prn_Set.frx":0446
         Left            =   120
         List            =   "frm_Prn_Set.frx":0448
         TabIndex        =   12
         Top             =   1320
         Width           =   975
      End
      Begin VB.ComboBox cmb_DatBits 
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frm_Prn_Set.frx":044A
         Left            =   1320
         List            =   "frm_Prn_Set.frx":044C
         TabIndex        =   11
         Top             =   600
         Width           =   975
      End
      Begin VB.ComboBox cmb_Baud 
         ForeColor       =   &H00C0C0C0&
         Height          =   315
         ItemData        =   "frm_Prn_Set.frx":044E
         Left            =   120
         List            =   "frm_Prn_Set.frx":0450
         TabIndex        =   10
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Stop Bits"
         Height          =   255
         Left            =   1320
         TabIndex        =   22
         Top             =   1080
         Width           =   855
      End
      Begin VB.Label Label3 
         Caption         =   "Parität"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Daten Bits"
         Height          =   255
         Left            =   1320
         TabIndex        =   20
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Baud"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   1215
      End
   End
   Begin VB.Frame Frame_Port 
      Caption         =   "Anschluss"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.OptionButton opt_Port 
         Caption         =   "USB"
         Height          =   255
         Index           =   7
         Left            =   1200
         TabIndex        =   8
         Top             =   1320
         Width           =   855
      End
      Begin VB.OptionButton opt_Port 
         Caption         =   "TCP/IP"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   7
         Top             =   1320
         Width           =   1095
      End
      Begin VB.OptionButton opt_Port 
         Caption         =   "COM4"
         Height          =   255
         Index           =   5
         Left            =   1200
         TabIndex        =   6
         Top             =   960
         Width           =   855
      End
      Begin VB.OptionButton opt_Port 
         Caption         =   "COM3"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   1095
      End
      Begin VB.OptionButton opt_Port 
         Caption         =   "COM2"
         Height          =   255
         Index           =   3
         Left            =   1200
         TabIndex        =   4
         Top             =   720
         Width           =   855
      End
      Begin VB.OptionButton opt_Port 
         Caption         =   "COM1"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   3
         Top             =   720
         Width           =   1095
      End
      Begin VB.OptionButton opt_Port 
         Caption         =   "LPT2"
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   735
      End
      Begin VB.OptionButton opt_Port 
         Caption         =   "LPT1"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frm_Prn_Set"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_Cancel_Click()

'Schließen
Me.Hide
End Sub

Private Sub cmd_OK_Click()

'Steuert die Übergabe der Variablen und Texte
Select Case SelCom
    Case 1: frm_Main.StatBar.Panels.Item(1).Text = "LPT1"
    Case 2: frm_Main.StatBar.Panels.Item(1).Text = "LPT2"
    Case 3 To 6: get_COM_Par
    Case 7: SetNet
    Case 8: frm_Main.StatBar.Panels.Item(1).Text = "USB"
End Select

Me.Hide
End Sub

Private Sub Form_Load()
FillCOM
End Sub


Private Sub opt_Port_Click(Index As Integer)

'Je nach gewählten Anschluss das Frame aktivieren oder deaktivieren
Select Case Index
    Case 0
        Frame_COM_Set.Enabled = False
        Frame_TCPIP.Enabled = False
        SetForeColorNet
        SetForeColorCom
        SelCom = 1
    Case 1
        Frame_COM_Set.Enabled = False
        Frame_TCPIP.Enabled = False
        SetForeColorNet
        SetForeColorCom
        SelCom = 2
    Case 2
        Frame_COM_Set.Enabled = True
        Frame_TCPIP.Enabled = False
        SetForeColorNet
        SetForeColorComB
        SelCom = 3
    Case 3
        Frame_COM_Set.Enabled = True
        Frame_TCPIP.Enabled = False
        SetForeColorComB
        SetForeColorNet
        SelCom = 4
    Case 4
        Frame_COM_Set.Enabled = True
        Frame_TCPIP.Enabled = False
        SetForeColorComB
        SetForeColorNet
        SelCom = 5
    Case 5
        Frame_COM_Set.Enabled = True
        Frame_TCPIP.Enabled = False
        SetForeColorComB
        SetForeColorNet
        SelCom = 6
    Case 6
        Frame_COM_Set.Enabled = False
        Frame_TCPIP.Enabled = True
        txt_Port.ForeColor = &H80000008
        txt_IP.ForeColor = &H80000008
        SetForeColorCom
        SelCom = 7
    Case 7
        Frame_COM_Set.Enabled = False
        Frame_TCPIP.Enabled = False
        SetForeColorCom
        SetForeColorNet
        SelCom = 8
        MsgBox err_com_set, vbExclamation, "Information"
End Select
    
End Sub

Private Sub txt_IP_KeyPress(KeyAscii As Integer)

'Eingabe überprüfen und nur Zahlen bzw. "." zulassen
    Select Case KeyAscii
        Case 8
        Case 46
        Case 48 To 57
        Case Else: KeyAscii = 0
    End Select

End Sub

Private Sub txt_Port_KeyPress(KeyAscii As Integer)

'Eingabe überprüfen und nur Zahlen zulassen
    Select Case KeyAscii
        Case 8
        Case 48 To 57
        Case Else: KeyAscii = 0
    End Select

End Sub
