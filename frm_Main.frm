VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "mscomm32.ocx"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frm_Main 
   Caption         =   "Zebra Control"
   ClientHeight    =   5130
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   6945
   Icon            =   "frm_Main.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   6945
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.Timer tmr_time 
      Interval        =   800
      Left            =   2280
      Top             =   960
   End
   Begin MSComDlg.CommonDialog Com_Dlg 
      Left            =   1680
      Top             =   960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSWinsockLib.Winsock MSSock 
      Left            =   1080
      Top             =   960
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSCommLib.MSComm MSCom 
      Left            =   360
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin MSComctlLib.StatusBar StatBar 
      Align           =   2  'Unten ausrichten
      Height          =   255
      Left            =   0
      TabIndex        =   1
      Top             =   4875
      Width           =   6945
      _ExtentX        =   12250
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   3528
            MinWidth        =   3528
            Text            =   "Port"
            TextSave        =   "Port"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "Num"
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   882
            MinWidth        =   882
            Text            =   "Cap"
            TextSave        =   "FEST"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            AutoSize        =   1
            Object.Width           =   6385
            MinWidth        =   1411
         EndProperty
      EndProperty
   End
   Begin RichTextLib.RichTextBox Rtx 
      Height          =   3800
      Left            =   0
      TabIndex        =   0
      Top             =   480
      Width           =   6375
      _ExtentX        =   11245
      _ExtentY        =   6694
      _Version        =   393217
      ScrollBars      =   3
      TextRTF         =   $"frm_Main.frx":0442
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00E0E0E0&
      X1              =   1200
      X2              =   1200
      Y1              =   0
      Y2              =   480
   End
   Begin VB.Image img_Prop 
      Height          =   240
      Left            =   1680
      Picture         =   "frm_Main.frx":04C4
      ToolTipText     =   "Druckereinstellungen"
      Top             =   120
      Width           =   240
   End
   Begin VB.Image img_New 
      Height          =   240
      Left            =   120
      Picture         =   "frm_Main.frx":05C6
      ToolTipText     =   "Neues Script"
      Top             =   120
      Width           =   240
   End
   Begin VB.Image img_Open 
      Height          =   240
      Left            =   480
      Picture         =   "frm_Main.frx":06C8
      ToolTipText     =   "Script öffnen"
      Top             =   120
      Width           =   240
   End
   Begin VB.Image img_Save 
      Height          =   240
      Left            =   840
      Picture         =   "frm_Main.frx":07CA
      ToolTipText     =   "Script speichern"
      Top             =   120
      Width           =   240
   End
   Begin VB.Image img_Print 
      Height          =   240
      Left            =   1320
      Picture         =   "frm_Main.frx":08CC
      ToolTipText     =   "Script drucken"
      Top             =   120
      Width           =   240
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0C0C0&
      FillColor       =   &H00C0C0C0&
      FillStyle       =   0  'Ausgefüllt
      Height          =   495
      Left            =   0
      Top             =   0
      Width           =   2295
   End
   Begin VB.Menu File 
      Caption         =   "1"
      Begin VB.Menu New 
         Caption         =   "2"
         Shortcut        =   ^N
      End
      Begin VB.Menu Open 
         Caption         =   "3"
         Shortcut        =   ^O
      End
      Begin VB.Menu Save 
         Caption         =   "4"
         Shortcut        =   ^S
      End
      Begin VB.Menu Leer1 
         Caption         =   "-"
      End
      Begin VB.Menu PrinterSet 
         Caption         =   "5"
         Shortcut        =   ^E
      End
      Begin VB.Menu Print 
         Caption         =   "5"
         Shortcut        =   ^P
      End
      Begin VB.Menu Leer 
         Caption         =   "-"
      End
      Begin VB.Menu Exit 
         Caption         =   "6"
      End
   End
   Begin VB.Menu zpl 
      Caption         =   "7"
      Begin VB.Menu Tool 
         Caption         =   "&Tools"
         Begin VB.Menu PrConf 
            Caption         =   "8"
            Shortcut        =   ^K
         End
         Begin VB.Menu MemoryList 
            Caption         =   "9"
         End
         Begin VB.Menu MediaCal 
            Caption         =   "10"
         End
         Begin VB.Menu MedProfile 
            Caption         =   "11"
         End
         Begin VB.Menu EnDump 
            Caption         =   "12"
         End
         Begin VB.Menu DisDump 
            Caption         =   "13"
         End
         Begin VB.Menu InitFlash 
            Caption         =   "13"
         End
         Begin VB.Menu dot6 
            Caption         =   "14"
         End
         Begin VB.Menu Dot12 
            Caption         =   "15"
         End
         Begin VB.Menu PassChange 
            Caption         =   "16"
         End
         Begin VB.Menu trenn 
            Caption         =   "-"
         End
         Begin VB.Menu TestCode 
            Caption         =   "17"
            Begin VB.Menu C39 
               Caption         =   "Code &39"
            End
            Begin VB.Menu C128 
               Caption         =   "Code &128"
            End
            Begin VB.Menu I25 
               Caption         =   "Interleaved &2/5"
            End
         End
      End
      Begin VB.Menu Med_Track 
         Caption         =   "18"
         Begin VB.Menu Cont 
            Caption         =   "19"
         End
         Begin VB.Menu Non_Cont 
            Caption         =   "20"
         End
      End
      Begin VB.Menu Med_Type 
         Caption         =   "21"
         Begin VB.Menu trans 
            Caption         =   "22"
         End
         Begin VB.Menu Direct 
            Caption         =   "23"
         End
      End
      Begin VB.Menu PM 
         Caption         =   "24"
         Begin VB.Menu Tear 
            Caption         =   "25"
         End
         Begin VB.Menu Peel 
            Caption         =   "26"
         End
         Begin VB.Menu Rewind 
            Caption         =   "27"
         End
         Begin VB.Menu Cutter 
            Caption         =   "28"
         End
         Begin VB.Menu Applicator 
            Caption         =   "29"
         End
      End
      Begin VB.Menu Lang 
         Caption         =   "30"
         Begin VB.Menu English 
            Caption         =   "31"
         End
         Begin VB.Menu German 
            Caption         =   "32"
         End
         Begin VB.Menu Italyan 
            Caption         =   "33"
         End
         Begin VB.Menu Spanish 
            Caption         =   "34"
         End
         Begin VB.Menu Spanish2 
            Caption         =   "35"
         End
         Begin VB.Menu Protugise 
            Caption         =   "36"
         End
         Begin VB.Menu France 
            Caption         =   "37"
         End
         Begin VB.Menu Norwayan 
            Caption         =   "38"
         End
         Begin VB.Menu Swedish 
            Caption         =   "39"
         End
         Begin VB.Menu Finnisch 
            Caption         =   "40"
         End
         Begin VB.Menu Dansk 
            Caption         =   "41"
         End
         Begin VB.Menu Netherlandish 
            Caption         =   "42"
         End
         Begin VB.Menu Japanish 
            Caption         =   "43"
         End
      End
      Begin VB.Menu setsens 
         Caption         =   "44"
         Begin VB.Menu transmi 
            Caption         =   "45"
         End
         Begin VB.Menu refle 
            Caption         =   "46"
         End
         Begin VB.Menu asel 
            Caption         =   "47"
         End
      End
      Begin VB.Menu mo 
         Caption         =   "&Motions..."
      End
   End
   Begin VB.Menu epl 
      Caption         =   "48"
      Begin VB.Menu pr_Conf 
         Caption         =   "49"
      End
      Begin VB.Menu set_Label 
         Caption         =   "50"
      End
      Begin VB.Menu cali 
         Caption         =   "51"
      End
   End
   Begin VB.Menu Men_Lan 
      Caption         =   "53"
      Begin VB.Menu M_Ger 
         Caption         =   "54"
      End
      Begin VB.Menu M_Eng 
         Caption         =   "55"
         Checked         =   -1  'True
      End
   End
   Begin VB.Menu Info 
      Caption         =   "&Info"
      Begin VB.Menu About 
         Caption         =   "52"
      End
   End
End
Attribute VB_Name = "frm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub About_Click()
frm_About.Show
End Sub

Private Sub Applicator_Click()
Rtx.Text = ZApp
End Sub

Private Sub asel_Click()
Rtx.Text = ZSAu
End Sub

Private Sub C128_Click()
Rtx.Text = Z128
End Sub

Private Sub C39_Click()
Rtx.Text = Z39
End Sub

Private Sub cali_Click()
Rtx.Text = ECal
End Sub



Private Sub Cutter_Click()
Rtx.Text = ZCutter
End Sub

Private Sub Dansk_Click()
Rtx.Text = ZDän
End Sub

Private Sub Direct_Click()
Rtx.Text = ZDirect
End Sub

Private Sub DisDump_Click()
Rtx.Text = ZDiDu
End Sub

Private Sub Dot12_Click()
Rtx.Text = ZSet12
End Sub

Private Sub dot6_Click()
Rtx.Text = ZSet6
End Sub

Private Sub EnDump_Click()
Rtx.Text = ZEnDu
End Sub

Private Sub English_Click()
Rtx.Text = ZEng
End Sub

Private Sub Exit_Click()
End
End Sub

Private Sub Finnisch_Click()
Rtx.Text = ZFin
End Sub

Private Sub Form_Load()

'Englisch als Startsprache festlegen
M_Eng_Click
'Aktuelle Uhrzeit anzeigen
StatBar.Panels.Item(4).Text = Date & " - " & Time

'LPT1 als Standard festlegen und in die Statusbar schreiben
SelCom = 1
frm_Main.StatBar.Panels.Item(1).Text = "LPT1"

'Größe anpassen
Me.Height = 6000
Me.Width = 6000

'ZPL-Variablen füllen
ZPL_Var

End Sub

Private Sub Form_Resize()

If Me.WindowState = vbMinimized Then Exit Sub

'Toolbar und Textbox der Größe der Form anpassen
On Error GoTo ToSmall
Rtx.Width = Me.Width - 100
Rtx.Height = Me.Height - 1500
Shape1.Width = Me.Width
Exit Sub

ToSmall:
Me.Height = 3000
Me.Width = 2400

End Sub

Private Sub Form_Terminate()
Unload frm_About
Unload frm_epl_lab
Unload frm_Motion
Unload frm_Pass
Unload frm_Prn_Set
Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload frm_About
Unload frm_epl_lab
Unload frm_Motion
Unload frm_Pass
Unload frm_Prn_Set
Unload Me
End Sub

Private Sub France_Click()
Rtx.Text = ZFra
End Sub

Private Sub German_Click()
Rtx.Text = ZDeu
End Sub

Private Sub I25_Click()
Rtx.Text = Z25
End Sub

Private Sub img_New_Click()
New_Click
End Sub

Private Sub img_Open_Click()
Open_Click
End Sub

Private Sub img_Print_Click()
Print_Click
End Sub

Private Sub img_Prop_Click()
PrinterSet_Click
End Sub

Private Sub img_Save_Click()
Save_Click
End Sub

Private Sub InitFlash_Click()
Rtx.Text = ZIniFl
End Sub

Private Sub Italyan_Click()
Rtx.Text = ZIta
End Sub

Private Sub Japanish_Click()
Rtx.Text = ZJap
End Sub

Private Sub M_Eng_Click()
La = 1
Lan_Eng
M_Ger.Checked = False
M_Eng.Checked = True

err_epl_lab = "Wrong or no value enterd"
err_com_set = "Use only a " & Chr$(34) & "Generic Text Only" & Chr$(34) & Chr$(10) & "Printer with USB"
err_con = "Connection Error, please check settings"
End Sub

Private Sub M_Ger_Click()
La = 0
Lan_Ger
M_Ger.Checked = True
M_Eng.Checked = False

err_epl_lab = "Falscher oder kein Wert eingegeben"
err_com_set = "Es kann nur ein " & Chr$(34) & "Generic Text Only" & Chr$(34) & Chr$(10) & "Drucker verwendet werden"
err_con = "Fehler bei der Verbindung bitte Einstellungen überprüfen"
End Sub

Private Sub MediaCal_Click()
Rtx.Text = ZMedCal
End Sub

Private Sub MedProfile_Click()
Rtx.Text = ZMedPro
End Sub

Private Sub MemoryList_Click()
Rtx.Text = ZMemList
End Sub

Private Sub mo_Click()
frm_Motion.Show
End Sub

Private Sub Netherlandish_Click()
Rtx.Text = ZHol
End Sub

Private Sub New_Click()
Rtx.Text = ""
End Sub

Private Sub Non_Cont_Click()
Rtx.Text = ZNCont
End Sub

Private Sub Cont_Click()
Rtx.Text = ZCont
End Sub

Private Sub Norwayan_Click()
Rtx.Text = ZNor
End Sub

Private Sub Open_Click()

On Error GoTo Cancel

'Dialog Konfigurieren
Com_Dlg.DialogTitle = "Script öffnen ..."
Com_Dlg.Filter = "*.*"
Com_Dlg.ShowOpen
FileName = Com_Dlg.FileName

'String an die RTF Box übergeben
Rtx.FileName = FileName
Exit Sub

Cancel:
Exit Sub

End Sub

Private Sub PassChange_Click()
'Form zum ändern zeigen
frm_Pass.Show

End Sub

Private Sub Peel_Click()
Rtx.Text = ZPeel
End Sub

Private Sub pr_Conf_Click()
Rtx.Text = ECon
End Sub

Private Sub PrConf_Click()
Rtx.Text = ZConf
End Sub

Private Sub Print_Click()

PrintOut

End Sub

Private Sub PrinterSet_Click()
'Druckereinstellungen öffnen
frm_Prn_Set.Show
End Sub

Private Sub Protugise_Click()
Rtx.Text = ZPor
End Sub

Private Sub refle_Click()
Rtx.Text = ZSRe
End Sub

Private Sub Rewind_Click()
Rtx.Text = ZRewind
End Sub

Private Sub Save_Click()

On Error GoTo Cancel

'Dialog Konfigurieren
Com_Dlg.DialogTitle = "Script speichern unter ..."
Com_Dlg.Filter = "*.*"
Com_Dlg.ShowSave
FileName = Com_Dlg.FileName

'String aus der RTF Box ausgeben
Open FileName For Output As #1
    Print #1, Rtx.Text
Close #1

Exit Sub

Cancel:
Exit Sub
    
End Sub

Private Sub set_Label_Click()
frm_epl_lab.Show
End Sub

Private Sub Spanish_Click()
Rtx.Text = ZSpa
End Sub

Private Sub Spanish2_Click()
Rtx.Text = ZSpa2
End Sub

Private Sub Swedish_Click()
Rtx.Text = ZSch
End Sub

Private Sub SystemInfo_Click()
MsgBox "Systeminfo"
End Sub

Private Sub Tear_Click()
Rtx.Text = ZTear
End Sub

Private Sub tmr_time_Timer()
StatBar.Panels.Item(4).Text = Date & " - " & Time
End Sub

Private Sub Trans_Click()
Rtx.Text = ZTrans
End Sub

Private Sub transmi_Click()
Rtx.Text = ZSTr
End Sub
