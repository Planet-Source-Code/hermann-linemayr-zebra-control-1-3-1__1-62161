VERSION 5.00
Begin VB.Form frm_Motion 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Power Up / Head Close Motion"
   ClientHeight    =   1785
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3690
   Icon            =   "frm_Motion.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1785
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.ComboBox cmb_Hc 
      Height          =   315
      Left            =   1440
      TabIndex        =   1
      Text            =   "Calibrate"
      Top             =   720
      Width           =   1935
   End
   Begin VB.ComboBox cmb_Pu 
      Height          =   315
      ItemData        =   "frm_Motion.frx":0442
      Left            =   1440
      List            =   "frm_Motion.frx":0444
      TabIndex        =   0
      Text            =   "Calibrate"
      Top             =   240
      Width           =   1935
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "Abbrechen"
      Height          =   375
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "Head Close"
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Power Up"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   360
      Width           =   1335
   End
End
Attribute VB_Name = "frm_Motion"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub Form_Load()
Select Case La
    Case 0: Fill_ger
    Case 1: Fill_eng
End Select
End Sub

Private Sub OKButton_Click()

Dim VPu, VHc As String

'Auswählen der Power Up Variable
Select Case cmb_Pu.ListIndex
    Case 0: VPu = "F"
    Case 1: VPu = "C"
    Case 2: VPu = "L"
    Case 3: VPu = "N"
End Select

'Auswählen der Head Close Variable
Select Case cmb_Hc.ListIndex
    Case 0: VHc = "F"
    Case 1: VHc = "C"
    Case 2: VHc = "L"
    Case 3: VHc = "N"
End Select

'Wenn nichts ausgewählt wurde auf Calibrate setzen
If VPu = "" Then VPu = "C"
If VHc = "" Then VHc = "C"

ZMot = "^XA" & Chr$(13) & Chr$(10) & "^MF" & VPu & "," & VHc & Chr$(13) & Chr$(10) & _
"^JUS" & Chr$(13) & Chr$(10) & "^XZ"

frm_Main.Rtx.Text = ZMot

Unload Me
End Sub

Public Function Fill_eng()
        'Power Up Box füllen
        cmb_Pu.Clear
        cmb_Pu.Text = "Calibrate"
        cmb_Pu.AddItem "Feed"
        cmb_Pu.AddItem "Calibrate"
        cmb_Pu.AddItem "Length"
        cmb_Pu.AddItem "No Motion"
        
        'Head Close Box füllen
        cmb_Hc.Clear
        cmb_Hc.Text = "Calibrate"
        cmb_Hc.AddItem "Feed"
        cmb_Hc.AddItem "Calibrate"
        cmb_Hc.AddItem "Length"
        cmb_Hc.AddItem "No Motion"
End Function

Public Function Fill_ger()
        'Power Up Box füllen
        cmb_Pu.Clear
        cmb_Pu.Text = "Kalibrieren"
        cmb_Pu.AddItem "Vorschub"
        cmb_Pu.AddItem "Kalibrieren"
        cmb_Pu.AddItem "Länge"
        cmb_Pu.AddItem "Keine Bewegung"
        
        'Head Close Box füllen
        cmb_Hc.Clear
        cmb_Hc.Text = "Kalibrieren"
        cmb_Hc.AddItem "Vorschub"
        cmb_Hc.AddItem "Kalibrieren"
        cmb_Hc.AddItem "Länge"
        cmb_Hc.AddItem "Keine Bewegung"
End Function
