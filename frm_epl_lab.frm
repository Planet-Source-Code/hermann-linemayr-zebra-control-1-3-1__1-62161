VERSION 5.00
Begin VB.Form frm_epl_lab 
   BorderStyle     =   3  'Fester Dialog
   Caption         =   "Generate a Label"
   ClientHeight    =   2820
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   6000
   Icon            =   "frm_epl_lab.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2820
   ScaleWidth      =   6000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'Bildschirmmitte
   Begin VB.CheckBox opt_Black 
      Caption         =   "&Black Mark"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   1800
      Width           =   1335
   End
   Begin VB.CheckBox opt_Therm 
      Caption         =   "&Thermal Direct"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1440
      Width           =   1335
   End
   Begin VB.TextBox txt_S 
      Height          =   285
      Left            =   1320
      TabIndex        =   2
      Top             =   960
      Width           =   975
   End
   Begin VB.TextBox txt_H 
      Height          =   285
      Left            =   1320
      TabIndex        =   1
      Top             =   600
      Width           =   975
   End
   Begin VB.TextBox txt_B 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   240
      Width           =   975
   End
   Begin VB.CommandButton CancelButton 
      Cancel          =   -1  'True
      Caption         =   "&Abbrechen"
      Height          =   375
      Left            =   1560
      TabIndex        =   6
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton OKButton 
      Caption         =   "&OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   2280
      Width           =   1215
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   4200
      X2              =   5040
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   3840
      X2              =   4080
      Y1              =   2280
      Y2              =   2280
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   1215
      Index           =   1
      Left            =   3240
      Shape           =   4  'Gerundetes Rechteck
      Top             =   2280
      Width           =   1335
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   4320
      X2              =   5400
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line3 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   3840
      X2              =   4080
      Y1              =   2160
      Y2              =   2160
   End
   Begin VB.Line Line8 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   5040
      X2              =   5040
      Y1              =   2280
      Y2              =   2520
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   1215
      Index           =   0
      Left            =   3240
      Shape           =   4  'Gerundetes Rechteck
      Top             =   960
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   3960
      X2              =   3960
      Y1              =   2040
      Y2              =   2400
   End
   Begin VB.Label Label1 
      Caption         =   "Zwischenraum"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   12
      Top             =   2520
      Width           =   1095
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   3240
      X2              =   3240
      Y1              =   360
      Y2              =   720
   End
   Begin VB.Line Line4 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   1
      X1              =   4560
      X2              =   4560
      Y1              =   360
      Y2              =   720
   End
   Begin VB.Line Line5 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   3120
      X2              =   4680
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Zentriert
      BackStyle       =   0  'Transparent
      Caption         =   "Breite"
      Height          =   255
      Index           =   1
      Left            =   3240
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   1215
      Index           =   3
      Left            =   3240
      Shape           =   4  'Gerundetes Rechteck
      Top             =   -360
      Width           =   1335
   End
   Begin VB.Line Line6 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      Index           =   0
      X1              =   4320
      X2              =   5400
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line7 
      BorderColor     =   &H000000FF&
      BorderWidth     =   2
      X1              =   5280
      X2              =   5280
      Y1              =   840
      Y2              =   2280
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Höhe"
      Height          =   255
      Index           =   1
      Left            =   4800
      TabIndex        =   10
      Top             =   1440
      Width           =   735
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Rechts
      Caption         =   "Spalt (mm)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   9
      Top             =   960
      Width           =   1095
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Rechts
      Caption         =   "Höhe (mm)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   8
      Top             =   600
      Width           =   1095
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Rechts
      Caption         =   "Breite (mm)"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   240
      Width           =   1095
   End
   Begin VB.Shape Shape2 
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Ausgefüllt
      Height          =   4575
      Left            =   3120
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "frm_epl_lab"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CancelButton_Click()
Unload Me
End Sub

Private Sub OKButton_Click()

On Error GoTo FehlerWert 'Bei falscher Eingbabe Felhermeldung ausgeben

Dim mmB, mmH, mmS, dB, dH, dS As Double
Dim TStat, BStat As String
Dim oldstring, newletter, oldletter, newstring As String

    oldletter = "."     '. durch , ersetzen damit gerechnet werden kann
    newletter = ","
    
    oldstring = txt_B.Text      'Im Wert den "." durch einen "," ersetzen
    newstring = Replace(oldstring, newletter, oldletter)
    txt_B.Text = newstring

    oldstring = txt_H.Text
    newstring = Replace(oldstring, newletter, oldletter)
    txt_H.Text = newstring
    
    oldstring = txt_S.Text
    newstring = Replace(oldstring, newletter, oldletter)
    txt_S.Text = newstring


    mmB = txt_B.Text   'Variablen mit den Daten aus den Eingabefeldern füllen
    mmH = txt_H.Text
    mmS = txt_S.Text
    
    dB = mmB * 8    'Umrechnen der Werte aus den Variaben in DOT
    dH = mmH * 8
    dS = mmS * 8
    
    oldletter = ","     ', durch . ersetzen da EPL Kommas mit . macht
    newletter = "."
    
    oldstring = dB      'Im umgerechneten Breite Wert den "," durch einen "." ersetzen
    newstring = Replace(oldstring, newletter, oldletter)
    dB = newstring
    
    oldstring = dH
    newstring = Replace(oldstring, newletter, oldletter)
    dH = newstring
    
    oldstring = dS
    newstring = Replace(oldstring, newletter, oldletter)
    dS = newstring


PrintMode = opt_Therm.Value 'Überprüfen ob Direct oder Thermal

    Select Case PrintMode
        Case 1
            TStat = "D"
        Case Else
            TStat = ""
    End Select
    
Gap = opt_Black.Value 'Überprüfen ob Blackmark oder Normal
    
    Select Case Gap
        Case 1
            BStat = "S"
        Case Else
            BStat = ""
    End Select
    
ELab = Chr$(13) & Chr$(10) & "^@" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "O" _
& TStat & BStat & Chr$(13) & Chr$(10) & "N" & Chr$(13) & Chr$(10) & "D10" _
& Chr$(13) & Chr$(10) & "S2" & Chr$(13) & Chr$(10) & "ZT" & Chr$(13) & Chr$(10) & _
"Q" & dH & "," & dS & Chr$(13) & Chr$(10) & "q" & dB & Chr$(13) & _
Chr$(10) & "A20,20,0,2,3,3,N," & Chr(34) & " BSR " & Chr(34) & Chr$(13) & Chr$(10) & _
"A20,70,0,2,3,3,R," & Chr(34) & " idware " & Chr(34) & Chr$(13) & Chr$(10) & "P1" & Chr$(13) & Chr$(10)

frm_Main.Rtx.Text = ELab

Unload Me

Exit Sub

FehlerWert:
MsgBox err_epl_lab, vbCritical, "Fehler" 'Fehlermeldung


End Sub


Private Sub txt_B_KeyPress(KeyAscii As Integer)

'Eingabe überprüfen und nur Zahlen bzw. "." und "," zulassen
    Select Case KeyAscii
        Case 8
        Case 44
        Case 46
        Case 48 To 57
        Case Else: KeyAscii = 0
    End Select

End Sub

Private Sub txt_H_KeyPress(KeyAscii As Integer)

'Eingabe überprüfen und nur Zahlen bzw. "." und "," zulassen
    Select Case KeyAscii
        Case 8
        Case 44
        Case 46
        Case 48 To 57
        Case Else: KeyAscii = 0
    End Select

End Sub


Private Sub txt_S_KeyPress(KeyAscii As Integer)
'Eingabe überprüfen und nur Zahlen bzw. "." und "," zulassen
    Select Case KeyAscii
        Case 8
        Case 44
        Case 46
        Case 48 To 57
        Case Else: KeyAscii = 0
    End Select
End Sub
