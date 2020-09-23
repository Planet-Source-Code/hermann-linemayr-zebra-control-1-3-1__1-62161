Attribute VB_Name = "Port_Functions"
Public Function SetForeColorNet()

'Den Netz Textfeldern eine graue Schrift verpassen
With frm_Prn_Set
    .txt_Port.ForeColor = &HC0C0C0
    .txt_IP.ForeColor = &HC0C0C0
End With

End Function

Public Function SetForeColorCom()

'Den COM Feldern eine graue Schrift verpassen
With frm_Prn_Set
    .cmb_Baud.ForeColor = &HC0C0C0
    .cmb_DatBits.ForeColor = &HC0C0C0
    .cmb_StopBits.ForeColor = &HC0C0C0
    .cmb_Par.ForeColor = &HC0C0C0
End With

End Function

Public Function SetForeColorComB()

'Den COM Feldern eine graue Schrift verpassen
With frm_Prn_Set
    .cmb_Baud.ForeColor = &H80000008
    .cmb_DatBits.ForeColor = &H80000008
    .cmb_StopBits.ForeColor = &H80000008
    .cmb_Par.ForeColor = &H80000008
End With

End Function

Public Function get_COM_Par()

With frm_Prn_Set
    
    'Parität in Kurzzeichen an die Variable Par übergeben
    Select Case .cmb_Par.ListIndex
        Case 0: Parity = "E" 'Even
        Case 1: Parity = "O" 'Odd
        Case 2: Parity = "N" 'None
        Case 3: Parity = "M" 'Mark
        Case 4: Parity = "S" 'Space
    End Select
    
    'Stop Bit an die Variable StopBit übergeben
    StopBit = .cmb_StopBits.Text
    
    'Daten Bit an die Variable DataBit übergeben
    DataBit = .cmb_DatBits.Text
    
    'Baudrate an die Variable Baud übergeben
    Baud = .cmb_Baud.Text

    'Den COM_String aus den einzelnen Parametern zusammensetzen
    COM_String = Baud & "," & Parity & "," & DataBit & "," & StopBit
    
    'Die Bezeichnung für die Statusbar übergeben
    Select Case SelCom
        Case 3: COM_Bez = "COM1"
        Case 4: COM_Bez = "COM2"
        Case 5: COM_Bez = "COM3"
        Case 6: COM_Bez = "COM4"
    End Select


'COM Port und Parameter in die Statusbar schreiben
frm_Main.StatBar.Panels.Item(1).Text = COM_Bez & " " & COM_String

'Neue Werte in die Registry schreiben
CreateKey "HKCU\Software\BSR\ZPL2\Baud", .cmb_Baud.Text
CreateKey "HKCU\Software\BSR\ZPL2\Par", .cmb_Par.Text
CreateKey "HKCU\Software\BSR\ZPL2\Dat", .cmb_DatBits.Text
CreateKey "HKCU\Software\BSR\ZPL2\Stop", .cmb_StopBits.Text


End With

End Function

Public Function SetNet()
On Error GoTo VFehler:

With frm_Prn_Set
    'Variablen übergeben
    IP = .txt_IP.Text
    Port = .txt_Port.Text
End With
    
With frm_Main
    'Wenn offene Verbindung dann schließen
    If .MSSock.State <> 0 Then .MSSock.Close
    
    'Verbindung aufbauen
    .MSSock.Connect IP, Port
End With

With frm_Main
    .StatBar.Panels.Item(1).Text = "IP " & frm_Prn_Set.txt_IP.Text & ":" & frm_Prn_Set.txt_Port.Text
End With

Exit Function
    
VFehler:
    MsgBox err_con, vbCritical, "Fehler!"

End Function
