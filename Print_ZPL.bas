Attribute VB_Name = "Print_ZPL"

Public Function PrintOut()

With frm_Main
        
    'Prüfen welcher Port zur Ausgabe gewählt wurde
    Select Case SelCom
        Case 1: GoTo PrintLPT1
        Case 2: GoTo PrintLPT2
        Case 3: GoTo PrintCOM1
        Case 4: GoTo PrintCOM2
        Case 5: GoTo PrintCOM3
        Case 6: GoTo PrintCOM4
        Case 7: GoTo PrintNet
        Case 8: GoTo PrintUSB
    End Select
    Exit Function
    
    
PrintLPT1:
        Open "LPT1" For Output As #1
            Print #1, .Rtx.Text
        Close #1
    Exit Function

PrintLPT2:
        Open "LPT2" For Output As #1
            Print #1, .Rtx.Text
        Close #1
    Exit Function
    
PrintCOM1:
        .MSCom.CommPort = 1
        .MSCom.Settings = COM_String
        .MSCom.PortOpen = True
        .MSCom.Output = .Rtx.Text
        .MSCom.PortOpen = False
    Exit Function
    
PrintCOM2:
        .MSCom.CommPort = 2
        .MSCom.Settings = COM_String
        .MSCom.PortOpen = True
        .MSCom.Output = .Rtx.Text
        .MSCom.PortOpen = False
    Exit Function
    
PrintCOM3:
        .MSCom.CommPort = 3
        .MSCom.Settings = COM_String
        .MSCom.PortOpen = True
        .MSCom.Output = .Rtx.Text
        .MSCom.PortOpen = False
    Exit Function

PrintCOM4:
        .MSCom.CommPort = 4
        .MSCom.Settings = COM_String
        .MSCom.PortOpen = True
        .MSCom.Output = .Rtx.Text
        .MSCom.PortOpen = False
    Exit Function
    
PrintNet:
        On Error GoTo FehlerV:
        'Daten an Socket übergeben und senden
        .MSSock.SendData .Rtx.Text
        DoEvents
    Exit Function

PrintUSB:
    On Error GoTo ErrorUSB:
        .Com_Dlg.CancelError = True
        .Com_Dlg.DialogTitle = "Drucker wählen"
        .Com_Dlg.PrinterDefault = False
        .Com_Dlg.Flags = cdlPDReturnDC
        .Com_Dlg.ShowPrinter
        .Rtx.SelPrint .Com_Dlg.hDC
    Exit Function

End With

ErrorUSB:
Exit Function

FehlerV:
MsgBox err_con, vbCritical, "Fehler"
frm_Prn_Set.Show

End Function
