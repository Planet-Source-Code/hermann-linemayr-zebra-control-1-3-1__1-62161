Attribute VB_Name = "Fill_Items"
Public Function FillCOM()

With frm_Prn_Set
    With .cmb_Baud 'Baudrate
        .Clear
        .AddItem "1200"
        .AddItem "2400"
        .AddItem "4800"
        .AddItem "9600"
        .AddItem "19200"
        .AddItem "38400"
        .AddItem "57600"
        .AddItem "115200"
        .Text = "9600"
    End With
    
    With .cmb_DatBits 'Daten Bits
        .Clear
        .AddItem "7"
        .AddItem "8"
        .Text = "8"
    End With
    
    With .cmb_Par 'Parität
        .Clear
        .AddItem "Even"
        .AddItem "Odd"
        .AddItem "None"
        .AddItem "Mark"
        .AddItem "Space"
        .Text = "None"
    End With
    
    With .cmb_StopBits 'Stop Bits
        .Clear
        .AddItem "1"
        .AddItem "2"
        .Text = "1"
    End With
    
    With .txt_IP 'IP-Adresse
        .Text = "192.168.1.111"
    End With
    
    With .txt_Port 'Port
        .Text = "9100"
    End With
    
    With .opt_Port(0) 'LPT1 auswählen
        .Enabled = True
    End With

    'Variablen für den COM Standard setzen
    Baud = 9600
    DataBit = 8
    StopBit = 1
    Parity = "N"
    


End With

End Function
