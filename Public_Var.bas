Attribute VB_Name = "Public_Var"
'Variablen für COM-Port
Public Baud, DataBit, StopBit, SelCom As Integer
Public Parity, COM_String, COM_Bez As String
Public IP, Port As String

'Sonstige Variablen
Public Pass As String
Public FileName As String
Public La As Byte

'Variablen für die ZPL Strings
Public ZCont, ZNCont As String
Public ZTear, ZPeel, ZRewind, ZCutter, ZApp As String
Public ZDirect, ZTrans As String
Public ZEng, ZSpa, ZSpa2, ZFra, ZDeu, ZIta, ZNor, ZPor, ZSch, ZDän, ZHol, ZFin, ZJap As String
Public ZIniFl, ZMedCal, ZEnDu, ZDiDu, ZMedPro, ZSet6, ZSet12, ZConf, ZMemList As String
Public Z128, Z39, Z25 As String
Public ZMot As String
Public ZSTr, ZSRe, ZSAu As String

'Variablen für die EPL Strings
Public ECon, ECal, ELab As String

'Variablen für Sprache
Public err_epl_lab As String
Public err_com_set As String
Public err_con As String





Public Function ZPL_Var()
'Continuous
ZCont = "^XA" & Chr$(13) & Chr$(10) & "^MNN" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Non Continuous
ZNCont = "^XA" & Chr$(13) & Chr$(10) & "^MNY" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"

'Tear Off
ZTear = "^XA" & Chr$(13) & Chr$(10) & "^MMT" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Peel
ZPeel = "^XA" & Chr$(13) & Chr$(10) & "^MMP" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Rewind
ZRewind = "^XA" & Chr$(13) & Chr$(10) & "^MMR" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Cutter
ZCutter = "^XA" & Chr$(13) & Chr$(10) & "^MMC" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Applicator
ZApp = "^XA" & Chr$(13) & Chr$(10) & "^MMA" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"

'Direct Thermal
ZDirect = "^XA" & Chr$(13) & Chr$(10) & "^MTD" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Thermal Transfer
ZTrans = "^XA" & Chr$(13) & Chr$(10) & "^MTT" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"

'English
ZEng = "^XA" & Chr$(13) & Chr$(10) & "^KL1" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Spanisch
ZSpa = "^XA" & Chr$(13) & Chr$(10) & "^KL2" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Französisch
ZFra = "^XA" & Chr$(13) & Chr$(10) & "^KL3" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Deutsch
ZDeu = "^XA" & Chr$(13) & Chr$(10) & "^KL4" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Italienisch
ZIta = "^XA" & Chr$(13) & Chr$(10) & "^KL5" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Norwegisch
ZNor = "^XA" & Chr$(13) & Chr$(10) & "^KL6" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Portugisisch
ZPor = "^XA" & Chr$(13) & Chr$(10) & "^KL7" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Schwedisch
ZSch = "^XA" & Chr$(13) & Chr$(10) & "^KL8" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Dänisch
ZDän = "^XA" & Chr$(13) & Chr$(10) & "^KL9" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Spanisch 2
ZSpa2 = "^XA" & Chr$(13) & Chr$(10) & "^KL10" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Holländisch
ZHol = "^XA" & Chr$(13) & Chr$(10) & "^KL11" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Finnisch
ZFin = "^XA" & Chr$(13) & Chr$(10) & "^KL12" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Japanisch
ZJap = "^XA" & Chr$(13) & Chr$(10) & "^KL13" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"

'Init Flash Memory
ZIniFl = "^XA" & Chr$(13) & Chr$(10) & "^JBE" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Media Calibration
ZMedCal = "^XA" & Chr$(13) & Chr$(10) & "~JC" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Enable Dump
ZEnDu = "^XA" & Chr$(13) & Chr$(10) & "~JD" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Disable Dump
ZDiDu = "^XA" & Chr$(13) & Chr$(10) & "~JE" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Media Profile
ZMedPro = "^XA" & Chr$(13) & Chr$(10) & "~JG" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'12 dot
ZSet12 = "^XA" & Chr$(13) & Chr$(10) & "^JMA" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'6 dot
ZSet6 = "^XA" & Chr$(13) & Chr$(10) & "^JMB" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Config
ZConf = "~WC"
'Memory List
ZMemList = "^XA" & Chr$(13) & Chr$(10) & "^WDB:*.*" & Chr$(13) & Chr$(10) & "^XZ" & Chr$(13) & Chr$(10) & _
"^XA" & Chr$(13) & Chr$(10) & "^WDE:*.*" & Chr$(13) & Chr$(10) & "^XZ" & Chr$(13) & Chr$(10) & _
"^XA" & Chr$(13) & Chr$(10) & "^WDR:*.*" & Chr$(13) & Chr$(10) & "^XZ" & Chr$(13) & Chr$(10) & _
"^XA" & Chr$(13) & Chr$(10) & "^WDZ:*.*" & Chr$(13) & Chr$(10) & "^XZ"

'Code128
Z128 = "^XA" & Chr$(13) & Chr$(10) & "^FO50,50" & Chr$(13) & Chr$(10) & "^BY3^BCN,100,Y,N,N" & Chr$(13) & Chr$(10) & "^FD1234ABCD^FS" & Chr$(13) & Chr$(10) & "^XZ"
'Code39
Z39 = "^XA" & Chr$(13) & Chr$(10) & "^FO50,50" & Chr$(13) & Chr$(10) & "^BY3^B3N,N,150,Y,N" & Chr$(13) & Chr$(10) & "^FD123ABC^FS" & Chr$(13) & Chr$(10) & "^XZ"
'Inteleaved 2/5
Z25 = "^XA" & Chr$(13) & Chr$(10) & "^FO50,50" & Chr$(13) & Chr$(10) & "^BY3^B2N,150,Y,N,N" & Chr$(13) & Chr$(10) & "^FD123456^FS" & Chr$(13) & Chr$(10) & "^XZ"

'Transmissive Sensor
ZSTr = "^XA" & Chr$(13) & Chr$(10) & "^JST" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Reflective Sensor
ZSRe = "^XA" & Chr$(13) & Chr$(10) & "^JSR" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"
'Autosensing
ZSAu = "^XA" & Chr$(13) & Chr$(10) & "^JSA" & Chr$(13) & Chr$(10) & "^JUS" & Chr$(13) & Chr$(10) & "^XZ"

'##############################################

'Ab hier EPL

'Print Config
ECon = Chr$(13) & Chr$(10) & "U" & Chr$(13) & Chr$(10)
'Calibrate
ECal = Chr$(13) & Chr$(10) & "^@" & Chr$(13) & Chr$(10) & "N" & Chr$(13) & Chr$(10) & "ZB" & Chr$(13) & Chr$(10) & _
"A300,10,0,3,1,1,N," & Chr$(34) & "Reset durchgeführt" & Chr$(34) & Chr$(13) & Chr$(10) & _
"A300,30,0,3,1,1,N," & Chr$(34) & "Sensoren werden" & Chr$(34) & Chr$(13) & Chr$(10) & _
"A300,60,0,3,1,1,N," & Chr$(34) & "eingemessen!" & Chr$(34) & Chr$(13) & Chr$(10) & _
"P1" & Chr$(13) & Chr$(10) & "xa" & Chr$(13) & Chr$(10) & "Y96,N,8,1" & Chr$(13) & Chr$(10) & _
"N" & Chr$(13) & Chr$(10) & "ZB" & Chr$(13) & Chr$(10) & _
"A300,10,0,3,1,1,N," & Chr$(34) & "RS232 >>> 96,n,8,1" & Chr$(34) & Chr$(13) & Chr$(10) & _
"A300,30,0,3,1,1,N," & Chr$(34) & "Sensor eingemessen" & Chr$(34) & Chr$(13) & Chr$(10) & _
"P1" & Chr$(13) & Chr$(10) & Chr$(13) & Chr$(10) & "EPL2" & Chr$(13) & Chr$(10) & _
"N" & Chr$(13) & Chr$(10) & "ZB" & Chr$(13) & Chr$(10) & _
"A300,10,0,3,1,1,N," & Chr$(34) & "EPL2 Mode" & Chr$(34) & Chr$(13) & Chr$(10) & _
"A300,30,0,3,1,1,N," & Chr$(34) & "eingeschalten" & Chr$(34) & Chr$(13) & Chr$(10) & _
"P1" & Chr$(13) & Chr$(10) & "U" & Chr$(13) & Chr$(10) & "N" & Chr$(13) & Chr$(10) & _
"ZB" & Chr$(13) & Chr$(10) & _
"A300,10,0,3,1,1,N," & Chr$(34) & "Gerät mit Netzschalter" & Chr$(34) & Chr$(13) & Chr$(10) & _
"A300,30,0,3,1,1,N," & Chr$(34) & "AUS und EINSCHALTEN" & Chr$(34) & Chr$(13) & Chr$(10) & _
"A300,60,0,3,1,1,N," & Chr$(34) & "!! Danach ist das !!" & Chr$(34) & Chr$(13) & Chr$(10) & _
"A300,90,0,3,1,1,N," & Chr$(34) & "!!  Gerät bereit  !!" & Chr$(34) & Chr$(13) & Chr$(10) & _
"P1" & Chr$(13) & Chr$(10)




End Function

'Registry ein- und auslesen
Public Function CreateKey(Folder As String, Value As String)

Dim x As Object
On Error Resume Next
Set x = CreateObject("wscript.shell")
x.RegWrite Folder, Value

End Function

Public Function ReadKey(Value As String) As String

Dim x As Object
On Error Resume Next
Set x = CreateObject("wscript.shell")
y = x.RegRead(Value)
ReadKey = y
End Function

Public Function Replace(oldstring, newletter, oldletter) As String

    Dim i As Integer
    i = 1


    Do While InStr(i, oldstring, oldletter, vbTextCompare) <> 0
        Replace = Replace & Mid(oldstring, i, InStr(i, oldstring, oldletter, vbTextCompare) - i) & newletter
        i = InStr(i, oldstring, oldletter, vbTextCompare) + Len(oldletter)
    Loop

    Replace = Replace & Right(oldstring, Len(oldstring) - i + 1)
    
   
End Function
