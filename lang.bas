Attribute VB_Name = "lang"
Public Function Lan_Ger()

'Menüführung Deutsch

'Main Form
With frm_Main
    
    'Menü Datei
    .File.Caption = "&Datei"
    .New.Caption = "&Neu"
    .Open.Caption = "Ö&ffnen"
    .Save.Caption = "&Speichern"
    .PrinterSet.Caption = "Drucker&einrichtung"
    .Print.Caption = "&Drucken"
    .Exit.Caption = "&Beenden"
    
    'Menü ZPL
    
    'Menü Tools
        .zpl.Caption = "&ZPL Skripte"
        .PrConf.Caption = "&Konfigurations Ausdruck"
        .MemoryList.Caption = "Speicherauf&listung"
        .MediaCal.Caption = "&Material kalibrieren"
        .MedProfile.Caption = "&Sensorprofil drucken"
        .EnDump.Caption = "Dump Modus &aktivieren"
        .DisDump.Caption = "Dump Modus &deaktivieren"
        .InitFlash.Caption = "&Flash Speicher löschen"
        .dot6.Caption = "Emulierung für &6 Dot"
        .Dot12.Caption = "Emulierung für &12 Dot"
        .PassChange.Caption = "&Passwort ändern"
        .TestCode.Caption = "&Testcode drucken"
    
    'Menü Media Tracking
    .Med_Track.Caption = "Medien &Art"
        .Cont.Caption = "&Endlos"
        .Non_Cont.Caption = "&Nicht Endlos"
    
    'Menü Media Type
    .Med_Type.Caption = "&Druckmethode"
        .trans.Caption = "Thermo &Transfer"
        .Direct.Caption = "Thermo &Direkt"
    
    'Menü Print Mode
    .PM.Caption = "Druckm&odus"
        .Tear.Caption = "Abb&reissen"
        .Peel.Caption = "Ab&schälen"
        .Rewind.Caption = "Auf&wickeln"
        .Cutter.Caption = "Absch&neiden"
        .Applicator.Caption = "&Applikator"
        
    'Menü für Sprache
    .Lang.Caption = "&Sprache"
        .English.Caption = "&Englisch"
        .German.Caption = "&Deutsch"
        .Italyan.Caption = "&Italienisch"
        .Spanish.Caption = "&Spanisch"
        .Spanish2.Caption = "Sp&anisch 2"
        .Protugise.Caption = "&Portugiesisch"
        .France.Caption = "&Französisch"
        .Norwayan.Caption = "N&orwegisch"
        .Swedish.Caption = "Sch&wedisch"
        .Finnisch.Caption = "Fi&nnisch"
        .Dansk.Caption = "Dänis&ch"
        .Netherlandish.Caption = "&Holländisch"
        .Japanish.Caption = "&Japanisch"
        
    'Menü Sensor
    .setsens.Caption = "S&ensor Auswahl"
        .transmi.Caption = "&Emitter/Empfänger"
        .refle.Caption = "6Reflexion"
        .asel.Caption = "&Autom. Auswahl"
        
    'Menü EPL
    .epl.Caption = "&EPL Skripte"
        .pr_Conf.Caption = "&Konfigurations Ausdruck"
        .set_Label.Caption = "&Etiketteneinrichtung"
        .cali.Caption = "&Material kalibrieren"
        
    'Menü Sprache
    .Men_Lan.Caption = "&Sprache"
        .M_Ger.Caption = "&Deutsch"
        .M_Eng.Caption = "&Englisch"
        
    'Menü Info
        .About.Caption = "Ü&ber.."
        
'-------------------------------------------

    'Tool Tip für Toolbarbuttons
    .img_New.ToolTipText = "neues Skript"
    .img_Open.ToolTipText = "Skript öffnen"
    .img_Print.ToolTipText = "Skript drucken"
    .img_Prop.ToolTipText = "Druckkereigenschaften"
    .img_Save.ToolTipText = "Skript speichern"
    
End With

'-------------------------------------------

'EPL Label Form
With frm_epl_lab
    
    .Caption = "Label erstellen"
    .Label1(0).Caption = "Breite (mm)"
    .Label2(0).Caption = "Höhe (mm)"
    .Label3(0).Caption = "Spalt (mm)"
    
    .Label1(1).Caption = "Zwischenraum"
    .Label2(1).Caption = "Breite"
    .Label3(1).Caption = "Höhe"
    
    .opt_Therm.Caption = "&Thermo Direkt"
    .CancelButton.Caption = "&Abbrechen"

End With

'-------------------------------------------

'Motion Form
With frm_Motion
    
    .Caption = "Motions"
    .Label1.Caption = "Einschalten"
    .Label2.Caption = "Kopf schließen"
    .CancelButton.Caption = "&Abbrechen"
    
End With

'-------------------------------------------

'Pass Form
With frm_Pass
    
    .Caption = "Passwort"
    .Label1.Caption = "Passwort (0000 - 9999)"
    .cmd_Set.Caption = "&OK"

End With

'-------------------------------------------

'Printer Setting Form
With frm_Prn_Set

    .Caption = "Druckereinstellungen"
    .Frame_COM_Set.Caption = "COM Einstellungen"
    .Frame_Port.Caption = "Anschluss"
    .Label3.Caption = "Parität"
    .Label5.Caption = "Adresse"
    .cmd_Cancel.Caption = "&Abbrechen"

End With

End Function

Public Function Lan_Eng()

'Menüführung Englisch

'Main Form
With frm_Main
    'Menu Datei
    .File.Caption = "&File"
    .New.Caption = "&New"
    .Open.Caption = "&Open"
    .Save.Caption = "&Save"
    .PrinterSet.Caption = "Print&ersettings"
    .Print.Caption = "&Print"
    .Exit.Caption = "E&xit"
    
    'Mneü ZPL
        'Menü Tools
        .zpl.Caption = "&ZPL Scripts"
        .PrConf.Caption = "Print &Config"
        .MemoryList.Caption = "List &Memory"
        .MediaCal.Caption = "Medi&a Calibration"
        .MedProfile.Caption = "Me&dia Profile"
        .EnDump.Caption = "&Enable Dump Mode"
        .DisDump.Caption = "Disable D&ump Mode"
        .InitFlash.Caption = "&Init Flash Memory"
        .dot6.Caption = "&Set 6 Dot"
        .Dot12.Caption = "Se&t 12 Dot"
        .PassChange.Caption = "Change &Password"
        .TestCode.Caption = "Print Testc&ode"
    
    'Menü Media Tracking
    .Med_Track.Caption = "Media T&racking"
        .Cont.Caption = "&Continuous"
        .Non_Cont.Caption = "&Non Continuous"
        
    'Menü Media Type
    .Med_Type.Caption = "Media T&ype"
        .trans.Caption = "&Thermal Transfer"
        .Direct.Caption = "&Direct Thermal"
    
    'Menü Print Mode
    .PM.Caption = "&Print Mode"
        .Tear.Caption = "&Tear Off"
        .Peel.Caption = "&Peel Off"
        .Rewind.Caption = "&Rewind"
        .Cutter.Caption = "&Cutter"
        .Applicator.Caption = "&Applicator"
        
    'Menü für Sprache
    .Lang.Caption = "L&anguage"
        .English.Caption = "&English"
        .German.Caption = "&German"
        .Italyan.Caption = "&Italien"
        .Spanish.Caption = "&Spanish"
        .Spanish2.Caption = "S&panish 2"
        .Protugise.Caption = "&Portuguese"
        .France.Caption = "&French"
        .Norwayan.Caption = "&Norwegian"
        .Swedish.Caption = "S&wedish"
        .Finnisch.Caption = "Finnis&h"
        .Dansk.Caption = "&Danish"
        .Netherlandish.Caption = "D&utch"
        .Japanish.Caption = "&Japan"
 
    'Menü Sensor
    .setsens.Caption = "&Set Sensor"
        .transmi.Caption = "&Transmissive"
        .refle.Caption = "&Reflective"
        .asel.Caption = "&Auto Select"
        
    'Menü EPL
    .epl.Caption = "&EPL Scripts"
        .pr_Conf.Caption = "&Print Config"
        .set_Label.Caption = "&Set Label"
        .cali.Caption = "Media &Calibration"
        
    'Menü Sprache
    .Men_Lan.Caption = "&Language"
        .M_Ger.Caption = "&German"
        .M_Eng.Caption = "&English"
        
    'Menü Info
        .About.Caption = "&About..."

'-------------------------------------------

    'Tool Tip für Toolbarbuttons
    .img_New.ToolTipText = "New Script"
    .img_Open.ToolTipText = "Open Script"
    .img_Print.ToolTipText = "Print Script"
    .img_Prop.ToolTipText = "Printersettings"
    .img_Save.ToolTipText = "Save Script"

End With

'-------------------------------------------

'EPL Label Form
With frm_epl_lab
    
    .Caption = "Create a label"
    .Label1(0).Caption = "Width (mm)"
    .Label2(0).Caption = "Height (mm)"
    .Label3(0).Caption = "Space (mm)"
    
    .Label1(1).Caption = "Space"
    .Label2(1).Caption = "Width"
    .Label3(1).Caption = "Height"
    
    .opt_Therm.Caption = "&Direct Thermal"
    .CancelButton.Caption = "&Cancel"

End With

'-------------------------------------------

'Motion Form
With frm_Motion
    
    .Caption = "Power Up / Head Close Motion"
    .Label1.Caption = "Power Up"
    .Label2.Caption = "Head Close"
    .CancelButton.Caption = "&Cancel"
    
End With
    
'-------------------------------------------

'Pass Form
With frm_Pass
    
    .Caption = "Password"
    .Label1.Caption = "Password (0000 - 9999)"
    .cmd_Set.Caption = "&OK"

End With
    
'-------------------------------------------

'Printer Setting Form
With frm_Prn_Set

    .Caption = "Communication Setting"
    .Frame_COM_Set.Caption = "COM Settings"
    .Frame_Port.Caption = "Connection"
    .Label3.Caption = "Parity"
    .Label5.Caption = "Adress"
    .cmd_Cancel.Caption = "&Cancel"

End With
    
End Function
