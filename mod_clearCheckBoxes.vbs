
Private Sub ClearCheckBoxes()
'-----------------------------------------------------------------------------------------------
' [VBA]
'   Name........: ClearCheckBoxes:
'   Description.: limpeza dos checkbox para novo preenchimento
'   ------------:
'   Author......: Rafael Silva // rafael.dsilva@claro.com // rafael.neromad@gmail.com
'   Commentaries: Contato: (19) 991 704 394
'
'-----------------------------------------------------------------------------------------------
    Dim chkBox As Excel.CheckBox
    
    Application.ScreenUpdating = False
    
    For Each chkBox In FRONT.CheckBoxes
    
            chkBox.Value = xlOff
            
    Next chkBox
    
    FRONT.Range("E10") = ""
    FRONT.Range("E13") = ""
    FRONT.Range("E14") = ""
    FRONT.Range("E15") = ""
    FRONT.Range("I5") = ""
    FRONT.Range("L10") = ""
    FRONT.Range("L12") = "sem_email"
    FRONT.Range("L14") = ""
    FRONT.Range("L16") = ""
    FRONT.Range("L20") = ""
    FRONT.Range("L21") = ""
    FRONT.Range("L22") = ""
    
    Application.ScreenUpdating = True
    
End Sub