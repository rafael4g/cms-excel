Public Sub widescreen()
'-----------------------------------------------------------------------------------------------
' [VBA]
'   Name........: widescreen:
'   Description.: tela cheia para melhor visualização do preenchimento
'   ------------:
'   Author......: Rafael Silva // rafael.dsilva@claro.com // rafael.neromad@gmail.com
'   Commentaries: Contato: (19) 991 704 394
'
'-----------------------------------------------------------------------------------------------
    Application.ScreenUpdating = False
    
    'Application.DisplayFormulaBar = True
    ActiveWindow.DisplayHeadings = True
    
    Application.DisplayFullScreen = False
    
    Range("A1").Select
    
End Sub