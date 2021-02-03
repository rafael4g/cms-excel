Function handleFormule(cel As Range)
'-----------------------------------------------------------------------------------------------
' [VBA]
'   Name........: handleFormule:
'   Description.: pega a formula com a linguagem DAX
'   ------------:
'   Author......: Rafael Silva // rafael.dsilva@claro.com // rafael.neromad@gmail.com
'   Commentaries: Contato: (19) 991 704 394
'
'-----------------------------------------------------------------------------------------------
handleFormule = cel.FormulaR1C1
End Function