Public Sub handleIncludeFormule()
'-----------------------------------------------------------------------------------------------
' [VBA]
'   Name........: handleIncludeFormule:
'   Description.: preenche a coluna A na sheet "BASE" para controle usado no preenchimento
'   ------------:
'   Author......: Rafael Silva // rafael.dsilva@claro.com // rafael.neromad@gmail.com
'   Commentaries: Contato: (19) 991 704 394
'
'-----------------------------------------------------------------------------------------------
BASE.Select
Dim i As Integer: i = 2
'/*Colando valores para atualizar a formula da coluna*/
Dim linha As Long: linha = BASE.Range("Q1048576").End(xlUp).Row
BASE.Range("A2:A" & linha + 1).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Delete Shift:=xlUp

Do While BASE.Cells(i, 17) <> ""

    BASE.Cells(i, 1).FormulaR1C1 = "=IF( LEN(RC[1])>2, COUNTIF(R2C[1]:RC[1],""<>""&""""), """" )"
    
    i = i + 1
    
Loop

FRONT.Select

End Sub