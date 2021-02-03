Function handleHack(objQuery As Integer) As Integer
'-----------------------------------------------------------------------------------------------
' [VBA]
'   Name........: handleHack:
'   Description.: função para identificar a linha correta para o update
'   ------------:
'   Author......: Rafael Silva // rafael.dsilva@claro.com // rafael.neromad@gmail.com
'   Commentaries: Contato: (19) 991 704 394
'
'-----------------------------------------------------------------------------------------------
Dim z As Integer: z = 2
    
    Do While BASE.Cells(z, 17) <> ""
    
        If BASE.Cells(z, 1) = objQuery Then
        
        handleHack = z
        
        Exit Function
        
        End If
        
        z = z + 1
        
    Loop

End Function
