Private Sub ComboBox1_Click()
'-----------------------------------------------------------------------------------------------
' [VBA]
'   Name........: ComboBox1_Click:
'   Description.: função para o click do ComboBox
'   ------------:
'   Author......: Rafael Silva // rafael.dsilva@claro.com // rafael.neromad@gmail.com
'   Commentaries: Contato: (19) 991 704 394
'
'-----------------------------------------------------------------------------------------------
    FRONT.Cells(3, 5) = ComboBox1.Value
    
    FRONT.Cells(5, 4).Select
    
    Call ClearCheckBoxesStatic
    
    If CheckBox1.Value Then
    
        FRONT.Cells(2, 2) = ComboBox1.ListIndex + 1
    
        Call populateCheckBoxUpdated
    
    Else
    
        FRONT.Cells(2, 2) = ComboBox1.ListIndex + 2
    
        Call populateCheckBox
    
    End If

End Sub