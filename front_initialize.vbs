Private Sub UserForm_Initialize()
'-----------------------------------------------------------------------------------------------
' [VBA]
'   Name........: UserForm_Initialize:
'   Description.: Popula o COMBO-BOX
'   ------------:
'   Author......: Rafael Silva // rafael.dsilva@claro.com // rafael.neromad@gmail.com
'   Commentaries: Contato: (19) 991 704 394
'
'-----------------------------------------------------------------------------------------------
  Dim i As Long: i = 2
  
  Dim objText As String
  
    Call ClearCheckBoxesStatic
    
    ComboBox1.Clear
    
    Do While BASE.Cells(i, 17) <> ""
    
        objText = ""
        
        If BASE.Cells(i, 1) <> "" Then
        
            If Left(BASE.Cells(i, 20), 1) = "N" Then
            
                objText = "*" & BASE.Cells(i, 17).Value & "*"
        
                ComboBox1.AddItem objText
            
            Else
            
                objText = "-" & BASE.Cells(i, 17).Value & "-"
        
                ComboBox1.AddItem objText
            
            End If
        
        Else
        
        ComboBox1.AddItem BASE.Cells(i, 17).Value
        
        End If
        
    i = i + 1
    
    Loop
 

End Sub