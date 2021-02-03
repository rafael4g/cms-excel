
Private Sub CheckBox1_Click()
'-----------------------------------------------------------------------------------------------
' [VBA]
'   Name........: CheckBox1_Click:
'   Description.: função para o click do checkbox
'   ------------:
'   Author......: Rafael Silva // rafael.dsilva@claro.com // rafael.neromad@gmail.com
'   Commentaries: Contato: (19) 991 704 394
'
'-----------------------------------------------------------------------------------------------
  Dim i As Long: i = 2
  
  Dim objText As String
  
    FRONT.Range("E3") = ""
    
    Call ClearCheckBoxesStatic
    
    ComboBox1.Clear
    
    If CheckBox1.Value Then
    
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
                
            End If
            
        i = i + 1
        
        Loop
        
    Else
        
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
   
    End If
   
End Sub