Private Sub Workbook_Open()
'-----------------------------------------------------------------------------------------------
' [VBA]
'   Name........: Workbook_Open:
'   Description.: Popula as formulas da aba BASE
'   ------------:
'   Author......: Rafael Silva // rafael.dsilva@claro.com // rafael.neromad@gmail.com
'   Commentaries: Contato: (19) 991 704 394
'
'-----------------------------------------------------------------------------------------------
'Call handleIncludeFormule

  Dim i As Long: i = 2
  
    FRONT.ComboBox1.Clear
    
    Do While BASE.Cells(i, 17) <> ""
       
        FRONT.ComboBox1.AddItem BASE.Cells(i, 17).Value
       
    i = i + 1
    Loop
    
    Call server.handleIncludeFormule
    
End Sub