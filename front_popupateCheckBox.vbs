Public Sub populateCheckBox()
'-----------------------------------------------------------------------------------------------
' [VBA]
'   Name........: populateCheckBox:
'   Description.: atualiza os checkBoxs para consulta
'   ------------:
'   Author......: Rafael Silva // rafael.dsilva@claro.com // rafael.neromad@gmail.com
'   Commentaries: Contato: (19) 991 704 394
'
'-----------------------------------------------------------------------------------------------
Dim objLineUpdate As Integer: objLineUpdate = FRONT.Range("B2").Value 'Celula com INDEX para update de informação

If FRONT.Range("E3") = "" Then Exit Sub

'ultima atualização
FRONT.Range("i5").Value = BASE.Cells(objLineUpdate, 2)

'contado com sucesso coluna T = 20
If BASE.Cells(objLineUpdate, 20) = "SIM" Then

         FRONT.Range("AO1").Value = True
         
ElseIf BASE.Cells(objLineUpdate, 20) = "NÃO" Then

         FRONT.Range("AP1").Value = True
         
End If


'motivos coluna u = 21
    If BASE.Cells(objLineUpdate, 21) = "SEM INTERESSE" Then
    
        FRONT.Range("AO2") = True
        
    ElseIf BASE.Cells(objLineUpdate, 21) = "PEDIU RETORNO" Then
    
        FRONT.Range("AP3") = True
        
    ElseIf BASE.Cells(objLineUpdate, 21) = "REPOSICIONADO CRN" Then
    
        FRONT.Range("AO4") = True
         
    ElseIf BASE.Cells(objLineUpdate, 21) = "REPOSICIONADO PROJETO" Then
    
        FRONT.Range("AP2") = True
         
    ElseIf BASE.Cells(objLineUpdate, 21) = "CTR CANCELADO" Then
    
        FRONT.Range("AP3") = True
         
    ElseIf BASE.Cells(objLineUpdate, 21) = "CTR DESCONECTADO INVOLUNTARIO-INAD" Then
    
        FRONT.Range("AP4") = True
         
    ElseIf BASE.Cells(objLineUpdate, 21) = "CTR DESCONECTADO VOLUNTARIO-OPCAO" Then
    
        FRONT.Range("AP5") = True
        
    ElseIf BASE.Cells(objLineUpdate, 21) = "REPOSICIONADO CRN RETENCAO" Then
        
        FRONT.Range("AS2") = True
             
    ElseIf BASE.Cells(objLineUpdate, 21) = "SEM SUCESSO" Then
    
        FRONT.Range("AS3") = True
                          
    ElseIf BASE.Cells(objLineUpdate, 21) = "OFERTA AQUISICAO" Then
    
        FRONT.Range("AS4") = True
             
    End If


'produtos inclusos coluna v = 22
    FRONT.Range("E10").Value = BASE.Cells(objLineUpdate, 22)

'movimentação bl coluna w = 23
    If BASE.Cells(objLineUpdate, 23) = "SIM" Then
    
        FRONT.Range("AO5") = True
         
    ElseIf BASE.Cells(objLineUpdate, 23) = "NÃO" Then
    
        FRONT.Range("AP6") = True
    
    End If


'velocidade aceita coluna x = 24
    If BASE.Cells(objLineUpdate, 24) = "60M" Then
    
        FRONT.Range("AO6") = True
         
    ElseIf BASE.Cells(objLineUpdate, 24) = "120M" Then
    
        FRONT.Range("AP7") = True
         
    ElseIf BASE.Cells(objLineUpdate, 24) = "240M" Then
    
        FRONT.Range("AP8") = True
         
    End If
    


'Delta's colunas y = 25, z = 26, ac = 29, ad = 30
    FRONT.Range("E13") = BASE.Cells(objLineUpdate, 25).Value 'BL
    
    FRONT.Range("E14") = BASE.Cells(objLineUpdate, 26).Value 'FONE
    
    FRONT.Range("E15") = BASE.Cells(objLineUpdate, 29).Value 'TV
    'FRONT.Range("E16") = BASE.Cells(objLineUpdate, 30).Value 'TOTAL


'movimentação tv coluna aa = 27
    If BASE.Cells(objLineUpdate, 27) = "SIM" Then
    
        FRONT.Range("AO7") = True
         
    ElseIf BASE.Cells(objLineUpdate, 27) = "SIM" Then
    
        FRONT.Range("AP9") = True
         
    End If
    


'pacote de tv ab = 28
    If BASE.Cells(objLineUpdate, 28) = "INICIAL HD" Then
    
        FRONT.Range("AO8") = True
         
    ElseIf BASE.Cells(objLineUpdate, 28) = "NET FACIL TURBO HD" Then
    
        FRONT.Range("AO9") = True
         
    ElseIf BASE.Cells(objLineUpdate, 28) = "NET MIX HD" Then
    
        FRONT.Range("AP10") = True
         
    ElseIf BASE.Cells(objLineUpdate, 28) = "NET TOP HD" Then
    
        FRONT.Range("AP11") = True
         
    End If
    


'possui celular de alguma operadora ae = 31
    If BASE.Cells(objLineUpdate, 31) = "INICIAL HD" Then
    
        FRONT.Range("AO10") = True
         
    ElseIf BASE.Cells(objLineUpdate, 31) = "NET FACIL TURBO HD" Then
    
        FRONT.Range("AO11") = True
         
    ElseIf BASE.Cells(objLineUpdate, 31) = "NET MIX HD" Then
    
        FRONT.Range("AP12") = True
         
    ElseIf BASE.Cells(objLineUpdate, 31) = "NET TOP HD" Then
    
        FRONT.Range("AP13") = True
   
    End If
    


'esta satisfeito com a operadora atual af = 32
    If BASE.Cells(objLineUpdate, 32) = "SIM" Then
    
        FRONT.Range("AO12") = True
         
    ElseIf BASE.Cells(objLineUpdate, 32) = "NÃO" Then
    
        FRONT.Range("AP14") = True
         
    End If
    


'adquiriu movel ag = 33
    If BASE.Cells(objLineUpdate, 33) = "SIM" Then
    
        FRONT.Range("AO13") = True
         
    ElseIf BASE.Cells(objLineUpdate, 33) = "NÃO" Then
    
        FRONT.Range("AP15") = True
   
    End If
    


'plano aceito ah = 34
    If BASE.Cells(objLineUpdate, 34) = "CONTROLE APP 3GB" Then
    
        FRONT.Range("AO14") = True
         
    ElseIf BASE.Cells(objLineUpdate, 34) = "CONTROLE APP 4GB" Then
    
        FRONT.Range("AO15") = True
         
    ElseIf BASE.Cells(objLineUpdate, 34) = "CONTROLE APP 5GB" Then
    
        FRONT.Range("AO16") = True
         
    ElseIf BASE.Cells(objLineUpdate, 34) = "8GB + MIN ILIMITADOS" Then
    
        FRONT.Range("AO17") = True
         
    ElseIf BASE.Cells(objLineUpdate, 34) = "10GB + MIN ILIMITADOS" Then
    
        FRONT.Range("AP16") = True
         
    ElseIf BASE.Cells(objLineUpdate, 34) = "15GB + MIN ILIMITADOS" Then
    
        FRONT.Range("AP17") = True
         
    ElseIf BASE.Cells(objLineUpdate, 34) = "50GB + MIN ILIMITADOS" Then
    
        FRONT.Range("AP18") = True
         
    ElseIf BASE.Cells(objLineUpdate, 34) = "100GB + MIN ILIMITADOS" Then
    
        FRONT.Range("AP19") = True

    End If
    


'motivo de nao aceite celular ai = 35
    If BASE.Cells(objLineUpdate, 35) = "COMUNIDADE" Then
    
        FRONT.Range("AO18") = True
         
    ElseIf BASE.Cells(objLineUpdate, 35) = "FIDELIDADE" Then
    
        FRONT.Range("AP20") = True
         
    ElseIf BASE.Cells(objLineUpdate, 35) = "PRE-PAGO" Then
    
        FRONT.Range("AP21") = True

    End If
    


'quantidade de ctto realizados colunas aj = 36
    FRONT.Range("L10") = BASE.Cells(objLineUpdate, 36).Value


'confirma email ak = 37
    If BASE.Cells(objLineUpdate, 37) = "SIM" Then
    
        FRONT.Range("AO19") = True
         
    ElseIf BASE.Cells(objLineUpdate, 37) = "NAO" Then
    
        FRONT.Range("AP22") = True
         
    End If



'melhor email para contato com cliente al = 38
    FRONT.Range("L12") = BASE.Cells(objLineUpdate, 38).Value

'cliente baixou app minha claro residencial am = 39
    If BASE.Cells(objLineUpdate, 39) = "SIM" Then
    
        FRONT.Range("AO20") = True
         
    ElseIf BASE.Cells(objLineUpdate, 39) = "NAO" Then
    
        FRONT.Range("AP23") = True

    End If
    


'tem alguma sugestao para claro an = 40
     FRONT.Range("L14") = BASE.Cells(objLineUpdate, 40).Value

'cliente indimplente ao = 41
    If BASE.Cells(objLineUpdate, 41) = "SIM" Then
    
        FRONT.Range("AO21") = True
         
    ElseIf BASE.Cells(objLineUpdate, 41) = "NAO" Then
    
        FRONT.Range("AP24") = True

    End If
    


'se inadimplente data de retorno ap = 42
     FRONT.Range("L16") = BASE.Cells(objLineUpdate, 42).Value


'necessario troca de aparelho aq = 43
    If BASE.Cells(objLineUpdate, 43) = "SIM" Then
    
        FRONT.Range("AO22") = True
            
    ElseIf BASE.Cells(objLineUpdate, 43) = "NAO" Then
    
        FRONT.Range("AP25") = True

    End If


'necessario troca de aparelho ar = 44
    If BASE.Cells(objLineUpdate, 44) = "VISITA TECNICA" Then
    
        FRONT.Range("AO23") = True
         
    ElseIf BASE.Cells(objLineUpdate, 44) = "ATIVACAO CHIP" Then
    
        FRONT.Range("AO24") = True
         
    ElseIf BASE.Cells(objLineUpdate, 44) = "ISENCAO DE TAXA" Then
    
        FRONT.Range("AP26") = True

    End If


'data agendamento as = 45
     FRONT.Range("L20") = BASE.Cells(objLineUpdate, 45).Value
    
'periodo at = 46
     FRONT.Range("L21") = BASE.Cells(objLineUpdate, 46).Value
    
'observação au = 47
     FRONT.Range("L22") = BASE.Cells(objLineUpdate, 47).Value


End Sub