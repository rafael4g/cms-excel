Public Sub execUpdateData()
'-----------------------------------------------------------------------------------------------
' [VBA]
'   Name........: execUpdateData:
'   Description.: atualiza as linhas da base
'   ------------:
'   Author......: Rafael Silva // rafael.dsilva@claro.com // rafael.neromad@gmail.com
'   Commentaries: Contato: (19) 991 704 394
'
'-----------------------------------------------------------------------------------------------
Dim objLineUpdate As Integer

    If FRONT.Range("E3") = "" Then
    
        MsgBox "Selecione um Contrato...", vbCritical, "SALVANDO..."
     
        Exit Sub
    
    End If

    If FRONT.CheckBox1.Value Then
    
        objLineUpdate = handleHacked(FRONT.Range("B2").Value) 'Celula com INDEX para update de informação já alteradas
    
    Else
    
        objLineUpdate = FRONT.Range("B2").Value 'Celula com INDEX para update de nova informação
    
    End If

'data de atualizacao
BASE.Cells(objLineUpdate, 2) = Now()

'contado com sucesso coluna T = 20
    If FRONT.Range("AR1").Value > 1 Then
    
       FRONT.Range("D5").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "CONTATO COM SUCESSO"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO1"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 20) = "SIM"
             
        ElseIf Left(FRONT.Range("AP1"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 20) = "NÃO"
             
        Else
    
        BASE.Cells(objLineUpdate, 20) = "SEM PREENCHIMENTO"
        
        End If
        
    End If


'motivos coluna u = 21
    If FRONT.Range("AR2").Value > 1 Then
    
       FRONT.Range("D6").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "MOTIVOS"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO2"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 21) = "SEM INTERESSE"
             
        ElseIf Left(FRONT.Range("AP3"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 21) = "PEDIU RETORNO"
             
        ElseIf Left(FRONT.Range("AO4"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 21) = "REPOSICIONADO CRN"
             
        ElseIf Left(FRONT.Range("AP2"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 21) = "REPOSICIONADO PROJETO"
             
        ElseIf Left(FRONT.Range("AP3"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 21) = "CTR CANCELADO"
             
        ElseIf Left(FRONT.Range("AP4"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 21) = "CTR DESCONECTADO INVOLUNTARIO-INAD"
             
        ElseIf Left(FRONT.Range("AP5"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 21) = "CTR DESCONECTADO VOLUNTARIO-OPCAO"
             
        ElseIf Left(FRONT.Range("AS2"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 21) = "REPOSICIONADO CRN RETENCAO"
             
        ElseIf Left(FRONT.Range("AS3"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 21) = "SEM SUCESSO"
                          
        ElseIf Left(FRONT.Range("AS4"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 21) = "OFERTA AQUISICAO"
                                       
                          
        Else
        
        BASE.Cells(objLineUpdate, 21) = "SEM PREENCHIMENTO"
        
        End If
        
    End If

'produtos inclusos coluna v = 22
    BASE.Cells(objLineUpdate, 22) = FRONT.Range("E10").Value

'movimentação bl coluna w = 23
    If FRONT.Range("AR6").Value > 1 Then
    
       FRONT.Range("D11").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "MOVIMENTACAO BL"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO5"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 23) = "SIM"
             
        ElseIf Left(FRONT.Range("AP6"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 23) = "NÃO"
             
        Else
        
        BASE.Cells(objLineUpdate, 23) = "SEM PREENCHIMENTO"
        
        End If
        
    End If

'velocidade aceita coluna x = 24
    If FRONT.Range("AR7").Value > 1 Then
    
       FRONT.Range("D12").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "VELOCIDADE ACEITA"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO6"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 24) = "60M"
             
        ElseIf Left(FRONT.Range("AP7"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 24) = "120M"
             
        ElseIf Left(FRONT.Range("AP8"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 24) = "240M"
             
        Else
        
        BASE.Cells(objLineUpdate, 24) = "SEM PREENCHIMENTO"
        
        End If
        
    End If

'Delta's colunas y = 25, z = 26, ac = 29, ad = 30
    BASE.Cells(objLineUpdate, 25) = FRONT.Range("E13").Value 'BL
    
    BASE.Cells(objLineUpdate, 26) = FRONT.Range("E14").Value 'FONE
    
    BASE.Cells(objLineUpdate, 29) = FRONT.Range("E15").Value 'TV
    
    BASE.Cells(objLineUpdate, 30) = FRONT.Range("E16").Value 'TOTAL


'movimentação tv coluna aa = 27
    If FRONT.Range("AR9").Value > 1 Then
    
       FRONT.Range("D17").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "MOVIMENTACAO TV"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO7"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 27) = "SIM"
             
        ElseIf Left(FRONT.Range("AP9"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 27) = "SIM"
             
        Else
        
             BASE.Cells(objLineUpdate, 27) = "SEM PREENCHIMENTO"
             
        End If
        
    End If


'pacote de tv ab = 28
    If FRONT.Range("AR10").Value > 1 Then
    
       FRONT.Range("D18").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "PACOTE DE TV"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO8"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 28) = "INICIAL HD"
             
        ElseIf Left(FRONT.Range("AO9"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 28) = "NET FACIL TURBO HD"
             
        ElseIf Left(FRONT.Range("AP10"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 28) = "NET MIX HD"
             
        ElseIf Left(FRONT.Range("AP11"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 28) = "NET TOP HD"
             
        Else
        
             BASE.Cells(objLineUpdate, 28) = "SEM PREENCHIMENTO"
             
        End If
        
    End If


'possui celular de alguma operadora ae = 31
    If FRONT.Range("AR12").Value > 1 Then
    
       FRONT.Range("D20").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "POSSUI CELULAR"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO10"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 31) = "INICIAL HD"
             
        ElseIf Left(FRONT.Range("AO11"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 31) = "NET FACIL TURBO HD"
             
        ElseIf Left(FRONT.Range("AP12"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 31) = "NET MIX HD"
             
        ElseIf Left(FRONT.Range("AP13"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 31) = "NET TOP HD"
             
        Else
        
             BASE.Cells(objLineUpdate, 31) = "SEM PREENCHIMENTO"
             
        End If
        
    End If


'esta satisfeito com a operadora atual af = 32
    If FRONT.Range("AR14").Value > 1 Then
    
       FRONT.Range("K3").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "ESTA SATISFEITO"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO12"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 32) = "SIM"
             
        ElseIf Left(FRONT.Range("AP14"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 32) = "NÃO"
             
        Else
        
             BASE.Cells(objLineUpdate, 32) = "SEM PREENCHIMENTO"
             
        End If
        
    End If



'adquiriu movel ag = 33
    If FRONT.Range("AR15").Value > 1 Then
    
       FRONT.Range("K4").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "ADQUIRIU MOVEL"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO13"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 33) = "SIM"
             
        ElseIf Left(FRONT.Range("AP15"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 33) = "NÃO"
             
        Else
             BASE.Cells(objLineUpdate, 33) = "SEM PREENCHIMENTO"
             
        End If
        
    End If



'plano aceito ah = 34
    If FRONT.Range("AR16").Value > 1 Then
    
       FRONT.Range("K5").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "PLANO ACEITO"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO14"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 34) = "CONTROLE APP 3GB"
             
        ElseIf Left(FRONT.Range("AO15"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 34) = "CONTROLE APP 4GB"
             
        ElseIf Left(FRONT.Range("AO16"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 34) = "CONTROLE APP 5GB"
             
        ElseIf Left(FRONT.Range("AO17"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 34) = "8GB + MIN ILIMITADOS"
             
        ElseIf Left(FRONT.Range("AP16"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 34) = "10GB + MIN ILIMITADOS"
             
        ElseIf Left(FRONT.Range("AP17"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 34) = "15GB + MIN ILIMITADOS"
             
        ElseIf Left(FRONT.Range("AP18"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 34) = "50GB + MIN ILIMITADOS"
             
        ElseIf Left(FRONT.Range("AP19"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 34) = "100GB + MIN ILIMITADOS"
             
        Else
        
             BASE.Cells(objLineUpdate, 34) = "SEM PREENCHIMENTO"
             
        End If
        
    End If



'motivo de nao aceite celular ai = 35
    If FRONT.Range("AR20").Value > 1 Then
    
       FRONT.Range("K9").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "MOTIVO DE NAO ACEITE CELULAR"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO18"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 35) = "COMUNIDADE"
             
        ElseIf Left(FRONT.Range("AP20"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 35) = "FIDELIDADE"
             
        ElseIf Left(FRONT.Range("AP21"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 35) = "PRE-PAGO"
             
        Else
        
             BASE.Cells(objLineUpdate, 35) = "SEM PREENCHIMENTO"
             
        End If
        
    End If


'quantidade de ctto realizados colunas aj = 36
    BASE.Cells(objLineUpdate, 36) = FRONT.Range("L10").Value


'confirma email ak = 37
    If FRONT.Range("AR22").Value > 1 Then
    
       FRONT.Range("K11").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "CONFIRMA EMAIL CLIENTE"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO19"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 37) = "SIM"
             
        ElseIf Left(FRONT.Range("AP22"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 37) = "NAO"
             
        Else
        
             BASE.Cells(objLineUpdate, 37) = "SEM PREENCHIMENTO"
             
        End If
        
    End If


'melhor email para contato com cliente al = 38
    BASE.Cells(objLineUpdate, 38) = FRONT.Range("L12").Value

'cliente baixou app minha claro residencial am = 39
    If FRONT.Range("AR23").Value > 1 Then
    
       FRONT.Range("K13").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "APP MINHA CLARO RESIDENCIAL"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO20"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 39) = "SIM"
             
        ElseIf Left(FRONT.Range("AP23"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 39) = "NAO"
             
        Else
        
             BASE.Cells(objLineUpdate, 39) = "SEM PREENCHIMENTO"
             
        End If
        
    End If


'tem alguma sugestao para claro an = 40
    BASE.Cells(objLineUpdate, 40) = FRONT.Range("L14").Value

'cliente indimplente ao = 41
    If FRONT.Range("AR24").Value > 1 Then
    
       FRONT.Range("K13").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "APP MINHA CLARO RESIDENCIAL"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO21"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 41) = "SIM"
             
        ElseIf Left(FRONT.Range("AP24"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 41) = "NAO"
             
        Else
        
             BASE.Cells(objLineUpdate, 41) = "SEM PREENCHIMENTO"
             
        End If
        
    End If


'se inadimplente data de retorno ap = 42
    BASE.Cells(objLineUpdate, 42) = FRONT.Range("L16").Value



'necessario troca de aparelho aq = 43
    If FRONT.Range("AR25").Value > 1 Then
    
       FRONT.Range("K17").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "NECESSARIO TROCA DE APARELHO"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO22"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 43) = "SIM"
             
        ElseIf Left(FRONT.Range("AP25"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 43) = "NAO"
             
        Else
        
             BASE.Cells(objLineUpdate, 43) = "SEM PREENCHIMENTO"
             
        End If
        
    End If


'necessario troca de aparelho ar = 44
    If FRONT.Range("AR25").Value > 1 Then
    
       FRONT.Range("K17").Select
       
       MsgBox "Verificar check box assinalados...", vbCritical, "NECESSARIO TROCA DE APARELHO"
       
       Exit Sub
       
       Else
       
        If Left(FRONT.Range("AO23"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 44) = "VISITA TECNICA"
             
        ElseIf Left(FRONT.Range("AO24"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 44) = "ATIVACAO CHIP"
             
        ElseIf Left(FRONT.Range("AP26"), 1) = "V" Then
        
             BASE.Cells(objLineUpdate, 44) = "ISENCAO DE TAXA"
             
        Else
        
             BASE.Cells(objLineUpdate, 44) = "SEM PREENCHIMENTO"
             
        End If
        
    End If


'data agendamento as = 45
    BASE.Cells(objLineUpdate, 45) = FRONT.Range("L20").Value
    
'periodo at = 46
    BASE.Cells(objLineUpdate, 46) = FRONT.Range("L21").Value
    
'observação au = 47
    BASE.Cells(objLineUpdate, 47) = FRONT.Range("L22").Value

Application.DisplayAlerts = False
ActiveWorkbook.Save
Application.DisplayAlerts = True
MsgBox "ATUALIZADO", vbApplicationModal, "Salvando..."
Call ClearCheckBoxesStatic
Call ComboBox1.Clear

Call CheckBox1_Click

End Sub