Attribute VB_Name = "modDesignacao"
Private mcrg As clsCargo
Sub IncluirDesignacao(Optional ByVal Teste As Boolean = False)

    Titulo = "INCLUIR DESIGNACAO"
    
    Set mcrg = New clsCargo
    
    mcrg.CriaObjeto gdsgDesigncao.Cargo
    
    
    
    VerficaCamposObrigatorios
    
    If Not gsspSisap.VerificaTituloTela(Titulo) Then
        Designacoes
    End If
    
    Tela_01
    
    Tela_02
    
    Tela_03
    
    ConfirmaCargaHoraria Teste
    
    Set mcrg = Nothing
End Sub

Private Function VerficaCamposObrigatorios()

    With gdsgDesigncao
    
        If .UnidadeAdministrativa.UnidadeAdmCod < 100 Then
            gsspSisap.JanelaErro ("A unidade administrativa não é válida!")
            gsspSisap.EncerraSisap
        ElseIf .SituacaoExercicio = 0 Then
            gsspSisap.JanelaErro ("A situação de exercício não é válida!")
            gsspSisap.EncerraSisap
        ElseIf .DataInicial = 0 Then
            gsspSisap.JanelaErro ("A data inicial não pode estar em branco!")
            gsspSisap.EncerraSisap
        ElseIf .DataFinal = 0 Then
            gsspSisap.JanelaErro ("A data final não pode estar em branco!")
            gsspSisap.EncerraSisap
         ElseIf .CargasHorarias.Count = 0 Then
            gsspSisap.JanelaErro ("Verifique a Carga Horária inserida!")
            gsspSisap.EncerraSisap
        End If
    
    End With

End Function

Private Function Tela_01()
    
    With gsspSisap
        
        .EnviaOpcao 1
        .Enter
        .EnviaMaspDv gdsvServidor.MaspDv
        .ProximoCampo 6
        .EnviaAdm gdsvServidor.Admisao
        .EnviaData gdsgDesigncao.DataInicial
        .EnviaData gdsgDesigncao.DataFinal
        
        Do
            .Enter 1, 997
        Loop Until .PegaCampo(8, 10, 3) = "NATUREZA"
    
    End With
    

End Function

Private Function Tela_02()

       
    With gsspSisap
    
        
        .EnviaNumero gdsgDesigncao.CargasHorarias.Item(1).CodNatureza, 3
        .ProximoCampo 3
        .Envia mcrg.Carreira
        .ProximoCampo
        .Envia mcrg.Nivel
        .ProximoCampo
        .Envia mcrg.SimboloVencimento
        .ProximoCampo
        .Envia mcrg.Grau
        .ProximoCampo
        .EnviaNumero gdsgDesigncao.CategoriaProfisisonal, 5
        .EnviaNumero gdsgDesigncao.SituacaoExercicio, 2
        .EnviaNumero gdsgDesigncao.UnidadeAdministrativa.MunicipioCod, 4
        .Envia "MG"
        .EnviaNumero gdsgDesigncao.UnidadeAdministrativa.UnidadeAdmCod, 8
        .Envia gdsgDesigncao.UnidadeAdministrativa.ZonaRural
        .ProximoCampo
        .EnviaNumero 12
        .EnviaNumero mcrg.ClassificaoOrcamentaria( _
            gdsgDesigncao.CargasHorarias.Item(1).Nivel, _
            gdsgDesigncao.CargasHorarias.Item(1).Modalidade), 12
        
        .EnviaNumero 1
        Do
            .Enter
        Loop Until .VerificaTituloTela("INFORMAR CARGA HORARIA")
    End With
    

End Function

Private Function Tela_03()

    If mcrg.Carreira = "PEB" Or mcrg.Carreira = "EEB" Then
        CargaHorariaMagisterio
    Else
        CargaHorariaAdministrativo
    End If

End Function

Private Function GeraSubstituto(CargaHoraria As clsCargaHoraria) As Boolean
    
    NaturezasSubstituicao = intArray(2, 8, 10, 19, 44, 92, 35, 70, _
        84, 86, 88, 90, 84, 26, 32, 37, 53, 57, 77)
    
    For i = 0 To UBound(NaturezasSubstituicao)
        GeraSubstituto = _
            IIf(NaturezasSubstituicao(i) = CargaHoraria.CodNatureza _
                And CargaHoraria.CodGrupo = 7, True, False)
                
        If GeraSubstituto Then
            Exit For
        End If
    Next i
    
End Function

Private Function PreencheSubstituto(ByVal CargaHoraria As clsCargaHoraria)
    With gsspSisap
            If GeraSubstituto(CargaHoraria) Then
            
                If gdsgDesigncao.SubstituidoMaspDv = 0 Then
                    .JanelaErro "É obrigatório um substituto!"
                    .EncerraSisap
                End If
                
                .EnviaMaspDv gdsgDesigncao.SubstituidoMaspDv
                .EnviaAdm gdsgDesigncao.SubstituidoAdmissao
                .EnviaNumero gdsgDesigncao.SubstituidoGrupoNatureza, 2
            Else
                .ProximoCampo 4
            End If
    End With
End Function

Private Function CargaHorariaMagisterio()

    With gsspSisap
        For i = 1 To gdsgDesigncao.CargasHorarias.Count
     
            Set CargaHoraria = gdsgDesigncao.CargasHorarias.Item(i)
        
            If i <> 1 Then
                .EnviaNumero CargaHoraria.CodGrupo, 2
                .EnviaNumero CargaHoraria.CodNatureza, 3
                .EnviaData gdsgDesigncao.DataInicial
                .EnviaData gdsgDesigncao.DataFinal
            End If
            
            .EnviaNumero CargaHoraria.Tipo, 2
            .EnviaNumero CargaHoraria.Nivel, 2
            .EnviaNumero CargaHoraria.Modalidade, 2
            .EnviaNumero CargaHoraria.Materia, 5
            .EnviaNumero CargaHoraria.QuantidadeAulas, 2
            .EnviaNumero CargaHoraria.Turno, 2
            
            If i <> 1 Then 'Pula Unidade administrativa
                .ProximoCampo
            End If
            
            PreencheSubstituto CargaHoraria
            
            If i Mod 3 = 0 Then
                .F8
            End If
        
        Next i
        
        If i > 3 Then
            .F7 (i \ 3) 'calcula páginas que tem que voltar
        End If
    End With

End Function

Private Function ConfirmaCargaHoraria(Optional ByVal Teste As Boolean = False)
        With gsspSisap
            If .Enter(1, 8) = 8 Then
                If Not Teste Then
                    .JanelaInformacao "Confira os dados de Designação antes de confirmar a inclusão."
                    
                    If gdsgDesigncao.DataInicial < DateSerial(Year(Date), Month(Date), 1) Then
                        .JanelaAlerta "O servidor tem Acerto para ser conferido!"
                        Debug.Print "Total de Aulas RB :"; gdsgDesigncao.CargasHorarias.TotalRB
                    End If
                End If
            End If
        End With
End Function

Private Function CargaHorariaAdministrativo()
    With gsspSisap
    
        Set CargaHoraria = gdsgDesigncao.CargasHorarias.Item(1)
        
        .EnviaNumero CargaHoraria.Tipo, 2
        .EnviaNumero CargaHoraria.QuantidadeAulas, 2
        .EnviaNumero CargaHoraria.Turno, 2
        
        PreencheSubstituto CargaHoraria
                        
    End With
End Function

Public Function IncluirDesligamentoDesignado()
    
    Dim CodNatureza As Integer
    Dim DataDesligamento As Date

On Error GoTo ValorInvalido

    NavIncluirDesligamentoDesignado

ValorInvalido:
    resposta = InputBox("Código da Natureza do Desligamento", _
        "Desligamento", "10")
        
    CodNatureza = CInt(resposta)
    
    strDate = InputBox("Por favor entre com a data do Desligamento:", _
                        "Desligamento", Format(Now(), "dd/mm/yyyy"))
    DataDesligamento = CDate(strDate)
    
    gsspSisap.PrimeiroCampo
    gsspSisap.EnviaNumero CodNatureza, 3
    gsspSisap.ProximoCampo 3
    gsspSisap.EnviaData DataDesligamento
    If gsspSisap.Enter(1, 8) = 8 Then
        gsspSisap.JanelaAlerta "Confirme os dados de Desligamento"
    End If
    
End Function
Public Function EnviaCargaHorariaParaAcerto()

    gdsgDesigncao.TotalAulasRB = gdsgDesigncao.CargasHorarias.TotalRB
    gdsgDesigncao.TotalAulasEC = gdsgDesigncao.CargasHorarias.TotalEC

End Function

Public Function EnviaVerbasDeAcerto()

    NavLancamentoCargoCodigoRecebimento
    
    EnviaVerbaRB
    EnviaVerbaEC
    EnviaVerbaAbonoRB
    EnviaVerbaAbonoEC
End Function

Private Function EnviaVerbaRB()
    If Not gdsgDesigncao.VerbaRB = 0 Then
        EncontraProximaVerba
        With gsspSisap
            .Incluir
            .EnviaNumero gdsgDesigncao.VerbaRB
            .EnviaData PrimerioDiaMes
            .ProximoCampo 4
            .Envia Format(gdsgDesigncao.VerbaRBValor, "###,###.00")
            .ProximoCampo 2
            .Envia "?"
            .Enter 2, 0
            .Envia Format(gdsgDesigncao.VerbaRBValor, "#.00")
            .F5 1, 415
        End With
    End If
End Function

Private Function EnviaVerbaEC()

    If Not gdsgDesigncao.VerbaEC = 0 Then
        EncontraProximaVerba
        With gsspSisap
            .Incluir
            .EnviaNumero gdsgDesigncao.VerbaEC
            .EnviaData PrimerioDiaMes
            .ProximoCampo 4
            .Envia Format(gdsgDesigncao.VerbaECValor, "###,###.00")
            .ProximoCampo 2
            .Envia "?"
            .Enter 2, 0
            .Envia Format(gdsgDesigncao.VerbaECValor, "#.00")
            .F5 1, 415
        End With
    End If
End Function
Private Function EnviaVerbaAbonoRB()
    If Not gdsgDesigncao.VerbaAbonoRB = 0 Then
        EncontraProximaVerba
         With gsspSisap
            .Incluir
            .EnviaNumero gdsgDesigncao.VerbaAbonoRB
            .EnviaData PrimerioDiaMes
            .ProximoCampo 4
            .Envia Format(gdsgDesigncao.VerbaAbonoRBValor, "###,###.00")
            .ProximoCampo 2
            .Envia "?"
            .Enter 2, 0
            .Envia Format(gdsgDesigncao.VerbaAbonoRBValor, "#.00")
            .F5 1, 415
        End With
    End If
End Function
Private Function EnviaVerbaAbonoEC()
    If Not gdsgDesigncao.VerbaAbonoEC = 0 Then
        EncontraProximaVerba
        With gsspSisap
          .Incluir
          .EnviaNumero gdsgDesigncao.VerbaAbonoEC
          .EnviaData PrimerioDiaMes
          .ProximoCampo 4
          .Envia Format(gdsgDesigncao.VerbaAbonoECValor, "###,###.00")
          .ProximoCampo 2
          .Envia "?"
          .Enter 2, 0
          .Envia Format(gdsgDesigncao.VerbaAbonoECValor, "#.00")
          .F5 1, 415
        End With
    End If
End Function
Private Function EncontraProximaVerba()
    Dim Verba As Long
    Dim i As Long
    
    i = 9
    
    Verba = -1
    
    Do
        Verba = gsspSisap.PegaVerba(i, 5)
        If Not Verba = 0 Then
            gsspSisap.ProximaLinha
            i = i + 1
        End If
    Loop While Not Verba = 0
End Function
