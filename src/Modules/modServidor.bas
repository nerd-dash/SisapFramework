Attribute VB_Name = "modServidor"

Public Sub ServidorLimpaDados()
    
    gdsvServidor.Nome = ""
    gdsvServidor.Cargo = ""
    gdsvServidor.Lotacao = ""
    gdsvServidor.Exercicio = ""
    gdsvServidor.SituacaoFuncional = ""
    gdsvServidor.SituacaoServidor = ""
    gdsvServidor.CodSituacaoFuncional = 0
    gdsvServidor.CodSituacaoServidor = 0
    gdsvServidor.DataAposentadoria = 0

End Sub

Public Sub ServidorBuscaNome()

    With gsspSisap
        NavDadosFuncionais
        gdsvServidor.Nome = .PegaCampo(72, 7, 9)
    End With
    
End Sub



Public Sub ServidorBuscaCargo(Optional ByVal DataReferencia As Date)
    
    With gsspSisap
    
        If EntraCargoAtivo(DataReferencia) Then
            gdsvServidor.Cargo = _
                .PegaCampo(7, 14, 46) & _
                .PegaCampo(1, 14, 65)
        End If
    End With
    
    Debug.Print "Buscando Cargo " & gdsvServidor.Cargo

End Sub

Sub ServidorBuscaExercicio()
    
    With gsspSisap
        PesquisarServidor
        If EntraCargoAtivo Then
            .MarcarOpcao 20, 61
            .Enter
        End If
    End With

End Sub

Sub RotinaPegaLotacao(Optional ByVal DataReferencia As Date)
    
    gdsvServidor.Lotacao = RetornaUnidadeAdministrativadeLotacao(DataReferencia)

End Sub

Sub RotinaPegaExercicio(Optional ByVal DataReferencia As Date)
    
    gdsvServidor.Exercicio = RetornaUnidadeAdministrativadeExercico(DataReferencia).UnidadeAdmNome

End Sub

Public Function EntraCargoAtivo( _
    Optional ByVal DataReferencia As Date) As Boolean
   
    With gsspSisap
        
        NavDadosFuncionais
        
        gdsvServidor.CodSituacaoFuncional = .PegaCampoNumerico(2, 9, 18)
        gdsvServidor.CodSituacaoServidor = .PegaCampoNumerico(2, 10, 18)
        gdsvServidor.SituacaoFuncional = .PegaCampo(35, 9, 23)
        gdsvServidor.SituacaoServidor = .PegaCampo(57, 10, 23)
        
        EntraCargoAtivo = IdentificarCargo(14, 18, 24, 7, DataReferencia)
   End With
End Function



Public Sub PesquisarAfastamentos(Optional ByVal MaspDv As Long = 0, _
            Optional ByVal Admissao As Integer = 0)
            
    If MaspDv = 0 Then
        MaspDv = gdsvServidor.MaspDv
    End If
    
    If Admissao = 0 Then
        Admissao = gdsvServidor.Admisao
    End If

    Titulo = "PESQUISAR AFASTAMENTOS DO SERVIDOR"
    With gsspSisap
        Do While Not (.PegaCampo(34, 4, 24) = Titulo _
        And .PegaCampo(12, 7, 24) = "AFASTAMENTOS")
            PesquisaDadosServidor
            .EnviaOpcao 2
            .Enter 1, 2
        Loop
        
        .EnviaOpcao 1
        .Enter
        .PrimeiroCampo
        .EnviaMaspDv MaspDv
        .EnviaAdm Admissao
        .Enter 'Tem que pegar servidor inexistente
        
    End With
    
End Sub
Public Function EvolucaoCarreira( _
    Optional ByVal DataReferencia As Date)

    Titulo = "PESQUISAR SERVIDOR PUBLICO"
   
    With gsspSisap
        If Not .VerificaTituloTela(Titulo) Then
            EntraCargoAtivo DataReferencia
        End If
        .MarcarOpcao 21, 19
        .Enter 1, 205
    End With

End Function

Public Function PesquisarCargaHoraria()

    Titulo = "PESQUISAR CARGA HORARIA"
   
    With gsspSisap
        If Not .VerificaTituloTela(Titulo) Then
            CargaHorariaSEE
            .EnviaOpcao 7
            .Enter
        Else
            .F9
        End If
    End With

End Function

Public Function PesquisarCargaHorariaVigente( _
    Optional ByVal Data As Date = DATA_EM_ABERTO, _
    Optional ByVal MaspDv As Long = 0, _
    Optional ByVal Admissao As Integer = 0, _
    Optional ByVal ColunaDataInicial As Integer, _
    Optional ByVal ColunaDataFinal As Integer)

    If Data = DATA_EM_ABERTO Then
        Data = Date
    End If
    
    If MaspDv = 0 Then
        MaspDv = gdsvServidor.MaspDv
    End If
    
    If Admissao = 0 Then
        Admissao = gdsvServidor.Admisao
    End If

    Titulo = "PESQUISAR C. H. DE SERVIDOR POR UNIDADE VIGENTE EM: " _
    & Format(Data, "mm") & "/" & Format(Data, "yyyy")
    
    With gsspSisap
        If Not .VerificaTituloTela(Titulo) Then
            PesquisarCargaHoraria
            .EnviaOpcao 2
            .Enter
            .Envia Format(Data, "mmyyyy")
            .Enter
            .EnviaMaspDv MaspDv
            .ProximoCampo 8
            .EnviaAdm Admissao
            .Enter
            IdentificarCargo 12, 20, 24, 8, Data, _
                ColunaDataInicial, ColunaDataFinal
        End If
        
    End With

End Function


Public Function ServidorBuscaPosicionamento( _
    Optional ByVal Data As Date = DATA_EM_ABERTO) As String

       
    Titulo = "PESQUISAR HISTORICO DE CARGOS SERVIDOR"
    
    If Data = DATA_EM_ABERTO Then
        Data = Date
    End If
       
    With gsspSisap
        If Not .VerificaTituloTela(Titulo) Then
            EvolucaoCarreira Data
        Else
            .F2
            EvolucaoCarreira Data
        End If
        
        Dim AchouPosicionamento As Boolean
        AchouPosicionamento = False
        
        Do
            
            For Linha = 11 To 16 Step 5
            
                
                DataInicio = .PegaData(10, Linha, 50)
                DataFinal = .PegaData(10, Linha, 71)
                
                If Data >= DataInicio And Data <= DataFinal Then
                    ServidorBuscaPosicionamento = .PegaCampo(8, Linha + 2, 35) _
                        & .PegaCampo(1, Linha + 2, 56)
                    AchouPosicionamento = True
                    Exit For
                End If
                
            Next Linha
            
             If Not AchouPosicionamento Then
                If .F8(1, 9) = 9 Then
                    .JanelaErro "Não foi possível encontrar o Posicionamento na data informada! Por favor verifique!"
                    .EncerraSisap
                End If
            End If
            
        Loop While AchouPosicionamento = False
    
    End With
        
End Function

Public Function ServidorBuscaCargaHoraria( _
    Optional ByVal Data As Date) _
        As clschsCargaHoraria

    Set ServidorBuscaCargaHoraria = New clschsCargaHoraria
    
    If Data = DATA_VAZIA Then
        Data = Date
    End If

    PesquisarCargaHorariaVigente Data, 0, 0, 52, 65
    
    With gsspSisap

        Dim AchouPosicionamento As Boolean
        Dim CargaHoraria As New clsCargaHoraria
        
        AchouPosicionamento = False
        
        LinhaInicial = 7
        LinhaFinal = 21
        
        Do
            For i = LinhaInicial To LinhaFinal Step 5
                Set CargaHoraria = PegaCargaHoraria(i, 15)
                
                If CargaHoraria.Tipo <> 0 Then
                
                    Debug.Print "Carga Horaria Tipo " & CargaHoraria.Tipo & " " & _
                    CargaHoraria.Descricao & " " & _
                    CargaHoraria.QuantidadeAulas & " hrs"
                    
                    ServidorBuscaCargaHoraria.Add CargaHoraria
                    
                End If
               
            Next i
             
            If Not AchouPosicionamento Then
                If .F8(1, 9) = 9 Then
                    AchouPosicionamento = True
                End If
            End If

    
        Loop While AchouPosicionamento = False
        
        .F2
                
    End With
    
End Function


Private Function PegaCargaHoraria(ByVal LinhaInicial As Integer, _
                                    ByVal ColunaInicial As Integer) As clsCargaHoraria
            
            Set PegaCargaHoraria = New clsCargaHoraria
        
            With gsspSisap
                
            
                Set PegaCargaHoraria = New clsCargaHoraria
                PegaCargaHoraria.CodGrupo = .PegaCampoNumerico(2, LinhaInicial, ColunaInicial + 2)
                
                PegaCargaHoraria.CodNatureza = .PegaCampoNumerico(3, LinhaInicial, ColunaInicial + 5)
                PegaCargaHoraria.Descricao = .PegaCampo(55, LinhaInicial, ColunaInicial + 10)
                
                PegaCargaHoraria.DataInicio = .PegaData(14, LinhaInicial + 1, ColunaInicial + 2)
                PegaCargaHoraria.DataFimPrevisto = .PegaData(14, LinhaInicial + 1, 42)
                PegaCargaHoraria.DataFimEfetivo = .PegaData(14, LinhaInicial + 1, 66)
                
                PegaCargaHoraria.Tipo = .PegaCampoNumerico(2, LinhaInicial + 2, ColunaInicial)
                PegaCargaHoraria.Nivel = .PegaCampoNumerico(2, LinhaInicial + 2, ColunaInicial + 16)
                PegaCargaHoraria.Modalidade = .PegaCampoNumerico(2, LinhaInicial + 2, ColunaInicial + 26)
                PegaCargaHoraria.Materia = .PegaCampoNumerico(5, LinhaInicial + 2, ColunaInicial + 38)
                PegaCargaHoraria.QuantidadeAulas = .PegaCampoNumerico(2, LinhaInicial + 2, ColunaInicial + 54)
                PegaCargaHoraria.Turno = .PegaCampoNumerico(2, LinhaInicial + 2, ColunaInicial + 64)
                
                PegaCargaHoraria.UnidadeAdministrativa = .PegaCampoNumerico(8, LinhaInicial + 3, ColunaInicial)
                PegaCargaHoraria.SubstitutoMaspDv = .PegaMaspDv(LinhaInicial + 3, ColunaInicial + 21)
                PegaCargaHoraria.SubstitutoAdm = .PegaCampoNumerico(1, LinhaInicial + 3, ColunaInicial + 41)
                PegaCargaHoraria.SubstitutoGrupoNatureza = .PegaCampoNumerico(2, LinhaInicial + 3, ColunaInicial + 61)
                
                
            End With
    
End Function


Private Function PegaUnidadeAdministrativa(ByVal LinhaInicial As Integer, ByVal ColunaInicial As Integer, _
                                Optional ByVal Exercicio As Boolean) As clsUnidadeAdministrativa
            
            Set PegaUnidadeAdministrativa = New clsUnidadeAdministrativa
        
            With gsspSisap
            
                PegaUnidadeAdministrativa.NaturezaGrupo = .PegaCampoNumerico(2, LinhaInicial, ColunaInicial)
                PegaUnidadeAdministrativa.NaturezaTipo = .PegaCampoNumerico(3, LinhaInicial, ColunaInicial + 3)
                PegaUnidadeAdministrativa.NaturezaDescricao = .PegaCampo(58, LinhaInicial, ColunaInicial + 9)
               
                PegaUnidadeAdministrativa.DataPublicacao = .PegaData(14, LinhaInicial + 2, ColunaInicial + 1)
               
                If Exercicio Then
                    PegaUnidadeAdministrativa.DataProrrogacao = .PegaData(14, LinhaInicial + 2, ColunaInicial + 20)
                    PegaUnidadeAdministrativa.DataInicio = .PegaData(14, LinhaInicial + 2, ColunaInicial + 40)
                    PegaUnidadeAdministrativa.DataFinal = .PegaData(14, LinhaInicial + 2, ColunaInicial + 57)
                Else
                    PegaUnidadeAdministrativa.DataInicio = .PegaData(14, LinhaInicial + 2, ColunaInicial + 24)
                    PegaUnidadeAdministrativa.DataFinal = .PegaData(14, LinhaInicial + 2, ColunaInicial + 44)
                    
                End If
                    
                
                PegaUnidadeAdministrativa.RegionalCod = .PegaCampoNumerico(2, LinhaInicial + 6, ColunaInicial)
                PegaUnidadeAdministrativa.RegionalNome = .PegaCampo(50, LinhaInicial + 6, ColunaInicial + 5)
                
                PegaUnidadeAdministrativa.UnidadeAdmCod = .PegaCampoNumerico(7, LinhaInicial + 7, ColunaInicial + 2)
                PegaUnidadeAdministrativa.UnidadeAdmNome = .PegaCampo(43, LinhaInicial + 7, ColunaInicial + 12)
                PegaUnidadeAdministrativa.ZonaRural = .PegaCampo(1, LinhaInicial + 7, ColunaInicial + 65) = "S"

                
                PegaUnidadeAdministrativa.MunicipioCod = .PegaCampoNumerico(4, LinhaInicial + 8, ColunaInicial + 1)
                PegaUnidadeAdministrativa.MunicipioNome = .PegaCampo(43, LinhaInicial + 8, ColunaInicial + 8)
            
                PegaUnidadeAdministrativa.DistritoCod = .PegaCampoNumerico(2, LinhaInicial + 9, ColunaInicial + 1)
                PegaUnidadeAdministrativa.DistritoNome = .PegaCampo(50, LinhaInicial + 9, ColunaInicial + 5)

                                
            End With
    
End Function


Public Function IdentificarCargo(ByVal LinhaInicial As Integer, _
                ByVal LinhaFinal As Integer, _
                Optional ByVal ColunaNatureza As Integer = 24#, _
                Optional ByVal ColunaOpcao As Integer = 8#, _
                Optional ByVal Data As Date = DATA_EM_ABERTO, _
                Optional ByVal ColunaDataInicial As Integer = 54#, _
                Optional ByVal ColunaDataFinal As Integer = 67#, _
                Optional ByVal Titulo As String) As Boolean
    
    Dim strDate As String
    
    If Titulo = vbNullString Then
        Titulo = "IDENTIFICAR CARGO"
    End If
    
    Dim AchouCargoAtivo As Boolean
    AchouCargoAtivo = False
    ContinuaProcurarCargo = True
    
    With gsspSisap
    
        If .VerificaTituloTela(Titulo) Then
             
            If Data = DATA_VAZIA Then
                Data = DATA_EM_ABERTO
            End If
            
            Debug.Print "Buscando Cargo em : " & Data
            
            If gdsvServidor.CodSituacaoServidor = 7 Then
            
                If gdsvServidor.DataAposentadoria = DATA_VAZIA Then
                   Do
                        strDate = InputBox("Por favor entre com a data da Aposentadoria:", _
                        "Servidor Aposentado", Format(Now(), "dd/mm/yyyy"))
                        If IsDate(strDate) Then
                            Data = DateAdd("d", -1, CDate(strDate))
                            gdsvServidor.DataAposentadoria = CDate(strDate)
                        ElseIf strDate = vbNullString Then
                            gsspSisap.EncerraSisap
                        Else
                            MsgBox "Data Inválida"
                        End If
                    Loop While gdsvServidor.DataAposentadoria <> CDate(strDate)
                Else
                    Data = DateAdd("d", -1, gdsvServidor.DataAposentadoria)
                End If
                
            ElseIf gdsvServidor.CodSituacaoServidor = 2 Then
            
                ContinuaProcurarCargo = True
                Data = .PegaData(10, LinhaInicial, ColunaDataFinal)
                Debug.Print "O Servidor está desligado no momento"
                
            End If
            
            Do While ContinuaProcurarCargo = True
                For i = LinhaInicial To LinhaFinal
                    Natureza = .PegaCampoNumerico(2, i, ColunaNatureza)
                    DataInicio = .PegaData(10, i, ColunaDataInicial)
                    DataFinal = .PegaData(10, i, ColunaDataFinal)
                    
                    If (Natureza = 6 Or Natureza = 7) _
                    And DataEstaEntre(Data, _
                        DataInicio, _
                        DataFinal) Then
                        
                        .MarcarOpcao i, ColunaOpcao
                        .Enter
                        AchouCargoAtivo = True
                        ContinuaProcurarCargo = False
                        Exit For
                    ElseIf Natureza = 0 Then
                        AchouCargoAtivo = True
                        ContinuaProcurarCargo = False
                        Exit For
                    End If
                Next i
                
                If ContinuaProcurarCargo Then
                    If .F8(1, 205) = 9 Then
                        .JanelaErro "Não foi possível encontrar um cargo Ativo para o Servidor! Por favor verifique!"
                        .EncerraSisap
                    End If
                End If
            Loop
        End If
        
   End With
   
   IdentificarCargo = AchouCargoAtivo

End Function

Public Function RetornaUnidadeAdministrativadeExercico( _
    Optional ByVal DataReferencia As Date) As clsUnidadeAdministrativa

    Set RetornaUnidadeAdministrativadeExercico = New clsUnidadeAdministrativa
    
    With gsspSisap
           
        EntraCargoAtivo DataReferencia
        
        If gdsvServidor.CodSituacaoServidor <> 2 Then
            .MarcarOpcao 20, 61
            .Enter 1, 187
            Set RetornaUnidadeAdministrativadeExercico = PegaUnidadeAdministrativa(9, 13, True)
        End If
    End With
    
    Debug.Print "Número de Unidade Adminstrativa da Exercicio : " & RetornaUnidadeAdministrativadeExercico.UnidadeAdmCod
End Function

Public Function RetornaUnidadeAdministrativadeLotacao( _
    Optional ByVal DataReferencia As Date) As Long
    
    With gsspSisap
           
        EntraCargoAtivo DataReferencia
        
        If gdsvServidor.CodSituacaoServidor <> 2 Then
            .MarcarOpcao 21, 3
            .Enter 1, 187
            RetornaUnidadeAdministrativadeLotacao = _
            PegaUnidadeAdministrativa(9, 13, False).UnidadeAdmCod
        End If
    End With
    
    Debug.Print "Número de Unidade Adminstrativa da Lotação : " & RetornaUnidadeAdministrativadeLotacao
End Function



Public Function BuscaDadosAtuaisServidor()
            Application.ScreenUpdating = False
                If gdsvServidor.MaspDv > 0 And _
                    gdsvServidor.Admisao > 0 Then
                    'Servidor.LimpaFormulario
                    ServidorBuscaNome
                    ServidorBuscaCargo
                    RotinaPegaLotacao
                    RotinaPegaExercicio
                    If gdsvServidor.CodSituacaoServidor = 7 Then
                        gdsvServidor.MostraLinha 10
                        gfpcFPConversao.DataAposentadoria = _
                        gdsvServidor.DataAposentadoria
                    Else
                        gdsvServidor.EscondeLinha 10
                    End If
                End If
            Application.ScreenUpdating = True
End Function


