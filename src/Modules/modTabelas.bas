Attribute VB_Name = "modTabelas"

Public Function RetornaVencimento(ByVal Cargo_ As String, _
    Optional ByVal Data As Date = DATA_EM_ABERTO, _
    Optional ByVal Atualizado As Boolean, _
    Optional ByVal CargaHoraria As Integer, _
    Optional ByVal Subsidio As String) As clsVencimento

    Dim Cargo As clsCargo
    
    Set Cargo = New clsCargo
    
    Cargo.CriaObjeto Cargo_
    
    Set RetornaVencimento = New clsVencimento
    
    If Subsidio = "" Then
        Subsidio = RetornaVencimento.DefineSubisido(Data)
    End If
    
    If CargaHoraria = 0 Then
        CargaHoraria = Cargo.CargaHoraria
    End If
    
    With gsspSisap
        
        PesquisaSimboloVencimento

        .PrimeiroCampo
        .Envia (Cargo.SimboloVencimento)
        .ProximoCampo
        .Envia Subsidio
        .Enter
            
        Dim AchouVencimento As Boolean
        Dim Vencimento As clsVencimento
                
        AchouVencimento = False
        
        Do
            Linha = 10
            Coluna = 7
            
            For i = Linha To 20
            
                Set Vencimento = PegaVencimento(i)
                
                If Vencimento.Cargo.Grau = Cargo.Grau And _
                Vencimento.Cargo.CargaHoraria = CargaHoraria And _
                ValidacaoDataFinal(Atualizado, Vencimento, Data) Then
                    
                    Set RetornaVencimento = Vencimento
                    AchouCargoAtivo = True
                    Exit For
                 
                End If
                
            Next i
            
            If Not AchouPosicionamento Then
                If .F8(1, 9) = 9 Then
                    AchouPosicionamento = True
                End If
            End If
            
        Loop While AchouCargoAtivo = False
    End With
    
    RetornaVencimento.Subsidio = Subsidio
    
    Set Cargo = Nothing
End Function

Private Function ValidacaoDataFinal(ByVal Atualizado As Boolean, _
    ByRef Vencimento As clsVencimento, _
    ByVal Data As Date) As Boolean
    
    If Atualizado Then
        ValidacaoDataFinal = Vencimento.DataFinal = DATA_EM_ABERTO
    Else
         ValidacaoDataFinal = Vencimento.DataInicio <= Data And _
                (Vencimento.DataFinal >= Data Or Vencimento.DataFinal = DATA_EM_ABERTO)
    End If
End Function

Public Function PesquisaTabelasProdemge()
    Titulo = "PESQUISAR TABELAS PRODEMGE"
    With gsspSisap
        If Not .VerificaTituloTela(Titulo) Then
            PesquisaTabelasSisap
        End If
        
        .PrimeiroCampo
        .EnviaOpcao 21
        .Enter
        
    End With
End Function

Public Function PesquisaTabelasCarreiras()
    Titulo = "PESQUISAR TABELAS CARREIRAS"
    With gsspSisap
        If Not .VerificaTituloTela(Titulo) Then
            PesquisaTabelasProdemge
        End If
        
        .PrimeiroCampo
        .EnviaOpcao 9
        .Enter
        
    End With
End Function

Public Function PesquisaSimboloVencimento()
    Titulo = "PESQUISA SIMBOLO DE VENCIMENTO"
    With gsspSisap
        If Not .VerificaTituloTela(Titulo) Then
            PesquisaTabelasCarreiras
        Else
            .F9
        End If
    End With
End Function

Public Function PegaVencimento(ByVal Linha As Integer, Optional Coluna As Integer = 7) As clsVencimento

    Set PegaVencimento = New clsVencimento
    Set PegaVencimento.Cargo = New clsCargo
    
    With gsspSisap
    
                PegaVencimento.Cargo.Grau = .PegaCampo(1, Linha, Coluna)
                PegaVencimento.Cargo.CargaHoraria = .PegaCampoNumerico(2, Linha, Coluna + 7)
                PegaVencimento.DataInicio = .PegaData(10, Linha, Coluna + 48)
                PegaVencimento.DataFinal = .PegaData(10, Linha, Coluna + 62)
                PegaVencimento.Valor = .PegaCampoMoeda(13, True, Linha, Coluna + 30)
                'PegaVencimento.DefineSubisido (PegaVencimento.DataInicio)
                
                
    End With
End Function


