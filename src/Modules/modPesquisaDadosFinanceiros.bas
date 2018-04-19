Attribute VB_Name = "modPesquisaDadosFinanceiros"
Public Function DadosFinanceirosMesesAnteriores(Optional ByVal _
    Data As Date)
    
    If Data = DATA_VAZIA Then
        Data = Date
    End If
    
    Titulo = "DADOS FINANCEIROS MESES ANTERIORES"
    
    With gsspSisap
        
        If Not .VerificaTituloTela(Titulo) Then
            PesquisaDeDadosFinanceiros
        Else
            .F2
        End If
        
        .EnviaOpcao 3
        .Enter
        .EnviaMaspDv gdsvServidor.MaspDv
        .EnviaAdm gdsvServidor.Admisao
        .Enter
        'Span mes e ano
        .Envia Format(Data, "mmyyyy")
        .Enter
        'Seleciona o cargo
        IdentificarCargo 12, 21, 50, 7, Data, 57, 68, _
            "PESQUISA DADOS FINANCEIROS MESES ANTERIORES"
    End With
End Function

Public Function PesquisaHistoricoPagamento(Optional ByVal _
    Data As Date)
    
    PesquisaDeDadosFinanceiros
    
    With gsspSisap
        .EnviaOpcao 5
        .Enter
        .EnviaMaspDv gdsvServidor.MaspDv
        .Enter
         If Data <> DATA_VAZIA Then
            .Envia Format(Data, "mmyyyy")
         End If
        
        .Enter
        If .PegaCampo(8, 18, 4) = "SERVIDOR" Then
            .Enter
        End If
        
    End With
    

End Function



