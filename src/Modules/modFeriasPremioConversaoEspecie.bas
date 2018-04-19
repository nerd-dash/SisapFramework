Attribute VB_Name = "modFeriasPremioConversaoEspecie"
Public Sub FPConversaoCalcula()
        
        If gfpcFPConversao.SaldoTotalEmDias > 0 And _
            gfpcFPConversao.DataPublicacao <> DATA_VAZIA Then

            Application.ScreenUpdating = False
            Debug.Print "Buscando dados de Posicionamento para Conversão FP"
            FPConversaoBuscaPosicionamentoDataAfastamento
            
            Debug.Print "Buscando dados de Carga Horária para Conversão FP"
            modFeriasPremioConversaoEspecie.FPConversaoBuscaCargaHoraria
            
            Debug.Print "Buscando dados de Vencimento para Conversão FP"
            modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB
            
            Debug.Print "Abrindo Verbas Atuais para preenchimento"
            DadosFinanceirosMesesAnteriores
            
            gsspSisap.JanelaAlerta "Verifique as verbas do servidor."
            Application.ScreenUpdating = True
            
        Else
            gsspSisap.JanelaErro "Verifique a data de publicação e o saldo de FP!"
        End If
    
End Sub
Public Function FPConversaoBuscaCargaHoraria()
    
    Dim CodsRB As New Collection

    Dim CargasHorarias As clschsCargaHoraria
    Dim VesperaAfastamento As Date
    
    
    With gfpcFPConversao
    
    VesperaAfastamento = DateAdd("d", -1, .DataAposentadoria)
    
    Set CargasHorarias = ServidorBuscaCargaHoraria(VesperaAfastamento)
    
        .CargaHorariaRB = CargasHorarias.TotalRB(VesperaAfastamento)
        .CargaHorariaEC = CargasHorarias.TotalEC(VesperaAfastamento)
        .CargaHorariaEX = CargasHorarias.TotalEX(VesperaAfastamento)
        .CargaHorariaECEX = CargasHorarias.TotalECEX(VesperaAfastamento)
        
    End With
End Function

Public Function FPConversaoBuscaVencimentoRB( _
    Optional ByVal CargaHoraria As Integer, _
    Optional ByVal Subsidio As String)
    
    With gfpcFPConversao
       .VencimentoRB = RetornaVencimento(.CargoDataAfastamento, _
        DateAdd("d", -1, .DataAposentadoria), True, _
            CargaHoraria, Subsidio).Valor
    End With

End Function
Public Function FPConversaoBuscaPosicionamentoDataAfastamento()
    With gfpcFPConversao
       .CargoDataAfastamento = ServidorBuscaPosicionamento(DateAdd("d", -1, .DataAposentadoria))
    End With
End Function





