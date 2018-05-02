Attribute VB_Name = "modPesquisaDadosFinanceiros"
Public Function DadosFinanceirosMesesAnteriores(Optional ByVal _
    Data As Date)
    
    NavDadosFinanceirosMesesAnteriores Data
  
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



