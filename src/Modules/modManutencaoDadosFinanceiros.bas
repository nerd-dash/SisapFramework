Attribute VB_Name = "modManutencaoDadosFinanceiros"
Public Sub LancamentoCargoCodigoRecebimento()


    Titulo = "MANUTENCAO DADO FINANCEIRO"
    With gsspSisap
        If Not .VerificaTituloTela(Titulo) Then
            ManutencaoDeDadosFinanceiros
            .PrimeiroCampo
            .EnviaOpcao 1
            .Enter
        Else
            .F9
        End If
        
        .EnviaMaspDv gdsvServidor.MaspDv
        .EnviaAdm gdsvServidor.Admisao
        .Enter 2
        
    End With
    
End Sub

