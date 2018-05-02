Attribute VB_Name = "modSisap"

Public Sub PesquisaDadosServidor()
    With gsspSisap
        If Not gsspSisap.Tela.Indice = 2 Then
            .AcessaComando "PESQUISA DADOS.SERVIDOR SISAP"
            .PrimeiroCampo
        End If
    End With
End Sub

Public Sub Designacoes()
    With gsspSisap
         If Not gsspSisap.Tela.Indice = 10 Then
            .AcessaComando "DESIGNACAO"
            .PrimeiroCampo
        End If
    End With
End Sub

Public Sub ManutencaoDeDadosFinanceiros()
    
    With gsspSisap
        If Not gsspSisap.Tela.Indice = 38 Then
            .AcessaComando "MANUTENCAO DADOS FINANCEIROS"
        End If
    End With
End Sub

Public Sub PesquisaDadosPessoais()
    
    With gsspSisap
        If Not gsspSisap.Tela.Indice = 27 Then
            .AcessaComando "PESQUISA DADOS.PESSOAIS"
        End If
    End With
End Sub

Public Sub PesquisaDeDadosFinanceiros()
    Titulo = "PESQUISA DADOS FINANCEIROS"
    With gsspSisap
        If Not gsspSisap.Tela.Indice = 35 Then
            .AcessaComando Titulo, "* DADOS"
        End If
    End With
End Sub

Public Sub CargaHorariaSEE()
    Titulo = "CARGA HORARIA - SEE"
    With gsspSisap
        If Not .PegaCampo(Len(Titulo), 4, 31) = Titulo Then
            .AcessaComando "CARGA HORARIA SEE", "CARGA *"
        End If
    End With
End Sub

Public Sub DesativarAssitMedicaIPSEMG()
    With gsspSisap
        If Not gsspSisap.Tela.Indice = 42 Then
            .AcessaComando "DESATIVAR ASSIST.MEDICA -IPSEMG"
        End If
    End With
End Sub

Public Sub PesquisaTabelasSisap()
    With gsspSisap
        If Not gsspSisap.Tela.Indice = 98 Then
            .AcessaComando "PESQUISA TABELAS SISAP"
        End If
    End With
End Sub

Public Sub PesquisarVinculados()

     With gsspSisap
         If Not gsspSisap.Tela.Indice = 75 Then
            .AcessaComando "PESQUISA VINCULADOS"
            .PrimeiroCampo
        End If
    End With

End Sub

Public Sub MudancaSituacaoServidor()
    With gsspSisap
         If Not gsspSisap.Tela.Indice = 81 Then
            .AcessaComando "MUDANCA SITUACAO EXERCICIO."
            .PrimeiroCampo
        End If
    End With
End Sub
Public Sub PesquisaInspecaoMedica()
    With gsspSisap
         If Not gsspSisap.Tela.Indice = 81 Then
            .AcessaComando "PESQUISA INSPECAO.MEDICA"
            .PrimeiroCampo
        End If
    End With
End Sub



