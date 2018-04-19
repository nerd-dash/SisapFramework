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
    Titulo = "PESQUISAR DADOS PESSOAIS"
    With gsspSisap
        If Not .PegaCampo(Len(Titulo), 4, 26) = Titulo Then
            .AcessaComando "PESQUISA DADOS.PESSOAIS"
        End If
    End With
End Sub

Public Sub PesquisaDeDadosFinanceiros()
    Titulo = "PESQUISA DADOS FINANCEIROS"
    With gsspSisap
        If Not .PegaCampo(Len(Titulo), 4, 29) = Titulo Then
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
    Titulo = "DESATIVAR ASSIST.MEDICA -IPSEMG"
    With gsspSisap
        If Not .VerificaTituloTela( _
            "MANUTENCAO ASSISTENCIA MEDICA IPSEMG") Then
            .AcessaComando Titulo
        End If
    End With
End Sub

Public Sub PesquisaTabelasSisap()
    Titulo = "PESQUISAR TABELAS"
    With gsspSisap
        If Not .VerificaTituloTela(Titulo) Then
            .AcessaComando "PESQUISA TABELAS SISAP"
        End If
    End With
End Sub





