Attribute VB_Name = "modNavegador"
Option Explicit

Private Caminho As clstelsTela
Private Tela As clsTela

Public Function NavDadosFuncionais() As Integer
    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(4)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,2,3,34,4
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = New clsTela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDadosServidor
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 8
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then 'TELA 3
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .ProximoCampo 8
            .EnviaAdm gdsvServidor.Admisao
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Caminho.Item(4).Equals(Tela) Then 'TELA 4
            Debug.Print "Navegou até os Dados do Servidor"
        End If
        
    End With
    
End Function

Public Function NavEntraCargoAtivo(Optional ByVal DataReferencia As Date _
) As Integer
    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(5)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,2,3,4,5,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Tela.Indice < 4 Then 'TELA 1
            NavDadosFuncionais
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(4).Equals(Tela) Then 'TELA 4
            IdentificarCargo 14, 18, 24, 7, DataReferencia
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(5).Equals(Tela) Then 'Fixing
            Debug.Print "Navegou até o Cargo Ativo"
        End If
        
    End With
    
End Function

Public Function NavIncluirDesligamentoDesignado()

 With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(17)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,10,15,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            Designacoes
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 4
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 15 Then 'Inserir masp e adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .EnviaAdm gdsvServidor.Admisao
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 16 Then 'Identifica Cargo
            IdentificarCargo 12, 21, 24, 8, DATA_EM_ABERTO, 52, 65
            GoTo EncontaPosicaoNoCaminho
         ElseIf Caminho.Item(3).Equals(Tela) Then 'TELA 4
            Debug.Print "Navegou Inclusão de Desligamento de Designado"
        End If
        
    End With
    
End Function

Public Function NavLancamentoCargoCodigoRecebimento()

    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(40)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,38,39,40,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            ManutencaoDeDadosFinanceiros
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 1
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Caminho.Item(3).Equals(Tela) Then 'Envia masp e adm
            .EnviaMaspDv gdsvServidor.MaspDv
            .EnviaAdm gdsvServidor.Admisao
            .Enter 2
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(4).Equals(Tela) Then
            Debug.Print "Navegou Inclusão de Verbas"
        End If
        
    End With

End Function


Public Function NavPesquisarFeriasPremio(ByVal MaspDv As Long, _
    ByVal Admissao As Integer, _
    Optional ByVal DataReferencia As Date)
    
    If DataReferencia = DATA_VAZIA Then
        DataReferencia = Date
    End If

    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(46)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,2,46,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDadosServidor
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 12
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Tela.Indice = 45 Then 'Inserir masp e adm
            .EnviaMaspDv MaspDv
            .EnviaAdm Admissao
            .Enter 1
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 16 Then 'Identifica Cargo
            IdentificarCargo 12, 21, 24, 8, DataReferencia, 52, 65
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then
            Debug.Print "Navegou até Pesquisa Férias Prêmio"
        End If
        
    End With

End Function
