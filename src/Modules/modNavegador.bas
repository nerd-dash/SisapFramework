Attribute VB_Name = "modNavegador"
Option Explicit

Private Caminho As clstelsTela
Private Tela As clsTela

Public Function NavDadosFuncionais() As Integer
    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(4)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,2,3,4
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
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
        Else
            gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 08 - Dados Funcionais"

        End If
        
    End With
    
End Function

Public Function NavEntraCargoAtivo(Optional ByVal Data As Date) As Boolean
    With gsspSisap
    
        If Data = DATA_VAZIA Then
                Data = DATA_EM_ABERTO
        End If
        
        NavDadosFuncionais
        
        NavEntraCargoAtivo = EntraCargoAtivo(Data)
                
        Set Tela = gsspSisap.Tela
        
        If Tela.Indice = 5 Then 'TELA 3
            Debug.Print "Navegou até o Cargo Ativo"
        Else
            gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " O Cargo Ativo"
        End If
        
    End With
    
End Function
Public Function NavEvolucaoCarreira( _
    Optional ByVal DataReferencia As Date)
   
    With gsspSisap
        NavEntraCargoAtivo DataReferencia
        .MarcarOpcao 21, 19
        .Enter 1, 205
    End With

End Function

Public Function NavSituacaoExercicio()
   
    With gsspSisap
        NavDadosFuncionais
        .MarcarOpcao 20, 49
        .Enter 1, 205
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
        Else
            gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 12 - Férias Prêmio"
        End If
        
    End With

End Function

Public Function NavPesquisaDadosFinanceirosCargoRecebimento()

    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(62)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,35,61,62,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDeDadosFinanceiros
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 1
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Caminho.Item(3).Equals(Tela) Then 'Envia masp e adm
            .EnviaMaspDv gdsvServidor.MaspDv
            .EnviaAdm gdsvServidor.Admisao
            .Enter 1
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(4).Equals(Tela) Then
            Debug.Print "Navegou Pesquisa Dados Financeiro Cargo Recebimento"
        Else
            gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 01 - Dados Financeiros MêS Atual - Cargo Recebimento"
        End If
        
    End With

End Function

Public Function NavContaBancaria()

   With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(53)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,27,53,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDadosPessoais
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 3
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Tela.Indice = 28 Then 'Envia masp e adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .Enter 1
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then
            Debug.Print "Navegou Pesquisa Dados Bancários"
        Else
        gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 03 - Dados Bancarios"
            
        End If
        
    End With
End Function

Public Function NavPesquisaDadosFinanceirosCargoRecebimentoMesAnterior()

    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(64)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,35,64,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDeDadosFinanceiros
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 3
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Caminho.Item(3).Equals(Tela) Then 'Envia masp e adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .EnviaAdm gdsvServidor.Admisao
            .Enter 1
        Else
            gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 03 - Dados Financeiros Meses Anteriores"
        End If
        
    End With

End Function

Public Function NavHistoricoDePagamento()

    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(37)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,35,36,37
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDeDadosFinanceiros
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 5
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Caminho.Item(3).Equals(Tela) Then 'Envia masp e adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .Enter 2
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(4).Equals(Tela) Then
            Debug.Print "Navegou Histórico Financeiro"
        Else
            gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 05 - Histórico de Pagamento"
        End If
        
    End With

End Function

Public Function NavOcorrencias()

    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(68)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,35,36,37
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDeDadosFinanceiros
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 9
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Caminho.Item(3).Equals(Tela) Then 'Envia masp e adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .EnviaAdm gdsvServidor.Admisao
            .Enter
            Debug.Print "Navegou Ocorrências"
        Else
            gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 09 - Ocorrencias"
        End If
        
    End With

End Function

Public Function NavLiquidoBancario()

    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(70)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,35,71,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDeDadosFinanceiros
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 16
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Caminho.Item(3).Equals(Tela) Then 'Envia masp e adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .Enter
            Debug.Print "Navegou Liquido Bancario"
        Else
            gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 16 - Liquido Bancario"
        End If
        
    End With

End Function

Public Function NavPagamentoSuspensoPorMasp()

    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(73)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,35,73,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDeDadosFinanceiros
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 17
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Tela.Indice = 72 Then 'Envia masp e adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .EnviaAdm gdsvServidor.Admisao
            .Enter 1, 4
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then
            Debug.Print "Navegou Liquido Bancario"
        Else
            gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 16 - Liquido Bancario"
        End If
        
    End With

End Function

Public Function NavDadosPessoais()

   With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(29)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,27,29,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDadosPessoais
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 1
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Tela.Indice = 28 Then 'Envia masp e adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .Enter 1
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then
            Debug.Print "Navegou Pesquisa Pessoais"
        Else
        gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 01 - Dados Pessoais"
            
        End If
        
    End With
End Function

Public Function NavDocumentos()

   With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(74)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,27,29,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDadosPessoais
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 2
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Tela.Indice = 28 Then 'Envia masp e adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .Enter 1
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then
            Debug.Print "Navegou Documentos"
        Else
        gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 02 - Documentos"
            
        End If
        
    End With
End Function

Public Function NavVinculadosPorRepresentante()

   With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(76)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,75,76,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisarVinculados
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 3
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .Enter 1, 0
            Debug.Print "Navegou Vinculados Por Representante"
        Else
        gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 03 - Vinculados Por Representante"
            
        End If
        
    End With
End Function

Public Function NavEndereco()

   With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(31)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,27,31,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDadosPessoais
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 4
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Tela.Indice = 28 Then 'Envia masp e adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .Enter 1
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then
            Debug.Print "Navegou Endereço"
        Else
        gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 04 - Endereço"
            
        End If
        
    End With
End Function

Public Function NavFormacaoEscolar()

   With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(78)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,2,78,
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDadosServidor
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 28
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Tela.Indice = 77 Then 'Envia masp e adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .Enter 1
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then
            Debug.Print "Navegou Formação Escolar"
        Else
        gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 28 - Pesquisar Formação Escolar"
            
        End If
        
    End With
End Function

Public Function NavAfastamentos()

   With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(21)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,2,79,21,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDadosServidor
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 2
            .Enter 1, 2
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 1
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 80 Then 'Envia masp e adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .EnviaAdm gdsvServidor.Admisao
            .Enter 1
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(4).Equals(Tela) Then
            Debug.Print "Navegou Afastamentos"
        Else
        gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 02 - Afastamentos"
            
        End If
        
    End With
End Function

Public Function NavPublicacaoInspecaoMedica()

   With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(84)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,81,82,84

EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaInspecaoMedica
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 81 Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 3
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 3
            .PrimeiroCampo
            .EnviaOpcao 2
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 83 Then 'Envia masp e adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .Enter 1
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 34 Then 'Envia masp e adm
            .PrimeiroCampo
            ServidorSelecionarAdmissao
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then
            Debug.Print "Navegou Publicacao Inspecao Medica"
        Else
        gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 03 - Publicacao Inspecao Medica"
            
        End If

    End With
End Function

Public Function NavPesquisarCargaHorariaVigente( _
    Optional ByVal Data As Date = DATA_EM_ABERTO, _
    Optional ByVal MaspDv As Long = 0, _
    Optional ByVal Admissao As Integer = 0, _
    Optional ByVal ColunaDataInicial As Integer = 52, _
    Optional ByVal ColunaDataFinal As Integer = 65)

    If Data = DATA_EM_ABERTO Then
        Data = Date
    End If
    
    If MaspDv = 0 Then
        MaspDv = gdsvServidor.MaspDv
    End If
    
    If Admissao = 0 Then
        Admissao = gdsvServidor.Admisao
    End If
    
    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(88)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,32,33,87,
                
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            CargaHorariaSEE
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 7
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 2
            .Enter
            .Envia Format(Data, "mmyyyy")
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 86 Then 'Envia masp e adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .ProximoCampo 8
            .EnviaAdm gdsvServidor.Admisao
            .Enter 1
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 16 Then 'Envia masp e adm
            IdentificarCargo 12, 20, 24, 8, Data, _
                ColunaDataInicial, ColunaDataFinal
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(4).Equals(Tela) Then
            Debug.Print "Navegou Carga Horaria Vigente em " & Format(Data, "mm/yyyy")
        Else
        gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 07 - Pesquisar Carga Horaria"
            
        End If
        
    End With
    

End Function

Public Function NavExercicios()
        
    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(90)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,2,89,90,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDadosServidor
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 9
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then 'TELA 3
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .Enter
            GoTo EncontaPosicaoNoCaminho
         ElseIf Tela.Indice = 34 Then 'Envia masp e adm
            .PrimeiroCampo
            ServidorSelecionarAdmissao
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(4).Equals(Tela) Then 'TELA 4
            Debug.Print "Navegou Exercícios"
        Else
            gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 09 - Exercícios"

        End If
        
    End With
    
End Function

Public Function NavFaltasConsolidadas()
    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(92)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,2,91,92,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDadosServidor
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 10
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then 'TELA 3
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .EnviaAdm gdsvServidor.Admisao
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 22 Or _
            Caminho.Item(4).Equals(Tela) Then 'TELA 4
            Debug.Print "Navegou Faltas"
        Else
            gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 10 -  Faltas"

        End If
        
    End With
    
End Function

Public Function NavFeriasRegulamentares()
    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(94)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,2,91,92,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'Tela do Menu Principal
            PesquisaDadosServidor
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then
            'Menu Pesquisa dados do Serv.
            .PrimeiroCampo
            .EnviaOpcao 13
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 93 Then 'Masp e Adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .EnviaAdm gdsvServidor.Admisao
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then
            Debug.Print "Navegou Férias Regulamentares"
        Else
            gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 13 -  Férias Regulamentares"
        End If
        
    End With
    
End Function

Public Function NavDesativarAssitMedicaIPSEMG()
    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(96)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,42,96,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'Tela do Menu Principal
            DesativarAssitMedicaIPSEMG
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then
            'Menu Pesquisa dados do Serv.
            .PrimeiroCampo
            .Envia "01P"
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 95 Then 'Masp e Adm
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .EnviaAdm gdsvServidor.Admisao
            .Enter 1, 4
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then
            Debug.Print "Navegou Férias Regulamentares"
        Else
            gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 13 - Férias Regulamentares"
        End If
        
    End With
    
End Function

Public Function NavSimboloVencimento()
    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(100)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,98,99,100,
        
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'Tela do Menu Principal
            PesquisaTabelasSisap
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then
            'Menu Pesquisa dados do Serv.
            .PrimeiroCampo
            .EnviaOpcao 21
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then
            'Menu Pesquisa dados do Serv.
            .PrimeiroCampo
            .EnviaOpcao 9
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(4).Equals(Tela) Then
            Debug.Print "Navegou Simbolo Vencimento"
        Else
            gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 09 - Simbolo Vencimento"
        End If
        
    End With
    
End Function

Public Function NavPesquisaPorUnidadeSEEResumida()

   With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(102)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,2,101,102,
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDadosServidor
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 20
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 1
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(4).Equals(Tela) Then
            Debug.Print "Por Unidade (SEE) Resumida"
        Else
        gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 20 - Por Unidade (SEE) Resumida"
            
        End If
        
    End With
End Function

Public Function NavDadosFinanceirosMesesAnteriores(Optional ByVal _
    Data As Date)
    
    If Data = DATA_VAZIA Then
        Data = Date
    End If

   With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(103)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,35,64,103,
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            PesquisaDeDadosFinanceiros
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 3
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .EnviaAdm gdsvServidor.Admisao
            .Enter
            .Envia Format(Data, "mmyyyy")
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 26 Then
            IdentificarCargo 12, 21, 50, 7, Data, 57, 68, _
            "PESQUISA DADOS FINANCEIROS MESES ANTERIORES"
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(4).Equals(Tela) Then
            Debug.Print "Dados Financeiros Meses Anteriores"
        Else
        gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 03 - Dados Financeiros Meses Anteriores"
            
        End If
        
    End With
End Function

Public Function NavPesquisarMudancaSituacaoExercicio()
    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(107)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,109,105,106,107,
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            MudancaSituacaoServidor
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 1
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 4
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(4).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .EnviaAdm gdsvServidor.Admisao
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(5).Equals(Tela) Then
            Debug.Print "Pesquisar Mudanca Situacao Exercicio"
        Else
        gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 04 - Pesquisar Mudanca Situacao Exercicio"
            
        End If
        
    End With
End Function

Public Function NavPesquisarAjustamentoFuncional()
    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(114)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,109,110,114,
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            MudancaSituacaoServidor
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 3
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 4
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 111 Then 'TELA 2
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .EnviaAdm gdsvServidor.Admisao
            .Enter 1, 222
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 113 Then
            Debug.Print "Pesquisar Ajustamento Funcional - Sem ajustamento"
        ElseIf Caminho.Item(4).Equals(Tela) Then
            Debug.Print "Pesquisar Ajustamento Funcional"
        Else
        gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 04 - Pesquisar Ajustamento Funcional"
            
        End If
        
    End With
End Function

Public Function NavPesquisarFuncaoEducacao()
    With gsspSisap
        
        Set Tela = gnavNavegador.BuscaTela(117)
        
        gnavNavegador.VoltaAncestralEmComum Tela
        
        Set Caminho = gnavNavegador.CaminhoParaTela(Tela)
        '1,109,115,118,
        
EncontaPosicaoNoCaminho:
        
        Set Tela = gsspSisap.Tela
        
        If Caminho.Item(1).Equals(Tela) Then 'TELA 1
            MudancaSituacaoServidor
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(2).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 2
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(3).Equals(Tela) Then 'TELA 2
            .PrimeiroCampo
            .EnviaOpcao 5
            .Enter
            GoTo EncontaPosicaoNoCaminho
        ElseIf Tela.Indice = 116 Then 'TELA 2
            .PrimeiroCampo
            .EnviaMaspDv gdsvServidor.MaspDv
            .EnviaAdm gdsvServidor.Admisao
            .Enter 1, 222
            GoTo EncontaPosicaoNoCaminho
        ElseIf Caminho.Item(4).Equals(Tela) Then
            Debug.Print "Pesquisar Funcao Educação"
        Else
        gsspSisap.JanelaAlerta "Não foi possível navegar até : " & vbNewLine _
                & " 05 - Pesquisar Funcao Educação"
            
        End If
        
    End With
End Function


