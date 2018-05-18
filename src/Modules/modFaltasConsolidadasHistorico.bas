Attribute VB_Name = "modFaltasConsolidadasHistorico"
Option Explicit

Private Faltas As IFaltas
Private Relatorio As IRelatorio

    
Sub CadastraFaltasConsolidadasNoSistema()
    Set Faltas = New clswsFaltasConsHistorico
    Set Relatorio = Faltas
    
    Dim Falta As IFalta
    Dim Resposta As Integer
    Dim Dados(10) As Variant
    
    Dim tmpMaspDv As Long
    Dim tmpAdm As Integer
    Dim MesAnterior As Date
    
    Dim ContadorLinhas As Long
    
    'Fazer dupla verificação se quer enviar
    MsgBox "Verifique os dados de Faltas que serão inseridos!", vbExclamation, _
        "Cadastro de Faltas Consolidadas"
    
    Resposta = MsgBox("Você tem certeza que deseja cadastrar as Faltas?" & vbNewLine & _
                                    "Não será possível impedir essa ação!", _
                                     vbYesNo + vbQuestion, _
                                     "Cadastro de Faltas Consolidadas")
    EventosHabilitados False
    
    tmpMaspDv = gdsvServidor.MaspDv
    tmpAdm = gdsvServidor.Admisao
    
    Set gdsvServidor = New clswsDadosServidor
                                     
    If Resposta = vbYes Then
    
        ContadorLinhas = 0
        
        With gsspSisap
            For Each Falta In Faltas.Faltas
            
                ContadorLinhas = ContadorLinhas + 1
                
                If Not Falta.MaspDv = 0 Then
                
                    gdsvServidor.MaspDv = Falta.MaspDv
                    gdsvServidor.Admisao = Falta.Adm
                    
                    NavIncluirHistoricoFaltasConsolidadas Falta.Apuracao
                    
                    .PrimeiroCampo
                    .Envia Format(Falta.Apuracao, "mmyyyy")
                    .EnviaOpcao Falta.Tipo
                    
                    Select Case Falta.Tipo
                    
                    Case 4, 24
                        .Envia Format(Falta.Reposicao, "mmyyyy")
                    End Select
                    
                    .Enter
                    .PrimeiroCampo
                    
                    If Not Falta.Quantidade = 0 Then
                        'Not Falta.NaturezaQuantidade = 0 Then
                        .EnviaNumero Falta.Quantidade, 3
                        

                    Else
                        .ProximoCampo
                    End If
                    
                    If Not Falta.Complementar = 0 Then
                        'Not Falta.NaturezaComplementar = 0 Then
                        .EnviaNumero Falta.Complementar, 3

                    End If
                    
                    Dados(1) = .PegaCampo(50, 7, 8)
                    Dados(2) = Falta.MaspDv
                    Dados(3) = Falta.Adm
                    Dados(4) = Falta.Apuracao
                    Dados(5) = Falta.Tipo
                    Dados(6) = Falta.Quantidade
                    Dados(7) = Falta.NaturezaQuantidade
                    Dados(8) = Falta.Complementar
                    Dados(9) = Falta.NaturezaComplementar
                    Dados(10) = Falta.Reposicao
                    
                    .Enter 1, 8
                    .F5
                    Relatorio.Inserir Dados
                    Relatorio.ApagaLinhaTabela ContadorLinhas
                End If
            Next Falta
   
        End With
        
        gdsvServidor.MaspDv = tmpMaspDv
        gdsvServidor.Admisao = tmpAdm
        
        ImprimeRelatorioDeFaltasConsolidadas
               
        MsgBox "Todas as faltas foram cadastradas com sucesso!", vbInformation, _
            "Cadastro de Faltas Consolidadas"
        
        Faltas.Limpar
        
    End If
    
    EventosHabilitados True
    
   
End Sub

Sub ImprimeRelatorioDeFaltasConsolidadas()
    Set Relatorio = New clswsFaltasConsHistorico
    Relatorio.Imprimir
End Sub


