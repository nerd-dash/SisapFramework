Attribute VB_Name = "modVerbas"
Public Function PegaVerbasCargoRecebimento( _
    Optional ByVal LinhaInicial As Integer = 9, _
    Optional ByVal LinhaFinal As Integer = 20)
    
    NavPesquisaDadosFinanceirosCargoRecebimento
    Dim Vantagens As New clswsAcertoVantagem
    Dim Desconto As New clswsAcertoDesconto
    
    PegaVerbas Vantagens
    gsspSisap.Enter
    PegaVerbas Desconto
    
End Function

Private Function PegaVerbas(ByRef acerto As IVerbas, _
 Optional ByVal LinhaInicial As Integer = 9, _
    Optional ByVal LinhaFinal As Integer = 20)

    Dim Verba As clsVerba
     
    With gsspSisap
    
        Dim AchouPosicionamento As Boolean
        
        acerto.Limpa
        
        Do
            For i = LinhaInicial To LinhaFinal
        
                Set Verba = New clsVerba
                
                Verba.Verba = .PegaVerba(i, 5)
                
                If Verba.Verba <> 0 Then
                
                    Verba.Operacao = .PegaCampo(1, i, 3)
                    Verba.DataInicio = .PegaData(14, i, 11)
                    Verba.DataFim = .PegaData(14, i, 25)
                    Verba.QtdEspecif = .PegaCampoMoeda(11, True, i, 40)
                    Verba.Valor = .PegaCampoMoeda(10, True, i, 52)
                    Verba.Vigencia = .PegaData(14, i, 63)
                    
                    acerto.Add Verba
                Else
                    AchouPosicionamento = True
                End If
            Next i
            
           If Not AchouPosicionamento Then
                If .F8(1, 9) = 9 Then
                    AchouPosicionamento = True
                End If
            End If
        
        Loop While AchouPosicionamento = False
        

    
    End With


End Function
Public Function CalculaVerbaIPSEMG(ByVal Valor As Double) As Double
    
    Dim dbl As Double
    
    dbl = (Valor * 3.2) / 100
    
    If dbl <= 30 Then
        dbl = 30
    ElseIf dbl >= 250 Then
        dbl = 250
    End If
    
    CalculaVerbaIPSEMG = dbl
    
End Function
