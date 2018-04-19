Attribute VB_Name = "TesteVerbas"
Sub TesteVerbas()
  
    Dim verba As New clsVerba
    Dim Vantagens As New clsAcertoVantagem
    
    With gsspSisap
    
    
    
        For i = 9 To 20
        
            Set verba = New clsVerba
            
            verba.verba = .PegaVerba(i, 5)
            
            If verba.verba <> 0 Then
            
                verba.Operacao = .PegaCampo(1, i, 3)
                verba.DataInicio = .PegaData(14, i, 11)
                verba.DataFim = .PegaData(14, i, 25)
                verba.QtdEspecif = .PegaCampoMoeda(11, True, i, 40)
                verba.Valor = .PegaCampoMoeda(10, True, i, 52)
                verba.Vigencia = .PegaData(14, i, 63)
                
                Vantagens.Add verba
            End If
        Next i
    
    End With

End Sub

