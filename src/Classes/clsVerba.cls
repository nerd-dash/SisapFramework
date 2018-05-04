VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVerba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pOperacao As String
Private pVerba As Long
Private pDataInicio As Date
Private pDataFim As Date
Private pValor As Double
Private pQtdEspecif As Double
Private pVigencia As Date

Public Property Get Operacao() As String

    Operacao = pOperacao

End Property


Public Property Get Verba() As Long

    Verba = pVerba

End Property

Public Property Get DataInicio() As Date

    DataInicio = pDataInicio
    
End Property


Public Property Get DataFim() As Date

    DataFim = pDataFim
    
End Property

Public Property Get Vigencia() As Date

    Vigencia = pVigencia
    
End Property

Public Property Get Valor() As Double

    Valor = pValor

End Property

Public Property Get QtdEspecif() As Double

    QtdEspecif = pQtdEspecif

End Property

Public Property Let Operacao(Operacao_ As String)

    pOperacao = Operacao_

End Property


Public Property Let Verba(Verba_ As Long)

    pVerba = Verba_

End Property

Public Property Let DataInicio(DataInicio_ As Date)

    pDataInicio = DataInicio_
    
End Property


Public Property Let DataFim(DataFim_ As Date)

    pDataFim = DataFim_
    
End Property

Public Property Let Vigencia(Vigencia_ As Date)

    pVigencia = Vigencia_
    
End Property

Public Property Let Valor(Valor_ As Double)

    pValor = Valor_

End Property

Public Property Let QtdEspecif(QtdEspecif_ As Double)

    pQtdEspecif = QtdEspecif_

End Property

Public Function PreencheVerba(ByVal Linha As Long)
        
        Dim Verba As Long
        
        With gsspSisap
            
            Verba = .PegaVerba(Linha, 5)
            
            If Verba <> 0 Then
                pVerba = Verba
                pOperacao = .PegaCampo(1, Linha, 3)
                pDataInicio = .PegaData(14, Linha, 11)
                pDataFim = .PegaData(14, Linha, 25)
                pQtdEspecif = .PegaCampoMoeda(11, True, Linha, 40)
                pValor = .PegaCampoMoeda(10, True, Linha, 52)
                pVigencia = .PegaData(14, Linha, 63)
            End If
        End With
End Function
