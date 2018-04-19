VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsAcertoVantagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Planilha As wsAcertoDesignacao
Private pVerbas As New clsvrbVerba
Private pLinha As Long
Dim Tabela As ListObject

Private Sub ClassInitialize()
    Set Tabela = Planilha.ListObjects("Acerto.Vantagens")
    pLinha = pVerbas.Count + 1
    
End Sub

Public Function Add(Item As clsVerba)
    
    
    pVerbas.Add Item
 
    Tabela.DataBodyRange(Linha, 1) = Item.Operacao
    Tabela.DataBodyRange(Linha, 2) = Item.verba
    Tabela.DataBodyRange(Linha, 3) = IIf(Item.DataInicio = DATA_EM_ABERTO, "", Item.DataInicio)
    Tabela.DataBodyRange(Linha, 4) = IIf(Item.DataFim = DATA_EM_ABERTO, "", Item.DataFim)
    Tabela.DataBodyRange(Linha, 5) = IIf(Item.QtdEspecif = 0, "", Item.QtdEspecif)
    Tabela.DataBodyRange(Linha, 6) = IIf(Item.Valor = 0, "", Item.Valor)
    Tabela.DataBodyRange(Linha, 7) = IIf(Item.Vigencia = DATA_EM_ABERTO, "", Item.Vigencia)
    
    Linha = Linha + 1
    
End Function
Public Function Remove(NomeOuIndex As String)
    pVerbas.Remove NomeOuIndex
End Function

Public Property Get Count() As Long
    Count = pVerbas.Count
End Property

Public Property Get Item(ByVal NomeOuIndex As Integer) As clsVerba
    Set Item = pVerbas.Item(NomeOuIndex)
End Property

