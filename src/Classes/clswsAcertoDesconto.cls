VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clswsAcertoDesconto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Implements IVerbas

Private Planilha As Worksheet

Private Tabela As ListObject
Private Linha As ListRow
Private pVerbas As Collection


Private Sub Class_Initialize()
    Set Planilha = wsAcertoDescontos
    Set Tabela = Planilha.ListObjects("Acerto.Descontos")
    Set pVerbas = New Collection
End Sub

Public Function IVerbas_Add(Item As clsVerba)
    
    Planilha.Unprotect
    
    pVerbas.Add Item
    
    Set Linha = Tabela.ListRows.Add
    With Linha
 
        .Range(1) = Item.Operacao
        .Range(2) = Item.Verba
        .Range(3) = IIf(Item.DataInicio = DATA_EM_ABERTO, "", Item.DataInicio)
        .Range(4) = IIf(Item.DataFim = DATA_EM_ABERTO, "", Item.DataFim)
        .Range(5) = IIf(Item.QtdEspecif = 0, "", Item.QtdEspecif)
        .Range(6) = IIf(Item.Valor = 0, "", Item.Valor)
        .Range(7) = IIf(Item.Vigencia = DATA_EM_ABERTO, "", Item.Vigencia)
        
    End With
    
    Tabela.DataBodyRange.Locked = False
    
    IVerbas_AtualizaFormulario
   
    Planilha.Protect
    
End Function
Public Function IVerbas_Remove(NomeOuIndex As String)
    pVerbas.Remove NomeOuIndex
End Function

Public Function IVerbas_Limpa()
    Planilha.Unprotect
    
    Do While Not Tabela.ListRows.Count = 0
        Tabela.ListRows(Tabela.ListRows.Count).Delete
    Loop
    
    IVerbas_AtualizaFormulario
   
    Planilha.Protect
    
End Function

Public Property Get IVerbas_Count() As Long
    IVerbas_Count = pVerbas.Count
End Property

Public Property Get IVerbas_Item(ByVal NomeOuIndex As Integer) As clsVerba
    Set Item = pVerbas.Item(NomeOuIndex)
End Property

Public Property Get IVerbas_Verbas() As Collection
    
    Dim Verba As clsVerba
    Dim Linha As ListRow
    
    IVerbas_Clear
    
    'popula a colection com dados da tabela
    For Each Linha In Tabela.ListRows
    
        Set Verba = New clsVerba
        
        Verba.Operacao = Linha.Range(1)
        Verba.Verba = Linha.Range(2)
        Verba.DataInicio = Linha.Range(3)
        Verba.DataFim = Linha.Range(4)
        Verba.QtdEspecif = Linha.Range(5)
        Verba.Valor = Linha.Range(6)
        Verba.Vigencia = Linha.Range(7)
        
        pVerbas.Add Verba
    Next Linha
    
    Set IVerbas_Verbas = pVerbas
End Property


Public Function IVerbas_Clear()
    Set pVerbas = Nothing
    Set pVerbas = New Collection
End Function
Public Function IVerbas_AtualizaFormulario()
    
    Dim ultimaLinhaTabela As Long
    
    Planilha.Unprotect
    
    ultimaLinhaTabela = Tabela.Range.row + _
        IIf(Tabela.ListRows.Count = 0, 1, Tabela.ListRows.Count)

    Planilha.Activate
    gestEstilo.Fundo Planilha.Range("A1:N175")
    gestEstilo.FormularioFundo Planilha.Range(Cells(2, 2), _
        Cells(ultimaLinhaTabela + 1, 13))
    gestEstilo.FormularioTitulo Planilha.Range("C3:L3")

    Planilha.Protect

End Function

Private Function Corrige5647()

    Dim i As Long
    Dim Verba As clsVerba
    
    Set Verba = New clsVerba
    
    
    i = 9
    
    gsspSisap.PrimeiroCampo
    
    Do
        Verba.PreencheVerba i
        
        If Not Verba.Verba = 0 Then
            If Verba.Verba = 5647 And _
                Verba.QtdEspecif <> 1 Then
                    gsspSisap.Alterar
                    gsspSisap.ProximoCampo 7
                    gsspSisap.EnviaNumero 1
                    gsspSisap.Enter 1, 8
                    Exit Do
            End If
            gsspSisap.ProximaLinha
            i = i + 1
        End If
    Loop While Not Verba.Verba = 0




End Function

