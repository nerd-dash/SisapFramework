VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clswsFaltasConsHistorico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IRelatorio
Implements IFaltas

Private Planilha As Worksheet

Private Tabela As ListObject
Private Linha As ListRow
Private pFaltas As Collection
Private Relatorio As ListObject


Private Sub Class_Initialize()
    Set Planilha = wsFaltasConsolidadasHistorico
    Set Tabela = Planilha.ListObjects("FaltasConsolidadasHistorico")
    Set Relatorio = Planilha.ListObjects("RelatorioFaltasConsolidadasHistorico")
    IFaltas_PegaFaltasDaTabela
End Sub

Public Function IFaltas_Add(Item As IFalta)
    
    Planilha.Unprotect
    
    pFaltas.Add Item
    
    Set Linha = Tabela.ListRows.Add
    With Linha
 
        .Range(1) = Item.MaspDv
        .Range(2) = Item.Adm
        .Range(3) = Item.Apuracao
        .Range(4) = Item.Tipo
        .Range(5) = Item.Quantidade
        .Range(6) = Item.NaturezaQuantidade
        .Range(7) = Item.Complementar
        .Range(8) = Item.NaturezaComplementar
        .Range(9) = Item.Reposicao
        
    End With
    
    Tabela.DataBodyRange.Locked = False

   
    Planilha.Protect
    
End Function
Public Function IFaltas_Remove(NomeOuIndex As String)
    pFaltas.Remove NomeOuIndex
End Function

Public Function IFaltas_Limpar()
    IRelatorio_Limpar
End Function

Public Function IRelatorio_Limpar()
    Planilha.Unprotect
    
    With Tabela
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.ClearContents
        End If
    End With
    
    With Relatorio
        If Not .DataBodyRange Is Nothing Then
            .DataBodyRange.Delete
        End If
    End With
    
    Tabela.DataBodyRange.Locked = False
    
    Planilha.Protect
    
End Function

Public Property Get IFaltas_Count() As Long
    IRelatorio_Count = pFaltas.Count
End Property

Public Property Get IFaltas_Item(ByVal NomeOuIndex As Integer) As IFalta
    Set IRelatorio_Item = pFaltas.Item(NomeOuIndex)
End Property

Public Property Get IFaltas_Faltas() As Collection
    IFaltas_PegaFaltasDaTabela
    Set IFaltas_Faltas = pFaltas
End Property


Public Function IFaltas_Clear()
    Set pFaltas = Nothing
    Set pFaltas = New Collection
End Function
Public Function IRelatorio_AtualizaFormulario()
    
    Planilha.Unprotect
    
    Dim ultimaLinhaTabela As Long

    ultimaLinhaTabela = Tabela.Range.row + _
        IIf(Tabela.ListRows.Count = 0, 1, Tabela.ListRows.Count)

    Planilha.Activate
    gestEstilo.Fundo Planilha.Range("A1:N1175")
    gestEstilo.FormularioFundo Planilha.Range(Cells(2, 2), _
        Cells(ultimaLinhaTabela + 1, 12))
    gestEstilo.FormularioTitulo Planilha.Range("C3:k3")
    gestEstilo.FormularioGrupo Planilha.Range("C5:k5")

    Planilha.Protect
    
End Function

Private Function IFaltas_PegaFaltasDaTabela()

    Dim Falta As IFalta
    Dim Linha As ListRow
    
    IFaltas_Clear
    
    'popula a colection com dados da tabela
    For Each Linha In Tabela.ListRows
        
        'Por enquanto toda falta vai ser magistério para manter simples
        Set Falta = New clsFaltaMagisterio
        
        Falta.MaspDv = Linha.Range(1)
        Falta.Adm = Linha.Range(2)
        Falta.Apuracao = Linha.Range(3)
        Falta.Tipo = Linha.Range(4)
        Falta.Quantidade = Linha.Range(5)
        Falta.NaturezaQuantidade = Linha.Range(6)
        Falta.Complementar = Linha.Range(7)
        Falta.NaturezaComplementar = Linha.Range(8)
        Falta.Reposicao = Linha.Range(9)
        
        pFaltas.Add Falta
    Next Linha
    
End Function

Public Function IRelatorio_Imprimir()

    ActiveSheet.PageSetup.PrintArea = Relatorio.Range.Address
    EventosHabilitados False
    
    Dim strFilename     As String
    Dim strSaveToDirectory   As String
    
    strFilename = gdsvServidor.MaspDv & gdsvServidor.Admisao _
            & " " & gdsvServidor.Nome
    strFilename = Replace(strFilename, " ", "_")
    
    strFilename = ActiveSheet.Name & "-" & _
    "-" & Format(Now(), "yyyy-mm-dd-hh-nn") & ".pdf"
    
    ActiveSheet.Range("Area_de_impressao").ExportAsFixedFormat _
            Type:=xlTypePDF, _
            Filename:=strFilename, _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=True, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=True
    Debug.Print "Salvando : " & strFilename
    EventosHabilitados True
    
End Function

Public Function IRelatorio_Inserir(Dados() As Variant)
    
    Dim i As Long
    Dim Dado As Variant
    
    Planilha.Unprotect
    i = 0
    Set Linha = Relatorio.ListRows.Add
    With Linha
        
        For Each Dado In Dados
            .Range(i) = Dado
            i = i + 1
        Next Dado
        
    End With
    
    Tabela.DataBodyRange.Locked = False
    Planilha.Protect
        
End Function

Public Function IRelatorio_ApagaLinhaTabela(ByVal Linha As Long)
    Tabela.ListRows(Linha).Range.ClearContents
End Function
