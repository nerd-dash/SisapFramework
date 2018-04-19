VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clschsCargaHoraria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pCargasHoraria As Collection

Private Sub Class_Initialize()
    Set pCargasHoraria = New Collection
End Sub
Public Sub Add(Item As clsCargaHoraria)
    pCargasHoraria.Add Item
End Sub
Public Sub Remove(NomeOuIndex As String)
    pCargasHoraria.Remove NomeOuIndex
End Sub

Public Property Get Count() As Long
    Count = pCargasHoraria.Count
End Property

Public Property Get Item(ByVal NomeOuIndex As Integer) As clsCargaHoraria
    Set Item = pCargasHoraria.Item(NomeOuIndex)
End Property


Property Get TotalRB(Optional ByVal DataReferencia As Date) As Integer

        If DataReferencia = DATA_VAZIA Then
            DataReferencia = DATA_EM_ABERTO
        End If
        
        Dim CargaHoraria As clsCargaHoraria
        
        Dim TiposRB() As Integer
    
        TotalRB = 0
    
        TiposRB = intArray(1, 7, 8, 10, 11, 12, 13, 14, 17, 20, 21, 33, 5, 30, 31)
        
        For Each CargaHoraria In pCargasHoraria
        
            If DataEstaEntre(DataReferencia, _
                CargaHoraria.DataInicio, _
                CargaHoraria.DataFimEfetivo) Then
                
                 For Each cod In TiposRB
                 
                     If cod = CargaHoraria.Tipo And _
                        CargaHoraria.CodGrupo <> 36 Then 'Ignora apostilas
                        TotalRB = TotalRB + CargaHoraria.QuantidadeAulas
                        Exit For
                     End If
                 
                 Next cod
            End If
            
        Next CargaHoraria
    
End Property


Property Get TotalEC(Optional ByVal DataReferencia As Date) As Integer

        TotalEC = 0

        If DataReferencia = DATA_VAZIA Then
            DataReferencia = DATA_EM_ABERTO
        End If
        
        Dim CargaHoraria As clsCargaHoraria
         
        For Each CargaHoraria In pCargasHoraria
        
            If DataEstaEntre(DataReferencia, _
                CargaHoraria.DataInicio, _
                CargaHoraria.DataFimEfetivo) Then
                
                If CargaHoraria.Tipo = 2 Then
                    TotalEC = TotalEC + CargaHoraria.QuantidadeAulas
                End If
                
            End If
            
        Next CargaHoraria
End Property


Property Get TotalEX(Optional ByVal DataReferencia As Date) As Integer
        
        TotalEX = 0
        
        If DataReferencia = DATA_VAZIA Then
            DataReferencia = DATA_EM_ABERTO
        End If
        
        Dim CargaHoraria As clsCargaHoraria
         
        For Each CargaHoraria In pCargasHoraria
        
            If DataEstaEntre(DataReferencia, _
                CargaHoraria.DataInicio, _
                CargaHoraria.DataFimEfetivo) Then
                
                If CargaHoraria.Tipo = 22 Then
                    TotalEX = TotalEX + CargaHoraria.QuantidadeAulas
                End If
                
            End If
            
        Next CargaHoraria
End Property


Property Get TotalECEX(Optional ByVal DataReferencia As Date) As Integer

        TotalECEX = 0
        
        If DataReferencia = DATA_VAZIA Then
            DataReferencia = DATA_EM_ABERTO
        End If
        
        Dim CargaHoraria As clsCargaHoraria
         
        For Each CargaHoraria In pCargasHoraria
        
            If DataEstaEntre(DataReferencia, _
                CargaHoraria.DataInicio, _
                CargaHoraria.DataFimEfetivo) Then
                
                If CargaHoraria.Tipo = 23 Then
                    TotalECEX = TotalECEX + CargaHoraria.QuantidadeAulas
                End If
                
            End If
            
        Next CargaHoraria
End Property




