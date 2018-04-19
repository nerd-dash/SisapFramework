VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clstelsTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pTela As Collection


Private Sub Class_Initialize()
    Set pTela = New Collection
End Sub
Public Sub Add(Item As clsTela)
    pTela.Add Item
End Sub
Public Sub Remove(NomeOuIndex As String)
    pTela.Remove NomeOuIndex
End Sub

Public Property Get Count() As Long
    Count = pTela.Count
End Property

Public Property Get Item(ByVal NomeOuIndex As Integer) As clsTela
    Set Item = pTela.Item(NomeOuIndex)
End Property

Public Property Get Collection() As Collection
    Set Collection = pTela
End Property

Public Property Let Collection(Telas As Collection)
    Set pTela = Telas
End Property


