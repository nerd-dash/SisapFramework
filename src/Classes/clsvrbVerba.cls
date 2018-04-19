VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsvrbVerba"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private pVerbas As Collection


Private Sub Class_Initialize()
    Set pVerbas = New Collection
End Sub
Public Sub Add(Item As clsVerba)
    pVerbas.Add Item
End Sub


Public Sub Remove(NomeOuIndex As String)
    pVerbas.Remove NomeOuIndex
End Sub

Public Property Get Count() As Long
    Count = pVerbas.Count
End Property

Public Property Get Item(ByVal NomeOuIndex As Integer) As clsVerba

    Set Item = pVerbas.Item(NomeOuIndex)
End Property

