VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IVerbas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Function Add(Item As clsVerba)
End Function

Public Function Remove(NomeOuIndex As String)
End Function

Public Property Get Count() As Long
End Property

Public Property Get Item(ByVal NomeOuIndex As Integer) As clsVerba
End Property

Public Function Limpa()
End Function

Public Property Get Verbas() As Collection
End Property

Public Function Clear()
End Function
Public Function AtualizaFormulario()
End Function
