VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IFaltas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private pFaltas As Collection

Public Function Add(Item As IFalta)
End Function

Public Function Remove(NomeOuIndex As String)
End Function

Public Property Get Count() As Long
End Property

Public Property Get Item(ByVal NomeOuIndex As Integer) As IFalta
End Property

Public Function Limpar()
End Function

Public Property Get Faltas() As Collection
End Property

Public Function Clear()
End Function


