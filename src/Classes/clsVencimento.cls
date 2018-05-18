VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsVencimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Cargo As clsCargo
Public Valor As Double
Public DataInicio As Date
Public DataFinal As Date
Private pSubsidio As String

Public Property Get Subsidio() As String
    Subsidio = pSubsidio
End Property

Public Property Let Subsidio(Subsidio_ As String)
    pSubsidio = Subsidio_
End Property


Public Function DefineSubisido(ByVal DataVencimento As Date) As String
    Dim Resposta As Integer
    
    If pSubsidio = "" Then
   
     Select Case DataVencimento
     
         Case Is < #1/1/2011#
             pSubsidio = "N"
         Case Is < #1/1/2012#
             Resposta = _
                 MsgBox("Em " & Format(DataVencimento, "dd/mm/yyyy") & _
                 " o servidor estava no SUBSÍDIO ?", _
                 vbYesNo + vbQuestion, _
                 "Tipo de Vencimento")
             If Resposta = vbYes Then
                 pSubsidio = "S"
             Else
                 pSubsidio = "N"
             End If
             
         Case Is < #6/1/2015#
             pSubsidio = "S"
         Case Else
             pSubsidio = "V"
     End Select
    
    Debug.Print "Vencimento do Tipo " & pSubsidio
    
   End If
   
   DefineSubisido = pSubsidio
   
End Function

