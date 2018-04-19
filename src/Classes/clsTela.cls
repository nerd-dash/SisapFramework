VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private pIndice As Long
Private pTitulo As String
Private pCodPrograma As String
Private pCodTela As String
Private pProximaTela As Long
Private pAnteriorTela As Long


Public Property Get Indice() As Long
    Indice = pIndice
End Property

Public Property Let Indice(ByVal Indice_ As Long)
    pIndice = Indice_
End Property

Public Property Get Titulo() As String
    Titulo = pTitulo
End Property

Public Property Let Titulo(ByVal Titulo_ As String)
    pTitulo = Titulo_
End Property

Public Property Get CodPrograma() As String
    CodPrograma = pCodPrograma
End Property

Public Property Let CodPrograma(ByVal CodPrograma_ As String)
    pCodPrograma = CodPrograma_
End Property

Public Property Get CodTela() As String
    CodTela = pCodTela
End Property

Public Property Let CodTela(ByVal CodTela_ As String)
    pCodTela = CodTela_
End Property

Public Property Get ProximaTela() As Long
    ProximaTela = pProximaTela
End Property

Public Property Let ProximaTela(ByVal ProximaTela_ As Long)
    pProximaTela = ProximaTela_
End Property

Public Property Get AnteriorTela() As Long
    AnteriorTela = pAnteriorTela
End Property

Public Property Let AnteriorTela(ByVal AnteriorTela_ As Long)
    pAnteriorTela = AnteriorTela_
End Property

Public Function Equals(ByRef Tela As clsTela) As Boolean
    
    If Tela.CodPrograma = pCodPrograma And _
    Tela.CodTela = pCodTela And _
    Tela.Titulo = pTitulo Then
        pAnteriorTela = Tela.AnteriorTela
        pProximaTela = Tela.ProximaTela
        pIndice = Tela.Indice
        Equals = True
    Else
        Equals = False
    End If
    
End Function

Private Sub Class_Initialize()
    
    pCodTela = gsspSisap.PegaCampo(8, 2, 3)
    pCodPrograma = gsspSisap.PegaCampo(8, 3, 3)
    pTitulo = gsspSisap.PegaTituloTela

End Sub
