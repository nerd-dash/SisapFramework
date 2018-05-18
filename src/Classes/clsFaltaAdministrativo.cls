VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFaltaAdministrativo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IFalta

Private pApuracao As Date
Private pReposicao As Date
Private pTipo As Integer
Private pDias As Integer
Private pHorasDescontar As Integer

Public Property Get IFalta_Apuracao() As Date
End Property

Public Property Let IFalta_Apuracao(ByVal Data As Date)
End Property

Public Property Get IFalta_Reposicao() As Date
End Property

Public Property Let IFalta_Reposicao(ByVal Data As Date)
End Property

Public Property Get IFalta_Tipo() As Integer
End Property
Public Property Let IFalta_Tipo(ByVal Tipo As Integer)
End Property

Public Property Get IFalta_Dias() As Integer
End Property
Public Property Let IFalta_Dias(ByVal Dias As Integer)
End Property

Public Property Get HorasDescontar() As Integer
End Property
Public Property Let HorasDescontar(ByVal Horas As Integer)
End Property

Public Property Get IFalta_Adm() As Integer
    IFalta_Adm = pAdm
End Property
Public Property Let IFalta_Adm(ByVal Adm As Integer)
    pAdm = Adm
End Property

Public Property Get IFalta_MaspDv() As Long
    IFalta_MaspDv = pMaspDv
End Property
Public Property Let IFalta_MaspDv(ByVal MaspDv As Long)
    pMaspDv = MaspDv
End Property
