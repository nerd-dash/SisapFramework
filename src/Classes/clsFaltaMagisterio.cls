VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFaltaMagisterio"
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
Private pRB As Integer
Private pExEC As Integer
Private pMaspDv As Long
Private pAdm As Integer
Private pNaturezaRB As Long
Private pNaturezaExEC As Long
Private pCarreira As String
Private pCadastradasRB As Integer
Private pCadastradasExEC As Integer
Private pMinistradasRB As Integer
Private pMinistradasExEC As Integer


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

Public Property Get IFalta_Apuracao() As Date
    IFalta_Apuracao = pApuracao
End Property

Public Property Let IFalta_Apuracao(ByVal Data As Date)
    pApuracao = Data
End Property

Public Property Get IFalta_Reposicao() As Date
    IFalta_Reposicao = pReposicao
End Property

Public Property Let IFalta_Reposicao(ByVal Data As Date)
    pReposicao = Data
End Property

Public Property Get IFalta_Tipo() As Integer
    IFalta_Tipo = pTipo
End Property
Public Property Let IFalta_Tipo(ByVal Tipo As Integer)
    pTipo = Tipo
End Property

Public Property Get IFalta_Dias() As Integer
    IFalta_Dias = pDias
End Property
Public Property Let IFalta_Dias(ByVal Dias As Integer)
    pDias = Dias
End Property

Public Property Get IFalta_Quantidade() As Integer
    IFalta_Quantidade = pRB
End Property
Public Property Let IFalta_Quantidade(ByVal RB_em_Horas As Integer)
    Dim Limiar As Double
    Dim ParteFracionaria As Double
    Dim ParteInteira As Integer
        
On Error GoTo LogaErro
        
    pRB = RB_em_Horas
    
    If RB_em_Horas = 1 Then
        pDias = 0
    Else
    
        Limiar = RB_em_Horas / 3.6
        
        ParteInteira = Int(Limiar)
        
        ParteFracionaria = Limiar - ParteInteira
        
        If ParteFracionaria >= 0.2 Then
            pDias = ParteInteira + 1
        Else
            pDias = ParteInteira
        End If
        
    End If
Exit Property

LogaErro:

    Dim Erro As IErro
    Set Erro = New clsErro
    
    Erro.LogaErro Err, "Let IFalta_Quantidade"
    
    
End Property

Public Property Get IFalta_Complementar() As Integer
    IFalta_Complementar = pExEC
End Property
Public Property Let IFalta_Complementar(ByVal Horas_Aula_de_Ex_ou_EC As Integer)
    pExEC = Horas_Aula_de_Ex_ou_EC
End Property

Public Property Get IFalta_NaturezaQuantidade() As Long
    IFalta_NaturezaQuantidade = pNaturezaRB
End Property
Public Property Let IFalta_NaturezaQuantidade(ByVal NaturezaRB As Long)
    pNaturezaRB = NaturezaRB
End Property

Public Property Get IFalta_NaturezaComplementar() As Long
    IFalta_NaturezaComplementar = pNaturezaExEC
End Property
Public Property Let IFalta_NaturezaComplementar(ByVal NaturezaExEC As Long)
    pNaturezaExEC = NaturezaExEC
End Property

Public Property Get IFalta_NaturezaCarreira() As Long
    IFalta_NaturezaCarreira = pCarreira
End Property
Public Property Let IFalta_NaturezaCarreira(ByVal Carreira As Long)
    pCarreira = Carreira
End Property

Public Property Get IFalta_Carreira() As String
    IFalta_Carreira = pCarreira
End Property
Public Property Let IFalta_Carreira(ByVal Carreira_ As String)
    pCarreira = Carreira_
End Property

Public Property Get CadastradasRB() As Integer
    CadastradasRB = pCadastradasRB
End Property
Public Property Let CadastradasRB(ByVal Aulas As Integer)
    pCadastradasRB = Aulas
End Property

Public Property Get CadastradasExEC() As Long
    CadastradasExEC = pCadastradasExEC
End Property
Public Property Let CadastradasExEC(ByVal Aulas As Long)
    pCadastradasExEC = Aulas
End Property

Public Property Get MinistradasRB() As Long
    MinistradasRB = pMinistradasRB
End Property
Public Property Let MinistradasRB(ByVal Aulas As Long)
    pMinistradasRB = Aulas
End Property

Public Property Get MinistradasExEC() As Long
    MinistradasExEC = pMinistradasExEC
End Property
Public Property Let MinistradasExEC(ByVal Aulas As Long)
    pMinistradasExEC = Aulas
End Property
