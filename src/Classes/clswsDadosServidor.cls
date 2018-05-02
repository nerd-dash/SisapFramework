VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clswsDadosServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Planilha As wsDadosServidor

Private Sub Class_Initialize()
    Set Planilha = wsDadosServidor
End Sub

Public Property Get MaspDv() As Long
    MaspDv = Planilha.[Servidor.MaspDv].Value2
End Property

Public Property Let MaspDv(MaspDv_ As Long)
    Planilha.[Servidor.MaspDv] = MaspDv_
    
End Property
Public Property Get Admisao() As Integer
    Admisao = Planilha.[Servidor.Admissao].Value2
End Property

Public Property Let Admisao(Admissao_ As Integer)
    Planilha.[Servidor.Admissao] = Admissao_
End Property

Public Property Get Nome() As String
    Nome = Planilha.[Servidor.Nome].Value2
End Property

Public Property Let Nome(ByVal Nome_ As String)
    Planilha.Unprotect
    Application.[Servidor.Nome] = Nome_
    Planilha.Protect
End Property
Public Property Get Cargo() As String
    Cargo = Planilha.[Servidor.Cargo].Value2
End Property
Public Property Let Cargo(ByVal Cargo_ As String)
    Planilha.Unprotect
    Planilha.[Servidor.Cargo] = Cargo_
    Planilha.Protect
End Property

Public Property Get Lotacao() As String
    Lotacao = Planilha.[Servidor.Lotacao].Value2
End Property

Public Property Let Lotacao(ByVal Lotacao_ As String)
    Planilha.Unprotect
    Application.[Servidor.Lotacao] = Lotacao_
    Planilha.Protect
End Property

Public Property Get Exercicio() As String
    Exercicio = Planilha.[Servidor.Exercicio].Value2
End Property

Public Property Let Exercicio(ByVal Exercicio_ As String)
    Planilha.Unprotect
    Application.[Servidor.Exercicio] = Exercicio_
    Planilha.Protect
End Property

Public Property Get SituacaoFuncional() As String
    SituacaoFuncional = Planilha.[Servidor.SituacaoFuncional].Value2
End Property

Public Property Let SituacaoFuncional( _
    ByVal SituacaoFuncional_Servidor As String)
    Planilha.Unprotect
    Application.[Servidor.SituacaoFuncional] = SituacaoFuncional_Servidor
    Planilha.Protect
End Property

Public Property Get SituacaoServidor() As String
    SituacaoServidor = Planilha.[Servidor.SituacaoServidor].Value2
End Property

Public Property Let SituacaoServidor( _
    ByVal SituacaoServidor_Servidor As String)
    Planilha.Unprotect
    Application.[Servidor.SituacaoServidor] = SituacaoServidor_Servidor
    Planilha.Protect
End Property

Public Property Get CodSituacaoServidor() As Integer
    CodSituacaoServidor = Planilha.[Servidor.CodSituacaoServidor].Value2
End Property

Public Property Let CodSituacaoServidor( _
    ByVal CodSituacaoServidor_ As Integer)
    Planilha.Unprotect
    Application.[Servidor.CodSituacaoServidor] = CodSituacaoServidor_
    Planilha.Protect
End Property

Public Property Get CodSituacaoFuncional() As Integer
    CodSituacaoFuncional = Planilha.[Servidor.CodSituacaoFuncional].Value2
End Property

Public Property Let CodSituacaoFuncional( _
    ByVal CodSituacaoFuncional_ As Integer)
    Planilha.Unprotect
    Application.[Servidor.CodSituacaoFuncional] = CodSituacaoFuncional_
    Planilha.Protect
End Property

Public Property Get DataAposentadoria() As Date
    DataAposentadoria = Planilha.[Servidor.DataAposentadoria].Value2
End Property

Public Property Let DataAposentadoria(ByVal DataAposentadoria_ As Date)
    Planilha.Unprotect
    Application.[Servidor.DataAposentadoria] = DataAposentadoria_
    Planilha.Protect
End Property

Public Function MostraLinha(Linha As Integer)
    Planilha.Unprotect
    Planilha.Rows(Linha).EntireRow.Hidden = False
    Planilha.Protect
End Function

Public Function EscondeLinha(Linha As Integer)
    Planilha.Unprotect
    Planilha.Rows(Linha).EntireRow.Hidden = True
    Planilha.Protect
End Function

Public Function LimpaFormulario()

   
    
    Planilha.Unprotect
    Planilha.[Servidor.Cargo] = ""
    Planilha.[Servidor.Lotacao] = ""
    Planilha.[Servidor.Exercicio] = ""
    Planilha.[Servidor.SituacaoFuncional] = ""
    Planilha.[Servidor.SituacaoServidor] = ""
    Planilha.[Servidor.Nome] = ""
    Planilha.[Servidor.DataAposentadoria] = DATA_VAZIA
    
    Planilha.Protect
End Function

