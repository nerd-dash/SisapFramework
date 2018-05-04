Attribute VB_Name = "modRecursos"
'$Header: $
'******************************************************************************
Option Explicit
'Esse módulo contém todas as globais, constantes e funções auxiliares
    'da aplicação
    
''Variáveis Globais
Public gdsgDesigncao As New clswsDesignacao
Public gfpcFPConversao As New clswsFeriasPremioConversao
Public gdsvServidor As New clswsDadosServidor
Public gsspSisap As New clsSisap
Public gnavNavegador As New clsNavegador
Public gestEstilo As New clsEstilos

''Constantes Globais
Public Const DATA_EM_ABERTO As Date = #12/31/2999#
Public Const DATA_VAZIA As Date = #12:00:00 AM#
Public Const JANELA_SISAP As String = "SISAP"
Public Const TITULO_JANELA_SISAP As String = _
    JANELA_SISAP & " - bhmvsb.prodemge.gov.br"
    
Public Declare PtrSafe Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long



Public Property Get glngPID() As Long
    glngPID = wsDadosFormularios.[frmLogin.PID]
End Property

Public Property Let glngPID(ByVal PID As Long)
    wsDadosFormularios.[frmLogin.PID] = PID
End Property


''Procedure para Testes rápido
Sub A_TableTest()

    PegaVerbasCargoRecebimento

End Sub


''Funções Auxiliares
Public Function intArray(ParamArray numbers() As Variant) As Integer()
   
   Dim i As Long
   Dim intArry() As Integer
   
   For i = LBound(numbers) To UBound(numbers)
        ReDim Preserve intArry(i)
        intArry(i) = CInt(numbers(i))
   Next i
    
    intArray = intArry
End Function

Public Sub PreparaRelease()


    LimpaTodosFormularios
    
    wsGeral.Range("I10") = "Alterar"
    wsGeral.Range("T10") = "0"
    wsDadosFormularios.[frmLogin.Masp] = ""
    wsDadosFormularios.[frmLogin.Senha] = ""
    wsDadosFormularios.[frmLogin.Top] = ""
    wsDadosFormularios.[frmLogin.Left] = ""
    wsDadosFormularios.[frmLogin.PID] = ""
    wsDadosFormularios.[frmLogin.Impressora] = "YQPF"
    wsDadosFormularios.[frmLogin.LembrarSenha] = False
    
    wsDadosServidor.[Servidor.MaspDv] = ""
    wsDadosServidor.[Servidor.Admissao] = ""

    
End Sub

Public Function LimpaTodosFormularios()

    gdsvServidor.LimpaFormulario
    gfpcFPConversao.NovaConversao
    gdsgDesigncao.NovaDesigncao
    wsDadosFormularios.Range("B1:z23").Value2 = ""
    
End Function

Public Function DataEstaEntre(ByVal DataTestada As Date, _
                        ByVal DataInicial As Date, _
                        ByVal DataFinal As Date) As Boolean
                        
        If DataInicial = DATA_VAZIA And DataFinal = DATA_VAZIA Then
         DataEstaEntre = True
        Else
            DataEstaEntre = DataTestada >= DataInicial And _
                DataTestada <= DataFinal
        End If
        
End Function

Public Function PrimerioDiaMes()
    PrimerioDiaMes = DateSerial(Year(Date), Month(Date), 1)
End Function

Public Sub MyAppActive(Handle As Long)
    Dim lngStatus As Long
    lngStatus = SetForegroundWindow(Handle)
End Sub


Public Function EventosHabilitados(ByVal bool As Boolean)
    Application.EnableEvents = bool
    Application.ScreenUpdating = bool
End Function

