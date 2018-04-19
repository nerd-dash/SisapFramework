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
Public glngPid As Long
Public gnavNavegador As New clsNavegador

''Constantes Globais
Public Const DATA_EM_ABERTO As Date = #12/31/2999#
Public Const DATA_VAZIA As Date = #12:00:00 AM#
Public Const JANELA_SISAP As String = "pw3270:A"
Public Const TITULO_JANELA_SISAP As String = _
    JANELA_SISAP & " - bhmvsb.prodemge.gov.br"


''Procedure para Testes rápido
Sub A_TableTest()

   EnviaVerbasDeAcerto
    
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

    wsGeral.Range("I10") = "Alterar"
    wsGeral.Range("T10") = "0"
    wsDadosOcultos.[Taxador.Login.Masp] = ""
    wsDadosOcultos.[Taxador.Login.Senha] = ""
    wsDadosOcultos.[Taxador.Login.Impressora] = "YQPF"
    wsDadosOcultos.[Taxador.Login.LembrarSenha] = False
    
    LimpaTodosFormularios
    
    wsDadosServidor.[Servidor.MaspDv] = ""
    wsDadosServidor.[Servidor.Admissao] = ""
    wsGeral.Range("I10").Select
    
End Sub

Public Function LimpaTodosFormularios()

    gdsvServidor.LimpaFormulario
    gfpcFPConversao.NovaConversao
    gdsgDesigncao.NovaDesigncao
    
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
