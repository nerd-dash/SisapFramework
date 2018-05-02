VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsSisap"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare PtrSafe Function hllapi_init Lib "libhllapi.dll" (ByVal tp As String) As Long
Private Declare PtrSafe Function hllapi_deinit Lib "libhllapi.dll" () As Long
Private Declare PtrSafe Function hllapi_get_revision Lib "libhllapi.dll" () As Long
Private Declare PtrSafe Function hllapi_connect Lib "libhllapi.dll" (ByVal uri As String, ByVal wait As Integer) As Long
Private Declare PtrSafe Function hllapi_disconnect Lib "libhllapi.dll" () As Long
Private Declare PtrSafe Function hllapi_wait_for_ready Lib "libhllapi.dll" (ByVal timeout As Integer) As Long
Private Declare PtrSafe Function hllapi_wait Lib "libhllapi.dll" (ByVal timeout As Integer) As Long
Private Declare PtrSafe Function hllapi_get_message_id Lib "libhllapi.dll" () As Long
Private Declare PtrSafe Function hllapi_is_connected Lib "libhllapi.dll" () As Long
Private Declare PtrSafe Function hllapi_get_screen Lib "libhllapi.dll" (ByVal offset As Integer, ByVal text As String, ByVal length As Integer) As Long
Private Declare PtrSafe Function hllapi_get_screen_at Lib "libhllapi.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
Private Declare PtrSafe Function hllapi_set_text_at Lib "libhllapi.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
Private Declare PtrSafe Function hllapi_cmp_text_at Lib "libhllapi.dll" (ByVal row As Integer, ByVal col As Integer, ByVal text As String) As Long
Private Declare PtrSafe Function hllapi_find_text Lib "libhllapi.dll" (ByVal text As String) As Long
Private Declare PtrSafe Function hllapi_enter Lib "libhllapi.dll" () As Long
Private Declare PtrSafe Function hllapi_pfkey Lib "libhllapi.dll" (ByVal keycode As Integer) As Long
Private Declare PtrSafe Function hllapi_pakey Lib "libhllapi.dll" (ByVal keycode As Integer) As Long
Private Declare PtrSafe Function hllapi_set_session_parameter Lib "libhllapi.dll" (ByVal text As String, ByVal size As Integer, ByVal value As Integer) As Long
Private Declare PtrSafe Function hllapi_set_charset Lib "libhllapi.dll" (ByVal text As String, ByVal size As Integer, ByVal value As Integer) As Long
Private Declare PtrSafe Function hllapi_set_unlock_delay Lib "libhllapi.dll" (ByVal delay As Integer) As Long
Private Declare PtrSafe Function hllapi_getcursor Lib "libhllapi.dll" () As Long
Private Declare PtrSafe Function hllapi_setcursor Lib "libhllapi.dll" (ByVal addr As Integer) As Long
Private Declare PtrSafe Function hllapi_action Lib "libhllapi.dll" (ByVal text As String) As Integer



Private MsgErro As String
Private TempoEspera As Long
Public pnavNavegador As New clsNavegador
Public ptelTela As clsTela

Const NUMERO_LINHAS As Integer = 24
Const NUMERO_COLUNAS As Integer = 80
Const CARACTERES_TELA As Integer = NUMERO_LINHAS * NUMERO_COLUNAS

Const DIGITOS_MASPDV As Integer = 11
Const MASK_MASPDV As String = "0#######"

Const DIGITOS_DATA As Integer = 8
Const MASK_DATA As String = "ddmmyyyy"

Const DIGITOS_DATA_COM_BARRAS As Integer = 14
Const MASK_DATA_COM_BARRAS As String = "00/00/0000"

Const DIGITOS_MES_ANO As Integer = 6
Const MASK_MES_ANO As String = "0000"

Const DIGITOS_VERBA As Integer = 4
Const MASK_VERBA As String = "0000"

Const DIGITOS_OPCOES As Integer = 2
Const MASK_OPCOES As String = "00"

Const DIGITOS_CPF As Integer = 11
Const MASK_CPF As String = "00000000000"

Const DIGITOS_IDENTIFICACAO As Integer = 11
Const MASK_IDENTIFICACAO As String = "00000000000"

Const DIGITOS_COD_UNIDADE_ADMINISTRATIVA As Integer = 8
Const MASK_COD__UNIDADE_ADMINISTRATIVA As String = "0#######"

Const DIGITOS_REAIS As Integer = 10
Const MASK_REAIS As String = "0,00"

Const DIGITOS_REAIS_PONTUADO As Integer = 10
Const MASK_REAIS_PONTUADO As String = "#.###,00"

Private Sub Class_Initialize()

    MsgErro = Space(79)
    TempoEspera = 20
    
    If hllapi_init(JANELA_SISAP & ":a") <> 0 Then
        Call Shell("TaskKill /PID " & glngPID, vbHide)
        JanelaErro "Não foi possível encontrar a Tela do SISAP. Se ela já está aberta feche e abra uma nova tela e refaça a ação. "
        EncerraSisap
    End If
    
    Set ptelTela = New clsTela
    
    If Not ptelTela.CodTela = vbNullString Then
    'Evita problemas com as tela de Login
        pnavNavegador.LocalizaTela ptelTela
    End If

End Sub
Private Sub Class_Terminate()
    hllapi_deinit
End Sub
Public Property Get Tela() As clsTela

    Set ptelTela = New clsTela
    pnavNavegador.LocalizaTela ptelTela
    Set Tela = ptelTela

End Property

Public Property Let Tela(ByRef Tela_ As clsTela)

    pnavNavegador.LocalizaTela Tela_
    ptelTela = Tela_

End Property

Public Function EncerraSisap(Optional Alerta As Boolean = True)
    Class_Terminate
    If Alerta Then
        JanelaErro "O script teve de ser interrompido!"
    End If
    
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    
    End
End Function
Public Static Function Enter(Optional ByVal Repetir As Integer = 1, Optional ByVal IgnoraErro As Integer = -1) As Long
    
    Dim NumeroErro As Long
    
    For i = 1 To Repetir
       hllapi_enter
       NumeroErro = VerificaErroNoSisap(IgnoraErro)
    Next i
    
    Enter = NumeroErro
End Function
Public Function EstaNaTelaInicial() As Boolean

    EstaNaTelaInicial = PegaCampo(8, 5, 38) = "M E N U"

End Function
Public Function JanelaErro(str As String)

    MsgBox str, vbCritical

End Function

Public Function JanelaAlerta(str As String)

    MsgBox str, vbExclamation

End Function

Public Function JanelaInformacao(str As String)

    MsgBox str, vbInformation

End Function

Public Function JanelaConfirmacao(str As String)

    MsgBox str, vbApplicationModal

End Function


Public Function TelaInicial()
    If Me.F3(1, 0) <> 0 Then
        Me.F2 2
        TelaInicial
    End If
End Function

Public Function MostraPosicaoAtual()

    MsgBox "Linha " & Me.RetornaLinha(Me.PosicaoAtual) & " Coluna " & Me.RetornaColuna(Me.PosicaoAtual)

End Function

Public Function PosicaoAtual() As Long

    PosicaoAtual = hllapi_getcursor()

End Function

Public Static Function F2(Optional ByVal Repetir As Integer = 1, Optional ByVal IgnoraErro As Integer = -1) As Long
    Dim NumeroErro As Long
    For i = 1 To Repetir
        Me.VerificaTecladoTravado
        hllapi_pfkey 2
        NumeroErro = VerificaErroNoSisap(IgnoraErro)
    Next i
    F2 = NumeroErro
End Function


Public Static Function F3(Optional ByVal Repetir As Integer = 1, Optional ByVal IgnoraErro As Integer = -1) As Long
    Dim NumeroErro As Long
    For i = 1 To Repetir
        Me.VerificaTecladoTravado
        hllapi_pfkey 3
        NumeroErro = VerificaErroNoSisap(IgnoraErro)
    Next i
    F3 = NumeroErro
End Function

Public Static Function F4(Optional ByVal Repetir As Integer = 1, Optional ByVal IgnoraErro As Integer = -1) As Long
    Dim NumeroErro As Long
    For i = 1 To Repetir
        Me.VerificaTecladoTravado
        hllapi_pfkey 4
        NumeroErro = VerificaErroNoSisap(IgnoraErro)
    Next i
    F4 = NumeroErro
End Function

Public Static Function F5(Optional ByVal Repetir As Integer = 1, Optional ByVal IgnoraErro As Integer = -1) As Long
    Dim NumeroErro As Long
    For i = 1 To Repetir
        Me.VerificaTecladoTravado
        hllapi_pfkey 5
        NumeroErro = VerificaErroNoSisap(IgnoraErro)
    Next i
    F5 = NumeroErro
End Function



Public Static Function F7(Optional ByVal Repetir As Integer = 1, Optional ByVal IgnoraErro As Integer = -1) As Long
    Dim NumeroErro As Long
    For i = 1 To Repetir
        Me.VerificaTecladoTravado
        hllapi_pfkey 7
        NumeroErro = VerificaErroNoSisap(IgnoraErro)
    Next i
    F7 = NumeroErro
End Function

Public Static Function F8(Optional ByVal Repetir As Integer = 1, Optional ByVal IgnoraErro As Integer = -1) As Long
    Dim NumeroErro As Long
    For i = 1 To Repetir
        Me.VerificaTecladoTravado
        hllapi_pfkey 8
        NumeroErro = VerificaErroNoSisap(IgnoraErro)
    Next i
    F8 = NumeroErro
End Function

Public Static Function F9(Optional ByVal Repetir As Integer = 1, Optional ByVal IgnoraErro As Integer = -1) As Long
    Dim NumeroErro As Long
    For i = 1 To Repetir
        Me.VerificaTecladoTravado
        hllapi_pfkey 9
        NumeroErro = VerificaErroNoSisap(IgnoraErro)
    Next i
    F9 = NumeroErro
End Function
Public Static Function F10(Optional ByVal Repetir As Integer = 1, Optional ByVal IgnoraErro As Integer = -1) As Long
    Dim NumeroErro As Long
    For i = 1 To Repetir
        Me.VerificaTecladoTravado
        hllapi_pfkey 10
        NumeroErro = VerificaErroNoSisap(IgnoraErro)
    Next i
    F10 = NumeroErro
End Function


Public Static Function PrimeiroCampo()

    hllapi_action ("firstfield")

End Function

Public Static Function VoltaCampo(Optional ByVal Repetir As Integer = 1)
    For i = 1 To Repetir
        hllapi_action ("PreviousField")
    Next i
End Function


Public Static Function ProximoCampo(Optional ByVal Repetir As Integer = 1)
    For i = 1 To Repetir
        hllapi_action ("NextField")
    Next i
    
End Function
Public Static Function LimpaCampo()

    hllapi_action ("EraseEOF")

End Function
Public Static Function LimparTodosCampos()

    Me.PrimeiroCampo
    posicaoInicial = Me.PosicaoAtual
    
    Do
        Me.LimpaCampo
        Me.ProximoCampo
    
    Loop While Me.PosicaoAtual <> posicaoInicial
    Me.VerificaErroNoSisap

End Function

Public Static Function ProximaLinha(Optional ByVal Repetir As Integer = 1)
    
    novaLinha = 0
    
    For i = 1 To Repetir
    
        If (Me.PosicaoAtual + NUMERO_COLUNAS) > CARACTERES_TELA Then
            novaLinha = (Me.PosicaoAtual + NUMERO_COLUNAS) - CARACTERES_TELA
        Else
            novaLinha = (Me.PosicaoAtual + NUMERO_COLUNAS)
        End If
        
        hllapi_setcursor (novaLinha)
        
    Next i
    
    
End Function

Public Static Function LinhaAnterior(Optional ByVal Repetir As Integer = 1)
    
    novaLinha = 0
    
    For i = 1 To Repetir
    
        If (Me.PosicaoAtual - NUMERO_COLUNAS) < 0 Then
            novaLinha = CARACTERES_TELA - (NUMERO_COLUNAS - Me.PosicaoAtual)
        Else
            novaLinha = (Me.PosicaoAtual - NUMERO_COLUNAS)
        End If
        
        hllapi_setcursor (novaLinha)
        
    Next i
    
    
    
End Function

Public Function Envia(ByVal Comando As String, Optional ByVal Tamanho_doCampo As Integer = 0, Optional ByVal Repetir As Integer = 1)
    For i = 1 To Repetir
        If Tamanho_doCampo > Len(Comando) Then
            Espacos = Space(Tamanho_doCampo - Len(Comando))
            Comando = Comando & Espacos
        End If
        hllapi_set_text_at Me.RetornaLinha(Me.PosicaoAtual), Me.RetornaColuna(Me.PosicaoAtual), Comando
    Next i
End Function

Public Function EnviaNumero(ByVal Numero As Double, Optional ByVal Tamanho_doCampo As Integer = 0, Optional ByVal Repetir As Integer = 1)
    
    Formato = ""
    
    For i = 1 To Tamanho_doCampo
    
        Formato = "0" & Formato
    
    Next i
    
    Comando = Format(Numero, Formato)
    
    For i = 1 To Repetir
        hllapi_set_text_at Me.RetornaLinha(Me.PosicaoAtual), Me.RetornaColuna(Me.PosicaoAtual), Comando
    Next i
End Function

Public Function EnviaMaspDv(ByVal MaspDv As Long)
    Me.VerificaTecladoTravado
   hllapi_set_text_at Me.RetornaLinha(Me.PosicaoAtual), Me.RetornaColuna(Me.PosicaoAtual), Format(MaspDv, MASK_MASPDV)

End Function

Public Function EnviaAdm(ByVal Adm As Integer, Optional ByVal Repetir As Integer = 1)
    For i = 1 To Repetir
        hllapi_set_text_at Me.RetornaLinha(Me.PosicaoAtual), Me.RetornaColuna(Me.PosicaoAtual), Format(Adm, "#")
    Next i
End Function

Public Function EnviaData(ByVal Data As Date, Optional ByVal Repetir As Integer = 1)
    For i = 1 To Repetir
        hllapi_set_text_at Me.RetornaLinha(Me.PosicaoAtual), Me.RetornaColuna(Me.PosicaoAtual), Format(Data, MASK_DATA)
    Next i
End Function

Public Function EnviaOpcao(ByVal Opcao As Integer, Optional ByVal Repetir As Integer = 1)
    For i = 1 To Repetir
        hllapi_set_text_at Me.RetornaLinha(Me.PosicaoAtual), Me.RetornaColuna(Me.PosicaoAtual), Format(Opcao, "0#")
    Next i
End Function

Public Function MarcarOpcao(Optional ByVal Linha As Integer, Optional ByVal Coluna As Integer)
    Me.VerificaTecladoTravado
    
     If Linha & Coluna Then
        hllapi_set_text_at Linha, Coluna, "X"
    Else
        hllapi_set_text_at Me.RetornaLinha(Me.PosicaoAtual), _
        Me.RetornaColuna(Me.PosicaoAtual), "X"
    End If
    
End Function

Public Function Incluir()
    Me.VerificaTecladoTravado
    hllapi_set_text_at Me.RetornaLinha(Me.PosicaoAtual), _
    Me.RetornaColuna(Me.PosicaoAtual), "I"

End Function

Public Function Alterar()
    Me.VerificaTecladoTravado
    hllapi_set_text_at Me.RetornaLinha(Me.PosicaoAtual), _
    Me.RetornaColuna(Me.PosicaoAtual), "A"

End Function
Public Function Excluir()
    Me.VerificaTecladoTravado
    hllapi_set_text_at Me.RetornaLinha(Me.PosicaoAtual), _
    Me.RetornaColuna(Me.PosicaoAtual), "E"

End Function

Public Function Repetir()
   Me.VerificaTecladoTravado
   hllapi_set_text_at Me.RetornaLinha(Me.PosicaoAtual), _
   Me.RetornaColuna(Me.PosicaoAtual), "R"

End Function

Public Function RetornaLinha(ByVal posicao As Long) As Long

    Dim Linha As Long
    
    Linha = (posicao \ NUMERO_COLUNAS) + 1
       
    RetornaLinha = Linha
    
End Function

Public Function RetornaColuna(ByVal posicao As Long) As Long

    Dim Coluna As Long
    
    Coluna = (posicao Mod NUMERO_COLUNAS)
    
    If Coluna = 0 Then
       RetornaColuna = NUMERO_COLUNAS
    Else
        RetornaColuna = Coluna
    End If
    
End Function

Public Function PegaCampo(ByVal Tamanho As Integer, _
    Optional ByVal Linha As Integer, _
    Optional ByVal Coluna As Integer) As String
       
    PegaCampo = Space(Tamanho)
    
    If Linha & Coluna Then
        hllapi_get_screen_at Linha, Coluna, PegaCampo
    Else
        hllapi_get_screen_at Me.RetornaLinha(Me.PosicaoAtual), _
        Me.RetornaColuna(Me.PosicaoAtual), PegaCampo
   End If
    
    PegaCampo = UCase(Trim(PegaCampo))
    
End Function

Public Function DevolveNumero(ByVal Original As String) As Long
    Dim i As Integer
    Dim NumerosApenas As String

    NumerosApenas = ""
    For i = 1 To Len(Original)
        If IsNumeric(Mid(Original, i, 1)) Then
            NumerosApenas = NumerosApenas & Mid(Original, i, 1)
        End If
    Next i
    
    If NumerosApenas <> "" Then
        DevolveNumero = CLng(NumerosApenas)
    Else
        DevolveNumero = 0
    End If
    
End Function

Public Function PegaMaspDv(Optional ByVal Linha As Integer, _
    Optional ByVal Coluna As Integer) As Long
    
    Dim str As String
    
    str = Space(DIGITOS_MASPDV)
    
    If Linha & Coluna Then
        hllapi_get_screen_at Linha, Coluna, str
    Else
        hllapi_get_screen SISAP.PosicaoAtual - 1, str, -1
    End If
        
    PegaMaspDv = Me.DevolveNumero(str)
    
    
End Function

Public Function PegaData(Optional ByVal TamanhoCampo _
    As Integer = DIGITOS_DATA_COM_BARRAS, _
    Optional ByVal Linha As Integer, _
    Optional ByVal Coluna As Integer) As Date
    
    Dim str As String
    Dim Reg As New RegExp
    
    str = Space(TamanhoCampo)
    
    If Linha & Coluna Then
        hllapi_get_screen_at Linha, Coluna, str
    Else
        hllapi_get_screen SISAP.PosicaoAtual - 1, str, -1
    End If
    str = Replace(str, ".", "/")
    str = Replace(str, " ", "")
    
    Problema = 0
    
    Do Until IsDate(str)
    
        Select Case Problema
            Case Is = 0 'TEM ALGUMA SUJEIRA NA STRING
                With Reg
                    .Pattern = "[^\d/]+"
                    str = .Replace(str, "")
                    Problema = Problema + 1
                End With
                
            Case Is = 1 'A STRING VOLTOU APENAS COM BARRAS
                With Reg
                    .Pattern = "[/]+"
                    str = .Replace(str, "")
                    Problema = Problema + 1
                End With
                
            Case Is = 2 'A STRING VOLTOU VAZIA
                If str = "" Then
                    str = DATA_EM_ABERTO
                    Problem = Problema + 1
                End If
            Case Else 'NÃO FOI POSSIVEL CONVERTER
                JanelaErro ("Não foi possível converter a data da Linha " _
                 & Linha & " Coluna " & Coluna)
                EncerraSisap
        End Select
    Loop

    PegaData = CDate(str)
    

End Function

Public Function PegaVerba(Optional ByVal Linha As Integer, _
    Optional ByVal Coluna As Integer) As Long
    
    Dim str As String
    
    str = Space(DIGITOS_VERBA)
    
    If Linha & Coluna Then
        hllapi_get_screen_at Linha, Coluna, str
    Else
        hllapi_get_screen SISAP.PosicaoAtual - 1, str, -1
    End If
        
    PegaVerba = Me.DevolveNumero(str)
    
    
End Function


Public Function PegaCampoNumerico(ByVal Tamanho As Integer, _
    Optional ByVal Linha As Integer, _
    Optional ByVal Coluna As Integer) As Long
    
    Dim str As String
    
    str = Space(Tamanho)
    
    If Linha & Coluna Then
        hllapi_get_screen_at Linha, Coluna, str
    Else
        hllapi_get_screen SISAP.PosicaoAtual - 1, str, -1
    End If
        
    PegaCampoNumerico = Me.DevolveNumero(str)
    
    
End Function

Public Function PegaCampoMoeda(ByVal Tamanho As Integer, _
    ByVal Pontuado As Boolean, _
    Optional ByVal Linha As Integer, _
    Optional ByVal Coluna As Integer) As Double
    
    Dim str As String
    
    str = Space(Tamanho)
    
    If Linha & Coluna Then
        hllapi_get_screen_at Linha, Coluna, str
    Else
        hllapi_get_screen SISAP.PosicaoAtual - 1, str, -1
    End If
        
    str = Trim(str)
    
        
    If str = "" Then
        str = "0"
    End If
    
    PegaCampoMoeda = CDbl(str)
    
    
End Function


Public Function VerificaErroNoSisap(Optional ByVal IgnoraErro As Long = -1) As Long
    
    Dim Erro() As String
    Dim MsgErro As String
    
    MsgErro = Space(80)
    
    Me.VerificaTecladoTravado
       
    hllapi_get_screen_at 1, 1, MsgErro
    
    MsgErro = UCase(Trim(MsgErro))

    Erro() = Split(MsgErro, " ", 2)
    
    If (UBound(Erro) - LBound(Erro) + 1) > 1 Then
        If Erro(0) = "***" Then
            Erro(0) = "999"
        ElseIf Erro(0) = "ISHC001" Then
            Erro(0) = "998"
        ElseIf Erro(0) = "PRODEMGE" _
            Or Erro(0) = "GP32DC27" _
            Or Erro(0) = "GP32DC33" _
            Or Erro(0) = "ATENCAO!!!" Then
            Erro(0) = "997"
        End If
        
        If Not IsNumeric(Erro(0)) Then
            Debug.Print "Mensagem não tratavel " & Erro(0)
            VerificaErro = -1
        Else
            Debug.Print "Mensagem código " & Erro(0)
            VerificaErro = CLng(Erro(0))
        End If
        
        
        
        
        If Not VerificaErro = IgnoraErro And Not IgnoraErro = 0 Then
                    
            Select Case VerificaErro
            
                Case 0
                    'Mensagem Genérica
                
                Case 205
                    'JanelaInformacao Erro(1)
                    
                Case 295
                    'F8 para Próxima Tela
                    
                Case 187
                    JanelaErro Erro(1)
                    EncerraSisap
                    End
                    
                Case 998
                    JanelaErro "Um campo numérico contém algum caractere inválido. Verifique seu script."
                    EncerraSisap
                    End
                    
                Case 999
                    JanelaErro Erro(1)
                    EncerraSisap
                    End
                
                Case Else
                    JanelaErro Erro(1)
                    EncerraSisap
                
            End Select
            
        End If
    Else
        VerificaErro = 0
    End If
    
    Me.PrimeiroCampo
    
    VerificaErroNoSisap = VerificaErro
    

End Function

Public Function VerificaTecladoTravado()
       
   If hllapi_wait_for_ready(TempoEspera) <> 0 Then
    hllapi_action ("kybdreset")
    JanelaErro "O teclado travou. Pode ser que haja algum erro na rotina."
   End If
   
End Function

Public Function VerificaTituloTela(ByVal Titulo As String) As Boolean

    VerificaTituloTela = False
    
    If PegaCampo(50, 4, 12) = UCase(Trim(Left(Titulo, _
        Len(PegaCampo(50, 4, 12))))) Then
        VerificaTituloTela = True
    End If
    

End Function

Public Function PegaTituloTela() As String

    PegaTituloTela = PegaCampo(50, 4, 12)

End Function

Public Function AcessaComando(ByVal Comando As String, _
                                Optional ByVal Atalho As String)
    AchouComando = False
    
    TelaInicial
    Envia Comando
    
    If Enter(1, 999) = 999 Then
        LimpaCampo
        If Not IsEmpty(Atalho) Then
            Envia Atalho
        Else
            Envia "*"
        End If
        Enter 1, 999
    End If
    
    If PegaCampo(9, 6, 37) = "SELECIONE" Then
        Do
            PrimeiroCampo
            For Linha = 8 To 19
                If Comando = PegaCampo(50, Linha, 28) Then
                    MarcarOpcao
                    Enter
                    AchouComando = True
                    Linha = 19
                End If
                ProximoCampo
            Next Linha
            If Not AchouComando Then
                Enter
            End If
        Loop While AchouComando = False
    End If
    PrimeiroCampo
End Function






