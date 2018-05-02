VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMapaMenus 
   Caption         =   "Assistente de Inclusão de Telas no Mapa"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmMapaMenus.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmMapaMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Top As Double
Private Left As Double
Private Planilha As wsDadosFormularios

Private Sub UserForm_Initialize()

    Set Planilha = wsDadosFormularios
        
    Top = Planilha.[frmMapaMenus.Top].Value2
    Left = Planilha.[frmMapaMenus.Left].Value2
    
    With Me
        If Top = 0 And Left = 0 Then
            .Top = Application.Top
            .Left = Application.Left
        Else
            .Top = Top
            .Left = Left
        End If
    End With
    
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    With Me
        Planilha.[frmMapaMenus.Top].Value2 = .Top
        Planilha.[frmMapaMenus.Left].Value2 = .Left
    End With
End Sub
Private Sub btnEnviaAdm_Click()
    gsspSisap.EnviaAdm CInt(txtAdmissao.text)
End Sub


Private Sub btnPegaProximaTela_Click()
    Dim funcao As Integer
On Error GoTo ErrorHandler
    
    If txtFuncao.value = vbNullString Then
        txtFuncao.value = "0"
    End If
    
    funcao = CInt(txtFuncao.value)
    
    
    Set gnavNavegador = New clsNavegador
    wsMapaMenusSisap.Activate
    gnavNavegador.AdicionaProximaTela funcao
ErrorHandler:
End Sub

Private Sub btnEnviaMaspDv_Click()
    gsspSisap.EnviaMaspDv CLng(txtMaspDV)
End Sub

Private Sub btnEnviaOpcao_Click()
    gsspSisap.EnviaOpcao CInt(txtOpcao.text)
End Sub

Private Sub btnMarcaX_Click()
    gsspSisap.MarcarOpcao
End Sub

Private Sub btnEnter_Click()
    gsspSisap.Enter 1, 0
End Sub

Private Sub btnString_Click()
    gsspSisap.Envia txtString.text
End Sub

Private Sub CommandButton1_Click()
    gsspSisap.F2
    btnEnviaOpcao_Click
    btnEnter_Click
    btnEnviaMaspDv_Click
    btnEnviaAdm_Click
    btnString_Click
    btnEnter_Click
End Sub


Private Sub UserForm_Click()

End Sub
