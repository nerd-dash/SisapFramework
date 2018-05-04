VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmServidor 
   Caption         =   "Opções Rotinas - Servidor"
   ClientHeight    =   1350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5205
   OleObjectBlob   =   "frmServidor.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Top As Double
Private Left As Double
Private Planilha As wsDadosFormularios

Private Sub UserForm_Initialize()

    Set Planilha = wsDadosFormularios
        
    Top = Planilha.[frmServidor.Top].Value2
    Left = Planilha.[frmServidor.Left].Value2
    
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
        Planilha.[frmServidor.Top].Value2 = .Top
        Planilha.[frmServidor.Left].Value2 = .Left
    End With
End Sub

Private Sub btnCargaHoraria_Click()
    NavPesquisarCargaHorariaVigente Date, gdsvServidor.MaspDv, _
        gdsvServidor.Admisao, 52, 65
End Sub

Private Sub btnConsultaIpseng_Click()
    AssistenciaMedicaIpsemgDesativada
End Sub

Private Sub btnContaBancaria_Click()

   NavContaBancaria

End Sub

Private Sub btnPequisaDadosFinanceiros_Click()
    PesquisaHistoricoPagamento
End Sub

Private Sub btnPesquisaDadosServidor_Click()
    NavDadosFuncionais
End Sub
