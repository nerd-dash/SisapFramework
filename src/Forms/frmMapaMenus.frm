VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmMapaMenus 
   Caption         =   "Assistente de Inclusão de Telas no Mapa"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmMapaMenus.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmMapaMenus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnEnviaAdm_Click()
    gsspSisap.EnviaAdm CInt(txtAdmissao.text)
End Sub

Private Sub btnPegaProximaTela_Click()
    Set gnavNavegador = New clsNavegador
    wsMapaMenusSisap.Activate
    gnavNavegador.AdicionaProximaTela
End Sub

Private Sub btnEnviaMaspDv_Click()
    gsspSisap.EnviaMaspDv CLng(txtMaspDv)
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

Private Sub Label2_Click()

End Sub

Private Sub TextBox1_Change()

End Sub
