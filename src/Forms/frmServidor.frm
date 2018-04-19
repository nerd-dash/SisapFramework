VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmServidor 
   Caption         =   "Opções Rotinas - Servidor"
   ClientHeight    =   1350
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7845
   OleObjectBlob   =   "frmServidor.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmServidor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnCargaHoraria_Click()
    PesquisarCargaHorariaVigente Date, gdsvServidor.MaspDv, _
        gdsvServidor.Admisao, 52, 65
End Sub

Private Sub btnConsultaIpseng_Click()
    With gsspSisap
        DesativarAssitMedicaIPSEMG
        .Envia "01P"
        .Enter
        .EnviaMaspDv gdsvServidor.MaspDv
        .EnviaAdm gdsvServidor.Admisao
        .Enter 1, 4
    
    End With
    
    
End Sub

Private Sub btnContaBancaria_Click()

    With gsspSisap
        PesquisaDadosPessoais
        .EnviaOpcao 3
        .Enter
        .EnviaMaspDv gdsvServidor.MaspDv
        .Enter
    End With

End Sub

Private Sub btnPequisaDadosFinanceiros_Click()
    Pesquisa_Historico_Pagamento
End Sub

Private Sub btnPesquisaDadosServidor_Click()
    
    With gsspSisap
    
        PesquisaDadosServidor
        .Envia "08"
        .Enter
        .EnviaMaspDv gdsvServidor.MaspDv
        .ProximoCampo 8
        .EnviaAdm gdsvServidor.Admisao
        .Enter
    
    End With
End Sub
