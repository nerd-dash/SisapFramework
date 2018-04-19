VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDesignacao 
   Caption         =   "Designação"
   ClientHeight    =   2220
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4800
   OleObjectBlob   =   "frmDesignacao.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmDesignacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAfastamentoSubs_Click()
    
    Pesquisar_Afastamentos gdsgDesigncao.SubstituidoMaspDv, _
        gdsgDesigncao.SubstituidoAdmissao
End Sub

Private Sub btnCargaHorariaSubs_Click()
    Pesquisar_Carga_Horaria_Vigente Date, gdsgDesigncao.SubstituidoMaspDv, _
        gdsgDesigncao.SubstituidoAdmissao, 52, 65
End Sub

Private Sub btnDesiganacaoEnviar_Click()
    Application.ScreenUpdating = False
    IncluirDesignacao
    Application.ScreenUpdating = True
End Sub

Private Sub btnLimparDesigncao_Click()
    Application.ScreenUpdating = False
    gdsgDesigncao.NovaDesigncao
    Application.ScreenUpdating = True
End Sub

Private Sub btnDesligamento_Click()
    IncluirDesligamentoDesignado
End Sub

Private Sub btnEnviaVerbasAcerto_Click()
    EnviaVerbasDeAcerto
End Sub

Private Sub btnFPSubs_Click()
    With gdsgDesigncao
        If Not (.SubstituidoAdmissao = 0) And _
            Not (.SubstituidoMaspDv = 0) Then
            NavPesquisarFeriasPremio .SubstituidoMaspDv, _
                .SubstituidoAdmissao
        Else
            gsspSisap.JanelaAlerta "Verifique os dados do Substituto!"
        End If
    End With
End Sub
