VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDesignacao 
   Caption         =   "Designação"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   5040
   OleObjectBlob   =   "frmDesignacao.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmDesignacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Top As Double
Private Left As Double
Private Planilha As wsDadosFormularios

Private Sub btnSalvaAcerto_Click()
    ImprimeAcertoDesignacao
End Sub

Private Sub btnImprimeAcerto_Click()

End Sub

Private Sub btnNovaDesignacao_Click()
    gdsgDesigncao.NovaDesigncao
End Sub

Private Sub CommandButton2_Click()
    
End Sub

Private Sub UserForm_Initialize()

    Set Planilha = wsDadosFormularios
        
    Top = Planilha.[frmDesignacao.Top].Value2
    Left = Planilha.[frmDesignacao.Left].Value2
    
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
        Planilha.[frmDesignacao.Top].Value2 = .Top
        Planilha.[frmDesignacao.Left].Value2 = .Left
    End With
End Sub

Private Sub btnAfastamentoSubs_Click()
    
    PesquisarAfastamentos gdsgDesigncao.SubstituidoMaspDv, _
        gdsgDesigncao.SubstituidoAdmissao
End Sub

Private Sub btnCargaHorariaSubs_Click()
    NavPesquisarCargaHorariaVigente Date, gdsgDesigncao.SubstituidoMaspDv, _
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
    'PegaVerbasCargoRecebimento
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

Private Sub btnPlanilha_Click()
    ThisWorkbook.Unprotect
    
    Dim bool As Boolean
    bool = Not wsAcertoDesignacao.Visible

    wsAcertoDesignacao.Visible = bool
    
    If bool Then
        wsAcertoDesignacao.Activate
    Else
        wsDesignacao.Activate
    End If
    
    ThisWorkbook.Protect
    
End Sub

Private Sub Frame1_Click()

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Deactivate()

End Sub

