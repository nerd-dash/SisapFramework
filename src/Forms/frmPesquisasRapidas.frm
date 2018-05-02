VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPesquisasRapidas 
   Caption         =   "Pesquias Rápidas"
   ClientHeight    =   10290
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   7245
   OleObjectBlob   =   "frmPesquisasRapidas.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmPesquisasRapidas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Top As Double
Private Left As Double
Private Planilha As wsDadosFormularios
Private lngTempGMaspDv As Long
Private intTempGAdm As Integer

Private Sub Label3_Click()

End Sub

Private Sub tabCategorias_Change()

End Sub

Private Sub btnAfastamentos_Click()
On Error GoTo ErrorHandler

    AtualizaMaspDvAdm
    NavAfastamentos
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnDadosBancarios_Click()
On Error GoTo ErrorHandler

    AtualizaMaspDvAdm
    NavContaBancaria
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnDadosFinanceirosMesAtual_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavPesquisaDadosFinanceirosCargoRecebimento
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True

End Sub

Private Sub btnDadosFinanceirosMesesAnteriores_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavPesquisaDadosFinanceirosCargoRecebimentoMesAnterior
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnDadosFuncionais_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavDadosFuncionais
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnDadosPessoais_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavDadosPessoais
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True

End Sub

Private Sub btnDesativarAssitMedicaIpsemg_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavDesativarAssitMedicaIPSEMG
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True

End Sub

Private Sub btnDocumentos_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavDocumentos
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True

End Sub

Private Sub btnEndereco_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavEndereco
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnEvolucaoCarreira_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavEvolucaoCarreira
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnExercicio_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavExercicios
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnFaltas_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavFaltasConsolidadas
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnFeriasPremio_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavPesquisarFeriasPremio gdsvServidor.MaspDv, gdsvServidor.Admisao
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnFeriasRegulamentares_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavFeriasRegulamentares
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True

End Sub

Private Sub btnFormacaoEscolar_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavFormacaoEscolar
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnHistoricoPagamento_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavHistoricoDePagamento
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True

End Sub

Private Sub btnLiquidoBancario_Click()

On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavLiquidoBancario
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True


End Sub

Private Sub btnOcorrencias_Click()

On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavOcorrencias
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True


End Sub

Private Sub btnPagamentoSuspensoPorMasp_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavPagamentoSuspensoPorMasp
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True



End Sub


Private Sub btnPesquisarAjustamentoFuncional_Click()
On Error GoTo ErrorHandler

    AtualizaMaspDvAdm
    NavPesquisarAjustamentoFuncional
    
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnPesquisarCargaHoraria_Click()
On Error GoTo ErrorHandler

    AtualizaMaspDvAdm
    NavPesquisarCargaHorariaVigente
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnPesquisarFuncaoEducacao_Click()
On Error GoTo ErrorHandler

    AtualizaMaspDvAdm
    NavPesquisarFuncaoEducacao
    
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnPesquisarMudancaSituacaoExercicio_Click()
On Error GoTo ErrorHandler

    AtualizaMaspDvAdm
    NavPesquisarMudancaSituacaoExercicio
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True

End Sub

Private Sub btnPesquisaServidorPorUnidadeResumida_Click()
On Error GoTo ErrorHandler

    AtualizaMaspDvAdm
    NavPesquisaPorUnidadeSEEResumida
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnPublicacaoInspecaoMedia_Click()
On Error GoTo ErrorHandler

    AtualizaMaspDvAdm
    NavPublicacaoInspecaoMedica
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnSimboloVencimento_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavSimboloVencimento
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True

End Sub

Private Sub btnSituacaoExercicio_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavSituacaoExercicio
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub btnVinculados_Click()
On Error GoTo ErrorHandler
    
    AtualizaMaspDvAdm
    NavVinculadosPorRepresentante
    
    EventosHabilitados True
    
    Exit Sub
    
ErrorHandler:

    gsspSisap.JanelaAlerta "Não foi possível identificar o Masp e/ou admissão!"
    
    EventosHabilitados True
End Sub

Private Sub UserForm_Initialize()

    Set Planilha = wsDadosFormularios
        
    Top = Planilha.[frmPesquisasRapidas.Top].Value2
    Left = Planilha.[frmPesquisasRapidas.Left].Value2
    
    txtMaspDV = Planilha.[frmPesquisasRapidas.MaspDv].Value2
    txtAdm = Planilha.[frmPesquisasRapidas.Adm].Value2
     
    
    With Me
        If Top = 0 And Left = 0 Then
            .Top = Application.Top
            .Left = Application.Left
        Else
            .Top = Top
            .Left = Left
        End If
    End With
    
    lngTempGMaspDv = gdsvServidor.MaspDv
    intTempGAdm = gdsvServidor.Admisao
    
    
End Sub
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    With Me
        Planilha.[frmPesquisasRapidas.Top].Value2 = .Top
        Planilha.[frmPesquisasRapidas.Left].Value2 = .Left
    End With
End Sub

Private Function AtualizaMaspDvAdm()
    
    EventosHabilitados False
    
    gdsvServidor.MaspDv = CLng(txtMaspDV.value)
    gdsvServidor.Admisao = CInt(txtAdm.value)
    
    EventosHabilitados True
    
End Function
Private Sub UserForm_Terminate()

    Planilha.[frmPesquisasRapidas.MaspDv].Value2 = txtMaspDV
    Planilha.[frmPesquisasRapidas.Adm].Value2 = txtAdm

    EventosHabilitados False
    
    gdsvServidor.MaspDv = lngTempGMaspDv
    gdsvServidor.Admisao = intTempGAdm
    
    EventosHabilitados True
End Sub
