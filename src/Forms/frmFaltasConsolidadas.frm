VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFaltasConsolidadas 
   Caption         =   "Faltas Consolidadas"
   ClientHeight    =   780
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFaltasConsolidadas.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmFaltasConsolidadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Top As Double
Private Left As Double
Private Planilha As Worksheet

Private Sub btnEnviaFaltasConsolidadas_Click()
    If Application.ActiveSheet.Name = wsFaltasConsolidadas.Name Then
        modFaltasConsolidadas.CadastraFaltasConsolidadasNoSistema
    Else
        modFaltasConsolidadasHistorico.CadastraFaltasConsolidadasNoSistema
    End If
    
End Sub

Private Sub btnLimparFaltas_Click()
    Dim FaltasPlan As IRelatorio
    
    If Application.ActiveSheet.Name = wsFaltasConsolidadas.Name Then
        Set FaltasPlan = New clswsFaltasConsolidadas
    Else
        Set FaltasPlan = New clswsFaltasConsHistorico
    End If
    
    
    FaltasPlan.Limpar
    
End Sub

Private Sub UserForm_Activate()
    Set Planilha = wsDadosFormularios

    Top = Planilha.[frmFaltasConsolidadas.Top].Value2
    Left = Planilha.[frmFaltasConsolidadas.Left].Value2
    
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
        Planilha.[frmFaltasConsolidadas.Top].Value2 = .Top
        Planilha.[frmFaltasConsolidadas.Left].Value2 = .Left
    End With
End Sub
