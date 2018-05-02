VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmFeriasPremioConversao 
   Caption         =   "Férias Prêmio - Conversão em Espécie"
   ClientHeight    =   765
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmFeriasPremioConversao.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmFeriasPremioConversao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Top As Double
Private Left As Double
Private Planilha As wsDadosFormularios

Private Sub UserForm_Initialize()

    Set Planilha = wsDadosFormularios
        
    Top = Planilha.[frmFeriasPremioConversao.Top].Value2
    Left = Planilha.[frmFeriasPremioConversao.Left].Value2
    
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
        Planilha.[frmFeriasPremioConversao.Top].Value2 = .Top
        Planilha.[frmFeriasPremioConversao.Left].Value2 = .Left
    End With
End Sub
Private Sub btnFPConversaoCalcula_Click()
    FPConversaoCalcula
End Sub

