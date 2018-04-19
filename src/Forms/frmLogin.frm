VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Abrir Nova Tela do Sisap"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmLogin.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    
End Sub

Public Sub btnAbrirSisap_Click()

    Application.ScreenUpdating = False

    glngPid = Shell("C:\Program Files\pw3270\pw3270.exe", vbMinimizedNoFocus)
    
    Debug.Print "Abrindo instância do Pw3270 com PID :" & glngPid
    
    Application.wait (Now + TimeValue("0:00:05"))
    
    Set Planilha = Sheets("Dados Ocultos")
    
    
    Planilha.[Taxador.Login.Masp] = txtLoginMasp.text
    Planilha.[Taxador.Login.Impressora] = txtLoginImpressora.text
    
    If chkLembraSenha.value Then
        
        Planilha.[Taxador.Login.Senha] = txtLoginSenha.text
        Planilha.[Taxador.Login.LembrarSenha] = True
    Else
        Planilha.[Taxador.Login.Senha] = ""
        Planilha.[Taxador.Login.LembrarSenha] = False
    End If
    
    
    With gsspSisap
        
        .Envia "SISAP"
        Do
            .Enter 1, 997
        Loop While .PegaCampo(8, 1, 2) <> "PRODEMGE"
        
        .Envia txtLoginMasp.value
        
        If Len(txtLoginMasp.value) < 8 Then
            .ProximoCampo
        End If
        
        .Envia txtLoginSenha.value
        
        If Len(txtLoginSenha.value) < 8 Then
            .ProximoCampo
        End If
        .ProximoCampo
        .Envia txtLoginImpressora.value
        .Enter 1, 997
        
        .Envia "SIAP"
        .Enter 1, 997
        .Enter 1, 997
        .EncerraSisap False
    
    End With
    
    Application.ScreenUpdating = True
    
    
    'Call Shell("TaskKill /glngPid " & glngPid, vbHide)
    'Debug.Print "Finaliza instância do Pw3270 com glngPid :" & glngPid
    
    
     
     

End Sub

Private Sub CheckBox1_Click()

End Sub

Private Sub Label2_Click()

End Sub

Private Sub Label3_Click()

End Sub

Private Sub txtLoginMasp_Change()

End Sub

Private Sub txtLoginSenha_Change()

End Sub

Private Sub UserForm_Activate()
    Set Planilha = Sheets("Dados Ocultos")
    chkLembraSenha.value = Planilha.[Taxador.Login.LembrarSenha]
    txtLoginMasp.text = Planilha.[Taxador.Login.Masp]
    txtLoginSenha.text = Planilha.[Taxador.Login.Senha]
    txtLoginImpressora.text = Planilha.[Taxador.Login.Impressora]
End Sub

Private Sub UserForm_Click()

End Sub
