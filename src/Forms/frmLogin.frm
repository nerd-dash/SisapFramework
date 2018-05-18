VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmLogin 
   Caption         =   "Abrir Nova Tela do Sisap"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4425
   OleObjectBlob   =   "frmLogin.frx":0000
   ShowModal       =   0   'False
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Top As Double
Private Left As Double
Private Planilha As Worksheet


Public Sub btnAbrirSisap_Click()

    If Len(txtLoginMasp.value) > 0 And _
        Len(txtLoginSenha.value) > 0 Then
        Call Shell("TaskKill /PID " & glngPID, vbHide)
        
        glngPID = Shell("C:\Program Files\pw3270\pw3270.exe --session=" & JANELA_SISAP, vbMinimizedNoFocus)
        
        
        Debug.Print "Abrindo instância do Pw3270 com PID :" & glngPID
        
        Application.wait (Now + TimeValue("0:00:05"))
        
        Set Planilha = wsDadosFormularios
        
        
        Planilha.[frmLogin.Masp] = txtLoginMasp.Text
        Planilha.[frmLogin.Impressora] = txtLoginImpressora.Text
        
        If chkLembraSenha.value Then
            
            Planilha.[frmLogin.Senha] = txtLoginSenha.Text
            Planilha.[frmLogin.LembrarSenha] = True
        Else
            Planilha.[frmLogin.Senha] = ""
            Planilha.[frmLogin.LembrarSenha] = False
        End If
        
        
        With gsspSisap
            
            .Envia "SISAP"
            Do
                .Enter 1, 0
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
        
    
        
        
        'Call Shell("TaskKill /glngPid " & glngPid, vbHide)
        'Debug.Print "Finaliza instância do Pw3270 com glngPid :" & glngPid
        
    End If
     

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
    
    Set Planilha = wsDadosFormularios
    
    chkLembraSenha.value = Planilha.[frmLogin.LembrarSenha]
    txtLoginMasp.Text = Planilha.[frmLogin.Masp]
    txtLoginSenha.Text = Planilha.[frmLogin.Senha]
    txtLoginImpressora.Text = Planilha.[frmLogin.Impressora]
    
    Top = Planilha.[frmLogin.Top].Value2
    Left = Planilha.[frmLogin.Left].Value2
    
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

Private Sub UserForm_BeforeDragOver(ByVal Cancel As MSForms.ReturnBoolean, ByVal Control As MSForms.Control, ByVal Data As MSForms.DataObject, ByVal x As Single, ByVal Y As Single, ByVal State As MSForms.fmDragState, ByVal Effect As MSForms.ReturnEffect, ByVal Shift As Integer)
    
 

End Sub

Private Sub UserForm_Click()

End Sub

Private Sub UserForm_Deactivate()


    
End Sub

Private Sub UserForm_MouseMove(ByVal Button As Integer, ByVal Shift As Integer, ByVal x As Single, ByVal Y As Single)

End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    With Me
        Planilha.[frmLogin.Top].Value2 = .Top
        Planilha.[frmLogin.Left].Value2 = .Left
    End With
End Sub

Private Sub UserForm_Terminate()

End Sub
