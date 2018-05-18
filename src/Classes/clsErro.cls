VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Implements IErro

Option Explicit

Public Function IErro_LogaErro(ByRef Erro As ErrObject, ByVal Identificador As String)
    
    Open "ErrorLog.txt" For Append As #1

    Print #1, Now & vbTab & Identificador & vbTab _
        & Erro.Number & vbTab & Erro.Description

    Close #1
    
    Debug.Print "Logando Erro " & Identificador
    
    MsgBox "Ocorreu o Erro cód. " & Erro.Number & " " & Erro.Description, _
        vbCritical, Identificador
    
    
    gsspSisap.EncerraSisap True
    
    
End Function
