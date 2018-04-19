VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCargo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Codigo As String
Public Nivel As String
Public SimboloVencimento As String
Public Grau As String
Public CategoriaProfissional As Integer
Public CargaHoraria As Integer
Public Carreira As String



Public Function CriaObjeto(ByVal Cargo As String)
  
    Dim Reg As New RegExp
    Dim Cargos() As String
    
    With Reg
        .Global = True
        .Pattern = "(^[A-Z]+)(\d)([A-P])"
        If .Test(Cargo) Then
            Cargo = .Replace(UCase(Cargo), "$1-$2-$3")
            Cargos = Split(Cargo, "-")
            Me.Codigo = Cargos(0)
            Me.Nivel = Me.ParaRomano(Cargos(1))
            Me.SimboloVencimento = Cargos(0) & Cargos(1)
            Me.Grau = Cargos(2)
            Me.CategoriaProfissional = PegaCategoria(Me.Codigo)
            Me.CargaHoraria = PegaCargaHoraria(Me.Codigo)
            Me.Carreira = Left(Cargos(0), 3)
        Else
            gsspSisap.JanelaErro ("Não foi possível validar o simbolo(" & Cargo & ") do cargo inserido!")
            gsspSisap.EncerraSisap
        End If
        

    End With
    

End Function

Public Function ParaRomano(ByVal Num As Integer) As String

    Select Case Num
        Case 1
            ParaRomano = "I"
        Case 2
            ParaRomano = "II"
        Case 3
             ParaRomano = "III"
        Case 4
           ParaRomano = "IV"
        Case 5
           ParaRomano = "V"
        Case 6
           ParaRomano = "VI"
        Case Else
            ParaRomano = ""
    End Select
    
End Function

Public Function PegaCategoria(ByVal cod As String) As Integer

    Select Case cod
        Case "ASB", "ASBD"
            PegaCategoria = 206
        Case "ATB", "ATBD"
            PegaCategoria = 256
        Case "EEB", "EEBD"
            PegaCategoria = 104
        Case Else
            PegaCategoria = 0
    End Select
    
End Function

Public Function PegaCargaHoraria(ByVal cod As String) As Integer

    Select Case cod
        Case "ASB", "ASBD", "ATB", "ATBD"
            PegaCargaHoraria = 30
        Case "ANE", "ANEI", "TDE", "ATE"
            PegaCargaHoraria = 40
        Case Else
            PegaCargaHoraria = 24
    End Select
    
End Function

Public Function ClassificaoOrcamentaria(ByVal Nivel As Integer, ByVal Modalidade As Integer) As Double

    
    If Carreira = "ASB" Then
    
        ClassificaoOrcamentaria = 36102122137#
    
    ElseIf Carreira = "ATB" Then
    
    
        ClassificaoOrcamentaria = 36202122140#
           
    ElseIf Modalidade = 4 Then
    
        ClassificaoOrcamentaria = 36602122138#
    
     ElseIf Modalidade = 3 Then
    
        ClassificaoOrcamentaria = 36702122134#
        
    ElseIf Modalidade = 6 Then
    
        ClassificaoOrcamentaria = 36302122142#
        
    ElseIf Nivel = 4 Then
        
        If Carreira = "PEB" Then
            ClassificaoOrcamentaria = 36202122143#
        Else
            ClassificaoOrcamentaria = 36202122140#
        End If
        
    Else
        If Carreira = "PEB" Then
            ClassificaoOrcamentaria = 36102122144#
        Else
            ClassificaoOrcamentaria = 36102122137#
        End If
        
    End If
End Function



