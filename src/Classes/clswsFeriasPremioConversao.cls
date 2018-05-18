VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clswsFeriasPremioConversao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Planilha As wsFeriasPremioConversao

Public Property Get DataAposentadoria() As Date

    DataAposentadoria = Planilha.[FPConversaoEspecie.DataAposentadoria]

End Property


Public Property Let DataAposentadoria(DataAposentadoria_ As Date)
    Planilha.Unprotect
    Planilha.[FPConversaoEspecie.DataAposentadoria.Campo] = Format(DataAposentadoria_, "ddmmyy")
    Planilha.Protect
End Property


Public Property Get CargaHorariaRB() As Integer

    CargaHorariaRB = Planilha.[FPConversaoEspecie.CargaHorariaRB]

End Property

Public Property Let CargaHorariaRB(ByVal CH_ As Integer)
    Planilha.Unprotect
    Planilha.[FPConversaoEspecie.CargaHorariaRB] = CH_
    Planilha.Protect
End Property

Public Property Let CargaHorariaEC(ByVal CH_ As Integer)
    Planilha.Unprotect
    Planilha.[FPConversaoEspecie.CargaHorariaEC] = CH_
    Planilha.Protect
End Property

Public Property Get CargaHorariaEC() As Integer

    CargaHorariaEC = Planilha.[FPConversaoEspecie.CargaHorariaEC]

End Property
Public Property Let CargaHorariaEX(ByVal CH_ As Integer)
    Planilha.Unprotect
    Planilha.[FPConversaoEspecie.CargaHorariaEX] = CH_
    Planilha.Protect
End Property

Public Property Get CargaHorariaEX() As Integer

    CargaHorariaEX = Planilha.[FPConversaoEspecie.CargaHorariaEX]

End Property
Public Property Let CargaHorariaECEX(ByVal CH_ As Integer)
    Planilha.Unprotect
    Planilha.[FPConversaoEspecie.CargaHorariaECEX] = CH_
    Planilha.Protect
End Property

Public Property Get CargaHorariaECEX() As Integer

    CargaHorariaECEX = Planilha.[FPConversaoEspecie.CargaHorariaECEX]

End Property

Public Property Get VencimentoRB() As Double

    VencimentoRB = Planilha.[FPConversaoEspecie.VencimentoRB]

End Property
Public Property Let VencimentoRB(ByVal VencimentoRB_ As Double)
    Planilha.Unprotect
    Planilha.[FPConversaoEspecie.VencimentoRB] = VencimentoRB_
    Planilha.Protect
End Property

Public Property Get CargoDataAfastamento() As String

    CargoDataAfastamento = Planilha.[FPConversaoEspecie.Cargo.DataAposentadoria]

End Property
Public Property Let CargoDataAfastamento(ByVal CargoDataAfastamento_ As String)
    Planilha.Unprotect
    Planilha.[FPConversaoEspecie.Cargo.DataAposentadoria] = CargoDataAfastamento_
    Planilha.Protect
End Property

Public Property Get DataPublicacao() As Date

    DataPublicacao = Planilha.[FPConversaoEspecie.DataPublicacao]

End Property


Public Property Get SaldoTotalEmDias() As Integer

    SaldoTotalEmDias = Planilha.[FPConversaoEspecie.SaldoTotalEmDias]
    
End Property

Public Function NovaConversao()
    
    Planilha.Unprotect
    
    Planilha.[FPConversaoEspecie.Publicacao.Dia] = ""
    Planilha.[FPConversaoEspecie.Publicacao.Mes] = ""
    Planilha.[FPConversaoEspecie.Publicacao.Ano] = ""
    Planilha.[FPConversaoEspecie.Saldo.Meses] = ""
    Planilha.[FPConversaoEspecie.Saldo.Dias] = ""
    Planilha.[FPConversaoEspecie.Cargo.DataAposentadoria] = ""
    Planilha.[FPConversaoEspecie.CargaHorariaRB] = "00"
    Planilha.[FPConversaoEspecie.CargaHorariaEC] = "00"
    Planilha.[FPConversaoEspecie.CargaHorariaEX] = "00"
    Planilha.[FPConversaoEspecie.CargaHorariaECEX] = "00"
    Planilha.[FPConversaoEspecie.DataAposentadoria.Campo] = ""
    Planilha.[FPConversaoEspecie.OutosCampos.Verbas] = ""
    Planilha.[FPConversaoEspecie.OutosCampos.Quant] = ""
    Planilha.[FPConversaoEspecie.OutosCampos.Porcentagem] = ""
    Planilha.[FPConversaoEspecie.OutosCampos.Funcao] = ""
    Planilha.[FPConversaoEspecie.OutosCampos.Valores] = ""
    Planilha.[FPConversaoEspecie.OutosCampos.OutrasVant.Verba] = ""
    Planilha.[FPConversaoEspecie.OutosCampos.OutrasVant] = ""
    Planilha.[FPConversaoEspecie.VencimentoRB] = "-"
    Planilha.Protect
    
    

End Function

Private Sub Class_Initialize()
    Set Planilha = wsFeriasPremioConversao
End Sub
