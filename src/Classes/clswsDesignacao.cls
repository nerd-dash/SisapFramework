VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clswsDesignacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Planilha As wsDesignacao

Private Sub Class_Initialize()
    Set Planilha = wsDesignacao
End Sub

Public Property Get UnidadeAdministrativa() As clsUnidadeAdministrativa

    Set UnidadeAdministrativa = New clsUnidadeAdministrativa
    With UnidadeAdministrativa
    
        If Planilha.[Designacao.UnidadeAdministrativa] < 100 Then
            .UnidadeAdmCod = CLng("4" & Planilha.[Designacao.UnidadeAdministrativa].Value2)
        Else
            .UnidadeAdmCod = CLng("5" & Planilha.[Designacao.UnidadeAdministrativa].Value2)
        End If
        
        .MunicipioCod = CLng(Planilha.[Designacao.Municipio].Value2)
        .ZonaRural = Planilha.[Designacao.UnidadeAdministrativa.ZonaRural].Value2
    End With
End Property

Public Function LetUnidadeAdministrativa(str As String)
    Planilha.[Designacao.UnidadeAdministrativa] = str
End Function

Public Property Get Cargo() As String
    
    Cargo = Planilha.[Designacao.Cargo].Value2
    
End Property

Public Property Let Cargo(str As String)
    
    Planilha.[Designacao.Cargo] = str
    
End Property


Public Property Get CategoriaProfisisonal() As Integer
    
    CategoriaProfisisonal = Planilha.[Designacao.CategoriaProfisisonal].Value2
    
End Property

Public Property Let CategoriaProfisisonal(Num As Integer)
    
    Planilha.[Designacao.CategoriaProfisisonal] = Num
    
End Property

Public Property Get SituacaoExercicio() As Integer
    
    SituacaoExercicio = Planilha.[Designacao.SituacaoExercicio].Value2
    
End Property

Public Property Let SituacaoExercicio(Num As Integer)
    
    Planilha.[Designacao.SituacaoExercicio] = Num
    
End Property

Public Property Get DataInicial() As Date
    
    DataInicial = Planilha.[Designacao.DataInicial].Value2
    
End Property

Public Property Let DataInicial(Data As Date)
    
    Planilha.[Designacao.DataInicial] = Data
    
End Property


Public Property Get DataFinal() As Date
    
    DataFinal = Planilha.[Designacao.DataFinal].Value2
    
End Property

Public Property Let DataFinal(Data As Date)
    
    Planilha.[Designacao.DataFinal] = Data
    
End Property

Public Property Get SubstituidoMaspDv() As Long
    
    SubstituidoMaspDv = Planilha.[Designacao.SubstituidoMaspDv].Value2
    
End Property

Public Property Let SubstituidoMaspDv(Masp As Long)
    
    Planilha.[Designacao.SubstituidoMaspDv] = Masp
    
End Property

Public Property Get SubstituidoAdmissao() As Integer
    
    SubstituidoAdmissao = Planilha.[Designacao.SubstituidoAdmissao].Value2
    
End Property

Public Property Let SubstituidoAdmissao(Adm As Integer)
    
    Planilha.[Designacao.SubstituidoAdmissao] = Adm
    
End Property


Public Property Get SubstituidoGrupoNatureza() As Integer
    
    SubstituidoGrupoNatureza = _
        Planilha.[Designacao.SubstituidoGrupoNatureza].Value2
    
End Property

Public Property Let SubstituidoGrupoNatureza(Grupo As Integer)
    
    Planilha.[Designacao.SubstituidoGrupoNatureza] = Grupo
    
End Property

Public Property Get CargasHorarias() As clschsCargaHoraria
    
    Set CargasHorarias = New clschsCargaHoraria

    Set tbl = Range("Designacao.CargasHoraria")
    Dim Linha As Range
    Dim CargaHoraria As New clsCargaHoraria
    
    For Each Linha In tbl.Rows
    
    
        Set CargaHoraria = New clsCargaHoraria
        If Not IsEmpty(Linha.Value2(1, 1)) And _
            Not Linha.Value2(1, 1) = vbNullString Then
            CargaHoraria.CodGrupo = Linha.Value2(1, 1)
            CargaHoraria.CodNatureza = Linha.Value2(1, 2)
            CargaHoraria.Tipo = Linha.Value2(1, 3)
            CargaHoraria.Nivel = Linha.Value2(1, 4)
            CargaHoraria.Modalidade = Linha.Value2(1, 5)
            CargaHoraria.Materia = Linha.Value2(1, 6)
            CargaHoraria.QuantidadeAulas = Linha.Value2(1, 7)
            CargaHoraria.Turno = Linha.Value2(1, 8)
            
            CargasHorarias.Add CargaHoraria
        End If
        
    Next Linha
    
End Property

Public Property Let CargasHorarias(CargasHorarias_ As clschsCargaHoraria)

    Set tbl = Range("Designacao.CargasHoraria")
    Dim CargaHoraria As New clsCargaHoraria
    
    For Each Linha In tbl.Rows
      
        i = Linha.row - 19

        On Error GoTo TratarErro

        Linha.Value2(1, 2) = CargasHorarias_.Item(i).Cod_Natureza
        Linha.Value2(1, 3) = CargasHorarias_.Item(i).Tipo
        Linha.Value2(1, 4) = CargasHorarias_.Item(i).Nivel
        Linha.Value2(1, 5) = CargasHorarias_.Item(i).Modalidade
        Linha.Value2(1, 6) = CargasHorarias_.Item(i).Materia
        Linha.Value2(1, 7) = CargasHorarias_.Item(i).QuantidadeAulas
        Linha.Value2(1, 8) = CargasHorarias_.Item(i).Turno
        
    Next Linha
    
TratarErro:

    
End Property

Public Function NovaDesigncao()

    Planilha.[Designacao.UnidadeAdministrativa] = 0 '*
    Planilha.[Designacao.Cargo] = "-" '*
    Planilha.[Designacao.CategoriaProfisisonal] = ""
    Planilha.[Designacao.SituacaoExercicio] = "0" '*
    Planilha.[Designacao.DataInicial] = "" '*
    Planilha.[Designacao.DataFinal] = DateSerial(Year(Date), 12, 31) '*
    Planilha.[Designacao.SubstituidoMaspDv] = ""
    Planilha.[Designacao.SubstituidoAdmissao] = ""
    Planilha.[Designacao.SubstituidoGrupoNatureza] = ""
    
    Set tbl = Range("Designacao.CargasHoraria")
    For Each cell In tbl
        If Not cell.HasFormula Then
            cell.value = ""
        End If
    Next cell
    
    AppActivate Application.Caption
End Function

Public Property Get TotalAulasRB() As Integer

    TotalAulasRB = Planilha.[Designacao.Acerto.TotalAulasRB].Value2

End Property

Public Property Let TotalAulasRB(ByVal Aulas_ As Integer)
    wsAcertoDesignacao.Unprotect
    wsAcertoDesignacao.[Designacao.Acerto.TotalAulasRB] = Aulas_
    wsAcertoDesignacao.Protect
End Property

Public Property Get TotalAulasEC() As Integer

    TotalAulasEC = wsAcertoDesignacao.[Designacao.Acerto.TotalAulasEC].Value2

End Property

Public Property Let TotalAulasEC(ByVal Aulas_ As Integer)

    wsAcertoDesignacao.Unprotect
    wsAcertoDesignacao.[Designacao.Acerto.TotalAulasEC] = Aulas_
    wsAcertoDesignacao.Protect

End Property

Public Property Get VerbaRB() As Integer
    value = wsAcertoDesignacao.[Designacao.Acerto.VerbaRB].Value2
    VerbaRB = CInt(IIf(value = vbNullString, 0, value))

End Property
Public Property Get VerbaRBValor() As Double
    value = wsAcertoDesignacao.[Designacao.Acerto.VerbaRB.Valor].Value2
    VerbaRBValor = CDbl(IIf(value = vbNullString, 0, value))

End Property

Public Property Get VerbaEC() As Integer
    value = wsAcertoDesignacao.[Designacao.Acerto.VerbaEC].Value2
    VerbaEC = CInt(IIf(value = vbNullString, 0, value))

End Property
Public Property Get VerbaECValor() As Double
    value = wsAcertoDesignacao.[Designacao.Acerto.VerbaEC.Valor].Value2
    VerbaECValor = CDbl(IIf(value = vbNullString, 0, value))

End Property

Public Property Get VerbaAbonoRB() As Integer
    value = wsAcertoDesignacao.[Designacao.Acerto.VerbaAbonoRB].Value2
    VerbaAbonoRB = CInt(IIf(value = vbNullString, 0, value))

End Property
Public Property Get VerbaAbonoRBValor() As Double
    value = wsAcertoDesignacao.[Designacao.Acerto.VerbaAbonoRB.Valor].Value2
    VerbaAbonoRBValor = CDbl(IIf(value = vbNullString, 0, value))

End Property


Public Property Get VerbaAbonoEC() As Integer
    value = wsAcertoDesignacao.[Designacao.Acerto.VerbaAbonoEC].Value2
    VerbaAbonoEC = CInt(IIf(value = vbNullString, 0, value))

End Property
Public Property Get VerbaAbonoECValor() As Double
    value = wsAcertoDesignacao.[Designacao.Acerto.VerbaAbonoEC.Valor].Value2
    VerbaAbonoECValor = CDbl(IIf(value = vbNullString, 0, value))

End Property

Public Property Get VerbaValeTransporte() As Integer
    value = wsAcertoDesignacao.[Designacao.Acerto.VerbaValeTransporte].Value2
    VerbaRB = CInt(IIf(value = vbNullString, 0, value))

End Property
Public Property Get VerbaValeTransporteValor() As Double
    value = wsAcertoDesignacao.[Designacao.Acerto.VerbaValeTransporte.Valor].Value2
    VerbaValeTransporteValor = CDbl(IIf(value = vbNullString, 0, value))

End Property


