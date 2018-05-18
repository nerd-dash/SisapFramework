Attribute VB_Name = "TesteDesignicoes"
Sub TestesDesignicoes()

    Dim Ch1 As New clsCargaHoraria
    Dim Ch2 As New clsCargaHoraria
    Dim Ch3 As New clsCargaHoraria
    Dim Chs As New clschsCargaHoraria

    Application.ScreenUpdating = False
    

'##########################################################

    gdsgDesigncao.NovaDesigncao
    
    modRecursos.EventosHabilitados False
    
    gdsvServidor.MaspDv = 13169842
    gdsvServidor.Admisao = 3
    
    modRecursos.EventosHabilitados True
    
    'Vai buscar dados do servidor
    gdsgDesigncao.LetUnidadeAdministrativa "158330"
    gdsgDesigncao.Cargo = "PEBD1A"
    gdsgDesigncao.SituacaoExercicio = 9
    gdsgDesigncao.CategoriaProfisisonal = 0
    gdsgDesigncao.SubstituidoMaspDv = 0
    gdsgDesigncao.SubstituidoAdmissao = 0
    gdsgDesigncao.SubstituidoGrupoNatureza = 20
    gdsgDesigncao.DataInicial = Date
    gdsgDesigncao.DataFinal = _
        DateSerial(Year(Date), Month(Date) + 2, Day(Date))
    
    Set Ch1 = New clsCargaHoraria
    Set Ch2 = New clsCargaHoraria
    Set Chs = New clschsCargaHoraria
    
    Ch1.CodNatureza = 1
    Ch1.Tipo = 1
    Ch1.Nivel = 3
    Ch1.Modalidade = 2
    Ch1.Materia = 10100
    Ch1.QuantidadeAulas = 12
    Ch1.Turno = 11
    
    Chs.Add Ch1
    
    Ch2.CodNatureza = 1
    Ch2.Tipo = 1
    Ch2.Nivel = 4
    Ch2.Modalidade = 2
    Ch2.Materia = 20100
    Ch2.QuantidadeAulas = 4
    Ch2.Turno = 11
    
    Chs.Add Ch2

    gdsgDesigncao.CargasHorarias = Chs
    IncluirDesignacao

    
    Debug.Assert SISAP.Enter(1, 8) = 8
    
    ' ##################################
    
    gdsgDesigncao.NovaDesigncao
    
    gdsvServidor.MaspDv = 14571624
    gdsvServidor.Admisao = 9
    
    'Vai buscar dados do servidor
    gdsgDesigncao.LetUnidadeAdministrativa "159841"
    gdsgDesigncao.Cargo = "PEBD1A"
    gdsgDesigncao.SituacaoExercicio = 9
    gdsgDesigncao.CategoriaProfisisonal = 0
    gdsgDesigncao.SubstituidoMaspDv = 0
    gdsgDesigncao.SubstituidoAdmissao = 0
    gdsgDesigncao.SubstituidoGrupoNatureza = 0
    gdsgDesigncao.DataFinal = _
        DateSerial(Year(Date), Month(Date) + 2, Day(Date))
    
    Set Ch1 = New clsCargaHoraria
    Set Ch2 = New clsCargaHoraria
    Set Chs = New clschsCargaHoraria
    
    Ch1.CodNatureza = 1
    Ch1.Tipo = 1
    Ch1.Nivel = 4
    Ch1.Modalidade = 2
    Ch1.Materia = 20900
    Ch1.QuantidadeAulas = 4
    Ch1.Turno = 11
    
    Chs.Add Ch1
    
    gdsgDesigncao.CargasHorarias = Chs
    IncluirDesignacao
    
    Debug.Assert SISAP.Enter(1, 8) = 8
    
    
    ' ##################################
    
    gdsgDesigncao.NovaDesigncao
    
    gdsvServidor.MaspDv = 14544852
    gdsvServidor.Admisao = 9
    
    'Vai buscar dados do servidor
    gdsgDesigncao.LetUnidadeAdministrativa "160164"
    gdsgDesigncao.Cargo = "PEBS1A"
    gdsgDesigncao.SituacaoExercicio = 9
    gdsgDesigncao.CategoriaProfisisonal = 0
    gdsgDesigncao.SubstituidoMaspDv = 0
    gdsgDesigncao.SubstituidoAdmissao = 0
    gdsgDesigncao.SubstituidoGrupoNatureza = 0
    gdsgDesigncao.DataFinal = _
        DateSerial(Year(Date), Month(Date) + 2, Day(Date))
    
    Set Ch1 = New clsCargaHoraria
    Set Ch2 = New clsCargaHoraria
    Set Chs = New clschsCargaHoraria
    
    Ch1.CodNatureza = 36
    Ch1.Tipo = 1
    Ch1.Nivel = 4
    Ch1.Modalidade = 6
    Ch1.Materia = 31421
    Ch1.QuantidadeAulas = 5
    Ch1.Turno = 15
    
    Chs.Add Ch1
    
    gdsgDesigncao.CargasHorarias = Chs
    IncluirDesignacao
    
    Debug.Assert SISAP.Enter(1, 8) = 8
    
    ' ##################################
    
    gdsgDesigncao.NovaDesigncao
    
    gdsvServidor.MaspDv = 14198220
    gdsvServidor.Admisao = 2
    
    'Vai buscar dados do servidor
    gdsgDesigncao.LetUnidadeAdministrativa "160164"
    gdsgDesigncao.Cargo = "ATBD1A"
    gdsgDesigncao.SituacaoExercicio = 1
    gdsgDesigncao.CategoriaProfisisonal = 256
    gdsgDesigncao.SubstituidoMaspDv = 5476114
    gdsgDesigncao.SubstituidoAdmissao = 2
    gdsgDesigncao.SubstituidoGrupoNatureza = 14
    gdsgDesigncao.DataInicial = #3/16/2018#
    gdsgDesigncao.DataFinal = #12/31/2018#
    
    Set Ch1 = New clsCargaHoraria
    Set Ch2 = New clsCargaHoraria
    Set Chs = New clschsCargaHoraria
    
    Ch1.CodNatureza = 2
    Ch1.Tipo = 30
    Ch1.Nivel = 0
    Ch1.Modalidade = 0
    Ch1.Materia = 0
    Ch1.QuantidadeAulas = 30
    Ch1.Turno = 13
    
    Chs.Add Ch1
    
    gdsgDesigncao.CargasHorarias = Chs
    IncluirDesignacao
    
    Debug.Assert SISAP.Enter(1, 8) = 8
    
    
    ' ##################################
    
    gdsgDesigncao.NovaDesigncao
    
    gdsvServidor.MaspDv = 13319470
    gdsvServidor.Admisao = 2
    
    'Vai buscar dados do servidor
    gdsgDesigncao.LetUnidadeAdministrativa "159786"
    gdsgDesigncao.Cargo = "EEBD1A"
    gdsgDesigncao.SituacaoExercicio = 24
    gdsgDesigncao.CategoriaProfisisonal = 104
    gdsgDesigncao.SubstituidoMaspDv = 8964694
    gdsgDesigncao.SubstituidoAdmissao = 1
    gdsgDesigncao.SubstituidoGrupoNatureza = 1
    gdsgDesigncao.DataInicial = #4/1/2018#
    gdsgDesigncao.DataFinal = #5/31/2018#
    
    Set Ch1 = New clsCargaHoraria
    Set Ch2 = New clsCargaHoraria
    Set Chs = New clschsCargaHoraria
    
    Ch1.CodNatureza = 2
    Ch1.Tipo = 30
    Ch1.Nivel = 3
    Ch1.Modalidade = 2
    Ch1.Materia = 0
    Ch1.QuantidadeAulas = 24
    Ch1.Turno = 11
    
    Chs.Add Ch1
    
    gdsgDesigncao.CargasHorarias = Chs
    IncluirDesignacao
    
    Debug.Assert SISAP.Enter(1, 8) = 8
    
    Application.ScreenUpdating = True
    
    gdsgDesigncao.NovaDesigncao
End Sub


