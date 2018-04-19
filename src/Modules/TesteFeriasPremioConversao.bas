Attribute VB_Name = "TesteFeriasPremioConversao"
Public Sub TesteFPConversao()
    '##################
    Application.EnableEvents = False
    
    Before
    
    gdsvServidor.MaspDv = 3449907
    gdsvServidor.DataAposentadoria = #2/7/2011#
    gdsvServidor.Admisao = 1
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
        
    
    Debug.Print gdsvServidor.MaspDv
    
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "PEB1D"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 18
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 2
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 0
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0
    
    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB 0, "S"
    Debug.Assert gfpcFPConversao.VencimentoRB = 1567.21
  
    DadosFinanceirosMesesAnteriores
    
    Application.EnableEvents = True

    '##################
    Application.EnableEvents = False
    

    
    Before
    
    gdsvServidor.MaspDv = 2875235
    gdsvServidor.DataAposentadoria = #12/19/2011#
    gdsvServidor.Admisao = 2
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
    
    
    
    Debug.Print gdsvServidor.MaspDv
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "PEBT2G"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 18
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 0
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 0
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0

    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB 0, "S"
    Debug.Assert gfpcFPConversao.VencimentoRB = 1518.94


    Application.EnableEvents = True

    '##################
    Application.EnableEvents = False
    
    
    Before
    
    gdsvServidor.MaspDv = 3007176
    gdsvServidor.DataAposentadoria = #2/11/2011#
    gdsvServidor.Admisao = 1
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
    
    
    
    Debug.Print gdsvServidor.MaspDv
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "PEB2P"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 18
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 0
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 2
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0
    
    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB 0, "S"
    Debug.Assert gfpcFPConversao.VencimentoRB = 2261.93
    
    Application.EnableEvents = True
    
        '##################
    Application.EnableEvents = False
    
    
    Before
    
    gdsvServidor.MaspDv = 1144153
    gdsvServidor.DataAposentadoria = #6/22/2012#
    gdsvServidor.Admisao = 2
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
    
    
    
    Debug.Print gdsvServidor.MaspDv
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "EEB2D"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 40
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 0
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 0
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0
    

    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB 40
    Debug.Assert gfpcFPConversao.VencimentoRB = 2873.2

    Application.EnableEvents = True

    '##################
    Application.EnableEvents = False
    
    
    Before
    
    gdsvServidor.MaspDv = 3477403
    gdsvServidor.DataAposentadoria = #3/3/2011#
    gdsvServidor.Admisao = 1
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
    
    Debug.Print gdsvServidor.MaspDv
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "PEB2O"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 18
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 0
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 0
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0
    
    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB 0, "S"
    Debug.Assert gfpcFPConversao.VencimentoRB = 2206.76
    
    
    Application.EnableEvents = True
    '##################
    Application.EnableEvents = False
    
    
    Before
    
    gdsvServidor.MaspDv = 2498608
    gdsvServidor.DataAposentadoria = #3/12/2012#
    gdsvServidor.Admisao = 2
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
    
    
    
    Debug.Print gdsvServidor.MaspDv
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "PEB2I"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 18
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 2
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 0
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0
    
    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB
    Debug.Assert gfpcFPConversao.VencimentoRB = 1950.46
    
    
    Application.EnableEvents = True
    '##################
    Application.EnableEvents = False
    
    
    Before
    
    gdsvServidor.MaspDv = 3331923
    gdsvServidor.DataAposentadoria = #9/18/2012#
    gdsvServidor.Admisao = 2
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
    
    
    
    Debug.Print gdsvServidor.MaspDv
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "PEB1E"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 18
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 0
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 0
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0
    
    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB
    Debug.Assert gfpcFPConversao.VencimentoRB = 1606.37
    
    Application.EnableEvents = True
    '##################
    Application.EnableEvents = False
    
    
    Before
    
    gdsvServidor.MaspDv = 1849413
    gdsvServidor.DataAposentadoria = #8/2/2011#
    gdsvServidor.Admisao = 1
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
    
    
    
    Debug.Print gdsvServidor.MaspDv
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "ATB2L"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 30
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 0
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 0
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0
    
    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB 0, "S"
    Debug.Assert gfpcFPConversao.VencimentoRB = 1514.19
        
            
    Application.EnableEvents = True
       '##################
    Application.EnableEvents = False
    
    
    Before
    
    gdsvServidor.MaspDv = 6530877
    gdsvServidor.DataAposentadoria = #5/16/2011#
    gdsvServidor.Admisao = 1
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
    
    
    
    Debug.Print gdsvServidor.MaspDv
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "PEBT2D"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 18
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 0
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 0
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0
    
    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB 0, "S"
    Debug.Assert gfpcFPConversao.VencimentoRB = 1410.49
    
            
    Application.EnableEvents = True
    '##################
    Application.EnableEvents = False
    
    
    Before
    
    gdsvServidor.MaspDv = 3128626
    gdsvServidor.DataAposentadoria = #2/6/2012#
    gdsvServidor.Admisao = 1
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
    
    
    
    Debug.Print gdsvServidor.MaspDv
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "ASB3I"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 30
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 0
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 0
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0
    
    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB
    Debug.Assert gfpcFPConversao.VencimentoRB = 1225.05
    
        
    Application.EnableEvents = True
    '##################
    Application.EnableEvents = False
    
    
    Before
    
    gdsvServidor.MaspDv = 6040935
    gdsvServidor.DataAposentadoria = #6/16/2012#
    gdsvServidor.Admisao = 2
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
    
    
    
    Debug.Print gdsvServidor.MaspDv
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "ANE2N"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 40
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 0
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 0
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0
    
    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB
    Debug.Assert gfpcFPConversao.VencimentoRB = 3588.23
    
    
        
    Application.EnableEvents = True
    '##################
    Application.EnableEvents = False
  
    
    Before
    
    gdsvServidor.MaspDv = 3321239
    gdsvServidor.DataAposentadoria = #11/4/2011#
    gdsvServidor.Admisao = 1
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
    
    
    
    Debug.Print gdsvServidor.MaspDv
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "PEB4B"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 18
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 0
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 0
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0
    
    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB 0, "N"
    Debug.Assert gfpcFPConversao.VencimentoRB = 691.8
        

    Application.EnableEvents = True
    '##################
    Application.EnableEvents = False
    
    
    Before
    
    gdsvServidor.MaspDv = 1822303
    gdsvServidor.DataAposentadoria = #12/29/2011#
    gdsvServidor.Admisao = 2
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
    
    
    
    Debug.Print gdsvServidor.MaspDv
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "EEB2E"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 40
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 0
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 0
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0
    
    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB 40, "S"
    Debug.Assert gfpcFPConversao.VencimentoRB = 2945.03
    
        
    Application.EnableEvents = True
    '##################
    Application.EnableEvents = False
    
    
    Before
    
    gdsvServidor.MaspDv = 3127651
    gdsvServidor.DataAposentadoria = #8/6/2012#
    gdsvServidor.Admisao = 1
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
    
    
    
    Debug.Print gdsvServidor.MaspDv
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "ASB3G"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 30
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 0
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 0
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0
    
    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB
    Debug.Assert gfpcFPConversao.VencimentoRB = 1166.01
    
    Application.EnableEvents = True

 

    '##################
    Application.EnableEvents = False
    
    
    Before
    
    
    gdsvServidor.MaspDv = 3128485
    gdsvServidor.DataAposentadoria = #5/28/2012#
    gdsvServidor.Admisao = 1
    BuscaDadosAtuaisServidor
    gfpcFPConversao.DataAposentadoria = gdsvServidor.DataAposentadoria
    
    
    
    Debug.Print gdsvServidor.MaspDv
    
    FPConversaoBuscaPosicionamentoDataAfastamento
    Debug.Assert gfpcFPConversao.CargoDataAfastamento = "ATB4N"
    
    FPConversaoBuscaCargaHoraria
    Debug.Assert gfpcFPConversao.CargaHorariaRB = 30
    Debug.Assert gfpcFPConversao.CargaHorariaEC = 0
    Debug.Assert gfpcFPConversao.CargaHorariaEX = 0
    Debug.Assert gfpcFPConversao.CargaHorariaECEX = 0
    
    modFeriasPremioConversaoEspecie.FPConversaoBuscaVencimentoRB
    Debug.Assert gfpcFPConversao.VencimentoRB = 2079.55
    

    
    
End Sub

Private Function Before()
    With gsspSisap
        gdsvServidor.Nome = ""
        gdsvServidor.Cargo = ""
        gdsvServidor.Lotacao = ""
        gdsvServidor.Exercicio = ""
        gdsvServidor.SituacaoFuncional = ""
        gdsvServidor.SituacaoServidor = ""
        gdsvServidor.CodSituacaoFuncional = 0
    End With
    Debug.Print "FP Teste Before pronto!"
End Function





