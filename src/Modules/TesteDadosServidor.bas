Attribute VB_Name = "TesteDadosServidor"
Public Sub TesteDadosdoServidor()

'############################################################

    'gdsvServidor DESLIGADO
    
    modServidor.ServidorLimpaDados
    
    gdsvServidor.MaspDv = 12124921
    gdsvServidor.Admisao = 1
    
    Debug.Assert gdsvServidor.Nome = "MARIA EMILIA DE FREITAS PALHARES PRAIS"
    Debug.Assert gdsvServidor.Cargo = "ATBD1A"
    Debug.Assert gdsvServidor.Lotacao = "0"
    Debug.Assert gdsvServidor.Exercicio = "EE PROFESSOR MINERVINO CESARINO"
    

'############################################################

    'gdsvServidor APOSENTADO
    
    Application.EnableEvents = False

    
    '22.08.2013
    modServidor.ServidorLimpaDados
    
    gdsvServidor.MaspDv = 3789401
    gdsvServidor.Admisao = 2
    gdsvServidor.DataAposentadoria = #3/28/2013#
    
    modServidor.ServidorBuscaNome

    modServidor.ServidorBuscaCargo

    
    modServidor.RotinaPegaLotacao gdsvServidor.DataAposentadoria

    modServidor.RotinaPegaExercicio
    
    Debug.Assert gdsvServidor.Nome = "LUCILENE LEITE GUIMARAES"
    Debug.Assert gdsvServidor.Cargo = "PEB2O"
    Debug.Assert gdsvServidor.Lotacao = "159662"
    Debug.Assert gdsvServidor.Exercicio = "QUADRO TEMPORARIO - 1261"
    
    Application.EnableEvents = True

    gsspSisap.EncerraSisap
'############################################################

    'gdsvServidor DESIGNADO

 
    
    gdsvServidor.MaspDv = 12266870
    gdsvServidor.Admisao = 1

    
    Debug.Assert gdsvServidor.Nome = "JOSINEIDE BEZERRA NOBREGA DA SILVA"
    Debug.Assert gdsvServidor.Cargo = "PEBD1A"
    Debug.Assert gdsvServidor.Lotacao = "0"
    Debug.Assert gdsvServidor.Exercicio = "EE BRASIL"
    
'############################################################

    'gdsvServidor EFETIVO


    
    gdsvServidor.MaspDv = 13521935
    gdsvServidor.Admisao = 1


    Debug.Assert gdsvServidor.Nome = "FLAVIO ARANTES DO AMORIM BARCELOS"
    Debug.Assert gdsvServidor.Cargo = "ANE1B"
    Debug.Assert gdsvServidor.Lotacao = "25"
    Debug.Assert gdsvServidor.Exercicio = "39ª S R E - UBERABA - PORTE II"
     
    
'############################################################

    'gdsvServidor EM AFASTAMENTO PRELIMINAR
    
    gdsvServidor.MaspDv = 2879476
    gdsvServidor.Admisao = 2

    
    Debug.Assert gdsvServidor.Nome = "ROSA OLIVIA CAMILO RAMALHO"
    Debug.Assert gdsvServidor.Cargo = "PEB3P"
    Debug.Assert gdsvServidor.Lotacao = "158526"
    Debug.Assert gdsvServidor.Exercicio = "QUADRO TEMPORARIO - 1261"
  




End Sub

