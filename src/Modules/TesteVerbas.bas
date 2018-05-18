Attribute VB_Name = "TesteVerbas"
Sub TestaBuscarVerbas()

    Application.EnableEvents = False
    gdsvServidor.MaspDv = 3258126
    gdsvServidor.Admisao = 1
    
    PegaVerbasCargoRecebimento
    
    Dim acerto As IVerbas
    
    Set acerto = New clswsAcertoVantagem
    
    acerto.AtualizaFormulario
    
    Debug.Assert acerto.Verbas.Item(1).Verba = 26
    Debug.Assert acerto.Verbas.Item(2).Verba = 28
    Debug.Assert acerto.Verbas.Item(3).Verba = 42
   
    Set acerto = Nothing
    
    Set acerto = New clswsAcertoDesconto
    
    
    Debug.Assert acerto.Verbas.Item(1).Verba = 5502
    Debug.Assert acerto.Verbas.Item(2).Verba = 5647
    Debug.Assert acerto.Verbas.Item(3).Verba = 5643
    
   
    Application.EnableEvents = True
    
End Sub

Public Sub TestaLimpezaDasTabelas()
    Dim acerto As IVerbas
    
    Set acerto = New clswsAcertoVantagem
    acerto.Limpa

    Set acerto = Nothing
    Set acerto = New clswsAcertoDesconto
    acerto.Limpa
    
End Sub

