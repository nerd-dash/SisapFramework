Attribute VB_Name = "TesteFaltasConsolidadas"
Option Explicit

Sub TestaBuscaNaTabelaFaltas()

    Dim Faltas As IFaltas
    
    Set Faltas = New clswsFaltasConsolidadas
    
    Faltas.AtualizaFormulario
    
    Faltas.Limpar
    
End Sub
