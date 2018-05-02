VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEstilos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function Fundo(ByRef rng As Range)

    rng.Interior.ColorIndex = 1
    rng.Borders.LineStyle = xlNone
    
End Function

Public Function FormularioFundo(ByRef rng As Range)

    rng.Interior.ColorIndex = xlColorIndexNone
    rng.Borders.LineStyle = xlNone

    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = 48
    End With
    
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = 48
    End With
    
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = 48
    End With
 
   With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThick
        .ColorIndex = 48
    End With
    
End Function

Public Function FormularioTitulo(ByRef rng As Range)

    rng.Interior.ColorIndex = 15
    rng.Borders.LineStyle = xlNone

    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 1
    End With
    
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 1
    End With
    
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 1
    End With
 
   With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 1
    End With
    
End Function

Public Function FormularioGrupo(ByRef rng As Range)

    rng.Interior.Color = 14277081
    rng.Borders.LineStyle = xlNone

    With rng.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 1
    End With
    
    With rng.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 1
    End With
    
    With rng.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 1
    End With
 
   With rng.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlMedium
        .ColorIndex = 1
    End With
    
End Function
