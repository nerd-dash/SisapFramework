VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCargaHoraria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Private pDataInicio As Date
Private pDataFimPrevisto As Date
Private pDataFimEfetivo As Date
Private pTipo As Integer
Private pNivel As Integer
Private pModalidade As Integer
Private pMateria As Long
Private pQuantidadeAulas As Integer
Private pTurno As Integer
Private pUnidadeAdministrativa As Long
Private pSubstitutoMaspDv As Long
Private pSubstitutoAdm As Integer
Private pSubstitutoGrupoNatureza As Integer
Private pCodGrupo As Long
Private pCodNatureza As Long
Private pDescricao As String

Public Property Get DataInicio() As Date
    
DataInicio = pDataInicio
    
End Property
Public Property Get DataFimPrevisto() As Date
    
DataFimPrevisto = pDataFimPrevisto
    
End Property
Public Property Get DataFimEfetivo() As Date
    
DataFimEfetivo = pDataFimEfetivo
    
End Property
Public Property Get Tipo() As Integer
    
Tipo = pTipo
    
End Property
Public Property Get Nivel() As Integer
    
Nivel = pNivel
    
End Property
Public Property Get Modalidade() As Integer
    
Modalidade = pModalidade
    
End Property
Public Property Get Materia() As Long
    
Materia = pMateria
    
End Property
Public Property Get QuantidadeAulas() As Integer
    
QuantidadeAulas = pQuantidadeAulas
    
End Property
Public Property Get Turno() As Integer
    
Turno = pTurno
    
End Property
Public Property Get UnidadeAdministrativa() As Long
    
UnidadeAdministrativa = pUnidadeAdministrativa
    
End Property
Public Property Get SubstitutoMaspDv() As Long
    
SubstitutoMaspDv = pSubstitutoMaspDv
    
End Property
Public Property Get SubstitutoAdm() As Integer
    
SubstitutoAdm = pSubstitutoAdm
    
End Property
Public Property Get SubstitutoGrupoNatureza() As Integer
    
SubstitutoGrupoNatureza = pSubstitutoGrupoNatureza
    
End Property

Public Property Let DataInicio(ByVal DataInicio_ As Date)
    
pDataInicio = DataInicio_
    
End Property
Public Property Let DataFimPrevisto(ByVal DataFimPrevisto_ As Date)
    
pDataFimPrevisto = DataFimPrevisto_
    
End Property
Public Property Let DataFimEfetivo(ByVal DataFimEfetivo_ As Date)
    
pDataFimEfetivo = DataFimEfetivo_
    
End Property
Public Property Let Tipo(ByVal Tipo_ As Integer)
    
pTipo = Tipo_
    
End Property
Public Property Let Nivel(ByVal Nivel_ As Integer)
    
pNivel = Nivel_
    
End Property
Public Property Let Modalidade(ByVal Modalidade_ As Integer)
    
pModalidade = Modalidade_
    
End Property
Public Property Let Materia(ByVal Materia_ As Long)
    
pMateria = Materia_
    
End Property

Public Property Let QuantidadeAulas(ByVal QuantidadeAulas_ As Integer)
    
pQuantidadeAulas = QuantidadeAulas_
    
End Property

Public Property Let Turno(ByVal Turno_ As Integer)
    
pTurno = Turno_
    
End Property
Public Property Let UnidadeAdministrativa(ByVal UnidadeAdministrativa_ As Long)
    
pUnidadeAdministrativa = UnidadeAdministrativa_
    
End Property
Public Property Let SubstitutoMaspDv(ByVal SubstitutoMaspDv_ As Long)
    
pSubstitutoMaspDv = SubstitutoMaspDv_
    
End Property
Public Property Let SubstitutoAdm(ByVal SubstitutoAdm_ As Integer)
    
pSubstitutoAdm = SubstitutoAdm_
    
End Property
Public Property Let SubstitutoGrupoNatureza(ByVal SubstitutoGrupoNatureza_ As Integer)
    
pSubstitutoGrupoNatureza = SubstitutoGrupoNatureza_
    
End Property

Public Property Get CodGrupo() As Long

    CodGrupo = pCodGrupo

End Property

Public Property Let CodGrupo(ByVal CodGrupo_ As Long)

    pCodGrupo = CInt(CodGrupo_)

End Property

Public Property Get CodNatureza() As Long

    CodNatureza = pCodNatureza

End Property

Public Property Let CodNatureza(ByVal CodNatureza_ As Long)

    pCodNatureza = CInt(CodNatureza_)

End Property

Public Property Get Descricao() As String

    Descricao = pDescricao

End Property

Public Property Let Descricao(ByVal Descricao_ As String)

    pDescricao = Descricao_

End Property




