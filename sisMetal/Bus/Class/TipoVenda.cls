VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirTipoVenda(strDescricao As String)

  On Error GoTo trata
  Dim objTipoVenda As datSisMetal.clsTipoVenda
  '
  Set objTipoVenda = New datSisMetal.clsTipoVenda
  
  'Inserir TipoVenda
  objTipoVenda.InserirTipoVenda strDescricao
  '
  Set objTipoVenda = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarTipoVenda(lngTIPOVENDAID As Long, _
                                strDescricao As String)
                         
  On Error GoTo trata
  Dim objTipoVenda As datSisMetal.clsTipoVenda
  Set objTipoVenda = New datSisMetal.clsTipoVenda
  
  'Inserir TipoVenda
  objTipoVenda.AlterarTipoVenda lngTIPOVENDAID, _
                                        strDescricao

  '
  Set objTipoVenda = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarTipoVenda(lngTIPOVENDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objTipoVenda As datSisMetal.clsTipoVenda
  Set objTipoVenda = New datSisMetal.clsTipoVenda
  
  'Listar TipoVenda
  Set ListarTipoVenda = objTipoVenda.ListarTipoVenda(lngTIPOVENDAID)
  '
  Set objTipoVenda = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirTipoVenda(lngTIPOVENDAID As Long)

  On Error GoTo trata
  Dim objTipoVenda As datSisMetal.clsTipoVenda
  Set objTipoVenda = New datSisMetal.clsTipoVenda
  
  'Excluir TipoVenda
  objTipoVenda.ExlcuirTipoVenda lngTIPOVENDAID
  '
  Set objTipoVenda = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

