VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ListarTipoConvenioPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objTipoConvenio As datApler.clsTipoConvenio
  Set objTipoConvenio = New datApler.clsTipoConvenio
  
  'Listar TipoConvenio pela descrição
  Set ListarTipoConvenioPelaDesc = objTipoConvenio.ListarTipoConvenioPelaDesc(strDescricao)
  '
  Set objTipoConvenio = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
Public Sub InserirTipoConvenio(strDescricao As String)

  On Error GoTo trata
  Dim objTipoConvenio As datApler.clsTipoConvenio
  '
  Set objTipoConvenio = New datApler.clsTipoConvenio
  
  'Inserir TipoConvenio
  objTipoConvenio.InserirTipoConvenio strDescricao
  '
  Set objTipoConvenio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarTipoConvenio(lngTIPOCONVENIOID As Long, _
                               strDescricao As String)
                         
  On Error GoTo trata
  Dim objTipoConvenio As datApler.clsTipoConvenio
  Set objTipoConvenio = New datApler.clsTipoConvenio
  
  'Inserir TipoConvenio
  objTipoConvenio.AlterarTipoConvenio lngTIPOCONVENIOID, _
                                      strDescricao

  '
  Set objTipoConvenio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarTipoConvenio(lngTIPOCONVENIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objTipoConvenio As datApler.clsTipoConvenio
  Set objTipoConvenio = New datApler.clsTipoConvenio
  
  'Listar TipoConvenio
  Set ListarTipoConvenio = objTipoConvenio.ListarTipoConvenio(lngTIPOCONVENIOID)
  '
  Set objTipoConvenio = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirTipoConvenio(lngTIPOCONVENIOID As Long)

  On Error GoTo trata
  Dim objTipoConvenio As datApler.clsTipoConvenio
  Set objTipoConvenio = New datApler.clsTipoConvenio
  
  'Excluir TipoConvenio
  objTipoConvenio.ExlcuirTipoConvenio lngTIPOCONVENIOID
  '
  Set objTipoConvenio = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

