VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoDocumento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirTipoDocumento(strDescricao As String)

  On Error GoTo trata
  Dim objTipoDocumento As datSisMetal.clsTipoDocumento
  '
  Set objTipoDocumento = New datSisMetal.clsTipoDocumento
  
  'Inserir TipoDocumento
  objTipoDocumento.InserirTipoDocumento strDescricao
  '
  Set objTipoDocumento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarTipoDocumento(lngTIPODOCUMENTOID As Long, _
                                strDescricao As String)
                         
  On Error GoTo trata
  Dim objTipoDocumento As datSisMetal.clsTipoDocumento
  Set objTipoDocumento = New datSisMetal.clsTipoDocumento
  
  'Inserir TipoDocumento
  objTipoDocumento.AlterarTipoDocumento lngTIPODOCUMENTOID, _
                                        strDescricao

  '
  Set objTipoDocumento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarTipoDocumento(lngTIPODOCUMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objTipoDocumento As datSisMetal.clsTipoDocumento
  Set objTipoDocumento = New datSisMetal.clsTipoDocumento
  
  'Listar TipoDocumento
  Set ListarTipoDocumento = objTipoDocumento.ListarTipoDocumento(lngTIPODOCUMENTOID)
  '
  Set objTipoDocumento = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirTipoDocumento(lngTIPODOCUMENTOID As Long)

  On Error GoTo trata
  Dim objTipoDocumento As datSisMetal.clsTipoDocumento
  Set objTipoDocumento = New datSisMetal.clsTipoDocumento
  
  'Excluir TipoDocumento
  objTipoDocumento.ExlcuirTipoDocumento lngTIPODOCUMENTOID
  '
  Set objTipoDocumento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

