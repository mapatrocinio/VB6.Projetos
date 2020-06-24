VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoEstorno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Sub InserirTipoEstorno(strDescricao As String)

  On Error GoTo trata
  Dim objTipoEstorno As datSisMetal.clsTipoEstorno
  '
  Set objTipoEstorno = New datSisMetal.clsTipoEstorno
  
  'Inserir TipoEstorno
  objTipoEstorno.InserirTipoEstorno strDescricao
  '
  Set objTipoEstorno = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarTipoEstorno(lngTIPOESTORNOID As Long, _
                                strDescricao As String)
                         
  On Error GoTo trata
  Dim objTipoEstorno As datSisMetal.clsTipoEstorno
  Set objTipoEstorno = New datSisMetal.clsTipoEstorno
  
  'Inserir TipoEstorno
  objTipoEstorno.AlterarTipoEstorno lngTIPOESTORNOID, _
                                        strDescricao

  '
  Set objTipoEstorno = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarTipoEstorno(lngTIPOESTORNOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objTipoEstorno As datSisMetal.clsTipoEstorno
  Set objTipoEstorno = New datSisMetal.clsTipoEstorno
  
  'Listar TipoEstorno
  Set ListarTipoEstorno = objTipoEstorno.ListarTipoEstorno(lngTIPOESTORNOID)
  '
  Set objTipoEstorno = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirTipoEstorno(lngTIPOESTORNOID As Long)

  On Error GoTo trata
  Dim objTipoEstorno As datSisMetal.clsTipoEstorno
  Set objTipoEstorno = New datSisMetal.clsTipoEstorno
  
  'Excluir TipoEstorno
  objTipoEstorno.ExlcuirTipoEstorno lngTIPOESTORNOID
  '
  Set objTipoEstorno = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

