VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarTipoAjuste(lngTIPO_AJUSTEID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objTipoAjuste As datSisMetal.clsTipoAjuste
  Set objTipoAjuste = New datSisMetal.clsTipoAjuste
  
  'Listar TipoAjuste
  Set SelecionarTipoAjuste = objTipoAjuste.SelecionarTipoAjuste(lngTIPO_AJUSTEID)
  '
  Set objTipoAjuste = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirTipoAjuste(strTipo As String, _
                            strDescricao As String)
  
  On Error GoTo trata
  Dim objTipoAjuste As datSisMetal.clsTipoAjuste
  Set objTipoAjuste = New datSisMetal.clsTipoAjuste
  
  'Inserir TipoAjuste
  objTipoAjuste.IncluirTipoAjuste strTipo, _
                                strDescricao
  '
  Set objTipoAjuste = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarTipoAjuste(strTipo As String, _
                            strDescricao As String, _
                            lngTIPO_AJUSTEID As Long)
  On Error GoTo trata
  Dim objTipoAjuste As datSisMetal.clsTipoAjuste
  Set objTipoAjuste = New datSisMetal.clsTipoAjuste
  
  'Alterar TipoAjuste
  objTipoAjuste.AlterarTipoAjuste strTipo, _
                                strDescricao, _
                                lngTIPO_AJUSTEID

  '
  Set objTipoAjuste = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirTipoAjuste(lngTIPO_AJUSTEID As Long)
  On Error GoTo trata
  Dim objTipoAjuste As datSisMetal.clsTipoAjuste
  Set objTipoAjuste = New datSisMetal.clsTipoAjuste
  
  'Excluir TipoAjuste
  objTipoAjuste.ExcluirTipoAjuste lngTIPO_AJUSTEID

  '
  Set objTipoAjuste = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


