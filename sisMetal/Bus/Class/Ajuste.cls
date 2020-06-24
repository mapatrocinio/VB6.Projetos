VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAjuste"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirAjuste(ByVal lngAJUSTEID As Long)
  On Error GoTo trata
  '
  Dim objAjuste       As datSisMetal.clsAjuste
  '
  Set objAjuste = New datSisMetal.clsAjuste
  '
  'Excluir Ajuste
  objAjuste.ExcluirAjuste lngAJUSTEID
  '
  '
  Set objAjuste = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirAjuste(lngAJUSTEID As Long, _
                         lngTIPO_AJUSTEID As Long, _
                         strUsuario As String)
  On Error GoTo trata
  Dim objAjuste       As datSisMetal.clsAjuste
  Dim strData         As String
  '
  Set objAjuste = New datSisMetal.clsAjuste
  '
  strData = Format(Now, "DD/MM/YYYY hh:mm")
  'Inserir Ajuste
  objAjuste.InserirAjuste lngAJUSTEID, _
                          lngTIPO_AJUSTEID, _
                          strData, _
                          strUsuario
  '
  Set objAjuste = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarAjuste(lngAJUSTEID As Long, _
                         lngTIPO_AJUSTEID As Long, _
                         strUsuario As String)
  On Error GoTo trata
  Dim objAjuste As datSisMetal.clsAjuste
  '
  Set objAjuste = New datSisMetal.clsAjuste
  '
  'Alterar Ajuste
  objAjuste.AlterarAjuste lngAJUSTEID, _
                          lngTIPO_AJUSTEID, _
                          strUsuario
  '
  Set objAjuste = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarAjuste(lngAJUSTEID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objAjuste As datSisMetal.clsAjuste
  Set objAjuste = New datSisMetal.clsAjuste

  'Listar Ajuste
  Set ListarAjuste = objAjuste.ListarAjuste(lngAJUSTEID)
  '
  Set objAjuste = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function
