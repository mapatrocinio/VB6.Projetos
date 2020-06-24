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

Public Sub ExcluirAjuste(lngAJUSTEID As Long)
  On Error GoTo trata
  
  '
  Dim strSql As String

  strSql = "DELETE AJUSTE " & _
    " WHERE PKID = " & _
    Formata_Dados(lngAJUSTEID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAjuste.ExcluirAjuste]", _
            Err.Description
End Sub

Public Function ListarAjuste(lngAJUSTEID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT AJUSTE.*, TIPO_AJUSTE.DESCRICAO AS DESCR_AJUSTE " & _
    " FROM AJUSTE " & _
    " LEFT JOIN TIPO_AJUSTE ON TIPO_AJUSTE.PKID = AJUSTE.TIPO_AJUSTEID " & _
    " WHERE AJUSTE.PKID = " & _
    Formata_Dados(lngAJUSTEID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarAjuste = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAjuste.ListarAjuste]", _
            Err.Description
End Function


Public Sub InserirAjuste(lngAJUSTEID As Long, _
                         lngTIPO_AJUSTEID As Long, _
                         strData As String, _
                         strUsuario As String)
  On Error GoTo trata
  '
  Dim objRs As ADODB.Recordset
  Dim strSql As String
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO AJUSTE (TIPO_AJUSTEID, DATA, USUARIO) VALUES (" & _
    Formata_Dados(lngTIPO_AJUSTEID, tpDados_Longo) & ", " & _
    Formata_Dados(strData, tpDados_DataHora) & ", " & _
    Formata_Dados(strUsuario, tpDados_Texto) & ");"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngAJUSTEID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAjuste.InserirAjuste]", _
            Err.Description
End Sub

Public Sub AlterarAjuste(lngAJUSTEID As Long, _
                         lngTIPO_AJUSTEID As Long, _
                         strUsuario As String)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE AJUSTE SET " & _
    "TIPO_AJUSTEID = " & Formata_Dados(lngTIPO_AJUSTEID, tpDados_Longo) & _
    ", USUARIO = " & Formata_Dados(strUsuario, tpDados_Texto) & _
    " WHERE AJUSTE.PKID = " & Formata_Dados(lngAJUSTEID, tpDados_Longo)
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsAjuste.AlterarAjuste]", _
            Err.Description
End Sub
