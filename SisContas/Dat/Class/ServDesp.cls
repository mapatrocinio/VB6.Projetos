VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsServDesp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarServicoDespertador(lngSERVICODESPERTADORID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT DESPERTADOR.* " & _
    "FROM DESPERTADOR " & _
    "WHERE DESPERTADOR.PKID = " & _
    Formata_Dados(lngSERVICODESPERTADORID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarServicoDespertador = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsServicoDespertador.SelecionarServicoDespertador]", _
            Err.Description
End Function

Public Sub ExcluirServicoDespertador(lngSERVICODESPERTADORID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM DESPERTADOR " & _
    " WHERE PKID = " & lngSERVICODESPERTADORID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsServicoDespertador.ExcluirServicoDespertador]", _
            Err.Description
End Sub
Public Sub AlterarStatusServicoDespertador(lngSERVICODESPERTADORID As Long, _
                                           blnStatus As Boolean)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE DESPERTADOR SET " & _
    " STATUS = " & Formata_Dados(blnStatus, tpDados_Boolean, tpNulo_Aceita) & _
    " WHERE PKID = " & lngSERVICODESPERTADORID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsServicoDespertador.AlterarStatusServicoDespertador]", _
            Err.Description
End Sub

Public Sub AlterarServicoDespertador(lngSERVICODESPERTADORID As Long, _
                                     strData As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE DESPERTADOR SET " & _
    " DATAHORA = " & Formata_Dados(strData, tpDados_DataHora, tpNulo_Aceita) & _
    " WHERE PKID = " & lngSERVICODESPERTADORID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsServicoDespertador.AlterarServicoDespertador]", _
            Err.Description
End Sub
Public Sub IncluirServicoDespertador(lngLOCACAOID As Long, _
                                     strData As String, _
                                     blnStatus As Boolean)
  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "INSERT INTO DESPERTADOR (" & _
    "DATAHORA, ALOCACAOID, STATUS) VALUES (" & _
    Formata_Dados(strData, tpDados_DataHora, tpNulo_Aceita) & _
    ", " & Formata_Dados(lngLOCACAOID, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(blnStatus, tpDados_Boolean, tpNulo_Aceita) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsServicoDespertador.IncluirServicoDespertador]", _
            Err.Description
End Sub


