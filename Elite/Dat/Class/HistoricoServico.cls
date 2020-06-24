VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsHistoricoServico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Function ExcluirHistoricoServico(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  'Delete HISTORICOSERVICO
  strSql = "DELETE FROM HISTORICOSERVICO " & _
    "WHERE HISTORICOSERVICO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsHistoricoServico.ExcluirHistoricoServico]", _
            Err.Description
End Function


Public Function SelecionarHistoricoServicoPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT HISTORICOSERVICO.* " & _
    " FROM HISTORICOSERVICO " & _
    "WHERE HISTORICOSERVICO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarHistoricoServicoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsHistoricoServico.SelecionarHistoricoServicoPeloPkid]", _
            Err.Description
End Function

Public Sub AlterarHistoricoServico(lngPKID As Long, _
                                   strDataHora As String, _
                                   lngPACOTESERVICOID As Long, _
                                   strObservacao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE HISTORICOSERVICO SET " & _
    "DATAHORA = " & Formata_Dados(strDataHora, tpDados_DataHora) & ", " & _
    "PACOTESERVICOID = " & Formata_Dados(lngPACOTESERVICOID, tpDados_Longo) & ", " & _
    "OBSERVACAO = " & Formata_Dados(strObservacao, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsHistoricoServico.AlterarHistoricoServico]", _
            Err.Description
End Sub

Public Sub InserirHistoricoServico(lngPKID As Long, _
                                   strDataHora As String, _
                                   lngPACOTESERVICOID As Long, _
                                   strObservacao As String)

  On Error GoTo trata
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO HISTORICOSERVICO (DATAHORA, PACOTESERVICOID, OBSERVACAO) VALUES (" & _
    Formata_Dados(strDataHora, tpDados_DataHora) & ", " & _
    Formata_Dados(lngPACOTESERVICOID, tpDados_Longo) & ", " & _
    Formata_Dados(strObservacao, tpDados_Texto) & ");"
  '
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPKID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsHistoricoServico.InserirHistoricoServico]", _
            Err.Description
End Sub

