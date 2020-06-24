VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGRPagamento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True

Option Explicit

Public Function ExcluirGRPagamento(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM GRPAGAMENTO " & _
    "WHERE GRPAGAMENTO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGRPagamento.ExcluirGRPagamento]", _
            Err.Description
End Function


Public Function SelecionarGRPagamentoPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT GRPAGAMENTO.*, PRONTUARIO.NOME " & _
    "FROM GRPAGAMENTO INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GRPAGAMENTO.PRESTADORID " & _
    "WHERE GRPAGAMENTO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarGRPagamentoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGRPagamento.SelecionarGRPagamentoPeloPkid]", _
            Err.Description
End Function

'''Public Function SelecionarGRPagamentoPeloNumero(strNumero As String) As ADODB.Recordset
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "SELECT GRPAGAMENTO.*, PREDIO.NOME AS NOME_PREDIO " & _
'''    "FROM GRPAGAMENTO LEFT JOIN PREDIO ON PREDIO.PKID = GRPAGAMENTO.PREDIOID " & _
'''    "WHERE GRPAGAMENTO.NUMERO = " & _
'''    Formata_Dados(strNumero, tpDados_Texto)
'''  Set SelecionarGRPagamentoPeloNumero = RunSPReturnRS(strSql)
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsGRPagamento.SelecionarGRPagamentoPeloNome]", _
'''            Err.Description
'''End Function
'''

Public Sub AlterarGRPagamento(lngPKID As Long, _
                              lngPRESTADORID As Long, _
                              strDataInicio As String, _
                              strDataTermino As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE GRPAGAMENTO SET " & _
    "PRESTADORID = " & Formata_Dados(IIf(lngPRESTADORID = 0, "", lngPRESTADORID), tpDados_Longo) & ", " & _
    "DATAINICIO = " & Formata_Dados(strDataInicio, tpDados_DataHora) & ", " & _
    "DATATERMINO = " & Formata_Dados(strDataTermino, tpDados_DataHora) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGRPagamento.AlterarGRPagamento]", _
            Err.Description
End Sub

Public Sub InserirGRPagamento(lngPKID As Long, _
                              lngPRESTADORID As Long, _
                              strDataInicio As String, _
                              strDataTermino As String, _
                              strStatus As String, _
                              strStatusImpressao As String, _
                              strNomeUsu As String)

  On Error GoTo trata
  Dim strSql  As String
  Dim objRs   As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO GRPAGAMENTO (PRESTADORID, DATAINICIO, DATATERMINO, STATUS, STATUSIMPRESSAO, USUARIO) VALUES (" & _
    Formata_Dados(IIf(lngPRESTADORID = 0, "", lngPRESTADORID), tpDados_Longo) & ", " & _
    Formata_Dados(strDataInicio, tpDados_DataHora) & ", " & _
    Formata_Dados(strDataTermino, tpDados_DataHora) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ", " & _
    Formata_Dados(strStatusImpressao, tpDados_Texto) & ", " & _
    Formata_Dados(strNomeUsu, tpDados_Texto) & "); "
  
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPKID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGRPagamento.InserirGRPagamento]", _
            Err.Description
End Sub

'''Public Function SelecionarGRPagamento(lngPKID As Long) As ADODB.Recordset
'''  On Error GoTo trata
'''  '
'''  Dim strSql As String
'''
'''  strSql = "SELECT GRPAGAMENTO.*, CONTROLEACESSO.USUARIO " & _
'''    "FROM GRPAGAMENTO LEFT JOIN CONTROLEACESSO ON CONTROLEACESSO.PKID = GRPAGAMENTO.CONTROLEACESSOID " & _
'''    "WHERE GRPAGAMENTO.PKID = " & _
'''    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
'''  Set SelecionarGRPagamento = RunSPReturnRS(strSql)
'''  '
'''  Exit Function
'''trata:
'''  Err.Raise Err.Number, _
'''            Err.Source & ".[clsGRPagamento.SelecionarGRPagamento]", _
'''            Err.Description
'''End Function





