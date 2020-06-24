VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMovimentacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarMovimentacao(lngMOVIMENTACAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT MOVIMENTACAO.*, CONTADEBITO.DESCRICAO AS DESCRICAOCONTADEBITO, CONTACREDITO.DESCRICAO AS DESCRICAOCONTACREDITO " & _
    " FROM (MOVIMENTACAO LEFT JOIN CONTA AS CONTADEBITO ON CONTADEBITO.PKID = MOVIMENTACAO.CONTADEBITOID) " & _
    " LEFT JOIN CONTA AS CONTACREDITO ON CONTACREDITO.PKID = MOVIMENTACAO.CONTACREDITOID " & _
    " WHERE MOVIMENTACAO.PKID = " & _
    Formata_Dados(lngMOVIMENTACAOID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarMovimentacao = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsMovimentacao.SelecionarMovimentacao]"
End Function
Public Sub ExcluirMovimentacao(lngMOVIMENTACAOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM MOVIMENTACAO " & _
    " WHERE PKID = " & lngMOVIMENTACAOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsMovimentacao.ExcluirMovimentacao]"
End Sub

Public Sub AlterarMovimentacao(strData As String, _
                               strDocumento As String, _
                               strCONTADEBITOID As String, _
                               strCONTACREDITOID As String, _
                               strValor As String, _
                               strDescricao As String, _
                               lngMOVIMENTACAOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE MOVIMENTACAO SET " & _
    " DATA = " & Formata_Dados(strData, tpDados_DataHora, tpNulo_Aceita) & _
    ", DOCUMENTO = " & Formata_Dados(strDocumento, tpDados_Texto, tpNulo_Aceita) & _
    ", CONTADEBITOID = " & Formata_Dados(strCONTADEBITOID, tpDados_Longo, tpNulo_Aceita) & _
    ", CONTACREDITOID = " & Formata_Dados(strCONTACREDITOID, tpDados_Longo, tpNulo_Aceita) & _
    ", VALOR = " & Formata_Dados(strValor, tpDados_Moeda, tpNulo_Aceita) & _
    ", DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    " WHERE PKID = " & lngMOVIMENTACAOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsMovimentacao.AlterarMovimentacao]"
End Sub
Public Sub IncluirMovimentacao(strStatus As String, _
                               strData As String, _
                               strDocumento As String, _
                               strCONTADEBITOID As String, _
                               strCONTACREDITOID As String, _
                               strValor As String, _
                               strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO MOVIMENTACAO (" & _
    " STATUS, DATA, DOCUMENTO, CONTADEBITOID, CONTACREDITOID, VALOR, DESCRICAO) VALUES (" & _
    Formata_Dados(strStatus, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strData, tpDados_DataHora, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDocumento, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strCONTADEBITOID, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strCONTACREDITOID, tpDados_Longo, tpNulo_Aceita) & _
    ", " & Formata_Dados(strValor, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsMovimentacao.IncluirMovimentacao]"
End Sub

