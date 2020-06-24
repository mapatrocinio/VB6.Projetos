VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarConta(lngCONTAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CONTA.* " & _
    " FROM CONTA " & _
    " WHERE CONTA.PKID = " & _
    Formata_Dados(lngCONTAID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarConta = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsConta.SelecionarConta]"
End Function
Public Sub ExcluirConta(lngCONTAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM CONTA " & _
    " WHERE PKID = " & lngCONTAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsConta.ExcluirConta]"
End Sub

Public Sub AlterarConta(strValor As String, _
                        strData As String, _
                        strDescricao As String, _
                        lngCONTAID As Long, _
                        lngTIPOCONTAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE CONTA SET " & _
    " DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", VRSALDO = " & Formata_Dados(strValor, tpDados_Moeda, tpNulo_Aceita) & _
    ", DTSALDO = " & Formata_Dados(strData, tpDados_DataHora, tpNulo_Aceita) & _
    ", TIPOCONTAID = " & Formata_Dados(IIf(lngTIPOCONTAID = 0, "", lngTIPOCONTAID), tpDados_Longo, tpNulo_Aceita) & _
    " WHERE PKID = " & lngCONTAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsConta.AlterarConta]"
End Sub
Public Sub IncluirConta(strValor As String, _
                        strData As String, _
                        strDescricao As String, _
                        lngTIPOCONTAID As Long)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO CONTA (" & _
    " DESCRICAO, VRSALDO, DTSALDO, TIPOCONTAID) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strValor, tpDados_Moeda, tpNulo_Aceita) & _
    ", " & Formata_Dados(strData, tpDados_DataHora, tpNulo_Aceita) & _
    ", " & Formata_Dados(IIf(lngTIPOCONTAID = 0, "", lngTIPOCONTAID), tpDados_Longo, tpNulo_Aceita) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsConta.IncluirConta]"
End Sub

