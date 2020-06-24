VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGrupoDespesa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarGrupoDespesa(lngGrupoDespesaId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT GRUPODESPESA.* " & _
    " FROM GRUPODESPESA " & _
    " WHERE GRUPODESPESA.PKID = " & _
    Formata_Dados(lngGrupoDespesaId, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarGrupoDespesa = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsGrupoDespesa.SelecionarGrupoDespesa]"
End Function
Public Sub ExcluirGrupoDespesa(lngGrupoDespesaId As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM GRUPODESPESA " & _
    " WHERE PKID = " & lngGrupoDespesaId
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsGrupoDespesa.ExcluirGrupoDespesa]"
End Sub

Public Sub AlterarGrupoDespesa(lngGrupoDespesaId As Long, _
                               strCodigo As String, _
                               strDescricao As String, _
                               strTipo As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE GRUPODESPESA SET " & _
    " CODIGO = " & Formata_Dados(strCodigo, tpDados_Texto, tpNulo_Aceita) & _
    ", DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", TIPO = " & Formata_Dados(strTipo, tpDados_Texto, tpNulo_Aceita) & _
    " WHERE PKID = " & lngGrupoDespesaId
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsGrupoDespesa.AlterarGrupoDespesa]"
End Sub
Public Sub IncluirGrupoDespesa(strCodigo As String, _
                               strDescricao As String, _
                               strTipo As String)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO GRUPODESPESA (" & _
    " CODIGO, DESCRICAO, TIPO) VALUES (" & _
    Formata_Dados(strCodigo, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strTipo, tpDados_Texto, tpNulo_Aceita) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsGrupoDespesa.IncluirGrupoDespesa]"
End Sub

