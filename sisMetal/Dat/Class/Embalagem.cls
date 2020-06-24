VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmbalagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarEmbalagem(lngEMBALAGEMID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT EMBALAGEM.* " & _
    " FROM EMBALAGEM " & _
    " WHERE EMBALAGEM.PKID = " & _
    Formata_Dados(lngEMBALAGEMID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarEmbalagem = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsEmbalagem.SelecionarEmbalagem]"
End Function
Public Sub ExcluirEmbalagem(lngEMBALAGEMID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM EMBALAGEM " & _
    " WHERE PKID = " & lngEMBALAGEMID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsEmbalagem.ExcluirEmbalagem]"
End Sub

Public Sub AlterarEmbalagem(strSigla As String, _
                      strNome As String, _
                      lngEMBALAGEMID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE EMBALAGEM SET " & _
    " NOME = " & Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & _
    ", SIGLA = " & Formata_Dados(strSigla, tpDados_Texto, tpNulo_Aceita) & _
    " WHERE PKID = " & lngEMBALAGEMID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsEmbalagem.AlterarEmbalagem]"
End Sub
Public Sub IncluirEmbalagem(strSigla As String, _
                      strNome As String)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO EMBALAGEM (" & _
    " NOME, SIGLA) VALUES (" & _
    Formata_Dados(strNome, tpDados_Texto, tpNulo_Aceita) & _
    ", " & Formata_Dados(strSigla, tpDados_Texto, tpNulo_Aceita) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsEmbalagem.IncluirEmbalagem]"
End Sub

