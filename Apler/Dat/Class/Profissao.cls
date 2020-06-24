VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProfissao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirProfissao(strDescricao As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO PROFISSAO (DESCRICAO) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_Aceita) & ")"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProfissao.InserirProfissao]", _
            Err.Description
End Sub

Public Sub AlterarProfissao(lngPROFISSAOID As Long, _
                            strDescricao As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PROFISSAO SET DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita) & " " & _
    "WHERE PKID = " & lngPROFISSAOID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProfissao.AlterarProfissao]", _
            Err.Description
End Sub

Public Function ListarProfissao(lngPROFISSAOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PROFISSAO.* " & _
    "FROM PROFISSAO " & _
    "WHERE PROFISSAO.PKID = " & _
    Formata_Dados(lngPROFISSAOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarProfissao = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProfissao.ListarProfissao]", _
            Err.Description
End Function

Public Sub ExlcuirProfissao(lngPROFISSAOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM PROFISSAO WHERE PKID = " & _
    Formata_Dados(lngPROFISSAOID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProfissao.ExlcuirProfissao]", _
            Err.Description
End Sub


Public Function ListarProfissaoPelaDesc(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PROFISSAO.* " & _
    "FROM PROFISSAO " & _
    "WHERE PROFISSAO.DESCRICAO = " & _
    Formata_Dados(strDescricao, tpDados_Texto, tpNulo_NaoAceita)
  Set ListarProfissaoPelaDesc = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProfissao.ListarProfissaoPelaDesc]", _
            Err.Description
End Function

