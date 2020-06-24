VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFuncionario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarFuncionarioPeloPkid(lngFuncionarioId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT FUNCIONARIO.*, case FUNCIONARIO.NIVEL when 'ARQ' then 'ARQUIVISTA' when 'PRE' then 'PRESTADOR' when 'FIN' then 'FINANCEIRO' when 'ADM' then 'ADMINISTRADOR' when 'DIR' then 'DIRETOR' when 'GER' then 'GERENTE' when 'CAI' then 'CAIXA' when 'LAB' then 'LABORATÓRIO' end AS DESCNIVEL " & _
    "FROM FUNCIONARIO " & _
    "WHERE FUNCIONARIO.PRONTUARIOID = " & _
    Formata_Dados(lngFuncionarioId, tpDados_Longo)
  Set SelecionarFuncionarioPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFuncionario.SelecionarFuncionarioPeloPkid]", _
            Err.Description
End Function

Public Sub AlterarFuncionario(lngPRONTUARIOID As Long, _
                              strUsuario As String, _
                              strNivel As String, _
                              strSenha As String, _
                              strExcluido As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE FUNCIONARIO SET " & _
    "USUARIO = " & Formata_Dados(strUsuario, tpDados_Texto) & ", " & _
    "SENHA = " & Formata_Dados(strSenha, tpDados_Texto) & ", " & _
    "NIVEL = " & Formata_Dados(strNivel, tpDados_Texto) & ", " & _
    "INDEXCLUIDO = " & Formata_Dados(strExcluido, tpDados_Texto) & " " & _
    "WHERE PRONTUARIOID = " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFuncionario.AlterarFuncionario]", _
            Err.Description
End Sub

Public Sub InserirFuncionario(lngPRONTUARIOID As Long, _
                              strUsuario As String, _
                              strNivel As String, _
                              strSenha As String, _
                              strExcluido As String)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO FUNCIONARIO (PRONTUARIOID, USUARIO, SENHA, NIVEL, INDEXCLUIDO) VALUES (" & _
    " " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo) & _
    ", " & Formata_Dados(strUsuario, tpDados_Texto) & _
    ", " & Formata_Dados(strSenha, tpDados_Texto) & _
    ", " & Formata_Dados(strNivel, tpDados_Texto) & _
    ", " & Formata_Dados(strExcluido, tpDados_Texto) & ")"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFuncionario.InserirFuncionario]", _
            Err.Description
End Sub

