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
  '
  strSql = "SELECT FUNCIONARIO.*, case FUNCIONARIO.NIVEL when 'SEM' then 'SEM ACESSO AO SISTEMA' when 'FIN' then 'FINANCEIRO' when 'ADM' then 'ADMINISTRADOR' when 'DIR' then 'DIRETOR' when 'GER' then 'GERENTE' when 'CAI' then 'CAIXA' when 'ARR' then 'ARRECADADOR' when 'ATE' then 'ATENDENTE' when 'DON' then 'DONO' when 'LEI' then 'LEITURISTA' end AS DESCNIVEL " & _
    "FROM FUNCIONARIO " & _
    "WHERE FUNCIONARIO.PESSOAID = " & _
    Formata_Dados(lngFuncionarioId, tpDados_Longo)
  Set SelecionarFuncionarioPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFuncionario.SelecionarFuncionarioPeloPkid]", _
            Err.Description
End Function

Public Sub AlterarFuncionario(lngPESSOAID As Long, _
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
    "WHERE PESSOAID = " & Formata_Dados(lngPESSOAID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFuncionario.AlterarFuncionario]", _
            Err.Description
End Sub

Public Sub InserirFuncionario(lngPESSOAID As Long, _
                              strUsuario As String, _
                              strNivel As String, _
                              strSenha As String, _
                              strExcluido As String)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO FUNCIONARIO (PESSOAID, USUARIO, SENHA, NIVEL, INDEXCLUIDO) VALUES (" & _
    " " & Formata_Dados(lngPESSOAID, tpDados_Longo) & _
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

Public Sub InserirGerente(lngPESSOAID As Long)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO GERENTE (PESSOAID) VALUES (" & _
    " " & Formata_Dados(lngPESSOAID, tpDados_Longo) & ")"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFuncionario.InserirGerente]", _
            Err.Description
End Sub

Public Sub InserirAtendente(lngPESSOAID As Long)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO ATENDENTE (PESSOAID) VALUES (" & _
    " " & Formata_Dados(lngPESSOAID, tpDados_Longo) & ")"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFuncionario.InserirAtendente]", _
            Err.Description
End Sub

Public Sub InserirArrecadador(lngPESSOAID As Long)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO ARRECADADOR (PESSOAID) VALUES (" & _
    " " & Formata_Dados(lngPESSOAID, tpDados_Longo) & ")"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFuncionario.InserirArrecadador]", _
            Err.Description
End Sub

Public Sub InserirDono(lngPESSOAID As Long)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO DONO (PESSOAID) VALUES (" & _
    " " & Formata_Dados(lngPESSOAID, tpDados_Longo) & ")"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFuncionario.InserirDono]", _
            Err.Description
End Sub

Public Sub InserirLeiturista(lngPESSOAID As Long)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO LEITURISTA (PESSOAID) VALUES (" & _
    " " & Formata_Dados(lngPESSOAID, tpDados_Longo) & ")"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFuncionario.InserirLeiturista]", _
            Err.Description
End Sub
Public Sub InserirCaixa(lngPESSOAID As Long)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO CAIXA (PESSOAID) VALUES (" & _
    " " & Formata_Dados(lngPESSOAID, tpDados_Longo) & ")"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFuncionario.InserirCaixa]", _
            Err.Description
End Sub


