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
Public Function SelecionarFuncionarioPeloPkid(lngFUNCIONARIOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT FUNCIONARIO.*, case FUNCIONARIO.NIVEL when 'SEM' then 'SEM ACESSO AO SISTEMA' when 'COM' then 'COMPRADOR' when 'VEN' then 'VENDEDOR' when 'FIN' then 'FINANCEIRO' when 'ADM' then 'ADMINISTRADOR' when 'DIR' then 'DIRETOR' when 'GER' then 'GERENTE' when 'CAI' then 'CAIXA-RECEBIMENTO' when 'LOJ' then 'LOJA-VENDEDOR' end AS DESCNIVEL " & _
    "FROM FUNCIONARIO " & _
    "WHERE FUNCIONARIO.PESSOAID = " & _
    Formata_Dados(lngFUNCIONARIOID, tpDados_Longo)
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
                              strExcluido As String, _
                              strPercVenda As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE FUNCIONARIO SET " & _
    "USUARIO = " & Formata_Dados(strUsuario, tpDados_Texto) & ", " & _
    "SENHA = " & Formata_Dados(strSenha, tpDados_Texto) & ", " & _
    "NIVEL = " & Formata_Dados(strNivel, tpDados_Texto) & ", " & _
    "INDEXCLUIDO = " & Formata_Dados(strExcluido, tpDados_Texto) & ", " & _
    "PERC_VENDA = " & Formata_Dados(strPercVenda, tpDados_Moeda) & " " & _
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
                              strExcluido As String, _
                              strPercVenda As String)
  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO FUNCIONARIO (PESSOAID, USUARIO, SENHA, NIVEL, INDEXCLUIDO, PERC_VENDA) VALUES (" & _
    " " & Formata_Dados(lngPESSOAID, tpDados_Longo) & _
    ", " & Formata_Dados(strUsuario, tpDados_Texto) & _
    ", " & Formata_Dados(strSenha, tpDados_Texto) & _
    ", " & Formata_Dados(strNivel, tpDados_Texto) & _
    ", " & Formata_Dados(strExcluido, tpDados_Texto) & _
    ", " & Formata_Dados(strPercVenda, tpDados_Moeda) & ")"
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFuncionario.InserirFuncionario]", _
            Err.Description
End Sub

