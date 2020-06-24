VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPlano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True


Option Explicit

Public Function ExcluirPlano(lngPlanoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM PLANO " & _
    "WHERE PLANO.PKID = " & _
    Formata_Dados(lngPlanoId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPlano.ExcluirPlano]", _
            Err.Description
End Function

Public Function SelecionarPlanoPeloPkid(lngPlanoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PLANO.* " & _
    "FROM PLANO " & _
    "WHERE PLANO.PKID = " & _
    Formata_Dados(lngPlanoId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarPlanoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPlano.SelecionarPlanoPeloPkid]", _
            Err.Description
End Function

Public Function SelecionarPlanoPelaDescricao(strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PLANO.* " & _
    "FROM PLANO " & _
    "WHERE PLANO.DESCRICAO = " & _
      Formata_Dados(strDescricao, tpDados_Texto) & _
      " AND PLANO.STATUS = " & Formata_Dados("A", tpDados_Texto)
  Set SelecionarPlanoPelaDescricao = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPlano.SelecionarPlanoPelaDescricao]", _
            Err.Description
End Function


Public Sub AlterarPlano(lngPKID As Long, _
                        strDescricao As String, _
                        strObservacao As String, _
                        strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PLANO SET " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    "OBSERVACAO = " & Formata_Dados(strObservacao, tpDados_Texto) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPlano.AlterarPlano]", _
            Err.Description
End Sub

Public Sub InserirPlano(strDescricao As String, _
                        strObservacao As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO PLANO (DESCRICAO, OBSERVACAO, STATUS) VALUES (" & _
    Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    Formata_Dados(strObservacao, tpDados_Texto) & ", " & _
    Formata_Dados("A", tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPlano.InserirPlano]", _
            Err.Description
End Sub

Public Function SelecionarPlano(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PLANO.* " & _
    "FROM PLANO " & _
    "WHERE PLANO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarPlano = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPlano.SelecionarPlano]", _
            Err.Description
End Function



