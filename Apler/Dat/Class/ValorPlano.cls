VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsValorPlano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarValorPlanoPelaDescricao(lngTIPOVALORPLANOID As Long, _
                                                  strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT VALORPLANO.* " & _
    "FROM VALORPLANO " & _
    "WHERE VALORPLANO.DESCRICAO = " & _
      Formata_Dados(strDescricao, tpDados_Texto) & _
      " AND VALORPLANO.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
      " AND VALORPLANO.TIPOPLANOID = " & Formata_Dados(lngTIPOVALORPLANOID, tpDados_Longo)
  Set SelecionarValorPlanoPelaDescricao = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsValorPlano.SelecionarValorPlanoPelaDescricao]", _
            Err.Description
End Function

Public Function ExcluirValorPlano(lngVALORPLANOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM VALORPLANO " & _
    "WHERE VALORPLANO.PKID = " & _
    Formata_Dados(lngVALORPLANOID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsValorPlano.ExcluirValorPlano]", _
            Err.Description
End Function

Public Function SelecionarValorPlanoPeloPkid(lngVALORPLANOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT VALORPLANO.* " & _
    "FROM VALORPLANO " & _
    "WHERE VALORPLANO.PKID = " & _
    Formata_Dados(lngVALORPLANOID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarValorPlanoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsValorPlano.SelecionarValorPlanoPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarValorPlano(lngPKID As Long, _
                             strDescricao As String, _
                             strValorTitular As String, _
                             strValorDependente As String, _
                             strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE VALORPLANO SET " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    "VALORTITULAR = " & Formata_Dados(strValorTitular, tpDados_Moeda) & ", " & _
    "VALORDEPENDENTE = " & Formata_Dados(strValorDependente, tpDados_Moeda) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsValorPlano.AlterarValorPlano]", _
            Err.Description
End Sub

Public Sub InserirValorPlano(lngTIPOPLANOID As Long, _
                             strDescricao As String, _
                             strValorTitular As String, _
                             strValorDependente As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO VALORPLANO (TIPOPLANOID, DESCRICAO, VALORTITULAR, VALORDEPENDENTE, STATUS) VALUES (" & _
    Formata_Dados(lngTIPOPLANOID, tpDados_Longo) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    Formata_Dados(strValorTitular, tpDados_Moeda) & ", " & _
    Formata_Dados(strValorDependente, tpDados_Moeda) & ", " & _
    Formata_Dados("A", tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsValorPlano.InserirValorPlano]", _
            Err.Description
End Sub

Public Function SelecionarValorPlano(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT VALORPLANO.* " & _
    "FROM VALORPLANO " & _
    "WHERE VALORPLANO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarValorPlano = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsValorPlano.SelecionarValorPlano]", _
            Err.Description
End Function

