VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSubProdProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirSubProdProduto(lngSubProdProdutoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM SUBPRODUTO " & _
    "WHERE SUBPRODUTO.PKID = " & _
    Formata_Dados(lngSubProdProdutoId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSubProdProduto.ExcluirSubProdProduto]", _
            Err.Description
End Function

Public Function SelecionarSubProdProdutoPeloPkid(lngSubProdProdutoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT SUBPRODUTO.* " & _
    "FROM SUBPRODUTO " & _
    "WHERE SUBPRODUTO.PKID = " & _
    Formata_Dados(lngSubProdProdutoId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarSubProdProdutoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSubProdProduto.SelecionarSubProdProdutoPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarSubProdProduto(lngPKID As Long, _
                                 strDescricao As String, _
                                 strDataDesativacao As String, _
                                 strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE SUBPRODUTO SET " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    "DATADESATIVACAO = " & Formata_Dados(strDataDesativacao, tpDados_DataHora) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSubProdProduto.AlterarSubProdProduto]", _
            Err.Description
End Sub

Public Sub InserirSubProdProduto(lngPRODUTOID As Long, _
                                 strDescricao As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO SUBPRODUTO (PRODUTOID, DESCRICAO, STATUS) VALUES (" & _
    Formata_Dados(lngPRODUTOID, tpDados_Longo) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    Formata_Dados("A", tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSubProdProduto.InserirSubProdProduto]", _
            Err.Description
End Sub

Public Function SelecionarSubProdProduto(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT SUBPRODUTO.* " & _
    "FROM SUBPRODUTO " & _
    "WHERE SUBPRODUTO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarSubProdProduto = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSubProdProduto.SelecionarSubProdProduto]", _
            Err.Description
End Function

