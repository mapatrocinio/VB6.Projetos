VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsFaixaSubProduto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirFaixaSubProduto(lngFaixaSubProdutoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM FAIXA " & _
    "WHERE FAIXA.PKID = " & _
    Formata_Dados(lngFaixaSubProdutoId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFaixaSubProduto.ExcluirFaixaSubProduto]", _
            Err.Description
End Function

Public Function SelecionarFaixaSubProdutoPeloPkid(lngFaixaSubProdutoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT FAIXA.* " & _
    "FROM FAIXA " & _
    "WHERE FAIXA.PKID = " & _
    Formata_Dados(lngFaixaSubProdutoId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarFaixaSubProdutoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFaixaSubProduto.SelecionarFaixaSubProdutoPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarFaixaSubProduto(lngPKID As Long, _
                                  strDescricao As String, _
                                  strDatInicial As String, _
                                  strDataFinal As String, _
                                  strDataDesativacao As String, _
                                  strFxInicial As String, _
                                  strFxFinal As String, _
                                  strValor As String, _
                                  strPrCusto As String, _
                                  strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE FAIXA SET " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    "DATAINICIO = " & Formata_Dados(strDatInicial, tpDados_DataHora) & ", " & _
    "DATAFIM = " & Formata_Dados(strDataFinal, tpDados_DataHora) & ", " & _
    "DATADESATIVACAO = " & Formata_Dados(strDataDesativacao, tpDados_DataHora) & ", " & _
    "FXINICIAL = " & Formata_Dados(strFxInicial, tpDados_Longo) & ", " & _
    "FXFINAL = " & Formata_Dados(strFxFinal, tpDados_Longo) & ", " & _
    "VALOR = " & Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    "PRCUSTO = " & Formata_Dados(strPrCusto, tpDados_Moeda) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFaixaSubProduto.AlterarFaixaSubProduto]", _
            Err.Description
End Sub

Public Sub InserirFaixaSubProduto(lngSUBPRODUTOID As Long, _
                                  strDescricao As String, _
                                  strDatInicial As String, _
                                  strDataFinal As String, _
                                  strFxInicial As String, _
                                  strFxFinal As String, _
                                  strValor As String, _
                                  strPrCusto As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO FAIXA (SUBPRODUTOID, DESCRICAO, DATAINICIO, DATAFIM, FXINICIAL, FXFINAL, VALOR, PRCUSTO, STATUS) VALUES (" & _
    Formata_Dados(lngSUBPRODUTOID, tpDados_Longo) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    Formata_Dados(strDatInicial, tpDados_DataHora) & ", " & _
    Formata_Dados(strDataFinal, tpDados_DataHora) & ", " & _
    Formata_Dados(strFxInicial, tpDados_Longo) & ", " & _
    Formata_Dados(strFxFinal, tpDados_Longo) & ", " & _
    Formata_Dados(strValor, tpDados_Moeda) & ", " & _
    Formata_Dados(strPrCusto, tpDados_Moeda) & ", " & _
    Formata_Dados("A", tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFaixaSubProduto.InserirFaixaSubProduto]", _
            Err.Description
End Sub

Public Function SelecionarFaixaSubProduto(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT FAIXA.* " & _
    "FROM FAIXA " & _
    "WHERE FAIXA.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarFaixaSubProduto = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsFaixaSubProduto.SelecionarFaixaSubProduto]", _
            Err.Description
End Function

