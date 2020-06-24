VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinhaFornecedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirLinhaFornecedor(lngLinhaFornecedorId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM LINHA_FORNECEDOR " & _
    "WHERE LINHA_FORNECEDOR.PKID = " & _
    Formata_Dados(lngLinhaFornecedorId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLinhaFornecedor.ExcluirLinhaFornecedor]", _
            Err.Description
End Function

Public Function SelecionarLinhaFornecedorPeloPkid(lngLinhaFornecedorId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT LINHA_FORNECEDOR.*, TIPO_LINHA.NOME AS NOME_LINHA, LINHA.CODIGO AS CODIGO_LINHA  " & _
    "FROM LINHA_FORNECEDOR " & _
    " INNER JOIN LINHA ON LINHA.PKID = LINHA_FORNECEDOR.LINHAID " & _
    " INNER JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID " & _
    "WHERE LINHA_FORNECEDOR.PKID = " & _
    Formata_Dados(lngLinhaFornecedorId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarLinhaFornecedorPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLinhaFornecedor.SelecionarLinhaFornecedorPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarLinhaFornecedor(lngPKID As Long, _
                                  strCodigo As String, _
                                  strPeso As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE LINHA_FORNECEDOR SET " & _
    "CODIGO = " & Formata_Dados(strCodigo, tpDados_Texto) & ", " & _
    "PESO = " & Formata_Dados(strPeso, tpDados_Moeda) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLinhaFornecedor.AlterarLinhaFornecedor]", _
            Err.Description
End Sub

Public Sub InserirLinhaFornecedor(lngLINHAID As Long, _
                                  lngFORNECEDORID As Long, _
                                  strCodigo As String, _
                                  strPeso As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO LINHA_FORNECEDOR (LINHAID, FORNECEDORID, CODIGO, PESO) VALUES (" & _
    Formata_Dados(lngLINHAID, tpDados_Longo) & ", " & _
    Formata_Dados(lngFORNECEDORID, tpDados_Longo) & ", " & _
    Formata_Dados(strCodigo, tpDados_Texto) & ", " & _
    Formata_Dados(strPeso, tpDados_Moeda) & " " & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLinhaFornecedor.InserirLinhaFornecedor]", _
            Err.Description
End Sub

Public Function SelecionarLinhaFornecedor(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT LINHA_FORNECEDOR.* " & _
    "FROM LINHA_FORNECEDOR " & _
    "WHERE LINHA_FORNECEDOR.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarLinhaFornecedor = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLinhaFornecedor.SelecionarLinhaFornecedor]", _
            Err.Description
End Function

