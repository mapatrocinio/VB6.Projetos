VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProdutoConvenio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirProdutoConvenio(lngProdutoConvenioId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM PRODUTO " & _
    "WHERE PRODUTO.PKID = " & _
    Formata_Dados(lngProdutoConvenioId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProdutoConvenio.ExcluirProdutoConvenio]", _
            Err.Description
End Function

Public Function SelecionarProdutoConvenioPeloPkid(lngProdutoConvenioId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PRODUTO.* " & _
    "FROM PRODUTO " & _
    "WHERE PRODUTO.PKID = " & _
    Formata_Dados(lngProdutoConvenioId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarProdutoConvenioPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProdutoConvenio.SelecionarProdutoConvenioPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarProdutoConvenio(lngPKID As Long, _
                                  strDescricao As String, _
                                  strPercComissao As String, _
                                  strPercVitalicio As String, _
                                  strDataDesativacao As String, _
                                  strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PRODUTO SET " & _
    "DESCRICAO = " & Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    "PERCCOMISSAO = " & Formata_Dados(strPercComissao, tpDados_Moeda) & ", " & _
    "PERCVITALICIO = " & Formata_Dados(strPercVitalicio, tpDados_Moeda) & ", " & _
    "DATADESATIVACAO = " & Formata_Dados(strDataDesativacao, tpDados_DataHora) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProdutoConvenio.AlterarProdutoConvenio]", _
            Err.Description
End Sub

Public Sub InserirProdutoConvenio(lngCONVENIOID As Long, _
                                  strDescricao As String, _
                                  strPercComissao As String, _
                                  strPercVitalicio As String, _
                                  strDataDesativacao As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO PRODUTO (CONVENIOID, DESCRICAO, PERCCOMISSAO, PERCVITALICIO, DATADESATIVACAO, STATUS) VALUES (" & _
    Formata_Dados(lngCONVENIOID, tpDados_Longo) & ", " & _
    Formata_Dados(strDescricao, tpDados_Texto) & ", " & _
    Formata_Dados(strPercComissao, tpDados_Moeda) & ", " & _
    Formata_Dados(strPercVitalicio, tpDados_Moeda) & ", " & _
    Formata_Dados(strDataDesativacao, tpDados_DataHora) & ", " & _
    Formata_Dados("A", tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProdutoConvenio.InserirProdutoConvenio]", _
            Err.Description
End Sub

Public Function SelecionarProdutoConvenio(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PRODUTO.* " & _
    "FROM PRODUTO " & _
    "WHERE PRODUTO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarProdutoConvenio = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProdutoConvenio.SelecionarProdutoConvenio]", _
            Err.Description
End Function

