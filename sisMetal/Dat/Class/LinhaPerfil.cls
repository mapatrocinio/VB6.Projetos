VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinhaPerfil"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarLinha(lngLINHAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT LINHA.*, TIPO_LINHA.NOME, VARA.NOME AS NOME_VARA, VARA.VALOR AS VALOR_VARA " & _
    " FROM LINHA LEFT JOIN VARA ON VARA.PKID = LINHA.VARAID " & _
    " LEFT JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID " & _
    " WHERE LINHA.PKID = " & _
    Formata_Dados(lngLINHAID, tpDados_Longo, tpNulo_NaoAceita)
  Set SelecionarLinha = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsLinhaPerfil.SelecionarLinha]"
End Function
Public Sub ExcluirLinha(lngLINHAID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM LINHA " & _
    " WHERE PKID = " & lngLINHAID
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsLinhaPerfil.ExcluirLinha]"
End Sub

Public Sub AlterarLinha(lngTIPOLINHAID As Long, _
                        strCodigo As String, _
                        lngVARAID As Long, _
                        strPesoVara As String, _
                        lngLINHAID As Long, _
                        strNomeProduto As String, _
                        strLargura As String, _
                        strAba As String, _
                        strEspessura As String)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE LINHA SET " & _
    " TIPO_LINHAID = " & Formata_Dados(lngTIPOLINHAID, tpDados_Longo) & _
    ", CODIGO = " & Formata_Dados(strCodigo, tpDados_Texto) & _
    ", VARAID = " & Formata_Dados(lngVARAID, tpDados_Longo) & _
    ", PESO_VARA = " & Formata_Dados(strPesoVara, tpDados_Moeda) & _
    ", NOME_PRODUTO = " & Formata_Dados(strNomeProduto, tpDados_Texto) & _
    ", LARGURA = " & Formata_Dados(strLargura, tpDados_Moeda) & _
    ", ABA = " & Formata_Dados(strAba, tpDados_Moeda) & _
    ", ESPESSURA = " & Formata_Dados(strEspessura, tpDados_Moeda) & _
    " WHERE PKID = " & Formata_Dados(lngLINHAID, tpDados_Longo)
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsLinhaPerfil.AlterarLinha]"
End Sub
Public Sub IncluirLinha(lngTIPOLINHAID As Long, _
                        strCodigo As String, _
                        lngVARAID As Long, _
                        strPesoVara As String, _
                        strNomeProduto As String, _
                        strLargura As String, _
                        strAba As String, _
                        strEspessura As String)
  On Error GoTo trata
  '
  Dim strSql          As String
  '
  strSql = "INSERT INTO LINHA (" & _
    " TIPO_LINHAID, CODIGO, VARAID, PESO_VARA, NOME_PRODUTO, LARGURA, ABA, ESPESSURA) VALUES (" & _
    Formata_Dados(lngTIPOLINHAID, tpDados_Longo) & _
    ", " & Formata_Dados(strCodigo, tpDados_Texto) & _
    ", " & Formata_Dados(lngVARAID, tpDados_Longo) & _
    ", " & Formata_Dados(strPesoVara, tpDados_Moeda) & _
    ", " & Formata_Dados(strNomeProduto, tpDados_Texto) & _
    ", " & Formata_Dados(strLargura, tpDados_Moeda) & _
    ", " & Formata_Dados(strAba, tpDados_Moeda) & _
    ", " & Formata_Dados(strEspessura, tpDados_Moeda) & _
    ")"

  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, "[" & strClassName & ".clsLinhaPerfil.IncluirLinha]"
End Sub

Public Function CapturaItemLinha(ByVal strCodigo As String) As ADODB.Recordset
  '
  On Error GoTo trata
    
  Dim strSql  As String
  '
  strSql = "SELECT LINHA.*, TIPO_LINHA.NOME, VARA.VALOR FROM LINHA " & _
    " LEFT JOIN VARA ON VARA.PKID = LINHA.VARAID " & _
    " LEFT JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID " & _
    " WHERE " & _
    " LINHA.CODIGO LIKE " & Formata_Dados("%" & strCodigo & "%", tpDados_Texto) & _
    " OR TIPO_LINHA.NOME LIKE " & Formata_Dados("%" & strCodigo & "%", tpDados_Texto)
  Set CapturaItemLinha = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLinhaPerfil.CapturaItemLinha]", _
            Err.Description
End Function


