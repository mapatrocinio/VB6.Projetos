VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsItemOS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirItemOS(lngITEMOSID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM ITEM_OS " & _
    "WHERE ITEM_OS.PKID = " & _
    Formata_Dados(lngITEMOSID, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemOS.ExcluirItemOS]", _
            Err.Description
End Function

Public Function SelecionarItemOSPeloPkid(lngITEMOSID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ITEM_OS.*, TIPO_LINHA.NOME AS NOME_LINHA, LINHA.CODIGO AS CODIGO_LINHA " & _
    "FROM ITEM_OS " & _
    " INNER JOIN LINHA ON LINHA.PKID = ITEM_OS.LINHAID " & _
    " INNER JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID " & _
    "WHERE ITEM_OS.PKID = " & _
    Formata_Dados(lngITEMOSID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarItemOSPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemOS.SelecionarItemOSPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarItemOS(lngITEMOSID As Long, _
                         lngOSID As Long, _
                         lngLINHAID As Long, _
                         strQuantidade As String)

  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "EXEC SP_INSERE_ITEM_OS " & Formata_Dados(lngITEMOSID, tpDados_Longo) & _
      "," & Formata_Dados(lngOSID, tpDados_Longo) & _
      "," & Formata_Dados(lngLINHAID, tpDados_Longo) & _
      "," & Formata_Dados(strQuantidade, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemOS.AlterarItemOS]", _
            Err.Description
End Sub

Public Sub InserirItemOS(lngOSID As Long, _
                         lngLINHAID As Long, _
                         strQuantidade As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "EXEC SP_INSERE_ITEM_OS " & Formata_Dados("", tpDados_Longo) & _
      "," & Formata_Dados(lngOSID, tpDados_Longo) & _
      "," & Formata_Dados(lngLINHAID, tpDados_Longo) & _
      "," & Formata_Dados(strQuantidade, tpDados_Longo)
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemOS.InserirItemOS]", _
            Err.Description
End Sub

Public Function SelecionarItemOS(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT ITEM_OS.* " & _
    "FROM ITEM_OS " & _
    "WHERE ITEM_OS.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarItemOS = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsItemOS.SelecionarItemOS]", _
            Err.Description
End Function

