VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsParcela"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarParcela(lngParcelaId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PARCELA.* " & _
    " FROM PARCELA " & _
    " WHERE PARCELA.PKID = " & _
    Formata_Dados(lngParcelaId, tpDados_Longo, tpNulo_NaoAceita)
  
  Set SelecionarParcela = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsParcela.SelecionarParcela]", _
            Err.Description
End Function

Public Sub ExcluirParcela(lngParcelaId As Long)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "DELETE FROM PARCELA WHERE PKID = " & _
    Formata_Dados(lngParcelaId, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsParcela.ExcluirParcela]", _
            Err.Description
End Sub

Public Sub InserirParcela(lngContaCorrenteId As Long, _
                          strParcela As Long, _
                          strDtVencimento As String, _
                          strDtPagamento As String, _
                          strVrParcela As String, _
                          strVrPago As String)

  On Error GoTo trata
  '
  Dim strSql As String
      
  strSql = "INSERT INTO PARCELA " & _
    "(CONTACORRENTEID, PARCELA, DTVENCIMENTO, DTPAGAMENTO, VRPARCELA, VRPAGO) VALUES (" & _
    Formata_Dados(lngContaCorrenteId, tpDados_Longo) & _
    ", " & Formata_Dados(strParcela, tpDados_Longo) & _
    ", " & Formata_Dados(strDtVencimento, tpDados_DataHora) & _
    ", " & Formata_Dados(strDtPagamento, tpDados_DataHora) & _
    ", " & Formata_Dados(strVrParcela, tpDados_Moeda) & _
    ", " & Formata_Dados(strVrPago, tpDados_Moeda) & _
    ");"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsParcela.InserirParcela]", _
            Err.Description
End Sub


Public Sub AlterarParcela(lngParcelaId As Long, _
                          strDtVencimento As String, _
                          strDtPagamento As String, _
                          strVrParcela As String, _
                          strVrPago As String)

  On Error GoTo trata
  '
  Dim strSql As String
      
  strSql = "UPDATE PARCELA SET " & _
    "DTVENCIMENTO = " & Formata_Dados(strDtVencimento, tpDados_DataHora) & _
    ", DTPAGAMENTO = " & Formata_Dados(strDtPagamento, tpDados_DataHora) & _
    ", VRPARCELA = " & Formata_Dados(strVrParcela, tpDados_Moeda) & _
    ", VRPAGO = " & Formata_Dados(strVrPago, tpDados_Moeda) & _
    " WHERE PKID = " & Formata_Dados(lngParcelaId, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsParcela.AlterarParcela]", _
            Err.Description
End Sub


Public Sub ExcluirParcelasDaCC(lngCCId As Long)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "DELETE FROM PARCELA WHERE CONTACORRENTEID = " & _
    Formata_Dados(lngCCId, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsParcela.ExcluirParcelasDaCC]", _
            Err.Description
End Sub


