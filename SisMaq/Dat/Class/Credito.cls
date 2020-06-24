VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCredito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirCredito(lngCREDITOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM CREDITO " & _
    "WHERE CREDITO.PKID = " & _
    Formata_Dados(lngCREDITOID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCredito.ExcluirCredito]", _
            Err.Description
End Function


Public Function SelecionarCreditoPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT CREDITO.*, EQUIPAMENTO.NUMERO AS NUMERO_EQUIPAMENTO, BOLETOARREC.NUMERO AS NUMERO_BOLETO " & _
    "FROM CREDITO " & _
    " INNER JOIN MAQUINA ON MAQUINA.PKID = CREDITO.MAQUINAID " & _
    " INNER JOIN EQUIPAMENTO ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
    " INNER JOIN BOLETOARREC ON BOLETOARREC.PKID = CREDITO.BOLETOARRECID " & _
    "WHERE CREDITO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarCreditoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCredito.SelecionarCreditoPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarCredito(lngPKID As Long, _
                         strMedicao As String, _
                         strValorPago As String)

  On Error GoTo trata
  '
  Dim strSql As String
  'Inserir Credito
  strSql = "UPDATE CREDITO SET " & _
    " MEDICAO = " & Formata_Dados(strMedicao, tpDados_Moeda) & _
    ", VALORPAGO = " & Formata_Dados(strValorPago, tpDados_Moeda) & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCredito.AlterarCredito]", _
            Err.Description
End Sub

Public Sub InserirCredito(lngMAQUINAID As Long, _
                         lngBOLETOARRECID As Long, _
                         strNumero As String, _
                         strMedicao As String, _
                         strValorPago As String, _
                         strCoeficiente As String, _
                         strData As String, _
                         lngTURNOID As Long, _
                         strCredito As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = strSql & "INSERT INTO CREDITO (MAQUINAID, BOLETOARRECID, TURNOID, NUMERO, MEDICAO, VALORPAGO, COEFICIENTE, DATA, CREDITO) VALUES (" & _
    Formata_Dados(lngMAQUINAID, tpDados_Longo) & _
    ", " & Formata_Dados(lngBOLETOARRECID, tpDados_Longo) & _
    ", " & Formata_Dados(lngTURNOID, tpDados_Longo) & _
    ", " & Formata_Dados(strNumero, tpDados_Longo) & _
    ", " & Formata_Dados(strMedicao, tpDados_Moeda) & _
    ", " & Formata_Dados(strValorPago, tpDados_Moeda) & _
    ", " & Formata_Dados(strCoeficiente, tpDados_Moeda) & _
    ", " & Formata_Dados(strData, tpDados_DataHora) & _
    ", " & Formata_Dados(strCredito, tpDados_Moeda) & _
    ");"
    
  '
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsCredito.InserirCredito]", _
            Err.Description
End Sub





