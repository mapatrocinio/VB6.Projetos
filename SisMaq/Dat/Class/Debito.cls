VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsdebito"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirDebito(lngDEBITOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM DEBITO " & _
    "WHERE DEBITO.PKID = " & _
    Formata_Dados(lngDEBITOID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDebito.ExcluirDebito]", _
            Err.Description
End Function


Public Function SelecionarDebitoPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT DEBITO.*, EQUIPAMENTO.NUMERO AS NUMERO_EQUIPAMENTO, BOLETOATEND.NUMERO AS NUMERO_BOLETO " & _
    "FROM DEBITO " & _
    " INNER JOIN MAQUINA ON MAQUINA.PKID = DEBITO.MAQUINAID " & _
    " INNER JOIN EQUIPAMENTO ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
    " INNER JOIN BOLETOATEND ON BOLETOATEND.PKID = DEBITO.BOLETOATENDID " & _
    "WHERE DEBITO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarDebitoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDebito.SelecionarDebitoPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarDebito(lngPKID As Long, _
                         strMedicao As String, _
                         strValorPago As String)

  On Error GoTo trata
  '
  Dim strSql As String
  'Inserir Debito
  strSql = "UPDATE DEBITO SET " & _
    "MEDICAO = " & Formata_Dados(strMedicao, tpDados_Moeda) & _
    ", VALORPAGO = " & Formata_Dados(strValorPago, tpDados_Moeda) & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDebito.AlterarDebito]", _
            Err.Description
End Sub

Public Sub InserirDebito(lngMAQUINAID As Long, _
                         lngBOLETOATENDID As Long, _
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
  strSql = strSql & "INSERT INTO DEBITO (MAQUINAID, BOLETOATENDID, TURNOID, NUMERO, MEDICAO, VALORPAGO, COEFICIENTE, DATA, CREDITO) VALUES (" & _
    Formata_Dados(lngMAQUINAID, tpDados_Longo) & _
    ", " & Formata_Dados(lngBOLETOATENDID, tpDados_Longo) & _
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
            Err.Source & ".[clsDebito.InserirDebito]", _
            Err.Description
End Sub





