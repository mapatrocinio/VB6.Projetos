VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBoletoArrec"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirBoletoArrec(lngBOLETOARRECID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM BOLETOARREC " & _
    "WHERE BOLETOARREC.PKID = " & _
    Formata_Dados(lngBOLETOARRECID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBoletoArrec.ExcluirBoletoArrec]", _
            Err.Description
End Function


Public Function SelecionarBoletoArrecPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT BOLETOARREC.*, PESSOA.NOME AS DESC_ARRECADADOR " & _
    "FROM BOLETOARREC " & _
    " INNER JOIN CAIXAARREC ON CAIXAARREC.PKID = BOLETOARREC.CAIXAARRECID " & _
    " INNER JOIN PESSOA ON PESSOA.PKID = CAIXAARREC.ARRECADADORID " & _
    "WHERE BOLETOARREC.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarBoletoArrecPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBoletoArrec.SelecionarBoletoArrecPeloPkid]", _
            Err.Description
End Function

Public Sub AlterarStatusBoletoArrec(lngPKID As Long, _
                                    strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  'Inserir StatusBoletoArrec
  strSql = "UPDATE BOLETOARREC SET " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsStatusBoletoArrec.AlterarStatusBoletoArrec]", _
            Err.Description
End Sub

Public Sub AlterarBoletoArrec(lngPKID As Long, _
                              strNumero As String)

  On Error GoTo trata
  '
  Dim strSql As String
  'Inserir BoletoArrec
  strSql = "UPDATE BOLETOARREC SET " & _
    "NUMERO = " & Formata_Dados(strNumero, tpDados_Moeda) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBoletoArrec.AlterarBoletoArrec]", _
            Err.Description
End Sub

Public Sub InserirBoletoArrec(lngTURNOID As Long, _
                              lngCAIXAARRECID As Long, _
                              strNumero As String, _
                              strData As String, _
                              strStatus As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = strSql & "INSERT INTO BOLETOARREC (TURNOENTRADAID, CAIXAARRECID, NUMERO, DATAENTRADA, STATUS) VALUES (" & _
    Formata_Dados(lngTURNOID, tpDados_Longo) & ", " & _
    Formata_Dados(lngCAIXAARRECID, tpDados_Longo) & ", " & _
    Formata_Dados(strNumero, tpDados_Longo) & ", " & _
    Formata_Dados(strData, tpDados_DataHora) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ");"
    
  '
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBoletoArrec.InserirBoletoArrec]", _
            Err.Description
End Sub





