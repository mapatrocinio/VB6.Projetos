VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBoletoAtend"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirBoletoAtend(lngBOLETOATENDID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM BOLETOATEND " & _
    "WHERE BOLETOATEND.PKID = " & _
    Formata_Dados(lngBOLETOATENDID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBoletoAtend.ExcluirBoletoAtend]", _
            Err.Description
End Function


Public Function SelecionarBoletoAtendPeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT BOLETOATEND.*, PESSOA.NOME AS DESC_ATENDENTE " & _
    "FROM BOLETOATEND " & _
    " INNER JOIN CAIXAATEND ON CAIXAATEND.PKID = BOLETOATEND.CAIXAATENDID " & _
    " INNER JOIN PESSOA ON PESSOA.PKID = CAIXAATEND.ATENDENTEID " & _
    "WHERE BOLETOATEND.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarBoletoAtendPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBoletoAtend.SelecionarBoletoAtendPeloPkid]", _
            Err.Description
End Function

Public Sub AlterarStatusBoletoAtend(lngPKID As Long, _
                                    strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  'Inserir StatusBoletoAtend
  strSql = "UPDATE BOLETOATEND SET " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsStatusBoletoAtend.AlterarStatusBoletoAtend]", _
            Err.Description
End Sub

Public Sub AlterarBoletoAtend(lngPKID As Long, _
                              strNumero As String)

  On Error GoTo trata
  '
  Dim strSql As String
  'Inserir BoletoAtend
  strSql = "UPDATE BOLETOATEND SET " & _
    "NUMERO = " & Formata_Dados(strNumero, tpDados_Moeda) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBoletoAtend.AlterarBoletoAtend]", _
            Err.Description
End Sub

Public Sub InserirBoletoAtend(lngTURNOID As Long, _
                              lngCAIXAATENDID As Long, _
                              strNumero As String, _
                              strData As String, _
                              strStatus As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = strSql & "INSERT INTO BOLETOATEND (TURNOENTRADAID, CAIXAATENDID, NUMERO, DATAENTRADA, STATUS) VALUES (" & _
    Formata_Dados(lngTURNOID, tpDados_Longo) & ", " & _
    Formata_Dados(lngCAIXAATENDID, tpDados_Longo) & ", " & _
    Formata_Dados(strNumero, tpDados_Longo) & ", " & _
    Formata_Dados(strData, tpDados_DataHora) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ");"
    
  '
  RunSP strSql
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBoletoAtend.InserirBoletoAtend]", _
            Err.Description
End Sub





