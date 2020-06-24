VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSerie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirSerie(lngSERIEID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM SERIE " & _
    "WHERE SERIE.PKID = " & _
    Formata_Dados(lngSERIEID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSerie.ExcluirSerie]", _
            Err.Description
End Function


Public Function SelecionarSeriePeloPkid(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT SERIE.*, PESSOA.NOME AS DESC_DONO " & _
    "FROM SERIE LEFT JOIN PESSOA ON PESSOA.PKID = SERIE.DONOID " & _
    "WHERE SERIE.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarSeriePeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSerie.SelecionarSeriePeloPkid]", _
            Err.Description
End Function

Public Function SelecionarSeriePeloNumero(strNumero As String) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT SERIE.*, PESSOA.NOME AS NOME_PESSOA " & _
    "FROM SERIE LEFT JOIN PESSOA ON PESSOA.PKID = SERIE.DONOID " & _
    "WHERE SERIE.NUMERO = " & _
    Formata_Dados(strNumero, tpDados_Texto)
  Set SelecionarSeriePeloNumero = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSerie.SelecionarSeriePeloNumero]", _
            Err.Description
End Function


Public Sub AlterarSerie(lngPKID As Long, _
                        lngDONOID As Long, _
                        strNumero As String, _
                        strPercDono As String, _
                        strPercCasa As String, _
                        strCoeficiente As String, _
                        strStatus As String, _
                        strUsuario As String)

  On Error GoTo trata
  '
  Dim strSql As String
  'Inserir histórico
  strSql = "INSERT INTO SERIEHIST (SERIEID, DATA, NUMERO, PERCDONO, PERCCASA, COEFICIENTE, STATUS, USUARIO) " & _
    " SELECT PKID, GETDATE(), NUMERO, PERCDONO, PERCCASA, COEFICIENTE, STATUS," & Formata_Dados(strUsuario, tpDados_Texto) & " FROM SERIE " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  'Inserir série
  strSql = "UPDATE SERIE SET " & _
    "DONOID = " & Formata_Dados(lngDONOID, tpDados_Longo) & ", " & _
    "NUMERO = " & Formata_Dados(strNumero, tpDados_Texto) & ", " & _
    "PERCDONO = " & Formata_Dados(strPercDono, tpDados_Moeda) & ", " & _
    "PERCCASA = " & Formata_Dados(strPercCasa, tpDados_Moeda) & ", " & _
    "COEFICIENTE = " & Formata_Dados(strCoeficiente, tpDados_Moeda) & ", " & _
    "STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSerie.AlterarSerie]", _
            Err.Description
End Sub

Public Sub InserirSerie(ByRef lngPKID As Long, _
                        lngDONOID As Long, _
                        strNumero As String, _
                        strPercDono As String, _
                        strPercCasa As String, _
                        strCoeficiente As String, _
                        strStatus As String)

  On Error GoTo trata
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO SERIE (DONOID, NUMERO, PERCDONO, PERCCASA, COEFICIENTE, STATUS) VALUES (" & _
    Formata_Dados(lngDONOID, tpDados_Longo) & ", " & _
    Formata_Dados(strNumero, tpDados_Texto) & ", " & _
    Formata_Dados(strPercDono, tpDados_Moeda) & ", " & _
    Formata_Dados(strPercCasa, tpDados_Moeda) & ", " & _
    Formata_Dados(strCoeficiente, tpDados_Moeda) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ");"
  '
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPKID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsSerie.InserirSerie]", _
            Err.Description
End Sub





