VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrestProcedimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirPrestProcedimento(lngPrestProcedimentoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM PRESTADORPROCEDIMENTO " & _
    "WHERE PRESTADORPROCEDIMENTO.PKID = " & _
    Formata_Dados(lngPrestProcedimentoId, tpDados_Longo, tpNulo_Aceita)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPrestProcedimento.ExcluirPrestProcedimento]", _
            Err.Description
End Function

Public Function SelecionarPrestProcedimentoPeloPkid(lngPrestProcedimentoId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PRESTADORPROCEDIMENTO.*, PROCEDIMENTO.PROCEDIMENTO " & _
    "FROM PRESTADORPROCEDIMENTO INNER JOIN PROCEDIMENTO ON PROCEDIMENTO.PKID = PRESTADORPROCEDIMENTO.PROCEDIMENTOID " & _
    "WHERE PRESTADORPROCEDIMENTO.PKID = " & _
    Formata_Dados(lngPrestProcedimentoId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarPrestProcedimentoPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPrestProcedimento.SelecionarPrestProcedimentoPeloPkid]", _
            Err.Description
End Function


Public Sub AlterarPrestProcedimento(lngPKID As Long, _
                                    lngPROCEDIMENTOID As Long, _
                                    strPercCasa As String, _
                                    strPercPrestador As String, _
                                    strPercRX As String, _
                                    strPercTecRX As String, _
                                    strPercDonoUltra As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PRESTADORPROCEDIMENTO SET " & _
    "PROCEDIMENTOID = " & Formata_Dados(lngPROCEDIMENTOID, tpDados_Longo) & ", " & _
    "PERCCASA = " & Formata_Dados(strPercCasa, tpDados_Moeda) & ", " & _
    "PERCPRESTADOR = " & Formata_Dados(strPercPrestador, tpDados_Moeda) & ", " & _
    "PERCRX = " & Formata_Dados(strPercRX, tpDados_Moeda) & ", " & _
    "PERCTECRX = " & Formata_Dados(strPercTecRX, tpDados_Moeda) & ", " & _
    "PERCULTRA = " & Formata_Dados(strPercDonoUltra, tpDados_Moeda) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPrestProcedimento.AlterarPrestProcedimento]", _
            Err.Description
End Sub

Public Sub InserirPrestProcedimento(lngPRESTADORID As Long, _
                                    lngPROCEDIMENTOID As Long, _
                                    strPercCasa As String, _
                                    strPercPrestador As String, _
                                    strPercRX As String, _
                                    strPercTecRX As String, _
                                    strPercDonoUltra As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO PRESTADORPROCEDIMENTO (PRONTUARIOID, PROCEDIMENTOID, PERCCASA, PERCPRESTADOR, PERCRX, PERCTECRX, PERCULTRA) VALUES (" & _
    Formata_Dados(lngPRESTADORID, tpDados_Longo) & ", " & _
    Formata_Dados(lngPROCEDIMENTOID, tpDados_Longo) & ", " & _
    Formata_Dados(strPercCasa, tpDados_Moeda) & ", " & _
    Formata_Dados(strPercPrestador, tpDados_Moeda) & ", " & _
    Formata_Dados(strPercRX, tpDados_Moeda) & ", " & _
    Formata_Dados(strPercTecRX, tpDados_Moeda) & ", " & _
    Formata_Dados(strPercDonoUltra, tpDados_Moeda) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPrestProcedimento.InserirPrestProcedimento]", _
            Err.Description
End Sub

Public Function SelecionarPrestProcedimento(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT PRESTADORPROCEDIMENTO.* " & _
    "FROM PRESTADORPROCEDIMENTO " & _
    "WHERE PRESTADORPROCEDIMENTO.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarPrestProcedimento = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPrestProcedimento.SelecionarPrestProcedimento]", _
            Err.Description
End Function

