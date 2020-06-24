VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function SelecionarNFPeloPkid(lngNFID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  strSql = "SELECT NF.*, EMPRESA.NOME AS NOME_EMPRESA, CONTRATO.NUMERO AS CONTRATO_NUMERO, OBRA.DESCRICAO AS DESC_OBRA " & _
    "FROM NF INNER JOIN CONTRATO ON CONTRATO.PKID = NF.CONTRATOID " & _
    " INNER JOIN EMPRESA ON EMPRESA.PKID = CONTRATO.EMPRESAID " & _
    " INNER JOIN OBRA ON OBRA.PKID = NF.OBRAID " & _
    "WHERE NF.PKID = " & _
    Formata_Dados(lngNFID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarNFPeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsNF.SelecionarNFPeloPkid]", _
            Err.Description
End Function

Public Sub AlterarNF(lngPKID As Long, _
                     lngCONTRATOID As Long, _
                     strDataSaida As String, _
                     strDataIniCob As String, _
                     strNroRF As String, _
                     lngOBRAID As Long)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE NF SET " & _
    "CONTRATOID = " & Formata_Dados(lngCONTRATOID, tpDados_Longo) & ", " & _
    "DTSAIDA = " & Formata_Dados(strDataSaida, tpDados_DataHora) & ", " & _
    "DTINICIOCOB = " & Formata_Dados(strDataIniCob, tpDados_DataHora) & ", " & _
    "NUMERORF = " & Formata_Dados(strNroRF, tpDados_Texto) & ", " & _
    "OBRAID = " & Formata_Dados(lngOBRAID, tpDados_Longo) & " " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsNF.AlterarNF]", _
            Err.Description
End Sub

Public Sub InserirNF(ByRef lngPKID As Long, _
                     lngCONTRATOID As Long, _
                     strNumero As String, _
                     strSequencial As String, _
                     strAno As String, _
                     strData As String, _
                     strDataSaida As String, _
                     strDataIniCob As String, _
                     strNroRF As String, _
                     lngOBRAID As Long)
  On Error GoTo trata
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO NF (CONTRATOID, NUMERO, SEQUENCIAL, ANO, DATA, DTSAIDA, DTINICIOCOB, NUMERORF, OBRAID) VALUES (" & _
    " " & Formata_Dados(lngCONTRATOID, tpDados_Longo) & _
    ", " & Formata_Dados(strNumero, tpDados_Texto) & _
    ", " & Formata_Dados(strSequencial, tpDados_Longo) & _
    ", " & Formata_Dados(strAno, tpDados_Longo) & _
    ", " & Formata_Dados(strData, tpDados_DataHora) & _
    ", " & Formata_Dados(strDataSaida, tpDados_DataHora) & _
    ", " & Formata_Dados(strDataIniCob, tpDados_DataHora) & _
    ", " & Formata_Dados(strNroRF, tpDados_Texto) & _
    ", " & Formata_Dados(lngOBRAID, tpDados_Longo) & ");"
    
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPKID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsNF.InserirNF]", _
            Err.Description
End Sub

Public Sub AlterarStatusNF(lngNFID As Long, _
                           strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE NF SET "
  
  If Len(strStatus & "") > 0 Then
    strSql = strSql & " STATUS = " & Formata_Dados(strStatus, tpDados_Texto) & " "
  End If
  strSql = strSql & " WHERE PKID = " & Formata_Dados(lngNFID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsNF.AlterarStatusNF]", _
            Err.Description
End Sub
