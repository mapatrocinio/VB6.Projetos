VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsBM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function ExcluirBM(lngBMID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "DELETE FROM BM " & _
    "WHERE BM.PKID = " & _
    Formata_Dados(lngBMID, tpDados_Longo)
  RunSP strSql
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBM.ExcluirBM]", _
            Err.Description
End Function

Public Sub AlterarBM(lngPKID As Long, _
                     lngCONTRATOID As Long, _
                     strMedicao As String, _
                     strDataInicio As String, _
                     strDataTermino As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE BM SET " & _
    "CONTRATOID = " & Formata_Dados(lngCONTRATOID, tpDados_Longo) & _
    ", MEDICAO = " & Formata_Dados(strMedicao, tpDados_Texto) & _
    ", DATAINICIO = " & Formata_Dados(strDataInicio, tpDados_DataHora) & _
    ", DATATERMINO = " & Formata_Dados(strDataTermino, tpDados_DataHora) & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBM.AlterarBM]", _
            Err.Description
End Sub

Public Sub InserirBM(lngCONTRATOID As Long, _
                     strNumero As String, _
                     strMedicao As String, _
                     strDataEmissao As String, _
                     strDataInicio As String, _
                     strDataTermino As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO BM (CONTRATOID, NUMERO, MEDICAO, DATAEMISSAO, DATAINICIO, DATATERMINO) VALUES (" & _
    Formata_Dados(lngCONTRATOID, tpDados_Longo) & _
    ", " & Formata_Dados(strNumero, tpDados_Longo) & _
    ", " & Formata_Dados(strMedicao, tpDados_Texto) & _
    ", " & Formata_Dados(strDataEmissao, tpDados_DataHora) & _
    ", " & Formata_Dados(strDataInicio, tpDados_DataHora) & _
    ", " & Formata_Dados(strDataTermino, tpDados_DataHora) & _
    ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBM.InserirBM]", _
            Err.Description
End Sub

Public Function SelecionarBM(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT EMPRESA.NOME AS NOME_EMPRESA, CONTRATO.NUMERO AS NUMERO_CONTRATO, BM.* " & _
    "FROM BM " & _
    " INNER JOIN CONTRATO ON CONTRATO.PKID = BM.CONTRATOID " & _
    " INNER JOIN EMPRESA ON EMPRESA.PKID = CONTRATO.EMPRESAID " & _
    "WHERE BM.PKID = " & _
    Formata_Dados(lngPKID, tpDados_Longo)
  Set SelecionarBM = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsBM.SelecionarBM]", _
            Err.Description
End Function

