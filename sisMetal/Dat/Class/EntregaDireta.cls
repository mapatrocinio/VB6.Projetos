VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEntregaDireta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirEntregaDireta(lngENTREGADIRETAID As Long, _
                                strData As String)
  On Error GoTo trata
  
  '
  Dim strSql As String

  strSql = "DELETE ENTREGA_DIRETA " & _
    " WHERE PKID = " & _
    Formata_Dados(lngENTREGADIRETAID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntregaDireta.ExcluirEntregaDireta]", _
            Err.Description
End Sub

Public Function ListarEntregaDireta(lngENTREGADIRETAID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT ENTREGA_DIRETA.*, FORNECEDOR.NOME AS NOME_FORNECEDOR " & _
    " FROM ENTREGA_DIRETA " & _
    " LEFT JOIN LOJA AS FORNECEDOR ON ENTREGA_DIRETA.FORNECEDORID = FORNECEDOR.PKID " & _
    " WHERE ENTREGA_DIRETA.PKID = " & _
    Formata_Dados(lngENTREGADIRETAID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarEntregaDireta = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntregaDireta.ListarEntregaDireta]", _
            Err.Description
End Function


Public Sub InserirEntregaDireta(lngENTREGADIRETAID As Long, _
                                lngFORNECEDORID As Long, _
                                strAno As String, _
                                strNumroOS As String, _
                                strData As String)
  On Error GoTo trata
  '
  Dim objRs As ADODB.Recordset
  Dim strSql As String
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO ENTREGA_DIRETA (DATA, OS_NUMERO, OS_ANO, FORNECEDORID) VALUES (" & _
    Formata_Dados(strData, tpDados_DataHora) & ", " & _
    Formata_Dados(strNumroOS, tpDados_Longo) & ", " & _
    Formata_Dados(strAno, tpDados_Longo) & ", " & _
    Formata_Dados(lngFORNECEDORID, tpDados_Longo) & "); "
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngENTREGADIRETAID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntregaDireta.InserirEntregaDireta]", _
            Err.Description
End Sub

Public Sub AlterarEntregaDireta(lngENTREGADIRETAID As Long, _
                         lngFORNECEDORID As Long)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE ENTREGA_DIRETA SET " & _
    "FORNECEDORID = " & Formata_Dados(lngFORNECEDORID, tpDados_Longo) & _
    " WHERE ENTREGA_DIRETA.PKID = " & Formata_Dados(lngENTREGADIRETAID, tpDados_Longo)
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntregaDireta.AlterarEntregaDireta]", _
            Err.Description
End Sub

Public Function CapturarSequencialEntregaDireta(strAno As String, _
                                                lngFORNECEDORID As Long, _
                                                strNumeroOS As String) As Long
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SELECT PKID, SEQUENCIAL FROM SEQUENCIAL_ENTREGA_DIRETA " & _
      " WHERE ANO = " & Formata_Dados(strAno, tpDados_Longo) & _
      " AND FORNECEDORID = " & Formata_Dados(lngFORNECEDORID, tpDados_Longo)
  Set objRs = RunSPReturnRS(strSql)
  If objRs.EOF Then
    strNumeroOS = "1"
    '----------- INSERIR SEQUENCIAL
    strSql = "INSERT INTO SEQUENCIAL_ENTREGA_DIRETA (" & _
      " ANO, FORNECEDORID, SEQUENCIAL) VALUES (" & _
      Formata_Dados(strAno, tpDados_Longo) & _
      ", " & Formata_Dados(lngFORNECEDORID, tpDados_Longo) & _
      ", " & Formata_Dados(1, tpDados_Longo) & _
      ")"
    '
    RunSP strSql
    '
  Else
    strNumeroOS = objRs.Fields("SEQUENCIAL").Value + 1
    '----------- ATUALIZA SEQUENCIAL
    strSql = "UPDATE SEQUENCIAL_ENTREGA_DIRETA SET " & _
      "SEQUENCIAL = " & Formata_Dados(strNumeroOS, tpDados_Longo) & _
      " WHERE ANO = " & Formata_Dados(strAno, tpDados_Longo) & " " & _
      " AND FORNECEDORID = " & Formata_Dados(lngFORNECEDORID, tpDados_Longo)
    RunSP strSql
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsEntregaDireta.CapturarSequencialEntregaDireta]", _
            Err.Description
End Function
