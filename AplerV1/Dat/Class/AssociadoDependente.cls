VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAssociadoDependente"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarDependentePeloPkid(lngDependenteId As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "SELECT DEPENDENTE.*, GRAUPARENTESCO.DESCRICAO AS DESCR_GRAUPARENTESCO " & _
    "FROM DEPENDENTE LEFT JOIN GRAUPARENTESCO ON DEPENDENTE.GRAUPARENTESCOID = GRAUPARENTESCO.PKID " & _
    "WHERE DEPENDENTE.ASSOCIADOID = " & _
    Formata_Dados(lngDependenteId, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarDependentePeloPkid = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDependente.SelecionarDependentePeloPkid]", _
            Err.Description
End Function

Public Sub AlterarDependente(lngPKID As Long, _
                             lngGRAUPARENTESCOID As Long, _
                             strMatricula As String)


  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE DEPENDENTE SET " & _
    "GRAUPARENTESCOID = " & Formata_Dados(IIf(lngGRAUPARENTESCOID = 0, "", lngGRAUPARENTESCOID), tpDados_Longo) & ", " & _
    "MATRICULA = " & Formata_Dados(strMatricula, tpDados_Texto) & ", " & _
    "WHERE PKID = " & Formata_Dados(lngPKID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDependente.AlterarDependente]", _
            Err.Description
End Sub

Public Sub InserirDependente(lngPKID As Long, _
                             lngGRAUPARENTESCOID As Long, _
                             lngTITULARASSOCIADOID As Long, _
                             strMatricula As String)

  On Error GoTo trata
  Dim strSql As String
  '
  strSql = "INSERT INTO DEPENDENTE (ASSOCIADOID, GRAUPARENTESCOID, TITULARASSOCIADOID, MATRICULADEP) VALUES (" & _
    Formata_Dados(lngPKID, tpDados_Longo) & ", " & _
    Formata_Dados(IIf(lngGRAUPARENTESCOID = 0, "", lngGRAUPARENTESCOID), tpDados_Longo) & ", " & _
    Formata_Dados(IIf(lngTITULARASSOCIADOID = 0, "", lngTITULARASSOCIADOID), tpDados_Longo) & ", " & _
    Formata_Dados(strMatricula, tpDados_Texto) & ") "
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDependente.InserirDependente]", _
            Err.Description
End Sub

Public Function SelecionarDependente(lngPKID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String
  strSql = "SELECT DEPENDENTE.*, GRAUPARENTESCO.DESCRICAO AS DESCR_GRAUPARENTESCO " & _
    "FROM DEPENDENTE LEFT JOIN GRAUPARENTESCO ON DEPENDENTE.GRAUPARENTESCOID = GRAUPARENTESCO.PKID " & _
    "WHERE DEPENDENTE.ASSOCIADOID = " & _
    Formata_Dados(lngPKID, tpDados_Longo, tpNulo_Aceita)
  Set SelecionarDependente = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsDependente.SelecionarDependente]", _
            Err.Description
End Function

