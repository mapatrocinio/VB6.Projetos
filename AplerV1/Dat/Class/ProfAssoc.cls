VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProfAssoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub InserirProfAssoc(lngPROFISSAOID As Long, _
                            lngASSOCIADOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TAB_PROFASSOC (ASSOCIADOID, PROFISSAOID) VALUES (" & _
    Formata_Dados(lngASSOCIADOID, tpDados_Longo) & ", " & _
    Formata_Dados(lngPROFISSAOID, tpDados_Longo) & ");"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProfAssoc.InserirProfAssoc]", _
            Err.Description
End Sub


Public Sub ExcluirProfAssoc(lngPROFISSAOID As Long, _
                            lngASSOCIADOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "DELETE FROM TAB_PROFASSOC WHERE PROFISSAOID = " & _
    Formata_Dados(lngPROFISSAOID, tpDados_Longo) & _
    " AND ASSOCIADOID = " & _
    Formata_Dados(lngASSOCIADOID, tpDados_Longo) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsProfAssoc.ExcluirProfAssoc]", _
            Err.Description
End Sub

