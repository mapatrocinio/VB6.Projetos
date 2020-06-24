VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGRPgto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirGRPGTOGR(lngGRID As Long, _
                           lngGRPAGAMENTOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO GRPGTO (GRID, GRPAGAMENTOID) VALUES (" & _
    Formata_Dados(lngGRID, tpDados_Longo) & ", " & _
    Formata_Dados(lngGRPAGAMENTOID, tpDados_Longo) & ");"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGRPgto.InserirGRPGTOGR]", _
            Err.Description
End Sub

Public Sub ExlcuirGRPGTOGR(lngGRPGTOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "DELETE FROM GRPGTO WHERE PKID = " & _
    Formata_Dados(lngGRPGTOID, tpDados_Longo) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsGRPgto.ExlcuirGRPGTOGR]", _
            Err.Description
End Sub



