VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLinhaAssoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub InserirLinhaAssoc(lngLINHAID As Long, _
                             lngTITULARASSOCIADOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TAB_TITLINHA (TITULARASSOCIADOID, LINHAID) VALUES (" & _
    Formata_Dados(lngTITULARASSOCIADOID, tpDados_Longo) & ", " & _
    Formata_Dados(lngLINHAID, tpDados_Longo) & ");"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLinhaAssoc.InserirLinhaAssoc]", _
            Err.Description
End Sub


Public Sub ExcluirLinhaAssoc(lngLINHAID As Long, _
                             lngTITULARASSOCIADOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "DELETE FROM TAB_TITLINHA WHERE LINHAID = " & _
    Formata_Dados(lngLINHAID, tpDados_Longo) & _
    " AND TITULARASSOCIADOID = " & _
    Formata_Dados(lngTITULARASSOCIADOID, tpDados_Longo) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsLinhaAssoc.ExcluirLinhaAssoc]", _
            Err.Description
End Sub

