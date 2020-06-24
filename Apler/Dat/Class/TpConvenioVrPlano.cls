VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTpConvenioVrPlano"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Public Sub InserirTpConvenioVrPlano(lngVALORPLANOID As Long, _
                                    lngTIPOCONVENIOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO TAB_TPCONVENIOVRPLANO (TIPOCONVENIOID, VALORPLANOID) VALUES (" & _
    Formata_Dados(lngTIPOCONVENIOID, tpDados_Longo) & ", " & _
    Formata_Dados(lngVALORPLANOID, tpDados_Longo) & ");"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTpConvenioVrPlano.InserirTpConvenioVrPlano]", _
            Err.Description
End Sub


Public Sub ExcluirTpConvenioVrPlano(lngVALORPLANOID As Long, _
                                    lngTIPOCONVENIOID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "DELETE FROM TAB_TPCONVENIOVRPLANO WHERE VALORPLANOID = " & _
    Formata_Dados(lngVALORPLANOID, tpDados_Longo) & _
    " AND TIPOCONVENIOID = " & _
    Formata_Dados(lngTIPOCONVENIOID, tpDados_Longo) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsTpConvenioVrPlano.ExcluirTpConvenioVrPlano]", _
            Err.Description
End Sub

