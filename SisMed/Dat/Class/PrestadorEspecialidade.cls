VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrestadorEspecialidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirPrestadorEspecialidade(lngESPECIALIDADEID As Long, _
                                         lngPRESTADORID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "INSERT INTO PRESTADORESPECIALIDADE (PRONTUARIOID, ESPECIALIDADEID) VALUES (" & _
    Formata_Dados(lngPRESTADORID, tpDados_Longo, tpNulo_NaoAceita) & ", " & _
    Formata_Dados(lngESPECIALIDADEID, tpDados_Longo, tpNulo_NaoAceita) & ");"
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPrestadorEspecialidade.InserirPrestadorEspecialidade]", _
            Err.Description
End Sub

Public Sub ExlcuirPrestadorEspecialidade(lngESPECIALIDADEID As Long, _
                                         lngPRESTADORID As Long)
  On Error GoTo trata
  '
  Dim strSql As String
  '
  strSql = "DELETE FROM PRESTADORESPECIALIDADE WHERE ESPECIALIDADEID = " & _
    Formata_Dados(lngESPECIALIDADEID, tpDados_Longo, tpNulo_NaoAceita) & _
    " AND PRONTUARIOID = " & _
    Formata_Dados(lngPRESTADORID, tpDados_Longo, tpNulo_NaoAceita) & ";"
  
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPrestadorEspecialidade.ExlcuirPrestadorEspecialidade]", _
            Err.Description
End Sub


