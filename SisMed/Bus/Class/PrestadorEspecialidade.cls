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


Public Sub AssociarPrestadorAEspecialidade(lngESPECIALIDADEID As Long, _
                                           lngPRESTADORID As Long)
  On Error GoTo trata
  '
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objPrestadorEspecialidade As datSisMed.clsPrestadorEspecialidade
  '
  Set objPrestadorEspecialidade = New datSisMed.clsPrestadorEspecialidade
  '
  objPrestadorEspecialidade.InserirPrestadorEspecialidade lngESPECIALIDADEID, _
                                                          lngPRESTADORID
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub DesassociarPrestadorAEspecialidade(lngESPECIALIDADEID As Long, _
                                              lngPRESTADORID As Long)
  On Error GoTo trata
  '
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objPrestadorEspecialidade As datSisMed.clsPrestadorEspecialidade
  '
  Set objPrestadorEspecialidade = New datSisMed.clsPrestadorEspecialidade
  '
  objPrestadorEspecialidade.ExlcuirPrestadorEspecialidade lngESPECIALIDADEID, _
                                                          lngPRESTADORID
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

