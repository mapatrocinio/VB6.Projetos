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

Public Sub AssociarGRPGTOGR(lngGRID As Long, _
                           lngGRPAGAMENTOID As Long)
  On Error GoTo trata
  '
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objGRPgto As datSisMed.clsGRPgto
  '
  Set objGRPgto = New datSisMed.clsGRPgto
  '
  objGRPgto.InserirGRPGTOGR lngGRID, _
                           lngGRPAGAMENTOID
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub
Public Sub DesassociarGRPGTOGR(lngGRPGTOID As Long)
  On Error GoTo trata
  '
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objGRPgto As datSisMed.clsGRPgto
  '
  Set objGRPgto = New datSisMed.clsGRPgto
  '
  objGRPgto.ExlcuirGRPGTOGR lngGRPGTOID
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


