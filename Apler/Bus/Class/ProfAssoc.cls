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

Public Sub AssociarProfAoAssociado(lngPROFISSAOID As Long, _
                                   lngASSOCIADOID As Long)
  On Error GoTo trata
  '
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim objProfAssoc  As datApler.clsProfAssoc
  '
  Set objProfAssoc = New datApler.clsProfAssoc
  '
  objProfAssoc.InserirProfAssoc lngPROFISSAOID, _
                                lngASSOCIADOID
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub DesassociarProfDoAssociado(lngPROFISSAOID As Long, _
                                      lngASSOCIADOID As Long)
  On Error GoTo trata
  '
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim objProfAssoc  As datApler.clsProfAssoc
  '
  Set objProfAssoc = New datApler.clsProfAssoc
  '
  objProfAssoc.ExcluirProfAssoc lngPROFISSAOID, _
                                lngASSOCIADOID
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

