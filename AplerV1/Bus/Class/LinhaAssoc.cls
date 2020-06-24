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

Public Sub AssociarLinhaAoAssociado(lngLINHAID As Long, _
                                    lngTITULARASSOCIADOID As Long)
  On Error GoTo trata
  '
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim objLinhaAssoc   As datApler.clsLinhaAssoc
  '
  Set objLinhaAssoc = New datApler.clsLinhaAssoc
  '
  objLinhaAssoc.InserirLinhaAssoc lngLINHAID, _
                                  lngTITULARASSOCIADOID
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub DesassociarLinhaDoAssociado(lngLINHAID As Long, _
                                       lngTITULARASSOCIADOID As Long)
  On Error GoTo trata
  '
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim objLinhaAssoc   As datApler.clsLinhaAssoc
  '
  Set objLinhaAssoc = New datApler.clsLinhaAssoc
  '
  objLinhaAssoc.ExcluirLinhaAssoc lngLINHAID, _
                                  lngTITULARASSOCIADOID
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

