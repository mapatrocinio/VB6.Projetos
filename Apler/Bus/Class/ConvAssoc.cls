VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConvAssoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarConvAssocPeloPkid(lngConvAssocId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objConvAssoc As datApler.clsConvAssoc
  Set objConvAssoc = New datApler.clsConvAssoc
  
  'Listar ConvAssoc
  Set SelecionarConvAssocPeloPkid = objConvAssoc.SelecionarConvAssocPeloPkid(lngConvAssocId)
  '
  Set objConvAssoc = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirConvAssoc(lngConvAssocId As Long)
  On Error GoTo trata
  Dim objConvAssoc As datApler.clsConvAssoc
  Set objConvAssoc = New datApler.clsConvAssoc
  
  'Excluir ConvAssoc
  objConvAssoc.ExcluirConvAssoc lngConvAssocId
  '
  Set objConvAssoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirConvAssoc(lngASSOCIADOID As Long, _
                            lngPLANOCONVENIOID As Long, _
                            strDataInicio As String, _
                            strDataTermino As String)
  On Error GoTo trata
  Dim objConvAssoc     As datApler.clsConvAssoc
  '
  Set objConvAssoc = New datApler.clsConvAssoc
  'Inserir ConvAssoc
  objConvAssoc.InserirConvAssoc lngASSOCIADOID, _
                                lngPLANOCONVENIOID, _
                                strDataInicio, _
                                strDataTermino
  '
  Set objConvAssoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarConvAssoc(lngPKID As Long, _
                            lngPLANOCONVENIOID As Long, _
                            strDataInicio As String, _
                            strDataTermino As String)

  On Error GoTo trata
  Dim objConvAssoc     As datApler.clsConvAssoc
  '
  Set objConvAssoc = New datApler.clsConvAssoc
  'Atualizar ConvAssoc
  objConvAssoc.AlterarConvAssoc lngPKID, _
                                lngPLANOCONVENIOID, _
                                strDataInicio, _
                                strDataTermino
  '
  Set objConvAssoc = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

