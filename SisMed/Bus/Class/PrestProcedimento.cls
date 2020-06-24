VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrestProcedimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarPrestProcedimentoPeloPkid(lngPrestProcedimentoId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objPrestProcedimento As datSisMed.clsPrestProcedimento
  Set objPrestProcedimento = New datSisMed.clsPrestProcedimento
  
  'Listar PrestProcedimento
  Set SelecionarPrestProcedimentoPeloPkid = objPrestProcedimento.SelecionarPrestProcedimentoPeloPkid(lngPrestProcedimentoId)
  '
  Set objPrestProcedimento = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirPrestProcedimento(lngPrestProcedimentoId As Long)
  On Error GoTo trata
  Dim objPrestProcedimento As datSisMed.clsPrestProcedimento
  Set objPrestProcedimento = New datSisMed.clsPrestProcedimento
  
  'Excluir PrestProcedimento
  objPrestProcedimento.ExcluirPrestProcedimento lngPrestProcedimentoId
  '
  Set objPrestProcedimento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirPrestProcedimento(lngPRESTADORID As Long, _
                                    lngPROCEDIMENTOID As Long, _
                                    strPercCasa As String, _
                                    strPercPrestador As String, _
                                    strPercRX As String, _
                                    strPercTecRX As String, _
                                    strPercDonoUltra As String)
  On Error GoTo trata
  Dim objPrestProcedimento     As datSisMed.clsPrestProcedimento
  '
  Set objPrestProcedimento = New datSisMed.clsPrestProcedimento
  'Inserir PrestProcedimento
  objPrestProcedimento.InserirPrestProcedimento lngPRESTADORID, _
                                                lngPROCEDIMENTOID, _
                                                strPercCasa, _
                                                strPercPrestador, _
                                                strPercRX, _
                                                strPercTecRX, _
                                                strPercDonoUltra
  '
  Set objPrestProcedimento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarPrestProcedimento(lngPKID As Long, _
                                    lngPROCEDIMENTOID As Long, _
                                    strPercCasa As String, _
                                    strPercPrestador As String, _
                                    strPercRX As String, _
                                    strPercTecRX As String, _
                                    strPercDonoUltra As String)

  On Error GoTo trata
  Dim objPrestProcedimento     As datSisMed.clsPrestProcedimento
  '
  Set objPrestProcedimento = New datSisMed.clsPrestProcedimento
  'Atualizar PrestProcedimento
  objPrestProcedimento.AlterarPrestProcedimento lngPKID, _
                                                lngPROCEDIMENTOID, _
                                                strPercCasa, _
                                                strPercPrestador, _
                                                strPercRX, _
                                                strPercTecRX, _
                                                strPercDonoUltra
  '
  Set objPrestProcedimento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

