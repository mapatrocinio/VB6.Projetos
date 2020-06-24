VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcedimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirProcedimento(strProcedimento As String, _
                               strTipoProcedimentoId As String, _
                               strValor As String, _
                               strAceitaValor As String, _
                               strConsulta As String)

  On Error GoTo trata
  Dim objProcedimento As datSisMed.clsProcedimento
  '
  Set objProcedimento = New datSisMed.clsProcedimento
  
  'Inserir Procedimento
  objProcedimento.InserirProcedimento strProcedimento, _
                                      strTipoProcedimentoId, _
                                      strValor, _
                                      strAceitaValor, _
                                      strConsulta
  '
  Set objProcedimento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarProcedimento(lngPROCEDIMENTOID As Long, _
                               strProcedimento As String, _
                               strTipoProcedimentoId As String, _
                               strValor As String, _
                               strAceitaValor As String, _
                               strConsulta As String)
                         
  On Error GoTo trata
  Dim objProcedimento As datSisMed.clsProcedimento
  Set objProcedimento = New datSisMed.clsProcedimento
  
  'Inserir Procedimento
  objProcedimento.AlterarProcedimento lngPROCEDIMENTOID, _
                                      strProcedimento, _
                                      strTipoProcedimentoId, _
                                      strValor, _
                                      strAceitaValor, _
                                      strConsulta

  '
  Set objProcedimento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarProcedimento(lngPROCEDIMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objProcedimento As datSisMed.clsProcedimento
  Set objProcedimento = New datSisMed.clsProcedimento
  
  'Listar Procedimento
  Set ListarProcedimento = objProcedimento.ListarProcedimento(lngPROCEDIMENTOID)
  '
  Set objProcedimento = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirProcedimento(lngPROCEDIMENTOID As Long)

  On Error GoTo trata
  Dim objProcedimento As datSisMed.clsProcedimento
  Set objProcedimento = New datSisMed.clsProcedimento
  
  'Excluir Procedimento
  objProcedimento.ExlcuirProcedimento lngPROCEDIMENTOID
  '
  Set objProcedimento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

