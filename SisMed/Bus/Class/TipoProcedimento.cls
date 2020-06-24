VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTipoProcedimento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirTipoProcedimento(strTipoProcedimento As String)

  On Error GoTo trata
  Dim objTipoProcedimento As datSisMed.clsTipoProcedimento
  '
  Set objTipoProcedimento = New datSisMed.clsTipoProcedimento
  
  'Inserir TipoProcedimento
  objTipoProcedimento.InserirTipoProcedimento strTipoProcedimento
  '
  Set objTipoProcedimento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarTipoProcedimento(lngTIPOPROCEDIMENTOID As Long, _
                         strTipoProcedimento As String)
                         
  On Error GoTo trata
  Dim objTipoProcedimento As datSisMed.clsTipoProcedimento
  Set objTipoProcedimento = New datSisMed.clsTipoProcedimento
  
  'Inserir TipoProcedimento
  objTipoProcedimento.AlterarTipoProcedimento lngTIPOPROCEDIMENTOID, _
                          strTipoProcedimento

  '
  Set objTipoProcedimento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarTipoProcedimento(lngTIPOPROCEDIMENTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objTipoProcedimento As datSisMed.clsTipoProcedimento
  Set objTipoProcedimento = New datSisMed.clsTipoProcedimento
  
  'Listar TipoProcedimento
  Set ListarTipoProcedimento = objTipoProcedimento.ListarTipoProcedimento(lngTIPOPROCEDIMENTOID)
  '
  Set objTipoProcedimento = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirTipoProcedimento(lngTIPOPROCEDIMENTOID As Long)

  On Error GoTo trata
  Dim objTipoProcedimento As datSisMed.clsTipoProcedimento
  Set objTipoProcedimento = New datSisMed.clsTipoProcedimento
  
  'Excluir TipoProcedimento
  objTipoProcedimento.ExlcuirTipoProcedimento lngTIPOPROCEDIMENTOID
  '
  Set objTipoProcedimento = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

