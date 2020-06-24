VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProcReceita"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarProcReceitaPeloPkid(lngProcReceitaId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objProcReceita As datSisMed.clsProcReceita
  Set objProcReceita = New datSisMed.clsProcReceita
  
  'Listar ProcReceita
  Set SelecionarProcReceitaPeloPkid = objProcReceita.SelecionarProcReceitaPeloPkid(lngProcReceitaId)
  '
  Set objProcReceita = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirProcReceita(lngProcReceitaId As Long)
  On Error GoTo trata
  Dim objProcReceita As datSisMed.clsProcReceita
  Set objProcReceita = New datSisMed.clsProcReceita
  
  'Excluir ProcReceita
  objProcReceita.ExcluirProcReceita lngProcReceitaId
  '
  Set objProcReceita = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirProcReceita(lngPROCEDIMENTOID As Long, _
                              strTipo As String, _
                              strDescricao As String, _
                              strStatus As String)
  On Error GoTo trata
  Dim objProcReceita     As datSisMed.clsProcReceita
  '
  Set objProcReceita = New datSisMed.clsProcReceita
  'Inserir ProcReceita
  objProcReceita.InserirProcReceita lngPROCEDIMENTOID, _
                                    strTipo, _
                                    strDescricao, _
                                    strStatus
  '
  Set objProcReceita = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarProcReceita(lngPKID As Long, _
                              strTipo As String, _
                              strDescricao As String, _
                              strStatus As String)

  On Error GoTo trata
  Dim objProcReceita     As datSisMed.clsProcReceita
  '
  Set objProcReceita = New datSisMed.clsProcReceita
  'Atualizar ProcReceita
  objProcReceita.AlterarProcReceita lngPKID, _
                                    strTipo, _
                                    strDescricao, _
                                    strStatus
  '
  Set objProcReceita = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

