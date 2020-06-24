VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEmpresaCtrto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function SelecionarEmpresaCtrtoPelaDescricao(lngCONTRATOID As Long, _
                                                    strDescricao As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objEmpresaCtrto As datSisContas.clsEmpresaCtrto
  Set objEmpresaCtrto = New datSisContas.clsEmpresaCtrto
  
  'Listar EmpresaCtrto
  Set SelecionarEmpresaCtrtoPelaDescricao = objEmpresaCtrto.SelecionarEmpresaCtrtoPelaDescricao(lngCONTRATOID, _
                                                                                                strDescricao)
  '
  Set objEmpresaCtrto = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarEmpresaCtrtoPeloPkid(lngEMPRESACTRTOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objEmpresaCtrto As datSisContas.clsEmpresaCtrto
  Set objEmpresaCtrto = New datSisContas.clsEmpresaCtrto
  
  'Listar EmpresaCtrto
  Set SelecionarEmpresaCtrtoPeloPkid = objEmpresaCtrto.SelecionarEmpresaCtrtoPeloPkid(lngEMPRESACTRTOID)
  '
  Set objEmpresaCtrto = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirEmpresaCtrto(lngEMPRESACTRTOID As Long)
  On Error GoTo trata
  Dim objEmpresaCtrto As datSisContas.clsEmpresaCtrto
  Set objEmpresaCtrto = New datSisContas.clsEmpresaCtrto
  
  'Excluir EmpresaCtrto
  objEmpresaCtrto.ExcluirEmpresaCtrto lngEMPRESACTRTOID
  '
  Set objEmpresaCtrto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirEmpresaCtrto(lngCONTRATOID As Long, _
                               lngEMPRESAID As Long, _
                               strDescricao As String, _
                               strValor As String)
  On Error GoTo trata
  Dim objEmpresaCtrto     As datSisContas.clsEmpresaCtrto
  '
  Set objEmpresaCtrto = New datSisContas.clsEmpresaCtrto
  'Inserir EmpresaCtrto
  objEmpresaCtrto.InserirEmpresaCtrto lngCONTRATOID, _
                                      lngEMPRESAID, _
                                      strDescricao, _
                                      strValor
  '
  Set objEmpresaCtrto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarEmpresaCtrto(lngPKID As Long, _
                               lngEMPRESAID As Long, _
                               strDescricao As String, _
                               strValor As String, _
                               strStatus As String)

  On Error GoTo trata
  Dim objEmpresaCtrto     As datSisContas.clsEmpresaCtrto
  '
  Set objEmpresaCtrto = New datSisContas.clsEmpresaCtrto
  'Atualizar EmpresaCtrto
  objEmpresaCtrto.AlterarEmpresaCtrto lngPKID, _
                                      lngEMPRESAID, _
                                      strDescricao, _
                                      strValor, _
                                      strStatus
  '
  Set objEmpresaCtrto = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

