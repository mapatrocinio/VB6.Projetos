VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsContrato"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Public Function CapturaContrato(strCodigo As String, _
                                strEmpresa As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objContrato As datSisLoc.clsContrato
  Set objContrato = New datSisLoc.clsContrato
  
  'Capturar Contrato
  Set CapturaContrato = objContrato.CapturaContrato(strCodigo, _
                                                    strEmpresa)
  '
  Set objContrato = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarContrato(lngContratoId As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objContrato As datSisLoc.clsContrato
  Set objContrato = New datSisLoc.clsContrato
  
  'Listar Contrato
  Set SelecionarContrato = objContrato.SelecionarContrato(lngContratoId)
  '
  Set objContrato = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function SelecionarContratoPeloPkid(lngContratoId As Long, _
                                           lngOBRAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objContrato As datSisLoc.clsContrato
  Set objContrato = New datSisLoc.clsContrato
  
  'Listar Contrato
  Set SelecionarContratoPeloPkid = objContrato.SelecionarContratoPeloPkid(lngContratoId, _
                                                                          lngOBRAID)
  '
  Set objContrato = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirContrato(lngContratoId As Long)
  On Error GoTo trata
  Dim objContrato As datSisLoc.clsContrato
  Set objContrato = New datSisLoc.clsContrato
  
  'Excluir Contrato
  objContrato.ExcluirContrato lngContratoId
  '
  Set objContrato = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirContrato(strNumero As String, _
                           strSequencial As String, _
                           strAno As String, _
                           strDataInicial As String, _
                           strDataFim As String, _
                           strEmpresaId As String, _
                           strFuncionarioId As String)
  On Error GoTo trata
  Dim objContrato     As datSisLoc.clsContrato
  '
  Set objContrato = New datSisLoc.clsContrato
  'Inserir Contrato
  objContrato.InserirContrato strNumero, _
                              strSequencial, _
                              strAno, _
                              strDataInicial, _
                              strDataFim, _
                              strEmpresaId, _
                              strFuncionarioId
  '
  Set objContrato = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub AlterarContrato(lngPKID As Long, _
                           strDataInicial As String, _
                           strDataFim As String, _
                           strFuncionarioId As String)
  On Error GoTo trata
  Dim objContrato     As datSisLoc.clsContrato
  '
  Set objContrato = New datSisLoc.clsContrato
  'Atualizar Contrato
  objContrato.AlterarContrato lngPKID, _
                              strDataInicial, _
                              strDataFim, _
                              strFuncionarioId
  '
  Set objContrato = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

