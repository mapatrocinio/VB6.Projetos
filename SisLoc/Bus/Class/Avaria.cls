VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsAvaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirItemAvaria(lngITEMAVARIAID As Long)

  On Error GoTo trata
  Dim objAvaria As datSisLoc.clsAvaria
  '
  Set objAvaria = New datSisLoc.clsAvaria
  'Alterar Estoque Pela Avaria
  objAvaria.ExlcuirItemAvaria lngITEMAVARIAID
  '
  Set objAvaria = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirItemAvaria(lngITEMNFID As Long, _
                             lngAVARIAID As Long, _
                             strQuantidade As String, _
                             strValorIndenizacao As String)

  On Error GoTo trata
  Dim objAvaria As datSisLoc.clsAvaria
  '
  Set objAvaria = New datSisLoc.clsAvaria
  'Alterar Estoque Pela Avaria
  objAvaria.InserirItemAvaria lngITEMNFID, _
                              lngAVARIAID, _
                              strQuantidade, _
                              strValorIndenizacao
  '
  Set objAvaria = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarEstoquePelaAvaria(ByVal lngESTOQUEID As Long, _
                                    ByVal strQuantidade As String, _
                                    ByVal strStatus As String)

  On Error GoTo trata
  Dim objAvaria As datSisLoc.clsAvaria
  '
  Set objAvaria = New datSisLoc.clsAvaria
  'Alterar Estoque Pela Avaria
  objAvaria.AlterarEstoquePelaAvaria lngESTOQUEID, _
                                           strQuantidade, _
                                           strStatus
  '
  Set objAvaria = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub InserirAvaria(ByRef lngAVARIAID As Long, _
                         ByRef lngNFID As Long, _
                         ByRef strSequencial As String, _
                         strData As String, _
                         strDescricao As String)

  On Error GoTo trata
  Dim objAvaria As datSisLoc.clsAvaria
  '
  Set objAvaria = New datSisLoc.clsAvaria
  'Inserir Avaria
  strSequencial = objAvaria.CapturaSequencialAvaria(lngNFID)
  '
  objAvaria.InserirAvaria lngAVARIAID, _
                          lngNFID, _
                          strSequencial, _
                          strData, _
                          strDescricao
  '
  Set objAvaria = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarAvaria(lngAVARIAID As Long, _
                         strData As String, _
                         strDescricao As String)
                         
  On Error GoTo trata
  Dim objAvaria As datSisLoc.clsAvaria
  Set objAvaria = New datSisLoc.clsAvaria
  
  'Inserir Avaria
  objAvaria.AlterarAvaria lngAVARIAID, _
                          strData, _
                          strDescricao

  '
  Set objAvaria = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarAvariaPeloSeq(lngNFID As Long, _
                                    strSequencial As String) As ADODB.Recordset
  On Error GoTo trata
  Dim objAvaria As datSisLoc.clsAvaria
  Set objAvaria = New datSisLoc.clsAvaria
  
  'Listar Avaria
  Set ListarAvariaPeloSeq = objAvaria.ListarAvariaPeloSeq(lngNFID, _
                                                          strSequencial)
  '
  Set objAvaria = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Function ListarAvaria(lngAVARIAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objAvaria As datSisLoc.clsAvaria
  Set objAvaria = New datSisLoc.clsAvaria
  
  'Listar Avaria
  Set ListarAvaria = objAvaria.ListarAvaria(lngAVARIAID)
  '
  Set objAvaria = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirAvaria(lngAVARIAID As Long)

  On Error GoTo trata
  Dim objAvaria As datSisLoc.clsAvaria
  Set objAvaria = New datSisLoc.clsAvaria
  
  'Excluir Avaria
  objAvaria.ExlcuirAvaria lngAVARIAID
  '
  Set objAvaria = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

