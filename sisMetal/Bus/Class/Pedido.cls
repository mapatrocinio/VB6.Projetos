VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirPedido(ByVal lngPEDIDOID As Long, _
                         strStatus As String)
  On Error GoTo trata
  '
  Dim objPedido       As datSisMetal.clsPedido
  Dim strData         As String
  Dim strStatusFinal  As String
  '
  If strStatus = "E" Then
    strStatusFinal = "N"
  Else
    strStatusFinal = "S"
  End If
  
  strData = Format(Now, "DD/MM/YYYY hh:mm")
  Set objPedido = New datSisMetal.clsPedido
  '
  'Excluir Pedido
  objPedido.ExcluirPedido lngPEDIDOID, _
                          strData, _
                          strStatusFinal
  '
  '
  Set objPedido = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Sub InserirPedido(lngPEDIDOID As Long, _
                         lngFORNECEDORID As Long, _
                         lngANODIZADORAID As Long, _
                         lngFABRICAID As Long, _
                         strValorAluminio As String)
  On Error GoTo trata
  Dim objPedido       As datSisMetal.clsPedido
  Dim strAno          As String
  Dim strNumeroOS     As String
  Dim strData         As String
  Dim strStatus       As String
  Dim strCancelado    As String
  '
  Set objPedido = New datSisMetal.clsPedido
  '
  strData = Format(Now, "DD/MM/YYYY hh:mm")
  strAno = Format(Now, "YYYY")
  strStatus = "I"
  strCancelado = "N"
  'Capturar Sequencial do Pedido
  objPedido.CapturarSequencialPedido strAno, _
                                     lngFORNECEDORID, _
                                     strNumeroOS
  'Inserir Pedido
  objPedido.InserirPedido lngPEDIDOID, _
                          lngFORNECEDORID, _
                          lngANODIZADORAID, _
                          lngFABRICAID, _
                          strValorAluminio, _
                          strAno, _
                          strNumeroOS, _
                          strData, _
                          strStatus, _
                          strCancelado
  '
  Set objPedido = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarPedido(lngPEDIDOID As Long, _
                         lngFORNECEDORID As Long, _
                         lngANODIZADORAID As Long, _
                         lngFABRICAID As Long, _
                         strValorAluminio As String)
  On Error GoTo trata
  Dim objPedido As datSisMetal.clsPedido
  '
  Set objPedido = New datSisMetal.clsPedido
  '
  'Alterar Pedido
  objPedido.AlterarPedido lngPEDIDOID, _
                          lngFORNECEDORID, _
                          lngANODIZADORAID, _
                          lngFABRICAID, _
                          strValorAluminio
  '
  Set objPedido = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarPedido(lngPEDIDOID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objPedido As datSisMetal.clsPedido
  Set objPedido = New datSisMetal.clsPedido

  'Listar Pedido
  Set ListarPedido = objPedido.ListarPedido(lngPEDIDOID)
  '
  Set objPedido = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub AlterarStatusComprador(lngPEDIDOID As Long)
  On Error GoTo trata
  Dim objPedido As datSisMetal.clsPedido
  Dim strStatus As String
  Set objPedido = New datSisMetal.clsPedido
  '
  strStatus = "C"
  'Alterar Status para Comprador
  objPedido.AlterarStatusComprador lngPEDIDOID, _
                                   strStatus
  '
  Set objPedido = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarStatusFornecedor(lngPEDIDOID As Long)
  On Error GoTo trata
  Dim objPedido As datSisMetal.clsPedido
  Dim strData As String
  Dim strStatus As String
  Set objPedido = New datSisMetal.clsPedido
  '
  strStatus = "F"
  strData = Format(Now, "DD/MM/YYYY hh:mm")
  'Alterar Status para Fornecedor
  objPedido.AlterarStatusFornecedor lngPEDIDOID, _
                                    strStatus, _
                                    strData
  '
  Set objPedido = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ValidaPedidoFechado(lngPEDIDOID As Long) As Boolean
  On Error GoTo trata
  Dim objPedido As datSisMetal.clsPedido
  Dim blnRetorno As Boolean
  Dim objRs As ADODB.Recordset
  Set objPedido = New datSisMetal.clsPedido
  blnRetorno = True
  'Listar Pedido
  Set objRs = objPedido.ListarPedidoAberto(lngPEDIDOID)
  If Not objRs.EOF Then
    blnRetorno = False
  End If
  '
  objRs.Close
  Set objRs = Nothing
  Set objPedido = Nothing
  ValidaPedidoFechado = blnRetorno
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

