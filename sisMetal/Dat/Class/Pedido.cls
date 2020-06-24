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

Public Sub ExcluirPedido(lngPEDIDOID As Long, _
                         strData As String, _
                         strStatus As String)
  On Error GoTo trata
  
  '
  Dim strSql As String

  strSql = "UPDATE PEDIDO " & _
    " SET CANCELADO = " & Formata_Dados(strStatus, tpDados_Texto) & _
    ", DATA_CANCELAMENTO = " & Formata_Dados(strData, tpDados_DataHora) & _
    " WHERE PKID = " & _
    Formata_Dados(lngPEDIDOID, tpDados_Longo, tpNulo_NaoAceita)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedido.ExcluirPedido]", _
            Err.Description
End Sub

Public Function ListarPedido(lngPEDIDOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT PEDIDO.*, FORNECEDOR.NOME AS NOME_FORNECEDOR, " & _
    " ANODIZADORA.NOME AS NOME_ANODIZADORA, FABRICA.NOME AS NOME_FABRICA FROM PEDIDO " & _
    " LEFT JOIN LOJA AS FORNECEDOR ON PEDIDO.FORNECEDORID = FORNECEDOR.PKID " & _
    " LEFT JOIN LOJA AS ANODIZADORA ON PEDIDO.ANODIZADORAID = ANODIZADORA.PKID " & _
    " LEFT JOIN LOJA AS FABRICA ON PEDIDO.FABRICAID = FABRICA.PKID " & _
    " WHERE PEDIDO.PKID = " & _
    Formata_Dados(lngPEDIDOID, tpDados_Longo, tpNulo_NaoAceita)
  Set ListarPedido = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedido.ListarPedido]", _
            Err.Description
End Function

Public Function ListarPedidoAberto(lngPEDIDOID As Long) As ADODB.Recordset
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "SELECT PEDIDO.* " & _
    " FROM PEDIDO " & _
    " INNER JOIN ITEM_PEDIDO ON PEDIDO.PKID = ITEM_PEDIDO.PEDIDOID " & _
    " WHERE PEDIDO.PKID = " & _
    Formata_Dados(lngPEDIDOID, tpDados_Longo, tpNulo_NaoAceita) & _
    " AND ISNULL(PESO_INI,0) <> (ISNULL(PESO,0) + ISNULL(PESO_FAB,0)) "
  Set ListarPedidoAberto = RunSPReturnRS(strSql)
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedido.ListarPedidoAberto]", _
            Err.Description
End Function


Public Sub InserirPedido(lngPEDIDOID As Long, _
                         lngFORNECEDORID As Long, _
                         lngANODIZADORAID As Long, _
                         lngFABRICAID As Long, _
                         strValorAluminio As String, _
                         strAno As String, _
                         strNumroOS As String, _
                         strData As String, _
                         strStatus As String, _
                         strCancelado As String)
  On Error GoTo trata
  '
  Dim objRs As ADODB.Recordset
  Dim strSql As String
  '
  strSql = "SET NOCOUNT ON; "
  strSql = strSql & "INSERT INTO PEDIDO (DATA, OS_NUMERO, OS_ANO, FORNECEDORID, FABRICAID, ANODIZADORAID, STATUS, VALOR_ALUMINIO, CANCELADO) VALUES (" & _
    Formata_Dados(strData, tpDados_DataHora) & ", " & _
    Formata_Dados(strNumroOS, tpDados_Longo) & ", " & _
    Formata_Dados(strAno, tpDados_Longo) & ", " & _
    Formata_Dados(lngFORNECEDORID, tpDados_Longo) & ", " & _
    Formata_Dados(IIf(lngFABRICAID = 0, "", lngFABRICAID), tpDados_Longo) & ", " & _
    Formata_Dados(IIf(lngANODIZADORAID = 0, "", lngANODIZADORAID), tpDados_Longo) & ", " & _
    Formata_Dados(strStatus, tpDados_Texto) & ", " & _
    Formata_Dados(strValorAluminio, tpDados_Moeda) & ", " & _
    Formata_Dados(strCancelado, tpDados_Texto) & ");"
  strSql = strSql & "SELECT @@IDENTITY; SET NOCOUNT OFF;"
  '
  Set objRs = RunSPReturnRS(strSql)
  If Not objRs.EOF Then
    lngPEDIDOID = objRs.Fields(0).Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedido.InserirPedido]", _
            Err.Description
End Sub

Public Sub AlterarPedido(lngPEDIDOID As Long, _
                         lngFORNECEDORID As Long, _
                         lngANODIZADORAID As Long, _
                         lngFABRICAID As Long, _
                         strValorAluminio As String)
  On Error GoTo trata
  '
  Dim strSql As String

  strSql = "UPDATE PEDIDO SET " & _
    "FORNECEDORID = " & Formata_Dados(lngFORNECEDORID, tpDados_Longo) & _
    ", FABRICAID = " & Formata_Dados(IIf(lngFABRICAID = 0, "", lngFABRICAID), tpDados_Longo) & _
    ", ANODIZADORAID = " & Formata_Dados(IIf(lngANODIZADORAID = 0, "", lngANODIZADORAID), tpDados_Longo) & _
    ", VALOR_ALUMINIO = " & Formata_Dados(strValorAluminio, tpDados_Moeda) & _
    " WHERE PEDIDO.PKID = " & Formata_Dados(lngPEDIDOID, tpDados_Longo)
  '
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedido.AlterarPedido]", _
            Err.Description
End Sub

Public Function CapturarSequencialPedido(strAno As String, _
                                         lngFORNECEDORID As Long, _
                                         strNumeroOS As String) As Long
  On Error GoTo trata
  '
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  '
  strSql = "SELECT PKID, SEQUENCIAL FROM SEQUENCIAL_OS " & _
      " WHERE ANO = " & Formata_Dados(strAno, tpDados_Longo) & _
      " AND FORNECEDORID = " & Formata_Dados(lngFORNECEDORID, tpDados_Longo)
  Set objRs = RunSPReturnRS(strSql)
  If objRs.EOF Then
    strNumeroOS = "1"
    '----------- INSERIR SEQUENCIAL
    strSql = "INSERT INTO SEQUENCIAL_OS (" & _
      " ANO, FORNECEDORID, SEQUENCIAL) VALUES (" & _
      Formata_Dados(strAno, tpDados_Longo) & _
      ", " & Formata_Dados(lngFORNECEDORID, tpDados_Longo) & _
      ", " & Formata_Dados(1, tpDados_Longo) & _
      ")"
    '
    RunSP strSql
    '
  Else
    strNumeroOS = objRs.Fields("SEQUENCIAL").Value + 1
    '----------- ATUALIZA SEQUENCIAL
    strSql = "UPDATE SEQUENCIAL_OS SET " & _
      "SEQUENCIAL = " & Formata_Dados(strNumeroOS, tpDados_Longo) & _
      " WHERE ANO = " & Formata_Dados(strAno, tpDados_Longo) & " " & _
      " AND FORNECEDORID = " & Formata_Dados(lngFORNECEDORID, tpDados_Longo)
    RunSP strSql
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedido.CapturarSequencialPedido]", _
            Err.Description
End Function

Public Sub AlterarStatusComprador(lngPEDIDOID As Long, _
                                  strStatus As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PEDIDO SET "
  strSql = strSql & " STATUS = " & Formata_Dados(strStatus, tpDados_Texto)
  '
  strSql = strSql & " WHERE PKID = " & Formata_Dados(lngPEDIDOID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedido.AlterarStatusComprador]", _
            Err.Description
End Sub

Public Sub AlterarStatusFornecedor(lngPEDIDOID As Long, _
                                   strStatus As String, _
                                   strData As String)

  On Error GoTo trata
  '
  Dim strSql As String
  
  strSql = "UPDATE PEDIDO SET "
  strSql = strSql & " STATUS = " & Formata_Dados(strStatus, tpDados_Texto)
  strSql = strSql & ", DATA_ENVIO_FORN = " & Formata_Dados(strData, tpDados_DataHora)
  strSql = strSql & " WHERE PKID = " & Formata_Dados(lngPEDIDOID, tpDados_Longo)
  RunSP strSql
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[clsPedido.AlterarStatusFornecedor]", _
            Err.Description
End Sub



