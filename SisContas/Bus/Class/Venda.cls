VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsVenda"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub ExcluirTAB_VENDACARD(ByVal lngTAB_VENDACARDID As Long)
  Dim clsVend    As datSisContas.clsVenda
  '
  On Error GoTo trata
  Set clsVend = New datSisContas.clsVenda
  '
  'Para cada item do Pedido, mata Estorno, mesmo sabendo-se
  'que não irá haver estorno
  clsVend.ExcluirTAB_VENDACARD lngTAB_VENDACARDID
  '
  Set clsVend = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub EstornaTAB_VENDACARD(ByVal lngTAB_VENDACARDID As Long, _
                                ByVal lngVENDAID As Long, _
                                ByVal intITEM As Integer, _
                                lngQtdVendaAEstornar As Long)
  '
  Dim objRs     As ADODB.Recordset
  Dim strSql    As String
  Dim clsGer    As datSisContas.clsGeral
  Dim clsVend   As datSisContas.clsVenda
  Dim clsCard   As datSisContas.clsCardapio
  '
  On Error GoTo trata
  '
  Set clsGer = New datSisContas.clsGeral
  Set clsVend = New datSisContas.clsVenda
  Set clsCard = New datSisContas.clsCardapio
  '
  strSql = "Select Count(*) As Total From TAB_VENDACARD Where VENDAID = " & lngVENDAID & " And ITEM = " & intITEM + 1
  Set objRs = clsGer.ExecutarSQL(strSql)
  If objRs.Fields("Total").Value > 0 Then
    'Existe algo aonde vai entrar o Estorno, Então tem que Mover Itens
    clsVend.UpdateItemTAB_VENDACARD intITEM
  End If
  '
  objRs.Close
  Set objRs = Nothing
  '
  'Após ter aberto espaço na tabela, INSERIR ESTORNO
  'Antes pega Dados do Pedido Item
  strSql = "SELECT TAB_VENDACARD.PKID AS TABVENDACARDEXCID, CARDAPIO.VALOR AS VALORDOCARD, * FROM TAB_VENDACARD INNER JOIN CARDAPIO ON (TAB_VENDACARD.CARDAPIOID = CARDAPIO.PKID) WHERE TAB_VENDACARD.PKID = " & _
    lngTAB_VENDACARDID

  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    'Caso encontrou, Insere
    clsCard.InserirTAB_VENDACARD _
      objRs.Fields("VENDAID").Value, _
      objRs.Fields("CARDAPIOID").Value, _
      CStr((lngQtdVendaAEstornar * (-1))), _
      CStr(objRs.Fields("VALORDOCARD").Value * lngQtdVendaAEstornar * -1), _
      CStr(objRs.Fields("Item").Value + 1), _
      CStr(objRs.Fields("TABVENDACARDEXCID").Value), _
      ""

    '
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set clsGer = Nothing
  Set clsVend = Nothing
  Set clsCard = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Function VerificaEstorno(ByVal lngTAB_VENDACARDID As Long, _
                                strRet As String, _
                                lngQtdRestanteProduto As Long, _
                                curValorUnitaioItem As Currency) As Boolean
  
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim clsVend         As datSisContas.clsVenda
  Dim clsGer          As datSisContas.clsGeral
  Dim lngQtdProduto   As Long
  Dim lngQtdEstornada As Long
  '
  On Error GoTo trata
  '
  Set clsVend = New datSisContas.clsVenda
  Set clsGer = New datSisContas.clsGeral
  '
  strRet = ""
  Set objRs = clsVend.SelecionarTAB_VENDACARD(lngTAB_VENDACARDID)
  If objRs.EOF Then
    strRet = "Produto Não cadastrado"
  ElseIf IsNumeric(objRs.Fields("TAB_VENDACARDEXCID").Value) Then
    strRet = "Produto Não pode ser Excluído por ser um Estorno"
  Else 'OK
    lngQtdProduto = objRs.Fields("QUANTIDADE").Value
  End If
  '
  objRs.Close
  Set objRs = Nothing
  '
  If Len(Trim(strRet)) = 0 Then 'Não Houve Erro
    'Verifica se Existe Estorno para Item
    strSql = "Select SUM(QUANTIDADE) AS QTDESTORNADO From TAB_VENDACARD WHERE TAB_VENDACARDEXCID = " & lngTAB_VENDACARDID
    Set objRs = clsGer.ExecutarSQL(strSql)
    '
    If Not objRs.EOF Then
      If IsNumeric(objRs.Fields("QTDESTORNADO").Value) Then
        'Verifica Quantidade Lançada
        lngQtdEstornada = objRs.Fields("QTDESTORNADO").Value
        If lngQtdProduto + lngQtdEstornada <= 0 Then
          strRet = "Produto Não pode ser excluído, pois já foi totalmente estornado."
        End If
      End If
    End If
    '
    objRs.Close
    Set objRs = Nothing
    '
  End If
  If Len(Trim(strRet)) = 0 Then 'Não Houve Erro
    'Verifica Valor unitário do Item (só para retorno)
    strSql = "Select CARDAPIO.VALOR AS VALORUNITARIO From CARDAPIO INNER JOIN TAB_VENDACARD ON (CARDAPIO.PKID = TAB_VENDACARD.CARDAPIOID) WHERE TAB_VENDACARD.PKID = " & lngTAB_VENDACARDID
    Set objRs = clsGer.ExecutarSQL(strSql)
    '
    If Not objRs.EOF Then
      If IsNumeric(objRs.Fields("VALORUNITARIO").Value) Then
        'Verifica Quantidade Lançada
        curValorUnitaioItem = objRs.Fields("VALORUNITARIO").Value
      End If
    End If
    '
    objRs.Close
    Set objRs = Nothing
    '
  End If
  If Len(Trim(strRet)) = 0 Then
    'Ok
    VerificaEstorno = True
  Else
    VerificaEstorno = False
  End If
  Set clsVend = Nothing
  lngQtdRestanteProduto = lngQtdProduto + lngQtdEstornada
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirVenda(lngVENDAID As Long)

  On Error GoTo trata
  Dim clsVend As datSisContas.clsVenda
  Set clsVend = New datSisContas.clsVenda
  
  'Inserir Venda
  clsVend.ExlcuirVenda lngVENDAID
  '
  Set clsVend = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarVenda(lngVENDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsVend As datSisContas.clsVenda
  Set clsVend = New datSisContas.clsVenda
  
  'Listar Vendas
  Set ListarVenda = clsVend.ListarVenda(lngVENDAID)
  '
  Set clsVend = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub InserirVenda(lngVENDAID As Long, _
                        strData As String, _
                        strNome As String, _
                        strFuncao As String, _
                        lngTurnoId As Long, _
                        strDescricao As String, _
                        strVenda As String, _
                        strCobranca As String, _
                        strUsuario As String, _
                        lngSequencial As Long, _
                        lngCONFIGURACAOID As Long)
  On Error GoTo trata
  Dim clsVend As datSisContas.clsVenda
  Set clsVend = New datSisContas.clsVenda
  
  'Inserir Venda
  clsVend.InserirVenda lngVENDAID, _
                       strData, _
                       strNome, _
                       strFuncao, _
                       lngTurnoId, _
                       strDescricao, _
                       strVenda, _
                       strCobranca, _
                       strUsuario, _
                       lngSequencial, _
                       lngCONFIGURACAOID

  'Capturar PEDIDOID
  'lngVENDAID = clsVend.CapturarVENDAID(lngSequencial)
  '
  Set clsVend = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarVenda(lngVENDAID As Long, _
                        strNome As String, _
                        strFuncao As String, _
                        strDescricao As String, _
                        strVenda As String, _
                        strCobranca As String, _
                        lngCONFIGURACAOID As Long)
  On Error GoTo trata
  Dim clsVend As datSisContas.clsVenda
  Set clsVend = New datSisContas.clsVenda
  
  'Inserir Venda
  clsVend.AlterarVenda lngVENDAID, _
                       strNome, _
                       strFuncao, _
                       strDescricao, _
                       strVenda, _
                       strCobranca, _
                       lngCONFIGURACAOID

  '
  Set clsVend = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Function ListarTAB_VENDACARD(lngVENDAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsVend As datSisContas.clsVenda
  Set clsVend = New datSisContas.clsVenda
  
  'Listar Venda
  Set ListarTAB_VENDACARD = clsVend.ListarTAB_VENDACARD(lngVENDAID)
  '
  Set clsVend = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

