VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsExtraForaUnidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub InserirTAB_EXTRAUNIDADE(ByVal lngEXTRAUNIDADEID As Long, _
                                   ByVal strQuantidade As String, _
                                   ByVal strValor As String, _
                                   ByVal strDescricao As String, _
                                   ByVal strTAB_EXTRAUNIDADEEXCID As String)
  '
  On Error GoTo trata
    
  Dim clsExtra          As datSisContas.clsExtraForaUnidade
  Dim intITEM           As Integer
  '
  Set clsExtra = New datSisContas.clsExtraForaUnidade
  '
  intITEM = clsExtra.CapturaItemTAB_EXTRAUNIDADE(lngEXTRAUNIDADEID)
  '
  clsExtra.InserirTAB_EXTRAUNIDADE lngEXTRAUNIDADEID, _
                                   strQuantidade, _
                                   strValor, _
                                   CStr(intITEM), _
                                   strDescricao, _
                                   strTAB_EXTRAUNIDADEEXCID
  '
  Set clsExtra = Nothing
  '
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub ExcluirTAB_EXTRAUNIDADE(ByVal lngTAB_EXTRAUNIDADEID As Long)
  Dim clsExtra As datSisContas.clsExtraForaUnidade
  '
  On Error GoTo trata
  Set clsExtra = New datSisContas.clsExtraForaUnidade
  '
  'Para cada item do Etxra, mata Estorno, mesmo sabendo-se
  'que não irá haver estorno
  clsExtra.ExcluirTAB_EXTRAUNIDADE lngTAB_EXTRAUNIDADEID
  '
  Set clsExtra = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub EstornaTAB_EXTRAUNIDADE(ByVal lngTAB_EXTRAUNIDADEID As Long, _
                                   ByVal lngEXTRAUNIDADEID As Long, _
                                   ByVal intITEM As Integer, _
                                   lngQtdAEstornar As Long, _
                                   curValorAEstornar As Currency)
  '
  Dim objRs     As ADODB.Recordset
  Dim strSql    As String
  Dim clsGer    As datSisContas.clsGeral
  Dim clsExtra  As datSisContas.clsExtraForaUnidade
  '
  On Error GoTo trata
  '
  Set clsGer = New datSisContas.clsGeral
  Set clsExtra = New datSisContas.clsExtraForaUnidade
  '
  strSql = "Select Count(*) As Total From TAB_EXTRAUNIDADE Where EXTRAUNIDADEID = " & lngEXTRAUNIDADEID & " And ITEM = " & intITEM + 1
  Set objRs = clsGer.ExecutarSQL(strSql)
  If objRs.Fields("Total").Value > 0 Then
    'Existe algo aonde vai entrar o Estorno, Então tem que Mover Itens
    clsExtra.UpdateItemTAB_EXTRAUNIDADE intITEM
  End If
  '
  objRs.Close
  Set objRs = Nothing
  '
  'Após ter aberto espaço na tabela, INSERIR ESTORNO
  clsExtra.InserirTAB_EXTRAUNIDADE lngEXTRAUNIDADEID, _
                                   CStr((lngQtdAEstornar * (-1))), _
                                   CStr((curValorAEstornar * (-1))), _
                                   intITEM + 1, _
                                   "Estorno", _
                                   lngTAB_EXTRAUNIDADEID

  '
  Set clsGer = Nothing
  Set clsExtra = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Function VerificaEstorno(ByVal lngTAB_EXTRAUNIDADEID As Long, _
                                strRet As String) As Boolean
  
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim clsExtra        As datSisContas.clsExtraForaUnidade
  Dim clsVend         As datSisContas.clsVenda
  Dim clsGer          As datSisContas.clsGeral
  Dim lngQtdProduto   As Long
  Dim lngQtdEstornada As Long
  '
  On Error GoTo trata
  '
  Set clsExtra = New datSisContas.clsExtraForaUnidade
  Set clsGer = New datSisContas.clsGeral
  '
  strRet = ""
  Set objRs = clsExtra.SelecionarTAB_EXTRAUNIDADE(lngTAB_EXTRAUNIDADEID)
  If objRs.EOF Then
    strRet = "Item do Extra Não cadastrado"
  ElseIf IsNumeric(objRs.Fields("TAB_EXTRAUNIDADEEXCID").Value) Then
    strRet = "Item do Extra Não pode ser Excluído por ser um Estorno"
  Else 'OK
    lngQtdProduto = objRs.Fields("QUANTIDADE").Value
  End If
  '
  objRs.Close
  Set objRs = Nothing
  '
  If Len(Trim(strRet)) = 0 Then 'Não Houve Erro
    'Verifica se Existe Estorno para Item
    strSql = "Select SUM(QUANTIDADE) AS QTDESTORNADO From TAB_EXTRAUNIDADE WHERE TAB_EXTRAUNIDADEEXCID = " & lngTAB_EXTRAUNIDADEID
    Set objRs = clsGer.ExecutarSQL(strSql)
    '
    If Not objRs.EOF Then
      If IsNumeric(objRs.Fields("QTDESTORNADO").Value) Then
        'Verifica Quantidade Lançada
        lngQtdEstornada = objRs.Fields("QTDESTORNADO").Value
        If lngQtdProduto + lngQtdEstornada <= 0 Then
          strRet = "Item do Etxra não pode ser estornado, pois já foi totalmente estornado."
        End If
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
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub ExcluirExtra(lngEXTRAUNIDADEID As Long)

  On Error GoTo trata
  Dim clsExtra As datSisContas.clsExtraForaUnidade
  Set clsExtra = New datSisContas.clsExtraForaUnidade
  
  'Excluir Extra
  clsExtra.ExlcuirExtra lngEXTRAUNIDADEID
  '
  Set clsExtra = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Function ListarExtra(lngEXTRAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsExtra As datSisContas.clsExtraForaUnidade
  Set clsExtra = New datSisContas.clsExtraForaUnidade
  
  'Listar extra
  Set ListarExtra = clsExtra.ListarExtra(lngEXTRAID)
  '
  Set clsExtra = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

Public Sub InserirExtra(lngEXTRAID As Long, _
                        strData As String, _
                        lngTurnoId As Long, _
                        strDescricao As String, _
                        strTipo As String, _
                        lngSequencial As Long, _
                        lngCONFIGURACAOID As Long)
  On Error GoTo trata
  Dim clsExtra As datSisContas.clsExtraForaUnidade
  Set clsExtra = New datSisContas.clsExtraForaUnidade
  
  'Inserir Extra
  clsExtra.InserirExtra strData, _
                        lngTurnoId, _
                        strDescricao, _
                        strTipo, _
                        lngSequencial, _
                        lngCONFIGURACAOID

  'Capturar EXTRAID
  lngEXTRAID = clsExtra.CapturarEXTRAID(lngSequencial)
  '
  Set clsExtra = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub

Public Sub AlterarExtra(lngEXTRAID As Long, _
                        strDescricao As String, _
                        lngCONFIGURACAOID As Long)
  On Error GoTo trata
  Dim clsExtra As datSisContas.clsExtraForaUnidade
  Set clsExtra = New datSisContas.clsExtraForaUnidade
  
  'Alterar extra
  clsExtra.AlterarExtra lngEXTRAID, _
                        strDescricao, _
                        lngCONFIGURACAOID

  '
  Set clsExtra = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Sub


Public Function ListarTAB_EXTRAUNIDADE(lngEXTRAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim clsExtra As datSisContas.clsExtraForaUnidade
  Set clsExtra = New datSisContas.clsExtraForaUnidade
  
  'Listar TAB_EXTRAUNIDADE
  Set ListarTAB_EXTRAUNIDADE = clsExtra.ListarTAB_EXTRAUNIDADE(lngEXTRAID)
  '
  Set clsExtra = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Source, Err.Description
End Function

