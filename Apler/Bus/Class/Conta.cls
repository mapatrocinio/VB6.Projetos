VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsConta"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function RetornarSaldo(strDtInicial As String, _
                              strDtFinal As String, _
                              lngCONTAID As Long)
  '
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim objRsMovDebAnt  As ADODB.Recordset
  Dim objRsMovCredAnt As ADODB.Recordset
  Dim objRsMovDebPer  As ADODB.Recordset
  Dim objRsMovCredPer As ADODB.Recordset
  
  Dim clsGer          As busApler.clsGeral
  '
  Dim strDescricao          As String
  Dim curVrSaldo            As Currency
  Dim curVrSaldoPeriodo     As Currency
  Dim curVrSaldoTotal       As Currency
  Dim curVrSaldoTotalGeral  As Currency
  Dim datDataInicial        As Date
  '
  Dim curMovDebAnt    As Currency
  Dim curMovCredAnt   As Currency
  Dim curMovDebPer    As Currency
  Dim curMovCredPer   As Currency
  '
  On Error GoTo trata
  '
  Set clsGer = New busApler.clsGeral
  '
  'Inicia Data Inicial
  datDataInicial = CDate(Right(strDtInicial, 4) & "/" & Mid(strDtInicial, 4, 2) & "/" & Mid(strDtInicial, 1, 2))
  '
  strSql = "Select * " & _
    "From CONTA " & _
    "WHERE CONTA.PKID = " & lngCONTAID & _
    " ORDER BY CONTA.PKID;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  'objRs.Filter = strWhere
  '
  'Débito no periodo
  strSql = "Select MOVIMENTACAO.* " & _
            "FROM MOVIMENTACAO " & _
            "Where " & _
            " MOVIMENTACAO.DATA >= " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
            " AND MOVIMENTACAO.DATA <= " & Formata_Dados(strDtFinal, tpDados_DataHora, tpNulo_NaoAceita) & _
            " AND not CONTADEBITOID is NULL " & _
            " AND CONTADEBITOID = " & lngCONTAID & _
            " Order By CONTADEBITOID"
  '
  Set objRsMovDebPer = clsGer.ExecutarSQL(strSql)
  '
  'Crédito no periodo
  strSql = "Select MOVIMENTACAO.* " & _
            "FROM MOVIMENTACAO " & _
            "Where " & _
            " MOVIMENTACAO.DATA >= " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
            " AND MOVIMENTACAO.DATA <= " & Formata_Dados(strDtFinal, tpDados_DataHora, tpNulo_NaoAceita) & _
            " AND not CONTACREDITOID is NULL " & _
            " AND CONTACREDITOID = " & lngCONTAID & _
            " Order By CONTACREDITOID"
  '
  Set objRsMovCredPer = clsGer.ExecutarSQL(strSql)
  'Débito anterior
  strSql = "Select MOVIMENTACAO.* " & _
            "FROM MOVIMENTACAO " & _
            "Where " & _
            " MOVIMENTACAO.DATA < " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
            " AND not CONTADEBITOID is NULL " & _
            " AND CONTADEBITOID = " & lngCONTAID & _
            " Order By CONTADEBITOID"
  '
  Set objRsMovDebAnt = clsGer.ExecutarSQL(strSql)
  'Crédito anterior
  strSql = "Select MOVIMENTACAO.* " & _
            "FROM MOVIMENTACAO " & _
            "Where " & _
            " MOVIMENTACAO.DATA < " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
            " AND not CONTACREDITOID is NULL " & _
            " AND CONTACREDITOID = " & lngCONTAID & _
            " Order By CONTACREDITOID"
  '
  Set objRsMovCredAnt = clsGer.ExecutarSQL(strSql)
  '
  RetornarSaldo = 0
  If Not objRs.EOF Then   'se já houver algum item
    'Pegar Descrição e saldo da conta
    strDescricao = objRs.Fields("DESCRICAO").Value
    curVrSaldoTotal = 0
    curVrSaldo = 0
    curVrSaldoPeriodo = 0
    If Not IsDate(objRs.Fields("DTSALDO").Value) Then
      curVrSaldo = 0
    ElseIf objRs.Fields("DTSALDO").Value < datDataInicial Then
      curVrSaldo = IIf(Not IsNumeric(objRs.Fields("VRSALDO").Value), 0, objRs.Fields("VRSALDO").Value)
    Else
      curVrSaldoPeriodo = IIf(Not IsNumeric(objRs.Fields("VRSALDO").Value), 0, objRs.Fields("VRSALDO").Value)
    End If
    'Pegar Movimentação anterior Débito
    curMovDebAnt = 0
    If Not objRsMovDebAnt.EOF Then   'se já houver algum registro
      Do While objRsMovDebAnt.Fields("CONTADEBITOID").Value & "" = objRs.Fields("PKID").Value & ""
        If Not IsDate(objRs.Fields("DTSALDO").Value) Then
          curMovDebAnt = curMovDebAnt + objRsMovDebAnt.Fields("VALOR").Value
        ElseIf objRs.Fields("DTSALDO").Value < datDataInicial Then
          If IsDate(objRsMovDebAnt.Fields("DATA").Value) Then
            If objRsMovDebAnt.Fields("DATA").Value >= objRs.Fields("DTSALDO").Value Then
              curMovDebAnt = curMovDebAnt + objRsMovDebAnt.Fields("VALOR").Value
            End If
          End If
        End If
        objRsMovDebAnt.MoveNext
        If objRsMovDebAnt.EOF Then Exit Do
      Loop
    End If
    'Pegar Movimentação anterior Crédito
    curMovCredAnt = 0
    If Not objRsMovCredAnt.EOF Then   'se já houver algum registro
      Do While objRsMovCredAnt.Fields("CONTACREDITOID").Value & "" = objRs.Fields("PKID").Value & ""
        If Not IsDate(objRs.Fields("DTSALDO").Value) Then
          curMovCredAnt = curMovCredAnt + objRsMovCredAnt.Fields("VALOR").Value
        ElseIf objRs.Fields("DTSALDO").Value < datDataInicial Then
          If IsDate(objRsMovCredAnt.Fields("DATA").Value) Then
            If objRsMovCredAnt.Fields("DATA").Value >= objRs.Fields("DTSALDO").Value Then
              curMovCredAnt = curMovCredAnt + objRsMovCredAnt.Fields("VALOR").Value
            End If
          End If
        End If
        objRsMovCredAnt.MoveNext
        If objRsMovCredAnt.EOF Then Exit Do
      Loop
    End If
    'Valor saldo da movimentada
    curVrSaldo = curVrSaldo - curMovDebAnt + curMovCredAnt
    'If strpMostrarApenasSaldo = "N" Then
      'Valor total saldo da movimentada
      curVrSaldoTotal = curVrSaldo
    'End If
    curVrSaldoTotal = curVrSaldoTotal + curVrSaldoPeriodo
    'Pegar Movimentação no período Débito
    curMovDebPer = 0
    If Not objRsMovDebPer.EOF Then   'se já houver algum registro
      Do While objRsMovDebPer.Fields("CONTADEBITOID").Value & "" = objRs.Fields("PKID").Value & ""
        If Not IsDate(objRs.Fields("DTSALDO").Value) Then
          curMovDebPer = curMovDebPer + objRsMovDebPer.Fields("VALOR").Value
          'Popular matriz de movimentação
        Else
          If IsDate(objRsMovDebPer.Fields("DATA").Value) Then
            If objRsMovDebPer.Fields("DATA").Value >= objRs.Fields("DTSALDO").Value Then
              curMovDebPer = curMovDebPer + objRsMovDebPer.Fields("VALOR").Value
              'Popular matriz de movimentação
            End If
          End If
        End If
        objRsMovDebPer.MoveNext
        If objRsMovDebPer.EOF Then Exit Do
      Loop
    End If
    'Pegar Movimentação no período Crédito
    curMovCredPer = 0
    If Not objRsMovCredPer.EOF Then   'se já houver algum registro
      Do While objRsMovCredPer.Fields("CONTACREDITOID").Value & "" = objRs.Fields("PKID").Value & ""
        If Not IsDate(objRs.Fields("DTSALDO").Value) Then
          curMovCredPer = curMovCredPer + objRsMovCredPer.Fields("VALOR").Value
          'Popular matriz de movimentação
        Else
          If IsDate(objRsMovCredPer.Fields("DATA").Value) Then
            If objRsMovCredPer.Fields("DATA").Value >= objRs.Fields("DTSALDO").Value Then
              curMovCredPer = curMovCredPer + objRsMovCredPer.Fields("VALOR").Value
              'Popular matriz de movimentação
            End If
          End If
        End If
        objRsMovCredPer.MoveNext
        If objRsMovCredPer.EOF Then Exit Do
      Loop
    End If
    RetornarSaldo = curVrSaldoTotal - curMovDebPer + curMovCredPer
  End If
  Set clsGer = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Function SelecionarConta(lngCONTAID As Long) As ADODB.Recordset
  On Error GoTo trata
  Dim objConta As datApler.clsConta
  Set objConta = New datApler.clsConta
  
  'Listar Conta
  Set SelecionarConta = objConta.SelecionarConta(lngCONTAID)
  '
  Set objConta = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Function

Public Sub IncluirConta(strValor As String, _
                        strData As String, _
                        strDescricao As String, _
                        lngTIPOCONTAID As Long, _
                        lngParceiroId As Long)
  
  On Error GoTo trata
  Dim objConta As datApler.clsConta
  Set objConta = New datApler.clsConta
  
  'Inserir Conta
  objConta.IncluirConta strValor, _
                        strData, _
                        strDescricao, _
                        lngTIPOCONTAID, _
                        lngParceiroId
  '
  Set objConta = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub

Public Sub AlterarConta(strValor As String, _
                        strData As String, _
                        strDescricao As String, _
                        lngCONTAID As Long, _
                        lngTIPOCONTAID As Long)
  On Error GoTo trata
  Dim objConta As datApler.clsConta
  Set objConta = New datApler.clsConta
  
  'Alterar Conta
  objConta.AlterarConta strValor, _
                        strData, _
                        strDescricao, _
                        lngCONTAID, _
                        lngTIPOCONTAID

  '
  Set objConta = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub
Public Sub ExcluirConta(lngCONTAID As Long)
  On Error GoTo trata
  Dim objConta As datApler.clsConta
  Set objConta = New datApler.clsConta
  
  'Excluir Conta
  objConta.ExcluirConta lngCONTAID

  '
  Set objConta = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, Err.Description, Err.Source
End Sub


