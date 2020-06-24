Attribute VB_Name = "mdlUserFunctions"
Option Explicit
'''
'''Public Function RetornarSaldo(strDtInicial As String, _
'''                              strDtFinal As String, _
'''                              lngCONTAID As Long)
'''  '
'''  Dim strSql          As String
'''  Dim objRs           As ADODB.Recordset
'''  Dim objRsMovDebAnt  As ADODB.Recordset
'''  Dim objRsMovCredAnt As ADODB.Recordset
'''  Dim objRsMovDebPer  As ADODB.Recordset
'''  Dim objRsMovCredPer As ADODB.Recordset
'''
'''  Dim clsGer          As busSisContas.clsGeral
'''  '
'''  Dim strDescricao          As String
'''  Dim curVrSaldo            As Currency
'''  Dim curVrSaldoPeriodo     As Currency
'''  Dim curVrSaldoTotal       As Currency
'''  Dim curVrSaldoTotalGeral  As Currency
'''  Dim datDataInicial        As Date
'''  '
'''  Dim curMovDebAnt    As Currency
'''  Dim curMovCredAnt   As Currency
'''  Dim curMovDebPer    As Currency
'''  Dim curMovCredPer   As Currency
'''  '
'''  On Error GoTo trata
'''  AmpS
'''  '
'''  Set clsGer = New busSisContas.clsGeral
'''  '
'''  'Inicia Data Inicial
'''  datDataInicial = CDate(Right(strDtInicial, 4) & "/" & Mid(strDtInicial, 4, 2) & "/" & Mid(strDtInicial, 1, 2))
'''  '
'''  strSql = "Select * " & _
'''    "From CONTA " & _
'''    "WHERE CONTA.PKID = " & lngCONTAID & _
'''    " ORDER BY CONTA.PKID;"
'''  '
'''  Set objRs = clsGer.ExecutarSQL(strSql)
'''  'objRs.Filter = strWhere
'''  '
'''  'Débito no periodo
'''  strSql = "Select MOVIMENTACAO.* " & _
'''            "FROM MOVIMENTACAO " & _
'''            "Where " & _
'''            " MOVIMENTACAO.DATA >= " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
'''            " AND MOVIMENTACAO.DATA <= " & Formata_Dados(strDtFinal, tpDados_DataHora, tpNulo_NaoAceita) & _
'''            " AND CONTADEBITOID <> NULL " & _
'''            " AND CONTADEBITOID = " & lngCONTAID & _
'''            " Order By CONTADEBITOID"
'''  '
'''  Set objRsMovDebPer = clsGer.ExecutarSQL(strSql)
'''  '
'''  'Crédito no periodo
'''  strSql = "Select MOVIMENTACAO.* " & _
'''            "FROM MOVIMENTACAO " & _
'''            "Where " & _
'''            " MOVIMENTACAO.DATA >= " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
'''            " AND MOVIMENTACAO.DATA <= " & Formata_Dados(strDtFinal, tpDados_DataHora, tpNulo_NaoAceita) & _
'''            " AND CONTACREDITOID <> NULL " & _
'''            " AND CONTACREDITOID = " & lngCONTAID & _
'''            " Order By CONTACREDITOID"
'''  '
'''  Set objRsMovCredPer = clsGer.ExecutarSQL(strSql)
'''  'Débito anterior
'''  strSql = "Select MOVIMENTACAO.* " & _
'''            "FROM MOVIMENTACAO " & _
'''            "Where " & _
'''            " MOVIMENTACAO.DATA < " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
'''            " AND CONTADEBITOID <> NULL " & _
'''            " AND CONTADEBITOID = " & lngCONTAID & _
'''            " Order By CONTADEBITOID"
'''  '
'''  Set objRsMovDebAnt = clsGer.ExecutarSQL(strSql)
'''  'Crédito anterior
'''  strSql = "Select MOVIMENTACAO.* " & _
'''            "FROM MOVIMENTACAO " & _
'''            "Where " & _
'''            " MOVIMENTACAO.DATA < " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
'''            " AND CONTACREDITOID <> NULL " & _
'''            " AND CONTACREDITOID = " & lngCONTAID & _
'''            " Order By CONTACREDITOID"
'''  '
'''  Set objRsMovCredAnt = clsGer.ExecutarSQL(strSql)
'''  '
'''  RetornarSaldo = 0
'''  If Not objRs.EOF Then   'se já houver algum item
'''    'Pegar Descrição e saldo da conta
'''    strDescricao = objRs.Fields("DESCRICAO").Value
'''    curVrSaldoTotal = 0
'''    curVrSaldo = 0
'''    curVrSaldoPeriodo = 0
'''    If Not IsDate(objRs.Fields("DTSALDO").Value) Then
'''      curVrSaldo = 0
'''    ElseIf objRs.Fields("DTSALDO").Value < datDataInicial Then
'''      curVrSaldo = IIf(Not IsNumeric(objRs.Fields("VRSALDO").Value), 0, objRs.Fields("VRSALDO").Value)
'''    Else
'''      curVrSaldoPeriodo = IIf(Not IsNumeric(objRs.Fields("VRSALDO").Value), 0, objRs.Fields("VRSALDO").Value)
'''    End If
'''    'Pegar Movimentação anterior Débito
'''    curMovDebAnt = 0
'''    If Not objRsMovDebAnt.EOF Then   'se já houver algum registro
'''      Do While objRsMovDebAnt.Fields("CONTADEBITOID").Value & "" = objRs.Fields("PKID").Value & ""
'''        If Not IsDate(objRs.Fields("DTSALDO").Value) Then
'''          curMovDebAnt = curMovDebAnt + objRsMovDebAnt.Fields("VALOR").Value
'''        ElseIf objRs.Fields("DTSALDO").Value < datDataInicial Then
'''          If IsDate(objRsMovDebAnt.Fields("DATA").Value) Then
'''            If objRsMovDebAnt.Fields("DATA").Value >= objRs.Fields("DTSALDO").Value Then
'''              curMovDebAnt = curMovDebAnt + objRsMovDebAnt.Fields("VALOR").Value
'''            End If
'''          End If
'''        End If
'''        objRsMovDebAnt.MoveNext
'''        If objRsMovDebAnt.EOF Then Exit Do
'''      Loop
'''    End If
'''    'Pegar Movimentação anterior Crédito
'''    curMovCredAnt = 0
'''    If Not objRsMovCredAnt.EOF Then   'se já houver algum registro
'''      Do While objRsMovCredAnt.Fields("CONTACREDITOID").Value & "" = objRs.Fields("PKID").Value & ""
'''        If Not IsDate(objRs.Fields("DTSALDO").Value) Then
'''          curMovCredAnt = curMovCredAnt + objRsMovCredAnt.Fields("VALOR").Value
'''        ElseIf objRs.Fields("DTSALDO").Value < datDataInicial Then
'''          If IsDate(objRsMovCredAnt.Fields("DATA").Value) Then
'''            If objRsMovCredAnt.Fields("DATA").Value >= objRs.Fields("DTSALDO").Value Then
'''              curMovCredAnt = curMovCredAnt + objRsMovCredAnt.Fields("VALOR").Value
'''            End If
'''          End If
'''        End If
'''        objRsMovCredAnt.MoveNext
'''        If objRsMovCredAnt.EOF Then Exit Do
'''      Loop
'''    End If
'''    'Valor saldo da movimentada
'''    curVrSaldo = curVrSaldo - curMovDebAnt + curMovCredAnt
'''    'If strpMostrarApenasSaldo = "N" Then
'''      'Valor total saldo da movimentada
'''      curVrSaldoTotal = curVrSaldo
'''    'End If
'''    curVrSaldoTotal = curVrSaldoTotal + curVrSaldoPeriodo
'''    'Pegar Movimentação no período Débito
'''    curMovDebPer = 0
'''    If Not objRsMovDebPer.EOF Then   'se já houver algum registro
'''      Do While objRsMovDebPer.Fields("CONTADEBITOID").Value & "" = objRs.Fields("PKID").Value & ""
'''        If Not IsDate(objRs.Fields("DTSALDO").Value) Then
'''          curMovDebPer = curMovDebPer + objRsMovDebPer.Fields("VALOR").Value
'''          'Popular matriz de movimentação
'''        Else
'''          If IsDate(objRsMovDebPer.Fields("DATA").Value) Then
'''            If objRsMovDebPer.Fields("DATA").Value >= objRs.Fields("DTSALDO").Value Then
'''              curMovDebPer = curMovDebPer + objRsMovDebPer.Fields("VALOR").Value
'''              'Popular matriz de movimentação
'''            End If
'''          End If
'''        End If
'''        objRsMovDebPer.MoveNext
'''        If objRsMovDebPer.EOF Then Exit Do
'''      Loop
'''    End If
'''    'Pegar Movimentação no período Crédito
'''    curMovCredPer = 0
'''    If Not objRsMovCredPer.EOF Then   'se já houver algum registro
'''      Do While objRsMovCredPer.Fields("CONTACREDITOID").Value & "" = objRs.Fields("PKID").Value & ""
'''        If Not IsDate(objRs.Fields("DTSALDO").Value) Then
'''          curMovCredPer = curMovCredPer + objRsMovCredPer.Fields("VALOR").Value
'''          'Popular matriz de movimentação
'''        Else
'''          If IsDate(objRsMovCredPer.Fields("DATA").Value) Then
'''            If objRsMovCredPer.Fields("DATA").Value >= objRs.Fields("DTSALDO").Value Then
'''              curMovCredPer = curMovCredPer + objRsMovCredPer.Fields("VALOR").Value
'''              'Popular matriz de movimentação
'''            End If
'''          End If
'''        End If
'''        objRsMovCredPer.MoveNext
'''        If objRsMovCredPer.EOF Then Exit Do
'''      Loop
'''    End If
'''    RetornarSaldo = curVrSaldo - curMovDebPer + curMovCredPer
'''  End If
'''  Set clsGer = Nothing
'''  AmpN
'''  Exit Function
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Function

'Propósito: Retornar e Gravar o Sequencial geral
Public Function RetornaGravaSequencial(pCampo As String) As Long
  On Error GoTo trata
  '
  Dim objGeral  As busSisContas.clsGeral
  Dim objRs     As ADODB.Recordset
  Dim strSql    As String
  Dim lngSeq    As Long
  '
  Set objGeral = New busSisContas.clsGeral
  strSql = "SELECT IIF(ISNUMERIC(MAX(" & pCampo & ") + 1), MAX(" & pCampo & ") + 1, 1) AS SEQ " & _
    "From SEQUENCIAL;"
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If objRs.EOF Then
    lngSeq = 1
  ElseIf Not IsNumeric(objRs.Fields("SEQ").Value) Then
    lngSeq = 1
  Else
    lngSeq = objRs.Fields("SEQ").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'Caso tenha atingido o limite do sequencial, voltar a 1
  If lngSeq > 9999 Then
    lngSeq = 1
  End If
  
  strSql = "Select Count(*) as total from Sequencial"
  Set objRs = objGeral.ExecutarSQL(strSql)
  If objRs.Fields("total").Value = 0 Then
    strSql = "INSERT INTO SEQUENCIAL (" & pCampo & ") VALUES (" & lngSeq & ")"
  Else
    strSql = "UPDATE SEQUENCIAL SET " & pCampo & " = " & lngSeq
  End If
  
  '
  objRs.Close
  Set objRs = Nothing
  '---
  objGeral.ExecutarSQLAtualizacao (strSql)
  '
  RetornaGravaSequencial = CLng(lngSeq)
  '
  Set objGeral = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, "[mdlUserFunction.RetornaGravaSequencial]", Err.Description
End Function

Public Sub Ordenar_Matriz(ByRef Matriz_a_ser_ordenada() As String, _
                          ByRef Matriz_sem_ordem() As String, _
                          ByVal QtdLinhas As Long, _
                          ByVal Col_a_ser_ordenada As Integer, _
                          ByVal TipoOrdem As String, _
                          ByVal TipoDados As ADODB.DataTypeEnum)
  On Error GoTo trata
  Dim lngLinha    As Long
  Dim lngLinhaInt As Long
  Dim intColuna   As Integer
  Dim strColunaAux  As String
  
  Dim Valor_coluna_a_ser_ordenada
  Dim Valor_coluna_sem_ordem
  '
  
  'PASSO 1 - IGUALA OS VETORES
  For lngLinha = 0 To QtdLinhas - 1
    'Para cada linha
    For intColuna = 0 To UBound(Matriz_sem_ordem)
      'Para cada Coluna
      Matriz_a_ser_ordenada(intColuna, lngLinha) = Matriz_sem_ordem(intColuna, lngLinha)
    Next
  Next
  'PASSO 2 - PERCORRE VETOR A SER ORDENADO
  For lngLinha = 0 To QtdLinhas - 1
    'PASSO 3 - PARA CADA LINHA DO VETOR A SER ORDENADO,
    'PERCORRE VETOR SEM ORDEM NA PRÓXIMA LINHA ATÉ O FINAL
    For lngLinhaInt = (lngLinha + 1) To QtdLinhas - 1
      'PASSO 4 - COMPARA COM A COLUNA A SER ORDENADA
      'PASSO 5 - TROCA
      If TipoDados = ADODB.DataTypeEnum.adInteger Then
        Valor_coluna_a_ser_ordenada = CLng(IIf(IsNumeric(Matriz_a_ser_ordenada(Col_a_ser_ordenada - 1, lngLinha)), Matriz_a_ser_ordenada(Col_a_ser_ordenada - 1, lngLinha), 0))
        Valor_coluna_sem_ordem = CLng(IIf(IsNumeric(Matriz_sem_ordem(Col_a_ser_ordenada - 1, lngLinhaInt)), Matriz_sem_ordem(Col_a_ser_ordenada - 1, lngLinhaInt), 0))
      ElseIf TipoDados = ADODB.DataTypeEnum.adVarChar Then
        Valor_coluna_a_ser_ordenada = Matriz_a_ser_ordenada(Col_a_ser_ordenada - 1, lngLinha)
        Valor_coluna_sem_ordem = Matriz_sem_ordem(Col_a_ser_ordenada - 1, lngLinhaInt)
      ElseIf TipoDados = ADODB.DataTypeEnum.adDate Then
        Valor_coluna_a_ser_ordenada = CDate(Right(Matriz_a_ser_ordenada(Col_a_ser_ordenada - 1, lngLinha), 4) & "/" & Mid(Matriz_a_ser_ordenada(Col_a_ser_ordenada - 1, lngLinha), 4, 2) & "/" & Left(Matriz_a_ser_ordenada(Col_a_ser_ordenada - 1, lngLinha), 2))
        Valor_coluna_sem_ordem = CDate(Right(Matriz_sem_ordem(Col_a_ser_ordenada - 1, lngLinhaInt), 4) & "/" & Mid(Matriz_sem_ordem(Col_a_ser_ordenada - 1, lngLinhaInt), 4, 2) & "/" & Left(Matriz_sem_ordem(Col_a_ser_ordenada - 1, lngLinhaInt), 2))
      End If

      
      If IIf(TipoOrdem = "Descendente", Valor_coluna_sem_ordem > Valor_coluna_a_ser_ordenada, Valor_coluna_sem_ordem < Valor_coluna_a_ser_ordenada) Then
        For intColuna = 0 To UBound(Matriz_sem_ordem)
          'Para cada Coluna
          strColunaAux = Matriz_a_ser_ordenada(intColuna, lngLinha)
          Matriz_a_ser_ordenada(intColuna, lngLinha) = Matriz_sem_ordem(intColuna, lngLinhaInt)
          Matriz_sem_ordem(intColuna, lngLinhaInt) = strColunaAux
          Matriz_a_ser_ordenada(intColuna, lngLinhaInt) = strColunaAux
        Next
      End If
    Next
  Next
  Exit Sub
trata:
  Err.Raise Err.Number, "[mdlUserFunction.Ordenar_Matriz]", Err.Description
End Sub

'Proposito: Dado um mês,
'Retorna mês por extenso
Public Function Retorna_descr_mes(intMes As Integer) As String
  On Error GoTo trata
  Select Case intMes
  Case 1: Retorna_descr_mes = "Janeiro"
  Case 2: Retorna_descr_mes = "Fevereiro"
  Case 3: Retorna_descr_mes = "Março"
  Case 4: Retorna_descr_mes = "Abril"
  Case 5: Retorna_descr_mes = "Maio"
  Case 6: Retorna_descr_mes = "Junho"
  Case 7: Retorna_descr_mes = "Julho"
  Case 8: Retorna_descr_mes = "Agosto"
  Case 9: Retorna_descr_mes = "Setembro"
  Case 10: Retorna_descr_mes = "Outubro"
  Case 11: Retorna_descr_mes = "Novembro"
  Case 12: Retorna_descr_mes = "Dezembro"
  Case Else: Retorna_descr_mes = ""
  End Select
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[mdlUserFunctions.Retorna_descr_mes]"
End Function
'Proposito: Retornar último dia do mês
'Retorna DD
Public Function Retorna_ultimo_dia_do_mes(intMes As Integer, _
                                          intAno As Integer) As String
  On Error GoTo trata
  Dim intDia    As Integer
  Dim dtaTeste  As Date
  For intDia = 29 To 32
    dtaTeste = CDate(intAno & "/" & intMes & "/" & intDia)
  Next
  Exit Function
trata:
  Retorna_ultimo_dia_do_mes = Format(intDia - 1, "00")
End Function
'Proposito: Retornar o mês anterior
'Retorna MM/YYYY
Public Function Retorna_mes_ano_anterior(intMes As Integer, _
                                         intAno As Integer) As String
  On Error GoTo trata
  If intMes = 1 Then
    Retorna_mes_ano_anterior = Format(12, "00") & "/" & Format(intAno - 1, "0000")
  Else
    Retorna_mes_ano_anterior = Format(intMes - 1, "00") & "/" & Format(intAno, "0000")
  End If
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[mdlUserFunctions.Retorna_mes_ano_anterior]"
End Function
'Proposito: Retornar o Código do Turno Corrente
Public Function RetornaCodTurnoCorrente() As Long
  On Error GoTo trata
  'Retorna 0 - para Código de Erro
  Dim strSql      As String
  Dim objRs       As ADODB.Recordset
  Dim lngRetorno  As Long
  Dim objGeral    As busSisContas.clsGeral
  '
  Set objGeral = New busSisContas.clsGeral
  '
  strSql = "Select * from Turno Where Status = true"
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If objRs.EOF Then
    'Não há turno corrente cadastrado
    lngRetorno = 0
  ElseIf objRs.RecordCount > 1 Then
    'há mais de um turno corrente cadastrado
    lngRetorno = 0
  Else
    lngRetorno = objRs!PKID
  End If
  objRs.Close
  Set objRs = Nothing
  '
  RetornaCodTurnoCorrente = lngRetorno
  Set objGeral = Nothing
  Exit Function
trata:
  Err.Raise Err.Number, Err.Description, "[mdlFunction.RetornaCodTurnoCorrente]"
End Function

'Propósito: criptografar a senha do usuário armazenada no banco de dados
'Entrada: senha
'Retorna: senha
          'caso entrada seja não criptografada a saída é criptografada e vice-versa

Public Function Encripta(Senha As String) As String
Dim I As Integer
Dim str As String
  For I = 1 To Len(Senha)
    str = Mid(Senha, I, 1)
    str = 255 - Asc(str)
    Encripta = Encripta & Chr(str)
  Next I
End Function

'Propósito Abrir Registros do Sistema
'recebe parametro pAcao que assume
'0 - Captura parametros inicias
'1 -  grava últ usuário que
'2 - Grava BMP
'acessou o sistema
'Caso algum parametro seja nulo, regrava no Registro
Public Function CapturaParametrosRegistro(pAcao As Integer)
  Dim iLenCaminho, iLenArquivo As Long
  Dim iRet
  On Error GoTo RotErro
Repeticao:
  AmpS
  Select Case pAcao
  Case 0
    'Captura Banco de Dados
    gsBDadosPath = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="CaminhoDB")
    If Len(Trim(gsBDadosPath)) = 0 Then
      'Registro não esta gravado está em branco
      iRet = MsgBox("Não foi possível localizar o banco de dados " & _
        nomeBDados & ". Você deseja localizá-lo manualmente?", _
        vbQuestion + vbYesNo, TITULOSISTEMA)
      If iRet = vbYes Then
        frmMDI.CommonDialog1.ShowOpen
        iLenCaminho = Len(frmMDI.CommonDialog1.FileName)
        iLenArquivo = Len(frmMDI.CommonDialog1.FileTitle)
        gsBDadosPath = Left(frmMDI.CommonDialog1.FileName, iLenCaminho - iLenArquivo)
        SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                  Key:="CaminhoDB", setting:=gsBDadosPath
        GoTo Repeticao 'Captura Novamente os parametros até encontrar
                        'Parametros válidos
      Else
        End
      End If
    End If
    'Captura Caminho Crystal
    gsReportPath = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="CaminhoReport")
    If Len(Trim(gsReportPath)) = 0 Then
      'Registro não esta gravado está em branco
      iRet = MsgBox("Não foi possível localizar o caminho dos Formulários (.RPT)." & _
        " Você deseja localizá-lo manualmente?", _
        vbQuestion + vbYesNo, TITULOSISTEMA)
      If iRet = vbYes Then
        frmMDI.CommonDialog1.ShowOpen
        iLenCaminho = Len(frmMDI.CommonDialog1.FileName)
        iLenArquivo = Len(frmMDI.CommonDialog1.FileTitle)
        gsReportPath = Left(frmMDI.CommonDialog1.FileName, iLenCaminho - iLenArquivo)
        SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                  Key:="CaminhoReport", setting:=gsReportPath
        GoTo Repeticao 'Captura Novamente os parametros até encontrar
                        'Parametros válidos
      Else
        End
      End If
    End If
    'Captura Caminho App
    gsAppPath = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="CaminhoApp")
    If Len(Trim(gsAppPath)) = 0 Then
      'Registro não esta gravado está em branco
      iRet = MsgBox("Não foi possível localizar o caminho do Aplicativo (.EXE)." & _
        " Você deseja localizá-lo manualmente?", _
        vbQuestion + vbYesNo, TITULOSISTEMA)
      If iRet = vbYes Then
        frmMDI.CommonDialog1.ShowOpen
        iLenCaminho = Len(frmMDI.CommonDialog1.FileName)
        iLenArquivo = Len(frmMDI.CommonDialog1.FileTitle)
        gsAppPath = Left(frmMDI.CommonDialog1.FileName, iLenCaminho - iLenArquivo)
        SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                  Key:="CaminhoApp", setting:=gsAppPath
        GoTo Repeticao 'Captura Novamente os parametros até encontrar
                        'Parametros válidos
      Else
        End
      End If
    End If
    'Captura Nome do Usuário
    gsNomeUsu = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="Usuario")
    'Captura o Nome do Curso
    gsNomeEmpresa = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="Empresa")
    If Len(Trim(gsNomeEmpresa)) = 0 Then
      'Registro não esta gravado está em branco
      gsNomeEmpresa = "XXX"
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="Empresa", setting:=gsNomeEmpresa
      GoTo Repeticao 'Captura Novamente os parametros até encontrar
                      'Parametros válidos
    End If
    'Captura o Caminho dos bitmaps
    gsBMPPath = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="CaminhoBMP")
    If Len(Trim(gsBMPPath)) = 0 Then
      'Registro não esta gravado está em branco
      gsBMPPath = gsBDadosPath & "Images\BMP\"
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="CaminhoBMP", setting:=gsBMPPath
      GoTo Repeticao 'Captura Novamente os parametros até encontrar
                      'Parametros válidos
    End If
    'Captura o Caminho dos Icons
    gsIconsPath = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="CaminhoIcons")
    If Len(Trim(gsIconsPath)) = 0 Then
      'Registro não esta gravado está em branco
      gsIconsPath = gsBDadosPath & "Images\Icons\"
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="CaminhoIcons", setting:=gsIconsPath
      GoTo Repeticao 'Captura Novamente os parametros até encontrar
                      'Parametros válidos
    End If
    'Captura o Nome do BitMap
    gsBMP = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="BMP")
    'Captura o caminho do BackUp
    gsPathBackup = GetSetting(AppName:=TITULOSISTEMA, section:="Iniciar", _
                 Key:="CaminhoBackUp")
    If Len(Trim(gsPathBackup)) = 0 Then
      'Registro não esta gravado está em branco
      gsPathBackup = gsAppPath & "BackUp\"
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="CaminhoBackUp", setting:=gsPathBackup
      GoTo Repeticao 'Captura Novamente os parametros até encontrar
                      'Parametros válidos
    End If
  Case 1
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="Usuario", setting:=gsNomeUsu
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="Nivel", setting:=gsNivel
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="Ativo", setting:="S"

  Case 2
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="BMP", setting:=gsBMP
  Case 3
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="Nivel", setting:=""
      SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
                Key:="Ativo", setting:="N"
  
  End Select
  AmpN
  Exit Function
RotErro:
  AmpN
  frmMDI.CommonDialog1.ShowOpen
  iLenCaminho = Len(frmMDI.CommonDialog1.FileName)
  iLenArquivo = Len(frmMDI.CommonDialog1.FileTitle)
  gsBDadosPath = Left(frmMDI.CommonDialog1.FileName, iLenCaminho - iLenArquivo)
  SaveSetting AppName:=TITULOSISTEMA, section:="Iniciar", _
            Key:="CaminhoDB", setting:=gsBDadosPath
  
  AmpN
End Function


Public Function Formata_Dados(pValor As Variant, pTipoDados As TpTipoDados, pAceitaNulo As tpAceitaNulo, Optional pTamanhoCampo As Integer) As Variant
  On Error GoTo trata
  '
  Dim vRetorno As Variant
  Dim sData As String
  '
  Select Case pTipoDados
  Case TpTipoDados.tpDados_Boolean
    If pValor Then
      vRetorno = "True"
    Else
      vRetorno = "False"
    End If
  Case TpTipoDados.tpDados_Texto
    If Len(Trim(pValor & "")) = 0 Then
      If pAceitaNulo = tpNulo_Aceita Then
        vRetorno = "Null"
      Else
        vRetorno = "' '"
      End If
    Else
      vRetorno = "'" & Tira_Plic(Trim(pValor & "")) & "'"
    End If
  Case TpTipoDados.tpDados_Longo
    If Not IsNumeric(pValor) Then
      If pAceitaNulo = tpNulo_Aceita Then
        vRetorno = "Null"
      Else
        vRetorno = "0"
      End If
    Else
      vRetorno = CLng(pValor)
    End If
  Case TpTipoDados.tpDados_DataHora
    'Converter para Data
    sData = ""
    If Len(pValor & "") = 10 Then
      If Mid(pValor & "", 1, 2) <> "__" Then
        'Data no Formato DD/MM/YYYY
        sData = "#" & Mid(pValor, 4, 2) & "/" & Mid(pValor, 1, 2) & "/" & Mid(pValor, 7, 4) & "#"
      Else
        sData = "null"
      End If
    ElseIf Len(pValor & "") = 16 Then
      'Data no Formato DD/MM/YYYY hh:mm
      If Mid(pValor & "", 1, 2) <> "__" Then
        'Data no Formato DD/MM/YYYY
        sData = "#" & Mid(pValor, 4, 2) & "/" & Mid(pValor, 1, 2) & "/" & Mid(pValor, 7, 4) & " " & Mid(pValor, 12, 2) & ":" & Mid(pValor, 15, 2) & "#"
      Else
        sData = "null"
      End If
    Else
      sData = ""
    End If
    If Len(sData) = 0 Then
      If pAceitaNulo = tpNulo_Aceita Then
        vRetorno = "Null"
      Else
        vRetorno = "#01/01/1900#"
      End If
    Else
      vRetorno = sData
    End If
  Case TpTipoDados.tpDados_Moeda
    If Not IsNumeric(pValor) Then
      If pAceitaNulo = tpNulo_Aceita Then
        vRetorno = "Null"
      Else
        vRetorno = "0"
      End If
    Else
      vRetorno = Replace(pValor, ".", "")
      vRetorno = Replace(vRetorno, ",", ".")
    End If
  Case Else
  End Select
  '
  Formata_Dados = vRetorno
  '
  Exit Function
trata:
  
End Function
Public Function Tira_Plic(pValor As String) As Variant
  Tira_Plic = Replace(pValor, "'", "''")
End Function

Public Sub Main()
  On Error GoTo trata
  '
  frmSplash.Show
  frmSplash.Refresh
  '
  frmLogin.QuemChamou = 0
  Load frmLogin
  frmLogin.Show vbModal
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[mdlUserFunctions.Main]"
  End
End Sub

Public Sub TratarErro(ByVal pNumero As Long, _
                      ByVal pDescricao As String, _
                      ByVal pSource As String)
  '
  On Error Resume Next
  Dim strUsuario As String
  Dim intI       As Integer
    
  intI = FreeFile
  Open App.Path & "\SisContas.txt" For Append As #intI
  
  Print #intI, Format(Now(), "DD/MM/YYYY hh:mm") & ";" & pSource & ";" & pNumero & ";" & pDescricao
  Close #intI
  'mostrar Mensagem
  MsgBox "O Seginte Erro Ocorreu: " & vbCrLf & vbCrLf & _
    "Número: " & pNumero & vbCrLf & _
    "Descrição: " & pDescricao & vbCrLf & vbCrLf & _
    "Módulo: " & pSource & vbCrLf & _
    "Data/Hora: " & Format(Now(), "DD/MM/YYYY hh:mm") & vbCrLf & _
    "Erro gravado no arquivo: " & App.Path & "\SisContas.txt" & vbCrLf & vbCrLf & _
    "Caso o erro persista contacte o suporte e envie o arquivo acima, informando a data e hora acima informada da ocorrência deste erro.", vbCritical, TITULOSISTEMA
End Sub

Public Sub Ordenar_Matriz_Ncols(ByRef Matriz_a_ser_ordenada() As String, _
                                ByRef Matriz_sem_ordem() As String, _
                                ByVal QtdLinhas As Long, _
                                ByVal VetTipo_Coluna, _
                                ByVal VetTipo_Ordem, _
                                ByVal VetTipo_Dados)
  On Error GoTo trata
  Dim lngLinha    As Long
  Dim lngLinhaInt As Long
  Dim intColuna   As Integer
  Dim strColunaAux  As String
  Dim intColVet     As Integer
  
  Dim Valor_coluna_a_ser_ordenada
  Dim Valor_coluna_sem_ordem
  '
  'Dim Col_a_ser_ordenada
  'Dim TipoOrdem
  'Dim TipoDados
  'PASSO 1 - IGUALA OS VETORES
  For lngLinha = 0 To QtdLinhas - 1
    'Para cada linha
    For intColuna = 0 To UBound(Matriz_sem_ordem)
      'Para cada Coluna
      Matriz_a_ser_ordenada(intColuna, lngLinha) = Matriz_sem_ordem(intColuna, lngLinha)
    Next
  Next
  Dim blnTroca As Boolean
  Dim blnTroca1 As Boolean
  Dim blnTroca2 As Boolean
  'PASSO 2 - PERCORRE VETOR A SER ORDENADO
  For lngLinha = 0 To QtdLinhas - 1
    'PASSO 3 - PARA CADA LINHA DO VETOR A SER ORDENADO,
    'PERCORRE VETOR SEM ORDEM NA PRÓXIMA LINHA ATÉ O FINAL
    For lngLinhaInt = (lngLinha + 1) To QtdLinhas - 1
      'PARA CADA COLUNA A SER ORDENADA
      blnTroca = False
      blnTroca1 = False
      blnTroca2 = False
      
      For intColVet = 0 To UBound(VetTipo_Coluna)
        'PASSO 4 - COMPARA COM A COLUNA A SER ORDENADA
        If VetTipo_Dados(intColVet) = ADODB.DataTypeEnum.adInteger Then
          Valor_coluna_a_ser_ordenada = CLng(IIf(IsNumeric(Matriz_a_ser_ordenada(VetTipo_Coluna(intColVet), lngLinha)), Matriz_a_ser_ordenada(VetTipo_Coluna(intColVet), lngLinha), 0))
          Valor_coluna_sem_ordem = CLng(IIf(IsNumeric(Matriz_sem_ordem(VetTipo_Coluna(intColVet), lngLinhaInt)), Matriz_sem_ordem(VetTipo_Coluna(intColVet), lngLinhaInt), 0))
        ElseIf VetTipo_Dados(intColVet) = ADODB.DataTypeEnum.adDate Then
          Valor_coluna_a_ser_ordenada = CDate(Right(Matriz_a_ser_ordenada(VetTipo_Coluna(intColVet) - 1, lngLinha), 4) & "/" & Mid(Matriz_a_ser_ordenada(VetTipo_Coluna(intColVet) - 1, lngLinha), 4, 2) & "/" & Left(Matriz_a_ser_ordenada(VetTipo_Coluna(intColVet) - 1, lngLinha), 2))
          Valor_coluna_sem_ordem = CDate(Right(Matriz_sem_ordem(VetTipo_Coluna(intColVet) - 1, lngLinhaInt), 4) & "/" & Mid(Matriz_sem_ordem(VetTipo_Coluna(intColVet) - 1, lngLinhaInt), 4, 2) & "/" & Left(Matriz_sem_ordem(VetTipo_Coluna(intColVet) - 1, lngLinhaInt), 2))
          'Valor_coluna_a_ser_ordenada = Matriz_a_ser_ordenada(VetTipo_Coluna(intColVet), lngLinha)
          'Valor_coluna_sem_ordem = Matriz_sem_ordem(VetTipo_Coluna(intColVet), lngLinhaInt)
        ElseIf VetTipo_Dados(intColVet) = ADODB.DataTypeEnum.adVarChar Then
          Valor_coluna_a_ser_ordenada = Matriz_a_ser_ordenada(VetTipo_Coluna(intColVet), lngLinha)
          Valor_coluna_sem_ordem = Matriz_sem_ordem(VetTipo_Coluna(intColVet), lngLinhaInt)
        End If
        'PASSO 5 - TROCA
        If Valor_coluna_a_ser_ordenada = Valor_coluna_sem_ordem Then
          If intColVet = 0 Then
            blnTroca1 = True
          ElseIf intColVet = 1 Then
            blnTroca2 = True
          End If
        End If
        If IIf(VetTipo_Ordem(intColVet) = "Descendente", Valor_coluna_sem_ordem > Valor_coluna_a_ser_ordenada, Valor_coluna_sem_ordem < Valor_coluna_a_ser_ordenada) Then
          If intColVet = 0 Then
            blnTroca = True
          ElseIf intColVet = 1 And blnTroca1 Then
            blnTroca = True
          ElseIf intColVet = 2 And blnTroca1 And blnTroca2 Then
            blnTroca = True
          End If
          If blnTroca Then
            For intColuna = 0 To UBound(Matriz_sem_ordem)
              'Para cada Coluna
              strColunaAux = Matriz_a_ser_ordenada(intColuna, lngLinha)
              Matriz_a_ser_ordenada(intColuna, lngLinha) = Matriz_sem_ordem(intColuna, lngLinhaInt)
              Matriz_sem_ordem(intColuna, lngLinhaInt) = strColunaAux
              Matriz_a_ser_ordenada(intColuna, lngLinhaInt) = strColunaAux
            Next
            Exit For
          End If
        End If
      
      Next
    Next
  Next
  Exit Sub
trata:
  Err.Raise Err.Number, "[mdlUserFunction.Ordenar_Matriz]", Err.Description
End Sub

Public Sub TratarErroPrevisto(ByVal pDescricao As String, _
                              ByVal pSource As String)
  '
  On Error Resume Next
  'mostrar Mensagem
  MsgBox "Erro(s): " & vbCrLf & vbCrLf & _
    pDescricao '& vbCrLf & vbCrLf '& _
    '"Módulo: " & pSource & vbCrLf & vbCrLf & _
    '"Reavalie as informações e corrija os dados para que a alteração seja efetivada.", vbExclamation, TITULOSISTEMA
End Sub
Public Sub AmpS()
  Screen.MousePointer = vbHourglass
End Sub

Public Sub AmpN()
  Screen.MousePointer = vbDefault
End Sub

Public Sub Pintar_Controle(pControle As Variant, _
                            pCor As tpCorControle)
  On Error GoTo trata
  AmpS
  pControle.BackColor = pCor
  AmpN
  Exit Sub
trata:
  AmpN
  Err.Raise Err.Number, "[mdlUserFunction.Pintar_Controle]", Err.Description
End Sub

'Propósito: Centralizar um form MDI Child no form MDI.
'Entradas:  frmCenter - Form a centralizar

Public Sub CenterForm(frmCenter As Form)
   Dim intHeight      As Integer
   Dim intLeftOffset  As Integer
   Dim intTop         As Integer
   Dim intWidth       As Integer
   Dim intLeft        As Integer
   Dim intTopOffset   As Integer
   '
On Error GoTo trata
  AmpS
   If frmCenter.MDIChild = True Then
      intHeight = frmMDI.ScaleHeight
      intWidth = frmMDI.ScaleWidth
      intTop = frmMDI.Top
      intLeft = frmMDI.Left
   Else
      intHeight = Screen.Height
      intWidth = Screen.Width
      intTop = 0
      intLeft = 0
   End If

   'Calcula "left offset"
   intLeftOffset = ((intWidth - frmCenter.Width) / 2) + intLeft
   If (intLeftOffset + frmCenter.Width > Screen.Width) Or intLeftOffset < 100 Then
      intLeftOffset = 100
   End If

   'Calcula "top offset"
   intTopOffset = ((intHeight - frmCenter.Height) / 2) + intTop
   If (intTopOffset + frmCenter.Height > Screen.Height) Or intTopOffset < 100 Then
      intTopOffset = 100
   End If
   'Centraliza o form
  frmCenter.Move intLeftOffset, intTopOffset
  AmpN
  Exit Sub
trata:
  AmpN
  Err.Raise Err.Number, "[mdlUserFunction.CenterForm]", Err.Description
End Sub

Public Sub LerFiguras(pForm As Form, pOp As tpBmpForm, Optional pbtnOK As CommandButton, Optional pbtnCancelar As CommandButton, Optional pbtnFechar As CommandButton, Optional pbtnExcluir As CommandButton, Optional pbtnSenha As CommandButton, Optional pbtnIncluir As CommandButton, Optional pbtnAlterar As CommandButton, Optional pbtnFiltrar As CommandButton, Optional pbtnImprimir As CommandButton, Optional pbtnConsultar As CommandButton)
  On Error GoTo trata
    
  If pOp = tpBmpForm.tpBmp_Login Then
    pForm.Picture = LoadPicture(gsBMPPath & "Fundo.jpg")
  ElseIf pOp = tpBmpForm.tpBmp_MDI Then
    pForm.Picture = LoadPicture(gsBMPPath & "Fundo.jpg")
    pForm.Icon = LoadPicture(gsIconsPath & "Logo.ico")
  Else
    pForm.Icon = LoadPicture(gsIconsPath & "areatrab.ico")
  End If
  
  If Not (pbtnConsultar Is Nothing) Then
    pbtnConsultar.Picture = LoadPicture(gsIconsPath & "Procurar.ico")
    pbtnConsultar.DownPicture = LoadPicture(gsIconsPath & "ProcurarDown.ico")
    pbtnConsultar.ToolTipText = "Consultar"
  End If
  If Not (pbtnImprimir Is Nothing) Then
    pbtnImprimir.Picture = LoadPicture(gsIconsPath & "Impressora.ico")
    pbtnImprimir.DownPicture = LoadPicture(gsIconsPath & "ImpressoraDown.ico")
    pbtnImprimir.ToolTipText = "Imprimir"
  End If
  If Not (pbtnOK Is Nothing) Then
    pbtnOK.Picture = LoadPicture(gsIconsPath & "Ok.ico")
    pbtnOK.DownPicture = LoadPicture(gsIconsPath & "OkDown.ico")
    pbtnOK.ToolTipText = "Ok"
  End If
  If Not (pbtnAlterar Is Nothing) Then
    pbtnAlterar.Picture = LoadPicture(gsIconsPath & "Alterar.ico")
    pbtnAlterar.DownPicture = LoadPicture(gsIconsPath & "AlterarDown.ico")
    pbtnAlterar.ToolTipText = "Alterar"
  End If
  
  If Not (pbtnIncluir Is Nothing) Then
    pbtnIncluir.Picture = LoadPicture(gsIconsPath & "Incluir.ico")
    pbtnIncluir.DownPicture = LoadPicture(gsIconsPath & "IncluirDown.ico")
    pbtnIncluir.ToolTipText = "Incluir"
  End If
  
  If Not (pbtnCancelar Is Nothing) Then
    pbtnCancelar.Picture = LoadPicture(gsIconsPath & "Cancelar.ico")
    pbtnCancelar.DownPicture = LoadPicture(gsIconsPath & "CancelarDown.ico")
    pbtnCancelar.ToolTipText = "Cancelar"
  End If
  If Not (pbtnExcluir Is Nothing) Then
    pbtnExcluir.Picture = LoadPicture(gsIconsPath & "Excluir.ico")
    pbtnExcluir.DownPicture = LoadPicture(gsIconsPath & "ExcluirDown.ico")
    pbtnExcluir.ToolTipText = "Excluir"
  End If
  If Not (pbtnSenha Is Nothing) Then
    pbtnSenha.Picture = LoadPicture(gsIconsPath & "Senha.ico")
    pbtnSenha.DownPicture = LoadPicture(gsIconsPath & "SenhaDown.ico")
    pbtnSenha.ToolTipText = "Senha"
  End If
  If Not (pbtnFechar Is Nothing) Then
    pbtnFechar.Picture = LoadPicture(gsIconsPath & "Sair.ico")
    pbtnFechar.DownPicture = LoadPicture(gsIconsPath & "SairDown.ico")
    pbtnFechar.ToolTipText = "Sair"
  End If
  If Not (pbtnFiltrar Is Nothing) Then
    pbtnFiltrar.Picture = LoadPicture(gsIconsPath & "Filtrar.ico")
    pbtnFiltrar.DownPicture = LoadPicture(gsIconsPath & "FiltrarDown.ico")
    pbtnFiltrar.ToolTipText = "Aplicar Filtro"
  End If
  '
  Exit Sub
trata:
  AmpN
  Err.Raise Err.Number, "[mdlUserFunction.LerFiguras]", Err.Description
End Sub

Public Sub LerFigurasAvulsas(pbtn As CommandButton, pImagem As String, pImagemDown As String, pToolTipText As String)
  On Error GoTo trata
  '
  pbtn.Picture = LoadPicture(gsIconsPath & pImagem)
  pbtn.DownPicture = LoadPicture(gsIconsPath & pImagemDown)
  pbtn.ToolTipText = pToolTipText
  '
  Exit Sub
trata:
  AmpN
  Err.Raise Err.Number, "[mdlUserFunction.LerFigurasAvulsas]", Err.Description
End Sub

Public Sub LerFigurasAvulsasPicBox(ppicBox As PictureBox, pImagem As String, pToolTipText As String)
  On Error GoTo trata
  '
  ppicBox.Picture = LoadPicture(gsIconsPath & pImagem)
  ppicBox.ToolTipText = pToolTipText
  '
  Exit Sub
trata:
  AmpN
  Err.Raise Err.Number, "[mdlUserFunction.LerFigurasAvulsasPicBox]", Err.Description
End Sub

Public Sub Selecionar_Conteudo(pControle As Variant)
  On Error GoTo trata
  AmpS
  pControle.SelStart = 0
  pControle.SelLength = Len(pControle)
  AmpN
  Exit Sub
trata:
  AmpN
  Err.Raise Err.Number, "[mdlUserFunction.Selecionar_Conteudo]", Err.Description
End Sub


Public Function MakeBookmarkGeral(Index As Long) As Variant
  ' This support function is used only by the remaining
  ' support functions. It is not used directly by the
  ' unbound events. It is a good idea to create a
  ' MakeBookmark function such that all bookmarks can be
  ' created in the same way. Thus the method by which
  ' bookmarks are created is consistent and easy to
  ' modify. This function creates a bookmark when given
  ' an array row index.
  ' Since we have data stored in an array, we will just
  ' use the array index as our bookmark. We will convert
  ' it to a string first, using the CStr function.
  On Error GoTo trata
  MakeBookmarkGeral = CStr(Index)
  Exit Function
trata:
  Err.Raise Err.Number, "[mdlUserFunction.MakeBookmarkGeral]", Err.Description
End Function

Public Function IndexFromBookmarkGeral(Bookmark As Variant, _
        Offset As Long, intLINHASMATRIZ As Long) As Long
  ' This support function is used only by the remaining
  ' support functions. It is not used directly by the
  ' unbound events.
  ' IndexFromBookmark computes the row index that
  ' corresponds to a row that is (Offset) rows from the
  ' row specified by the Bookmark parameter. For example,
  ' if Bookmark refers to the index 50 of the dataset
  ' array and Offset = -10, then IndexFromBookmark will
  ' return 50 + (-10), or 40. Thus to get the index of
  ' the row specified by the bookmark itself, call
  ' IndexFromBookmark with an Offset of 0. If the given
  ' Bookmark is Null, it refers to BOF or EOF. If
  ' Offset < 0, the grid is requesting rows before the
  ' row specified by Bookmark, and so we must be at EOF
  ' because prior rows do not exist for BOF. Conversely,
  ' if Offset > 0, we are at BOF.
  On Error GoTo trata
  Dim Index As Long
  If IsNull(Bookmark) Then
    If Offset < 0 Then
      ' Bookmark refers to EOF. Since (MaxRow - 1)
      ' is the index of the last record, we can use
      ' an index of (MaxRow) to represent EOF.
      Index = intLINHASMATRIZ + Offset
    Else
      ' Bookmark refers to BOF. Since 0 is the index
      ' of the first record, we can use an index of
      ' -1 to represent BOF.
      Index = -1 + Offset
    End If
  Else
    ' Convert string to long integer
    Index = Val(Bookmark) + Offset
  End If
    
  ' Check to see if the row index is valid:
  '   (0 <= Index < MaxRow).
  ' If not, set it to a large negative number to
  ' indicate that it is invalid.
  If Index >= 0 And Index < intLINHASMATRIZ Then
    IndexFromBookmarkGeral = Index
  Else
    IndexFromBookmarkGeral = -9999
  End If
  '
  Exit Function
trata:
  Err.Raise Err.Number, "[mdlUserFunction.IndexFromBookmarkGeral]", Err.Description
End Function

Public Function GetRelativeBookmarkGeral(Bookmark As Variant, _
        Offset As Long, intLINHASMATRIZ As Long) As Variant
  ' GetRelativeBookmark is used to get a bookmark for a
  ' row that is a specified number of rows away from the
  ' given row. Offset specifies the number of rows to
  ' move. A positive Offset indicates that the desired
  ' row is after the one referred to by Bookmark, and a
  ' negative Offset means it is before the one referred
  ' to by Bookmark.
  On Error GoTo trata
  Dim Index As Long
    
  ' Compute the row index for the desired row
  Index = IndexFromBookmarkGeral(Bookmark, Offset, intLINHASMATRIZ)
  If Index < 0 Or Index >= intLINHASMATRIZ Then
    ' Index refers to a row before the first or after
    ' the last, so just return Null.
    GetRelativeBookmarkGeral = Null
  Else
    GetRelativeBookmarkGeral = MakeBookmarkGeral(Index)
  End If
  Exit Function
trata:
  Err.Raise Err.Number, "[mdlUserFunction.GetRelativeBookmarkGeral]", Err.Description
End Function

Public Function GetUserDataGeral(Bookmark As Variant, _
        Col As Integer, intCOLUNASMATRIZ As Long, intLINHASMATRIZ As Long, mtzMatriz) As Variant
  ' In this example, GetUserData is called by
  ' UnboundReadData to ask the user what data should be
  ' displayed in a specific cell in the grid. The grid
  ' row the cell is in is the one referred to by the
  ' Bookmark parameter, and the column it is in it given
  ' by the Col parameter. GetUserData is called on a
  ' cell-by-cell basis.
  
  On Error GoTo trata
  '
  Dim Index As Long

  ' Figure out which row the bookmark refers to
  Index = IndexFromBookmarkGeral(Bookmark, 0, intLINHASMATRIZ)
  If Index < 0 Or Index >= intLINHASMATRIZ Or _
      Col < 0 Or Col >= intCOLUNASMATRIZ Then
    ' Cell position is invalid, so just return null to
    ' indicate failure
    GetUserDataGeral = Null
  Else
    GetUserDataGeral = mtzMatriz(Col, Index)
  End If
  Exit Function
trata:
  Err.Raise Err.Number, "[mdlUserFunction.GetUserDataGeral]", Err.Description
End Function
Public Sub SetarFoco(objTarget As Object)
  On Error Resume Next
  If objTarget.Enabled = True Then objTarget.SetFocus
End Sub
Public Function Valida_Data(pMsk As MaskEdBox, pTipo As TpObriga) As Boolean
  Dim EData As Boolean
  EData = True
  'Verifica se Mmaskedit é data
  If Not IsDate(pMsk.Text) Then
    EData = False
  Else
    If CInt(Mid(pMsk.ClipText, 3, 2)) > 12 Then
      EData = False
    End If
  End If
  If pTipo = TpObrigatorio And Not (EData) Then
    'Campo é obrigatório e não é data
    Valida_Data = False
  ElseIf pTipo = TpNaoObrigatorio Then
    'Campo não é obrigatório
    If Len(pMsk.ClipText) <> 0 And Not EData Then
      'Digitou algo que não é data
      Valida_Data = False
    Else
      Valida_Data = True
    End If
  Else
    Valida_Data = True
  End If
End Function

Public Function Valida_Moeda(pMsk As MaskEdBox, pTipo As TpObriga) As Boolean
  Dim EValor As Boolean
  EValor = True
  'Verifica se Mmaskedit é valor
  If Not IsNumeric(pMsk.Text) Then
    EValor = False
  End If
  If pTipo = TpObrigatorio And Not (EValor) Then
    'Campo é obrigatório e não é Valor
    Valida_Moeda = False
  ElseIf pTipo = TpNaoObrigatorio Then
    'Campo não é obrigatório
    If Len(pMsk.ClipText) <> 0 And Not EValor Then
      'Digitou algo que não é Valor
      Valida_Moeda = False
    Else
      Valida_Moeda = True
    End If
  Else
    Valida_Moeda = True
  End If
End Function

Public Sub PreencheCombo(cbo, _
                         ByVal sSql As String, _
                         Optional TipoTodos As Boolean = True, _
                         Optional TipoBranco As Boolean = False)
' As ComboBox
Dim ssAux As ADODB.Recordset
Dim objGeral As busSisContas.clsGeral
On Error GoTo trata
  Set objGeral = New busSisContas.clsGeral
  Set ssAux = objGeral.ExecutarSQL(sSql)
  cbo.Clear
  If TipoBranco Then _
    cbo.AddItem ""
  If TipoTodos Then _
    cbo.AddItem "<TODOS>"
  While Not ssAux.EOF
    cbo.AddItem ssAux(0) & ""
    ssAux.MoveNext
  Wend
  Set objGeral = Nothing
  Exit Sub
trata:
  AmpN
  Err.Raise Err.Description, vbExclamation, TITULOSISTEMA
End Sub

Public Sub LimparCampoMask(objMask As MaskEdBox)
  Dim strMask As String
  With objMask
    strMask = .Mask
    .Mask = ""
    .Text = ""
    .Mask = strMask
  End With
End Sub
Public Sub LimparCampoCombo(objCbo As ComboBox)
  objCbo.Clear
End Sub
Public Sub LimparCampoTexto(objText As TextBox)
  objText.Text = ""
End Sub

Function INCLUIR_VALOR_NO_MASK(pMask As MaskEdBox, pValor As Variant, pTipo As tpMaskValor) As String
  On Error GoTo trata
  Dim strMask As String
  'Limpa Maskedit
  With pMask
    strMask = .Mask
    .Mask = ""
    .Text = ""
    .Mask = strMask
  End With
  '
  If pTipo = tpMaskValor.TpMaskData Then
    'Campo é data
    If Len(strMask) = 10 Then
      'Formato DD/MM/YYYY
      If Not IsNull(pValor) Then pMask.Text = Format(pValor, "DD/MM/YYYY")
    Else
      'Formato DD/MM/YYYY hh:mm
      If Not IsNull(pValor) Then pMask.Text = Format(pValor, "DD/MM/YYYY hh:mm")
    End If
  ElseIf pTipo = tpMaskValor.TpMaskLongo Then
    'Campo é Longo
    If Not IsNull(pValor) Then pMask.Text = Format(pValor, "#,##0")
    
  ElseIf pTipo = tpMaskValor.TpMaskMoeda Then
    'Campo é moeda
    If Not IsNull(pValor) Then pMask.Text = Format(pValor, "#,##0.00")
  Else
    'Campo é outros
    If Not IsNull(pValor) Then pMask.Text = pValor
  End If
  
  Exit Function
trata:
  AmpN
  Err.Raise Err.Number, "[mdlUserFunction.INCLUIR_VALOR_NO_MASK]", Err.Description
  
End Function


