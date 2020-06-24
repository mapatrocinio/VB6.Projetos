VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Form1"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrTempo 
      Interval        =   60000
      Left            =   240
      Top             =   510
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub AtualizaTelefonema()
  On Error GoTo trata
  Dim strSql                        As String
  Dim objRs                         As ADODB.Recordset
  Dim objRsInt                      As ADODB.Recordset
  Dim datDataHoraEntrada            As Date
  Dim datDataHoraFecha              As Date
  Dim datDataHoraCalcTelefonema     As Date
  Dim strRamalLocacao               As String
  Dim lngQtdTelefonemasInseridos    As Long
  Dim intArq                        As Integer
  Dim lngLine                       As Long
  Dim strLine                       As String
  'Campos do Arquivo
  Dim strRamal                      As String
  Dim strNumero                     As String
  Dim strHora                       As String
  Dim strDuracao                    As String
  Dim strData                       As String
  Dim strValor                      As String
  Dim strRegiao                     As String
  Dim objGeral                      As busSisMotel.clsGeral
  '
  Set objGeral = New busSisMotel.clsGeral
  'Fim dos Campos
  'Inicializa Variáveis de data
  datDataHoraEntrada = Now
  datDataHoraFecha = Now
  
  
  strSql = "Select LOCACAO.*, APARTAMENTO.RAMAL From LOCACAO INNER JOIN APARTAMENTO ON APARTAMENTO.PKID = LOCACAO.APARTAMENTOID WHERE LOCACAO.OCUPADO = " & Formata_Dados(True, tpDados_Boolean)
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  Do While Not objRs.EOF
    datDataHoraEntrada = objRs.Fields("DATAHORAENTRA").Value
    strRamalLocacao = objRs.Fields("RAMAL").Value & ""
    If objRs.Fields("SAIU").Value Then 'Já possui saída, Pegar dataHora do fechamento
      datDataHoraFecha = objRs.Fields("DTHORAFECHA").Value
    Else
      If IsNull(objRs.Fields("dtHoraCalcFecha").Value) Then
        datDataHoraFecha = Now
      Else
        datDataHoraFecha = objRs.Fields("dtHoraCalcFecha").Value
      End If
    End If
    '
    'De Posse das Horas, Pega Quantidade de ligações
    'já cadastradas para esta Unidade, Feitas pela mesa Telefonica
    strSql = "Select COUNT(*) FROM TELEFONEMA WHERE LOCACAOID = " & objRs.Fields("PKID").Value & " AND MESATEL = " & Formata_Dados(True, tpDados_Boolean)
    Set objRsInt = objGeral.ExecutarSQL(strSql)
    If objRsInt.EOF Then
      lngQtdTelefonemasInseridos = 0
    Else
      If Not IsNumeric(objRsInt(0)) Then
        lngQtdTelefonemasInseridos = 0
      Else
        lngQtdTelefonemasInseridos = objRsInt(0)
      End If
    End If
    objRsInt.Close
    Set objRsInt = Nothing
    'Abre Arquivo Texto
    intArq = FreeFile
    Open gsCaminhoMesaTel For Input As intArq
    '
    lngLine = 1
    Do While Not EOF(intArq)
      Line Input #intArq, strLine
      'Pegar Campos do Arquivo
      If True = True Then
        'novas definições para o capri
        strRamal = Trim(Mid(strLine, 1, 6))
        strNumero = Trim(Mid(strLine, 36, 18))
        strHora = Trim(Mid(strLine, 26, 5))
        strDuracao = Trim(Mid(strLine, 100, 5))
        strData = Trim(Mid(strLine, 13, 8))
        strData = Mid(strData, 1, 6) & "20" & Right(strData, 2)
        strValor = Trim(Mid(strLine, 105, 11))
        strRegiao = Trim(Mid(strLine, 54, 32))
        
      Else
        'antigo ainda não implementado, mas não descartado
        strRamal = Trim(Mid(strLine, 18, 5))
        strNumero = Trim(Mid(strLine, 24, 20))
        strHora = Trim(Mid(strLine, 45, 8))
        strDuracao = Trim(Mid(strLine, 54, 8))
        strData = Trim(Mid(strLine, 63, 10))
        strValor = Trim(Mid(strLine, 77, 8))
        strRegiao = Trim(Mid(strLine, 86, 25))
      End If
      '
      If IsDate(strData) Then
        'Não é Header, Valida Linha
        'PASSO 1 - VERIFICA SE É UNIDADE
        If strRamalLocacao = strRamal Then
          'Verifica se está no Período
          'Calcula a data/Hora do telefonema
          datDataHoraCalcTelefonema = CDate(Mid(strData, 7, 4) & "/" & Mid(strData, 4, 2) & "/" & Mid(strData, 1, 2) & " " & strHora)
          'Verifica se está no Intervalo
          If (datDataHoraCalcTelefonema >= datDataHoraEntrada) And (datDataHoraCalcTelefonema <= datDataHoraFecha) Then
            'está no Intervalo
            'Somente Insere se a linha corrente for maior que a quantidade
            'já inserida na tabela de telefonemas
            If lngLine > lngQtdTelefonemasInseridos Then
              'INSERE REGISTROS NA TABELA
              strSql = "INSERT INTO TELEFONEMA (QTDMINUTOS, VALOR, LOCACAOID, DATAHORA, DURACAO, REGIAO, NUMERO, MESATEL) VALUES (0, " & strValor & ", " & objRs.Fields("PKID").Value & ", " & Formata_Dados(strData & " " & Left(strHora, 5), tpDados_DataHora, tpNulo_NaoAceita) & ", " & _
                Formata_Dados(strDuracao, tpDados_Texto, tpNulo_NaoAceita, 8) & ", " & Formata_Dados(strRegiao, tpDados_Texto, tpNulo_NaoAceita, 25) & ", " & Formata_Dados(strNumero, tpDados_Texto, tpNulo_NaoAceita, 20) & ", True)"
              objGeral.ExecutarSQLAtualizacao strSql
            End If
            lngLine = lngLine + 1
          End If
          
        End If
      End If
    Loop
    Close intArq
    objRs.MoveNext
  Loop
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  Exit Sub
trata:
  MsgBox "o seguinte erro ocorreu : " & Err.Number & " - " & Err.Description
End Sub



Private Sub Form_Load()
  AtualizaTelefonema
End Sub

Private Sub tmrTempo_Timer()
  AtualizaTelefonema
End Sub
