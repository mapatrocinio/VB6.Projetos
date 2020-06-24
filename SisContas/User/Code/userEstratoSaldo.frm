VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmUserEstratoSaldo 
   Caption         =   "Visualização de saldo"
   ClientHeight    =   6045
   ClientLeft      =   2595
   ClientTop       =   3120
   ClientWidth     =   11415
   Icon            =   "userEstratoSaldo.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6045
   ScaleWidth      =   11415
   Begin VB.PictureBox picBotoes 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   11415
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   5145
      Width           =   11415
      Begin VB.PictureBox picAlinDir 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   912
         Left            =   0
         ScaleHeight     =   915
         ScaleWidth      =   11295
         TabIndex        =   1
         Top             =   0
         Width           =   11295
         Begin VB.CommandButton cmdImprimir 
            Height          =   735
            Left            =   8760
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   1215
         End
         Begin VB.CommandButton cmdFechar 
            Height          =   735
            Left            =   9960
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            Width           =   1215
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Align           =   1  'Align Top
      Height          =   4935
      Left            =   0
      OleObjectBlob   =   "userEstratoSaldo.frx":000C
      TabIndex        =   2
      Top             =   0
      Width           =   11415
   End
End
Attribute VB_Name = "frmUserEstratoSaldo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Dim blnPrimeiraVez        As Boolean
Public strpDataInicial    As String
Public strpDataFinal      As String
Public strpMostrarApenasSaldo As String
Private Matriz()          As String

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdImprimir_Click()
  On Error GoTo TratErro
  AmpS
  '
  'Cabeçalho do report
  grdGeral.PrintInfo.PageHeader = gsNomeEmpresa & " - Saldo de Contas"
  grdGeral.PrintInfo.PageHeader = grdGeral.PrintInfo.PageHeader & vbCrLf & "Emissão: " & Format(Now, "DD/MM/YYYY hh:mm") & " - Data " & strpDataInicial
  grdGeral.PrintInfo.RepeatColumnHeaders = True
  '
  grdGeral.PrintInfo.SettingsMarginBottom = 400
  grdGeral.PrintInfo.SettingsMarginLeft = 1000
  grdGeral.PrintInfo.SettingsMarginRight = 1000
  grdGeral.PrintInfo.SettingsMarginTop = 600
  grdGeral.PrintInfo.PreviewMaximize = True
  grdGeral.PrintInfo.SettingsOrientation = 2
  grdGeral.PrintInfo.PrintPreview
  '
  AmpN
  Exit Sub
  
TratErro:
  AmpN
  MsgBox "O seguinte Erro Ocorreu: " & Err.Description, vbOKOnly, TITULOSISTEMA

End Sub
Public Sub MontaMatriz(strDtInicial As String, _
                       strDtFinal As String, _
                       strMostrarApenasSaldo As String)
  '
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim objRsMovDebAnt  As ADODB.Recordset
  Dim objRsMovCredAnt As ADODB.Recordset
  Dim objRsMovDebPer  As ADODB.Recordset
  Dim objRsMovCredPer As ADODB.Recordset
  
  Dim intI            As Long
  Dim intJ            As Long
  Dim intJMov         As Long
  Dim intLinhas       As Long
  Dim clsGer          As busSisContas.clsGeral
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
  Dim MatrizSemOrdem()      As String
  Dim MatrizMovimentacao()      As String
  Dim MatrizMovimentacaoOrd()   As String
  Dim lngLinhasMov              As Long
  '
  On Error GoTo trata
  AmpS
  '
  Set clsGer = New busSisContas.clsGeral
  '
  'Inicia Data Inicial
  datDataInicial = CDate(Right(strDtInicial, 4) & "/" & Mid(strDtInicial, 4, 2) & "/" & Mid(strDtInicial, 1, 2))
  '
  strSql = "Select * " & _
    "From CONTA " & _
    " WHERE PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo) & _
    " Order By CONTA.PKID;"
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
            " AND MOVIMENTACAO.PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo) & _
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
            " AND MOVIMENTACAO.PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo) & _
            " Order By CONTACREDITOID"
  '
  Set objRsMovCredPer = clsGer.ExecutarSQL(strSql)
  'Débito anterior
  strSql = "Select MOVIMENTACAO.* " & _
            "FROM MOVIMENTACAO " & _
            "Where " & _
            " MOVIMENTACAO.DATA < " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
            " AND not CONTADEBITOID is NULL " & _
            " AND MOVIMENTACAO.PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo) & _
            " Order By CONTADEBITOID"
  '
  Set objRsMovDebAnt = clsGer.ExecutarSQL(strSql)
  'Crédito anterior
  strSql = "Select MOVIMENTACAO.* " & _
            "FROM MOVIMENTACAO " & _
            "Where " & _
            " MOVIMENTACAO.DATA < " & Formata_Dados(strDtInicial, tpDados_DataHora, tpNulo_NaoAceita) & _
            " AND not CONTACREDITOID is NULL " & _
            " AND MOVIMENTACAO.PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo) & _
            " Order By CONTACREDITOID"
  '
  Set objRsMovCredAnt = clsGer.ExecutarSQL(strSql)
  '
  ReDim MatrizSemOrdem(0 To COLUNASMATRIZ - 1, 0 To 0)
  '
  If Not objRs.EOF Then   'se já houver algum item
    LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then   'se já houver algum item
    intLinhas = 0
    curVrSaldoTotalGeral = 0
    For intI = 0 To LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
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
        If strpMostrarApenasSaldo = "N" Then
          'Valor total saldo da movimentada
          curVrSaldoTotal = curVrSaldo
        End If
        curVrSaldoTotal = curVrSaldoTotal + curVrSaldoPeriodo
        'Inicializa vetor
        ReDim MatrizMovimentacao(0 To 6, 0 To 0)
        lngLinhasMov = 0
        'Pegar Movimentação no período Débito
        curMovDebPer = 0
        If Not objRsMovDebPer.EOF Then   'se já houver algum registro
          Do While objRsMovDebPer.Fields("CONTADEBITOID").Value & "" = objRs.Fields("PKID").Value & ""
            If Not IsDate(objRs.Fields("DTSALDO").Value) Then
              curMovDebPer = curMovDebPer + objRsMovDebPer.Fields("VALOR").Value
              'Popular matriz de movimentação
              lngLinhasMov = lngLinhasMov + 1
              ReDim Preserve MatrizMovimentacao(0 To 6, 0 To lngLinhasMov - 1)
              MatrizMovimentacao(0, lngLinhasMov - 1) = objRsMovDebPer.Fields("CONTADEBITOID").Value & ""
              MatrizMovimentacao(1, lngLinhasMov - 1) = objRsMovDebPer.Fields("DATA").Value
              MatrizMovimentacao(2, lngLinhasMov - 1) = objRsMovDebPer.Fields("DOCUMENTO").Value & ""
              MatrizMovimentacao(3, lngLinhasMov - 1) = objRsMovDebPer.Fields("DESCRICAO").Value & ""
              MatrizMovimentacao(4, lngLinhasMov - 1) = IIf(objRsMovDebPer.Fields("VALOR").Value = 0, "", objRsMovDebPer.Fields("VALOR").Value)
              MatrizMovimentacao(5, lngLinhasMov - 1) = "D"
              MatrizMovimentacao(6, lngLinhasMov - 1) = curVrSaldoTotal
              '------
            Else
              If IsDate(objRsMovDebPer.Fields("DATA").Value) Then
                If objRsMovDebPer.Fields("DATA").Value >= objRs.Fields("DTSALDO").Value Then
                  curMovDebPer = curMovDebPer + objRsMovDebPer.Fields("VALOR").Value
                  'Popular matriz de movimentação
                  lngLinhasMov = lngLinhasMov + 1
                  ReDim Preserve MatrizMovimentacao(0 To 6, 0 To lngLinhasMov - 1)
                  MatrizMovimentacao(0, lngLinhasMov - 1) = objRsMovDebPer.Fields("CONTADEBITOID").Value & ""
                  MatrizMovimentacao(1, lngLinhasMov - 1) = objRsMovDebPer.Fields("DATA").Value
                  MatrizMovimentacao(2, lngLinhasMov - 1) = objRsMovDebPer.Fields("DOCUMENTO").Value & ""
                  MatrizMovimentacao(3, lngLinhasMov - 1) = objRsMovDebPer.Fields("DESCRICAO").Value & ""
                  MatrizMovimentacao(4, lngLinhasMov - 1) = IIf(objRsMovDebPer.Fields("VALOR").Value = 0, "", objRsMovDebPer.Fields("VALOR").Value)
                  MatrizMovimentacao(5, lngLinhasMov - 1) = "D"
                  MatrizMovimentacao(6, lngLinhasMov - 1) = curVrSaldoTotal
                  '------
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
              lngLinhasMov = lngLinhasMov + 1
              ReDim Preserve MatrizMovimentacao(0 To 6, 0 To lngLinhasMov - 1)
              MatrizMovimentacao(0, lngLinhasMov - 1) = objRsMovCredPer.Fields("CONTACREDITOID").Value & ""
              MatrizMovimentacao(1, lngLinhasMov - 1) = objRsMovCredPer.Fields("DATA").Value
              MatrizMovimentacao(2, lngLinhasMov - 1) = objRsMovCredPer.Fields("DOCUMENTO").Value & ""
              MatrizMovimentacao(3, lngLinhasMov - 1) = objRsMovCredPer.Fields("DESCRICAO").Value & ""
              MatrizMovimentacao(4, lngLinhasMov - 1) = IIf(objRsMovCredPer.Fields("VALOR").Value = 0, "", objRsMovCredPer.Fields("VALOR").Value)
              MatrizMovimentacao(5, lngLinhasMov - 1) = "C"
              MatrizMovimentacao(6, lngLinhasMov - 1) = curVrSaldoTotal
              '------
            Else
              If IsDate(objRsMovCredPer.Fields("DATA").Value) Then
                If objRsMovCredPer.Fields("DATA").Value >= objRs.Fields("DTSALDO").Value Then
                  curMovCredPer = curMovCredPer + objRsMovCredPer.Fields("VALOR").Value
                  'Popular matriz de movimentação
                  lngLinhasMov = lngLinhasMov + 1
                  ReDim Preserve MatrizMovimentacao(0 To 6, 0 To lngLinhasMov - 1)
                  MatrizMovimentacao(0, lngLinhasMov - 1) = objRsMovCredPer.Fields("CONTACREDITOID").Value & ""
                  MatrizMovimentacao(1, lngLinhasMov - 1) = objRsMovCredPer.Fields("DATA").Value
                  MatrizMovimentacao(2, lngLinhasMov - 1) = objRsMovCredPer.Fields("DOCUMENTO").Value & ""
                  MatrizMovimentacao(3, lngLinhasMov - 1) = objRsMovCredPer.Fields("DESCRICAO").Value & ""
                  MatrizMovimentacao(4, lngLinhasMov - 1) = IIf(objRsMovCredPer.Fields("VALOR").Value = 0, "", objRsMovCredPer.Fields("VALOR").Value)
                  MatrizMovimentacao(5, lngLinhasMov - 1) = "C"
                  MatrizMovimentacao(6, lngLinhasMov - 1) = curVrSaldoTotal
                  '------
                End If
              End If
            End If
            objRsMovCredPer.MoveNext
            If objRsMovCredPer.EOF Then Exit Do
          Loop
        End If
        'Monta Matriz
        intLinhas = intLinhas + 1
        ReDim Preserve MatrizSemOrdem(0 To COLUNASMATRIZ - 1, 0 To intLinhas - 1)
        'Descrição
        MatrizSemOrdem(0, intLinhas - 1) = strDescricao
        '
        'If strMostrarApenasSaldo = "N" Then
          'Preencher movimentação
          If lngLinhasMov > 0 Then
          'PASSO 1 - Ordenar pelo código
          ReDim MatrizMovimentacaoOrd(0 To 6, 0 To lngLinhasMov - 1)
          Ordenar_Matriz MatrizMovimentacaoOrd, _
                         MatrizMovimentacao, _
                         CLng(lngLinhasMov), _
                         2, _
                         "Ascendente", _
                         DataTypeEnum.adDate
        
          'PASSO 2 - Carregar linhas na tabela
          For intJMov = 0 To lngLinhasMov - 1
            'Para cada linha
            'Monta Matriz
            'Valor total saldo da movimentada
            curVrSaldoTotal = curVrSaldoTotal + IIf(MatrizMovimentacao(5, intJMov) = "D", -(IIf(IsNumeric(MatrizMovimentacaoOrd(4, intJMov)), MatrizMovimentacaoOrd(4, intJMov), 0)), IIf(IsNumeric(MatrizMovimentacaoOrd(4, intJMov)), MatrizMovimentacaoOrd(4, intJMov), 0))
            '
          Next
        End If
        'End If
        'Saldo
        MatrizSemOrdem(1, intLinhas - 1) = IIf(curVrSaldoTotal = 0, "", Format(curVrSaldoTotal, "###,##0.00"))
        'Soma total
        curVrSaldoTotalGeral = curVrSaldoTotalGeral + curVrSaldoTotal
'''        If lngLinhasMov > 0 Then
'''          'PASSO 3 - Carregar Saldo Total
'''          'Monta Matriz
'''          intLinhas = intLinhas + 1
'''          ReDim Preserve Matriz(0 To COLUNASMATRIZ - 1, 0 To intLinhas - 1)
'''          'Descrição
'''          Matriz(0, intLinhas - 1) = "      SALDO GERAL DA MOVIMENTAÇÃO "
'''          'Documento
'''          Matriz(1, intLinhas - 1) = ""
'''          'Data
'''          Matriz(2, intLinhas - 1) = ""
'''          'Dédito
'''          Matriz(3, intLinhas - 1) = ""
'''          'Crédito
'''          Matriz(4, intLinhas - 1) = ""
'''          'Saldo
'''          Matriz(5, intLinhas - 1) = IIf(curVrSaldoTotal = 0, "", curVrSaldoTotal)
'''          '
'''        End If
        objRs.MoveNext
      End If
      
      
    Next  'próxima linha matriz
    
    'Monta Matriz
    intLinhas = intLinhas + 1
    ReDim Preserve MatrizSemOrdem(0 To COLUNASMATRIZ - 1, 0 To intLinhas - 1)
    'Descrição
    MatrizSemOrdem(0, intLinhas - 1) = "| TOTAL GERAL |"
    '
    'Saldo
    MatrizSemOrdem(1, intLinhas - 1) = IIf(curVrSaldoTotalGeral = 0, "", Format(curVrSaldoTotalGeral, "###,##0.00"))
    
    LINHASMATRIZ = intLinhas
    If intLinhas > 0 Then
      'PASSO 1 - Ordenar pela descrição
      ReDim Matriz(0 To COLUNASMATRIZ - 1, 0 To intLinhas - 1)
      Ordenar_Matriz Matriz, _
                     MatrizSemOrdem, _
                     CLng(intLinhas), _
                     1, _
                     "Ascendente", _
                     DataTypeEnum.adVarChar
    End If
  End If
  Set clsGer = Nothing
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Activate()
  If blnPrimeiraVez Then
    DoEvents
    '
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0
  
    MontaMatriz strpDataInicial, _
                strpDataFinal, _
                strpMostrarApenasSaldo
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
    blnPrimeiraVez = False
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  ReDim Matriz(0 To 0, 0 To 0)
  LINHASMATRIZ = 0
  COLUNASMATRIZ = 0
End Sub

Private Sub grdGeral_UnboundReadDataEx( _
     ByVal RowBuf As TrueDBGrid60.RowBuffer, _
    StartLocation As Variant, ByVal Offset As Long, _
    ApproximatePosition As Long)
  ' UnboundReadData is fired by an unbound grid whenever
  ' it requires data for display. This event will fire
  ' when the grid is first shown, when Refresh or ReBind
  ' is used, when the grid is scrolled, and after a
  ' record in the grid is modified and the user commits
  ' the change by moving off of the current row. The
  ' grid fetches data in "chunks", and the number of rows
  ' the grid is asking for is given by RowBuf.RowCount.
  ' RowBuf is the row buffer where you place the data
  ' the bookmarks for the rows that the grid is
  ' requesting to display. It will also hold the number
  ' of rows that were successfully supplied to the grid.
  ' StartLocation is a vrtBookmark which, together with
  ' Offset, specifies the row for the programmer to start
  ' transferring data. A StartLocation of Null indicates
  ' a request for data from BOF or EOF.
  ' Offset specifies the relative position (from
  ' StartLocation) of the row for the programmer to start
  ' transferring data. A positive number indicates a
  ' forward relative position while a negative number
  ' indicates a backward relative position. Regardless
  ' of whether the rows to be read are before or after
  ' StartLocation, rows are always fetched going forward
  ' (this is why there is no ReadPriorRows parameter to
  ' the procedure).
  ' If you page down on the grid, for instance, the new
  ' top row of the grid will have an index greater than
  ' the StartLocation (Offset > 0). If you page up on
  ' the grid, the new index is less than that of
  ' StartLocation, so Offset < 0. If StartLocation is
  ' a vrtBookmark to row N, the grid always asks for row
  ' data in the following order:
  '   (N + Offset), (N + Offset + 1), (N + Offset + 2)...
  ' ApproximatePosition is a value you can set to indicate
  ' the ordinal position of (StartLocation + Offset).
  ' Setting this variable will enhance the ability of the
  ' grid to display its vertical scroll bar accurately.
  ' If the exact ordinal position of the new location is
  ' not known, you can set it to a reasonable,
  ' approximate value, or just ignore this parameter.
    
  On Error GoTo trata
  '
  Dim intColIndex      As Integer
  Dim intJ             As Integer
  Dim intRowsFetched   As Integer
  Dim intI             As Long
  Dim lngNewPosition   As Long
  Dim vrtBookmark      As Variant
  '
  intRowsFetched = 0
  For intI = 0 To RowBuf.RowCount - 1
    ' Get the vrtBookmark of the next available row
    vrtBookmark = GetRelativeBookmarkGeral(StartLocation, _
               Offset + intI, LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, COLUNASMATRIZ, LINHASMATRIZ, Matriz)
    Next intJ
  
    ' Set the vrtBookmark for the row
    RowBuf.Bookmark(intI) = vrtBookmark
  
    ' Increment the count of fetched rows
    intRowsFetched = intRowsFetched + 1
  Next intI
  
' Tell the grid how many rows were fetched
  RowBuf.RowCount = intRowsFetched
  
' Set the approximate scroll bar position. Only
' nonnegative values of IndexFromBookmark() are valid.
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserPedidoLis.grdGeral_UnboundReadDataEx]"
End Sub


Private Sub Form_Load()
On Error GoTo trata
  AmpS
  Me.Height = 6450
  Me.Width = 11535
  blnPrimeiraVez = True
  
  CenterForm Me
  'PreencheCombo cboEstInter, "SELECT DESCRICAO FROM  GRUPOESTOQUE ORDER BY DESCRICAO"
  
  Me.Caption = Me.Caption & " - " & strpDataInicial
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar, , , , , , cmdImprimir
  '
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
