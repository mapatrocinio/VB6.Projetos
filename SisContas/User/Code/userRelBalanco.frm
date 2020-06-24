VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmUserRelBalanco 
   Caption         =   "Balanço"
   ClientHeight    =   3840
   ClientLeft      =   450
   ClientTop       =   1920
   ClientWidth     =   6105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3840
   ScaleWidth      =   6105
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   9
      Top             =   120
      Width           =   5775
      Begin VB.OptionButton optSai2 
         Caption         =   "Impressora"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optSai1 
         Caption         =   "Vídeo"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   6105
      TabIndex        =   8
      Top             =   3075
      Width           =   6105
      Begin VB.CommandButton cmdRelatorio 
         Default         =   -1  'True
         Height          =   735
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Height          =   735
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   5775
      Begin VB.CheckBox chkSaldoAntCart 
         Caption         =   "Pegar saldo anteiror co cartão proveniente dos lançamentos"
         Height          =   435
         Left            =   1560
         TabIndex        =   11
         Top             =   990
         Value           =   1  'Checked
         Width           =   2805
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   255
         Index           =   1
         Left            =   1560
         TabIndex        =   2
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskVrSaldoAnterior 
         Height          =   255
         Left            =   1560
         TabIndex        =   3
         Top             =   720
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Saldo Anterior"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   720
         Width           =   1095
      End
      Begin VB.Label lblCliente 
         Caption         =   "Data Fianl : "
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Tag             =   "lblIdCliente"
         Top             =   360
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3360
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmUserRelBalanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdRelatorio_Click()
  Dim strDataIni      As String
  Dim strDataIniAntes As String
  Dim strDataFimAntes As String
  Dim lngQtdDiasNoMes As Long
  Dim objGeral        As busSisContas.clsGeral
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim curVrChqDevol   As Currency
  Dim curVrCarAnt     As Currency
  Dim curVrChqRecAnt  As Currency
  Dim curVrPenRecAnt  As Currency
  Dim curVrOutRen     As Currency
  Dim curVrPgto       As Currency
  Dim curVrTxCart     As Currency
  Dim strMesAtual     As String
  Dim strMesAnterior  As String
  '
  On Error GoTo TratErro
  AmpS
  
  If Not Valida_Data(mskData(1), TpObrigatorio) Then
    AmpN
    TratarErroPrevisto "Data Final Inválida", "[cmdRelatorio_Click]"
    SetarFoco mskData(1)
    Pintar_Controle mskData(1), tpCorContr_Erro
    Exit Sub
  ElseIf Not Valida_Moeda(mskVrSaldoAnterior, TpNaoObrigatorio) Then
    AmpN
    TratarErroPrevisto "Valor do Saldo Anterior inválido", "[cmdRelatorio_Click]"
    SetarFoco mskVrSaldoAnterior
    Pintar_Controle mskVrSaldoAnterior, tpCorContr_Erro
    Exit Sub
  End If
  '
  If optSai1.Value Then
    Report1.Destination = 0 'Video
  ElseIf optSai2.Value Then
    Report1.Destination = 1   'Impressora
  End If
  Report1.CopiesToPrinter = 1
  Report1.WindowState = crptMaximized
  'Captura data inicial e quantidade de dias no mes
  strDataIni = "01" & Right(mskData(1).Text, 8)
  lngQtdDiasNoMes = DateDiff("d", Right(strDataIni, 4) & "/" & Mid(strDataIni, 4, 2) & "/" & Left(strDataIni, 2), Right(mskData(1).Text, 4) & "/" & Mid(mskData(1).Text, 4, 2) & "/" & Left(mskData(1).Text, 2)) + 1
  Set objGeral = New busSisContas.clsGeral
  'Captura total de cheques devolvidos
  strSql = "SELECT Sum(CHEQUE.VALOR) As VALORCHEQUE " & _
    "FROM CHEQUE " & _
    "WHERE DTDEVOLUCAO >= " & Formata_Dados(strDataIni & " 00:00", tpDados_DataHora, tpNulo_Aceita) & _
    " AND DTDEVOLUCAO <= " & Formata_Dados(mskData(1).Text & " 23:59", tpDados_DataHora, tpNulo_Aceita) & _
    " AND CHEQUE.STATUS = 'D'"
  Set objRs = objGeral.ExecutarSQL(strSql)
  If objRs.EOF Then
    curVrChqDevol = 0
  ElseIf Not IsNumeric(objRs.Fields("VALORCHEQUE").Value) Then
    curVrChqDevol = 0
  Else
    curVrChqDevol = objRs.Fields("VALORCHEQUE").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Pegar faixa de data do mes anterior
  strDataIniAntes = "01/" & Retorna_mes_ano_anterior(Mid(mskData(1).Text, 4, 2), Right(mskData(1).Text, 4))
  strDataFimAntes = Retorna_ultimo_dia_do_mes(Mid(strDataIniAntes, 4, 2), Right(strDataIniAntes, 4)) & Right(strDataIniAntes, 8)
  'Captura total de cartões do mês anterior
  If chkSaldoAntCart.Value = 0 Then
    'Pegar salda anterior do cartão do SisContas
    strSql = "SELECT Sum(vw_cons_t_cred.VrCalcDescCartTxAdm) as VrCalcDescCartTxAdm, Sum(vw_cons_t_cred.PGTOCARTAO) As PGTOCARTAO " & _
      "FROM TURNO LEFT JOIN vw_cons_t_cred ON TURNO.PKID = vw_cons_t_cred.PKID " & _
      "WHERE TURNO.DATA >= " & Formata_Dados(strDataIniAntes & " 00:00", tpDados_DataHora, tpNulo_Aceita) & _
      " AND TURNO.DATA <= " & Formata_Dados(strDataFimAntes & " 23:59", tpDados_DataHora, tpNulo_Aceita)
  Else
    'Pegar salda anterior lançado no siscontas
    strSql = "SELECT Sum(DESPESA.VR_PAGO) As PGTOCARTAO, 0 AS VrCalcDescCartTxAdm " & _
      "FROM (GRUPODESPESA INNER JOIN SUBGRUPODESPESA ON GRUPODESPESA.PKID = SUBGRUPODESPESA.GRUPODESPESAID) " & _
      " INNER JOIN DESPESA ON SUBGRUPODESPESA.PKID = DESPESA.SUBGRUPODESPESAID " & _
      "WHERE DESPESA.DT_PAGAMENTO >= " & Formata_Dados(strDataIni & " 00:00", tpDados_DataHora, tpNulo_Aceita) & _
      " AND DESPESA.DT_PAGAMENTO <= " & Formata_Dados(mskData(1).Text & " 23:59", tpDados_DataHora, tpNulo_Aceita) & _
      " AND GRUPODESPESA.TIPO =  " & Formata_Dados("C", tpDados_Texto) & _
      " AND GRUPODESPESA.DESCRICAO = " & Formata_Dados("CARTÃO DE CRÉDITO", tpDados_Texto, tpNulo_Aceita)
  End If
  Set objRs = objGeral.ExecutarSQL(strSql)
  If objRs.EOF Then
    curVrCarAnt = 0
  ElseIf Not IsNumeric(objRs.Fields("PGTOCARTAO").Value) Then
    curVrCarAnt = 0
  Else
    curVrCarAnt = objRs.Fields("PGTOCARTAO").Value
  End If
  '---------
  If objRs.EOF Then
    curVrTxCart = 0
  ElseIf Not IsNumeric(objRs.Fields("VrCalcDescCartTxAdm").Value) Then
    curVrTxCart = 0
  Else
    curVrTxCart = objRs.Fields("VrCalcDescCartTxAdm").Value
  End If
    
  objRs.Close
  Set objRs = Nothing
  'Captura chqs resgatados do mês atual
  strSql = "SELECT Sum(CHEQUE.VALOR) As VALOR " & _
    "FROM CHEQUE " & _
    "WHERE CHEQUE.DTRECUPERACAO >= " & Formata_Dados(strDataIni & " 00:00", tpDados_DataHora, tpNulo_Aceita) & _
    " AND CHEQUE.DTRECUPERACAO <= " & Formata_Dados(mskData(1).Text & " 23:59", tpDados_DataHora, tpNulo_Aceita)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If objRs.EOF Then
    curVrChqRecAnt = 0
  ElseIf Not IsNumeric(objRs.Fields("VALOR").Value) Then
    curVrChqRecAnt = 0
  Else
    curVrChqRecAnt = objRs.Fields("VALOR").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Captura penhores resgatados do mês atual
  strSql = "SELECT Sum(PENHOR.VALOR) As VALOR " & _
    "FROM PENHOR " & _
    "WHERE PENHOR.DTRESGATE >= " & Formata_Dados(strDataIni & " 00:00", tpDados_DataHora, tpNulo_Aceita) & _
    " AND PENHOR.DTRESGATE <= " & Formata_Dados(mskData(1).Text & " 23:59", tpDados_DataHora, tpNulo_Aceita)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If objRs.EOF Then
    curVrPenRecAnt = 0
  ElseIf Not IsNumeric(objRs.Fields("VALOR").Value) Then
    curVrPenRecAnt = 0
  Else
    curVrPenRecAnt = objRs.Fields("VALOR").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Captura o valor de Pagamentos do mês atual
  strSql = "SELECT Sum(DESPESA.VR_PAGO) As VALOR " & _
    "FROM (GRUPODESPESA INNER JOIN SUBGRUPODESPESA ON GRUPODESPESA.PKID = SUBGRUPODESPESA.GRUPODESPESAID) " & _
    " INNER JOIN DESPESA ON SUBGRUPODESPESA.PKID = DESPESA.SUBGRUPODESPESAID " & _
    " INNER JOIN TURNO ON TURNO.PKID = DESPESA.TURNOID " & _
    "WHERE ((DESPESA.DT_PAGAMENTO >= " & Formata_Dados(strDataIni & " 00:00", tpDados_DataHora, tpNulo_Aceita) & _
    " AND DESPESA.DT_PAGAMENTO <= " & Formata_Dados(mskData(1).Text & " 23:59", tpDados_DataHora, tpNulo_Aceita) & _
    " AND DESPESA.TIPO = 'A') OR " & _
    "(TURNO.DATA >= " & Formata_Dados(strDataIni & " 00:00", tpDados_DataHora, tpNulo_Aceita) & _
    " AND TURNO.DATA <= " & Formata_Dados(mskData(1).Text & " 23:59", tpDados_DataHora, tpNulo_Aceita) & _
    " AND DESPESA.TIPO = 'T')) " & _
    " AND GRUPODESPESA.TIPO = " & Formata_Dados("D", tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If objRs.EOF Then
    curVrPgto = 0
  ElseIf Not IsNumeric(objRs.Fields("VALOR").Value) Then
    curVrPgto = 0
  Else
    curVrPgto = objRs.Fields("VALOR").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Captura o valor de outras rendas do mês atual
  strSql = "SELECT Sum(DESPESA.VR_PAGO) As VALOR " & _
    "FROM (GRUPODESPESA INNER JOIN SUBGRUPODESPESA ON GRUPODESPESA.PKID = SUBGRUPODESPESA.GRUPODESPESAID) " & _
    " INNER JOIN DESPESA ON SUBGRUPODESPESA.PKID = DESPESA.SUBGRUPODESPESAID " & _
    "WHERE DESPESA.DT_PAGAMENTO >= " & Formata_Dados(strDataIni & " 00:00", tpDados_DataHora, tpNulo_Aceita) & _
    " AND DESPESA.DT_PAGAMENTO <= " & Formata_Dados(mskData(1).Text & " 23:59", tpDados_DataHora, tpNulo_Aceita) & _
    " AND GRUPODESPESA.TIPO = " & Formata_Dados("C", tpDados_Texto) & _
    IIf(chkSaldoAntCart.Value = 0, "", " AND GRUPODESPESA.DESCRICAO <> " & Formata_Dados("CARTÃO DE CRÉDITO", tpDados_Texto, tpNulo_Aceita))
  Set objRs = objGeral.ExecutarSQL(strSql)
  If objRs.EOF Then
    curVrOutRen = 0
  ElseIf Not IsNumeric(objRs.Fields("VALOR").Value) Then
    curVrOutRen = 0
  Else
    curVrOutRen = objRs.Fields("VALOR").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Setar descrição dos meses na fórmula
  strMesAtual = Retorna_descr_mes(Mid(strDataIni, 4, 2))
  strMesAnterior = Retorna_descr_mes(Mid(strDataIniAntes, 4, 2))
  'Setar Formulas
  Report1.Formulas(0) = "DataIni = Date(" & Right(strDataIni, 4) & ", " & Mid(strDataIni, 4, 2) & ", " & Left(strDataIni, 2) & ")"
  Report1.Formulas(1) = "DataFim = Date(" & Right(mskData(1).Text, 4) & ", " & Mid(mskData(1).Text, 4, 2) & ", " & Left(mskData(1).Text, 2) & ")"
  Report1.Formulas(2) = "TotalDiasMesCorrente = " & lngQtdDiasNoMes
  'Report1.Formulas(3) = "TotalChqDevol = ccur(" & Formata_Dados(curVrChqDevol, tpDados_Moeda, tpNulo_Aceita) & ")"
  Report1.Formulas(3) = "TotalChqDevol = " & Formata_Dados(curVrChqDevol, tpDados_Moeda, tpNulo_Aceita) & ""
  Report1.Formulas(4) = "SACart = " & Formata_Dados(curVrCarAnt, tpDados_Moeda, tpNulo_Aceita)
  Report1.Formulas(5) = "SACheqResg = " & Formata_Dados(curVrChqRecAnt, tpDados_Moeda, tpNulo_Aceita)
  Report1.Formulas(6) = "SAPenResg = " & Formata_Dados(curVrPenRecAnt, tpDados_Moeda, tpNulo_Aceita)
  Report1.Formulas(7) = "SAOutrasRend = " & Formata_Dados(curVrOutRen, tpDados_Moeda, tpNulo_Aceita)
  Report1.Formulas(8) = "MesAtual = '" & strMesAtual & "'"
  Report1.Formulas(9) = "MesAnterior = '" & strMesAnterior & "'"
  'Report1.Formulas(10) = "TotalPagamentos = ccur(" & Formata_Dados(curVrPgto, tpDados_Moeda, tpNulo_Aceita) & ")"
  Report1.Formulas(10) = "TotalPagamentos = " & Formata_Dados(curVrPgto, tpDados_Moeda, tpNulo_Aceita) & ""
  'Report1.Formulas(11) = "TxCartao = ccur(" & Formata_Dados(curVrTxCart, tpDados_Moeda, tpNulo_Aceita) & ")"
  Report1.Formulas(11) = "TxCartao = " & Formata_Dados(curVrTxCart, tpDados_Moeda, tpNulo_Aceita) & ""
  Report1.Formulas(12) = "SASaldoAnterior = " & Formata_Dados(IIf(Not IsNumeric(mskVrSaldoAnterior.Text), 0, mskVrSaldoAnterior.Text), tpDados_Moeda, tpNulo_Aceita)
  
  Report1.Action = 1
  '
  AmpN
  Set objGeral = Nothing
  Exit Sub
  
TratErro:
  AmpN
  TratarErro Err.Number, Err.Description, "[frmUserRelBalanco.cmdRelatorio_Click]"
End Sub

Private Sub Form_Activate()
  mskData(1).SetFocus
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  AmpS
  CenterForm Me
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar, , , , , , cmdRelatorio
  '
  Report1.Connect = ConnectRpt
  '
  Report1.ReportFileName = gsReportPath & "Balanco.rpt"
  mskData(1).Text = Format(Now, "DD/MM/YYYY")
  '
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, "[frmUserRelBalanco.Form_Load]"
End Sub

Private Sub mskData_GotFocus(Index As Integer)
  Selecionar_Conteudo mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub

Private Sub mskVrSaldoAnterior_GotFocus()
  Selecionar_Conteudo mskVrSaldoAnterior
End Sub

Private Sub mskVrSaldoAnterior_LostFocus()
  Pintar_Controle mskVrSaldoAnterior, tpCorContr_Normal
End Sub

