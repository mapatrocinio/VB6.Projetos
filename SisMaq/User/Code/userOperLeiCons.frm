VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserOperLeiCons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Operacional Leiturista"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9540
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   9540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraUnidade 
      Caption         =   "Operacional Leiturista"
      Height          =   6015
      Left            =   60
      TabIndex        =   8
      Top             =   360
      Width           =   9435
      Begin TrueDBGrid60.TDBGrid grdLeiturista 
         Height          =   5595
         Left            =   90
         OleObjectBlob   =   "userOperLeiCons.frx":0000
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   300
         Width           =   9225
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Selecione a opção"
      Height          =   1125
      Left            =   60
      TabIndex        =   7
      Top             =   6420
      Width           =   9435
      Begin VB.CommandButton cmdSairSelecao 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   855
         Left            =   8430
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   180
         Width           =   900
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&A - Leitura                  "
         BeginProperty Font 
            Name            =   "Small Fonts"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   0
         Left            =   60
         TabIndex        =   3
         ToolTipText     =   "Boleto"
         Top             =   240
         Width           =   1305
      End
      Begin ComctlLib.StatusBar StatusBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1740
         Width           =   7500
         _ExtentX        =   13229
         _ExtentY        =   450
         SimpleText      =   ""
         _Version        =   327682
         BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
            NumPanels       =   5
            BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   6
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1940
               MinWidth        =   1940
               TextSave        =   "24/2/2011"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   5
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "23:48"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   1
               Alignment       =   1
               Bevel           =   2
               Enabled         =   0   'False
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "CAPS"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   2
               Alignment       =   1
               Bevel           =   2
               Enabled         =   0   'False
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "NUM"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   3
               Alignment       =   1
               AutoSize        =   2
               Bevel           =   2
               Enabled         =   0   'False
               Object.Width           =   1244
               MinWidth        =   1235
               TextSave        =   "INS"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.TextBox txtTurno 
      BackColor       =   &H00E0E0E0&
      Height          =   285
      Left            =   3570
      Locked          =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "txtTurno"
      Top             =   30
      Width           =   4815
   End
   Begin MSMask.MaskEdBox mskDataPrinc 
      Height          =   255
      Left            =   900
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   60
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
      MaxLength       =   10
      Format          =   "DD/MMM/YYYY"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label16 
      Caption         =   "Data"
      Height          =   255
      Left            =   60
      TabIndex        =   6
      Top             =   60
      Width           =   765
   End
   Begin VB.Label Label21 
      Caption         =   "Turno Corrente"
      Height          =   255
      Left            =   2190
      TabIndex        =   5
      Top             =   60
      Width           =   1245
   End
End
Attribute VB_Name = "frmUserOperLeiCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public nGrupo                         As Integer
Public blnRetorno                     As Boolean
Public blnFechar                      As Boolean
Private lngLEITURAID                  As Long
'''
'''Public objUserGRInc             As SisMaq.frmUserGRInc
'''Public objUserContaCorrente     As SisMaq.frmUserContaCorrente
'''
Public blnPrimeiraVez                 As Boolean 'Propósito: Preencher lista no combo

'Entrada Arrecadador
Private LEIT_COLUNASMATRIZ            As Long
Private LEIT_LINHASMATRIZ             As Long
Private LEIT_Matriz()                 As String


Public Sub Clique_botao(intIndice As Integer)
  If cmdSelecao(intIndice).Enabled = True And cmdSelecao(intIndice).Visible = True Then
    cmdSelecao_Click intIndice
  End If
End Sub



Private Sub Form_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  'Tratamento de tecla para verificação de chamada de Outras telas
  'verifica se tem permissão
  'Tudo ok, faz chamada
  If KeyAscii = 13 Then
    SendKeys "{tab}"
    Exit Sub
  End If
'''  Select Case KeyAscii
'''  Case 1
'''    'TURNO - ABERTURA/REIMPRESSÃO
'''    frmUserTurnoInc.Show vbModal 'Turno
'''    Form_Load
'''  Case 2
'''    'TURNO - FECHAMENTO
'''    FechamentoTurno
'''    Form_Load
'''  Case 3
'''    'DETALHAR ENTRADA
'''    frmUserEntradaLis.Show vbModal
'''    Form_Load
'''  Case 4
'''    'DETALHAR RETIRADA
'''    frmUserRetiradaLis.Show vbModal
'''    Form_Load
'''  Case 5
'''    'DETALHAR ENTRADA ARRECENTE
'''    frmUserEntradaAtendLis.Show vbModal
'''    Form_Load
'''  Case 6
'''    'DETALHAR BOLETO ARRECENTE
'''    frmUserBoletoAtendLis.Show vbModal
'''    Form_Load
'''  Case 4
'''    'ATUALIZAR
'''    Form_Load
'''  Case 5
'''    'CONSULTAR PRONTUÁRIO
'''    frmUserProntuarioGRCons.Show vbModal
'''    Form_Load
'''  Case 6
'''    'ZERAR SENHA
'''    frmUserZerarSenhaLis.Show vbModal
'''    Form_Load
'''  Case 7
'''    'CONSULTAR PROCEDIMENTO
'''    frmUserProcedimentoCons.indOrigem = 1
'''    frmUserProcedimentoCons.lngPRESTADORID = 0
'''    frmUserProcedimentoCons.Show vbModal
'''    Form_Load
'''  Case 8
'''    'CONSULTAR GR
'''    frmUserGRFinancCons.Show vbModal
'''    Form_Load
'''  End Select
  '
  'Trata_Matrizes_Totais
  'SetarFoco txtUsuario
  '
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             "[frmUserOperLeiCons.Form_KeyPress]"
End Sub

'''Private Sub cmdInfFinanc_Click()
'''  On Error GoTo trata
'''  'Chamar o form de Consulta/Visualização das Informações Financeiras.
'''  frmUserInfFinancLis.Show vbModal
'''  SetarFoco grdLeiturista
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, _
'''             Err.Description, _
'''             Err.Source
'''  AmpN
'''End Sub
'''
Private Sub cmdSairSelecao_Click()
  On Error GoTo trata
  blnFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Public Sub cmdSelecao_Click(Index As Integer)
  On Error GoTo trata
  nGrupo = Index
  'strNumeroAptoPrinc = optUnidade
  'If Not ValiCamposPrinc Then Exit Sub
  VerificaQuemChamou
  'Atualiza Valores
  '
  SetarFoco cmdSelecao(0)
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[cmdSelecao_Click]"
  frmMDI.tmrUnidade.Enabled = True
End Sub


Public Sub VerificaQuemChamou()
  Dim objUserLeituraInc As SisMaq.frmUserLeituraInc
  Dim objUserLeituraFechaInc As SisMaq.frmUserLeituraFechaInc
  Dim strMsg As String
  On Error GoTo trata
  '
  Select Case nGrupo

  Case 0
    'Leitura
'''    If RetornaCodTurnoCorrente = 0 Then
'''      MsgBox "Não há turno aberto. Por favor abra o turno antes de iniciar as atividades.", vbExclamation, TITULOSISTEMA
'''      SetarFoco cmdSelecao(0)
'''      Exit Sub
'''    End If
    '
    Set objUserLeituraInc = New SisMaq.frmUserLeituraInc
    objUserLeituraInc.Status = tpStatus_Incluir
    objUserLeituraInc.lngLEITURAID = 0
    objUserLeituraInc.Show vbModal
    Set objUserLeituraInc = Nothing
    
    Form_Load
    'Montar RecordSet
    LEIT_COLUNASMATRIZ = grdLeiturista.Columns.Count
    LEIT_LINHASMATRIZ = 0
    MontaLEIT_Matriz
    grdLeiturista.Bookmark = Null
    grdLeiturista.ReBind
    grdLeiturista.ApproxCount = LEIT_LINHASMATRIZ
    
  End Select
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  End
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim strSql            As String
  Dim datDataTurno      As Date
  Dim datDataIniAtual   As Date
  Dim datDataFimAtual   As Date
  Dim objGeral          As busSisMaq.clsGeral
  Dim objRs             As ADODB.Recordset
  '
'''  If RetornaCodTurnoCorrente(datDataTurno) = 0 Then
'''    TratarErroPrevisto "Não há turnos em aberto, favor informar ao Gerente para abrir o turno.", "Form_Load"
'''    End
'''  Else
    'OK Para turno
'''    datDataIniAtual = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now) & " 00:00:00")
'''    datDataFimAtual = DateAdd("d", 1, datDataIniAtual)
'''    If datDataTurno < datDataIniAtual Or datDataTurno >= datDataFimAtual Then
'''      TratarErroPrevisto "ATENÇÃO" & vbCrLf & vbCrLf & "A data do turno atual aberto não corresponde a data de hoje:" & vbCrLf & vbCrLf & "Data do turno --> " & Format(datDataTurno, "DD/MM/YYYY") & vbCrLf & "Data Atual --> " & Format(datDataIniAtual, "DD/MM/YYYY") & vbCrLf & vbCrLf & "Por favor, feche o turno e abra-o novamente.", "Form_Load"
'''    End If
'  End If

  blnFechar = False
  blnRetorno = False
  blnPrimeiraVez = True
  AmpS
'''  If Me.ActiveControl Is Nothing Then
'''    Me.Top = 580
'''    Me.Left = 1
'''    Me.WindowState = 2 'Maximizado
'''  End If
  Me.Height = 8145
  Me.Width = 9630
  CenterForm Me
  '
  LerFigurasAvulsas cmdSairSelecao, "Sair.ico", "SairDown.ico", "Sair"
  '
  txtTurno.Text = RetornaDescTurnoCorrente
  mskDataPrinc.Text = Format(Date, "DD/MM/YYYY")

  'Ok
  Set objGeral = New busSisMaq.clsGeral
  strSql = "Select LEITURA.PKID "
  strSql = strSql & " FROM LEITURA "
  strSql = strSql & " WHERE LEITURA.DATA =  " & Formata_Dados(mskDataPrinc.Text, tpDados_DataHora)
  strSql = strSql & " AND LEITURA.LEITURISTAID =  " & Formata_Dados(giFuncionarioId, tpDados_Longo)

  Set objRs = objGeral.ExecutarSQL(strSql)
  'Verifica se o boleto existe para o usuário
  If objRs.EOF Then
    lngLEITURAID = 0
  Else
    lngLEITURAID = objRs.Fields("PKID").Value & ""
  End If
  '
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  'NOVO BOTÕES NOVOS
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub

Private Sub grdLeiturista_UnboundReadDataEx( _
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
               Offset + intI, LEIT_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, LEIT_COLUNASMATRIZ, LEIT_LINHASMATRIZ, LEIT_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, LEIT_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserOperLeiCons.grdLeiturista_UnboundReadDataEx]"
End Sub

Private Sub Form_Activate()
  If blnPrimeiraVez Then
    DoEvents
    '
    'Montar RecordSet
    LEIT_COLUNASMATRIZ = grdLeiturista.Columns.Count
    LEIT_LINHASMATRIZ = 0
    MontaLEIT_Matriz
    grdLeiturista.Bookmark = Null
    grdLeiturista.ReBind
    grdLeiturista.ApproxCount = LEIT_LINHASMATRIZ
    '
    SetarFoco cmdSelecao(0)
  End If
End Sub
Public Sub MontaLEIT_Matriz()

  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMaq.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMaq.clsGeral
  '
  strSql = "SELECT LEITURAMAQUINA.PKID, MAQUINA.PKID, SERIE.NUMERO, TIPO.TIPO, EQUIPAMENTO.NUMERO, LEITURAMAQUINA.MEDICAOENTRADA, LEITURAMAQUINA.MEDICAOSAIDA "
  strSql = strSql & " FROM SERIE " & _
          " INNER JOIN EQUIPAMENTO ON SERIE.PKID = EQUIPAMENTO.SERIEID " & _
          " INNER JOIN MAQUINA ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
          " INNER JOIN TIPO ON TIPO.PKID = MAQUINA.TIPOID " & _
          " INNER JOIN LEITURAMAQUINA ON MAQUINA.PKID = LEITURAMAQUINA.MAQUINAID " & _
          " WHERE LEITURAMAQUINA.LEITURAID  = " & Formata_Dados(lngLEITURAID, tpDados_Longo) & _
          " ORDER BY SERIE.NUMERO, TIPO.TIPO, EQUIPAMENTO.NUMERO"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    LEIT_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim LEIT_Matriz(0 To LEIT_COLUNASMATRIZ - 1, 0 To LEIT_LINHASMATRIZ - 1)
  Else
    ReDim LEIT_Matriz(0 To LEIT_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To LEIT_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To LEIT_COLUNASMATRIZ - 1  'varre as colunas
          LEIT_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
