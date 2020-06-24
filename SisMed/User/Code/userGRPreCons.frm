VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserGRPreCons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Controle de Consulta de GR´s"
   ClientHeight    =   8715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12000
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8715
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox cboPeriodo 
      Height          =   315
      Left            =   6570
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   390
      Width           =   5325
   End
   Begin VB.ComboBox cboPrestador 
      Height          =   315
      Left            =   6570
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   60
      Width           =   5325
   End
   Begin VB.Frame fraUnidade 
      Caption         =   "GR´s"
      Height          =   5685
      Left            =   60
      TabIndex        =   12
      Top             =   690
      Width           =   11835
      Begin TrueDBGrid60.TDBGrid grdGeral 
         Height          =   5430
         Left            =   90
         OleObjectBlob   =   "userGRPreCons.frx":0000
         TabIndex        =   4
         Top             =   210
         Width           =   11580
      End
   End
   Begin VB.CommandButton cmdSairSelecao 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   855
      Left            =   11070
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   6510
      Width           =   900
   End
   Begin VB.Frame Frame8 
      Caption         =   "Selecione a opção"
      Height          =   1725
      Left            =   60
      TabIndex        =   11
      Top             =   6420
      Width           =   10935
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&C - Importar Receitas    "
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
         Index           =   2
         Left            =   2970
         TabIndex        =   7
         ToolTipText     =   "Atender GR"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&B - Consultar Financ.     "
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
         Index           =   1
         Left            =   1380
         TabIndex        =   6
         ToolTipText     =   "Atender GR"
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton cmdSelecao 
         Caption         =   "&A - Scanner               "
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
         TabIndex        =   5
         ToolTipText     =   "Atender GR"
         Top             =   240
         Width           =   1305
      End
      Begin ComctlLib.StatusBar StatusBar1 
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1350
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
               TextSave        =   "19/4/2012"
               Key             =   ""
               Object.Tag             =   ""
            EndProperty
            BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
               Style           =   5
               Alignment       =   1
               Bevel           =   2
               Object.Width           =   1235
               MinWidth        =   1235
               TextSave        =   "02:46"
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
   Begin VB.TextBox txtUsuario 
      BackColor       =   &H00E0E0E0&
      Height          =   288
      Left            =   1380
      Locked          =   -1  'True
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "txtUsuario"
      Top             =   30
      Width           =   1695
   End
   Begin MSMask.MaskEdBox mskDataPrinc 
      Height          =   255
      Left            =   3990
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   60
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   450
      _Version        =   393216
      BackColor       =   14737632
      Enabled         =   0   'False
      MaxLength       =   10
      Format          =   "DD/MMM/YYYY"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Caption         =   "Últimos"
      Height          =   255
      Left            =   5340
      TabIndex        =   20
      Top             =   420
      Width           =   1155
   End
   Begin VB.Label lblLivro 
      Caption         =   "Prestador"
      Height          =   255
      Left            =   5340
      TabIndex        =   19
      Top             =   60
      Width           =   1155
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00FF00FF&
      Caption         =   "Atendida a posteriore"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   4290
      TabIndex        =   18
      Top             =   8220
      Width           =   2295
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00808000&
      Caption         =   "Liberada a posteriore"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   2310
      TabIndex        =   17
      Top             =   8220
      Width           =   1935
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H00008000&
      Caption         =   "Atendida"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   8
      Left            =   1590
      TabIndex        =   16
      Top             =   8220
      Width           =   675
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      BackColor       =   &H0080FF80&
      Caption         =   "Fechada"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   870
      TabIndex        =   15
      Top             =   8220
      Width           =   675
   End
   Begin VB.Label lblCor 
      Alignment       =   2  'Center
      Caption         =   "Status :"
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   60
      TabIndex        =   14
      Top             =   8220
      Width           =   765
   End
   Begin VB.Label Label16 
      Caption         =   "Data"
      Height          =   255
      Left            =   3150
      TabIndex        =   10
      Top             =   60
      Width           =   735
   End
   Begin VB.Label Label22 
      Caption         =   "Usuário Logado"
      Height          =   255
      Left            =   30
      TabIndex        =   9
      Top             =   30
      Width           =   1215
   End
End
Attribute VB_Name = "frmUserGRPreCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'''
Option Explicit

Public nGrupo                   As Integer
Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public blnPrimeiraVez           As Boolean 'Propósito: Preencher lista no combo

Private COLUNASMATRIZ           As Long
Private LINHASMATRIZ            As Long
Private Matriz()                As String

Private datDataIniAtual         As Date
Private datDataFimAtual         As Date


Private Sub cboPrestador_Click()
  On Error GoTo trata
  If blnPrimeiraVez = False Then
    DoEvents
    '
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0

    MontaMatriz cboPrestador.Text, _
                cboPeriodo.Text
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
    blnPrimeiraVez = False
    SetarFoco grdGeral
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source

End Sub

Private Sub cboPeriodo_Click()
  On Error GoTo trata
  If blnPrimeiraVez = False Then
    DoEvents
    '
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0

    MontaMatriz cboPrestador.Text, _
                cboPeriodo.Text
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
    blnPrimeiraVez = False
    SetarFoco grdGeral
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source

End Sub

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
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0

  MontaMatriz cboPrestador.Text, _
              cboPeriodo.Text
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  grdGeral.ApproxCount = LINHASMATRIZ
  blnPrimeiraVez = False
  SetarFoco grdGeral
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[cmdSelecao_Click]"
  frmMDI.tmrUnidade.Enabled = True
End Sub


Public Sub VerificaQuemChamou()
  Dim objUserAtendimentoInc         As SisMed.frmUserAtendimentoInc
  Dim objFormGRPgtoInc              As SisMed.frmUserGRPgtoInc
  '
  Dim datData                       As Date
  Dim strDataIni                    As String
  Dim strDataFim                    As String
  Dim strSql                        As String
  Dim objRs                         As ADODB.Recordset
  Dim objGeral                      As busSisMed.clsGeral
  Dim lngGRPAGAMENTOID              As Long
  '
  On Error GoTo trata
  '
  Select Case nGrupo

  Case 0
  
    'Atendimento GR
    If Len(Trim(grdGeral.Columns("ID").Value & "")) = 0 Then
      MsgBox "Selecione uma GR para atendimento.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    If Trim(grdGeral.Columns("Status").Value & "") <> "F" Then
      If Trim(grdGeral.Columns("Status").Value & "") <> "L" Then
        If Trim(grdGeral.Columns("Status").Value & "") <> "A" Then
          If Trim(grdGeral.Columns("Status").Value & "") <> "P" Then
            MsgBox "Apenas GR´s fechadas podem ser atendidas.", vbExclamation, TITULOSISTEMA
            SetarFoco grdGeral
            Exit Sub
          End If
        End If
      End If
    End If
    If gsNivel = gsArquivista Or gsNivel = gsAdmin Then
      If Trim(grdGeral.Columns("INDSCANER").Value & "") <> "S" Then
        MsgBox "GR não atendida ou o prestador não trabalha com Scanner.", vbExclamation, TITULOSISTEMA
        SetarFoco grdGeral
        Exit Sub
      End If
    End If
    Set objUserAtendimentoInc = New SisMed.frmUserAtendimentoInc

    'objUserAtendimentoInc.Status = tpStatus_Incluir
    'Carrega campos
    objUserAtendimentoInc.strHora = Format(grdGeral.Columns("Hora").Value & "", "DD/MM/YYYY hh:mm")
    objUserAtendimentoInc.strSequencial = Format(grdGeral.Columns("Seq.").Value & "", "###,000")
    objUserAtendimentoInc.strProntuario = grdGeral.Columns("Prontuário").Value & ""
    objUserAtendimentoInc.strEspecialidade = grdGeral.Columns("Especialidade").Value & ""
    objUserAtendimentoInc.strPrestador = grdGeral.Columns("Prestador").Value & ""
    objUserAtendimentoInc.strSala = grdGeral.Columns("Sala").Value & ""
    objUserAtendimentoInc.strAtendente = grdGeral.Columns("Atendente").Value & ""
    objUserAtendimentoInc.strStatus = grdGeral.Columns("Status").Value & ""
    objUserAtendimentoInc.lngGRID = grdGeral.Columns("ID").Value
    objUserAtendimentoInc.Show vbModal
    Set objUserAtendimentoInc = Nothing
  
  Case 1
    'Consulta recebimento
    'Verifica se irá incluir ou alterar grpagaento
    datData = Now
    strDataIni = Format(datData, "DD/MM/YYYY")
    strDataFim = Format(DateAdd("d", 1, datData), "DD/MM/YYYY")
    'Verifica status
    Set objGeral = New busSisMed.clsGeral
    strSql = "SELECT * FROM GRPAGAMENTO " & _
      " WHERE GRPAGAMENTO.PRESTADORID = " & Formata_Dados(giFuncionarioId, tpDados_Longo) & _
      " AND GRPAGAMENTO.DATAINICIO = " & Formata_Dados(strDataIni, tpDados_DataHora) & _
      " AND GRPAGAMENTO.STATUS = " & Formata_Dados("PG", tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    lngGRPAGAMENTOID = 0
    If Not objRs.EOF Then
      'Já cadastrado
      lngGRPAGAMENTOID = objRs.Fields("PKID").Value
    Else
      'Não cadastrado
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      MsgBox "Não há GR´s lançada para o prestador " & gsNomeUsuCompleto, vbExclamation, TITULOSISTEMA
      SetarFoco grdGeral
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    '
    Set objFormGRPgtoInc = New SisMed.frmUserGRPgtoInc
    objFormGRPgtoInc.Status = tpStatus_Consultar
    objFormGRPgtoInc.icTipoGR = tpIcTipoGR_Prest
    objFormGRPgtoInc.strGR = "GR Paga a prestador"
    objFormGRPgtoInc.lngGRPAGAMENTOID = lngGRPAGAMENTOID
    '
    objFormGRPgtoInc.strDataIni = strDataIni
    objFormGRPgtoInc.strHoraIni = "00:00"
    objFormGRPgtoInc.strDataFim = strDataFim
    objFormGRPgtoInc.strHoraFim = "23:59"
    objFormGRPgtoInc.strPrestador = gsNomeUsuCompleto
    '
    objFormGRPgtoInc.Show vbModal
  
  Case 2
    'Importar receitas Scanneadas
    Importar_Receitas gsPathLocal
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
  '
  'OK Para turno
  datDataIniAtual = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now) & " 00:00:00")
  datDataFimAtual = DateAdd("d", 1, datDataIniAtual)
  '
  blnFechar = False
  blnRetorno = False
  blnPrimeiraVez = True
  AmpS
  If Me.ActiveControl Is Nothing Then
    Me.Top = 580
    Me.Left = 1
    Me.WindowState = 2 'Maximizado
  End If
  'Me.Height = 9195
  'Me.Width = 12090
  'CenterForm Metual
  LerFigurasAvulsas cmdSairSelecao, "Sair.ico", "SairDown.ico", "Sair"
  '
  strSql = "SELECT PRONTUARIO.NOME FROM PRONTUARIO " & _
      " INNER JOIN PRESTADOR ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID " & _
      " WHERE PRESTADOR.INDEXCLUIDO = " & Formata_Dados("N", tpDados_Texto) & _
      " ORDER BY PRONTUARIO.NOME;"
  If gsNivel = gsArquivista Or gsNivel = gsAdmin Then
    PreencheCombo cboPrestador, strSql, False, True
    cboPrestador.Enabled = True
    cboPeriodo.Enabled = True
  ElseIf gsNivel = gsPrestador Then
    PreencheCombo cboPrestador, strSql, False, True, strItemSel:=gsNomeUsuCompleto
    cboPrestador.Enabled = False
    cboPeriodo.Enabled = False
  End If
  cboPeriodo.AddItem "Data atual"
  cboPeriodo.AddItem "05 últimos dias"
  cboPeriodo.AddItem "10 últimos dias"
  cboPeriodo.AddItem "20 últimos dias"
  cboPeriodo.AddItem "30 últimos dias"
  cboPeriodo.AddItem "60 últimos dias"
  cboPeriodo.AddItem "90 últimos dias"
  cboPeriodo.Text = "Data atual"
  '
  txtUsuario.Text = gsNomeUsu
  mskDataPrinc.Text = Format(Date, "DD/MM/YYYY")
  'Configurações
  If gsNivel = gsArquivista Or gsNivel = gsAdmin Then
    cmdSelecao(0).Caption = "&A - Scanner               "
    cmdSelecao(1).Enabled = False
    cmdSelecao(2).Enabled = True
  ElseIf gsNivel = gsPrestador Then
    cmdSelecao(0).Caption = "&A - Atender GR        "
    cmdSelecao(1).Enabled = True
    cmdSelecao(2).Enabled = False
  End If
  'NOVO BOTÕES NOVOS
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
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
  TratarErro Err.Number, Err.Description, "[frmUserGRPreCons.grdGeral_UnboundReadDataEx]"
End Sub



Private Sub Form_Activate()
  If blnPrimeiraVez Then
    DoEvents
    '
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0

    MontaMatriz cboPrestador.Text, _
                cboPeriodo.Text
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
    blnPrimeiraVez = False
    SetarFoco grdGeral
  End If
End Sub

Public Sub MontaMatriz(strPrestador As String, _
                       strPeriodo As String)
  Dim strSql      As String
  Dim objRs       As ADODB.Recordset
  Dim intI        As Integer
  Dim intJ        As Integer
  Dim objGR       As busSisMed.clsGR
  Dim strPerFinal As String
  '
  AmpS
  On Error GoTo trata
  '
  Set objGR = New busSisMed.clsGR
  '
  'Tratar período
  If Not IsNumeric(Left(strPeriodo, 2)) Then
    strPerFinal = "01"
  Else
    strPerFinal = Left(strPeriodo, 2)
  End If
  'A data inicial passa a ser calculada de acordo com o período informado
  datDataIniAtual = DateAdd("d", CInt(strPerFinal) * -1, datDataFimAtual)
  Set objRs = objGR.CapturaGRTurnoCorrentePRE(giFuncionarioId, _
                                              Format(datDataIniAtual, "DD/MM/YYYY hh:mm"), _
                                              Format(datDataFimAtual, "DD/MM/YYYY hh:mm"), _
                                              gsNivel, _
                                              strPrestador)
  If Not objRs.EOF Then
    'objRs.Filter = "STATUS = 'F' OR STATUS = 'A'"
    LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim Matriz(0 To COLUNASMATRIZ - 1, 0 To LINHASMATRIZ - 1)
  Else
    ReDim Matriz(0 To COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To COLUNASMATRIZ - 1  'varre as colunas
          Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGR = Nothing
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
