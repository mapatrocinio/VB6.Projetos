VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserGRPagamentoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11070
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   11070
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4410
      Left            =   9210
      ScaleHeight     =   4410
      ScaleWidth      =   1860
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2925
         Left            =   90
         ScaleHeight     =   2865
         ScaleWidth      =   1605
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   1380
         Width           =   1665
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&Z"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1020
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4185
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   9015
      _ExtentX        =   15901
      _ExtentY        =   7382
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Cadastro"
      TabPicture(0)   =   "userGRPagamentoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&GR´s"
      TabPicture(1)   =   "userGRPagamentoInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdGR"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3705
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   8775
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   3465
            Index           =   0
            Left            =   120
            ScaleHeight     =   3465
            ScaleWidth      =   8595
            TabIndex        =   10
            TabStop         =   0   'False
            Top             =   150
            Width           =   8595
            Begin VB.ComboBox cboPrestador 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   690
               Width           =   5385
            End
            Begin MSMask.MaskEdBox mskDtInicio 
               Height          =   255
               Left            =   1320
               TabIndex        =   0
               Top             =   90
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDtTermino 
               Height          =   255
               Left            =   1320
               TabIndex        =   1
               Top             =   390
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   -2147483637
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskHoraInicio 
               Height          =   255
               Left            =   2700
               TabIndex        =   14
               Top             =   90
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   -2147483637
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               MaxLength       =   8
               Mask            =   "##:##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskHoraTermino 
               Height          =   255
               Left            =   2700
               TabIndex        =   15
               Top             =   390
               Width           =   915
               _ExtentX        =   1614
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   -2147483637
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               MaxLength       =   8
               Mask            =   "##:##:##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label7 
               Caption         =   "Dt. Término"
               Height          =   255
               Index           =   2
               Left            =   90
               TabIndex        =   13
               Top             =   390
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Prestador"
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   12
               Top             =   660
               Width           =   1455
            End
            Begin VB.Label Label7 
               Caption         =   "Dt. Início"
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   11
               Top             =   90
               Width           =   1215
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdGR 
         Height          =   3660
         Left            =   -74910
         OleObjectBlob   =   "userGRPagamentoInc.frx":0038
         TabIndex        =   16
         Top             =   390
         Width           =   8790
      End
   End
End
Attribute VB_Name = "frmUserGRPagamentoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Public blnPrimeiraVez           As Boolean
Public icTipoGR                 As tpIcTipoGR
Public strGR                    As String

Public lngPKID                  As Long

Dim GR_COLUNASMATRIZ         As Long
Dim GR_LINHASMATRIZ          As Long
Private GR_Matriz()          As String

Private Sub TratarCampos()
  Dim strStatusFuncao     As String
  Dim strSql              As String
  On Error GoTo trata
'''  Dim intTopAux As Integer
'''  intTopAux = 2940
  Me.Caption = Me.Caption & " " & strGR
  
  If icTipoGR = tpIcTipoGR.tpIcTipoGR_Prest Then
    strStatusFuncao = "0"
  ElseIf icTipoGR = tpIcTipoGR.tpIcTipoGR_DonoRX Then
    strStatusFuncao = "2"
  ElseIf icTipoGR = tpIcTipoGR.tpIcTipoGR_DonoUltra Then
    strStatusFuncao = "1"
  ElseIf icTipoGR = tpIcTipoGR.tpIcTipoGR_TecRX Then
    strStatusFuncao = "3"
  ElseIf icTipoGR = tpIcTipoGR.tpIcTipoGR_CancPont Then
    strStatusFuncao = "0"
  ElseIf icTipoGR = tpIcTipoGR.tpIcTipoGR_CancAut Then
    strStatusFuncao = "0"
  End If
  strSql = "Select PRONTUARIO.NOME " & _
        " FROM PRONTUARIO INNER JOIN PRESTADOR ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID " & _
        " INNER JOIN FUNCAO ON FUNCAO.PKID = PRESTADOR.FUNCAOID " & _
        " WHERE PRESTADOR.INDEXCLUIDO = " & Formata_Dados("N", tpDados_Texto) & _
        " AND FUNCAO.STATUS = " & Formata_Dados(strStatusFuncao, tpDados_Texto) & _
        " ORDER BY PRONTUARIO.NOME"
  PreencheCombo cboPrestador, strSql, False, True, True
  '
  INCLUIR_VALOR_NO_MASK mskHoraInicio, "00:00:00", TpMaskData
  INCLUIR_VALOR_NO_MASK mskHoraTermino, "23:59:59", TpMaskData
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserGRPAgamentoInc.TratarCampos]", _
            Err.Description
End Sub

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Prontuario
  'Dados cadastrais
  LimparCampoMask mskDtInicio
  LimparCampoMask mskDtTermino
  LimparCampoMask mskHoraInicio
  LimparCampoMask mskHoraTermino
  LimparCampoCombo cboPrestador
  
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserGRPAgamentoInc.LimparCampos]", _
            Err.Description
End Sub




Private Sub cboPrestador_LostFocus()
  Pintar_Controle cboPrestador, tpCorContr_Normal
End Sub

Private Sub cmdCancelar_Click()
  blnFechar = True
  blnRetorno = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdIncluir_Click()
  On Error GoTo trata
  Dim objFormGRPgtoInc     As SisMed.frmUserGRPgtoInc
  '
  Select Case tabDetalhes.Tab
  Case 1
    'GR´S
    Set objFormGRPgtoInc = New SisMed.frmUserGRPgtoInc
    objFormGRPgtoInc.Status = tpStatus_Incluir
    objFormGRPgtoInc.icTipoGR = icTipoGR
    objFormGRPgtoInc.strGR = strGR
    objFormGRPgtoInc.lngGRPAGAMENTOID = lngPKID
    '
    objFormGRPgtoInc.strDataIni = mskDtInicio.Text
    objFormGRPgtoInc.strHoraIni = mskHoraInicio.Text
    objFormGRPgtoInc.strDataFim = mskDtTermino.Text
    objFormGRPgtoInc.strHoraFim = mskHoraTermino.Text
    objFormGRPgtoInc.strPrestador = cboPrestador.Text
    '
    objFormGRPgtoInc.Show vbModal

    If objFormGRPgtoInc.blnRetorno Then
      GR_MontaMatriz
      grdGR.Bookmark = Null
      grdGR.ReBind
      grdGR.ApproxCount = GR_LINHASMATRIZ
    End If
    Set objFormGRPgtoInc = Nothing
    SetarFoco grdGR
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Resume Next
End Sub

Private Sub cmdOk_Click()
  Dim grdGRPagamento              As busSisMed.clsGRPagamento
  Dim objGeral                    As busSisMed.clsGeral
  Dim objRs                       As ADODB.Recordset
  Dim strSql                      As String
  Dim lngPRESTADORID              As Long
  Dim strStatus                   As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisMed.clsGeral
  Set grdGRPagamento = New busSisMed.clsGRPagamento
  'PRESTADOR
  lngPRESTADORID = 0
  strSql = "SELECT PRONTUARIO.PKID " & _
      " FROM PRONTUARIO INNER JOIN PRESTADOR ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID " & _
      " WHERE PRONTUARIO.NOME = " & Formata_Dados(cboPrestador.Text, tpDados_Texto) & _
      " AND PRESTADOR.INDEXCLUIDO = " & Formata_Dados("N", tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngPRESTADORID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Verifica status
  Select Case icTipoGR
  Case tpIcTipoGR_DonoRX: strStatus = "DR"
  Case tpIcTipoGR_DonoUltra: strStatus = "DU"
  Case tpIcTipoGR_Prest: strStatus = "PG"
  Case tpIcTipoGR_TecRX: strStatus = "TR"
  Case tpIcTipoGR_CancPont: strStatus = "CP"
  Case tpIcTipoGR_CancAut: strStatus = "CA"
  Case Else: strStatus = ""
  End Select
  INCLUIR_VALOR_NO_MASK mskDtTermino, mskDtInicio.Text, TpMaskData
  'Validar se prestador / DATA já cadastrado
'''  'Por nome
'''  strSql = "SELECT * FROM GRPAGAMENTO " & _
'''    " WHERE GRPAGAMENTO.PRESTADORID = " & Formata_Dados(lngPRESTADORID, tpDados_Longo) & _
'''    " AND ((GRPAGAMENTO.DATAINICIO <= " & Formata_Dados(mskDtInicio.Text, tpDados_DataHora) & _
'''    " AND GRPAGAMENTO.DATATERMINO > " & Formata_Dados(mskDtInicio.Text, tpDados_DataHora) & ")" & _
'''    " OR (GRPAGAMENTO.DATAINICIO <= " & Formata_Dados(mskDtTermino.Text, tpDados_DataHora) & _
'''    " AND GRPAGAMENTO.DATATERMINO > " & Formata_Dados(mskDtTermino.Text, tpDados_DataHora) & "))" & _
'''    " AND GRPAGAMENTO.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo) & _
'''    " AND GRPAGAMENTO.STATUS = " & Formata_Dados(strStatus, tpDados_Texto)
'''  Set objRs = objGeral.ExecutarSQL(strSql)
'''  If Not objRs.EOF Then
'''    Pintar_Controle mskDtInicio, tpCorContr_Erro
'''    TratarErroPrevisto "Pagamento já associado a este prestador / faixa de datas"
'''    objRs.Close
'''    Set objRs = Nothing
'''    Set objGeral = Nothing
'''    Set grdGRPagamento = Nothing
'''    cmdOk.Enabled = True
'''    SetarFoco mskDtInicio
'''    tabDetalhes.Tab = 0
'''    Exit Sub
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Prontuario
    grdGRPagamento.AlterarGRPagamento lngPKID, _
                                      lngPRESTADORID, _
                                      IIf(mskDtInicio.ClipText = "", "", mskDtInicio.Text) & " " & mskHoraInicio, _
                                      IIf(mskDtTermino.ClipText = "", "", mskDtTermino.Text) & " " & mskHoraTermino
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Prontuario
    grdGRPagamento.InserirGRPagamento lngPKID, _
                                      lngPRESTADORID, _
                                      IIf(mskDtInicio.ClipText = "", "", mskDtInicio.Text) & " " & mskHoraInicio, _
                                      IIf(mskDtTermino.ClipText = "", "", mskDtTermino.Text) & " " & mskHoraTermino, _
                                      strStatus, _
                                      "N", _
                                      gsNomeUsu
                                     
    '
  End If
  If Status = tpStatus_Alterar Then
    blnRetorno = True
    blnFechar = True
    Unload Me
  ElseIf Status = tpStatus_Incluir Then
    'Selecionar prontuario pelo nome
    Status = tpStatus_Alterar
    'Reload na tela
    Form_Load
    'Acerta tabs
    tabDetalhes.TabEnabled(1) = True
    tabDetalhes.Tab = 1
    blnRetorno = True
  End If

  Set grdGRPagamento = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  cmdOk.Enabled = True
End Sub


Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  If Not Valida_Data(mskDtInicio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data início válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
'''  If Not Valida_Data(mskDtTermino, TpObrigatorio, blnSetarFocoControle) Then
'''    strMsg = strMsg & "Preencher a data início válida" & vbCrLf
'''    tabDetalhes.Tab = 0
'''  End If
  If Not Valida_String(cboPrestador, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o prestador" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserGRPAgamentoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserGRPAgamentoInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco mskDtInicio
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserGRPAgamentoInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objGRPagamento          As busSisMed.clsGRPagamento
'''  Dim objFuncionario              As busSisMed.clsFuncionario
'''  Dim objPrestador           As busSisMed.clsPrestador
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 4890
  Me.Width = 11160
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, , , cmdIncluir
  '
  'Limpar Campos
  LimparCampos
  'Tratar campos
  TratarCampos
  '
  'Função

  tabDetalhes_Click 1
  If Status = tpStatus_Incluir Then
    '
    tabDetalhes.TabEnabled(1) = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    '-----------------------------
    'PRONTUARIO
    '------------------------------
    Set objGRPagamento = New busSisMed.clsGRPagamento
    Set objRs = objGRPagamento.SelecionarGRPagamentoPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      'Prontuario
      'Dados cadastrais
      INCLUIR_VALOR_NO_MASK mskDtInicio, objRs.Fields("DATAINICIO").Value, TpMaskData
      INCLUIR_VALOR_NO_MASK mskDtTermino, objRs.Fields("DATATERMINO").Value, TpMaskData
      If objRs.Fields("NOME").Value & "" <> "" Then
        cboPrestador.Text = objRs.Fields("NOME").Value & ""
      End If
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub




Private Sub grdGR_UnboundReadDataEx( _
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
               Offset + intI, GR_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, GR_COLUNASMATRIZ, GR_LINHASMATRIZ, GR_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, GR_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserGRPAgamentoInc.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub mskDtInicio_GotFocus()
  Seleciona_Conteudo_Controle mskDtInicio
End Sub
Private Sub mskDtInicio_LostFocus()
  Pintar_Controle mskDtInicio, tpCorContr_Normal
End Sub

Private Sub mskDtTermino_GotFocus()
  Seleciona_Conteudo_Controle mskDtTermino
End Sub
Private Sub mskDtTermino_LostFocus()
  Pintar_Controle mskDtTermino, tpCorContr_Normal
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    'Dados cadastrais
    grdGR.Enabled = False
    pictrava(0).Enabled = True
    '
    cmdOk.Enabled = True
    cmdIncluir.Enabled = False
    '
    SetarFoco mskDtInicio
  Case 1
    'GR´S
    grdGR.Enabled = True
    pictrava(0).Enabled = False
    '
    cmdOk.Enabled = False
    cmdIncluir.Enabled = True
    '
    'Montar RecordSet
    GR_COLUNASMATRIZ = grdGR.Columns.Count
    GR_LINHASMATRIZ = 0
    GR_MontaMatriz
    grdGR.Bookmark = Null
    grdGR.ReBind
    grdGR.ApproxCount = GR_LINHASMATRIZ
    '
    SetarFoco grdGR
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "frmUserGRPAgamentoInc.tabDetalhes"
  AmpN
End Sub

Public Sub GR_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  Dim strStatus As String
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMed.clsGeral
  '
  strSql = "SELECT GRPGTO.PKID, MAX(PRONTUARIO.NOME) AS NOME, MAX(GR.SEQUENCIAL) AS SEQUENCIAL, MAX(GR.SENHA) AS SENHA, MAX(GR.DATA) AS DATA, MAX(GRPAGAMENTO.DATAINICIO) AS DATAPGTO, SUM(GRPROCEDIMENTO.VALOR) AS VALOR " & _
      " FROM GRPAGAMENTO INNER JOIN GRPGTO ON GRPAGAMENTO.PKID = GRPGTO.GRPAGAMENTOID " & _
      " INNER JOIN GR ON GR.PKID = GRPGTO.GRID " & _
      " INNER JOIN GRPROCEDIMENTO ON GR.PKID = GRPROCEDIMENTO.GRID " & _
      " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = ATENDE.PRONTUARIOID " & _
      " WHERE GRPAGAMENTO.PKID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
      " GROUP BY GRPGTO.PKID;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    GR_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim GR_Matriz(0 To GR_COLUNASMATRIZ - 1, 0 To GR_LINHASMATRIZ - 1)
  Else
    ReDim GR_Matriz(0 To GR_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To GR_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To GR_COLUNASMATRIZ - 1  'varre as colunas
          GR_Matriz(intJ, intI) = objRs(intJ) & ""
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


