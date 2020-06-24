VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserGRPagamentoLis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de "
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9915
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   9915
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   675
      Left            =   0
      ScaleHeight     =   675
      ScaleWidth      =   9915
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   9915
      Begin VB.TextBox txtProntuario 
         Height          =   285
         Left            =   840
         TabIndex        =   1
         Text            =   "txtProntuario"
         Top             =   300
         Width           =   7575
      End
      Begin VB.CommandButton cmdNormal 
         Caption         =   "&Consultar"
         Height          =   255
         Index           =   0
         Left            =   8550
         TabIndex        =   2
         Top             =   300
         Width           =   1215
      End
      Begin MSMask.MaskEdBox mskDtInicio 
         Height          =   255
         Left            =   840
         TabIndex        =   0
         Top             =   30
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label6 
         Caption         =   "Prontuário"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   12
         Top             =   300
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "Data"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   11
         Top             =   30
         Width           =   645
      End
   End
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4305
      Left            =   8055
      ScaleHeight     =   4305
      ScaleWidth      =   1860
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   675
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   3795
         Left            =   90
         ScaleHeight     =   3735
         ScaleWidth      =   1635
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   270
         Width           =   1695
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdInserir 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Align           =   3  'Align Left
      Height          =   4305
      Left            =   0
      OleObjectBlob   =   "userGRPagamentoLis.frx":0000
      TabIndex        =   3
      Top             =   675
      Width           =   7980
   End
End
Attribute VB_Name = "frmUserGRPagamentoLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String
Public icTipoGR           As tpIcTipoGR
Public strGR              As String

Private Sub cmdAlterar_Click()
  Dim strDataIni As String
  '
  On Error GoTo trata
  If Not IsNumeric(grdGeral.Columns("ID").Value) Then
    MsgBox "Selecione um pagamento de GR !", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  
  
  frmUserGRPagamentoInc.Status = tpStatus_Alterar
  frmUserGRPagamentoInc.lngPKID = grdGeral.Columns("ID").Value
  frmUserGRPagamentoInc.icTipoGR = icTipoGR
  frmUserGRPagamentoInc.strGR = strGR
  frmUserGRPagamentoInc.Show vbModal
  
  If frmUserGRPagamentoInc.blnRetorno Then
    strDataIni = ""
    If mskDtInicio.Text <> "__/__/____" Then
      strDataIni = mskDtInicio.Text & " 00:00"
    Else
      strDataIni = ""
    End If
    
    MontaMatriz strDataIni, _
                txtProntuario.Text
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
  End If
  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub


Private Sub cmdExcluir_Click()
  Dim objGRPagamento As busSisMed.clsGRPagamento
  Dim objGer As busSisMed.clsGeral
  Dim objRs As ADODB.Recordset
  Dim strSql As String
  Dim strDataIni As String
  '
  On Error GoTo trata
  If Not IsNumeric(grdGeral.Columns("ID").Value) Then
    MsgBox "Selecione um pagamento de GR para exclusão!", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  '
  Set objGer = New busSisMed.clsGeral
  '
'''  strSql = "Select * from LOCACAO where CARTAOID = " & grdGeral.Columns("ID").Value
'''  Set objRs = objGer.ExecutarSQL(strSql)
'''  If Not objRs.EOF Then
'''    objRs.Close
'''    Set objRs = Nothing
'''    Set objGer = Nothing
'''    TratarErroPrevisto "Cartão não pode ser excluido, pois já está associado a locações.", "frmUserCartaoLis.cmdExcluir_Click"
'''    SetarFoco grdGeral
'''    Exit Sub
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
  '
  Set objGer = Nothing
  If MsgBox("Confirma exclusão do pagamento da GR " & grdGeral.Columns("Início").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdGeral
    Exit Sub
  End If
  'OK
  Set objGRPagamento = New busSisMed.clsGRPagamento
  
  objGRPagamento.ExcluirGRPagamento CLng(grdGeral.Columns("ID").Value)
  '
  strDataIni = ""
  If mskDtInicio.Text <> "__/__/____" Then
    strDataIni = mskDtInicio.Text & " 00:00"
  Else
    strDataIni = ""
  End If
  
  MontaMatriz strDataIni, _
              txtProntuario.Text
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  grdGeral.ApproxCount = LINHASMATRIZ
  
  Set objGRPagamento = Nothing
  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub cmdInserir_Click()
  Dim strDataIni As String
  On Error GoTo trata

  frmUserGRPagamentoInc.Status = tpStatus_Incluir
  frmUserGRPagamentoInc.lngPKID = 0
  frmUserGRPagamentoInc.icTipoGR = icTipoGR
  frmUserGRPagamentoInc.strGR = strGR
  frmUserGRPagamentoInc.Show vbModal
  
  If frmUserGRPagamentoInc.blnRetorno Then
    strDataIni = ""
    If mskDtInicio.Text <> "__/__/____" Then
      strDataIni = mskDtInicio.Text & " 00:00"
    Else
      strDataIni = ""
    End If
    
    MontaMatriz strDataIni, _
                txtProntuario.Text
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
  End If
  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub




Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  If Not Valida_Data(mskDtInicio, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data válida" & vbCrLf
  End If
  If mskDtInicio.Text = "__/__/____" And txtProntuario.Text = "" Then
    strMsg = strMsg & "Preencher a data ou prontuário" & vbCrLf
    SetarFoco mskDtInicio
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserGRPagamentoLis.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserGRPagamentoLis.ValidaCampos]", _
            Err.Description
End Function

Private Sub cmdNormal_Click(Index As Integer)
  Dim strDataIni             As String
  On Error GoTo trata
  '
  If Index = 0 Then
    If Not ValidaCampos Then
      Exit Sub
    End If
    strDataIni = ""
    If mskDtInicio.Text <> "__/__/____" Then
      strDataIni = mskDtInicio.Text & " 00:00"
    Else
      strDataIni = ""
    End If
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0
    MontaMatriz strDataIni, _
                txtProntuario.Text
                   
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
  
    grdGeral.SetFocus
  End If
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source, _
            Err.Description
End Sub


Private Sub Form_Activate()
  SetarFoco grdGeral
End Sub

Private Sub Form_Load()
  Dim lngHeight As Long
  Dim lngWidth As Long
  
  On Error GoTo trata
  
  AmpS
  
  If Me.ActiveControl Is Nothing Then
    'Tela
    Me.Height = Screen.Height - 1450
    Me.Width = 10005
    CenterForm Me
  End If
  
  Me.Caption = Me.Caption & strGR
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar, cmdExcluir, , cmdInserir, cmdAlterar
  
  LimparCampoMask mskDtInicio
  LimparCampoTexto txtProntuario
  INCLUIR_VALOR_NO_MASK mskDtInicio, Format(Now, "DD/MM/YYYY"), TpMaskData
  
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0

  MontaMatriz mskDtInicio
  grdGeral.ApproxCount = LINHASMATRIZ
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub


Public Sub MontaMatriz(strDataIni As String, _
                       Optional strProntuario As String)
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
  strSql = "SELECT GRPAGAMENTO.PKID, GRPAGAMENTO.DATAINICIO, GRPAGAMENTO.DATATERMINO, PRONTUARIO.NOME " & _
            "FROM GRPAGAMENTO INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GRPAGAMENTO.PRESTADORID " & _
            " WHERE GRPAGAMENTO.STATUS = " & Formata_Dados(strStatus, tpDados_Texto)
  '
  If strProntuario & "" <> "" Then
    strSql = strSql & " AND PRONTUARIO.NOME LIKE " & Formata_Dados(strProntuario & "%", tpDados_Texto)
  End If
  If strDataIni & "" <> "" Then
    'strSql = strSql & " AND (GRPAGAMENTO.DATAINICIO >= " & Formata_Dados(strDataIni, tpDados_DataHora) & _
                      " OR GRPAGAMENTO.DATATERMINO >= " & Formata_Dados(strDataIni, tpDados_DataHora) & ")"
    
    strSql = strSql & " AND GRPAGAMENTO.DATAINICIO = " & Formata_Dados(strDataIni, tpDados_DataHora)
  End If
  strSql = strSql & " ORDER BY GRPAGAMENTO.DATAINICIO DESC, PRONTUARIO.NOME"
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
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
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
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
  TratarErro Err.Number, Err.Description, "[frmUserGRPagamentoLis.grdGeral_UnboundReadDataEx]"
End Sub


Private Sub mskDtInicio_GotFocus()
  Seleciona_Conteudo_Controle mskDtInicio
End Sub
Private Sub mskDtInicio_LostFocus()
  Pintar_Controle mskDtInicio, tpCorContr_Normal
End Sub

Private Sub txtProntuario_GotFocus()
  Seleciona_Conteudo_Controle txtProntuario
End Sub
Private Sub txtProntuario_LostFocus()
  Pintar_Controle txtProntuario, tpCorContr_Normal
End Sub

