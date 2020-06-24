VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmUserEstoqueCons 
   Caption         =   "Consulta de itens do estoque geral"
   ClientHeight    =   6360
   ClientLeft      =   2595
   ClientTop       =   3120
   ClientWidth     =   7935
   Icon            =   "userEstoqueCons.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6360
   ScaleWidth      =   7935
   Begin VB.CommandButton cmdFiltrar 
      Caption         =   "&Z"
      Height          =   880
      Left            =   6570
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   30
      Width           =   1215
   End
   Begin VB.TextBox txtItem 
      Height          =   525
      Left            =   1800
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   120
      Width           =   4575
   End
   Begin VB.PictureBox picBotoes 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   7935
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   5460
      Width           =   7935
      Begin VB.PictureBox picAlinDir 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   912
         Left            =   0
         ScaleHeight     =   915
         ScaleWidth      =   7935
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   7932
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   6570
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   30
            Width           =   1215
         End
         Begin VB.CommandButton cmdConfirmar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   5370
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   30
            Width           =   1215
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Height          =   4335
      Left            =   0
      OleObjectBlob   =   "userEstoqueCons.frx":000C
      TabIndex        =   2
      Top             =   960
      Width           =   7935
   End
   Begin VB.Label Label1 
      Caption         =   "Informe a descrição do Item do Estoque"
      Height          =   375
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmUserEstoqueCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public QuemChamou As Integer
'Assume 0 Entrada de Material
'Assume 1 Inclusão de NF
Public strCodigoEstoque   As String

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String
Dim blnPrimeiraVez        As Boolean

Private Sub cmdCancelar_Click()
  strCodigoEstoque = ""
  Unload Me
End Sub

Private Sub cmdConfirmar_Click()
  If grdGeral.Columns(0).Value & "" = "" Then
    TratarErroPrevisto "Selecionar um produto", "cmdOK_Click"
    Pintar_Controle txtItem, tpCorContr_Erro
    SetarFoco txtItem
    Exit Sub
  End If
  Select Case QuemChamou
  Case 0 'Chamada de Entrada de Materiais
    frmUserEntradaMaterialInc.mskCodigo.Text = grdGeral.Columns(0).Value
  Case 1
    frmUserNFCons.objUserNFInc.txtPecaFim.Text = grdGeral.Columns(1).Value
  End Select
  '
  cmdCancelar_Click
End Sub

Private Sub cmdFiltrar_Click()
  Dim strWhere As String
  '
  If Len(Trim(txtItem.Text)) <> 0 Then
    strWhere = txtItem.Text
  End If
  '
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0

  MontaMatriz strWhere
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  grdGeral.ApproxCount = LINHASMATRIZ

  grdGeral.SetFocus
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

Private Sub Form_Activate()
  If blnPrimeiraVez Then
    DoEvents
    '
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0

    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
    blnPrimeiraVez = False
    '
    If LINHASMATRIZ = 0 Then
      SetarFoco txtItem
    Else
      SetarFoco grdGeral
    End If
  End If
End Sub
Public Sub MontaMatriz(Optional strWhere As String)
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisLoc.clsGeral
  Dim objEstoque  As busSisLoc.clsEstoque
  '
  AmpS
  On Error GoTo trata
  
  Set objEstoque = New busSisLoc.clsEstoque
  '
  If QuemChamou = 1 Then 'chamada da inclusão de NF
    If Len(Trim(strWhere)) = 0 Then
      Set objRs = objEstoque.CapturaEstoquePeloCodigo(strCodigoEstoque)
    Else
      Set objRs = objEstoque.CapturaEstoquePeloCodigo(strWhere)
    End If
  Else
    Set objRs = objEstoque.CapturaEstoquePeloCodigo(strWhere)
  End If
  'objRs.Filter = strWhere
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
  objRs.Close
  Set objRs = Nothing
  Set objEstoque = Nothing
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
Private Sub Form_Load()
  On Error GoTo trata
  AmpS
  Me.Height = 6765
  Me.Width = 8055
  blnPrimeiraVez = True
  
  CenterForm Me
  
  Me.Caption = Me.Caption
  LerFiguras Me, tpBmp_Vazio, cmdConfirmar, , cmdCancelar, , , , , cmdFiltrar
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub


Private Sub txtItem_GotFocus()
  Seleciona_Conteudo_Controle txtItem
End Sub
