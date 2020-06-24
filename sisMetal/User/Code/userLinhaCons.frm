VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmLinhaCons 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar linha"
   ClientHeight    =   6360
   ClientLeft      =   2580
   ClientTop       =   3105
   ClientWidth     =   7935
   Icon            =   "userLinhaCons.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6360
   ScaleWidth      =   7935
   Begin VB.CommandButton cmdNormal 
      Caption         =   "&Consultar"
      Height          =   255
      Index           =   0
      Left            =   6570
      TabIndex        =   8
      Top             =   120
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
      TabIndex        =   2
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
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   7932
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   6630
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdConfirmar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   5430
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   0
            Width           =   1215
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   $"userLinhaCons.frx":000C
            ForeColor       =   &H000000FF&
            Height          =   495
            Left            =   60
            TabIndex        =   7
            Top             =   60
            Width           =   5205
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Height          =   4335
      Left            =   0
      OleObjectBlob   =   "userLinhaCons.frx":009D
      TabIndex        =   1
      Top             =   960
      Width           =   7935
   End
   Begin VB.Label Label1 
      Caption         =   "Informe a descrição ou código da linha"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   1455
   End
End
Attribute VB_Name = "frmLinhaCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Enum tpIcOrigemLn
  tpIcOrigemLn_Ped = 0
  tpIcOrigemLn_Ins = 1
  tpIcOrigemLn_LnFr = 2
  tpIcOrigemLn_OS = 3
  tpIcOrigemLn_PedIt = 4
  tpIcOrigemLn_RelSalPerf = 5
  tpIcOrigemLn_Aju = 6
End Enum

Public intIcOrigemLn        As tpIcOrigemLn
Public strCodigoDescricao   As String
Dim blnFechar               As Boolean

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String
Dim blnPrimeiraVez        As Boolean


Private Sub cmdCancelar_Click()
  blnFechar = True
  strCodigoDescricao = ""
  Unload Me
End Sub

Private Sub cmdConfirmar_Click()
  On Error GoTo trata
  If Len(Trim(grdGeral.Columns(0).Value & "")) <> 0 Then
    If intIcOrigemLn = tpIcOrigemLn_Ped Then
      frmItemPedidoInc.txtCodigoFim.Text = grdGeral.Columns(0).Value
      frmItemPedidoInc.txtLinhaFim.Text = grdGeral.Columns(1).Value
      INCLUIR_VALOR_NO_MASK frmItemPedidoInc.mskComprimento, grdGeral.Columns(2).Value & "", TpMaskMoeda
    ElseIf intIcOrigemLn = tpIcOrigemLn_Ins Then
      frmInsumoInc.txtCodigoFim.Text = grdGeral.Columns(0).Value
      frmInsumoInc.txtLinhaFim.Text = grdGeral.Columns(1).Value
    ElseIf intIcOrigemLn = tpIcOrigemLn_LnFr Then
      frmLinhaFornecedorInc.txtCodigoFim.Text = grdGeral.Columns(0).Value
      frmLinhaFornecedorInc.txtLinhaFim.Text = grdGeral.Columns(1).Value
    ElseIf intIcOrigemLn = tpIcOrigemLn_OS Then
      frmItemOSInc.txtCodigoFim.Text = grdGeral.Columns(0).Value
      frmItemOSInc.txtLinhaFim.Text = grdGeral.Columns(1).Value
    ElseIf intIcOrigemLn = tpIcOrigemLn_PedIt Then
      frmPedidoItemPedidoInc.txtCodigoFim.Text = grdGeral.Columns(0).Value
      frmPedidoItemPedidoInc.txtLinhaFim.Text = grdGeral.Columns(1).Value
      frmPedidoItemPedidoInc.lngLINHAID = grdGeral.Columns(3).Value
    ElseIf intIcOrigemLn = tpIcOrigemLn_RelSalPerf Then
      frmRelSaldoPerfil.txtCodigoFim.Text = grdGeral.Columns(0).Value
      frmRelSaldoPerfil.txtLinhaFim.Text = grdGeral.Columns(1).Value
      frmRelSaldoPerfil.lngLINHAID = grdGeral.Columns(3).Value
    ElseIf intIcOrigemLn = tpIcOrigemLn_Aju Then
      frmAjusteInc.txtCodigoFim.Text = grdGeral.Columns(0).Value
      frmAjusteInc.txtLinhaFim.Text = grdGeral.Columns(1).Value
    End If
  End If
  blnFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub cmdNormal_Click(Index As Integer)
  Dim strWhere As String
  '
  If Len(Trim(txtItem.Text)) <> 0 Then
    strWhere = " TIPO_LINHA.NOME LIKE '%" & txtItem.Text & "%'" & _
        " OR LINHA.CODIGO LIKE '%" & txtItem.Text & "%'"
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

Public Sub MontaMatriz(Optional strWhere As String)
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMetal.clsGeral
  '
  AmpS
  On Error GoTo trata

  Set clsGer = New busSisMetal.clsGeral
  '
  strSql = "Select LINHA.CODIGO, TIPO_LINHA.NOME, VARA.VALOR, LINHA.PKID " & _
      " FROM LINHA " & _
      " LEFT JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID " & _
      " LEFT JOIN VARA ON VARA.PKID = LINHA.VARAID " & _
      " WHERE " & IIf(strWhere <> "", strWhere, "")
  strSql = strSql & " ORDER BY LINHA.CODIGO"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  '
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
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Activate()
  Dim strWhere As String
  If blnPrimeiraVez Then
    DoEvents
    '
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0
    strWhere = ""
    If Len(Trim(strCodigoDescricao)) <> 0 Then
      strWhere = " TIPO_LINHA.NOME LIKE '%" & strCodigoDescricao & "%'" & _
          " OR LINHA.CODIGO LIKE '%" & strCodigoDescricao & "%'"
    End If

    MontaMatriz strWhere
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = LINHASMATRIZ
    blnPrimeiraVez = False
    SetarFoco grdGeral
  End If
End Sub

Private Sub grdGeral_Click()
  'cmdConfirmar_Click
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
  TratarErro Err.Number, Err.Description, "[frmLinhaCons.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub Form_Load()
On Error GoTo trata
  AmpS
  Me.Height = 6765
  Me.Width = 8055
  blnPrimeiraVez = True
  blnFechar = False
  CenterForm Me
  LerFiguras Me, tpBmp_Vazio, cmdConfirmar, cmdCancelar
  txtItem.Text = strCodigoDescricao
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub



Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then
    Cancel = True
    Exit Sub
  End If
End Sub

Private Sub txtItem_GotFocus()
  Seleciona_Conteudo_Controle txtItem
End Sub
