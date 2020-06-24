VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "TDBG6.OCX"
Begin VB.Form frmUserDespesaLis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de despesas e receitas"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   11355
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4980
      Left            =   9495
      ScaleHeight     =   4980
      ScaleWidth      =   1860
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   3585
         Left            =   60
         ScaleHeight     =   3525
         ScaleWidth      =   1635
         TabIndex        =   6
         Top             =   1230
         Width           =   1695
         Begin VB.CommandButton cmdAlterar 
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdInserir 
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1800
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2640
            Width           =   1335
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Align           =   3  'Align Left
      Height          =   4980
      Left            =   0
      OleObjectBlob   =   "userDespesaCtaLis.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   9480
   End
End
Attribute VB_Name = "frmUserDespesaLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String
Public lngTURNOID         As Long
Public strTipo            As String
Public strTipoCtaPagas    As String




Private Sub cmdAlterar_Click()
  
  If Not IsNumeric(grdGeral.Columns("ID").Value) Then
    MsgBox "Selecione uma despesa !", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  If Trim(grdGeral.Columns("TIPO").Value & "") = "T" Then
    frmUserDespesaInc.Status = tpStatus_Alterar
  Else
    frmUserDespesaInc.Status = tpStatus_Alterar
  End If
  frmUserDespesaInc.lngDESPESAID = grdGeral.Columns("ID").Value
  frmUserDespesaInc.strTipo = Trim(grdGeral.Columns("TIPO").Value & "")
  frmUserDespesaInc.Show vbModal
  
  If frmUserDespesaInc.bRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If

End Sub

Private Sub cmdExcluir_Click()
  Dim objDespesa As busSisContas.clsDespesa
  '
  On Error GoTo trata
  If Len(Trim(grdGeral.Columns("ID").Value)) = 0 Then
    MsgBox "Selecione uma despesa para exclusão.", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  If Trim(grdGeral.Columns("TIPO").Value & "") = "T" Then
    MsgBox "Esta despesa foi criada pelo Sismotel, não poderá ser excluida.", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  '
  If MsgBox("Confirma exclusão da despesa " & grdGeral.Columns("Descrição").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then Exit Sub
  'OK
  Set objDespesa = New busSisContas.clsDespesa
  
  objDespesa.ExcluirDespesa CLng(grdGeral.Columns("ID").Value)
  '
  MontaMatriz
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  
  Set objDespesa = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub


Private Sub cmdInserir_Click()

  frmUserDespesaInc.Status = tpStatus_Incluir
  frmUserDespesaInc.strTipo = "A" 'Administração
  frmUserDespesaInc.Show vbModal
  
  If frmUserDespesaInc.bRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If

End Sub

Private Sub Form_Load()
On Error GoTo trata
  AmpS
  Me.Height = 5355
  Me.Width = 11445
  
  CenterForm Me
  
  Me.Caption = Me.Caption
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar, cmdExcluir, , cmdInserir, cmdAlterar
  
  
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0

  MontaMatriz
  grdGeral.ApproxCount = LINHASMATRIZ
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, Err.Source
End Sub


Public Sub MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisContas.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisContas.clsGeral
  '
  If strTipoCtaPagas = "N" Then
    strSql = "SELECT DESPESA.PKID, DESPESA.TIPO, DESPESA.SEQUENCIAL, IIF(DESPESA.TIPO='T', 'Telefonista', IIF(DESPESA.TIPO='A', 'Administração','')), GRUPODESPESA.CODIGO + SUBGRUPODESPESA.CODIGO + ' - ' + SUBGRUPODESPESA.DESCRICAO, DESPESA.DESCRICAO, FORMAT(DESPESA.VR_PAGO, '###,##0.00') " & _
              "FROM (DESPESA LEFT JOIN SUBGRUPODESPESA ON SUBGRUPODESPESA.PKID = DESPESA.SUBGRUPODESPESAID) " & _
              " LEFT JOIN GRUPODESPESA ON GRUPODESPESA.PKID = SUBGRUPODESPESA.GRUPODESPESAID " & IIf(strTipo = "A", " WHERE DESPESA.TIPO='A'", "")
  Else
    strSql = "SELECT DESPESA.PKID, DESPESA.TIPO, DESPESA.SEQUENCIAL, IIF(DESPESA.TIPO='T', 'Telefonista', IIF(DESPESA.TIPO='A', 'Administração','')), GRUPODESPESA.CODIGO + SUBGRUPODESPESA.CODIGO + ' - ' + SUBGRUPODESPESA.DESCRICAO, DESPESA.DESCRICAO, FORMAT(DESPESA.VR_PAGO, '###,##0.00') " & _
              "FROM (DESPESA LEFT JOIN SUBGRUPODESPESA ON SUBGRUPODESPESA.PKID = DESPESA.SUBGRUPODESPESAID) " & _
              " LEFT JOIN GRUPODESPESA ON GRUPODESPESA.PKID = SUBGRUPODESPESA.GRUPODESPESAID " & _
              " WHERE NOT ISDATE(DESPESA.DT_PAGAMENTO) " & _
              IIf(strTipo = "A", " AND DESPESA.TIPO='A'", "")
  End If
  strSql = strSql & " ORDER BY DESPESA.SEQUENCIAL DESC;"
  '
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
  TratarErro Err.Number, Err.Description, "[frmUserPedidoLis.grdGeral_UnboundReadDataEx]"
End Sub

