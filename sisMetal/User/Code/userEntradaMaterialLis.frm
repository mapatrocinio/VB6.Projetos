VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmEntradaMaterialLis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de entrada de material"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   10080
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4980
      Left            =   8220
      ScaleHeight     =   4980
      ScaleWidth      =   1860
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4695
         Left            =   60
         ScaleHeight     =   4635
         ScaleWidth      =   1635
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   180
         Width           =   1695
         Begin VB.CommandButton cmdFiltrar 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CommandButton cmdInserir 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1890
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2730
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   3630
            Width           =   1335
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Align           =   3  'Align Left
      Height          =   4980
      Left            =   0
      OleObjectBlob   =   "userEntradaMaterialLis.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   8160
   End
End
Attribute VB_Name = "frmEntradaMaterialLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String
Public strUnidade         As String
Private strNomeSuiteApto  As String
Private lngLOCACAOID      As Long
Private strStatus         As String
'strStatus Asssue C - COMPENSADO; D - DEVOLVIDO

Private Sub cmdAlterar_Click()
  On Error GoTo trata
  If Not IsNumeric(grdGeral.Columns("ID").Value) Then
    MsgBox "Selecione uma entrada de material !", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  
  frmEntradaMaterialInc.lngENTRADAMATERIALID = grdGeral.Columns("ID").Value
  frmEntradaMaterialInc.Status = tpStatus_Alterar
  frmEntradaMaterialInc.intQuemChamou = 1 'Chamada da Alteração (Exclusão)
  frmEntradaMaterialInc.Show vbModal
  
  If frmEntradaMaterialInc.blnRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If
  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Private Sub cmdExcluir_Click()
  Dim clsEntMat   As busSisMetal.clsEntradaMaterial
  Dim objRs       As ADODB.Recordset
  Dim strRetorno  As String
  '
  On Error GoTo trata
  If Len(Trim(grdGeral.Columns("ID").Value) & "") = 0 Then
    MsgBox "Selecione uma entrada de material.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  '
  Set clsEntMat = New busSisMetal.clsEntradaMaterial
  '
  strRetorno = clsEntMat.ValidarExclusaoEntradaMaterial(grdGeral.Columns("ID").Value)
  If strRetorno <> "" Then
    Set clsEntMat = Nothing
    TratarErroPrevisto "Não é possível excluir a entrada de material por constar referências nas tabelas de: " & vbCrLf & strRetorno & "." & vbCrLf & "Selecione o item, clique em alterar e em seguida exclua os itens da entrada de material.", "frmUserEntradaMaterialLis.cmdExcluir_Click"
    SetarFoco grdGeral
    Exit Sub
  End If
  If MsgBox("Confirma exclusão da entrada de material " & grdGeral.Columns("Descrição").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdGeral
    Exit Sub
  End If
  'OK
  clsEntMat.ExcluirENTRADAMATERIAL CLng(grdGeral.Columns("ID").Value)
  '
  MontaMatriz
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  
  Set clsEntMat = Nothing
  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub


Private Sub cmdFiltrar_Click()
  Dim strResp As String
  On Error GoTo trata
  strResp = InputBox("Entre com a descrição da entrada de material (deixe em branco para todos):", TITULOSISTEMA)
  '
  MontaMatriz strResp
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  '
  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub cmdInserir_Click()
  On Error GoTo trata

  frmEntradaMaterialInc.Status = tpStatus_Incluir
  frmEntradaMaterialInc.intQuemChamou = 0
  frmEntradaMaterialInc.lngENTRADAMATERIALID = 0
  frmEntradaMaterialInc.Show vbModal
  
  If frmEntradaMaterialInc.blnRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If
  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub Form_Load()
On Error GoTo trata
  AmpS
  Me.Height = 5355
  Me.Width = 10170
  
  CenterForm Me
  
  Me.Caption = Me.Caption
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar, cmdExcluir, , cmdInserir, cmdAlterar, cmdFiltrar
  
  
  COLUNASMATRIZ = 5
  LINHASMATRIZ = 0

  MontaMatriz
  grdGeral.ApproxCount = LINHASMATRIZ
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub


Public Sub MontaMatriz(Optional strSelecao As String)
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMetal.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMetal.clsGeral
  '
  strSql = "SELECT ENTRADAMATERIAL.PKID, ENTRADAMATERIAL.CODIGO,  ENTRADAMATERIAL.DATAAQUISICAO, min(ENTRADAMATERIAL.DESCRICAO)AS DESCRICAO, SUM(TAB_ENTRADAMATERIAL.QUANTIDADE) "
  strSql = strSql & " FROM ENTRADAMATERIAL LEFT JOIN TAB_ENTRADAMATERIAL ON ENTRADAMATERIAL.PKID = TAB_ENTRADAMATERIAL.ENTRADAMATERIALID " & _
            "GROUP BY ENTRADAMATERIAL.PKID, ENTRADAMATERIAL.CODIGO, ENTRADAMATERIAL.DATAAQUISICAO " & _
            "ORDER BY CODIGO DESC"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Len(Trim(strSelecao & "")) <> 0 Then
    objRs.Filter = "DESCRICAO LIKE '*" & strSelecao & "*'"
    If objRs.EOF Then
      MsgBox "Não foram encontrados itens para esta seleção", vbExclamation, TITULOSISTEMA
      Set objRs = clsGer.ExecutarSQL(strSql)
    End If
  End If
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


