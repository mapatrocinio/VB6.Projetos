VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmUserLinhaAssocInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Associação de linhas ao associado"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   6975
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   150
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Associar linhas ao associado"
      TabPicture(0)   =   "userLinhaAssocInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(17)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(16)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "grdLinhaAssoc"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "grdLinha"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "cmdCadastraItem(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCadastraItem(1)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCadastraItem(2)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdCadastraItem(3)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      Begin VB.CommandButton cmdCadastraItem 
         Caption         =   "<<"
         Height          =   375
         Index           =   3
         Left            =   4440
         TabIndex        =   4
         Top             =   1860
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastraItem 
         Caption         =   "<"
         Height          =   375
         Index           =   2
         Left            =   4440
         TabIndex        =   3
         Top             =   1500
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastraItem 
         Caption         =   ">>"
         Height          =   375
         Index           =   1
         Left            =   4440
         TabIndex        =   2
         Top             =   1140
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastraItem 
         Caption         =   ">"
         Height          =   375
         Index           =   0
         Left            =   4440
         TabIndex        =   1
         Top             =   780
         Width           =   375
      End
      Begin TrueDBGrid60.TDBGrid grdLinha 
         Height          =   4890
         Left            =   90
         OleObjectBlob   =   "userLinhaAssocInc.frx":001C
         TabIndex        =   0
         Top             =   780
         Width           =   4260
      End
      Begin TrueDBGrid60.TDBGrid grdLinhaAssoc 
         Height          =   4890
         Left            =   4890
         OleObjectBlob   =   "userLinhaAssocInc.frx":31D6
         TabIndex        =   5
         Top             =   780
         Width           =   4260
      End
      Begin VB.Label Label1 
         Caption         =   "* Aperte a tecla <CTRL> OU <SHIFT> + Botão direito do mouse para selecionar mais de um item do grid."
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   5820
         Width           =   8715
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "* As unidades serão incluidas/excluidas automaticamente após ser pressionado os botões >, >>, < ou <<."
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   6060
         Width           =   8715
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Linhas associadas ao associado"
         Height          =   195
         Index           =   16
         Left            =   4920
         TabIndex        =   11
         Top             =   420
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Linhas"
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   10
         Top             =   420
         Width           =   2655
      End
   End
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7245
      Left            =   9660
      ScaleHeight     =   7245
      ScaleWidth      =   1860
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   1215
         Left            =   120
         ScaleHeight     =   1155
         ScaleWidth      =   1545
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   5850
         Width           =   1605
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "frmUserLinhaAssocInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lngASSOCIADOID         As Long
Public strNomeAssociado       As String

'Variáveis para Grid
'
Dim LINHA_COLUNASMATRIZ        As Long
Dim LINHA_LINHASMATRIZ         As Long
Private LINHA_Matriz()         As String
'
Dim LINHAASSOC_COLUNASMATRIZ   As Long
Dim LINHAASSOC_LINHASMATRIZ    As Long
Private LINHAASSOC_Matriz()    As String

Dim blnFechar                 As Boolean
Public blnRetorno             As Boolean

Private Sub cmdCadastraItem_Click(Index As Integer)
  TratarAssociacao Index + 1
  SetarFoco grdLinha
End Sub



Public Sub LINHAASSOC_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busApler.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busApler.clsGeral
  '
  strSql = "SELECT TAB_TITLINHA.PKID, TITULAR.ASSOCIADOID, LINHA.PKID, LINHA.DESCRICAO " & _
           "FROM TAB_TITLINHA INNER JOIN TITULAR ON TITULAR.ASSOCIADOID = TAB_TITLINHA.TITULARASSOCIADOID " & _
           " INNER JOIN LINHA ON LINHA.PKID = TAB_TITLINHA.LINHAID " & _
           "WHERE TITULAR.ASSOCIADOID = " & Formata_Dados(lngASSOCIADOID, tpDados_Longo) & _
           " ORDER BY LINHA.DESCRICAO"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    LINHAASSOC_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim LINHAASSOC_Matriz(0 To LINHAASSOC_COLUNASMATRIZ - 1, 0 To LINHAASSOC_LINHASMATRIZ - 1)
  Else
    ReDim LINHAASSOC_Matriz(0 To LINHAASSOC_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To LINHAASSOC_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To LINHAASSOC_COLUNASMATRIZ - 1  'varre as colunas
          LINHAASSOC_Matriz(intJ, intI) = objRs(intJ) & ""
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

Public Sub LINHA_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busApler.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busApler.clsGeral
  '
  strSql = "SELECT LINHA.PKID, LINHA.DESCRICAO " & _
           "FROM LINHA " & _
           " WHERE LINHA.PKID NOT IN (SELECT TAB_TITLINHA.LINHAID FROM TAB_TITLINHA " & _
           "WHERE TAB_TITLINHA.TITULARASSOCIADOID = " & Formata_Dados(lngASSOCIADOID, tpDados_Longo) & _
           " AND TAB_TITLINHA.LINHAID = LINHA.PKID) " & _
           " ORDER BY LINHA.DESCRICAO"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    LINHA_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim LINHA_Matriz(0 To LINHA_COLUNASMATRIZ - 1, 0 To LINHA_LINHASMATRIZ - 1)
  Else
    ReDim LINHA_Matriz(0 To LINHA_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To LINHA_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To LINHA_COLUNASMATRIZ - 1  'varre as colunas
          LINHA_Matriz(intJ, intI) = objRs(intJ) & ""
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

Private Sub cmdFechar_Click()
  '
  blnFechar = True
  Unload Me
End Sub

Private Sub grdLinhaAssoc_UnboundReadDataEx( _
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
               Offset + intI, LINHAASSOC_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, LINHAASSOC_COLUNASMATRIZ, LINHAASSOC_LINHASMATRIZ, LINHAASSOC_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, LINHAASSOC_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmLinhaAssocInc.grdLinhaAssoc_UnboundReadDataEx]"
End Sub



Private Sub grdLinha_UnboundReadDataEx( _
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
               Offset + intI, LINHA_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, LINHA_COLUNASMATRIZ, LINHA_LINHASMATRIZ, LINHA_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, LINHA_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmLinhaAssocInc.grdLinha_UnboundReadDataEx]"
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs As ADODB.Recordset
  Dim strSql As String
  '
  blnFechar = False 'Não Pode Fechar pelo X
  blnRetorno = False
  AmpS
  Me.Height = 7620
  Me.Width = 11610
  CenterForm Me
  Me.Caption = Me.Caption & " - " & strNomeAssociado
  
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar
  '
  LINHA_COLUNASMATRIZ = grdLinha.Columns.Count
  LINHA_LINHASMATRIZ = 0
  LINHA_MontaMatriz
  grdLinha.ApproxCount = LINHA_LINHASMATRIZ
  '
  '
  LINHAASSOC_COLUNASMATRIZ = grdLinhaAssoc.Columns.Count
  LINHAASSOC_LINHASMATRIZ = 0
  LINHAASSOC_MontaMatriz
  grdLinhaAssoc.ApproxCount = LINHAASSOC_LINHASMATRIZ
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub TratarAssociacao(pIndice As Integer)
  On Error GoTo trata
  Dim intI          As Long
  Dim objLinhaAssoc As busApler.clsLinhaAssoc
  Dim lngRet        As Long
  Dim blnRet        As Boolean
  Dim intExc        As Long
  '
  Set objLinhaAssoc = New busApler.clsLinhaAssoc
  '
  blnRet = False
  intExc = 0
  '
  Select Case pIndice
  Case 1 'Cadastrar Selecionados
    For intI = 0 To grdLinha.SelBookmarks.Count - 1
      grdLinha.Bookmark = CLng(grdLinha.SelBookmarks.Item(intI))
      objLinhaAssoc.AssociarLinhaAoAssociado grdLinha.Columns("LINHAID").Text, lngASSOCIADOID
      blnRet = True
    Next
  Case 2 'Cadastrar Todos
    For intI = 0 To LINHA_LINHASMATRIZ - 1
      grdLinha.Bookmark = CLng(intI)
      objLinhaAssoc.AssociarLinhaAoAssociado grdLinha.Columns("LINHAID").Text, lngASSOCIADOID
      blnRet = True
    Next
  Case 3 'Retirar Selecionados
    For intI = 0 To grdLinhaAssoc.SelBookmarks.Count - 1
      grdLinhaAssoc.Bookmark = CLng(grdLinhaAssoc.SelBookmarks.Item(intI))
      objLinhaAssoc.DesassociarLinhaDoAssociado grdLinhaAssoc.Columns("LINHAID").Text, lngASSOCIADOID
      blnRet = True
    Next
  Case 4 'retirar Todos
    For intI = 0 To LINHAASSOC_LINHASMATRIZ - 1
      grdLinhaAssoc.Bookmark = CLng(intI)
      If IsNull(grdLinhaAssoc.Bookmark) Then grdLinhaAssoc.Bookmark = CLng(intI)
      objLinhaAssoc.DesassociarLinhaDoAssociado grdLinhaAssoc.Columns("LINHAID").Text, lngASSOCIADOID
      blnRet = True
    Next
  End Select
  '
  Set objLinhaAssoc = Nothing
    '
  If blnRet Then 'Houve Auteração, Atualiza grids
    blnRetorno = True
    '
    LINHA_COLUNASMATRIZ = grdLinha.Columns.Count
    LINHA_LINHASMATRIZ = 0
    LINHA_MontaMatriz
    grdLinha.ApproxCount = LINHA_LINHASMATRIZ
    grdLinha.Bookmark = Null
    grdLinha.ReBind
    '
    LINHAASSOC_COLUNASMATRIZ = grdLinhaAssoc.Columns.Count
    LINHAASSOC_LINHASMATRIZ = 0
    LINHAASSOC_MontaMatriz
    grdLinhaAssoc.ApproxCount = LINHAASSOC_LINHASMATRIZ
    grdLinhaAssoc.Bookmark = Null
    grdLinhaAssoc.ReBind
    '
  End If
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub
