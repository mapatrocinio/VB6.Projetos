VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmUserMovimentacaoLis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de Movimentações"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10260
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   10260
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4980
      Left            =   8400
      ScaleHeight     =   4980
      ScaleWidth      =   1860
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   3885
         Left            =   30
         ScaleHeight     =   3825
         ScaleWidth      =   1635
         TabIndex        =   6
         Top             =   1020
         Width           =   1695
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdInserir 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2760
            Width           =   1335
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Align           =   3  'Align Left
      Height          =   4980
      Left            =   0
      OleObjectBlob   =   "userMovimentacaoLis.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   8160
   End
End
Attribute VB_Name = "frmUserMovimentacaoLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String
Public lngTURNOID         As Long
Public strStatus          As String




Private Sub cmdAlterar_Click()
  
  If Not IsNumeric(grdGeral.Columns("ID").Value) Then
    MsgBox "Selecione uma movimentação !", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  frmUserMovimentacaoInc.Status = tpStatus_Alterar
  frmUserMovimentacaoInc.strStatus = strStatus
  frmUserMovimentacaoInc.lngMOVIMENTACAOID = grdGeral.Columns("ID").Value
  frmUserMovimentacaoInc.Show vbModal
  
  If frmUserMovimentacaoInc.bRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If

End Sub

Private Sub cmdExcluir_Click()
  Dim objMovimentacao      As busApler.clsMovimentacao
  '
  On Error GoTo trata
  If Len(Trim(grdGeral.Columns("ID").Value)) = 0 Then
    MsgBox "Selecione " & IIf(strStatus = "M", "uma movimentação ", "uma Conta") & " para exclusão.", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  '
  If MsgBox("Confirma exclusão " & IIf(strStatus = "M", "da movimentação ", IIf(strStatus = "D", "do ajuste de débito ", "do ajuste de crédito ")) & grdGeral.Columns("Descrição").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then Exit Sub
  'OK
  Set objMovimentacao = New busApler.clsMovimentacao
  
  objMovimentacao.ExcluirMovimentacao CLng(grdGeral.Columns("ID").Value)
  '
  MontaMatriz
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  
  Set objMovimentacao = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub


Private Sub cmdInserir_Click()

  frmUserMovimentacaoInc.Status = tpStatus_Incluir
  frmUserMovimentacaoInc.strStatus = strStatus
  frmUserMovimentacaoInc.Show vbModal
  If frmUserMovimentacaoInc.bRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If

End Sub

Private Sub Form_Load()
On Error GoTo trata
  AmpS
  Me.Height = 5355
  Me.Width = 10350
  
  CenterForm Me
  
  Me.Caption = Me.Caption
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar, cmdExcluir, , cmdInserir, cmdAlterar
    
  If strStatus = "M" Then
    Me.Caption = "Lista de Movimentações"
  ElseIf strStatus = "D" Then
    Me.Caption = "Lista de Ajustes - Débitos"
    grdGeral.Columns(2).Caption = "Conta Debitada"
    grdGeral.Columns(3).Visible = False
  ElseIf strStatus = "C" Then
    Me.Caption = "Lista de Ajustes - Créditos"
    grdGeral.Columns(2).Caption = "Conta Creditada"
    grdGeral.Columns(3).Visible = False
  End If
  
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
  Dim clsGer    As busApler.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busApler.clsGeral
  '
  If strStatus = "M" Then
    strSql = "SELECT MOVIMENTACAO.PKID, MOVIMENTACAO.DATA, CONTADEB.DESCRICAO, CONTACRED.DESCRICAO, MOVIMENTACAO.VALOR, MOVIMENTACAO.DESCRICAO " & _
              "FROM (MOVIMENTACAO LEFT JOIN CONTA AS CONTADEB ON CONTADEB.PKID = MOVIMENTACAO.CONTADEBITOID) " & _
              " LEFT JOIN CONTA AS CONTACRED ON CONTACRED.PKID = MOVIMENTACAO.CONTACREDITOID " & _
              " WHERE STATUS = " & Formata_Dados(strStatus, tpDados_Texto, tpNulo_NaoAceita) & _
              " AND MOVIMENTACAO.PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo) & _
              " ORDER BY MOVIMENTACAO.DATA DESC, MOVIMENTACAO.PKID DESC "
  ElseIf strStatus = "D" Then
    strSql = "SELECT MOVIMENTACAO.PKID, MOVIMENTACAO.DATA, CONTADEB.DESCRICAO, '', MOVIMENTACAO.VALOR, MOVIMENTACAO.DESCRICAO " & _
              "FROM MOVIMENTACAO LEFT JOIN CONTA AS CONTADEB ON CONTADEB.PKID = MOVIMENTACAO.CONTADEBITOID " & _
              " WHERE STATUS = " & Formata_Dados(strStatus, tpDados_Texto, tpNulo_NaoAceita) & _
              " AND MOVIMENTACAO.PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo) & _
              " ORDER BY MOVIMENTACAO.DATA DESC, MOVIMENTACAO.PKID DESC "
  ElseIf strStatus = "C" Then
    strSql = "SELECT MOVIMENTACAO.PKID, MOVIMENTACAO.DATA, CONTACRED.DESCRICAO, '', MOVIMENTACAO.VALOR, MOVIMENTACAO.DESCRICAO " & _
              "FROM MOVIMENTACAO LEFT JOIN CONTA AS CONTACRED ON CONTACRED.PKID = MOVIMENTACAO.CONTACREDITOID " & _
              " WHERE STATUS = " & Formata_Dados(strStatus, tpDados_Texto, tpNulo_NaoAceita) & _
              " AND MOVIMENTACAO.PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo) & _
              " ORDER BY MOVIMENTACAO.DATA DESC, MOVIMENTACAO.PKID DESC "
  End If
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
  TratarErro Err.Number, Err.Description, "[frmUserFormaPgto.grdGeral_UnboundReadDataEx]"
End Sub

