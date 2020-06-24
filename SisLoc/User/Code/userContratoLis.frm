VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmUserContratoLis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de contratos"
   ClientHeight    =   5475
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11730
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5475
   ScaleWidth      =   11730
   Begin VB.TextBox txtContrato 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   7260
      MaxLength       =   100
      TabIndex        =   1
      TabStop         =   0   'False
      Text            =   "txtContrato"
      Top             =   90
      Width           =   2475
   End
   Begin VB.ComboBox cboObra 
      Height          =   315
      Left            =   1380
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   4485
   End
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5475
      Left            =   9870
      ScaleHeight     =   5475
      ScaleWidth      =   1860
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2145
         Left            =   60
         ScaleHeight     =   2085
         ScaleWidth      =   1635
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   3240
         Width           =   1695
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   1020
            Width           =   1335
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Height          =   4980
      Left            =   0
      OleObjectBlob   =   "userContratoLis.frx":0000
      TabIndex        =   2
      Top             =   480
      Width           =   9780
   End
   Begin VB.Label Label5 
      Caption         =   "Contrato"
      Height          =   195
      Index           =   1
      Left            =   6000
      TabIndex        =   8
      Top             =   105
      Width           =   1215
   End
   Begin VB.Label Label5 
      Caption         =   "Obra"
      Height          =   195
      Index           =   24
      Left            =   120
      TabIndex        =   7
      Top             =   90
      Width           =   1215
   End
End
Attribute VB_Name = "frmUserContratoLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String
Public strStatus         As String
'strStatus Asssue A - AVARIA; D - DEVOLUÇÃO




Private Sub cboObra_Click()
  On Error GoTo trata
  Dim lngCONTRATOID     As Long
  Dim lngOBRAID         As Long
  Dim objRs             As ADODB.Recordset
  Dim objGeral          As busSisLoc.clsGeral
  Dim strSql            As String
  '
  If cboObra.Text = "" Then txtContrato.Text = ""
  
  'CONTRATO
  lngCONTRATOID = 0
  lngOBRAID = 0
  Set objGeral = New busSisLoc.clsGeral
  strSql = "SELECT OBRA.PKID, OBRA.CONTRATOID, CONTRATO.NUMERO " & _
        " FROM OBRA " & _
        " INNER JOIN CONTRATO ON CONTRATO.PKID = OBRA.CONTRATOID " & _
        " WHERE OBRA.DESCRICAO = " & Formata_Dados(cboObra.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngOBRAID = objRs.Fields("PKID").Value
    lngCONTRATOID = objRs.Fields("CONTRATOID").Value
    txtContrato.Text = objRs.Fields("NUMERO").Value & ""
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  '
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0
  MontaMatriz lngCONTRATOID & "", lngOBRAID & ""
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  grdGeral.ApproxCount = LINHASMATRIZ
  '
  Exit Sub
trata:
  TratarErro Err.Description, _
             Err.Description, _
             Err.Source
End Sub

Private Sub cboObra_LostFocus()
  Pintar_Controle cboObra, tpCorContr_Normal
End Sub

Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Dim blnSetarFocoControle  As Boolean
  Dim objGeral              As busSisLoc.clsGeral
  Dim objUserDevInc         As SisLoc.frmUserDevolucaoInc
  Dim lngCONTRATOID         As Long
  Dim lngOBRAID             As Long
  Dim strSql                As String
  Dim objRs                 As ADODB.Recordset
  '
  blnSetarFocoControle = True
  If Not Valida_String(cboObra, TpObrigatorio, blnSetarFocoControle) Then
    MsgBox "Selecione a obra !", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  'CONTRATO
  lngCONTRATOID = 0
  lngOBRAID = 0
  Set objGeral = New busSisLoc.clsGeral
  strSql = "SELECT OBRA.PKID, OBRA.CONTRATOID, CONTRATO.NUMERO " & _
        " FROM OBRA " & _
        " INNER JOIN CONTRATO ON CONTRATO.PKID = OBRA.CONTRATOID " & _
        " WHERE OBRA.DESCRICAO = " & Formata_Dados(cboObra.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngOBRAID = objRs.Fields("PKID").Value
    lngCONTRATOID = objRs.Fields("CONTRATOID").Value
    txtContrato.Text = objRs.Fields("NUMERO").Value & ""
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  '
  If strStatus = "D" Then
    Set objUserDevInc = New SisLoc.frmUserDevolucaoInc
    objUserDevInc.Status = tpStatus_Incluir
    objUserDevInc.lngCONTRATOID = lngCONTRATOID
    objUserDevInc.lngOBRAID = lngOBRAID
    objUserDevInc.blnPrimeiraVez = True
    objUserDevInc.Show vbModal
    Set objUserDevInc = Nothing
  End If
  'If frmUserEstoqueInc.bRetorno Then
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0
  MontaMatriz lngCONTRATOID & "", lngOBRAID & ""
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  grdGeral.ApproxCount = LINHASMATRIZ
  'End If
  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Description, _
             Err.Description, _
             Err.Source
                      
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim strSql      As String
  AmpS
  Me.Height = 5955
  Me.Width = 11820
  
  CenterForm Me
  
  Me.Caption = Me.Caption
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar, , , , cmdAlterar
  
  'Obra
  LimparCampoCombo cboObra
  LimparCampoTexto txtContrato
  strSql = "Select DESCRICAO from OBRA ORDER BY DESCRICAO"
  PreencheCombo cboObra, strSql, False, True
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub


Public Sub MontaMatriz(Optional strContratoId As String, _
                       Optional strObraId As String)
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisLoc.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisLoc.clsGeral
  '
  strSql = "SELECT ITEMNF.ESTOQUEID, MIN(ESTOQUE.CODIGO) AS CODIGO, MIN(ESTOQUE.DESCRICAO) AS DESCRICAO, SUM(ITEMNF.QUANTIDADE) AS TOTAL, ISNULL(SUM(vw_cons_baixa.QTD_DEVOL), 0) AS QTD_DEVOL, ISNULL(SUM(vw_cons_baixa.QTD_AVARIA), 0) AS QTD_AVARIA, (SUM(ITEMNF.QUANTIDADE) - ISNULL(SUM(vw_cons_baixa.QTD_DEVOL), 0) - ISNULL(SUM(vw_cons_baixa.QTD_AVARIA), 0)) AS QTD_REAL "
  strSql = strSql & " FROM ITEMNF INNER JOIN ESTOQUE ON ESTOQUE.PKID = ITEMNF.ESTOQUEID " & _
          " INNER JOIN NF ON NF.PKID = ITEMNF.NFID " & _
          " INNER JOIN vw_cons_baixa ON ITEMNF.PKID = vw_cons_baixa.ITEMNFID " & _
          "WHERE NF.CONTRATOID = " & Formata_Dados(strContratoId, tpDados_Longo) & _
          " AND NF.OBRAID = " & Formata_Dados(strObraId, tpDados_Longo) & _
          " AND NF.STATUS IN ('F', 'S') " & _
          " GROUP BY ITEMNF.ESTOQUEID " & _
          "  " & _
          " ORDER BY ESTOQUE.DESCRICAO;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
'''  If Len(Trim(strSelecao & "")) <> 0 Then
'''    objRs.Filter = "DESCRICAO LIKE '*" & strSelecao & "*'"
'''    If objRs.EOF Then
'''      MsgBox "Não foram encontrados itens para esta seleção", vbExclamation, TITULOSISTEMA
'''      Set objRs = clsGer.ExecutarSQL(strSql)
'''    End If
'''  End If
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


