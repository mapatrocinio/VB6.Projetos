VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmServicoAssoc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Associa��o de Servi�os ao Pacote"
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
      Top             =   120
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Associar Servi�os"
      TabPicture(0)   =   "userServicoAssoc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(17)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label1(16)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label1(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label1(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label6(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "grdUnidadeAssoc"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "grdUnidade"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cmdCadastraItem(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "cmdCadastraItem(1)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdCadastraItem(2)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "cmdCadastraItem(3)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "cboVeiculo"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).ControlCount=   12
      Begin VB.ComboBox cboVeiculo 
         Height          =   315
         Left            =   1470
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   360
         Width           =   5955
      End
      Begin VB.CommandButton cmdCadastraItem 
         Caption         =   "<<"
         Height          =   375
         Index           =   3
         Left            =   8940
         TabIndex        =   4
         Top             =   2100
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastraItem 
         Caption         =   "<"
         Height          =   375
         Index           =   2
         Left            =   8940
         TabIndex        =   2
         Top             =   1320
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastraItem 
         Caption         =   ">>"
         Height          =   375
         Index           =   1
         Left            =   8940
         TabIndex        =   3
         Top             =   1710
         Visible         =   0   'False
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastraItem 
         Caption         =   ">"
         Height          =   375
         Index           =   0
         Left            =   8940
         TabIndex        =   1
         Top             =   960
         Width           =   375
      End
      Begin TrueDBGrid60.TDBGrid grdUnidade 
         Height          =   2340
         Left            =   90
         OleObjectBlob   =   "userServicoAssoc.frx":001C
         TabIndex        =   0
         Top             =   960
         Width           =   8790
      End
      Begin TrueDBGrid60.TDBGrid grdUnidadeAssoc 
         Height          =   2670
         Left            =   90
         OleObjectBlob   =   "userServicoAssoc.frx":3D7C
         TabIndex        =   5
         Top             =   3570
         Width           =   9060
      End
      Begin VB.Label Label6 
         Caption         =   "Ve�culo"
         Height          =   255
         Index           =   0
         Left            =   150
         TabIndex        =   15
         Top             =   390
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "* Aperte a tecla <CTRL> OU <SHIFT> + Bot�o direito do mouse para selecionar mais de um item do grid."
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   6360
         Width           =   8715
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "* As unidades ser�o incluidas/excluidas automaticamente ap�s ser pressionado os bot�es >, >>, < ou <<."
         ForeColor       =   &H000000FF&
         Height          =   315
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   6600
         Width           =   8715
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label1 
         Caption         =   "Servi�os associados"
         Height          =   195
         Index           =   16
         Left            =   120
         TabIndex        =   11
         Top             =   3330
         Width           =   3375
      End
      Begin VB.Label Label1 
         Caption         =   "Servi�os n�o associados"
         Height          =   195
         Index           =   17
         Left            =   150
         TabIndex        =   10
         Top             =   690
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
Attribute VB_Name = "frmServicoAssoc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnRetorno               As Boolean
Dim blnFechar                   As Boolean
Public strCaption               As String
'''
'''Public strTipo                As String
Public lngPACOTEID              As Long
'Vari�veis para Grid
'
Dim UNID_COLUNASMATRIZ        As Long
Dim UNID_LINHASMATRIZ         As Long
Private UNID_Matriz()         As String
'
Dim UNIDASSOC_COLUNASMATRIZ   As Long
Dim UNIDASSOC_LINHASMATRIZ    As Long
Private UNIDASSOC_Matriz()    As String

Private blnMudarColTab As Boolean
Private blnEstaAlterando As Boolean

Private blnAtualizarAposLostFocus As Boolean


Private Sub cboVeiculo_LostFocus()
  Pintar_Controle cboVeiculo, tpCorContr_Normal
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs As ADODB.Recordset
  Dim strSql As String
  '
  blnAtualizarAposLostFocus = True 'Pronto para atualizar
  blnEstaAlterando = False
  blnMudarColTab = False
  blnFechar = False 'N�o Pode Fechar pelo X
  blnRetorno = False
  AmpS
  Me.Height = 7620
  Me.Width = 11610
  CenterForm Me
  Me.Caption = Me.Caption & " - " & strCaption
  'Combos
  'VAICULO
  strSql = "Select MODELO.NOME + ' (' + VEICULO.PLACA + ')' " & _
      " FROM VEICULO " & _
      " INNER JOIN MODELO ON MODELO.PKID = VEICULO.MODELOID " & _
      "ORDER BY MODELO.NOME, VEICULO.PLACA"
  PreencheCombo cboVeiculo, strSql, False, True
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar
  '
  UNID_COLUNASMATRIZ = grdUnidade.Columns.Count
  UNID_LINHASMATRIZ = 0
  UNID_MontaMatriz
  grdUnidade.ApproxCount = UNID_LINHASMATRIZ
  '
  grdUnidade.Columns(0).Locked = True
  '
  UNIDASSOC_COLUNASMATRIZ = grdUnidadeAssoc.Columns.Count
  UNIDASSOC_LINHASMATRIZ = 0
  UNIDASSOC_MontaMatriz
  grdUnidadeAssoc.ApproxCount = UNIDASSOC_LINHASMATRIZ
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub cmdCadastraItem_Click(Index As Integer)
  TratarAssociacao Index + 1
  SetarFoco grdUnidade
End Sub



Public Sub UNIDASSOC_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busElite.clsGeral
  '
  On Error GoTo trata

  Set clsGer = New busElite.clsGeral
  '
  strSql = "SELECT PACOTESERVICO.PKID, SERVICO.DATAHORA, " & _
           " AGENCIA.NOME + ' (' + dbo.formataCNPJ(AGENCIACNPJ.CNPJ) + ')', " & _
           " SERVICO.SOLICITANTE, SERVICO.STATUS " & _
           "FROM PACOTESERVICO " & _
           " INNER JOIN SERVICO ON SERVICO.PKID = PACOTESERVICO.SERVICOID " & _
           " INNER JOIN AGENCIACNPJ ON AGENCIACNPJ.PKID = SERVICO.AGENCIACNPJID " & _
           " INNER JOIN AGENCIA ON AGENCIA.PKID = AGENCIACNPJ.AGENCIAID " & _
           " WHERE PACOTESERVICO.PACOTEID = " & Formata_Dados(lngPACOTEID, tpDados_Longo) & _
           " AND PACOTESERVICO.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
           " ORDER BY SERVICO.DATAHORA DESC, AGENCIA.NOME, AGENCIACNPJ.CNPJ"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    UNIDASSOC_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim UNIDASSOC_Matriz(0 To UNIDASSOC_COLUNASMATRIZ - 1, 0 To UNIDASSOC_LINHASMATRIZ - 1)
  Else
    ReDim UNIDASSOC_Matriz(0 To UNIDASSOC_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se j� houver algum item
    For intI = 0 To UNIDASSOC_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda n�o se atingiu fim do recordset
        For intJ = 0 To UNIDASSOC_COLUNASMATRIZ - 1  'varre as colunas
          UNIDASSOC_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'pr�xima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Public Sub UNID_MontaMatriz()
  Dim strSql          As String
  Dim objRs           As ADODB.Recordset
  Dim intI            As Integer
  Dim intJ            As Integer
  Dim intCont         As Integer
  Dim objGeral        As busElite.clsGeral
  '
  On Error GoTo trata

  Set objGeral = New busElite.clsGeral

  strSql = "SELECT SERVICO.PKID, SERVICO.DATAHORA, " & _
           " AGENCIA.NOME + ' (' + dbo.formataCNPJ(AGENCIACNPJ.CNPJ) + ')', " & _
           " SERVICO.SOLICITANTE " & _
           "FROM SERVICO " & _
           " INNER JOIN AGENCIACNPJ ON AGENCIACNPJ.PKID = SERVICO.AGENCIACNPJID " & _
           " INNER JOIN AGENCIA ON AGENCIA.PKID = AGENCIACNPJ.AGENCIAID " & _
           " WHERE SERVICO.PKID NOT IN " & _
           "  (SELECT SERVICOID FROM PACOTESERVICO " & _
           "  WHERE PACOTESERVICO.STATUS = " & Formata_Dados("A", tpDados_Texto) & ")" & _
           " ORDER BY SERVICO.DATAHORA DESC, AGENCIA.NOME, AGENCIACNPJ.CNPJ"

  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then
    UNID_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim UNID_Matriz(0 To UNID_COLUNASMATRIZ - 1, 0 To UNID_LINHASMATRIZ - 1)
  Else
    ReDim UNID_Matriz(0 To UNID_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se j� houver algum item
    For intI = 0 To UNID_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda n�o se atingiu fim do recordset
        For intJ = 0 To UNID_COLUNASMATRIZ - 1  'varre as colunas
          UNID_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'pr�xima linha matriz
  End If
  Set objGeral = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdFechar_Click()
  '
  blnFechar = True
  Unload Me
End Sub


'''Private Function ValidaCampos() As Boolean
'''Dim Msg As String
'''Dim I As Integer
'''Dim sSql As String, rs As Recordset
''''
'''  '
'''  If Len(Msg) <> 0 Then
'''    MsgBox "Os seguintes erros ocorreram: " & vbCrLf & vbCrLf & Msg, vbExclamation, TITULOSISTEMA
'''    ValidaCampos = False
'''  Else
'''    ValidaCampos = True
'''  End If
'''End Function



Private Sub grdUnidadeAssoc_UnboundReadDataEx( _
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
               Offset + intI, UNIDASSOC_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, UNIDASSOC_COLUNASMATRIZ, UNIDASSOC_LINHASMATRIZ, UNIDASSOC_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, UNIDASSOC_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmServicoAssoc.grdUnidade_UnboundReadDataEx]"
End Sub



Private Sub grdUnidade_UnboundReadDataEx( _
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
               Offset + intI, UNID_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, UNID_COLUNASMATRIZ, UNID_LINHASMATRIZ, UNID_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, UNID_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmServicoAssoc.grdUnidade_UnboundReadDataEx]"
End Sub


Private Sub TratarAssociacao(pIndice As Integer)
  On Error GoTo trata
  Dim intI          As Long
  Dim objServico    As busElite.clsServico
  Dim objGeral      As busElite.clsGeral
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim lngRet        As Long
  Dim blnRet        As Boolean
  Dim intExc        As Long
  Dim objPacote     As busElite.clsPacote
  Dim strMsgErroExc As String
  Dim lngVEICULOID  As Long
  Dim strPlaca      As String
  '
  Set objServico = New busElite.clsServico
  '
  blnRet = False
  strMsgErroExc = ""
  intExc = 0
  '
  Select Case pIndice
  Case 1 'Cadastrar Selecionados
    'VEICULOID
    Set objGeral = New busElite.clsGeral
    strPlaca = ""
    strPlaca = Left(Right(cboVeiculo.Text, 9), 8)
    lngVEICULOID = 0
    strSql = "SELECT PKID FROM VEICULO WHERE VEICULO.PLACA = " & Formata_Dados(strPlaca, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngVEICULOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
    'Valida se obterve os campos com sucesso
    If lngVEICULOID = 0 Then
      Set objGeral = Nothing
      TratarErroPrevisto "Selecionar o ve�culo", "cmdOK_Click"
      Pintar_Controle cboVeiculo, tpCorContr_Erro
      SetarFoco cboVeiculo
      Exit Sub
    End If
    Set objGeral = Nothing
    For intI = 0 To grdUnidade.SelBookmarks.Count - 1
      grdUnidade.Bookmark = CLng(grdUnidade.SelBookmarks.Item(intI))
      'Verificar se item possui estoue suficiente
      objServico.AssociarServicoAoPacote lngPACOTEID, _
                                         grdUnidade.Columns("SERVICOID").Text, _
                                         lngVEICULOID
      blnRet = True
    Next
  Case 2 'Cadastrar Todos
    'VEICULOID
    strPlaca = ""
    strPlaca = Left(Right(cboVeiculo.Text, 9), 8)
    lngVEICULOID = 0
    strSql = "SELECT PKID FROM VEICULO WHERE VEICULO.PLACA = " & Formata_Dados(strPlaca, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngVEICULOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
    'Valida se obterve os campos com sucesso
    If lngVEICULOID = 0 Then
      Set objGeral = Nothing
      TratarErroPrevisto "Selecionar o ve�culo", "cmdOK_Click"
      Pintar_Controle cboVeiculo, tpCorContr_Erro
      SetarFoco cboVeiculo
      Exit Sub
    End If
    For intI = 0 To UNID_LINHASMATRIZ - 1
      grdUnidade.Bookmark = CLng(intI)
      objServico.AssociarServicoAoPacote lngPACOTEID, _
                                             grdUnidade.Columns("SERVICOID").Text, _
                                             lngVEICULOID
      blnRet = True
    Next
  Case 3 'Retirar Selecionados
    For intI = 0 To grdUnidadeAssoc.SelBookmarks.Count - 1
      grdUnidadeAssoc.Bookmark = CLng(grdUnidadeAssoc.SelBookmarks.Item(intI))
      If grdUnidadeAssoc.Columns("Status").Text & "" = "FINALIZADO" Then
        strMsgErroExc = strMsgErroExc & grdUnidadeAssoc.Columns("Data/hora").Text & _
                grdUnidadeAssoc.Columns("Ag�ncia").Text & _
                grdUnidadeAssoc.Columns("Solicitante").Text & vbCrLf
      Else
        objServico.DesassociarServicoAoPacote grdUnidadeAssoc.Columns("PACOTESERVICOID").Text
        blnRet = True
      End If
    Next
  Case 4 'retirar Todos
    For intI = 0 To UNIDASSOC_LINHASMATRIZ - 1
      grdUnidadeAssoc.Bookmark = CLng(intI)
      If IsNull(grdUnidadeAssoc.Bookmark) Then grdUnidadeAssoc.Bookmark = CLng(intI)
      If grdUnidadeAssoc.Columns("Status").Text & "" = "FINALIZADO" Then
        strMsgErroExc = strMsgErroExc & grdUnidadeAssoc.Columns("Data/hora").Text & _
                grdUnidadeAssoc.Columns("Ag�ncia").Text & _
                grdUnidadeAssoc.Columns("Solicitante").Text & vbCrLf
      Else
        objServico.DesassociarServicoAoPacote grdUnidadeAssoc.Columns("PACOTESERVICOID").Text
        blnRet = True
      End If
    Next
  End Select
  '
  If strMsgErroExc <> "" Then
    strMsgErroExc = "O(s) servi�o(s) abaxo n�o pode(m) ser excluido(s) pois j� forma finalizados: " & _
        vbCrLf & vbCrLf & strMsgErroExc & vbCrLf & vbCrLf
    If blnRet Then strMsgErroExc = strMsgErroExc & "O(s) demais foram dessassociados"
    TratarErroPrevisto strMsgErroExc, "[frmServicoAssoc.TratarAssociacao]"

  End If
  
  Set objServico = Nothing
    '
  If blnRet Then 'Houve Autera��o, Atualiza grids
    'Verifica se altera o status do pacote
    Set objPacote = New busElite.clsPacote
    objPacote.TratarStatus lngPACOTEID
    Set objPacote = Nothing
    
    blnRetorno = True
    '
    UNID_COLUNASMATRIZ = grdUnidade.Columns.Count
    UNID_LINHASMATRIZ = 0
    UNID_MontaMatriz
    grdUnidade.Bookmark = Null
    grdUnidade.ReBind
    '
    UNIDASSOC_COLUNASMATRIZ = grdUnidadeAssoc.Columns.Count
    UNIDASSOC_LINHASMATRIZ = 0
    UNIDASSOC_MontaMatriz
    grdUnidadeAssoc.Bookmark = Null
    grdUnidadeAssoc.ReBind
    '
  End If
  grdUnidade.ReBind
  grdUnidadeAssoc.ReBind
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub
''''''
''''''Private Function CompletaLanc(pValor As String, pSUBITEMPCOID As Long, pSUBOBRAID As String, pData As String, pValorDolar As Currency, pValorResiduo As Currency, ITEMPEDIDOMATERIALID As Long) As Boolean
''''''Dim rs As ADODB.Recordset
''''''Dim i As Integer
''''''Dim sSql, AuxSai As String
''''''Dim sMsgErro As String
''''''
''''''  Dim vrReal As Currency
''''''  Dim vrDolarAux As Currency
''''''  '
''''''  'Valor em Real � V�lido
''''''  vrReal = Formata_Dados_VB(Formata_Dados(pValor, tpDados_Moeda, tpNulo_NaoAceita), tpDados_Moeda)
''''''  vrDolarAux = vrReal / pValorDolar
''''''
''''''
''''''  SalvaLinha pValor, pSUBITEMPCOID, pSUBOBRAID, pData, vrDolarAux, pValorResiduo, ITEMPEDIDOMATERIALID
''''''
'''''''  MontaMatriz
'''''''  '
'''''''  ITEMPED_MontaMatriz
'''''''  grdUnidades.Bookmark = Null
'''''''  grdUnidades.ReBind
''''''
''''''  CompletaLanc = True
''''''
''''''End Function
