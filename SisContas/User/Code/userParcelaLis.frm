VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmUserParcelaLis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de parcelas"
   ClientHeight    =   4980
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7170
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4980
   ScaleWidth      =   7170
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4980
      Left            =   5310
      ScaleHeight     =   4980
      ScaleWidth      =   1860
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2175
         Left            =   60
         ScaleHeight     =   2115
         ScaleWidth      =   1635
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   2670
         Width           =   1695
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   180
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   1050
            Width           =   1335
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Align           =   3  'Align Left
      Height          =   4980
      Left            =   0
      OleObjectBlob   =   "userParcelaLis.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   5190
   End
End
Attribute VB_Name = "frmUserParcelaLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String
'
Public strNomeSuiteApto   As String
Public lngCCId            As Long



Private Sub cmdAlterar_Click()
  On Error GoTo trata
  If Not IsNumeric(grdGeral.Columns("PKID").Value & "") Then
    MsgBox "Selecione uma parcela !", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  
  frmUserParcelaInc.Status = tpStatus_Alterar
  frmUserParcelaInc.lngParcelaId = grdGeral.Columns("PKID").Value
  frmUserParcelaInc.strNumeroAptoPrinc = strNomeSuiteApto
  frmUserParcelaInc.Show vbModal
  
  If frmUserParcelaInc.blnRetorno Then
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
  Dim objTabFichaClieLoc  As busSisContas.clsTabFichaClieLoc
  Dim objFichaCliente     As busSisContas.clsFichaCliente
  Dim objRs               As ADODB.Recordset
  Dim strSql              As String
  Dim objGeral            As busSisContas.clsGeral
  '
  On Error GoTo trata
'''  If intChamada = 0 Then
'''    'Locação
'''    If Len(Trim(grdGeral.Columns("TabFichaClieLocId").Value & "")) = 0 Then
'''      MsgBox "Selecione um cliente para exclusão.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    ElseIf Trim(grdGeral.Columns("Tipo").Value & "") = "P" Then
'''      MsgBox "Cliente principal da locação não pode ser excluído.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''
'''  Else
'''    'Cadastro de Clientes
'''    If Len(Trim(grdGeral.Columns("FichaClienteId").Value & "")) = 0 Then
'''      MsgBox "Selecione um cliente para exclusão.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    Set objGeral = New busSisContas.clsGeral
'''    'Locações
'''    strSql = "SELECT * FROM TAB_FICHACLIELOC " & _
'''      " WHERE FICHACLIENTEID = " & Formata_Dados(grdGeral.Columns("FichaClienteId").Value, tpDados_Longo)
'''    Set objRs = objGeral.ExecutarSQL(strSql)
'''    If Not objRs.EOF Then
'''      objRs.Close
'''      Set objRs = Nothing
'''      Set objGeral = Nothing
'''      TratarErroPrevisto "O cliente não pode ser excluído, pois há locações associadas a ele."
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    'Reservas
'''    strSql = "SELECT * FROM RESERVA " & _
'''      " WHERE FICHACLIENTEID = " & Formata_Dados(grdGeral.Columns("FichaClienteId").Value, tpDados_Longo)
'''    Set objRs = objGeral.ExecutarSQL(strSql)
'''    If Not objRs.EOF Then
'''      objRs.Close
'''      Set objRs = Nothing
'''      Set objGeral = Nothing
'''      TratarErroPrevisto "O cliente não pode ser excluído, pois há reservas associadas a ele."
'''      SetarFoco grdGeral
'''      Exit Sub
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    '
'''    Set objGeral = Nothing
'''  End If
  '
  If MsgBox("Confirma exclusão do cliente " & grdGeral.Columns("Nome").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdGeral
    Exit Sub
  End If
  'OK
'''  If intChamada = 0 Then 'Chamada de Locação
'''    Set objTabFichaClieLoc = New busSisContas.clsTabFichaClieLoc
'''    objTabFichaClieLoc.ExcluirTabFichaClieLoc CLng(grdGeral.Columns("TabFichaClieLocId").Value)
'''    Set objTabFichaClieLoc = Nothing
'''  Else
'''    Set objFichaCliente = New busSisContas.clsFichaCliente
'''    objFichaCliente.ExcluirFichaCliente CLng(grdGeral.Columns("FichaClienteId").Value)
'''    Set objFichaCliente = Nothing
'''  End If
  
  '
  MontaMatriz
  grdGeral.Bookmark = Null
  grdGeral.ReBind

  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdFechar_Click()
   Unload Me
End Sub


Private Sub cmdFiltrar_Click()
  On Error GoTo trata

'''  frmUserFichaCliente.lngCLIENTEID = 0
'''  frmUserFichaCliente.Status = tpStatus_Consultar
'''  frmUserFichaCliente.Show vbModal
'''  If frmUserFichaCliente.bRetorno Then
'''    MontaMatriz
'''    grdGeral.Bookmark = Null
'''    grdGeral.ReBind
'''  End If
'''  SetarFoco grdGeral
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub cmdInserir_Click()
  On Error GoTo trata
'''  frmUserFichaClienteInc.Status = tpStatus_Incluir
'''  frmUserFichaClienteInc.intChamada = intChamada
'''  frmUserFichaClienteInc.lngTabFichaClienteId = 0
'''  frmUserFichaClienteInc.lngCCId = lngCCId
'''  frmUserFichaClienteInc.strNumeroAptoPrinc = strNomeSuiteApto
'''  frmUserFichaClienteInc.Show vbModal

  If frmUserFichaClienteInc.blnRetorno Then
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
  Me.Width = 7260
  
  CenterForm Me
  
'''  If intChamada = 0 Then
'''    cmdFiltrar.Enabled = False
'''  Else
'''    cmdFiltrar.Enabled = True
'''    grdGeral.Columns(2).Visible = False
'''  End If
  
  Me.Caption = Me.Caption & " - " & strNomeSuiteApto
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar, , , , cmdAlterar
  
  
  COLUNASMATRIZ = grdGeral.Columns.Count
  LINHASMATRIZ = 0

  MontaMatriz
  grdGeral.ApproxCount = LINHASMATRIZ
  
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
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
  strSql = "SELECT PARCELA.PKID, PARCELA.PARCELA, PARCELA.DTVENCIMENTO, PARCELA.VRPARCELA " & _
            "FROM PARCELA " & _
            " WHERE PARCELA.CONTACORRENTEID = " & Formata_Dados(lngCCId, tpDados_Longo) & _
            " ORDER BY PARCELA.PARCELA;"
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

Private Sub Form_Unload(Cancel As Integer)
  Dim strSql            As String
  Dim objRs             As ADODB.Recordset
  Dim objGeral          As busSisContas.clsGeral
  Dim curVrCC           As Currency
  Dim curVrTotParc      As Currency
  Dim strMsgErro        As String
  '
  
  On Error GoTo trata
  Set objGeral = New busSisContas.clsGeral
  curVrCC = 0
  curVrTotParc = 0
  strSql = "SELECT min(CONTACORRENTE.VALOR) AS VRCC, SUM(PARCELA.VRPARCELA) AS VRTOTPARCELA " & _
    "FROM CONTACORRENTE INNER JOIN PARCELA ON CONTACORRENTE.PKID = PARCELA.CONTACORRENTEID " & _
    "WHERE CONTACORRENTE.PKID = " & Formata_Dados(lngCCId, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    If IsNumeric(objRs.Fields("VRCC").Value) Then
      curVrCC = objRs.Fields("VRCC").Value
    End If
    If IsNumeric(objRs.Fields("VRTOTPARCELA").Value) Then
      curVrTotParc = objRs.Fields("VRTOTPARCELA").Value
    End If
  End If
  If curVrCC > curVrTotParc Then
    strMsgErro = "ATENÇÃO:" & vbCrLf & vbCrLf & "O valor total do pagamento [R$ " & _
      Format(curVrCC, "###,##0.00") & "] não pode ser maior que o valor das parcelas [R$ " & _
      Format(curVrTotParc, "###,##0.00") & "]" & vbCrLf & vbCrLf & _
      "Ajuste os valores das parcelas para que se possa fechar a soma."
    TratarErroPrevisto strMsgErro
    Cancel = 1
  End If
  
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
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
  TratarErro Err.Number, Err.Description, "[frmUserFichaCliente.grdGeral_UnboundReadDataEx]"
End Sub




