VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmServicoCons 
   Caption         =   "Consulta de Serviços do Pacote - "
   ClientHeight    =   6360
   ClientLeft      =   2595
   ClientTop       =   3120
   ClientWidth     =   9885
   Icon            =   "userServicoCons.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6360
   ScaleWidth      =   9885
   Begin VB.PictureBox picBotoes 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   9885
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   5460
      Width           =   9885
      Begin VB.PictureBox picAlinDir 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   912
         Left            =   0
         ScaleHeight     =   915
         ScaleWidth      =   9795
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   0
         Width           =   9795
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   8370
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdConfirmar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   7170
            Style           =   1  'Graphical
            TabIndex        =   0
            Top             =   0
            Width           =   1215
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdServico 
      Align           =   1  'Align Top
      Height          =   5310
      Left            =   0
      OleObjectBlob   =   "userServicoCons.frx":000C
      TabIndex        =   4
      Top             =   0
      Width           =   9885
   End
End
Attribute VB_Name = "frmServicoCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public intQuemChamou    As Integer
'0 = FINALIZAR SERVIÇO
'1 = OBSERVAR SERVIÇO
'2 = TRANSFERIR SERVIÇO

Public lngPACOTEID      As Long
Public strPacote        As String



Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String
Dim blnPrimeiraVez        As Boolean


Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdConfirmar_Click()
  On Error GoTo trata
  Select Case intQuemChamou
  Case 0 'Chamada de Finalizar Serviço
    If Not IsNumeric(grdServico.Columns("SERVICOID").Value) Then
      MsgBox "Selecione um serviço!", vbExclamation, TITULOSISTEMA
      SetarFoco grdServico
      Exit Sub
    End If
    
    frmServicoFimInc.lngSERVICOID = grdServico.Columns("SERVICOID").Value
    frmServicoFimInc.lngPACOTEID = lngPACOTEID
    frmServicoFimInc.Status = tpStatus_Alterar
    frmServicoFimInc.Show vbModal
    
    If frmServicoFimInc.blnRetorno Then
      '
      COLUNASMATRIZ = grdServico.Columns.Count
      LINHASMATRIZ = 0
  
      MontaMatriz
      grdServico.Bookmark = Null
      grdServico.ReBind
      grdServico.ApproxCount = LINHASMATRIZ
    End If
    SetarFoco grdServico
  Case 1 'Chamada de Observar serviço
    If Not IsNumeric(grdServico.Columns("SERVICOID").Value) Then
      MsgBox "Selecione um serviço!", vbExclamation, TITULOSISTEMA
      SetarFoco grdServico
      Exit Sub
    End If
    
    frmHistoricoServicoLis.lngSERVICOID = grdServico.Columns("SERVICOID").Value
    frmHistoricoServicoLis.lngPACOTEID = lngPACOTEID
    frmHistoricoServicoLis.lngPACOTESERVICOID = grdServico.Columns("PACOTESERVICOID").Value
    frmHistoricoServicoLis.strCapction = grdServico.Columns(2).Value & " - " & grdServico.Columns(3).Value
    frmHistoricoServicoLis.Show vbModal
    
    If frmServicoFimInc.blnRetorno Then
      '
      COLUNASMATRIZ = grdServico.Columns.Count
      LINHASMATRIZ = 0
  
      MontaMatriz
      grdServico.Bookmark = Null
      grdServico.ReBind
      grdServico.ApproxCount = LINHASMATRIZ
    End If
    SetarFoco grdServico
    
  Case 2 'Chamada de Transferir Serviço
    If Not IsNumeric(grdServico.Columns("SERVICOID").Value) Then
      MsgBox "Selecione um serviço!", vbExclamation, TITULOSISTEMA
      SetarFoco grdServico
      Exit Sub
    ElseIf grdServico.Columns("STATUS").Value & "" = "FINALIZADO" Then
      MsgBox "Serviço já finalizado não pode ser transferido!", vbExclamation, TITULOSISTEMA
      SetarFoco grdServico
      Exit Sub
    End If
    
    frmServicoTransInc.lngSERVICOID = grdServico.Columns("SERVICOID").Value
    frmServicoTransInc.lngPACOTEID = lngPACOTEID
    frmServicoTransInc.lngPACOTESERVICOID = grdServico.Columns("PACOTESERVICOID").Value
    frmServicoTransInc.strServico = grdServico.Columns(2).Value & " - " & grdServico.Columns(3).Value
    frmServicoTransInc.strPacote = strPacote
    frmServicoTransInc.Show vbModal

    If frmServicoTransInc.blnRetorno Then
      '
      COLUNASMATRIZ = grdServico.Columns.Count
      LINHASMATRIZ = 0

      MontaMatriz
      grdServico.Bookmark = Null
      grdServico.ReBind
      grdServico.ApproxCount = LINHASMATRIZ
    End If
    SetarFoco grdServico
    
  End Select
  'cmdCancelar_Click
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmServicoCons.cmdConfirmar_Click]"
End Sub
Private Sub grdServico_UnboundReadDataEx( _
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
  TratarErro Err.Number, Err.Description, "[frmServicoCons.grdServico_UnboundReadDataEx]"
End Sub

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    DoEvents
    '
    COLUNASMATRIZ = grdServico.Columns.Count
    LINHASMATRIZ = 0

    MontaMatriz
    grdServico.Bookmark = Null
    grdServico.ReBind
    grdServico.ApproxCount = LINHASMATRIZ
    blnPrimeiraVez = False
    SetarFoco grdServico
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
Public Sub MontaMatriz(Optional strWhere As String)
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGeral    As busElite.clsGeral
  '
  AmpS
  On Error GoTo trata
  
  Set objGeral = New busElite.clsGeral
  '
  strSql = "SELECT SERVICO.PKID, PACOTESERVICO.PKID, SERVICO.DATAHORA, " & _
           " AGENCIA.NOME + ' (' + dbo.formataCNPJ(AGENCIACNPJ.CNPJ) + ')', " & _
           " SERVICO.SOLICITANTE, SERVICO.STATUS " & _
           "FROM PACOTESERVICO " & _
           " INNER JOIN SERVICO ON SERVICO.PKID = PACOTESERVICO.SERVICOID " & _
           " INNER JOIN AGENCIACNPJ ON AGENCIACNPJ.PKID = SERVICO.AGENCIACNPJID " & _
           " INNER JOIN AGENCIA ON AGENCIA.PKID = AGENCIACNPJ.AGENCIAID " & _
           " WHERE PACOTESERVICO.PACOTEID = " & Formata_Dados(lngPACOTEID, tpDados_Longo) & _
           " AND PACOTESERVICO.STATUS = " & Formata_Dados("A", tpDados_Texto)
           

  If Len(Trim(strWhere)) <> 0 Then
    strSql = strSql & strWhere
  End If
  strSql = strSql & " ORDER BY SERVICO.DATAHORA DESC, AGENCIA.NOME, AGENCIACNPJ.CNPJ;"
  '
  Set objRs = objGeral.ExecutarSQL(strSql)
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
  Set objGeral = Nothing
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
Private Sub Form_Load()
On Error GoTo trata
  Dim strSql As String
  AmpS
  Me.Height = 6765
  Me.Width = 10005
  blnPrimeiraVez = True
  
  CenterForm Me
  Select Case intQuemChamou
  Case 0:  Me.Caption = "Finalizar Serviços do Pacote - " & strPacote
  Case 1:  Me.Caption = "Observações dos Serviços do Pacote - " & strPacote
  Case 2:  Me.Caption = "Transferir Serviços do Pacote - " & strPacote
  End Select
  
  LerFiguras Me, tpBmp_Vazio, cmdConfirmar, cmdCancelar
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub


