VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmServicoTransInc 
   Caption         =   "Transferir Serviços do Pacote - "
   ClientHeight    =   6360
   ClientLeft      =   2595
   ClientTop       =   3120
   ClientWidth     =   9885
   Icon            =   "userServicoTransInc.frx":0000
   LinkTopic       =   "Form1"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6360
   ScaleWidth      =   9885
   Begin VB.ComboBox cboVeiculo 
      Height          =   315
      Left            =   1350
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   510
      Width           =   5955
   End
   Begin VB.TextBox txtServico 
      BackColor       =   &H00E0E0E0&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "txtServico"
      Top             =   180
      Width           =   8235
   End
   Begin VB.PictureBox picBotoes 
      Align           =   2  'Align Bottom
      BorderStyle     =   0  'None
      Height          =   900
      Left            =   0
      ScaleHeight     =   900
      ScaleWidth      =   9885
      TabIndex        =   3
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
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   0
         Width           =   9795
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   8370
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   0
            Width           =   1215
         End
         Begin VB.CommandButton cmdConfirmar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   7170
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   0
            Width           =   1215
         End
      End
   End
   Begin TrueDBGrid60.TDBGrid grdServico 
      Height          =   4440
      Left            =   0
      OleObjectBlob   =   "userServicoTransInc.frx":000C
      TabIndex        =   5
      Top             =   870
      Width           =   9885
   End
   Begin VB.Label Label6 
      Caption         =   "Veículo"
      Height          =   255
      Index           =   0
      Left            =   30
      TabIndex        =   8
      Top             =   540
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Serviço a ser transferido"
      Height          =   435
      Index           =   1
      Left            =   60
      TabIndex        =   6
      Top             =   105
      Width           =   1275
   End
End
Attribute VB_Name = "frmServicoTransInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public lngPACOTEID          As Long
Public lngPACOTESERVICOID   As Long
Public lngSERVICOID         As Long
Public strPacote            As String
Public strServico           As String
Public blnRetorno           As Boolean


Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String
Dim blnPrimeiraVez        As Boolean


Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdConfirmar_Click()
  Dim objGeral      As busElite.clsGeral
  Dim objRs         As ADODB.Recordset
  Dim objServico    As busElite.clsServico
  Dim objPacote     As busElite.clsPacote
  Dim lngVEICULOID  As Long
  Dim strPlaca      As String
  Dim strSql        As String

  On Error GoTo trata
    
  If Not IsNumeric(grdServico.Columns("PACOTEID").Value) Then
    MsgBox "Selecione um pacote!", vbExclamation, TITULOSISTEMA
    SetarFoco grdServico
    Exit Sub
  End If
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
    TratarErroPrevisto "Selecionar o veículo", "cmdOK_Click"
    Pintar_Controle cboVeiculo, tpCorContr_Erro
    SetarFoco cboVeiculo
    Exit Sub
  End If
  Set objGeral = Nothing
  '
  Set objServico = New busElite.clsServico
  Set objPacote = New busElite.clsPacote
  
  'ATUALIZAR STATUS DO SERVIÇO PARA INICIAL
  '------------------------------------------
  objServico.AlterarStatusServico lngSERVICOID, _
                                  "I"
                                  
  'ALTERAR O STATUS DO PACOTESERVICO PARA CANCELADO
  '------------------------------------------
  objServico.DesativarServicoDoPacote lngPACOTESERVICOID
  
  'TRATAR STATUS DO PACOTE CUJO SERVIÇO FOI TRANSFERIDO
  '------------------------------------------
  objPacote.TratarStatus lngPACOTEID


  'INCLUIR OU ALTERAR PACOTESERVICO PARA O NOVO PACOTE
  objServico.AssociarServicoAoPacote grdServico.Columns("PACOTEID").Value, _
                                     lngSERVICOID, _
                                     lngVEICULOID
  'TRATAR STATUS DO PACOTE CUJO SERVIÇO RECEBEU A TRANSFERENCIA
  '------------------------------------------
  objPacote.TratarStatus grdServico.Columns("PACOTEID").Value
  '
  Set objServico = Nothing
  Set objPacote = Nothing
  
  blnRetorno = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmServicoTransInc.cmdConfirmar_Click]"
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
  TratarErro Err.Number, Err.Description, "[frmServicoTransInc.grdServico_UnboundReadDataEx]"
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
  strSql = "SELECT PACOTE.PKID, PACOTE.DATAINICIO, PACOTE.DATATERMINO, PESSOA.NOME " & _
        "FROM PACOTE LEFT JOIN MOTORISTA ON MOTORISTA.PESSOAID = PACOTE.MOTORISTAID " & _
        " LEFT JOIN PESSOA ON PESSOA.PKID = MOTORISTA.PESSOAID " & _
        " WHERE (PACOTE.STATUS = " & Formata_Dados("I", tpDados_Texto) & _
        " OR PACOTE.STATUS = " & Formata_Dados("C", tpDados_Texto) & ")" & _
        " AND PACOTE.PKID <> " & Formata_Dados(lngPACOTEID, tpDados_Longo) & _
        " ORDER BY PACOTE.DATAINICIO DESC, PACOTE.DATATERMINO DESC;"
           

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
  blnRetorno = False
  'Combos
  'VAICULO
  strSql = "Select MODELO.NOME + ' (' + VEICULO.PLACA + ')' " & _
      " FROM VEICULO " & _
      " INNER JOIN MODELO ON MODELO.PKID = VEICULO.MODELOID " & _
      "ORDER BY MODELO.NOME, VEICULO.PLACA"
  PreencheCombo cboVeiculo, strSql, False, True
  '
  CenterForm Me
  Me.Caption = "Transferir Serviços do Pacote - " & strPacote
  txtServico.Text = strServico
  LerFiguras Me, tpBmp_Vazio, cmdConfirmar, cmdCancelar
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub


