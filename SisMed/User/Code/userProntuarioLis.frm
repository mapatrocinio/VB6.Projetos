VERSION 5.00
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmUserProntuarioLis 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Lista de associados"
   ClientHeight    =   7410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7410
   ScaleWidth      =   10080
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7410
      Left            =   8220
      ScaleHeight     =   7410
      ScaleWidth      =   1860
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4665
         Left            =   90
         ScaleHeight     =   4605
         ScaleWidth      =   1635
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1695
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&X"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdInserir 
            Caption         =   "&Z"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   2730
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   3600
            Width           =   1335
         End
      End
      Begin Crystal.CrystalReport Report1 
         Left            =   0
         Top             =   0
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
   End
   Begin TrueDBGrid60.TDBGrid grdGeral 
      Align           =   3  'Align Left
      Height          =   7410
      Left            =   0
      OleObjectBlob   =   "userProntuarioLis.frx":0000
      TabIndex        =   0
      Top             =   0
      Width           =   8160
   End
End
Attribute VB_Name = "frmUserProntuarioLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public IcProntuario             As tpIcProntuario
Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String


Private Sub cmdAlterar_Click()
  On Error GoTo trata
  If Not IsNumeric(grdGeral.Columns("PKID").Value & "") Then
    MsgBox "Selecione um prestador !", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  
  frmUserProntuarioInc.Status = tpStatus_Alterar
  frmUserProntuarioInc.lngPKID = grdGeral.Columns("PKID").Value
  frmUserProntuarioInc.IcProntuario = IcProntuario
  frmUserProntuarioInc.intQuemChamou = 0
  frmUserProntuarioInc.Show vbModal
  
  If frmUserProntuarioInc.blnRetorno Then
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
  Dim objProntuario        As busSisMed.clsProntuario
  Dim objRs               As ADODB.Recordset
  Dim strSql              As String
  Dim objGeral            As busSisMed.clsGeral
  '
  On Error GoTo trata
  'Exclus�o de prontu�rio
  If Not IsNumeric(grdGeral.Columns("PKID").Value & "") Then
    MsgBox "Selecione um prontu�rio para exclus�o.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  '
  If MsgBox("ATEN��O: A exclus�o do prontu�rio remover� todas associa��es." & vbCrLf & "Caso queira voc� pode apenas alter�-lo e selecionar a op��o exclu�do, isso ir� exclu�-lo logicamente, mantendo suas informa��es na base de dados." & vbCrLf & "Confirma exclus�o do associado " & grdGeral.Columns("Nome").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdGeral
    Exit Sub
  End If
  'OK
  Set objProntuario = New busSisMed.clsProntuario
  objProntuario.ExcluirProntuario CLng(grdGeral.Columns("PKID").Value)
  Set objProntuario = Nothing
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



Private Sub cmdImprimir_Click()
  On Error GoTo TratErro
  AmpS
  '
  Report1.Destination = 0 'Video
  Report1.CopiesToPrinter = 1
  Report1.WindowState = crptMaximized
  '
  Report1.Action = 1
  '
  AmpN
  Exit Sub
  
TratErro:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Private Sub cmdInserir_Click()
  On Error GoTo trata
  frmUserProntuarioInc.Status = tpStatus_Incluir
  frmUserProntuarioInc.IcProntuario = IcProntuario
  frmUserProntuarioInc.lngPKID = 0
  frmUserProntuarioInc.intQuemChamou = 0
  frmUserProntuarioInc.Show vbModal

  If frmUserProntuarioInc.blnRetorno Then
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
  Me.Height = 7890
  Me.Width = 10170
  
  CenterForm Me
  
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar, cmdExcluir, , cmdInserir, cmdAlterar, , cmdImprimir
  
  If IcProntuario = tpIcProntuario_Func Then
    Me.Caption = "Lista de Funcion�rio"
  ElseIf IcProntuario = tpIcProntuario_Pac Then
    Me.Caption = "Lista de Paciente"
  ElseIf IcProntuario = tpIcProntuario_Prest Then
    Me.Caption = "Lista de Prestador"
  End If
  '
  Report1.Connect = ConnectRpt
  Report1.ReportFileName = gsReportPath & "ListaProntuario.rpt"
  '
  If gsNivel = gsArquivista Then
    cmdAlterar.Enabled = False
    cmdInserir.Enabled = False
    cmdExcluir.Enabled = False
    cmdImprimir.Enabled = True
  Else
    cmdAlterar.Enabled = True
    cmdInserir.Enabled = True
    cmdExcluir.Enabled = True
    cmdImprimir.Enabled = False
  End If
  '
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
  Dim intI      As Long
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMed.clsGeral
  '
  'Chamada do Cadastro de prontu�rio
  strSql = "SELECT PRONTUARIO.PKID, PRONTUARIO.NOME, PRONTUARIO.PKID, CASE TIPO_PESSOA WHEN 'F' THEN PRONTUARIO.CPF WHEN 'J' THEN PRONTUARIO.CNPJ ELSE '' END, PRONTUARIO.DTNASCIMENTO FROM PRONTUARIO "
  
  If IcProntuario = tpIcProntuario_Func Then
    strSql = strSql & " INNER JOIN FUNCIONARIO ON PRONTUARIO.PKID = FUNCIONARIO.PRONTUARIOID "
  ElseIf IcProntuario = tpIcProntuario_Pac Then
    strSql = strSql & " INNER JOIN PACIENTE ON PRONTUARIO.PKID = PACIENTE.PRONTUARIOID "
  ElseIf IcProntuario = tpIcProntuario_Prest Then
    strSql = strSql & " INNER JOIN PRESTADOR ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID "
  End If
  strSql = strSql & " ORDER BY PRONTUARIO.NOME"
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
  If Not objRs.EOF Then   'se j� houver algum item
    For intI = 0 To LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda n�o se atingiu fim do recordset
        For intJ = 0 To COLUNASMATRIZ - 1  'varre as colunas
          Matriz(intJ, intI) = objRs(intJ) & ""
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
  TratarErro Err.Number, Err.Description, "[frmUserProntuario.grdGeral_UnboundReadDataEx]"
End Sub




