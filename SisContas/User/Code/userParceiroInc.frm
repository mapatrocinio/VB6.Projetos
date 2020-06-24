VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmUserParceiroInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de Parceiros"
   ClientHeight    =   4665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4665
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4665
      Left            =   8520
      ScaleHeight     =   4665
      ScaleWidth      =   1860
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2955
         Left            =   0
         ScaleHeight     =   2895
         ScaleWidth      =   1605
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   1560
         Width           =   1665
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   990
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4395
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   7752
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Parceiro"
      TabPicture(0)   =   "userParceiroInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Usuários"
      TabPicture(1)   =   "userParceiroInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdUsuario"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   1755
         Left            =   150
         TabIndex        =   7
         Top             =   360
         Width           =   7545
         Begin VB.TextBox txtDescricao 
            Height          =   285
            Left            =   1680
            MaxLength       =   100
            TabIndex        =   0
            Top             =   360
            Width           =   5535
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Parceiro"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   8
            Top             =   360
            Width           =   1095
         End
      End
      Begin TrueDBGrid60.TDBGrid grdUsuario 
         Height          =   3795
         Left            =   -74880
         OleObjectBlob   =   "userParceiroInc.frx":0038
         TabIndex        =   9
         Top             =   480
         Width           =   7335
      End
   End
End
Attribute VB_Name = "frmUserParceiroInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                 As tpStatus
Public lngPARCEIROID          As Long
Public bRetorno               As Boolean
Public bFechar                As Boolean

'Variáveis para Grid

Dim USU_COLUNASMATRIZ         As Long
Dim USU_LINHASMATRIZ          As Long
Private USU_Matriz()          As String


Public Sub USU_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGeral  As busSisContas.clsGeral
  '
  On Error GoTo trata
  
  Set objGeral = New busSisContas.clsGeral
  '
  strSql = "SELECT CONTROLEACESSO.USUARIO, TAB_CONTROL_PARC.PKID, TAB_CONTROL_PARC.PARCEIROID, CONTROLEACESSO.PKID AS CONTROLEACESSOID " & _
            "FROM TAB_CONTROL_PARC INNER JOIN  CONTROLEACESSO ON CONTROLEACESSO.PKID =  TAB_CONTROL_PARC.CONTROLEACESSOID " & _
            "WHERE TAB_CONTROL_PARC.PARCEIROID = " & Formata_Dados(lngPARCEIROID, tpDados_Longo) & " " & _
            "ORDER BY CONTROLEACESSO.USUARIO"
  '
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    USU_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim USU_Matriz(0 To USU_COLUNASMATRIZ - 1, 0 To USU_LINHASMATRIZ - 1)
  Else
    ReDim USU_Matriz(0 To USU_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To USU_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To USU_COLUNASMATRIZ - 1  'varre as colunas
          USU_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set objGeral = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub


Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Dim strTipo As String
  Dim objRs As ADODB.Recordset
  '
  Select Case tabDetalhes.Tab
  Case 1
    frmuserControlParcInc.strParceiro = txtDescricao.Text
    frmuserControlParcInc.lngPARCEIROID = lngPARCEIROID
    frmuserControlParcInc.Show vbModal
    If frmuserControlParcInc.bRetorno Then
      USU_COLUNASMATRIZ = grdUsuario.Columns.Count
      USU_LINHASMATRIZ = 0
      USU_MontaMatriz
      grdUsuario.Bookmark = Null
      grdUsuario.ReBind
      grdUsuario.ApproxCount = USU_LINHASMATRIZ
    End If
    SetarFoco grdUsuario
  End Select
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub grdUsuario_UnboundReadDataEx( _
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
               Offset + intI, USU_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, USU_COLUNASMATRIZ, USU_LINHASMATRIZ, USU_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, USU_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserParceiroInc.grdUsuario_UnboundReadDataEx]"
End Sub


Private Sub cmdCancelar_Click()
  '
  bFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error GoTo trata
  Dim strSql          As String
  Dim strMsgErro      As String
  Dim objRs           As ADODB.Recordset
  Dim objParceiro     As busSisContas.clsParceiro
  '
  'Código para Inclusão
  'Validações Básicas
  If Not ValidaCampos Then Exit Sub
  Set objParceiro = New busSisContas.clsParceiro
  '
  If Status = tpStatus_Incluir Then
    'Inserir PARCEIRO
    objParceiro.InserirParceiro txtDescricao.Text
    '
    Set objRs = objParceiro.ListarParceiroPelaDesc(txtDescricao.Text)
    If Not objRs.EOF Then
      lngPARCEIROID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    'Associar todas as unidades ao novo estoque cadastrado
    Set objParceiro = Nothing
    '
    bFechar = True
    bRetorno = True
    Unload Me
  Else
    'Alterar PARCEIRO
    objParceiro.AlterarParceiro txtDescricao.Text, _
                                lngPARCEIROID
    '
    Set objParceiro = Nothing
    '
    bFechar = True
    bRetorno = True
    Unload Me
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg     As String
  
  If Len(Trim(txtDescricao.Text)) = 0 Then
    strMsg = strMsg & "A descrição do parceiro é inválida." & vbCrLf
    Pintar_Controle txtDescricao, tpCorContr_Erro
  End If
  '
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserParceiroInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim objParceiro   As busSisContas.clsParceiro
  Dim strTipo       As String
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 5145
  Me.Width = 10470
  CenterForm Me
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, , , , cmdAlterar
  '
  '
  tabDetalhes_Click 0
  '
  If Status = tpStatus_Alterar Then
    '
    'Pega Dados do Banco de dados
    Set objParceiro = New busSisContas.clsParceiro
    Set objRs = objParceiro.ListarParceiro(lngPARCEIROID)
    '
    If Not objRs.EOF Then
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      '
    End If
    objRs.Close
    Set objRs = Nothing
    Set objParceiro = Nothing
  Else
    tabDetalhes.TabEnabled(1) = False
  End If

  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Unload Me
End Sub



Private Sub Form_Unload(Cancel As Integer)
  If Not bFechar Then Cancel = True
End Sub



Private Sub tabDetalhes_Click(PreviousTab As Integer)
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    'dados principais do Pedido
    Frame1.Enabled = True
    grdUsuario.Enabled = False
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdAlterar.Enabled = False
    SetarFoco txtDescricao
  Case 1
    Frame1.Enabled = False
    grdUsuario.Enabled = True
    '
    'Inclusão de Usuários
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdAlterar.Enabled = True
    'Montar RecordSet
    USU_COLUNASMATRIZ = grdUsuario.Columns.Count
    USU_LINHASMATRIZ = 0
    USU_MontaMatriz
    grdUsuario.Bookmark = Null
    grdUsuario.ReBind
    grdUsuario.ApproxCount = USU_LINHASMATRIZ
    SetarFoco grdUsuario
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "SisMotel.frmUserParceiroInc.tabDetalhes"
  AmpN
End Sub


Private Sub txtDescricao_GotFocus()
  Selecionar_Conteudo txtDescricao
End Sub

Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub





