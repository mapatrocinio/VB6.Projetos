VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserGrupoDespesaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Grupo de Despesa"
   ClientHeight    =   5055
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5055
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5055
      Left            =   8250
      ScaleHeight     =   5055
      ScaleWidth      =   1860
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4725
         Left            =   120
         ScaleHeight     =   4665
         ScaleWidth      =   1605
         TabIndex        =   6
         Top             =   180
         Width           =   1665
         Begin VB.CommandButton cmdInserir 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   3630
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4785
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   8440
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do Grupo"
      TabPicture(0)   =   "userGrupoDespesaInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Sub Grupos"
      TabPicture(1)   =   "userGrupoDespesaInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdGeral"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Informações cadastrais"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3135
         Left            =   120
         TabIndex        =   8
         Top             =   360
         Width           =   7335
         Begin VB.Frame Frame5 
            Caption         =   "Sexo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   840
            TabIndex        =   9
            Top             =   3480
            Width           =   2295
         End
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1080
            Width           =   2535
         End
         Begin VB.TextBox txtDescricao 
            Height          =   285
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   1
            Top             =   720
            Width           =   5175
         End
         Begin MSMask.MaskEdBox mskGrupo 
            Height          =   255
            Left            =   1560
            TabIndex        =   0
            Top             =   360
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   450
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   4
            Mask            =   "####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblLivro 
            Caption         =   "Tipo"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label lblCheque 
            Caption         =   "Descrição"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Código"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   360
            Width           =   1215
         End
      End
      Begin TrueDBGrid60.TDBGrid grdGeral 
         Height          =   3300
         Left            =   -74880
         OleObjectBlob   =   "userGrupoDespesaInc.frx":0038
         TabIndex        =   13
         Top             =   480
         Width           =   7545
      End
   End
End
Attribute VB_Name = "frmUserGrupoDespesaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                         As tpStatus
Public lngGRUPODESPESAID              As Long
Public bRetorno                       As Boolean
Public blnPrimeiraVez                 As Boolean
Public bFechar                        As Boolean
Dim COLUNASMATRIZ         As Long
Dim LINHASMATRIZ          As Long
Private Matriz()          As String

Private Sub cboTipo_LostFocus()
  Pintar_Controle cboTipo, tpCorContr_Normal
End Sub

Private Sub cmdAlterar_Click()
  
  If Not IsNumeric(grdGeral.Columns("ID").Value) Then
    MsgBox "Selecione um sub grupo de despesa !", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  frmUserSubGrupoDespesaInc.Status = tpStatus_Alterar
  frmUserSubGrupoDespesaInc.lngGRUPODESPESAID = lngGRUPODESPESAID
  frmUserSubGrupoDespesaInc.lngSUBGRUPODESPESAID = CLng(grdGeral.Columns("ID").Value)
  frmUserSubGrupoDespesaInc.Show vbModal
  
  If frmUserSubGrupoDespesaInc.bRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If

End Sub

Private Sub cmdCancelar_Click()
  bFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdExcluir_Click()
  Dim objSubGrupoDespesa    As busSisMaq.clsSubGrupoDespesa
  Dim objGeral              As busSisMaq.clsGeral
  Dim objRs                 As ADODB.Recordset
  Dim strSql                As String
  '
  On Error GoTo trata
  If Len(Trim(grdGeral.Columns("ID").Value)) = 0 Then
    MsgBox "Selecione um sub grupo de despesa.", vbExclamation, TITULOSISTEMA
    Exit Sub
  End If
  If MsgBox("Confirma exclusão do sub grupo de despesa " & grdGeral.Columns("Descrição").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then Exit Sub
  'OK
  Set objGeral = New busSisMaq.clsGeral
  strSql = "SELECT * FROM DESPESA WHERE SUBGRUPODESPESAID = " & CLng(grdGeral.Columns("ID").Value)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    TratarErroPrevisto "Não é possível excluir o sub grupo de despesa, por constar despesas lançadas para ele.", "[cmdExcluir_Click]"
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  Set objSubGrupoDespesa = New busSisMaq.clsSubGrupoDespesa
  
  objSubGrupoDespesa.ExcluirSubGrupoDespesa CLng(grdGeral.Columns("ID").Value)
  '
  MontaMatriz
  grdGeral.Bookmark = Null
  grdGeral.ReBind
  
  Set objSubGrupoDespesa = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdInserir_Click()
  frmUserSubGrupoDespesaInc.Status = tpStatus_Incluir
  frmUserSubGrupoDespesaInc.lngGRUPODESPESAID = lngGRUPODESPESAID
  frmUserSubGrupoDespesaInc.Show vbModal
  
  If frmUserSubGrupoDespesaInc.bRetorno Then
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
  End If

End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim objGrupoDespesa         As busSisMaq.clsGrupoDespesa
  Dim clsGer                  As busSisMaq.clsGeral
  Dim strTipo                 As String
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração
    If Not ValidaCampos Then Exit Sub
    'Valida se grupo da despesa já cadastrada
    Set clsGer = New busSisMaq.clsGeral
    strSql = "Select PKID From GRUPODESPESA WHERE CODIGO = " & Formata_Dados(mskGrupo.Text, tpDados_Texto, tpNulo_NaoAceita) & _
      " AND PKID <> " & Formata_Dados(lngGRUPODESPESAID, tpDados_Longo)
    Set objRs = clsGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set clsGer = Nothing
      TratarErroPrevisto "Código do grupo de despesa já cadastrado", "cmdOK_Click"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    
    strSql = "Select PKID From GRUPODESPESA WHERE DESCRICAO = " & Formata_Dados(txtDescricao.Text, tpDados_Texto, tpNulo_NaoAceita) & _
      " AND PKID <> " & Formata_Dados(lngGRUPODESPESAID, tpDados_Longo)
    Set objRs = clsGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set clsGer = Nothing
      TratarErroPrevisto "Descrição do grupo de despesa já cadastrada", "cmdOK_Click"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing

    Set objGrupoDespesa = New busSisMaq.clsGrupoDespesa
    strTipo = Left(cboTipo.Text, 1)
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objGrupoDespesa.AlterarGrupoDespesa lngGRUPODESPESAID, _
                                          mskGrupo.Text, _
                                          txtDescricao.Text, _
                                          strTipo
      bRetorno = True
      tabDetalhes.Tab = 1
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      objGrupoDespesa.IncluirGrupoDespesa mskGrupo.Text, _
                                          txtDescricao.Text, _
                                          strTipo
      'CAPTURAR PKID
      strSql = "Select PKID From GRUPODESPESA WHERE CODIGO = " & Formata_Dados(mskGrupo.Text, tpDados_Texto)
      Set objRs = clsGer.ExecutarSQL(strSql)
      If Not objRs.EOF Then
        lngGRUPODESPESAID = objRs.Fields("PKID").Value
      End If
      objRs.Close
      Set objRs = Nothing
      
      cmdExcluir.Enabled = True
      cmdInserir.Enabled = True
      cmdAlterar.Enabled = True
      tabDetalhes.TabEnabled(1) = True
      tabDetalhes.Tab = 1
      Status = tpStatus_Alterar
      bRetorno = True
    End If
    Set objGrupoDespesa = Nothing
    'bFechar = True
    'Unload Me
  End Select
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg     As String
  '
  If Not IsNumeric(mskGrupo.Text) Or Len(mskGrupo.ClipText) <> 4 Then
    strMsg = strMsg & "Informar o código do grupo da despesa válido" & vbCrLf
    Pintar_Controle mskGrupo, tpCorContr_Erro
  End If
  If Len(txtDescricao.Text) = 0 Then
    strMsg = strMsg & "Informar a descrição válida" & vbCrLf
    Pintar_Controle txtDescricao, tpCorContr_Erro
  End If
  If Len(cboTipo.Text) = 0 Then
    strMsg = strMsg & "Selecionar o tipo do grupo de despesa" & vbCrLf
    Pintar_Controle cboTipo, tpCorContr_Erro
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserGrupoDespesaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco mskGrupo
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserGrupoDespesaInc.Form_Activate]"
End Sub


Public Sub MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMaq.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMaq.clsGeral
  '
  strSql = "SELECT SUBGRUPODESPESA.PKID, SUBGRUPODESPESA.CODIGO, SUBGRUPODESPESA.DESCRICAO FROM SUBGRUPODESPESA " & _
      " WHERE GRUPODESPESAID = " & lngGRUPODESPESAID & _
      " ORDER BY SUBGRUPODESPESA.CODIGO;"
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


Private Sub Form_Load()
On Error GoTo trata
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim objGrupoDespesa As busSisMaq.clsGrupoDespesa
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 5535
  Me.Width = 10200
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , cmdInserir, cmdAlterar
  '
  cboTipo.AddItem ""
  cboTipo.AddItem "DÉBITO"
  cboTipo.AddItem "CRÉDITO"
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    LimparCampoMask mskGrupo
    LimparCampoTexto txtDescricao
    '
    cmdExcluir.Enabled = False
    cmdInserir.Enabled = False
    cmdAlterar.Enabled = False
    tabDetalhes.TabEnabled(1) = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objGrupoDespesa = New busSisMaq.clsGrupoDespesa
    Set objRs = objGrupoDespesa.SelecionarGrupoDespesa(lngGRUPODESPESAID)
    '
    If Not objRs.EOF Then
      INCLUIR_VALOR_NO_MASK mskGrupo, objRs.Fields("CODIGO").Value & "", TpMaskOutros
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      If objRs.Fields("TIPO").Value & "" = "D" Then
        cboTipo.Text = "DÉBITO"
      ElseIf objRs.Fields("TIPO").Value & "" = "C" Then
        cboTipo.Text = "CRÉDITO"
      End If
    End If
    Set objGrupoDespesa = Nothing
    cmdInserir.Enabled = False
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdOk.Enabled = True
  End If
  '
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If Not bFechar Then Cancel = True
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
  TratarErro Err.Number, Err.Description, "[frmUserGrupoDespesaInc.grdGeral_UnboundReadDataEx]"
End Sub



Private Sub mskGrupo_GotFocus()
  Selecionar_Conteudo mskGrupo
End Sub

Private Sub mskGrupo_LostFocus()
  Pintar_Controle mskGrupo, tpCorContr_Normal
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    cmdInserir.Enabled = False
    cmdExcluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdOk.Enabled = True
  Case 1
    COLUNASMATRIZ = grdGeral.Columns.Count
    LINHASMATRIZ = 0
  
    MontaMatriz
    grdGeral.ApproxCount = LINHASMATRIZ
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    '
    cmdInserir.Enabled = True
    cmdExcluir.Enabled = True
    cmdAlterar.Enabled = True
    cmdOk.Enabled = False
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub


Private Sub txtDescricao_GotFocus()
  Selecionar_Conteudo txtDescricao
End Sub

Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub

