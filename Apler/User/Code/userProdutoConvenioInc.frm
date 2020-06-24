VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserProdutoConvenioInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Produto de Convênio"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4890
      Left            =   8430
      ScaleHeight     =   4890
      ScaleWidth      =   1860
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4635
         Left            =   90
         ScaleHeight     =   4575
         ScaleWidth      =   1605
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   150
         Width           =   1665
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   3570
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4665
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   8229
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do produto"
      TabPicture(0)   =   "userProdutoConvenioInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Sub-produto"
      TabPicture(1)   =   "userProdutoConvenioInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdSubProduto"
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2655
         Left            =   120
         TabIndex        =   11
         Top             =   420
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2385
            Index           =   0
            Left            =   120
            ScaleHeight     =   2385
            ScaleWidth      =   7575
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtConvenio 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtConvenio"
               Top             =   90
               Width           =   6075
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1290
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   1020
               Width           =   2235
               Begin VB.OptionButton optStatus 
                  Caption         =   "Inativo"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   5
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optStatus 
                  Caption         =   "Ativo"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   4
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   825
               End
            End
            Begin VB.TextBox txtProdutoConvenio 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   1
               Text            =   "txtProdutoConvenio"
               Top             =   390
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskPercComissao 
               Height          =   255
               Left            =   1320
               TabIndex        =   2
               Top             =   720
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskPercVitalicio 
               Height          =   255
               Left            =   5820
               TabIndex        =   3
               Top             =   720
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "% Vitalicio"
               Height          =   195
               Index           =   2
               Left            =   4560
               TabIndex        =   18
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Convênio"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   17
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "% Comissão"
               Height          =   195
               Index           =   21
               Left            =   60
               TabIndex        =   16
               Top             =   735
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Status"
               Height          =   315
               Index           =   5
               Left            =   60
               TabIndex        =   14
               Top             =   1050
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Produto Convênio"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   13
               Top             =   450
               Width           =   1215
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdSubProduto 
         Height          =   4155
         Left            =   -74940
         OleObjectBlob   =   "userProdutoConvenioInc.frx":0038
         TabIndex        =   19
         Top             =   390
         Width           =   7965
      End
   End
End
Attribute VB_Name = "frmUserProdutoConvenioInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngCONVENIOID            As Long
Public strDescrConvenio         As String

Private blnPrimeiraVez          As Boolean

Dim SPROD_COLUNASMATRIZ         As Long
Dim SPROD_LINHASMATRIZ          As Long
Private SPROD_Matriz()          As String

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Valor Produto Convênio
  LimparCampoTexto txtConvenio
  LimparCampoTexto txtProdutoConvenio
  LimparCampoMask mskPercComissao
  LimparCampoMask mskPercVitalicio
  optStatus(0).Value = False
  optStatus(1).Value = False
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserProdutoConvenioInc.LimparCampos]", _
            Err.Description
End Sub

Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 1
    'Sub-Produto
    If Not IsNumeric(grdSubProduto.Columns("PKID").Value & "") Then
      MsgBox "Selecione um sub-produto !", vbExclamation, TITULOSISTEMA
      SetarFoco grdSubProduto
      Exit Sub
    End If

    frmUserSubProdProdutoInc.lngPKID = grdSubProduto.Columns("PKID").Value
    frmUserSubProdProdutoInc.lngPRODUTOID = lngPKID
    frmUserSubProdProdutoInc.strDescrConvenio = txtConvenio.Text
    frmUserSubProdProdutoInc.strDescrProduto = txtProdutoConvenio.Text
    frmUserSubProdProdutoInc.Status = tpStatus_Alterar
    frmUserSubProdProdutoInc.Show vbModal

    If frmUserSubProdProdutoInc.blnRetorno Then
      SPROD_MontaMatriz
      grdSubProduto.Bookmark = Null
      grdSubProduto.ReBind
      grdSubProduto.ApproxCount = SPROD_LINHASMATRIZ
    End If
    SetarFoco grdSubProduto
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdCancelar_Click()
  blnFechar = True
  blnRetorno = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub





Private Sub cmdExcluir_Click()
  On Error GoTo trata
  Dim objSubProdProduto       As busApler.clsSubProdProduto
  'Dim objProdutoConvenio     As busApler.clsProdutoConvenio
  '
  Select Case tabDetalhes.Tab
  Case 1 'Sub-Produto
    '
    If Len(Trim(grdSubProduto.Columns("PKID").Value & "")) = 0 Then
      MsgBox "Selecione um Sub-Produto do produto/convênio.", vbExclamation, TITULOSISTEMA
      SetarFoco grdSubProduto
      Exit Sub
    End If
    '
    Set objSubProdProduto = New busApler.clsSubProdProduto
    '
    If MsgBox("Confirma exclusão do sub-produto " & grdSubProduto.Columns("Sub-Produto").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdSubProduto
      Exit Sub
    End If
    'OK
    objSubProdProduto.ExcluirSubProdProduto CLng(grdSubProduto.Columns("PKID").Value)
    '
    SPROD_MontaMatriz
    grdSubProduto.Bookmark = Null
    grdSubProduto.ReBind
    grdSubProduto.ApproxCount = SPROD_LINHASMATRIZ

    Set objSubProdProduto = Nothing
    SetarFoco grdSubProduto
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdIncluir_Click()
  On Error GoTo trata
  Dim objForm As Form
  '
  Select Case tabDetalhes.Tab
  Case 1 'Sub-Produto
    frmUserSubProdProdutoInc.Status = tpStatus_Incluir
    frmUserSubProdProdutoInc.lngPKID = 0
    frmUserSubProdProdutoInc.lngPRODUTOID = lngPKID
    frmUserSubProdProdutoInc.strDescrConvenio = txtConvenio.Text
    frmUserSubProdProdutoInc.strDescrProduto = txtProdutoConvenio.Text
    frmUserSubProdProdutoInc.Show vbModal

    If frmUserSubProdProdutoInc.blnRetorno Then
      SPROD_MontaMatriz
      grdSubProduto.Bookmark = Null
      grdSubProduto.ReBind
      grdSubProduto.ApproxCount = SPROD_LINHASMATRIZ
    End If
    SetarFoco grdSubProduto
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdOK_Click()
  Dim objProdutoConvenio        As busApler.clsProdutoConvenio
  Dim objGeral                  As busApler.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim strStatus                 As String
  Dim strDatDesativacao         As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busApler.clsGeral
  Set objProdutoConvenio = New busApler.clsProdutoConvenio
  'Status
  If optStatus(0).Value Then
    strStatus = "A"
  Else
    strStatus = "I"
  End If

  'Validar se produto convênio já cadastrado
  strSql = "SELECT * FROM PRODUTO " & _
    " WHERE PRODUTO.DESCRICAO = " & Formata_Dados(txtProdutoConvenio.Text, tpDados_Texto) & _
    " AND PRODUTO.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtProdutoConvenio, tpCorContr_Erro
    TratarErroPrevisto "Produto do convênio já cadastrado"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objProdutoConvenio = Nothing
    cmdOk.Enabled = True
    SetarFoco txtProdutoConvenio
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar ProdutoConvenio
    strDatDesativacao = ""
    If strStatus = "I" Then
      strDatDesativacao = Format(Now, "DD/MM/YYYY hh:mm")
    End If
    objProdutoConvenio.AlterarProdutoConvenio lngPKID, _
                                              txtProdutoConvenio.Text, _
                                              mskPercComissao.Text, _
                                              mskPercVitalicio.Text, _
                                              strDatDesativacao, _
                                              strStatus
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir ProdutoConvenio
    objProdutoConvenio.InserirProdutoConvenio lngCONVENIOID, _
                                              txtProdutoConvenio.Text, _
                                              mskPercComissao.Text, _
                                              mskPercVitalicio.Text, _
                                              strDatDesativacao
  End If
  Set objProdutoConvenio = Nothing
  blnRetorno = True
  blnFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  cmdOk.Enabled = True
End Sub


Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  If Not Valida_String(txtProdutoConvenio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o produto do convênio" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskPercComissao, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o percentual da comissão" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskPercVitalicio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o percentual vitalício" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Option(optStatus, blnSetarFocoControle) Then
    strMsg = strMsg & "Slecionar o status" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserProdutoConvenioInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserProdutoConvenioInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco txtProdutoConvenio
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserProdutoConvenioInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objProdutoConvenio      As busApler.clsProdutoConvenio
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 5370
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , cmdIncluir, cmdAlterar
  '
  'Limpar Campos
  LimparCampos
  txtConvenio.Text = strDescrConvenio
  tabDetalhes_Click 1
  If Status = tpStatus_Incluir Then
    '
    optStatus(0).Value = True
    'Visible
    optStatus(0).Visible = False
    optStatus(1).Visible = False
    Label5(5).Visible = False
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objProdutoConvenio = New busApler.clsProdutoConvenio
    Set objRs = objProdutoConvenio.SelecionarProdutoConvenioPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      txtProdutoConvenio.Text = objRs.Fields("DESCRICAO").Value & ""
      INCLUIR_VALOR_NO_MASK mskPercComissao, objRs.Fields("PERCCOMISSAO").Value & "", TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskPercVitalicio, objRs.Fields("PERCVITALICIO").Value & "", TpMaskMoeda
      If objRs.Fields("STATUS").Value & "" = "A" Then
        optStatus(0).Value = True
        optStatus(1).Value = False
      ElseIf objRs.Fields("STATUS").Value & "" = "I" Then
        optStatus(0).Value = False
        optStatus(1).Value = True
      Else
        optStatus(0).Value = False
        optStatus(1).Value = False
      End If
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objProdutoConvenio = Nothing
    'Visible
    optStatus(0).Visible = True
    optStatus(1).Visible = True
    Label5(5).Visible = True
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = True
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
  If Not blnFechar Then Cancel = True
End Sub

Private Sub grdSubProduto_UnboundReadDataEx( _
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
               Offset + intI, SPROD_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, SPROD_COLUNASMATRIZ, SPROD_LINHASMATRIZ, SPROD_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, SPROD_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserConvenioInc.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub mskPercVitalicio_GotFocus()
  Seleciona_Conteudo_Controle mskPercVitalicio
End Sub
Private Sub mskPercVitalicio_LostFocus()
  Pintar_Controle mskPercVitalicio, tpCorContr_Normal
End Sub
Private Sub mskPercComissao_GotFocus()
  Seleciona_Conteudo_Controle mskPercComissao
End Sub
Private Sub mskPercComissao_LostFocus()
  Pintar_Controle mskPercComissao, tpCorContr_Normal
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    grdSubProduto.Enabled = False
    pictrava(0).Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco txtProdutoConvenio
  Case 1
    'Sub-Produto
    grdSubProduto.Enabled = True
    pictrava(0).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = True
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    'Montar RecordSet
    SPROD_COLUNASMATRIZ = grdSubProduto.Columns.Count
    SPROD_LINHASMATRIZ = 0
    SPROD_MontaMatriz
    grdSubProduto.Bookmark = Null
    grdSubProduto.ReBind
    grdSubProduto.ApproxCount = SPROD_LINHASMATRIZ
    '
    SetarFoco grdSubProduto
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "Apler.frmUserProdutoConvenioInc.tabDetalhes"
  AmpN
End Sub

Private Sub txtProdutoConvenio_GotFocus()
  Seleciona_Conteudo_Controle txtProdutoConvenio
End Sub
Private Sub txtProdutoConvenio_LostFocus()
  Pintar_Controle txtProdutoConvenio, tpCorContr_Normal
End Sub


Public Sub SPROD_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busApler.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busApler.clsGeral
  '
  strSql = "SELECT SUBPRODUTO.PKID, SUBPRODUTO.DESCRICAO, SUBPRODUTO.DATADESATIVACAO, CASE SUBPRODUTO.STATUS WHEN 'A' THEN 'Ativo' ELSE 'Inativo' END  " & _
          "FROM SUBPRODUTO " & _
          "WHERE SUBPRODUTO.PRODUTOID = " & lngPKID & _
          " ORDER BY SUBPRODUTO.DESCRICAO"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    SPROD_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim SPROD_Matriz(0 To SPROD_COLUNASMATRIZ - 1, 0 To SPROD_LINHASMATRIZ - 1)
  Else
    ReDim SPROD_Matriz(0 To SPROD_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To SPROD_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To SPROD_COLUNASMATRIZ - 1  'varre as colunas
          SPROD_Matriz(intJ, intI) = objRs(intJ) & ""
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


