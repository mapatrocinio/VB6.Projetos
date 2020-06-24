VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserConvAssocInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de convênio para associado"
   ClientHeight    =   5040
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5040
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5040
      Left            =   8430
      ScaleHeight     =   5040
      ScaleWidth      =   1860
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4605
         Left            =   90
         ScaleHeight     =   4545
         ScaleWidth      =   1605
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   300
         Width           =   1665
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   3570
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4815
      Left            =   120
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userConvAssocInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Beneficiários"
      TabPicture(1)   =   "userConvAssocInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdBeneficiario"
      Tab(1).Control(0).Enabled=   0   'False
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
         TabIndex        =   13
         Top             =   420
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2385
            Index           =   0
            Left            =   120
            ScaleHeight     =   2385
            ScaleWidth      =   7575
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   180
            Width           =   7575
            Begin VB.ComboBox cboPlano 
               Height          =   315
               Left            =   1290
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   420
               Width           =   6105
            End
            Begin VB.TextBox txtAssociado 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtAssociado"
               Top             =   90
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskDtInicio 
               Height          =   255
               Left            =   1290
               TabIndex        =   2
               Top             =   780
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDtTermino 
               Height          =   255
               Left            =   1290
               TabIndex        =   3
               Top             =   1080
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Término"
               Height          =   195
               Index           =   2
               Left            =   60
               TabIndex        =   18
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Associado"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   17
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Início"
               Height          =   195
               Index           =   21
               Left            =   60
               TabIndex        =   16
               Top             =   795
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Plano"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   15
               Top             =   450
               Width           =   1215
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdBeneficiario 
         Height          =   4245
         Left            =   -74940
         OleObjectBlob   =   "userConvAssocInc.frx":0038
         TabIndex        =   4
         Top             =   390
         Width           =   7905
      End
   End
End
Attribute VB_Name = "frmUserConvAssocInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngASSOCIADOID           As Long
Public strNomeAssociado         As String

Private blnPrimeiraVez          As Boolean

Dim BEN_COLUNASMATRIZ           As Long
Dim BEN_LINHASMATRIZ              As Long
Private BEN_Matriz()            As String

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Valor ConvAssoc
  LimparCampoTexto txtAssociado
  LimparCampoCombo cboPlano
  LimparCampoMask mskDtInicio
  LimparCampoMask mskDtTermino
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserConvAssocInc.LimparCampos]", _
            Err.Description
End Sub

Private Sub cboPlano_LostFocus()
  Pintar_Controle cboPlano, tpCorContr_Normal
End Sub



Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 1
    If Not IsNumeric(grdBeneficiario.Columns("PKID").Value & "") Then
      MsgBox "Selecione um Beneficiário !", vbExclamation, TITULOSISTEMA
      SetarFoco grdBeneficiario
      Exit Sub
    End If

    frmUserBeneficiarioInc.lngPKID = grdBeneficiario.Columns("PKID").Value
    frmUserBeneficiarioInc.lngTABCONVASSOCID = lngPKID
    frmUserBeneficiarioInc.strNomeAssociado = txtAssociado.Text
    frmUserBeneficiarioInc.strNomeConvenio = cboPlano.Text
    frmUserBeneficiarioInc.Status = tpStatus_Alterar
    frmUserBeneficiarioInc.Show vbModal
    '
    If frmUserBeneficiarioInc.blnRetorno Then
      BEN_MontaMatriz
      grdBeneficiario.Bookmark = Null
      grdBeneficiario.ReBind
      grdBeneficiario.ApproxCount = BEN_LINHASMATRIZ
    End If
    SetarFoco grdBeneficiario
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
  Dim objBeneficiario     As busApler.clsBeneficiario
  '
  Select Case tabDetalhes.Tab
  Case 1 'Exclusão de Beneficiario
    '
    If Len(Trim(grdBeneficiario.Columns("PKID").Value & "")) = 0 Then
      MsgBox "Selecione um Beneficiário.", vbExclamation, TITULOSISTEMA
      SetarFoco grdBeneficiario
      Exit Sub
    End If
    '
    Set objBeneficiario = New busApler.clsBeneficiario
    '
    If MsgBox("Confirma exclusão do Beneficiário " & grdBeneficiario.Columns("Nome").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdBeneficiario
      Exit Sub
    End If
    'OK
    objBeneficiario.ExcluirBeneficiario CLng(grdBeneficiario.Columns("PKID").Value)
    '
    BEN_MontaMatriz
    grdBeneficiario.Bookmark = Null
    grdBeneficiario.ReBind
    grdBeneficiario.ApproxCount = BEN_LINHASMATRIZ

    Set objBeneficiario = Nothing
    SetarFoco grdBeneficiario
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
  Case 1
    frmUserBeneficiarioInc.Status = tpStatus_Incluir
    frmUserBeneficiarioInc.lngPKID = 0
    frmUserBeneficiarioInc.lngTABCONVASSOCID = lngPKID
    frmUserBeneficiarioInc.strNomeAssociado = txtAssociado.Text
    frmUserBeneficiarioInc.strNomeConvenio = cboPlano.Text
    frmUserBeneficiarioInc.Show vbModal

    If frmUserBeneficiarioInc.blnRetorno Then
      BEN_MontaMatriz
      grdBeneficiario.Bookmark = Null
      grdBeneficiario.ReBind
      grdBeneficiario.ApproxCount = BEN_LINHASMATRIZ
    End If
    SetarFoco grdBeneficiario
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdOK_Click()
  Dim objConvAssoc              As busApler.clsConvAssoc
  Dim objGeral                  As busApler.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngPLANOCONVENIOID        As Long
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busApler.clsGeral
  Set objConvAssoc = New busApler.clsConvAssoc
  'PLANO CONVENIO
  lngPLANOCONVENIOID = 0
  strSql = "SELECT PKID FROM PLANOCONVENIO WHERE NOME = " & Formata_Dados(cboPlano.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngPLANOCONVENIOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'Validar se valor plano já cadastrado
  strSql = "SELECT * FROM TAB_CONVASSOC " & _
    " WHERE TAB_CONVASSOC.PLANOCONVENIOID = " & Formata_Dados(lngPLANOCONVENIOID, tpDados_Longo) & _
    " AND TAB_CONVASSOC.ASSOCIADOID = " & Formata_Dados(lngASSOCIADOID, tpDados_Longo) & _
    " AND TAB_CONVASSOC.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle cboPlano, tpCorContr_Erro
    TratarErroPrevisto "Plano já associado ao associado"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objConvAssoc = Nothing
    cmdOk.Enabled = True
    SetarFoco cboPlano
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar ConvAssoc
    objConvAssoc.AlterarConvAssoc lngPKID, _
                                  lngPLANOCONVENIOID, _
                                  mskDtInicio.Text, _
                                  IIf(mskDtTermino.ClipText = "", "", mskDtTermino.Text)
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir ConvAssoc
    objConvAssoc.InserirConvAssoc lngASSOCIADOID, _
                                  lngPLANOCONVENIOID, _
                                  mskDtInicio.Text, _
                                  IIf(mskDtTermino.ClipText = "", "", mskDtTermino.Text)
  End If
  Set objConvAssoc = Nothing
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
  If Not Valida_String(cboPlano, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o plano" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskDtInicio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de início válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskDtTermino, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de término válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserConvAssocInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserConvAssocInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco cboPlano
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserConvAssocInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objConvAssoc            As busApler.clsConvAssoc
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 5520
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , cmdIncluir, cmdAlterar
  '
  'Limpar Campos
  LimparCampos
  'Beneficiario Convênio
  strSql = "Select PLANOCONVENIO.NOME from PLANOCONVENIO " & _
    "INNER JOIN CONVENIO ON CONVENIO.PKID = PLANOCONVENIO.CONVENIOID " & _
    "WHERE EXISTS (SELECT TIPOCONVENIOID FROM TAB_TPCONVENIOVRPLANO T " & _
    "     INNER JOIN ASSOCIADO ON T.VALORPLANOID = ASSOCIADO.VALORPLANOID WHERE CONVENIO.TIPOCONVENIOID = T.TIPOCONVENIOID " & _
    "     AND ASSOCIADO.PKID = " & Formata_Dados(lngASSOCIADOID, tpDados_Longo) & ")" & _
    " OR EXISTS (SELECT PKID FROM TAB_CONVASSOC T WHERE PLANOCONVENIO.PKID = T.PLANOCONVENIOID " & _
    "     AND T.ASSOCIADOID = " & Formata_Dados(lngASSOCIADOID, tpDados_Longo) & ")" & _
    "ORDER BY PLANOCONVENIO.NOME"
  
  PreencheCombo cboPlano, strSql, False, True
  '
  txtAssociado.Text = strNomeAssociado
  tabDetalhes_Click 1
  If Status = tpStatus_Incluir Then
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = False
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objConvAssoc = New busApler.clsConvAssoc
    Set objRs = objConvAssoc.SelecionarConvAssocPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      If objRs.Fields("DESCR_PLANOCONVENIO").Value & "" <> "" Then
        cboPlano.Text = objRs.Fields("DESCR_PLANOCONVENIO").Value & ""
      End If
      INCLUIR_VALOR_NO_MASK mskDtInicio, objRs.Fields("DATAINICIO").Value & "", TpMaskData
      INCLUIR_VALOR_NO_MASK mskDtTermino, objRs.Fields("DATATERMINO").Value & "", TpMaskData
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objConvAssoc = Nothing
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


Private Sub grdBeneficiario_UnboundReadDataEx( _
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
               Offset + intI, BEN_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, BEN_COLUNASMATRIZ, BEN_LINHASMATRIZ, BEN_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, BEN_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserConvAssocInc.grdGeral_UnboundReadDataEx]"
End Sub


Private Sub mskDtInicio_GotFocus()
  Seleciona_Conteudo_Controle mskDtInicio
End Sub
Private Sub mskDtInicio_LostFocus()
  Pintar_Controle mskDtInicio, tpCorContr_Normal
End Sub

Private Sub mskDtTermino_GotFocus()
  Seleciona_Conteudo_Controle mskDtTermino
End Sub
Private Sub mskDtTermino_LostFocus()
  Pintar_Controle mskDtTermino, tpCorContr_Normal
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    grdBeneficiario.Enabled = False
    pictrava(0).Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco cboPlano
  Case 1
    grdBeneficiario.Enabled = True
    pictrava(0).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = True
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    'Montar RecordSet
    BEN_COLUNASMATRIZ = grdBeneficiario.Columns.Count
    BEN_LINHASMATRIZ = 0
    BEN_MontaMatriz
    grdBeneficiario.Bookmark = Null
    grdBeneficiario.ReBind
    grdBeneficiario.ApproxCount = BEN_LINHASMATRIZ
    '
    SetarFoco grdBeneficiario
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "Apler.frmUserBeneficiarioInc.tabDetalhes"
  AmpN
End Sub


Public Sub BEN_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busApler.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busApler.clsGeral
  '
  strSql = "SELECT BENEFICIARIO.PKID, BENEFICIARIO.NOME, BENEFICIARIO.CPF, BENEFICIARIO.DATANASCIMENTO " & _
          "FROM BENEFICIARIO " & _
          "WHERE BENEFICIARIO.TABCONVASSOCID = " & lngPKID & _
          " ORDER BY BENEFICIARIO.NOME"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    BEN_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim BEN_Matriz(0 To BEN_COLUNASMATRIZ - 1, 0 To BEN_LINHASMATRIZ - 1)
  Else
    ReDim BEN_Matriz(0 To BEN_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To BEN_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To BEN_COLUNASMATRIZ - 1  'varre as colunas
          BEN_Matriz(intJ, intI) = objRs(intJ) & ""
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

