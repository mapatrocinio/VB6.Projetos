VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserContratoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de contrato de empresa"
   ClientHeight    =   5010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5010
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5010
      Left            =   8430
      ScaleHeight     =   5010
      ScaleWidth      =   1860
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4635
         Left            =   90
         ScaleHeight     =   4575
         ScaleWidth      =   1605
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   270
         Width           =   1665
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3570
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4785
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   90
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   8440
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userContratoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Obra"
      TabPicture(1)   =   "userContratoInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdObra"
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
         Height          =   3165
         Left            =   120
         TabIndex        =   15
         Top             =   420
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2925
            Index           =   0
            Left            =   120
            ScaleHeight     =   2925
            ScaleWidth      =   7575
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtAno 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2460
               MaxLength       =   100
               TabIndex        =   3
               Text            =   "txtAno"
               Top             =   780
               Width           =   615
            End
            Begin VB.TextBox txtSequencial 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   2
               Text            =   "txtSequencial"
               Top             =   780
               Width           =   1125
            End
            Begin VB.ComboBox cboFuncionario 
               BackColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   1320
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   420
               Width           =   6105
            End
            Begin VB.TextBox txtEmpresa 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtEmpresa"
               Top             =   90
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskInicio 
               Height          =   255
               Left            =   1320
               TabIndex        =   4
               Top             =   1110
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskFim 
               Height          =   255
               Left            =   5820
               TabIndex        =   5
               Top             =   1110
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Número"
               Height          =   195
               Index           =   5
               Left            =   60
               TabIndex        =   21
               Top             =   765
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Funcionário"
               Height          =   195
               Index           =   24
               Left            =   60
               TabIndex        =   20
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Fim"
               Height          =   195
               Index           =   7
               Left            =   4560
               TabIndex        =   19
               Top             =   1110
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Início"
               Height          =   195
               Index           =   6
               Left            =   60
               TabIndex        =   18
               Top             =   1110
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Empresa"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   17
               Top             =   105
               Width           =   1215
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdObra 
         Height          =   4215
         Left            =   -74910
         OleObjectBlob   =   "userContratoInc.frx":0038
         TabIndex        =   6
         Top             =   390
         Width           =   7965
      End
   End
End
Attribute VB_Name = "frmUserContratoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngEMPRESAID            As Long
Public strDescrEmpresa         As String

Private blnPrimeiraVez          As Boolean

Dim OBRA_COLUNASMATRIZ         As Long
Dim OBRA_LINHASMATRIZ          As Long
Private OBRA_Matriz()          As String

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Contrato
  LimparCampoTexto txtEmpresa
  LimparCampoCombo cboFuncionario
  LimparCampoTexto txtSequencial
  LimparCampoTexto txtAno
  LimparCampoMask mskInicio
  LimparCampoMask mskFim
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserObraInc.LimparCampos]", _
            Err.Description
End Sub

Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 1
    'Obra
    If Not IsNumeric(grdObra.Columns("PKID").Value & "") Then
      MsgBox "Selecione uma obra!", vbExclamation, TITULOSISTEMA
      SetarFoco grdObra
      Exit Sub
    End If
    frmUserObraInc.Status = tpStatus_Alterar
    frmUserObraInc.lngPKID = grdObra.Columns("PKID").Value
    frmUserObraInc.lngCONTRATOID = lngPKID
    frmUserObraInc.strDescrEmpresa = txtEmpresa.Text
    frmUserObraInc.strDescrContrato = Format(txtSequencial.Text, "0000") & "/" & txtAno.Text
    frmUserObraInc.Show vbModal
    '
    If frmUserObraInc.blnRetorno Then
      OBRA_MontaMatriz
      grdObra.Bookmark = Null
      grdObra.ReBind
      grdObra.ApproxCount = OBRA_LINHASMATRIZ
    End If
    SetarFoco grdObra
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
  Dim objObra     As busSisLoc.clsObra
  Dim objGeral    As busSisLoc.clsGeral
  Dim objRs       As ADODB.Recordset
  Dim strSql      As String
  '
  Select Case tabDetalhes.Tab
  Case 1 'Exclusão de Obra
    '
    If Len(Trim(grdObra.Columns("PKID").Value & "")) = 0 Then
      MsgBox "Selecione uma obra do contrato.", vbExclamation, TITULOSISTEMA
      SetarFoco grdObra
      Exit Sub
    End If
    '
    Set objGeral = New busSisLoc.clsGeral
    'NF
    strSql = "SELECT * FROM NF WHERE OBRAID = " & Formata_Dados(grdObra.Columns("PKID").Value, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      TratarErroPrevisto "Não é possível excluir a obra, pois existem NFs associadas a ele.", "[cmdExcluir_Click]"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    'DEVOLUCAO
    strSql = "SELECT * FROM DEVOLUCAO WHERE OBRAID = " & Formata_Dados(grdObra.Columns("PKID").Value, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      TratarErroPrevisto "Não é possível excluir a obra, pois existem devoluções associadas a ele.", "[cmdExcluir_Click]"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    '
    If MsgBox("Confirma exclusão da obra " & grdObra.Columns("Obra").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdObra
      Exit Sub
    End If
    'OK
    Set objObra = New busSisLoc.clsObra
    objObra.ExcluirObra CLng(grdObra.Columns("PKID").Value)
    '
    OBRA_MontaMatriz
    grdObra.Bookmark = Null
    grdObra.ReBind
    grdObra.ApproxCount = OBRA_LINHASMATRIZ

    Set objObra = Nothing
    SetarFoco grdObra
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
  Case 1 'Obra
    frmUserObraInc.Status = tpStatus_Incluir
    frmUserObraInc.lngPKID = 0
    frmUserObraInc.lngCONTRATOID = lngPKID
    frmUserObraInc.strDescrEmpresa = txtEmpresa.Text
    frmUserObraInc.strDescrContrato = Format(txtSequencial.Text, "0000") & "/" & txtAno.Text
    frmUserObraInc.Show vbModal

    If frmUserObraInc.blnRetorno Then
      OBRA_MontaMatriz
      grdObra.Bookmark = Null
      grdObra.ReBind
      grdObra.ApproxCount = OBRA_LINHASMATRIZ
    End If
    SetarFoco grdObra
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdOK_Click()
  Dim objContrato               As busSisLoc.clsContrato
  Dim objGeral                  As busSisLoc.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngFUNCIONARIOID          As Long
  Dim strNumero                 As String
  Dim strSequencial             As String
  Dim lngAno                    As Long
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisLoc.clsGeral
  Set objContrato = New busSisLoc.clsContrato
  'FUNCIONARIO
  lngFUNCIONARIOID = 0
  strSql = "SELECT PKID FROM PESSOA WHERE NOME = " & Formata_Dados(cboFuncionario.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngFUNCIONARIOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing

  'Validar se contrato já cadastrado
  strSql = "SELECT * FROM CONTRATO " & _
    " WHERE CONTRATO.SEQUENCIAL = " & Formata_Dados(txtSequencial.Text, tpDados_Longo) & _
    " AND CONTRATO.ANO = " & Formata_Dados(txtAno.Text, tpDados_Longo) & _
    " AND CONTRATO.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtSequencial, tpCorContr_Erro
    TratarErroPrevisto "Contrato já cadastrado"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objContrato = Nothing
    cmdOk.Enabled = True
    SetarFoco txtSequencial
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Contrato
    objContrato.AlterarContrato lngPKID, _
                                mskInicio.Text, _
                                mskFim.Text, _
                                lngFUNCIONARIOID & ""
    '
  ElseIf Status = tpStatus_Incluir Then
    'Obter dados do contrato
    'lngAno = Right(mskInicio.Text, 4)
    'strSequencial = RetornaGravaCampoSequencialCtrto("SEQUENCIAL", lngAno) & ""
    lngAno = txtAno.Text
    strSequencial = txtSequencial.Text
    strNumero = "RF" & Format(strSequencial, "0000") & "/" & lngAno

    'Inserir Contrato
    objContrato.InserirContrato strNumero, _
                                strSequencial, _
                                lngAno & "", _
                                mskInicio.Text, _
                                mskFim.Text, _
                                lngEMPRESAID & "", _
                                lngFUNCIONARIOID & ""
  End If
  Set objContrato = Nothing
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
  If Not Valida_Moeda(txtSequencial, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o sequencial válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(txtAno, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o ano válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  
  If strMsg = "" Then
    If Len(txtAno.Text) <> 4 Then
      SetarFoco txtAno
      blnSetarFocoControle = False
      strMsg = strMsg & "Preencher o ano válido" & vbCrLf
      tabDetalhes.Tab = 0
    End If
  End If
  If Not Valida_Data(mskInicio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de início" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskFim, TpnaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de fim válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserObraInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserObraInc.ValidaCampos]", _
            Err.Description
End Function

Public Sub OBRA_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisLoc.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisLoc.clsGeral
  '
  strSql = "SELECT OBRA.PKID, OBRA.DESCRICAO " & _
          "FROM OBRA " & _
          "WHERE OBRA.CONTRATOID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
          " ORDER BY OBRA.DESCRICAO"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    OBRA_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim OBRA_Matriz(0 To OBRA_COLUNASMATRIZ - 1, 0 To OBRA_LINHASMATRIZ - 1)
  Else
    ReDim OBRA_Matriz(0 To OBRA_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To OBRA_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To OBRA_COLUNASMATRIZ - 1  'varre as colunas
          OBRA_Matriz(intJ, intI) = objRs(intJ) & ""
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


Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    If Status = tpStatus_Alterar Then
      SetarFoco mskInicio
    Else
      SetarFoco txtSequencial
    End If
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserObraInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objContrato             As busSisLoc.clsContrato
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 5490
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , cmdIncluir, cmdAlterar
  '
  'Limpar Campos
  LimparCampos
  'Fucnionário
  strSql = "Select NOME from PESSOA ORDER BY NOME"
  PreencheCombo cboFuncionario, strSql, False, True
  
  txtEmpresa.Text = strDescrEmpresa
  INCLUIR_VALOR_NO_COMBO gsNomeUsuCompleto & "", cboFuncionario
  '
  txtAno.Enabled = True
  txtSequencial.Enabled = True
  '
  tabDetalhes_Click 1
  If Status = tpStatus_Incluir Then
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    txtAno.Enabled = False
    txtSequencial.Enabled = False
    Set objContrato = New busSisLoc.clsContrato
    Set objRs = objContrato.SelecionarContrato(lngPKID)
    '
    If Not objRs.EOF Then
      
      INCLUIR_VALOR_NO_COMBO objRs.Fields("FUNCIONARIO").Value & "", cboFuncionario
      txtSequencial.Text = objRs.Fields("SEQUENCIAL").Value & ""
      txtAno.Text = objRs.Fields("ANO").Value & ""
      INCLUIR_VALOR_NO_MASK mskInicio, objRs.Fields("DATAINICIO").Value & "", TpMaskData
      INCLUIR_VALOR_NO_MASK mskFim, objRs.Fields("DATAFIM").Value & "", TpMaskData
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objContrato = Nothing
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

Private Sub grdObra_UnboundReadDataEx( _
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
               Offset + intI, OBRA_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, OBRA_COLUNASMATRIZ, OBRA_LINHASMATRIZ, OBRA_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, OBRA_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserObraInc.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub mskFim_GotFocus()
  Seleciona_Conteudo_Controle mskFim
End Sub
Private Sub mskFim_LostFocus()
  Pintar_Controle mskFim, tpCorContr_Normal
End Sub

Private Sub mskInicio_GotFocus()
  Seleciona_Conteudo_Controle mskInicio
End Sub
Private Sub mskInicio_LostFocus()
  Pintar_Controle mskInicio, tpCorContr_Normal
End Sub


Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    grdObra.Enabled = False
    pictrava(0).Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    If Status = tpStatus_Alterar Then
      SetarFoco mskInicio
    Else
      SetarFoco txtSequencial
    End If
  Case 1
    'Obra
    grdObra.Enabled = True
    pictrava(0).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = True
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    'Montar RecordSet
    OBRA_COLUNASMATRIZ = grdObra.Columns.Count
    OBRA_LINHASMATRIZ = 0
    OBRA_MontaMatriz
    grdObra.Bookmark = Null
    grdObra.ReBind
    grdObra.ApproxCount = OBRA_LINHASMATRIZ
    '
    SetarFoco grdObra
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "SisLoc.frmUserObraInc.tabDetalhes"
  AmpN
End Sub

Private Sub txtAno_GotFocus()
  Seleciona_Conteudo_Controle txtAno
End Sub
Private Sub txtAno_LostFocus()
  Pintar_Controle txtAno, tpCorContr_Normal
End Sub

Private Sub txtSequencial_GotFocus()
  Seleciona_Conteudo_Controle txtSequencial
End Sub
Private Sub txtSequencial_LostFocus()
  Pintar_Controle txtSequencial, tpCorContr_Normal
End Sub
