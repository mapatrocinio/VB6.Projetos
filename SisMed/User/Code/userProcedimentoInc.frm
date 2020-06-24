VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserProcedimentoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Procedimento"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5715
      Left            =   8520
      ScaleHeight     =   5715
      ScaleWidth      =   1860
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4695
         Left            =   120
         ScaleHeight     =   4635
         ScaleWidth      =   1605
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   810
         Width           =   1665
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5415
      Left            =   120
      TabIndex        =   16
      TabStop         =   0   'False
      Top             =   90
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9551
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do procedimento"
      TabPicture(0)   =   "userProcedimentoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picTrava(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Receita"
      TabPicture(1)   =   "userProcedimentoInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdReceita"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.PictureBox picTrava 
         BorderStyle     =   0  'None
         Height          =   1695
         Index           =   0
         Left            =   120
         ScaleHeight     =   1695
         ScaleWidth      =   7695
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   480
         Width           =   7695
         Begin VB.Frame Frame3 
            BackColor       =   &H00C0C0C0&
            Caption         =   "Cadastro"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1665
            Left            =   0
            TabIndex        =   13
            Top             =   0
            Width           =   7695
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   5250
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   23
               Top             =   1230
               Width           =   2235
               Begin VB.OptionButton optConsulta 
                  Caption         =   "Sim"
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  TabIndex        =   5
                  Top             =   60
                  Width           =   645
               End
               Begin VB.OptionButton optConsulta 
                  Caption         =   "Não"
                  Height          =   195
                  Index           =   1
                  Left            =   780
                  TabIndex        =   6
                  Top             =   60
                  Width           =   645
               End
            End
            Begin VB.PictureBox Picture3 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1560
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   21
               Top             =   1230
               Width           =   2235
               Begin VB.OptionButton optAceitaValor 
                  Caption         =   "Não"
                  Height          =   195
                  Index           =   1
                  Left            =   780
                  TabIndex        =   4
                  Top             =   60
                  Width           =   645
               End
               Begin VB.OptionButton optAceitaValor 
                  Caption         =   "Sim"
                  Height          =   195
                  Index           =   0
                  Left            =   0
                  TabIndex        =   3
                  Top             =   60
                  Width           =   645
               End
            End
            Begin VB.ComboBox cboTipo 
               Height          =   315
               Left            =   1545
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   570
               Width           =   5925
            End
            Begin VB.TextBox txtProcedimento 
               Height          =   285
               Left            =   1560
               MaxLength       =   100
               TabIndex        =   0
               Text            =   "txtProcedimento"
               Top             =   240
               Width           =   5895
            End
            Begin MSMask.MaskEdBox mskValor 
               Height          =   255
               Left            =   1545
               TabIndex        =   2
               Top             =   930
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;-#,##0.00"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Consulta?"
               Height          =   255
               Index           =   0
               Left            =   3810
               TabIndex        =   24
               Top             =   1260
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Aceita Valor?"
               Height          =   255
               Index           =   12
               Left            =   120
               TabIndex        =   22
               Top             =   1260
               Width           =   1215
            End
            Begin VB.Label lblPercentual 
               Caption         =   "Valor"
               Height          =   255
               Left            =   120
               TabIndex        =   20
               Top             =   930
               Width           =   1455
            End
            Begin VB.Label Label6 
               Caption         =   "Tipo"
               Height          =   255
               Index           =   9
               Left            =   120
               TabIndex        =   19
               Top             =   600
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Procedimento"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   18
               Top             =   240
               Width           =   1455
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdReceita 
         Height          =   4725
         Left            =   -74940
         OleObjectBlob   =   "userProcedimentoInc.frx":0038
         TabIndex        =   7
         Top             =   390
         Width           =   7905
      End
   End
End
Attribute VB_Name = "frmUserProcedimentoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngPROCEDIMENTOID                As Long
Public blnRetorno                   As Boolean
Public blnFechar                    As Boolean
Private blnPrimeiraVez            As Boolean

Dim REC_COLUNASMATRIZ         As Long
Dim REC_LINHASMATRIZ          As Long
Private REC_Matriz()          As String


Private Sub cboTipo_LostFocus()
  Pintar_Controle cboTipo, tpCorContr_Normal
End Sub

Private Sub cmdCancelar_Click()
  blnFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim objProcedimento               As busSisMed.clsProcedimento
  Dim objGer                  As busSisMed.clsGeral
  Dim strTipoProcedimentoId        As String
  Dim objGeral                As busSisMed.clsGeral
  Dim strAceitaValor          As String
  Dim strConsulta             As String

  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração de Grupo cardápio
    If Not ValidaCampos Then Exit Sub
    'Valida se Grupo cardápio já cadastrado
    Set objGer = New busSisMed.clsGeral
    strSql = "Select * From PROCEDIMENTO WHERE PROCEDIMENTO = " & Formata_Dados(txtProcedimento.Text, tpDados_Texto, tpNulo_Aceita) & _
      " AND PKID <> " & Formata_Dados(lngPROCEDIMENTOID, tpDados_Longo, tpNulo_NaoAceita)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGer = Nothing
      TratarErroPrevisto "Procedimento já cadastrado", "cmdOK_Click"
      Pintar_Controle txtProcedimento, tpCorContr_Erro
      SetarFoco txtProcedimento
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    'Obter Tipo procedimento
    Set objGeral = New busSisMed.clsGeral
    strSql = "SELECT PKID FROM TIPOPROCEDIMENTO WHERE TIPOPROCEDIMENTO = " & Formata_Dados(cboTipo.Text, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    strTipoProcedimentoId = ""
    If Not objRs.EOF Then
      strTipoProcedimentoId = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGer = Nothing
    '
    Set objProcedimento = New busSisMed.clsProcedimento
    '
    If optAceitaValor(0).Value Then
      strAceitaValor = "S"
    ElseIf optAceitaValor(1).Value Then
      strAceitaValor = "N"
    Else
      strAceitaValor = ""
    End If
    If optConsulta(0).Value Then
      strConsulta = "S"
    ElseIf optConsulta(1).Value Then
      strConsulta = "N"
    Else
      strConsulta = ""
    End If
    
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      '
      objProcedimento.AlterarProcedimento lngPROCEDIMENTOID, _
                                          txtProcedimento.Text, _
                                          strTipoProcedimentoId, _
                                          mskValor.ClipText, _
                                          strAceitaValor, _
                                          strConsulta
                            
      blnRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Pega Informações para inserir
      '
      objProcedimento.InserirProcedimento txtProcedimento.Text, _
                                          strTipoProcedimentoId, _
                                          mskValor.ClipText, _
                                          strAceitaValor, _
                                          strConsulta
      '
    End If
    Set objProcedimento = Nothing
  End Select
  If Status = tpStatus_Alterar Then
    blnRetorno = True
    blnFechar = True
    Unload Me
  ElseIf Status = tpStatus_Incluir Then
    'Selecionar prontuario pelo nome
    Status = tpStatus_Alterar
    'Reload na tela
    Form_Load
    'Acerta tabs
    blnRetorno = True
  End If
  tabDetalhes.Tab = 1
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg     As String
  Dim blnSetarFoco  As Boolean
  '
  blnSetarFoco = True
  If Not Valida_String(txtProcedimento, TpObrigatorio, blnSetarFoco) Then
    strMsg = strMsg & "Informar o Procedimento" & vbCrLf
  End If
'''  If Not Valida_String(cboTipo, TpObrigatorio, blnSetarFoco) Then
'''    strMsg = strMsg & "Selecionar o Tipo de procedimento" & vbCrLf
'''  End If
  If Not Valida_Moeda(mskValor, TpObrigatorio, blnSetarFoco) Then
    strMsg = strMsg & "Informar o valor válido" & vbCrLf
  End If
  If Not Valida_Option(optAceitaValor, blnSetarFoco) Then
    strMsg = strMsg & "Selecionar se procedimento aceita insersão manual de valor na GR" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Option(optConsulta, blnSetarFoco) Then
    strMsg = strMsg & "Selecionar se procedimento é uma consulta" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserProcedimentoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Function



Private Sub grdReceita_UnboundReadDataEx( _
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
               Offset + intI, REC_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, REC_COLUNASMATRIZ, REC_LINHASMATRIZ, REC_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, REC_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserProcedimentoInc.grdGeral_UnboundReadDataEx]"
End Sub


Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    tabDetalhes.Tab = 0
    SetarFoco txtProcedimento
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserProcedimentoInc.Form_Activate]"
End Sub

Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Dim objFormProcReceita As SisMed.frmUserProcReceitaInc
  Select Case tabDetalhes.Tab
  Case 1
    'Proc Receita
    If Not IsNumeric(grdReceita.Columns("PKID").Value & "") Then
      MsgBox "Selecione um tipo de procedimento!", vbExclamation, TITULOSISTEMA
      SetarFoco grdReceita
      Exit Sub
    End If

    Set objFormProcReceita = New SisMed.frmUserProcReceitaInc
    objFormProcReceita.Status = tpStatus_Alterar
    objFormProcReceita.lngPKID = grdReceita.Columns("PKID").Value
    objFormProcReceita.lngPROCEDIMENTOID = lngPROCEDIMENTOID
    objFormProcReceita.strNomeProcedimento = txtProcedimento.Text
    objFormProcReceita.Show vbModal
    If objFormProcReceita.blnRetorno Then
      REC_MontaMatriz
      grdReceita.Bookmark = Null
      grdReceita.ReBind
      grdReceita.ApproxCount = REC_LINHASMATRIZ
    End If
    Set objFormProcReceita = Nothing
    SetarFoco grdReceita
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdExcluir_Click()
  On Error GoTo trata
  Dim objProcReceita      As busSisMed.clsProcReceita
  '
  Select Case tabDetalhes.Tab
  Case 1 'Exclusão de associado
    If Not IsNumeric(grdReceita.Columns("PKID").Value & "") Then
      MsgBox "Selecione um tipo de procedimento para exclusão.", vbExclamation, TITULOSISTEMA
      SetarFoco grdReceita
      Exit Sub
    End If
    '
    If MsgBox("Confirma exclusão do tipo Procedimento " & grdReceita.Columns("Tipo").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdReceita
      Exit Sub
    End If
    Set objProcReceita = New busSisMed.clsProcReceita
    objProcReceita.ExcluirProcReceita CLng(grdReceita.Columns("PKID").Value)
    Set objProcReceita = Nothing
    '
    REC_MontaMatriz
    grdReceita.Bookmark = Null
    grdReceita.ReBind
    SetarFoco grdReceita

  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdIncluir_Click()
  On Error GoTo trata
  Dim objFormProcReceita As SisMed.frmUserProcReceitaInc
  '
  Select Case tabDetalhes.Tab
  Case 1
    'PROCEDIMENTO RECEITA
    Set objFormProcReceita = New SisMed.frmUserProcReceitaInc
    objFormProcReceita.Status = tpStatus_Incluir
    objFormProcReceita.lngPKID = 0
    objFormProcReceita.lngPROCEDIMENTOID = lngPROCEDIMENTOID
    objFormProcReceita.strNomeProcedimento = txtProcedimento.Text
    objFormProcReceita.Show vbModal
    If objFormProcReceita.blnRetorno Then
      REC_MontaMatriz
      grdReceita.Bookmark = Null
      grdReceita.ReBind
      grdReceita.ApproxCount = REC_LINHASMATRIZ
    End If
    Set objFormProcReceita = Nothing
    SetarFoco grdReceita
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Resume Next
End Sub

Public Sub REC_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMed.clsGeral
  '
  strSql = "SELECT PROCEDIMENTORECEITA.PKID, PROCEDIMENTORECEITA.TIPO " & _
          "FROM PROCEDIMENTORECEITA " & _
          "WHERE PROCEDIMENTORECEITA.PROCEDIMENTOID = " & Formata_Dados(lngPROCEDIMENTOID, tpDados_Longo) & _
          " ORDER BY PROCEDIMENTORECEITA.TIPO"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    REC_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim REC_Matriz(0 To REC_COLUNASMATRIZ - 1, 0 To REC_LINHASMATRIZ - 1)
  Else
    ReDim REC_Matriz(0 To REC_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To REC_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To REC_COLUNASMATRIZ - 1  'varre as colunas
          REC_Matriz(intJ, intI) = objRs(intJ) & ""
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
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim objProcedimento  As busSisMed.clsProcedimento
  '
  blnFechar = False
  blnRetorno = False
  AmpS
  Me.Height = 6195
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , cmdIncluir, cmdAlterar
  '
  strSql = "SELECT TIPOPROCEDIMENTO FROM TIPOPROCEDIMENTO ORDER BY TIPOPROCEDIMENTO"
  PreencheCombo cboTipo, strSql, False, True
  '
  optAceitaValor(0).Value = False
  optAceitaValor(1).Value = False
  optConsulta(0).Value = False
  optConsulta(1).Value = False
  
  tabDetalhes_Click 1
  
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    LimparCampoTexto txtProcedimento
    LimparCampoMask mskValor
    optAceitaValor(1).Value = True
    optConsulta(1).Value = True
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = False
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objProcedimento = New busSisMed.clsProcedimento
    Set objRs = objProcedimento.ListarProcedimento(lngPROCEDIMENTOID)
    '
    If Not objRs.EOF Then
      txtProcedimento.Text = objRs.Fields("PROCEDIMENTO").Value & ""
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR").Value, TpMaskMoeda
      If objRs.Fields("TIPOPROCEDIMENTO").Value & "" <> "" Then
        cboTipo = objRs.Fields("TIPOPROCEDIMENTO").Value & ""
      End If
      If objRs.Fields("INDACEITAVALOR").Value & "" = "S" Then
        optAceitaValor(0).Value = True
        optAceitaValor(1).Value = False
      ElseIf objRs.Fields("INDACEITAVALOR").Value & "" = "N" Then
        optAceitaValor(0).Value = False
        optAceitaValor(1).Value = True
      Else
        optAceitaValor(0).Value = False
        optAceitaValor(1).Value = False
      End If
      If objRs.Fields("INDCONSULTA").Value & "" = "S" Then
        optConsulta(0).Value = True
        optConsulta(1).Value = False
      ElseIf objRs.Fields("INDCONSULTA").Value & "" = "N" Then
        optConsulta(0).Value = False
        optConsulta(1).Value = True
      Else
        optConsulta(0).Value = False
        optConsulta(1).Value = False
      End If
      
      '
    End If
    Set objProcedimento = Nothing
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = True
  End If
  
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub

Private Sub mskValor_GotFocus()
  Selecionar_Conteudo mskValor
End Sub

Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    'Dados cadastrais
    grdReceita.Enabled = False
    Frame3.Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco txtProcedimento
  Case 1
    'Receita
    grdReceita.Enabled = True
    Frame3.Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = True
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    '
    '
    'Montar RecordSet
    REC_COLUNASMATRIZ = grdReceita.Columns.Count
    REC_LINHASMATRIZ = 0
    REC_MontaMatriz
    grdReceita.Bookmark = Null
    grdReceita.ReBind
    grdReceita.ApproxCount = REC_LINHASMATRIZ
    '
    SetarFoco grdReceita
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "frmUserProcedimentoInc.tabDetalhes"
  AmpN
End Sub

Private Sub txtProcedimento_GotFocus()
  Selecionar_Conteudo txtProcedimento
End Sub

Private Sub txtProcedimento_LostFocus()
  Pintar_Controle txtProcedimento, tpCorContr_Normal
End Sub

