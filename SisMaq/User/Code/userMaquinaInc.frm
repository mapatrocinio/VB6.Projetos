VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserMaquinaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Máquina"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5610
      Left            =   8430
      ScaleHeight     =   5610
      ScaleWidth      =   1860
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2055
         Left            =   90
         ScaleHeight     =   1995
         ScaleWidth      =   1605
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1665
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   990
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5295
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userMaquinaInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Histórico Máquina"
      TabPicture(1)   =   "userMaquinaInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdMaquina"
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
         Height          =   4755
         Left            =   120
         TabIndex        =   10
         Top             =   330
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   4515
            Index           =   0
            Left            =   120
            ScaleHeight     =   4515
            ScaleWidth      =   7575
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtSerie 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1350
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtSerie"
               Top             =   90
               Width           =   6075
            End
            Begin VB.ComboBox cboTipo 
               Height          =   315
               Left            =   1350
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   780
               Width           =   6105
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1320
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   14
               TabStop         =   0   'False
               Top             =   1410
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
            Begin VB.TextBox txtNumero 
               Height          =   285
               Left            =   1350
               MaxLength       =   100
               TabIndex        =   1
               Text            =   "txtNumero"
               Top             =   435
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskCoeficiente 
               Height          =   255
               Left            =   1350
               TabIndex        =   3
               Top             =   1140
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0.0000;($#,##0.0000)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Série"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   18
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Coeficiente"
               Height          =   285
               Index           =   2
               Left            =   90
               TabIndex        =   17
               Top             =   1110
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Tipo"
               Height          =   285
               Index           =   24
               Left            =   90
               TabIndex        =   15
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Status"
               Height          =   315
               Index           =   5
               Left            =   90
               TabIndex        =   13
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Número"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   12
               Top             =   480
               Width           =   1215
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdMaquina 
         Height          =   4545
         Left            =   -74910
         OleObjectBlob   =   "userMaquinaInc.frx":0038
         TabIndex        =   19
         Top             =   390
         Width           =   7965
      End
   End
End
Attribute VB_Name = "frmUserMaquinaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngSERIEID               As Long
Public strDescSerie             As String
Public lngTIPOANTERIORID        As Long

Private blnPrimeiraVez          As Boolean

Dim MAQ_COLUNASMATRIZ           As Long
Dim MAQ_LINHASMATRIZ            As Long

Private MAQ_Matriz()            As String

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Maquina
  LimparCampoTexto txtNumero
  LimparCampoCombo cboTipo
  LimparCampoMask mskCoeficiente
  LimparCampoTexto txtSerie
  optStatus(0).Value = False
  optStatus(1).Value = False
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserMaquinaInc.LimparCampos]", _
            Err.Description
End Sub


Private Sub cboTipo_LostFocus()
  Pintar_Controle cboTipo, tpCorContr_Normal
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

Private Sub cmdOK_Click()
  Dim objMaquina                As busSisMaq.clsMaquina
  Dim objGeral                  As busSisMaq.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngTIPOID                 As Long
  Dim lngMAQUINAID              As Long
  Dim strStatus                 As String
  Dim strDataInicio             As String
  Dim strDataTermino            As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisMaq.clsGeral
  Set objMaquina = New busSisMaq.clsMaquina
  'MAQUINA
  lngMAQUINAID = 0
  strSql = "SELECT PKID FROM MAQUINA WHERE MAQUINA.EQUIPAMENTOID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
      " AND MAQUINA.STATUS = " & Formata_Dados("A", tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngMAQUINAID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'TIPO
  lngTIPOID = 0
  strSql = "SELECT PKID FROM TIPO WHERE TIPO.TIPO = " & Formata_Dados(cboTipo.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngTIPOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Status
  If optStatus(0).Value Then
    strStatus = "A"
  Else
    strStatus = "I"
  End If
  If lngTIPOID <> lngTIPOANTERIORID And Status = tpStatus_Alterar Then
    'Alterou tipo
    If MsgBox("ATENÇÃO: Você está alterando o tipo da máquina, isso irá gerar uma nova máquina. Confirma esta mudança?", vbYesNo, TITULOSISTEMA) = vbNo Then
      cmdOk.Enabled = True
      SetarFoco txtNumero
      tabDetalhes.Tab = 0
      Exit Sub
    End If
    'OK
  End If
  'Validar se funcionário já cadastrado
  strSql = "SELECT * FROM EQUIPAMENTO " & _
    " WHERE EQUIPAMENTO.NUMERO = " & Formata_Dados(txtNumero.Text, tpDados_Texto) & _
    " AND EQUIPAMENTO.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo) & _
    " AND EQUIPAMENTO.STATUS = " & Formata_Dados("A", tpDados_Texto)
    
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtNumero, tpCorContr_Erro
    TratarErroPrevisto "Maquina já cadastrada"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objMaquina = Nothing
    cmdOk.Enabled = True
    SetarFoco txtNumero
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Equipamento
    objMaquina.AlterarEquipamento lngPKID, _
                                  lngSERIEID, _
                                  txtNumero.Text, _
                                  mskCoeficiente.ClipText, _
                                  strStatus
    If lngTIPOID = lngTIPOANTERIORID Then
      'Alterar Máquina
      objMaquina.AlterarMaquina lngMAQUINAID, _
                                lngTIPOID, _
                                "", _
                                strStatus
    Else
      'Alterar Máquina para inativa
      strDataTermino = Format(Now, "DD/MM/YYYY hh:mm")
      strDataInicio = strDataTermino
      objMaquina.AlterarMaquina lngMAQUINAID, _
                                lngTIPOANTERIORID, _
                                strDataTermino, _
                                "I", _
                                gsNomeUsu
      
      'Inserir Máquina ativa
      objMaquina.InserirMaquina lngPKID, _
                                lngTIPOID, _
                                strDataInicio, _
                                strStatus
    End If
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Equipamento
    strDataInicio = Format(Now, "DD/MM/YYYY hh:mm")
    objMaquina.InserirEquipamento lngPKID, _
                                  lngSERIEID, _
                                  txtNumero.Text, _
                                  mskCoeficiente.ClipText, _
                                  strStatus, _
                                  lngTIPOID, _
                                  strDataInicio, _
                                  strStatus
    blnRetorno = True
    blnFechar = True
    Unload Me
  End If
  Set objMaquina = Nothing
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
  If Not Valida_String(txtNumero, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o nome" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(cboTipo, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o tipo" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskCoeficiente, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o coeficiente válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Option(optStatus, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o status" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserMaquinaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserMaquinaInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco txtNumero
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserMaquinaInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objMaquina             As busSisMaq.clsMaquina
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 6090
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  lngTIPOANTERIORID = 0
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  txtSerie.Text = strDescSerie
  'Tipo
  strSql = "Select TIPO.TIPO FROM TIPO " & _
      " ORDER BY TIPO.TIPO"
  PreencheCombo cboTipo, strSql, False, True
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
    Set objMaquina = New busSisMaq.clsMaquina
    Set objRs = objMaquina.SelecionarEquipamentoPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      lngTIPOANTERIORID = objRs.Fields("TIPOID").Value & ""
      txtNumero.Text = objRs.Fields("NUMERO").Value & ""
      cboTipo.Text = objRs.Fields("DESC_TIPO").Value & ""
      INCLUIR_VALOR_NO_MASK mskCoeficiente, objRs.Fields("COEFICIENTE").Value, TpMaskMoeda
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
    Set objMaquina = Nothing
    'Visible
    optStatus(0).Visible = True
    optStatus(1).Visible = True
    Label5(5).Visible = True
    '
    tabDetalhes.TabEnabled(0) = True
    If gsNivel <> gsAdmin And gsNivel <> gsDiretor Then
      tabDetalhes.TabEnabled(1) = False
    Else
      tabDetalhes.TabEnabled(1) = True
    End If
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


Private Sub mskCoeficiente_GotFocus()
  Seleciona_Conteudo_Controle mskCoeficiente
End Sub
Private Sub mskCoeficiente_LostFocus()
  Pintar_Controle mskCoeficiente, tpCorContr_Normal
End Sub


Private Sub txtNumero_GotFocus()
  Seleciona_Conteudo_Controle txtNumero
End Sub
Private Sub txtNumero_LostFocus()
  Pintar_Controle txtNumero, tpCorContr_Normal
End Sub


Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    grdMaquina.Enabled = False
    pictrava(0).Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    '
    SetarFoco txtNumero
  Case 1
    'Máquina
    grdMaquina.Enabled = True
    pictrava(0).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    'Montar RecordSet
    MAQ_COLUNASMATRIZ = grdMaquina.Columns.Count
    MAQ_LINHASMATRIZ = 0
    MAQ_MontaMatriz
    grdMaquina.Bookmark = Null
    grdMaquina.ReBind
    grdMaquina.ApproxCount = MAQ_LINHASMATRIZ
    '
    SetarFoco grdMaquina
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "SisMaq.frmUserMaquinaInc.tabDetalhes"
  AmpN
End Sub


Private Sub grdMaquina_UnboundReadDataEx( _
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
               Offset + intI, MAQ_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, MAQ_COLUNASMATRIZ, MAQ_LINHASMATRIZ, MAQ_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, MAQ_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserMaquinaInc.grdGeral_UnboundReadDataEx]"
End Sub

Public Sub MAQ_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMaq.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMaq.clsGeral
  '
  strSql = "SELECT EQUIPAMENTO.PKID, EQUIPAMENTO.NUMERO, MAQUINA.USUARIO, MAQUINA.INICIO, MAQUINA.TERMINO, EQUIPAMENTO.COEFICIENTE, TIPO.TIPO " & _
          "FROM EQUIPAMENTO INNER JOIN MAQUINA ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
          "         AND MAQUINA.STATUS = " & Formata_Dados("I", tpDados_Texto) & _
          " INNER JOIN TIPO ON TIPO.PKID = MAQUINA.TIPOID " & _
          "WHERE EQUIPAMENTO.PKID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
          " ORDER BY EQUIPAMENTO.NUMERO"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    MAQ_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim MAQ_Matriz(0 To MAQ_COLUNASMATRIZ - 1, 0 To MAQ_LINHASMATRIZ - 1)
  Else
    ReDim MAQ_Matriz(0 To MAQ_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To MAQ_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To MAQ_COLUNASMATRIZ - 1  'varre as colunas
          MAQ_Matriz(intJ, intI) = objRs(intJ) & ""
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
