VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmUserSalaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de sala"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4800
      Left            =   8430
      ScaleHeight     =   4800
      ScaleWidth      =   1860
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4575
         Left            =   90
         ScaleHeight     =   4515
         ScaleWidth      =   1605
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   120
         Width           =   1665
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2700
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3540
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1830
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   90
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4575
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   8070
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userSalaInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Grade de atendimento"
      TabPicture(1)   =   "userSalaInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdAtendimento"
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
         Height          =   3795
         Left            =   120
         TabIndex        =   14
         Top             =   360
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   3555
            Index           =   0
            Left            =   120
            ScaleHeight     =   3555
            ScaleWidth      =   7575
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1290
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   21
               TabStop         =   0   'False
               Top             =   1350
               Width           =   2235
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
               Begin VB.OptionButton optStatus 
                  Caption         =   "Inativo"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   5
                  Top             =   0
                  Width           =   1095
               End
            End
            Begin VB.TextBox txtTelefone 
               Height          =   285
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   3
               Text            =   "txtTelefone"
               Top             =   1020
               Width           =   2175
            End
            Begin VB.TextBox txtAndar 
               Height          =   285
               Left            =   1320
               MaxLength       =   10
               TabIndex        =   2
               Text            =   "txtAndar"
               Top             =   720
               Width           =   2175
            End
            Begin VB.ComboBox cboPredio 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   390
               Width           =   6105
            End
            Begin VB.TextBox txtNumero 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   0
               Text            =   "txtNumero"
               Top             =   75
               Width           =   2175
            End
            Begin VB.Label Label5 
               Caption         =   "Status"
               Height          =   315
               Index           =   5
               Left            =   60
               TabIndex        =   22
               Top             =   1380
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Telefone"
               Height          =   195
               Index           =   29
               Left            =   60
               TabIndex        =   19
               Top             =   1020
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Andar"
               Height          =   195
               Index           =   27
               Left            =   60
               TabIndex        =   18
               Top             =   690
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Prédio"
               Height          =   195
               Index           =   24
               Left            =   60
               TabIndex        =   17
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Número"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   16
               Top             =   120
               Width           =   1215
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdAtendimento 
         Height          =   3945
         Left            =   -74850
         OleObjectBlob   =   "userSalaInc.frx":0038
         TabIndex        =   6
         Top             =   420
         Width           =   7965
      End
   End
End
Attribute VB_Name = "frmUserSalaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Public lngPKID                  As Long

Private blnPrimeiraVez          As Boolean

Dim ATEND_COLUNASMATRIZ         As Long
Dim ATEND_LINHASMATRIZ          As Long
Private ATEND_Matriz()          As String

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Sala
  LimparCampoTexto txtNumero
  LimparCampoTexto txtAndar
  LimparCampoTexto txtTelefone
  LimparCampoCombo cboPredio
  optStatus(0).Value = False
  optStatus(1).Value = False
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserSalaInc.LimparCampos]", _
            Err.Description
End Sub


Private Sub cboPredio_LostFocus()
  Pintar_Controle cboPredio, tpCorContr_Normal
End Sub

Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 1
    If Not IsNumeric(grdAtendimento.Columns("PKID").Value & "") Then
      MsgBox "Selecione um atendimento !", vbExclamation, TITULOSISTEMA
      SetarFoco grdAtendimento
      Exit Sub
    End If

    frmUserAtendeInc.lngPKID = grdAtendimento.Columns("PKID").Value
    frmUserAtendeInc.lngSALAID = lngPKID
    frmUserAtendeInc.strDescrSala = cboPredio.Text & " - " & txtNumero.Text
    frmUserAtendeInc.Status = tpStatus_Alterar
    frmUserAtendeInc.Show vbModal

    If frmUserAtendeInc.blnRetorno Then
      CarregaHistoricoReceita
      grdAtendimento.Bookmark = Null
      grdAtendimento.ReBind
      grdAtendimento.ApproxCount = ATEND_LINHASMATRIZ
    End If
    SetarFoco grdAtendimento
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
  Dim objAtende     As busSisMed.clsAtende
  '
  Select Case tabDetalhes.Tab
  Case 1 'Exclusão de grade de atendimento
    '
    If Len(Trim(grdAtendimento.Columns("PKID").Value & "")) = 0 Then
      MsgBox "Selecione um atendimento.", vbExclamation, TITULOSISTEMA
      SetarFoco grdAtendimento
      Exit Sub
    End If
    '
    Set objAtende = New busSisMed.clsAtende
    '
    If MsgBox("Confirma exclusão do ítem da grade de atendimento " & grdAtendimento.Columns("Dia da Semana").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdAtendimento
      Exit Sub
    End If
    'OK
    objAtende.ExcluirAtende CLng(grdAtendimento.Columns("PKID").Value)
    '
    CarregaHistoricoReceita
    grdAtendimento.Bookmark = Null
    grdAtendimento.ReBind
    grdAtendimento.ApproxCount = ATEND_LINHASMATRIZ

    Set objAtende = Nothing
    SetarFoco grdAtendimento
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
    frmUserAtendeInc.Status = tpStatus_Incluir
    frmUserAtendeInc.lngPKID = 0
    frmUserAtendeInc.lngSALAID = lngPKID
    frmUserAtendeInc.strDescrSala = cboPredio.Text & " - " & txtNumero.Text
    frmUserAtendeInc.Show vbModal

    If frmUserAtendeInc.blnRetorno Then
      CarregaHistoricoReceita
      grdAtendimento.Bookmark = Null
      grdAtendimento.ReBind
      grdAtendimento.ApproxCount = ATEND_LINHASMATRIZ
    End If
    SetarFoco grdAtendimento
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdOk_Click()
  Dim objSala               As busSisMed.clsSala
  Dim objGeral                  As busSisMed.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngPREDIOID         As Long
  Dim strStatus                 As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisMed.clsGeral
  Set objSala = New busSisMed.clsSala
  'PRÉDIO
  lngPREDIOID = 0
  strSql = "SELECT PKID FROM PREDIO WHERE NOME = " & Formata_Dados(cboPredio.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngPREDIOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Status
  If optStatus(0).Value Then
    strStatus = "A"
  Else
    strStatus = "I"
  End If
  'Validar se sala já cadastrada
  strSql = "SELECT * FROM SALA " & _
    " WHERE SALA.NUMERO = " & Formata_Dados(txtNumero.Text, tpDados_Texto) & _
    " AND SALA.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtNumero, tpCorContr_Erro
    TratarErroPrevisto "Sala já cadastrada"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objSala = Nothing
    cmdOk.Enabled = True
    SetarFoco txtNumero
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Sala
    objSala.AlterarSala lngPKID, _
                        lngPREDIOID, _
                        txtNumero.Text, _
                        txtAndar.Text, _
                        txtTelefone.Text, _
                        strStatus
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Sala
    objSala.InserirSala lngPREDIOID, _
                        txtNumero.Text, _
                        txtAndar.Text, _
                        txtTelefone.Text, _
                        strStatus
    blnRetorno = True
    'Selecionar plano cadastrado
    Set objRs = objSala.SelecionarSalaPeloNumero(txtNumero.Text)
    If Not objRs.EOF Then
      'Captura dados para entrar em modo de alteração
      lngPKID = objRs.Fields("PKID")
      Status = tpStatus_Alterar
      'Reload na tela
      Form_Load
      'Acerta tabs
      tabDetalhes.Tab = 1
      blnRetorno = True
    Else
      blnRetorno = True
      blnFechar = True
      Unload Me
    End If
  End If
  Set objSala = Nothing
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
    strMsg = strMsg & "Preencher o número" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(cboPredio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o prédio" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Option(optStatus, blnSetarFocoControle) Then
    strMsg = strMsg & "Slecionar o status" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserSalaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserSalaInc.ValidaCampos]", _
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
  TratarErro Err.Number, Err.Description, "[frmUserSalaInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objSala             As busSisMed.clsSala
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 5280
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , cmdIncluir, cmdAlterar
  '
  'Limpar Campos
  LimparCampos
  'Tipo de Convênio
  strSql = "Select NOME from PREDIO ORDER BY NOME"
  PreencheCombo cboPredio, strSql, False, True, True
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
    Set objSala = New busSisMed.clsSala
    Set objRs = objSala.SelecionarSalaPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      txtNumero.Text = objRs.Fields("NUMERO").Value & ""
      txtAndar.Text = objRs.Fields("ANDAR").Value & ""
      txtTelefone.Text = objRs.Fields("TELEFONE").Value & ""
      cboPredio.Text = objRs.Fields("NOME_PREDIO").Value & ""
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
    Set objSala = Nothing
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = True
    'Visible
    optStatus(0).Visible = True
    optStatus(1).Visible = True
    Label5(5).Visible = True
    '
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


Private Sub txtAndar_GotFocus()
  Seleciona_Conteudo_Controle txtAndar
End Sub
Private Sub txtAndar_LostFocus()
  Pintar_Controle txtAndar, tpCorContr_Normal
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
    grdAtendimento.Enabled = False
    pictrava(0).Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco txtNumero
  Case 1
    grdAtendimento.Enabled = True
    pictrava(0).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = True
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    'Montar RecordSet
    ATEND_COLUNASMATRIZ = grdAtendimento.Columns.Count
    ATEND_LINHASMATRIZ = 0
    CarregaHistoricoReceita
    grdAtendimento.Bookmark = Null
    grdAtendimento.ReBind
    grdAtendimento.ApproxCount = ATEND_LINHASMATRIZ
    '
    SetarFoco grdAtendimento
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "Apler.frmUserSalaInc.tabDetalhes"
  AmpN
End Sub


Private Sub grdAtendimento_UnboundReadDataEx( _
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
               Offset + intI, ATEND_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ATEND_COLUNASMATRIZ, ATEND_LINHASMATRIZ, ATEND_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ATEND_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserPlanoInc.grdGeral_UnboundReadDataEx]"
End Sub

Public Sub CarregaHistoricoReceita()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMed.clsGeral
  '
  strSql = "SELECT ATENDE.PKID, DIASDASEMANA.DIADASEMANA, ATENDE.HORAINICIO, ATENDE.HORATERMINO, PRONTUARIO.NOME, CASE ATENDE.STATUS WHEN 'A' THEN 'Ativo' ELSE 'Inativo' END  " & _
          "FROM ATENDE " & _
          " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = ATENDE.PRONTUARIOID " & _
          " INNER JOIN DIASDASEMANA ON DIASDASEMANA.PKID = ATENDE.DIASDASEMANAID " & _
          "WHERE ATENDE.SALAID = " & lngPKID & _
          " ORDER BY DIASDASEMANA.CODIGO, ATENDE.HORAINICIO, ATENDE.HORATERMINO, PRONTUARIO.NOME  "

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    ATEND_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim ATEND_Matriz(0 To ATEND_COLUNASMATRIZ - 1, 0 To ATEND_LINHASMATRIZ - 1)
  Else
    ReDim ATEND_Matriz(0 To ATEND_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To ATEND_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To ATEND_COLUNASMATRIZ - 1  'varre as colunas
          ATEND_Matriz(intJ, intI) = objRs(intJ) & ""
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

Private Sub txtTelefone_GotFocus()
  Seleciona_Conteudo_Controle txtTelefone
End Sub
Private Sub txtTelefone_LostFocus()
  Pintar_Controle txtTelefone, tpCorContr_Normal
End Sub
