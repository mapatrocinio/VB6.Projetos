VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmUserTurnoInc 
   Caption         =   "Abertura/Fechamento de Turno"
   ClientHeight    =   5415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9675
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   9675
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   0
      ScaleHeight     =   1095
      ScaleWidth      =   9675
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   4320
      Width           =   9675
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   880
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   90
         Width           =   1215
      End
      Begin VB.CommandButton cmdConfirmar 
         Caption         =   "ENTER"
         Default         =   -1  'True
         Height          =   880
         Left            =   3810
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   90
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4110
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   9375
      _ExtentX        =   16536
      _ExtentY        =   7250
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   529
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Cadastro"
      TabPicture(0)   =   "userTurnoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAluno"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Reimpressão de Turnos"
      TabPicture(1)   =   "userTurnoInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label3"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "grdGeral"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "cmdImprimir"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Picture4"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "chkExibeOutros"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).ControlCount=   6
      Begin VB.CheckBox chkExibeOutros 
         Caption         =   "Exibir turnos de outros funcionários"
         Height          =   195
         Left            =   -74880
         TabIndex        =   7
         Top             =   3390
         Width           =   2865
      End
      Begin VB.PictureBox Picture4 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   -69960
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   3600
         Width           =   495
      End
      Begin VB.CommandButton cmdImprimir 
         Caption         =   "&Z"
         Height          =   880
         Left            =   -66990
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   3120
         Width           =   1215
      End
      Begin VB.Frame fraAluno 
         Height          =   3435
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   6855
         Begin VB.PictureBox Picture3 
            AutoSize        =   -1  'True
            BorderStyle     =   0  'None
            Height          =   495
            Left            =   4320
            ScaleHeight     =   495
            ScaleWidth      =   495
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   2880
            Width           =   495
         End
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   2415
            Index           =   0
            Left            =   120
            ScaleHeight     =   2415
            ScaleWidth      =   6615
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   6615
            Begin VB.TextBox txtTurno 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   525
               Left            =   1440
               Locked          =   -1  'True
               MultiLine       =   -1  'True
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "userTurnoInc.frx":0038
               Top             =   0
               Width           =   4935
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Informações do novo turno"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1695
               Left            =   120
               TabIndex        =   13
               Top             =   600
               Width           =   6255
               Begin VB.ComboBox cboDiaDaSemana 
                  DataSource      =   "dcPeriodo"
                  Height          =   315
                  Left            =   1560
                  Style           =   2  'Dropdown List
                  TabIndex        =   3
                  Top             =   960
                  Width           =   4335
               End
               Begin VB.ComboBox cboPeriodo 
                  DataSource      =   "dcPeriodo"
                  Height          =   315
                  Left            =   1560
                  Style           =   2  'Dropdown List
                  TabIndex        =   2
                  Top             =   600
                  Width           =   4335
               End
               Begin VB.PictureBox Picture2 
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   120
                  ScaleHeight     =   255
                  ScaleWidth      =   3255
                  TabIndex        =   16
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   3255
                  Begin MSMask.MaskEdBox mskData 
                     Height          =   255
                     Index           =   0
                     Left            =   1440
                     TabIndex        =   1
                     TabStop         =   0   'False
                     Top             =   0
                     Width           =   1695
                     _ExtentX        =   2990
                     _ExtentY        =   450
                     _Version        =   393216
                     BackColor       =   14737632
                     AutoTab         =   -1  'True
                     MaxLength       =   16
                     Mask            =   "##/##/#### ##:##"
                     PromptChar      =   "_"
                  End
                  Begin VB.Label Da 
                     Caption         =   "Dt./Hr.  Turno"
                     Height          =   255
                     Left            =   0
                     TabIndex        =   17
                     Top             =   0
                     Width           =   1335
                  End
               End
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
                  TabIndex        =   14
                  Top             =   3480
                  Width           =   2295
               End
               Begin VB.Label Label1 
                  Caption         =   "Dia da Semana"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   19
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.Label Label8 
                  Caption         =   "Período"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   15
                  Top             =   600
                  Width           =   735
               End
            End
            Begin VB.Label Label21 
               Caption         =   "Turno Corrente"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   0
               TabIndex        =   18
               Top             =   0
               Width           =   1455
            End
         End
         Begin VB.Label lblImpressao 
            Caption         =   "Aguarde o fim da Impressão. .  ."
            Height          =   255
            Left            =   1560
            TabIndex        =   20
            Top             =   3000
            Width           =   2295
         End
      End
      Begin TrueDBGrid60.TDBGrid grdGeral 
         Height          =   2595
         Left            =   -74880
         OleObjectBlob   =   "userTurnoInc.frx":0041
         TabIndex        =   6
         Top             =   480
         Width           =   9135
      End
      Begin VB.Label Label3 
         Caption         =   "Exibe últimos 500 turnos"
         ForeColor       =   &H000000FF&
         Height          =   285
         Left            =   -74880
         TabIndex        =   24
         Top             =   3150
         Width           =   1875
      End
      Begin VB.Label Label2 
         Caption         =   "Aguarde o fim da Impressão. .  ."
         Height          =   255
         Left            =   -69690
         TabIndex        =   23
         Top             =   3690
         Width           =   2295
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   9300
      Top             =   4050
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmUserTurnoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim GER_COLUNASMATRIZ         As Long
Dim GER_LINHASMATRIZ          As Long
Private GER_Matriz()          As String

Option Explicit

Private Sub cboDiaDaSemana_LostFocus()
  Pintar_Controle cboDiaDaSemana, tpCorContr_Normal
End Sub

Private Sub cboPeriodo_LostFocus()
  Pintar_Controle cboPeriodo, tpCorContr_Normal
End Sub

Private Sub chkExibeOutros_Click()
  On Error GoTo trata
  tabDetalhes_Click 0
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub cmdCancelar_Click()
  On Error GoTo trata
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

Private Sub cmdConfirmar_Click()
  On Error GoTo trata
  Dim lngTURNOANTERIORID  As Long
  Dim strMsg              As String
  Dim objTurno            As busSisMaq.clsTurno
  Dim objGeral            As busSisMaq.clsGeral
  Dim objEntrada          As busSisMaq.clsEntrada
  Dim strSql              As String
  Dim curVrCaixa          As Currency
  '
  gsNomeUsuLib = ""
  If Not ValidaCampos Then Exit Sub
  If RetornaCodTurnoCorrente <> 0 Then
    TratarErroPrevisto "Há turno aberto. Favor fechá-lo", "frmUserTurnoInc.cmdFechar_Click"
    Exit Sub
  End If
  
  'Data1.Recordset!StatusOp = iStatusOp
  If MsgBox("Confirma abertura do Turno?", vbYesNo, TITULOSISTEMA) = vbNo Then Exit Sub
  'Encerrar Turno
  '
'''  If gsNomeUsuLib = "" Then
'''    If gbPedirSenhaFechaTurno = True Then
'''      '----------------------------
'''      '----------------------------
'''      'Pede Senha Superior (Diretor, Gerente ou Administrador
'''      If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''        'Só pede senha superior se quem estiver logado não for superior
'''        frmUserLoginSup.Show vbModal
'''
'''        If Len(Trim(gsNomeUsuLib)) = 0 Then
'''          strMsg = "Para efetuar o Fechamento/Abertura do Turno é necessário a Confirmação com senha superior."
'''          TratarErroPrevisto strMsg, "cmdConfirmar_Click"
'''          Exit Sub
'''        End If
'''        '
'''        'Capturou Nome do Usuário, continua processo de Sangria
'''      Else
'''        gsNomeUsuLib = gsNomeUsu
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''    End If
'''  End If
  'Tratamento dos botões
  tabDetalhes.TabEnabled(1) = False
  cmdConfirmar.Enabled = False
  cmdCancelar.Enabled = False
  '
  Picture3.Visible = True
  lblImpressao.Visible = True
  '
  DoEvents
  '
  '------------------------------------
  'ABERTURA
  '------------------------------------
  Set objTurno = New busSisMaq.clsTurno
  objTurno.InserirTurno giFuncionarioId, _
                        cboPeriodo.ItemData(cboPeriodo.ListIndex), _
                        mskData(0).Text, _
                        True, _
                        cboDiaDaSemana.ItemData(cboDiaDaSemana.ListIndex)
                        
  '
'''  'Verifica lançamento de valor inicial do caixa
'''  If IsNumeric(mskVrCaixa.Text) Then
'''    curVrCaixa = CCur(mskVrCaixa.ClipText)
'''    If curVrCaixa <> 0 Then
'''      'Entrou com valor do caixa, insere valor inicial
'''      Set objEntrada = New busSisMaq.clsEntrada
'''      objEntrada.InserirEntrada RetornaCodTurnoCorrente, _
'''                                mskVrCaixa.ClipText, _
'''                                "T", _
'''                                Format(Now, "DD/MM/YYYY hh:mm")
'''      Set objEntrada = Nothing
'''    End If
'''  End If
  '------------
  'INCLUIR LOG
  'INCLUI_LOG_UNIDADE MODOINSERIR, 0, "Abertura de Turno", RetornaDescTurnoCorrente, "", "", "", gsNomeUsuLib
  'Rotina para Impressão do Fechamento do turno
  '------------
  '
  '
  'IMP_COMPROV_FECHA_TURNO lngTURNOANTERIORID, gsNomeEmpresa, 1
  '
  MsgBox "O turno foi aberto com sucesso!", vbExclamation, TITULOSISTEMA
  '
  Set objTurno = Nothing
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Private Sub cmdImprimir_Click()
  On Error GoTo trata
  Dim lngTURNOANTERIORID  As Long
  Dim strMsg              As String
  Dim strSql              As String
  Dim objGeral            As busSisMaq.clsGeral
  '
  If Not IsNumeric(grdGeral.Columns("PKID").Value) Then
    MsgBox "Selecione um turno da lista.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGeral
    Exit Sub
  End If
  'Data1.Recordset!StatusOp = iStatusOp
  If MsgBox("Confirma a impressão do turno " & grdGeral.Columns("Data").Value & " - " & grdGeral.Columns("Prestador").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdGeral
    Exit Sub
  End If
  'Encerrar Turno
  '
  'trata botões
  tabDetalhes.TabEnabled(0) = False
  Picture4.Visible = True
  Label2.Visible = True
  '
  DoEvents
  '
  lngTURNOANTERIORID = grdGeral.Columns("PKID").Value
  '
  'IMP_COMPROV_FECHA_TURNO lngTURNOANTERIORID, gsNomeEmpresa, 1
  '
  Report1.Connect = ConnectRpt
  Report1.ReportFileName = gsReportPath & "FechaTurno.rpt"
  '
  Report1.Destination = 0 'Video
  Report1.CopiesToPrinter = 1
  Report1.WindowState = crptMaximized
  '
  Report1.Formulas(0) = "TurnoId = " & lngTURNOANTERIORID
  '
  Report1.Action = 1
  '
  MsgBox "Impressão efetuada com sucesso !", vbExclamation, TITULOSISTEMA
  'trata botões
  tabDetalhes.TabEnabled(0) = True
  Picture4.Visible = False
  Label2.Visible = False
  '
  Set objGeral = Nothing
  '
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
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
  strSql = "Select TOP 500 TURNO.PKID, TURNO.DATA, DIASDASEMANA.DIADASEMANA, CONVERT(VARCHAR, PERIODO.PERIODO) + ' - ' + CONVERT(VARCHAR, PERIODO.inicio) + ' às ' + CONVERT(VARCHAR, PERIODO.termino) AS PERIODO, FUNCIONARIO.USUARIO "
  strSql = strSql & " FROM PERIODO INNER JOIN TURNO ON PERIODO.PKID = TURNO.PERIODOID " & _
      " INNER JOIN DIASDASEMANA ON TURNO.DIASDASEMANAID = DIASDASEMANA.PKID " & _
      " INNER JOIN FUNCIONARIO ON FUNCIONARIO.PESSOAID = TURNO.CAIXAID "
  'strSql = strSql & " WHERE TURNO.STATUS <> " & Formata_Dados(True, tpDados_Boolean) & " "
  If chkExibeOutros.Value = False Then
    strSql = strSql & " AND TURNO.CAIXAID = " & Formata_Dados(giFuncionarioId, tpDados_Longo)
  End If
  strSql = strSql & " ORDER BY DATA DESC;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    GER_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim GER_Matriz(0 To GER_COLUNASMATRIZ - 1, 0 To GER_LINHASMATRIZ - 1)
  Else
    ReDim GER_Matriz(0 To GER_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To GER_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To GER_COLUNASMATRIZ - 1  'varre as colunas
          GER_Matriz(intJ, intI) = objRs(intJ) & ""
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
  AmpS
  Me.Width = 9795
  Me.Height = 5820
  CenterForm Me
  '
  Picture3.Visible = False
  lblImpressao.Visible = False
  '
  Picture4.Visible = False
  Label2.Visible = False
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar, , , , , , cmdImprimir
  LerFigurasAvulsasPicBox Picture3, "impressora.ico", "Aguarde o fim da impressão.. .  ."
  LerFigurasAvulsasPicBox Picture4, "impressora.ico", "Aguarde o fim da impressão.. .  ."
  LerFigurasAvulsas cmdConfirmar, "Ok.ico", "OkDown.ico", "Abrir turno"
  
  '
  CarregaPeriodo
  CarregaDiaDaSemana
  '
  txtTurno.Text = RetornaDescTurnoCorrente
  mskData(0).Text = DataHoraAtualFormatada
  SetarDiaDaSemana cboDiaDaSemana, _
                     Now
  
  '
  'mskVrCaixa.Text = IIf(Not IsNumeric(gcVrIniCaixa), 0, Format(gcVrIniCaixa, "###,##0.00"))
  'mskVrCaixa.Text = IIf(Not IsNumeric(0), 0, Format(0, "###,##0.00"))
  'mskVrCaixa.Enabled = True
  'Label6.Enabled = True
  '
  tabDetalhes_Click 0
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub


Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim Msg As String
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  Dim objGer As busSisMaq.clsGeral
  '
  '
  If Len(cboPeriodo.Text) = 0 Then
    Msg = Msg & "Selecionar um período" & vbCrLf
    Pintar_Controle cboPeriodo, tpCorContr_Erro
  End If
  '
  If Len(cboDiaDaSemana.Text) = 0 Then
    Msg = Msg & "Selecionar um dia da semana" & vbCrLf
    Pintar_Controle cboDiaDaSemana, tpCorContr_Erro
  End If
  '
  Set objGer = New busSisMaq.clsGeral
  '
  If Len(Msg) = 0 Then
    strSql = "SELECT * " & _
      "FROM TURNO " & _
      "WHERE TURNO.PERIODOID = " & cboPeriodo.ItemData(cboPeriodo.ListIndex) & _
      " AND TURNO.DATA >= " & Formata_Dados(Left(mskData(0).Text, 10) & " 00:00", tpDados_DataHora, tpNulo_Aceita) & _
      " AND TURNO.DATA <= " & Formata_Dados(Left(mskData(0).Text, 10) & " 23:59", tpDados_DataHora, tpNulo_Aceita) '& _
      '" AND TURNO.GERENTEID = " & Formata_Dados(giFuncionarioId, tpDados_Longo)

    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      If MsgBox("Este turno já foi aberto hoje, tem certeza de que deseja continuar?", vbYesNo, TITULOSISTEMA) = vbNo Then
        Msg = "Este turno já foi aberto hoje"
      End If
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  'NOVO obrigar a lançar leitura especial final para o turno
  If Len(Msg) = 0 Then
    'Verifica se FOI LANÇADO LEITURA ESPECIAL FINAL

    strSql = "SELECT * " & _
      "FROM LEITURAFECHA " & _
      " INNER JOIN PERIODO ON PERIODO.PKID = LEITURAFECHA.PERIODOID " & _
      "WHERE LEITURAFECHA.DATA >= " & Formata_Dados(Left(mskData(0).Text, 10) & " 00:00", tpDados_DataHora, tpNulo_Aceita) & _
      " AND LEITURAFECHA.DATA <= " & Formata_Dados(Left(mskData(0).Text, 10) & " 23:59", tpDados_DataHora, tpNulo_Aceita) & _
      " AND PERIODO.PERIODO = " & Formata_Dados(cboPeriodo.ItemData(cboPeriodo.ListIndex), tpDados_Longo) & _
      " AND LEITURAFECHA.STATUS = " & Formata_Dados("I", tpDados_Texto)
    Set objRs = objGer.ExecutarSQL(strSql)
    If objRs.EOF Then
      Msg = "É necessário entrar com a leitura inicial antes de abrir o turno."
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  
  Set objGer = Nothing
  If Len(Msg) <> 0 Then
    TratarErroPrevisto Msg, "ValidaCampos"
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

'Propósito: Carregar combos com dias da semanas
Public Sub CarregaDiaDaSemana()
  On Error GoTo trata
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objGeral  As busSisMaq.clsGeral
  Dim lngIndice As Long
  '
  Set objGeral = New busSisMaq.clsGeral
  '
  strSql = "Select * From DIASDASEMANA ORDER BY DIASDASEMANA.CODIGO"
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  lngIndice = 0
  Do While Not objRs.EOF
    cboDiaDaSemana.AddItem objRs.Fields("DiaDaSemana").Value & ""
    cboDiaDaSemana.ItemData(lngIndice) = objRs.Fields("PKID").Value
    objRs.MoveNext
    lngIndice = lngIndice + 1
  Loop
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserTurnoInc.CarregaDiaDaSemana]", _
            Err.Description
End Sub


'Propósito: Carregar combos com período
Public Sub CarregaPeriodo()
  On Error GoTo trata
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objGeral  As busSisMaq.clsGeral
  Dim lngIndice As Long
  '
  Set objGeral = New busSisMaq.clsGeral
  '
  strSql = "Select * From Periodo Order By Periodo"
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  lngIndice = 0
  Do While Not objRs.EOF
    cboPeriodo.AddItem objRs.Fields("PERIODO").Value & " - " & objRs.Fields("Inicio").Value & " até " & objRs.Fields("Termino").Value
    cboPeriodo.ItemData(lngIndice) = objRs.Fields("PKID").Value
    objRs.MoveNext
    lngIndice = lngIndice + 1
  Loop
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserTurnoInc.CarregaPeriodo]", _
            Err.Description
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
               Offset + intI, GER_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, GER_COLUNASMATRIZ, GER_LINHASMATRIZ, GER_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, GER_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserTurnoInc.grdGeral_UnboundReadDataEx]"
End Sub


Private Sub tabDetalhes_Click(PreviousTab As Integer)
  On Error GoTo trata
  Dim sSql As String
  Select Case tabDetalhes.Tab
  Case 0
    fraAluno.Enabled = True
    grdGeral.Enabled = False
    cmdImprimir.Enabled = False
    '
    cmdConfirmar.Enabled = True
    SetarFoco cboPeriodo
  Case 1
    fraAluno.Enabled = False
    grdGeral.Enabled = True
    cmdImprimir.Enabled = True
    '
    cmdConfirmar.Enabled = False
    'Montar RecordSet
    GER_COLUNASMATRIZ = grdGeral.Columns.Count
    GER_LINHASMATRIZ = 0
    MontaMatriz
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = GER_LINHASMATRIZ
    SetarFoco grdGeral
  
    SetarFoco grdGeral
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub



'Proposito: Retornar a descrição do Turno Corrente e emitir msg de erro
'para usuário
Public Function RetornaDescTurnoCorrente(Optional TURNOID As Long) As String
  On Error GoTo trata
  'Retorna "", caso não encontre
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim strRetorno    As String
  Dim strErro       As String
  Dim objGeral      As busSisMaq.clsGeral
  Dim strDescrTurno As String
  '
  Set objGeral = New busSisMaq.clsGeral
  '
  strSql = "Select Data, DIASDASEMANA.DIADASEMANA, Periodo, inicio, termino " & _
    "FROM (PERIODO INNER JOIN TURNO ON PERIODO.PKID = TURNO.PERIODOID) INNER JOIN DIASDASEMANA ON TURNO.DIASDASEMANAID = DIASDASEMANA.PKID " & _
    "WHERE " & IIf(TURNOID <> 0, "TURNO.PKID = " & Formata_Dados(TURNOID, tpDados_Longo) & ";", "Status = " & Formata_Dados(True, tpDados_Boolean) & " AND TURNO.CAIXAID = " & Formata_Dados(giFuncionarioId, tpDados_Longo))

  'ASSUME 0 - TODOS OS DIAS / 1-  FIM DE SEMANA / 2 - FERIADO / 3 - DIAS DE SEMANA / 4 - SEGUNDA / 5 - TERÇA  / 6 - QUARTA / 7 - QUINTA  / 8 - SEXTA / 9 - SÁBADO  / 10 - DOMINGO / 11 - ESPECIAL
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If objRs.EOF Then
    'Não há turno corrente cadastrado
    strRetorno = ""
    strErro = "Não há Turno aberto no Sistema"
  ElseIf objRs.RecordCount > 1 Then
    'há mais de um turno corrente cadastrado
    strRetorno = ""
    '
    strErro = "Há mais de um turno aberto no sistema:" & vbCrLf & vbCrLf
    Do While Not objRs.EOF
      strDescrTurno = Format(objRs.Fields("Data").Value, "DD/MM/YYYY") & " / " & objRs.Fields("DIADASEMANA").Value & " - Período " & objRs.Fields("Periodo").Value & " de " & objRs.Fields("inicio").Value & " as " & objRs.Fields("termino").Value
      strRetorno = strDescrTurno & vbCrLf
      objRs.MoveNext
    Loop
  Else
    strDescrTurno = Format(objRs.Fields("Data").Value, "DD/MM/YYYY") & " / " & objRs.Fields("DIADASEMANA").Value & " - Período " & objRs.Fields("Periodo").Value & " de " & objRs.Fields("inicio").Value & " as " & objRs.Fields("termino").Value
    strRetorno = strDescrTurno
    strErro = ""
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  '
  RetornaDescTurnoCorrente = strRetorno
  'Emite Msg de Erro
  'If Len(strErro) <> 0 Then Err.Raise 1, , strErro
  Exit Function
trata:
  Err.Raise Err.Number, _
            "[mdlUserFunction.RetornaDescTurnoCorrente]", _
            Err.Description
End Function


