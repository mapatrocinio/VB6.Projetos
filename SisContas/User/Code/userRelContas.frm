VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmUserRelContas 
   Caption         =   "Demonstrativo de Contas"
   ClientHeight    =   4665
   ClientLeft      =   450
   ClientTop       =   1920
   ClientWidth     =   6105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4665
   ScaleWidth      =   6105
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   5775
      Begin VB.OptionButton optSai2 
         Caption         =   "Impressora"
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
         Left            =   120
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optSai1 
         Caption         =   "Vídeo"
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
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   0
      ScaleHeight     =   855
      ScaleWidth      =   6105
      TabIndex        =   14
      Top             =   3810
      Width           =   6105
      Begin VB.CommandButton cmdRelatorio 
         Default         =   -1  'True
         Height          =   735
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   30
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Height          =   735
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   30
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   120
      TabIndex        =   12
      Top             =   1200
      Width           =   5775
      Begin VB.ComboBox cboTipo 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   1800
         Width           =   3135
      End
      Begin VB.ComboBox cboOperador 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1440
         Width           =   735
      End
      Begin VB.ComboBox cboSubGrupo 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1080
         Width           =   3135
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   3135
      End
      Begin VB.OptionButton optConta 
         Caption         =   "Contas a Pagar"
         Height          =   195
         Index           =   1
         Left            =   1920
         TabIndex        =   9
         Top             =   2160
         Width           =   1695
      End
      Begin VB.OptionButton optConta 
         Caption         =   "Contas Pagas"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Top             =   2160
         Value           =   -1  'True
         Width           =   1695
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskValor 
         Height          =   255
         Left            =   2160
         TabIndex        =   7
         Top             =   1440
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "Tipo"
         Height          =   255
         Left            =   240
         TabIndex        =   21
         Top             =   1800
         Width           =   735
      End
      Begin VB.Label Label3 
         Caption         =   "Sub-Grupo"
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label4 
         Caption         =   "Valor :"
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Grupo"
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Até"
         Height          =   255
         Left            =   2640
         TabIndex        =   15
         Tag             =   "lblIdCliente"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCliente 
         Caption         =   "Período : "
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   13
         Tag             =   "lblIdCliente"
         Top             =   360
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3360
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmUserRelContas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboGrupo_Click()
  Dim strSql As String
  On Error GoTo trata
  If cboGrupo.Text = "<TODOS>" Then
    strSql = "Select SUBGRUPODESPESA.CODIGO + ' - ' + SUBGRUPODESPESA.DESCRICAO " & _
      " FROM SUBGRUPODESPESA " & _
      " ORDER BY SUBGRUPODESPESA.CODIGO "
  Else
    strSql = "Select SUBGRUPODESPESA.CODIGO + ' - ' + SUBGRUPODESPESA.DESCRICAO " & _
      " FROM GRUPODESPESA INNER JOIN SUBGRUPODESPESA ON GRUPODESPESA.PKID = SUBGRUPODESPESA.GRUPODESPESAID" & _
      " WHERE GRUPODESPESA.CODIGO = " & Formata_Dados(Left(cboGrupo.Text, 4), tpDados_Texto, tpNulo_Aceita) & _
      " ORDER BY SUBGRUPODESPESA.CODIGO "
  End If
  PreencheCombo cboSubGrupo, strSql, True, False
  
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub


Private Sub cboOperador_LostFocus()
  Pintar_Controle cboOperador, tpCorContr_Normal
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdRelatorio_Click()
  On Error GoTo TratErro
  AmpS
  
  If Not Valida_Data(mskData(0), TpObrigatorio) Then
    AmpN
    TratarErroPrevisto "Data Inicial Inválida", "[cmdRelatorio_Click]"
    SetarFoco mskData(0)
    Pintar_Controle mskData(0), tpCorContr_Erro
    Exit Sub
  ElseIf Not Valida_Data(mskData(1), TpObrigatorio) Then
    AmpN
    TratarErroPrevisto "Data Final Inválida", "[cmdRelatorio_Click]"
    SetarFoco mskData(1)
    Pintar_Controle mskData(1), tpCorContr_Erro
    Exit Sub
  ElseIf Not Valida_Moeda(mskValor, TpNaoObrigatorio) Then
    AmpN
    TratarErroPrevisto "Valor Inválido", "[cmdRelatorio_Click]"
    SetarFoco mskValor
    Pintar_Controle mskValor, tpCorContr_Erro
    Exit Sub
  ElseIf Len(mskValor.ClipText) > 0 And cboOperador.Text = "" Then
    AmpN
    TratarErroPrevisto "Selecione um operador", "[cmdRelatorio_Click]"
    SetarFoco cboOperador
    Pintar_Controle cboOperador, tpCorContr_Erro
    Exit Sub
  End If

  '
  If optSai1.Value Then
    Report1.Destination = 0 'Video
  ElseIf optSai2.Value Then
    Report1.Destination = 1   'Impressora
  End If
  Report1.CopiesToPrinter = 1
  Report1.WindowState = crptMaximized
  '
  If optConta(0).Value Then
    'Contas Pagas
    Report1.ReportFileName = gsReportPath & "DemoContasPagas.rpt"
  Else
    'Contas a Pagar
    Report1.ReportFileName = gsReportPath & "DemoContasAPagar.rpt"
  End If
  
  Report1.Formulas(0) = "DataIni = Date(" & Right(mskData(0).Text, 4) & ", " & Mid(mskData(0).Text, 4, 2) & ", " & Left(mskData(0).Text, 2) & ")"
  Report1.Formulas(1) = "DataFim = Date(" & Right(mskData(1).Text, 4) & ", " & Mid(mskData(1).Text, 4, 2) & ", " & Left(mskData(1).Text, 2) & ")"
  If cboGrupo.Text = "" Or cboGrupo.Text = "<TODOS>" Then
    Report1.Formulas(2) = "CodigoGrupo = True = true"
  Else
    Report1.Formulas(2) = "CodigoGrupo = {GRUPODESPESA.CODIGO} = '" & Left(cboGrupo.Text, 4) & "'"
  End If
  If cboSubGrupo.Text = "" Or cboSubGrupo.Text = "<TODOS>" Then
    Report1.Formulas(3) = "CodigoSubGrupo = True = true"
  Else
    Report1.Formulas(3) = "CodigoSubGrupo = {SUBGRUPODESPESA.CODIGO} = '" & Left(cboSubGrupo.Text, 4) & "'"
  End If
  If Len(mskValor.ClipText) = 0 Then
    Report1.Formulas(4) = "Valor = True = true"
  Else
    If optConta(0).Value Then
      Report1.Formulas(4) = "Valor = {DESPESA.VR_PAGO} " & cboOperador.Text & " " & Formata_Dados(mskValor.Text, tpDados_Moeda, tpNulo_Aceita)
    Else
      Report1.Formulas(4) = "Valor = {DESPESA.VR_PAGAR} " & cboOperador.Text & " " & Formata_Dados(mskValor.Text, tpDados_Moeda, tpNulo_Aceita)
    End If
  End If
  If optConta(0).Value Then
    If cboTipo.Text = "" Or cboTipo.Text = "<TODOS>" Then
      Report1.Formulas(5) = "DespesaTipo = True = true"
    Else
      Report1.Formulas(5) = "DespesaTipo = {DESPESA.TIPO} = '" & IIf(cboTipo.Text = "Administração", "A", "T") & "'"
    End If
  End If
  Report1.Formulas(6) = "ParceiroId = " & glParceiroId
  Report1.Action = 1
  '
  AmpN
  Exit Sub
  
TratErro:
  AmpN
  TratarErro Err.Number, Err.Description, "[frmUserRelContas.cmdRelatorio_Click]"
End Sub

Private Sub Form_Activate()
  mskData(0).SetFocus
End Sub

Private Sub Form_Load()
  Dim strSql As String
  On Error GoTo trata
  AmpS
  CenterForm Me
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar, , , , , , cmdRelatorio
  '
  Report1.Connect = ConnectRpt
  '
  mskData(0).Text = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
  mskData(1).Text = Format(Now, "DD/MM/YYYY")
  '
  strSql = "Select GRUPODESPESA.CODIGO + ' - ' + GRUPODESPESA.DESCRICAO " & _
    " FROM GRUPODESPESA " & _
    " WHERE GRUPODESPESA.PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo) & _
    " AND GRUPODESPESA.TIPO = " & Formata_Dados("D", tpDados_Texto) & _
    " ORDER BY GRUPODESPESA.CODIGO "
  PreencheCombo cboGrupo, strSql, True, False
  strSql = "Select SUBGRUPODESPESA.CODIGO + ' - ' + SUBGRUPODESPESA.DESCRICAO " & _
    " FROM SUBGRUPODESPESA " & _
    " ORDER BY SUBGRUPODESPESA.CODIGO "
  PreencheCombo cboSubGrupo, strSql, True, False
  cboOperador.Clear
  cboOperador.AddItem ""
  cboOperador.AddItem "="
  cboOperador.AddItem ">"
  cboOperador.AddItem ">="
  cboOperador.AddItem "<"
  cboOperador.AddItem "<="
  cboOperador.AddItem "<>"
  cboTipo.Clear
  cboTipo.AddItem "<TODOS>"
  cboTipo.AddItem "Administração"
  cboTipo.AddItem "Telefonia"
  
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, "[frmUserRelContas.Form_Load]"
End Sub

Private Sub mskValor_GotFocus()
  Selecionar_Conteudo mskValor
End Sub

Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub

Private Sub mskData_GotFocus(Index As Integer)
  Selecionar_Conteudo mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub



