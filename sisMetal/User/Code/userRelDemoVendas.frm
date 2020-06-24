VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRelDemoVendas 
   Caption         =   "Demonstrativo de Vendas"
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
      TabIndex        =   11
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
      TabIndex        =   9
      Top             =   3810
      Width           =   6105
      Begin VB.CommandButton cmdRelatorio 
         Default         =   -1  'True
         Height          =   735
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   30
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Height          =   735
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   6
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
      TabIndex        =   7
      Top             =   1200
      Width           =   5775
      Begin VB.ComboBox cboVendedor 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   720
         Width           =   3135
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
      Begin VB.Label Label2 
         Caption         =   "Vendedor"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Até"
         Height          =   255
         Left            =   2640
         TabIndex        =   10
         Tag             =   "lblIdCliente"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCliente 
         Caption         =   "Período : "
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
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
Attribute VB_Name = "frmRelDemoVendas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdRelatorio_Click()
  On Error GoTo TratErro
  AmpS
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngVENDEDORID             As Long
  '
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
  ElseIf (gsNivel = "LOJ") And (cboVendedor.Text = "" Or cboVendedor.Text = "<TODOS>") Then
    'Loja é obrigatório selecionar o vendedor
    AmpN
    TratarErroPrevisto "Selecionar o Vendedor", "[cmdRelatorio_Click]"
    SetarFoco cboVendedor
    Exit Sub
    
  End If
  'obter campos
  Set objGeral = New busSisMetal.clsGeral
  'VENDEDOR
  lngVENDEDORID = 0
  If cboVendedor.Text <> "" And cboVendedor.Text <> "<TODOS>" Then
    strSql = "SELECT PKID FROM PESSOA " & _
        " INNER JOIN FUNCIONARIO ON PESSOA.PKID = FUNCIONARIO.PESSOAID " & _
        " WHERE NOME = " & Formata_Dados(cboVendedor.Text, tpDados_Texto) & _
        " AND FUNCIONARIO.NIVEL = " & Formata_Dados("LOJ", tpDados_Texto)
    '
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngVENDEDORID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  '
  Set objGeral = Nothing
  '
  If optSai1.Value Then
    Report1.Destination = 0 'Video
  ElseIf optSai2.Value Then
    Report1.Destination = 1   'Impressora
  End If
  Report1.CopiesToPrinter = 1
  Report1.WindowState = crptMaximized
  '
  'Contas Pagas
  Report1.ReportFileName = gsReportPath & "DemoVendas.rpt"
  '
  Report1.Formulas(0) = "DataIni = Date(" & Right(mskData(0).Text, 4) & ", " & Mid(mskData(0).Text, 4, 2) & ", " & Left(mskData(0).Text, 2) & ")"
  Report1.Formulas(1) = "DataFim = Date(" & Right(mskData(1).Text, 4) & ", " & Mid(mskData(1).Text, 4, 2) & ", " & Left(mskData(1).Text, 2) & ")"
  If cboVendedor.Text = "" Or cboVendedor.Text = "<TODOS>" Then
    Report1.Formulas(2) = "VENDEDORID = True = true"
  Else
    Report1.Formulas(2) = "VENDEDORID = {PESSOA.PKID} = " & lngVENDEDORID
  End If
  Report1.Action = 1
  '
  AmpN
  Exit Sub
  
TratErro:
  AmpN
  TratarErro Err.Number, Err.Description, "[frmRelDemoVendas.cmdRelatorio_Click]"
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
  strSql = "Select PESSOA.NOME " & _
    " FROM PESSOA INNER JOIN FUNCIONARIO ON PESSOA.PKID = FUNCIONARIO.PESSOAID "
  strSql = strSql & " WHERE FUNCIONARIO.NIVEL = " & Formata_Dados("LOJ", tpDados_Texto)
  If gsNivel = "LOJ" Then
    strSql = strSql & " AND PESSOA.PKID = " & Formata_Dados(giFuncionarioId, tpDados_Longo)
  End If
  strSql = strSql & " ORDER BY PESSOA.NOME "
  PreencheCombo cboVendedor, strSql, True, True
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, "[frmRelDemoVendas.Form_Load]"
End Sub

Private Sub mskData_GotFocus(Index As Integer)
  Selecionar_Conteudo mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub



