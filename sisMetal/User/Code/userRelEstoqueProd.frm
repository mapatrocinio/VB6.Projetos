VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRelEstoqueProd 
   Caption         =   "Demonstrativo de Estoque de Produtos"
   ClientHeight    =   3675
   ClientLeft      =   450
   ClientTop       =   1920
   ClientWidth     =   6105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3675
   ScaleWidth      =   6105
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   8
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
      TabIndex        =   7
      Top             =   2820
      Width           =   6105
      Begin VB.CommandButton cmdRelatorio 
         Default         =   -1  'True
         Height          =   735
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   30
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Height          =   735
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   5
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
      Height          =   1485
      Left            =   120
      TabIndex        =   6
      Top             =   1200
      Width           =   5775
      Begin VB.ComboBox cboFamilia 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   540
         Width           =   4425
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   4425
      End
      Begin VB.Label Label3 
         Caption         =   "Família"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   540
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Grupo"
         Height          =   255
         Left            =   240
         TabIndex        =   9
         Top             =   180
         Width           =   735
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
Attribute VB_Name = "frmRelEstoqueProd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdRelatorio_Click()
  On Error GoTo TratErro
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngGRUPOID                As Long
  Dim lngFAMILIAID              As Long
  
  AmpS
  '
  If optSai1.Value Then
    Report1.Destination = 0 'Video
  ElseIf optSai2.Value Then
    Report1.Destination = 1   'Impressora
  End If
  'obter campos
  Set objGeral = New busSisMetal.clsGeral
  'GRUPO
  lngGRUPOID = 0
  strSql = "SELECT PKID FROM GRUPO_PRODUTO WHERE NOME = " & Formata_Dados(cboGrupo.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngGRUPOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'FAMILIA
  lngFAMILIAID = 0
  strSql = "SELECT PKID FROM FAMILIAPRODUTOS WHERE DESCRICAO = " & Formata_Dados(cboFamilia.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngFAMILIAID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set objGeral = Nothing
  
  Report1.CopiesToPrinter = 1
  Report1.WindowState = crptMaximized
  '
  Report1.ReportFileName = gsReportPath & "EstoqueProduto.rpt"
  '
  If lngGRUPOID = 0 Then
    Report1.Formulas(2) = "GRUPOID = True = true"
  Else
    Report1.Formulas(2) = "GRUPOID = {GRUPO_PRODUTO.PKID} = " & lngGRUPOID
  End If
  If lngFAMILIAID = 0 Then
    Report1.Formulas(3) = "FAMILIAID = True = true"
  Else
    Report1.Formulas(3) = "FAMILIAID = {FAMILIAPRODUTOS.PKID} = " & lngFAMILIAID
  End If
  Report1.Action = 1
  '
  AmpN
  Exit Sub
  
TratErro:
  AmpN
  TratarErro Err.Number, Err.Description, "[frmRelEstoqueProd.cmdRelatorio_Click]"
End Sub

Private Sub Form_Activate()
  cboGrupo.SetFocus
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
  strSql = "Select GRUPO_PRODUTO.NOME " & _
    " FROM GRUPO_PRODUTO " & _
    " ORDER BY GRUPO_PRODUTO.NOME "
  PreencheCombo cboGrupo, strSql, False, True
  strSql = "Select FAMILIAPRODUTOS.DESCRICAO " & _
    " FROM FAMILIAPRODUTOS " & _
    " ORDER BY FAMILIAPRODUTOS.DESCRICAO "
  PreencheCombo cboFamilia, strSql, False, True
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, "[frmRelEstoqueProd.Form_Load]"
End Sub
