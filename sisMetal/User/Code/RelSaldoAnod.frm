VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRelSaldoAnod 
   Caption         =   "Saldo de Perfil na Anodizadora"
   ClientHeight    =   3825
   ClientLeft      =   450
   ClientTop       =   1920
   ClientWidth     =   6105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3825
   ScaleWidth      =   6105
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5775
      Begin VB.OptionButton optSai1 
         Caption         =   "Vídeo"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optSai2 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   240
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
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   6105
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2850
      Width           =   6105
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   880
         Left            =   3150
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdRelatorio 
         Caption         =   "ENTER"
         Default         =   -1  'True
         Height          =   880
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   60
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
      Height          =   1815
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   5775
      Begin VB.ComboBox cboAnodizadora 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   240
         Width           =   4335
      End
      Begin VB.Label lblCliente 
         Caption         =   "Anodizadora: "
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   7
         Tag             =   "lblIdCliente"
         Top             =   240
         Width           =   975
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
Attribute VB_Name = "frmRelSaldoAnod"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboAnodizadora_LostFocus()
  Pintar_Controle cboAnodizadora, tpCorContr_Normal
End Sub

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdRelatorio_Click()
  Dim objGer              As busSisMetal.clsGeral
  Dim lngANODIZADORAID    As Long
  '
  On Error GoTo TratErro
  AmpS
  '
  Set objGer = New busSisMetal.clsGeral
  'ANODIZADORA
  lngANODIZADORAID = 0
  strSql = "SELECT LOJA.PKID FROM LOJA " & _
    " INNER JOIN ANODIZADORA ON ANODIZADORA.LOJAID = LOJA.PKID " & _
    " WHERE LOJA.NOME = " & Formata_Dados(cboAnodizadora.Text, tpDados_Texto)
  Set objRs = objGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngANODIZADORAID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGer = Nothing
  '
  If optSai1.Value Then
    Report1.Destination = 0 'Video
  ElseIf optSai2.Value Then
    Report1.Destination = 1   'Impressora
  End If
  Report1.CopiesToPrinter = 1
  Report1.WindowState = crptMaximized
  '
  Report1.Formulas(0) = "ANODIZADORAID = " & IIf(lngANODIZADORAID = 0, "true = true", "{VW_CONS_ESTOQUE_PERFIL.ANODIZADORAID} = " & lngANODIZADORAID) & ""
  '
  Report1.Action = 1
  '
  AmpN
  Exit Sub
  
TratErro:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Private Sub Form_Activate()
  SetarFoco cboAnodizadora
End Sub

Private Sub Form_Load()
  On Error GoTo RotErro
  AmpS
  Me.Height = 4335
  Me.Width = 6225
  CenterForm Me
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar, , , , , , cmdRelatorio
  '
  strSql = "SELECT NOME FROM LOJA " & _
      " INNER JOIN ANODIZADORA ON LOJA.PKID = ANODIZADORA.LOJAID " & _
      " ORDER BY LOJA.NOME "
  PreencheCombo cboAnodizadora, strSql, False, True
  Report1.Connect = ConnectRpt
  Report1.ReportFileName = gsReportPath & "EstoqueAnodizadora.rpt"
  '
  AmpN
  Exit Sub
RotErro:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub
