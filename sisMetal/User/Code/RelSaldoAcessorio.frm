VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmRelSaldoAcessorio 
   Caption         =   "Saldo de Acessórios"
   ClientHeight    =   3825
   ClientLeft      =   450
   ClientTop       =   1920
   ClientWidth     =   7605
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3825
   ScaleWidth      =   7605
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   60
      TabIndex        =   8
      Top             =   120
      Width           =   7395
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
      ScaleWidth      =   7605
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   2850
      Width           =   7605
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   880
         Left            =   3960
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdRelatorio 
         Caption         =   "ENTER"
         Default         =   -1  'True
         Height          =   880
         Left            =   2760
         Style           =   1  'Graphical
         TabIndex        =   4
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
      TabIndex        =   6
      Top             =   840
      Width           =   7335
      Begin VB.CheckBox chkAbaixo 
         Caption         =   "Abaixo do estoque mínimo"
         Height          =   375
         Left            =   1350
         TabIndex        =   3
         Top             =   510
         Width           =   3255
      End
      Begin VB.ComboBox cboGrupo 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   3855
      End
      Begin VB.Label Label5 
         Caption         =   "Status"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   10
         Top             =   570
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Grupo"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1215
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
Attribute VB_Name = "frmRelSaldoAcessorio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cboGrupo_LostFocus()
  Pintar_Controle cboGrupo, tpCorContr_Normal
End Sub


Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdRelatorio_Click()
  On Error GoTo TratErro
  Dim objGeral                  As busSisMetal.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngGRUPOID                As Long
  AmpS
  '
  'obter campos
  Set objGeral = New busSisMetal.clsGeral
  'GRUPO
  lngGRUPOID = 0
  strSql = "SELECT PKID, NOME FROM GRUPO WHERE NOME = " & Formata_Dados(cboGrupo.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngGRUPOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
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
  Report1.Formulas(0) = "GRUPOID = " & IIf(lngGRUPOID = 0, "true = true", "{GRUPO.PKID} = " & lngGRUPOID) & ""
  Report1.Formulas(1) = "TOTAL = " & IIf(chkAbaixo.Value = 0, "true = true", "{@Balanco} <= 0 ")
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
  SetarFoco cboGrupo
End Sub

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Ítens do relatório
  LimparCampoCombo cboGrupo
  LimparCampoCheck chkAbaixo
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmRelSaldoAcessorio.LimparCampos]", _
            Err.Description
End Sub

Private Sub Form_Load()
  On Error GoTo RotErro
  AmpS
  Me.Height = 4335
  Me.Width = 7725
  CenterForm Me
  '
  LimparCampos
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar, , , , , , cmdRelatorio
  '
  'Grupo
  strSql = "Select NOME from GRUPO ORDER BY NOME"
  PreencheCombo cboGrupo, strSql, False, True
  '
  Report1.Connect = ConnectRpt
  Report1.ReportFileName = gsReportPath & "EstoqueAcessorio.rpt"
  '
  AmpN
  Exit Sub
RotErro:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub


