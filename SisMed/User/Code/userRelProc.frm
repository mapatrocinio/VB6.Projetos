VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmUserRelProc 
   Caption         =   "Prestadores x Procedimentos"
   ClientHeight    =   4485
   ClientLeft      =   450
   ClientTop       =   1920
   ClientWidth     =   6105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4485
   ScaleWidth      =   6105
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   9
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
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   3510
      Width           =   6105
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   880
         Left            =   3150
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdRelatorio 
         Caption         =   "ENTER"
         Default         =   -1  'True
         Height          =   880
         Left            =   1950
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
      Height          =   2535
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Width           =   5775
      Begin VB.TextBox txtProcedimento 
         Height          =   315
         Left            =   1320
         TabIndex        =   3
         Top             =   540
         Width           =   4155
      End
      Begin VB.ComboBox cboPrestador 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   180
         Width           =   4185
      End
      Begin VB.Label lblCliente 
         Caption         =   "Prestador :"
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   10
         Tag             =   "lblIdCliente"
         Top             =   180
         Width           =   975
      End
      Begin VB.Label lblCliente 
         Caption         =   "Procedimento :"
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   8
         Tag             =   "lblIdCliente"
         Top             =   570
         Width           =   1275
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3360
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmUserRelProc"
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
  '
  If optSai1.Value Then
    Report1.Destination = 0 'Video
  ElseIf optSai2.Value Then
    Report1.Destination = 1   'Impressora
  End If
  Report1.CopiesToPrinter = 1
  Report1.WindowState = crptMaximized
  '
  Report1.Formulas(0) = "Procedimento = '" & txtProcedimento.Text & "*'"
  Report1.Formulas(1) = "Prestador = '" & IIf(cboPrestador.Text = "<TODOS>", "*", cboPrestador.Text) & "'"
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
  SetarFoco cboPrestador
End Sub

Private Sub Form_Load()
  On Error GoTo RotErro
  AmpS
  CenterForm Me
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar, , , , , , cmdRelatorio
  '
  Report1.Connect = ConnectRpt
  Report1.ReportFileName = gsReportPath & "Procedimento.rpt"
  '
  'Prestadores
  strSql = "SELECT PRONTUARIO.NOME FROM PRONTUARIO " & _
    " INNER JOIN PRESTADOR ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID " & _
    " ORDER BY PRONTUARIO.NOME"
  PreencheCombo cboPrestador, strSql, True, False, True
  AmpN
  Exit Sub
RotErro:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

