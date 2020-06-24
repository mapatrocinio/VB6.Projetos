VERSION 5.00
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "CRYSTL32.OCX"
Begin VB.Form frmUserRelGrupoSubGrupo 
   Caption         =   "Listagem de Grupo / Sub Grupo"
   ClientHeight    =   3840
   ClientLeft      =   450
   ClientTop       =   1920
   ClientWidth     =   6105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3840
   ScaleWidth      =   6105
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   4
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
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   6105
      TabIndex        =   3
      Top             =   3075
      Width           =   6105
      Begin VB.CommandButton cmdRelatorio 
         Default         =   -1  'True
         Height          =   735
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Height          =   735
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   0
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
      Height          =   1695
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   5775
      Begin VB.CheckBox chkMostrarSubGrupo 
         Caption         =   "Mostrar Sub Grupo"
         Height          =   255
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Value           =   1  'Checked
         Width           =   1695
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
Attribute VB_Name = "frmUserRelGrupoSubGrupo"
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
  Set objGeral = New busSisContas.clsGeral
  '
  Report1.CopiesToPrinter = 1
  Report1.WindowState = crptMaximized
  '
  Report1.Formulas(0) = "MostrarSubGrupo = " & IIf(chkMostrarSubGrupo.Value = 1, "true", "false")
  
  Report1.Action = 1
  '
  AmpN
  Set objGeral = Nothing
  Exit Sub
  
TratErro:
  AmpN
  TratarErro Err.Number, Err.Description, "[frmUserRelGrupoSubGrupo.cmdRelatorio_Click]"
End Sub

Private Sub Form_Activate()
  SetarFoco chkMostrarSubGrupo
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  AmpS
  CenterForm Me
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar, , , , , , cmdRelatorio
  '
  Report1.Connect = ConnectRpt
  '
  Report1.ReportFileName = gsReportPath & "ListaGrupoSubGrupo.rpt"
  '
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, "[frmUserRelGrupoSubGrupo.Form_Load]"
End Sub

