VERSION 5.00
Begin VB.Form frmUserAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sobre - Sistema Folha de Pagamento"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   270
   ClientWidth     =   5715
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   5715
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOk 
      Caption         =   "ENTER"
      Default         =   -1  'True
      Height          =   880
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3030
      Width           =   1215
   End
   Begin VB.Label lblVersao 
      BackStyle       =   0  'Transparent
      Caption         =   "Versão : "
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
      Left            =   1560
      TabIndex        =   4
      Top             =   2160
      Width           =   3855
   End
   Begin VB.Label lblDatahora 
      BackStyle       =   0  'Transparent
      Caption         =   "Data e Hora : "
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
      Left            =   1560
      TabIndex        =   3
      Top             =   1920
      Width           =   3855
   End
   Begin VB.Label lblEmpresa 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   2
      Top             =   1440
      Width           =   3735
   End
   Begin VB.Label lblTitulo 
      BackStyle       =   0  'Transparent
      Caption         =   "Título"
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
      Left            =   1560
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Empresa:"
      Height          =   255
      Left            =   1560
      TabIndex        =   0
      Top             =   960
      Width           =   2535
   End
End
Attribute VB_Name = "frmUserAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub cmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  CenterForm Me
  '
  Me.Caption = "Sobre " & TITULOSISTEMA & " / " & gsNomeEmpresa
  lblEmpresa.Caption = gsNomeEmpresa
  lblTitulo.Caption = TITULOSISTEMA
  lblDatahora.Caption = lblDatahora.Caption & Format(Now, "DD/MM/YYYY hh:mm")
  lblVersao.Caption = lblVersao.Caption & App.Major & "." & App.Minor
  '
  LerFiguras Me, tpBmp_Login, pbtnFechar:=Me.cmdOk
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub

