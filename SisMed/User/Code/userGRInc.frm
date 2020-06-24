VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserGRInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de Guia de Recolhimento"
   ClientHeight    =   7905
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7905
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7905
      Left            =   8520
      ScaleHeight     =   7905
      ScaleWidth      =   1860
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4725
         Left            =   90
         ScaleHeight     =   4665
         ScaleWidth      =   1605
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   3120
         Width           =   1665
         Begin VB.CommandButton cmdPagamento 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   7755
      Left            =   120
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   13679
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da GR"
      TabPicture(0)   =   "userGRInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraProf"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraProf 
         Height          =   7395
         Left            =   120
         TabIndex        =   27
         Top             =   330
         Width           =   7935
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   7125
            Index           =   0
            Left            =   120
            ScaleHeight     =   7125
            ScaleWidth      =   7785
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   240
            Width           =   7785
            Begin VB.Frame Frame4 
               Height          =   1725
               Left            =   0
               TabIndex        =   51
               Top             =   5130
               Width           =   7695
               Begin TrueDBGrid60.TDBGrid grdProcedimento 
                  Height          =   1515
                  Left            =   60
                  OleObjectBlob   =   "userGRInc.frx":001C
                  TabIndex        =   18
                  Top             =   150
                  Width           =   7545
               End
            End
            Begin VB.Frame Frame2 
               Height          =   1425
               Left            =   0
               TabIndex        =   43
               Top             =   -90
               Width           =   7695
               Begin VB.PictureBox picTrava 
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   1185
                  Index           =   3
                  Left            =   90
                  ScaleHeight     =   1185
                  ScaleWidth      =   7425
                  TabIndex        =   44
                  TabStop         =   0   'False
                  Top             =   150
                  Width           =   7425
                  Begin VB.TextBox txtTurno 
                     BackColor       =   &H00E0E0E0&
                     Height          =   288
                     Left            =   1200
                     Locked          =   -1  'True
                     TabIndex        =   1
                     TabStop         =   0   'False
                     Text            =   "txtTurno"
                     Top             =   300
                     Width           =   6135
                  End
                  Begin VB.TextBox txtCaixa 
                     BackColor       =   &H00E0E0E0&
                     Height          =   288
                     Left            =   1200
                     Locked          =   -1  'True
                     TabIndex        =   2
                     TabStop         =   0   'False
                     Text            =   "txtCaixa"
                     Top             =   600
                     Width           =   6135
                  End
                  Begin VB.TextBox txtDiaDaSemana 
                     BackColor       =   &H00E0E0E0&
                     Height          =   288
                     Left            =   1200
                     Locked          =   -1  'True
                     TabIndex        =   0
                     TabStop         =   0   'False
                     Text            =   "txtDiaDaSemana"
                     Top             =   0
                     Width           =   6135
                  End
                  Begin VB.PictureBox Picture2 
                     BorderStyle     =   0  'None
                     Enabled         =   0   'False
                     Height          =   255
                     Left            =   4440
                     ScaleHeight     =   255
                     ScaleWidth      =   3045
                     TabIndex        =   45
                     TabStop         =   0   'False
                     Top             =   900
                     Width           =   3045
                     Begin MSMask.MaskEdBox mskData 
                        Height          =   255
                        Index           =   0
                        Left            =   1200
                        TabIndex        =   4
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
                     Begin VB.Label Label2 
                        Caption         =   "Data"
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
                        TabIndex        =   46
                        Top             =   0
                        Width           =   615
                     End
                  End
                  Begin VB.TextBox txtSequencial 
                     BackColor       =   &H00E0E0E0&
                     Height          =   285
                     Left            =   1200
                     Locked          =   -1  'True
                     TabIndex        =   3
                     TabStop         =   0   'False
                     Text            =   "txtSequencial"
                     Top             =   900
                     Width           =   1455
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Turno"
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
                     Index           =   7
                     Left            =   0
                     TabIndex        =   50
                     Top             =   300
                     Width           =   1215
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Caixa"
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
                     Index           =   5
                     Left            =   0
                     TabIndex        =   49
                     Top             =   600
                     Width           =   1215
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Dia"
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
                     Index           =   1
                     Left            =   0
                     TabIndex        =   48
                     Top             =   0
                     Width           =   1215
                  End
                  Begin VB.Label Label44 
                     Caption         =   "Sequencial"
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
                     TabIndex        =   47
                     Top             =   900
                     Width           =   1065
                  End
               End
            End
            Begin VB.Frame Frame1 
               Height          =   1155
               Left            =   0
               TabIndex        =   30
               Top             =   3990
               Width           =   7695
               Begin VB.PictureBox picTrava 
                  BorderStyle     =   0  'None
                  Height          =   975
                  Index           =   2
                  Left            =   60
                  ScaleHeight     =   975
                  ScaleWidth      =   7575
                  TabIndex        =   32
                  TabStop         =   0   'False
                  Top             =   150
                  Width           =   7575
                  Begin VB.TextBox txtProcedimento 
                     Height          =   285
                     Left            =   1260
                     MaxLength       =   100
                     TabIndex        =   14
                     Top             =   0
                     Width           =   6135
                  End
                  Begin VB.TextBox txtProcedimentoFim 
                     BackColor       =   &H00E0E0E0&
                     Height          =   285
                     Left            =   1260
                     Locked          =   -1  'True
                     MaxLength       =   100
                     TabIndex        =   15
                     TabStop         =   0   'False
                     Top             =   300
                     Width           =   6135
                  End
                  Begin MSMask.MaskEdBox mskQuantidade 
                     Height          =   255
                     Left            =   2520
                     TabIndex        =   17
                     Top             =   600
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   450
                     _Version        =   393216
                     Format          =   "#,##0;($#,##0)"
                     PromptChar      =   "_"
                  End
                  Begin MSMask.MaskEdBox mskValor 
                     Height          =   255
                     Left            =   1260
                     TabIndex        =   16
                     Top             =   600
                     Width           =   1245
                     _ExtentX        =   2196
                     _ExtentY        =   450
                     _Version        =   393216
                     BackColor       =   14737632
                     Enabled         =   0   'False
                     Format          =   "#,##0.00;($#,##0.00)"
                     PromptChar      =   "_"
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Procedimento"
                     Height          =   255
                     Index           =   0
                     Left            =   30
                     TabIndex        =   35
                     Top             =   30
                     Width           =   1095
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Procedimento"
                     Height          =   255
                     Index           =   3
                     Left            =   30
                     TabIndex        =   34
                     Top             =   300
                     Width           =   1095
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Valor/Qtd"
                     Height          =   255
                     Index           =   1
                     Left            =   30
                     TabIndex        =   33
                     Top             =   600
                     Width           =   1095
                  End
               End
            End
            Begin VB.Frame Frame3 
               BackColor       =   &H00C0C0C0&
               Caption         =   "Cadastro"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2655
               Left            =   0
               TabIndex        =   29
               Top             =   1350
               Width           =   7695
               Begin VB.PictureBox picTrava 
                  BackColor       =   &H00C0C0C0&
                  BorderStyle     =   0  'None
                  Height          =   2415
                  Index           =   1
                  Left            =   120
                  ScaleHeight     =   2415
                  ScaleWidth      =   7455
                  TabIndex        =   36
                  TabStop         =   0   'False
                  Top             =   180
                  Width           =   7455
                  Begin VB.TextBox txtDescricao 
                     Height          =   585
                     Left            =   1230
                     MaxLength       =   255
                     MultiLine       =   -1  'True
                     TabIndex        =   13
                     Text            =   "userGRInc.frx":4E7D
                     Top             =   1830
                     Width           =   6135
                  End
                  Begin VB.TextBox txtPrestEspec 
                     Height          =   285
                     Left            =   1230
                     MaxLength       =   100
                     TabIndex        =   5
                     Text            =   "txtPrestEspec"
                     Top             =   30
                     Width           =   6135
                  End
                  Begin VB.TextBox txtPrestador 
                     BackColor       =   &H00E0E0E0&
                     Height          =   288
                     Left            =   1230
                     Locked          =   -1  'True
                     TabIndex        =   6
                     TabStop         =   0   'False
                     Text            =   "txtPrestador"
                     Top             =   330
                     Width           =   6135
                  End
                  Begin VB.TextBox txtEspecialidade 
                     BackColor       =   &H00E0E0E0&
                     Height          =   288
                     Left            =   1230
                     Locked          =   -1  'True
                     TabIndex        =   7
                     TabStop         =   0   'False
                     Text            =   "txtEspecialidade"
                     Top             =   630
                     Width           =   6135
                  End
                  Begin VB.TextBox txtSala 
                     BackColor       =   &H00E0E0E0&
                     Height          =   288
                     Left            =   1230
                     Locked          =   -1  'True
                     TabIndex        =   8
                     TabStop         =   0   'False
                     Text            =   "txtSala"
                     Top             =   930
                     Width           =   3495
                  End
                  Begin VB.TextBox txtPeriodo 
                     BackColor       =   &H00E0E0E0&
                     Height          =   288
                     Left            =   4740
                     Locked          =   -1  'True
                     TabIndex        =   9
                     TabStop         =   0   'False
                     Text            =   "txtPeriodo"
                     Top             =   930
                     Width           =   2625
                  End
                  Begin VB.TextBox txtProntuario 
                     Height          =   285
                     Left            =   1230
                     MaxLength       =   100
                     TabIndex        =   10
                     Text            =   "txtProntuario"
                     Top             =   1230
                     Width           =   6135
                  End
                  Begin VB.TextBox txtProntuarioFim 
                     BackColor       =   &H00E0E0E0&
                     Height          =   288
                     Left            =   1230
                     Locked          =   -1  'True
                     TabIndex        =   11
                     TabStop         =   0   'False
                     Text            =   "txtProntuarioFim"
                     Top             =   1530
                     Width           =   4995
                  End
                  Begin MSMask.MaskEdBox mskDataNascFim 
                     Height          =   285
                     Left            =   6210
                     TabIndex        =   12
                     TabStop         =   0   'False
                     Top             =   1530
                     Width           =   1155
                     _ExtentX        =   2037
                     _ExtentY        =   503
                     _Version        =   393216
                     BackColor       =   14737632
                     AutoTab         =   -1  'True
                     MaxLength       =   10
                     Mask            =   "##/##/####"
                     PromptChar      =   "_"
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Descrição"
                     Height          =   255
                     Index           =   9
                     Left            =   30
                     TabIndex        =   56
                     Top             =   1830
                     Width           =   1215
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Prest/Espec"
                     Height          =   255
                     Index           =   3
                     Left            =   30
                     TabIndex        =   42
                     Top             =   30
                     Width           =   1215
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Prestador"
                     Height          =   255
                     Index           =   2
                     Left            =   30
                     TabIndex        =   41
                     Top             =   330
                     Width           =   1215
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Especialidade"
                     Height          =   255
                     Index           =   6
                     Left            =   30
                     TabIndex        =   40
                     Top             =   630
                     Width           =   1215
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Sala/Período"
                     Height          =   255
                     Index           =   8
                     Left            =   30
                     TabIndex        =   39
                     Top             =   930
                     Width           =   1215
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Prontuário"
                     Height          =   255
                     Index           =   0
                     Left            =   30
                     TabIndex        =   38
                     Top             =   1230
                     Width           =   1215
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Prontuário"
                     Height          =   255
                     Index           =   4
                     Left            =   30
                     TabIndex        =   37
                     Top             =   1530
                     Width           =   1215
                  End
               End
            End
            Begin ComctlLib.StatusBar StatusBar1 
               Height          =   255
               Left            =   5970
               TabIndex        =   31
               Top             =   6840
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   450
               SimpleText      =   ""
               _Version        =   327682
               BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
                  NumPanels       =   5
                  BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
                     Style           =   6
                     Alignment       =   1
                     Bevel           =   2
                     Object.Width           =   1940
                     MinWidth        =   1940
                     TextSave        =   "1/10/2011"
                     Key             =   ""
                     Object.Tag             =   ""
                  EndProperty
                  BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
                     Style           =   5
                     Alignment       =   1
                     Bevel           =   2
                     Object.Width           =   1235
                     MinWidth        =   1235
                     TextSave        =   "17:16"
                     Key             =   ""
                     Object.Tag             =   ""
                  EndProperty
                  BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
                     Style           =   1
                     Alignment       =   1
                     Bevel           =   2
                     Enabled         =   0   'False
                     Object.Visible         =   0   'False
                     Object.Width           =   1235
                     MinWidth        =   1235
                     TextSave        =   "CAPS"
                     Key             =   ""
                     Object.Tag             =   ""
                  EndProperty
                  BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
                     Style           =   2
                     Alignment       =   1
                     Bevel           =   2
                     Enabled         =   0   'False
                     Object.Visible         =   0   'False
                     Object.Width           =   1235
                     MinWidth        =   1235
                     TextSave        =   "NUM"
                     Key             =   ""
                     Object.Tag             =   ""
                  EndProperty
                  BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
                     Style           =   3
                     Alignment       =   1
                     AutoSize        =   2
                     Bevel           =   2
                     Enabled         =   0   'False
                     Object.Visible         =   0   'False
                     Object.Width           =   1244
                     MinWidth        =   1235
                     TextSave        =   "INS"
                     Key             =   ""
                     Object.Tag             =   ""
                  EndProperty
               EndProperty
            End
            Begin VB.Label lblCor 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFF80&
               Caption         =   "Movimento após o fechamento"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   990
               TabIndex        =   55
               Top             =   6870
               Width           =   2625
            End
            Begin VB.Label lblCor 
               Alignment       =   2  'Center
               Caption         =   "Status :"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   0
               Left            =   0
               TabIndex        =   54
               Top             =   6870
               Width           =   765
            End
            Begin VB.Label lblCor 
               Alignment       =   2  'Center
               Caption         =   "Impressão:"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   2
               Left            =   3780
               TabIndex        =   53
               Top             =   6870
               Width           =   915
            End
            Begin VB.Label lblCor 
               Alignment       =   2  'Center
               BackColor       =   &H0080FF80&
               Caption         =   "Sim"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   3
               Left            =   4860
               TabIndex        =   52
               Top             =   6870
               Width           =   525
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserGRInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum tpIcEstadoGR
  tpIcEstadoGR_Inic = 0
  tpIcEstadoGR_Proc = 1
  tpIcEstadoGR_Con = 2
End Enum

Public Status                 As tpStatus
Public IcEstadoGR             As tpIcEstadoGR
Public lngGRID                As Long
Public blnRetorno             As Boolean
Public blnFechar              As Boolean
Public strAceitaValor         As String
'Public intQuemChamou          As Integer
'Public strMotivo              As String
'Public lngQtdRestanteProdReal As Long
Private blnPrimeiraVez        As Boolean
Private strCortesiaGR         As String
Private strStatusGR           As String
'Private blnMostrarAlertaProd  As Boolean
'Private blnAlterouItem        As Boolean

'Variáveis para Grid ObraEngenheiro
Dim PROC_COLUNASMATRIZ         As Long
Dim PROC_LINHASMATRIZ          As Long
Private PROC_Matriz()          As String




Public Sub PROC_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGer    As busSisMed.clsGeral
  '
  On Error GoTo trata

  Set objGer = New busSisMed.clsGeral
  '
  strSql = "SELECT GRPROCEDIMENTO.QTD, GRPROCEDIMENTO.VALOR, PROCEDIMENTO.PROCEDIMENTO, GRPROCEDIMENTO.PKID, GRPROCEDIMENTO.GRID, GRPROCEDIMENTO.PROCEDIMENTOID " & _
          "FROM GRPROCEDIMENTO INNER JOIN  PROCEDIMENTO ON GRPROCEDIMENTO.PROCEDIMENTOID =  PROCEDIMENTO.PKID " & _
          "WHERE GRPROCEDIMENTO.GRID = " & lngGRID & " " & _
          "ORDER BY GRPROCEDIMENTO.PKID"

  '
  Set objRs = objGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    PROC_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim PROC_Matriz(0 To PROC_COLUNASMATRIZ - 1, 0 To PROC_LINHASMATRIZ - 1)
  Else
    ReDim PROC_Matriz(0 To PROC_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To PROC_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To PROC_COLUNASMATRIZ - 1  'varre as colunas
          PROC_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set objGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub


Private Sub cmdExcluir_Click()
  Dim objGR               As busSisMed.clsGR
  Dim objGer              As busSisMed.clsGeral
  Dim objRs               As ADODB.Recordset
  Dim strSql              As String
  '
  On Error GoTo trata
  If Len(Trim(grdProcedimento.Columns("Procedimento").Value & "")) = 0 Then
    MsgBox "Selecione um Procedimento para exclusão.", vbExclamation, TITULOSISTEMA
    SetarFoco grdProcedimento
    Exit Sub
  End If
  '
  Set objGer = New busSisMed.clsGeral
  '
  If MsgBox("Confirma exclusão do Procedimento " & grdProcedimento.Columns("Procedimento").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdProcedimento
    Exit Sub
  End If
  'OK
  Set objGR = New busSisMed.clsGR
  
  objGR.ExcluirGRPROCEDIMENTO CLng(grdProcedimento.Columns("GRPROCEDIMENTOID").Value)
  'Verifica movimento após o fechamento
  VerificaMovAposFecha lngGRID
  '
  'Montar RecordSet
  PROC_COLUNASMATRIZ = grdProcedimento.Columns.Count
  PROC_LINHASMATRIZ = 0
  PROC_MontaMatriz
  grdProcedimento.Bookmark = Null
  grdProcedimento.ReBind
  SetarFoco txtProcedimento
  Form_Load
  Set objGR = Nothing
  SetarFoco grdProcedimento
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub





Private Sub cmdPagamento_Click()
  Dim objUserContaCorrente  As SisMed.frmUserContaCorrente
  Dim strSql                As String
  Dim objRs                 As ADODB.Recordset
  Dim objGeral              As busSisMed.clsGeral
  Dim lngTotal              As Long
  Dim objGR                 As busSisMed.clsGR
  On Error GoTo trata
  'CAPTURA TOTAL
  lngTotal = 0
  Set objGeral = New busSisMed.clsGeral
  strSql = "SELECT COUNT(PKID) AS TOTAL " & _
    " FROM GRPROCEDIMENTO " & _
    " WHERE GRPROCEDIMENTO.GRID = " & Formata_Dados(lngGRID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    If IsNumeric(objRs.Fields("TOTAL").Value) Then
      lngTotal = objRs.Fields("TOTAL").Value
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If lngTotal = 0 Then
    MsgBox "Não há itens lançados nesta GR.", vbExclamation, TITULOSISTEMA
    SetarFoco txtProcedimento
    Exit Sub
  End If
  If strCortesiaGR = "S" And strStatusGR & "" = "I" Then
    'Cortesia com status Inicial
    MsgBox "Atenção. esta é uma GR de cortesia para funcionário, por isso não há pagamento para ela. Seu status será alterado para [fchado].", vbExclamation, TITULOSISTEMA
    Set objGR = New busSisMed.clsGR
    objGR.AlterarStatusGR lngGRID, _
                          "F", _
                          ""
    Set objGR = Nothing
    '
    Form_Load
  Else
    '
    Set objUserContaCorrente = New frmUserContaCorrente
    objUserContaCorrente.lngCCID = 0
    objUserContaCorrente.lngGRID = lngGRID
    objUserContaCorrente.intGrupo = 0
    objUserContaCorrente.strFuncionarioNome = gsNomeUsuCompleto
    If Status = tpStatus_Consultar Then
      objUserContaCorrente.Status = tpStatus_Consultar
    Else
      objUserContaCorrente.Status = tpStatus_Incluir
    End If
    objUserContaCorrente.strStatusLanc = "RC"
    objUserContaCorrente.strNivelAcesso = ""
    objUserContaCorrente.Show vbModal
    If objUserContaCorrente.blnRetorno = True Then
      Form_Load
    End If
    Set objUserContaCorrente = Nothing
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserGRInc.cmdExcluir_Click]"
End Sub

'''
'''Private Sub cmdExcluir_Click()
'''  Dim objRs                     As ADODB.Recordset
'''  Dim strSql                    As String
'''  Dim intI                      As Integer
'''  Dim strMsg                    As String
'''  Dim strMsgErro                As String
'''  '
'''  Dim strQtd                    As String
'''  Dim strValor                  As String
'''  Dim strItem                   As String
'''  Dim strDesc                   As String
'''  Dim clsPed                    As busSisMed.clsPedido
'''  Dim clsLoc                    As busSisMed.clsLocacao
'''  Dim clsvend                   As busSisMed.clsVenda
'''  Dim lngQtdRestanteProd        As Long
'''  Dim curValorUnitarioProd      As Currency
'''  '
'''  On Error GoTo trata
'''  '
'''  Select Case tabDetalhes.Tab
'''  Case 0 'Inclusão/Alteração de Venda
'''  Case 2
'''    If grdProcedimento.Columns(1).Text = "" Then
'''      MsgBox "Selecione um item da venda para exclui-lo/extorna-lo.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdProcedimento
'''      Exit Sub
'''    End If
'''    Set clsvend = New busSisMed.clsVenda
'''    Set clsPed = New busSisMed.clsPedido
'''    Set clsLoc = New busSisMed.clsLocacao
'''    'VERIFICAR SE ESSA FUNCIONALIDADE IRÁ ENTRAR DEPOIS
''''''    If Not clsPed.TemPermExcluirPedido(gsNivel, lngGRID, strMsgErro) Then
''''''      MsgBox strMsgErro, vbExclamation, TITULOSISTEMA
''''''      Set clsPed = Nothing
''''''      Set clsLoc = Nothing
''''''      Exit Sub
''''''    End If
'''    If MsgBox("Deseja excluir/estornar o item do cardápio nro. " & grdProcedimento.Columns(2).Text & " desta venda ?", vbYesNo, TITULOSISTEMA) = vbYes Then
''''''      If intQuemChamou = 1 Then 'Chamada da Exclusão
''''''        '----------------------------
''''''        '----------------------------
''''''        'Pede Senha Superior (Diretor, Gerente ou Administrador
''''''        'Apenas na alteração, na Inclusão não
''''''        If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
''''''          'Só pede senha superior se quem estiver logado não for superior
''''''          frmUserLoginSup.Show vbModal
''''''
''''''          If Len(Trim(gsNomeUsuLib)) = 0 Then
''''''            Set clsvend = Nothing
''''''            Set clsLoc = Nothing
''''''            strMsg = "Para efetuar a exclusão/estorno de itens da venda é necessário a confirmação com senha superior."
''''''            TratarErroPrevisto strMsg, "frmUserVendaInc.cmdExcluir_Click"
''''''            SetarFoco grdProcedimento
''''''            Exit Sub
''''''          End If
''''''          '
''''''          'Capturou Nome do Usuário, continua processo de Sangria
''''''        Else
''''''          gsNomeUsuLib = gsNomeUsu
''''''        End If
''''''      Else
''''''        gsNomeUsuLib = gsNomeUsu
''''''      End If
'''
'''      'Independente da Chamada ou se motel trabalha ou não com estorno
'''      'Verificar se produto que está sendo estornado já não foi estornado ou se não é o próprio estorno
'''      If Not clsvend.VerificaEstorno(CLng(grdProcedimento.Columns("TAB_VENDAITEMID").Text), strMsgErro, lngQtdRestanteProd, curValorUnitarioProd) Then
'''        Set clsvend = Nothing
'''        Set clsLoc = Nothing
'''        Set clsPed = Nothing
'''        TratarErroPrevisto strMsgErro, "userVendaInc.cmdExcluir_Click"
'''        SetarFoco grdProcedimento
'''        Exit Sub
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''      frmUserTextoExc.QuemChamou = 4 'Chamada de Venda
'''      frmUserTextoExc.QuantidadeASerEstornada = lngQtdRestanteProd
'''      frmUserTextoExc.Show vbModal
'''      If Len(strMotivo & "") = 0 Then
'''        SetarFoco grdProcedimento
'''        Exit Sub
'''      End If
'''      If grdProcedimento.Columns(1).Text <> "" Then
'''
'''
'''        'caso trabalhe com estorno de mercadoria,
'''        'pede a quantidade a ser estornada,
'''        'estorna do estoque
'''        'Guarda os Valores para impressão
'''
'''        If gbTrabComEstorno Then
'''          clsvend.EstornaGRPROCEDIMENTO CLng(grdProcedimento.Columns("TAB_VENDAITEMID").Text), lngGRID, CInt(grdProcedimento.Columns(0).Text), lngQtdRestanteProdReal
'''        End If
'''        'Carrega variáveis para impressão
'''        strQtd = CStr(Format(lngQtdRestanteProdReal, "###,##0"))
'''        strValor = CStr(Format(curValorUnitarioProd * lngQtdRestanteProdReal, "###,##0.00"))
'''        strItem = grdProcedimento.Columns(0).Text
'''        strDesc = grdProcedimento.Columns(2).Text
'''        strDesc = strDesc & " - " & grdProcedimento.Columns(4).Text
'''        '
'''        clsPed.RetornaProdutoEstoque grdProcedimento.Columns(2).Text, CStr(IIf(gbTrabComEstorno, lngQtdRestanteProdReal, lngQtdRestanteProd)), 0
'''        '
'''        If Not gbTrabComEstorno Then 'Caso não trabalhe com estorno, deleta
'''          clsvend.ExcluirGRPROCEDIMENTO grdProcedimento.Columns("TAB_VENDAITEMID").Text
'''          'clsPed.ExcluirPedidoSeVazio lngPEDIDOID
'''        End If
''''''        If intQuemChamou = 1 Then 'Só imprime na alteração
''''''          'Imprimir Exc Venda
''''''          IMP_COMPROV_CANC_VENDA lngGRID, gsNomeEmpresa, 1, strQtd, strValor, strItem, strDesc, strMotivo
''''''        End If
'''        '
'''        blnAlterouItem = True
'''        Set clsPed = Nothing
'''        Set clsLoc = Nothing
'''        Set clsvend = Nothing
'''      End If
'''    End If
'''    'Montar RecordSet
'''    PROC_COLUNASMATRIZ = 8
'''    PROC_LINHASMATRIZ = 0
'''    PROC_MontaMatriz
'''    grdProcedimento.Bookmark = Null
'''    grdProcedimento.ReBind
'''    '
'''    SetarFoco grdProcedimento
'''  End Select
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, "[frmUserVendaInc.cmdExcluir_Click]"
'''End Sub
'''


'''End Sub
'''
Private Sub grdProcedimento_UnboundReadDataEx( _
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
               Offset + intI, PROC_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, PROC_COLUNASMATRIZ, PROC_LINHASMATRIZ, PROC_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, PROC_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserGRInc.grdProcedimento_UnboundReadDataEx]"
End Sub

Private Sub cmdCancelar_Click()
'''  'Rotina para Impressão do Pedido
'''  '------------
'''  'Verificar antes se tem algo a ser impresso
'''  Dim objRs       As ADODB.Recordset
'''  Dim objRsVda    As ADODB.Recordset
'''  Dim strSql      As String
'''  Dim objVenda     As busSisMed.clsVenda
'''  Dim strCobranca As String
'''  Dim objCC       As busSisMed.clsContaCorrente
'''  Dim objForm     As SisMotel.frmUserLocContaCorrente
'''  Dim lngCCId     As Long
'''  Dim curVrVenda  As Currency
  '
  On Error GoTo trata
  '
  blnFechar = True
  blnRetorno = True
'''  Set objVenda = New busSisMed.clsVenda
'''  '
'''  Set objRs = objVenda.ListarGRPROCEDIMENTO(lngGRID)
'''  '
'''  If Not objRs.EOF Then
'''    'Existem Itens de pedido associado ao pedido
'''    'Então imprimir pedido
'''    '
'''    If (intQuemChamou = 0) Or (intQuemChamou = 1 And blnAlterouItem = True) Then 'Chamada original de inclusão
'''      'Antes da impressão
'''      'Novo Preparar para inclusão de pagamento para despesa
'''      If optCobranca(0).Value Then
'''        strCobranca = "S"
'''      Else
'''        strCobranca = "N"
'''      End If
'''      If strCobranca = "S" Then
'''        If chkPgto.Value = False Then
'''          'Pagamento em apenas uma forma de pagamento
'''          'Assumir pagamento em dinheiro
'''          curVrVenda = 0
'''          Set objRsVda = objVenda.ListarVenda(lngGRID)
'''          If Not objRsVda.EOF Then
'''            curVrVenda = objRsVda.Fields("VR_TOT_VENDA").Value
'''          End If
'''          Set objRsVda = Nothing
'''          Set objCC = New busSisMed.clsContaCorrente
'''          lngCCId = objCC.InserirCC(lngGRID, _
'''                            RetornaCodTurnoCorrente, _
'''                            Format(Now, "DD/MM/YYYY hh:mm"), _
'''                            Format(curVrVenda, "##0.00"), _
'''                            "C", _
'''                            "ES", _
'''                            "VD", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            gsNomeUsu, _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "0", _
'''                            "", _
'''                            "")
'''          Set objCC = Nothing
'''        Else
'''          Set objForm = New frmUserLocContaCorrente
'''          objForm.lngLOCDESPVDAEXTID = lngGRID
'''          objForm.intGrupo = 0
'''          objForm.strNumeroAptoPrinc = ""
'''          objForm.Status = tpStatus_Incluir
'''          objForm.strStatusLanc = "VD"
'''          objForm.Show vbModal
'''          Set objForm = Nothing
'''        End If
'''      End If
'''      '
'''      IMP_COMP_VENDA lngGRID, gsNomeEmpresa
'''      '----- Imprimir Impressora Fiscal
'''
'''      If optCobranca(0).Value Then '= "S" Then  'Apenas Imprime Venda Cobradas
'''        If gbTrabComImpFiscal Then
'''          If blnImprimirCupomFiscal = True And intQuemChamou = 0 Then 'imprime cupom fiscal apenas na inclusão da venda
'''            IMP_CUPOM_FISCAL_VENDA lngGRID, gsNomeEmpresa
'''          End If
'''        End If
'''      End If
'''    End If
'''  Else
'''    If MsgBox("Atenção: Não foram lançados itens nesta venda. Caso confirme, a venda não será impressa. Deseja continuar ?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      objRs.Close
'''      Set objRs = Nothing
'''      blnImprimirCupomFiscal = True
'''      blnFechar = False
'''      Exit Sub
'''    End If
'''  End If
'''  '
'''  objRs.Close
'''  Set objRs = Nothing
'''  '
'''  Set objVenda = Nothing
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdOk_Click()
  Dim objGR                       As busSisMed.clsGR
  Dim objGeral                    As busSisMed.clsGeral
  Dim objRs                       As ADODB.Recordset
  Dim strSql                      As String
  '
  Dim lngPRONTUARIOID     As Long
  Dim lngTURNOID          As Long
  Dim lngATENDEID         As Long
  Dim lngPRESTADORID      As Long
  Dim lngESPECIALIDADEID  As Long
  Dim lngSALAID           As Long
  Dim lngDIASDASEMANAID   As Long
  Dim lngSequencial       As Long
  Dim lngSequencialSenha  As Long
  Dim strHoraIni          As String
  Dim strHoraFim          As String
  Dim strCortesia         As String
  Dim strUsuLib           As String
  '
  On Error GoTo trata
  'cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisMed.clsGeral
  Set objGR = New busSisMed.clsGR
  'PRONTUARIO
  lngPRONTUARIOID = 0
  strSql = "SELECT PKID FROM PRONTUARIO " & _
        " WHERE NOME = " & Formata_Dados(txtProntuarioFim.Text, tpDados_Texto) & _
        " AND DTNASCIMENTO = " & Formata_Dados(mskDataNascFim.Text, tpDados_DataHora)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngPRONTUARIOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  strCortesia = "N"
  strUsuLib = ""
  If Status = tpStatus_Incluir Then
    'FUNCIONÁRIO
    strSql = "SELECT * FROM FUNCIONARIO " & _
          " WHERE PRONTUARIOID = " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      'NOVO TRATAMENTO DE CORTESIA PARA FUNCIONÁRIO
      If MsgBox("Prontuário é um funcionário. Deseja lançar cortesia para o funcionário " & txtProntuarioFim.Text & "?", vbYesNo) = vbYes Then
        'DESEJA INCLUIR CORTESIA
        'Pedir senha superior
        '----------------------------
        '----------------------------
        'Pede Senha Superior (Diretor, Gerente ou Administrador
        If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
          'Só pede senha superior se quem estiver logado não for superior
          frmUserLoginSup.Show vbModal
  
          If Len(Trim(gsNomeUsuLib)) = 0 Then
            Pintar_Controle txtPrestEspec, tpCorContr_Erro
            TratarErroPrevisto "Para lançar uma cortesia para funcionário é necessário a Confirmação com senha superior."
            Set objGeral = Nothing
            Set objGR = Nothing
            cmdOk.Enabled = True
            SetarFoco txtPrestEspec
            Exit Sub
          End If
          '
          'Capturou Nome do Usuário, continua processo de Sangria
        Else
          gsNomeUsuLib = gsNomeUsu
        End If
        strUsuLib = gsNomeUsuLib
        strCortesia = "S"
        '--------------------------------
        '--------------------------------
      Else
        strCortesia = "N"
        strUsuLib = ""
      End If
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  'TURNO
  lngTURNOID = RetornaCodTurnoCorrente
  'Hora Inicial
  strHoraIni = Left(txtPeriodo, 5)
  'Hora Final
  strHoraFim = Right(txtPeriodo, 5)
  'PRESTADOR
  lngPRESTADORID = 0
  strSql = "SELECT PRONTUARIO.PKID FROM PRONTUARIO INNER JOIN PRESTADOR ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID " & _
    "WHERE PRONTUARIO.NOME = " & Formata_Dados(txtPrestador.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngPRESTADORID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'ESPECIALIDADE
  lngESPECIALIDADEID = 0
  strSql = "SELECT ESPECIALIDADE.PKID FROM ESPECIALIDADE " & _
    "WHERE ESPECIALIDADE.ESPECIALIDADE = " & Formata_Dados(txtEspecialidade.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngESPECIALIDADEID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'SALA
  lngSALAID = 0
  strSql = "SELECT SALA.PKID FROM SALA " & _
    "WHERE SALA.NUMERO = " & Formata_Dados(txtSala.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngSALAID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'DIA DA SEMANA
  lngDIASDASEMANAID = Retorna_DIADASEMANA_Data(Now)
  'ATENDE
  lngATENDEID = 0
  strSql = "SELECT ATENDE.PKID FROM ATENDE " & _
      "WHERE PRONTUARIOID = " & Formata_Dados(lngPRESTADORID, tpDados_Longo) & _
      " AND DIASDASEMANAID = " & Formata_Dados(lngDIASDASEMANAID, tpDados_Longo) & _
      " AND SALAID = " & Formata_Dados(lngSALAID, tpDados_Longo) & _
      " AND HORAINICIO = " & Formata_Dados(strHoraIni, tpDados_Texto) & _
      " AND HORATERMINO = " & Formata_Dados(strHoraFim, tpDados_Texto) & _
      " AND ATENDE.STATUS = " & Formata_Dados("A", tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngATENDEID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  
  If lngATENDEID = 0 Or lngPRONTUARIOID = 0 Then
    Pintar_Controle txtPrestEspec, tpCorContr_Erro
    TratarErroPrevisto "Prontuario / Sala não cadastrado"
    Set objGeral = Nothing
    Set objGR = Nothing
    cmdOk.Enabled = True
    SetarFoco txtPrestEspec
    Exit Sub
  End If
  If Status = tpStatus_Alterar Then
    'Alterar GR
    objGR.AlterarGR lngGRID, _
                    lngPRONTUARIOID, _
                    lngATENDEID, _
                    lngESPECIALIDADEID, _
                    txtDescricao.Text
    'Verifica MOV
    VerificaMovAposFecha lngGRID
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir GR
    'Sequencial
    lngSequencial = RetornaGravaCampoSequencial("SEQUENCIAL", lngTURNOID)
    lngSequencialSenha = RetornaGravaCampoSequencialSenha("SEQUENCIAL", lngPRESTADORID)
    objGR.InserirGR lngGRID, _
                    lngPRONTUARIOID, _
                    lngTURNOID, _
                    IIf(gsNivel <> gsLaboratorio, "", lngTURNOID & ""), _
                    lngATENDEID, _
                    lngESPECIALIDADEID, _
                    lngSequencial & "", _
                    lngSequencialSenha & "", _
                    Format(Now, "DD/MM/YYYY hh:mm"), _
                    "I", _
                    "N", _
                    giFuncionarioId, _
                    strCortesia, _
                    strUsuLib, _
                    txtDescricao.Text

  End If
  'Verificação
  If Status = tpStatus_Alterar Then
    'Selecionar prontuario pelo nome
    Status = tpStatus_Alterar
    IcEstadoGR = tpIcEstadoGR_Proc
    'Reload na tela
    Form_Load
    'Acerta tabs
    blnRetorno = True
  ElseIf Status = tpStatus_Incluir Then
    'Selecionar prontuario pelo nome
    Status = tpStatus_Alterar
    IcEstadoGR = tpIcEstadoGR_Proc
    'Reload na tela
    Form_Load
    'Acerta tabs
    blnRetorno = True
  End If
  'cmdOk.Enabled = True
  SetarFoco txtProcedimento
  Set objGR = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  cmdOk.Enabled = True
End Sub

Private Sub mskQuantidade_GotFocus()
  Selecionar_Conteudo mskQuantidade
End Sub

Private Sub mskQuantidade_KeyUp(KeyCode As Integer, Shift As Integer)
  On Error GoTo trata
  Dim objUserProcedimentoCons As SisMed.frmUserProcedimentoCons
  Dim objGR                   As busSisMed.clsGR
  Dim objGeral                As busSisMed.clsGeral
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim lngPROCEDIMENTOID       As Long
  Dim curVALORPROC            As Currency
  Dim curVALORFINAL           As Currency
  Dim curVALORCORT            As Currency
  Dim curPERCCASA             As Currency
  
  '
  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub
  If KeyCode = 13 Then Exit Sub
  '
  Pintar_Controle txtProcedimento, tpCorContr_Normal
  If Not ValidaCamposProcedimento Then
    Exit Sub
  End If
  
  Set objGeral = New busSisMed.clsGeral
  '
  'PROCEDIMENTO
  lngPROCEDIMENTOID = 0
  strSql = "SELECT PROCEDIMENTO.PKID, PROCEDIMENTO.VALOR FROM PROCEDIMENTO " & _
    "WHERE PROCEDIMENTO.PROCEDIMENTO = " & Formata_Dados(txtProcedimentoFim.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngPROCEDIMENTOID = objRs.Fields("PKID").Value
    curVALORPROC = IIf(IsNull(objRs.Fields("VALOR").Value), 0, objRs.Fields("VALOR").Value)
  End If
  objRs.Close
  Set objRs = Nothing
  '
  If lngPROCEDIMENTOID = 0 Then
    TratarErroPrevisto "Procedimento não cadastrado."
    Pintar_Controle txtProcedimento, tpCorContr_Erro
    SetarFoco txtProcedimento
    Exit Sub
  End If
  '
  'PERCENTUAL DA CASA PARA CORTESIA
  curPERCCASA = 0
  curVALORCORT = 0
  If strCortesiaGR = "S" Then
    strSql = "SELECT ISNULL(PRESTADORPROCEDIMENTO.PERCCASA, 0) AS PERCCASA FROM GR " & _
              "INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
              "INNER JOIN PRESTADORPROCEDIMENTO ON PRESTADORPROCEDIMENTO.PRONTUARIOID = ATENDE.PRONTUARIOID " & _
              "WHERE GR.PKID = " & Formata_Dados(lngGRID, tpDados_Longo) & _
              " AND PRESTADORPROCEDIMENTO.PROCEDIMENTOID = " & Formata_Dados(lngPROCEDIMENTOID, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      curPERCCASA = objRs.Fields("PERCCASA").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
  End If
  Set objGeral = Nothing
  'Calculo do Valor a ser cobrado
  curVALORFINAL = curVALORPROC * (CLng(mskQuantidade.Text))
  If strCortesiaGR = "S" Then
    curVALORCORT = (curPERCCASA * curVALORFINAL / 100)
    curVALORFINAL = curVALORFINAL - (curPERCCASA * curVALORFINAL / 100)
    '----------------------------------------------
    'NOVO
    curVALORCORT = curVALORCORT + curVALORFINAL
    curVALORFINAL = 0
  End If
  If mskValor.Enabled = True Then
    curVALORFINAL = CCur(IIf(Not IsNumeric(mskValor.Text), 0, mskValor.Text)) * (CLng(mskQuantidade.Text))
  End If
  
  'Inclusão de proceidmentos
  Set objGR = New busSisMed.clsGR
  '
  objGR.InserirGRPROCEDIMENTO lngGRID, _
                              lngPROCEDIMENTOID, _
                              mskQuantidade.Text, _
                              Format(curVALORFINAL, "###,##0.00") & "", _
                              Format(curVALORCORT, "###,##0.00") & ""
  '
  Set objGR = Nothing
  VerificaMovAposFecha lngGRID
  'cmdOk.Default = True
  'Novo procedimento
  txtProcedimento.Text = ""
  txtProcedimentoFim.Text = ""
  LimparCampoMask mskQuantidade
  LimparCampoMask mskValor
  TratarAcValorProco "N", _
                     mskValor, _
                     mskQuantidade, _
                     False
  '
  'Montar RecordSet
  PROC_COLUNASMATRIZ = grdProcedimento.Columns.Count
  PROC_LINHASMATRIZ = 0
  PROC_MontaMatriz
  grdProcedimento.Bookmark = Null
  grdProcedimento.ReBind
  SetarFoco txtProcedimento
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub mskQuantidade_LostFocus()
  Pintar_Controle mskQuantidade, tpCorContr_Normal
End Sub

Private Sub txtDescricao_GotFocus()
  Selecionar_Conteudo txtDescricao
End Sub

Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub

Private Sub txtProcedimento_GotFocus()
  Selecionar_Conteudo txtProcedimento
End Sub



'''
'''Private Function ValidaCamposItens() As Boolean
'''  Dim strMsg     As String
'''
'''  If Not IsNumeric(mskQuantidade.Text) Then
'''    strMsg = strMsg & "A Quantidade do produto é inválida" & vbCrLf
'''    Pintar_Controle mskQuantidade, tpCorContr_Erro
'''  ElseIf CLng(mskQuantidade.Text) = 0 Then
'''    strMsg = strMsg & "A Quantidade do produto não pode ser igual a zero" & vbCrLf
'''    Pintar_Controle mskQuantidade, tpCorContr_Erro
'''  End If
'''  '
'''  If Len(mskCodigo.ClipText) <> 6 Then
'''    strMsg = strMsg & "Digitar o tamanho do código do produto com 5 dígitos" & vbCrLf
'''    Pintar_Controle mskCodigo, tpCorContr_Erro
'''  End If
'''  '
'''  If Not IsNumeric(Right(mskCodigo.ClipText, 4)) Then
'''    strMsg = strMsg & "Digitar os últimos 4 dígitos do produto com valor númerico" & vbCrLf
'''    Pintar_Controle mskCodigo, tpCorContr_Erro
'''  End If
'''  '
'''  If Len(strMsg) <> 0 Then
'''    TratarErroPrevisto strMsg, "[frmUserVendaInc.ValidaCampos]"
'''    ValidaCamposItens = False
'''  Else
'''    ValidaCamposItens = True
'''  End If
'''End Function
'''
'''Private Function ValidaCamposVenda() As Boolean
'''  Dim strMsg     As String
'''  '
'''  If Len(txtPrestEspec.Text) = 0 Then
'''    strMsg = strMsg & "Informar o Nome" & vbCrLf
'''    Pintar_Controle txtPrestEspec, tpCorContr_Erro
'''  End If
'''  '
'''  If Len(txtFuncao.Text) = 0 Then
'''    strMsg = strMsg & "Informar a Função" & vbCrLf
'''    Pintar_Controle txtFuncao, tpCorContr_Erro
'''  End If
'''  '
'''  If Not optVenda(0).Value And Not optVenda(1).Value Then
'''    strMsg = strMsg & "Selecionar a venda" & vbCrLf
'''  End If
'''  '
'''  If Not optCobranca(0).Value And Not optCobranca(1).Value Then
'''    strMsg = strMsg & "Selecionar a cobrança" & vbCrLf
'''  End If
'''  '
'''  If Len(cboConfiguracao.Text) = 0 Then
'''    strMsg = strMsg & "Selecionar a configuração" & vbCrLf
'''    Pintar_Controle cboConfiguracao, tpCorContr_Erro
'''  End If
'''  If Len(strMsg) <> 0 Then
'''    TratarErroPrevisto strMsg, "[frmUserVendaInc.ValidaCamposVenda]"
'''    ValidaCamposVenda = False
'''  Else
'''    ValidaCamposVenda = True
'''  End If
'''End Function
'''
'''Private Sub cmdPedido_Click()
'''  On Error GoTo trata
'''
'''  frmUserCardapioCons.QuemChamou = 1
'''  frmUserCardapioCons.sCobranca = IIf(optCobranca(0).Value, "S", "N")
'''  frmUserCardapioCons.Show vbModal
'''
''''  frmConsItemCard.QuemChamou = 1
''''  frmConsItemCard.sCobranca = IIf(optCobranca(0).Value, "S", "N")
''''  frmConsItemCard.Show vbModal
'''  If Len(mskQuantidade.ClipText) = 0 Then
'''    SetarFoco mskQuantidade
'''  ElseIf Len(txtCodigo.Text) = 0 Then
'''    SetarFoco txtCodigo
'''  Else
'''    SetarFoco txtDescricaoItem
'''  End If
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub
'''
Private Sub Form_Activate()
  Dim datDataTurno      As Date
  Dim datDataIniAtual   As Date
  Dim datDataFimAtual   As Date
  On Error GoTo trata
  If blnPrimeiraVez Then
    '
    If Status = tpStatus_Incluir Then
      If RetornaCodTurnoCorrente(datDataTurno) = 0 Then
        TratarErroPrevisto "Não há turnos em aberto, favor abrir um turno antes de incluir as GR´s", "Form_Load"
        Unload Me
      Else
        'OK Para turno
        datDataIniAtual = CDate(Day(Now) & "/" & Month(Now) & "/" & Year(Now) & " 00:00:00")
        datDataFimAtual = DateAdd("d", 1, datDataIniAtual)
        If datDataTurno < datDataIniAtual Or datDataTurno >= datDataFimAtual Then
          TratarErroPrevisto "ATENÇÃO" & vbCrLf & vbCrLf & "A data do turno atual aberto não corresponde a data de hoje:" & vbCrLf & vbCrLf & "Data do turno --> " & Format(datDataTurno, "DD/MM/YYYY") & vbCrLf & "Data Atual --> " & Format(datDataIniAtual, "DD/MM/YYYY") & vbCrLf & vbCrLf & "Por favor, feche o turno e abra-o novamente para lançar as GR´s.", "Form_Load"
          Unload Me
        End If
      End If
    End If
    'Seta foco no grid
    'Montar RecordSet
    PROC_COLUNASMATRIZ = grdProcedimento.Columns.Count
    PROC_LINHASMATRIZ = 0
    PROC_MontaMatriz
    grdProcedimento.Bookmark = Null
    grdProcedimento.ReBind
    '
    tabDetalhes.Tab = 0
    blnPrimeiraVez = False
    SetarFoco txtPrestEspec
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserVendaInc.Form_Activate]"
End Sub

'''
'''
'''Private Sub txtProcedimento_GotFocus()
'''  cmdOk.Default = False
'''  Selecionar_Conteudo txtProcedimento
'''End Sub
'''



Private Sub txtPrestEspec_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtProcedimento_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtProcedimento_LostFocus()
  On Error GoTo trata
  Dim objUserProcedimentoCons As SisMed.frmUserProcedimentoCons
  Dim objGR                   As busSisMed.clsGR
  Dim objGeral                As busSisMed.clsGeral
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim lngPRESTADORID          As Long
  '
  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub
  If Me.ActiveControl.Name = "grdProcedimento" Then Exit Sub
  If Me.ActiveControl.Name = "cmdPagamento" Then Exit Sub
  If Me.ActiveControl.Name = "cmdExcluir" Then Exit Sub
  If Me.ActiveControl.Name = "cmdImprimir" Then Exit Sub
  '
  Pintar_Controle txtProcedimento, tpCorContr_Normal
  If Len(txtProcedimento.Text) = 0 Then
    TratarErroPrevisto "Entre com o procedimento."
    Pintar_Controle txtProcedimento, tpCorContr_Erro
    SetarFoco txtProcedimento
    Exit Sub
  End If
  Set objGR = New busSisMed.clsGR
  Set objGeral = New busSisMed.clsGeral
  '
  'PRESTADOR
  lngPRESTADORID = 0
  strSql = "SELECT PRONTUARIO.PKID FROM PRONTUARIO INNER JOIN PRESTADOR ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID " & _
    "WHERE PRONTUARIO.NOME = " & Formata_Dados(txtPrestador.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngPRESTADORID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  '
  If lngPRESTADORID = 0 Then
    TratarErroPrevisto "Prestador não cadastrado."
    Pintar_Controle txtProcedimento, tpCorContr_Erro
    SetarFoco txtProcedimento
    Exit Sub
  End If
  
  Set objRs = objGR.CapturaProcedimento(txtProcedimento.Text, _
                                        lngPRESTADORID)
  If objRs.EOF Then
    'Novo : apresentar tela para seleção do prontuário
    Set objUserProcedimentoCons = New SisMed.frmUserProcedimentoCons
    objUserProcedimentoCons.strProcedimento = txtProcedimento.Text
    objUserProcedimentoCons.lngPRESTADORID = lngPRESTADORID
    objUserProcedimentoCons.indOrigem = 0
    objUserProcedimentoCons.Show vbModal

    If objUserProcedimentoCons.strProcedimento = "" Then
      txtProcedimento.Text = ""
      txtProcedimentoFim.Text = ""
      LimparCampoMask mskValor
      TratarErroPrevisto "Selecione um procedimento"
      TratarAcValorProco "N", _
                         mskValor, _
                         mskQuantidade, _
                         False
      Pintar_Controle txtProcedimento, tpCorContr_Erro
      SetarFoco txtProcedimento
      Exit Sub
    Else
      'SetarFoco mskQuantidade
      TratarAcValorProco strAceitaValor, _
                         mskValor, _
                         mskQuantidade, _
                         True
    End If
    Set objUserProcedimentoCons = Nothing
  Else
    If objRs.RecordCount = 1 Then
      txtProcedimentoFim = objRs.Fields("PROCEDIMENTO").Value & ""
      TratarAcValorProco objRs.Fields("INDACEITAVALOR").Value & "", _
                         mskValor, _
                         mskQuantidade
    Else
      'Novo : apresentar tela para seleção do prontuário
      Set objUserProcedimentoCons = New frmUserProcedimentoCons
      objUserProcedimentoCons.strProcedimento = txtProcedimento.Text
      objUserProcedimentoCons.lngPRESTADORID = lngPRESTADORID
      objUserProcedimentoCons.indOrigem = 0
      objUserProcedimentoCons.Show vbModal

      If objUserProcedimentoCons.strProcedimento = "" Then
        txtProcedimento.Text = ""
        txtProcedimentoFim.Text = ""
        LimparCampoMask mskValor
        TratarErroPrevisto "Selecione um procedimento"
        TratarAcValorProco "N", _
                           mskValor, _
                           mskQuantidade, _
                           False
        Pintar_Controle txtProcedimento, tpCorContr_Erro
        SetarFoco txtProcedimento
        Exit Sub
      Else
        'SetarFoco mskQuantidade
        TratarAcValorProco strAceitaValor, _
                           mskValor, _
                           mskQuantidade, _
                           True
        'Tratar Valor
      End If
      Set objUserProcedimentoCons = Nothing
    End If
  End If
  '
  objRs.Close
  Set objRs = Nothing
  Set objGR = Nothing
  'cmdOk.Default = True
  
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

'''
'''Private Sub mskCodigo_Change()
''''''  On Error GoTo trata
''''''  Dim clsCard   As busSisMed.clsCardapio
''''''  Dim objRs     As ADODB.Recordset
''''''  If Len(mskCodigo.ClipText) <> 6 Then
''''''    txtProduto.Text = ""
''''''    Exit Sub
''''''  End If
''''''  Set clsCard = New busSisMed.clsCardapio
''''''  '
''''''  Set objRs = clsCard.CapturaItemCardapio(mskCodigo.ClipText)
''''''  If objRs.EOF Then
''''''    txtProduto.Text = ""
''''''  Else
''''''    txtProduto.Text = objRs.Fields("DESCRICAO").Value
''''''  End If
''''''  '
''''''  objRs.Close
''''''  Set objRs = Nothing
''''''  Set clsCard = Nothing
''''''  Exit Sub
''''''trata:
''''''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub
'''
'''Private Sub mskCodigo_KeyPress(KeyAscii As Integer)
'''  On Error GoTo trata
'''  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub
'''
'''Private Sub mskCodigo_LostFocus()
'''  On Error GoTo trata
'''  mskCodigo.Text = UCase(mskCodigo.Text)
'''  Pintar_Controle mskCodigo, tpCorContr_Normal
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub
'''
Private Sub Form_Load()
On Error GoTo trata
  Dim objRs     As ADODB.Recordset
  Dim strSql    As String
'''  Dim clsLoc    As busSisMed.clsLocacao
'''  Dim clsvend   As busSisMed.clsVenda
  Dim objGR As busSisMed.clsGR
  blnPrimeiraVez = True
  blnFechar = False
  blnRetorno = False
  strCortesiaGR = "N"
'''  blnImprimirCupomFiscal = True
'''  blnMostrarAlertaProd = True
'''  blnAlterouItem = False
  AmpS
  Me.Height = 8385
  Me.Width = 10470
  CenterForm Me
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , , , , cmdImprimir
  LerFigurasAvulsas cmdPagamento, "InfFinanc.ico", "InfFinancDown.ico", "Recebimento"
  '
  'Limpar Campos
  LimparCampos
  'Tratar campos
  TratarCampos
  '
'''  If Not gbTrabComVendasExt Then
'''    optVenda(1).Enabled = False
'''  End If
  'Configuracao
'''  tabDetalhes_Click 0
  If Status = tpStatus_Incluir Then
    '
    lblCor(0).Visible = False
    lblCor(1).Visible = False
    lblCor(2).Visible = False
    lblCor(3).Visible = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    '-----------------------------
    'GR
    '------------------------------
    lblCor(0).Visible = True
    lblCor(1).Visible = True
    lblCor(2).Visible = True
    lblCor(3).Visible = True
    '
    Set objGR = New busSisMed.clsGR
    Set objRs = objGR.SelecionarGRPeloPkid(lngGRID)
    '
    If Not objRs.EOF Then
      'GR
      strCortesiaGR = objRs.Fields("INDCORTESIA").Value & ""
      strStatusGR = objRs.Fields("STATUS").Value & ""
      'Cabeçalho
      txtTurno.Text = RetornaDescTurnoCorrente(objRs.Fields("TURNOID").Value & "")
      txtDiaDaSemana.Text = Retorna_DIADASEMANA_Descr(objRs.Fields("DATA").Value)
      txtSequencial = objRs.Fields("SEQUENCIAL").Value & ""
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DATA").Value, TpMaskData
      txtCaixa.Text = Retorna_CAIXA_Nome(giFuncionarioId)
      'GR
      txtPrestEspec.Text = objRs.Fields("PREST_PRESTADOR").Value & ""
      txtPrestador.Text = objRs.Fields("PREST_PRESTADOR").Value & ""
      txtEspecialidade.Text = objRs.Fields("ESPEC_ESPECIALIDADE").Value & ""
      txtSala.Text = objRs.Fields("SALA_NUMERO").Value & ""
      txtPeriodo.Text = objRs.Fields("PERIODO_PERIODO").Value & ""
      txtProntuario.Text = objRs.Fields("NOME").Value & ""
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      txtProntuarioFim.Text = objRs.Fields("NOME").Value & ""
      INCLUIR_VALOR_NO_MASK mskDataNascFim, Format(objRs.Fields("DTNASCIMENTO").Value, "DD/MM/YYYY"), TpMaskData
      '
      TratarStatus objRs.Fields("STATUS").Value & "", _
                   objRs.Fields("STATUSIMPRESSAO").Value & "", _
                   lblCor(1), _
                   lblCor(3)
    End If
    objRs.Close
    Set objRs = Nothing
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
'  If Not blnFechar Then Cancel = True
End Sub



'''Private Sub tabDetalhes_Click(PreviousTab As Integer)
'''  Dim strMsgErro    As String
'''  Dim strCobranca   As String
'''  '
'''  On Error GoTo trata
'''  Select Case tabDetalhes.Tab
'''  Case 0
'''    'dados principais da venda
'''    blnMostrarAlertaProd = False
'''    If Status = tpStatus_Incluir Then
'''      picTrava(0).Enabled = True
'''    Else
'''      picTrava(0).Enabled = False
'''    End If
'''    Frame1.Enabled = False
'''    grdProcedimento.Enabled = False
'''    '
'''    If intQuemChamou = 1 Then 'Alteração
'''      cmdOk.Enabled = False
'''      cmdCancelar.Enabled = True
'''      cmdPedido.Enabled = False
'''      cmdExcluir.Enabled = False
'''      cmdImprimir.Enabled = False
'''    Else
'''      cmdOk.Enabled = True
'''      cmdCancelar.Enabled = True
'''      cmdPedido.Enabled = False
'''      cmdExcluir.Enabled = False
'''      cmdImprimir.Enabled = False
'''    End If
'''    SetarFoco txtNome
'''  Case 1
'''    blnMostrarAlertaProd = True
'''    picTrava(0).Enabled = False
'''    Frame1.Enabled = True
'''    grdProcedimento.Enabled = False
'''    '
'''    If optCobranca(0).Value Then
'''      strCobranca = "S"
'''    Else
'''      strCobranca = "N"
'''    End If
'''    If strCobranca = "N" Then 'Só pede senha para vendas não cobradas
'''      If gbPedirSenhaVdaDiretoria Then
'''        'Pedir senha superior
'''        '----------------------------
'''        '----------------------------
'''        'Pede Senha Superior (Diretor, Gerente ou Administrador
'''        If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''          'Só pede senha superior se quem estiver logado não for superior
'''          frmUserLoginSup.Show vbModal
'''
'''          If Len(Trim(gsNomeUsuLib)) = 0 Then
'''            strMsgErro = "Para efetuar uma venda interna não cobrada é necessário a Confirmação com senha superior."
'''            MsgBox strMsgErro, vbExclamation, TITULOSISTEMA
'''            tabDetalhes.Tab = 0
'''            Exit Sub
'''          End If
'''          '
'''          'Capturou Nome do Usuário, continua processo de Sangria
'''        Else
'''          gsNomeUsuLib = gsNomeUsu
'''        End If
'''        '--------------------------------
'''        '--------------------------------
'''      End If
'''    End If
'''    'Inclusão de Iten do Pedido
'''    cmdCancelar.Enabled = True
'''    cmdExcluir.Enabled = False
'''    'If intQuemChamou = 1 Then
'''    If Status = tpStatus_Consultar Then
'''      cmdImprimir.Enabled = False
'''      cmdPedido.Enabled = False
'''      cmdOk.Enabled = False
'''    Else
'''      cmdImprimir.Enabled = False
'''      cmdPedido.Enabled = True
'''      cmdOk.Enabled = True
'''    End If
'''    '
'''    SetarFoco txtProcedimento
'''  Case 2
'''    blnMostrarAlertaProd = False
'''    picTrava(0).Enabled = False
'''    Frame1.Enabled = False
'''    grdProcedimento.Enabled = True
'''    '
'''    'visualização dos Itens do pedido
'''    cmdOk.Enabled = False
'''    cmdCancelar.Enabled = True
'''    cmdPedido.Enabled = False
'''    'If intQuemChamou = 1 Then 'Alteração
'''    If Status = tpStatus_Consultar Then
'''      cmdImprimir.Enabled = False
'''      cmdExcluir.Enabled = False
'''    Else 'inclusão
'''      cmdImprimir.Enabled = False
'''      cmdExcluir.Enabled = True
'''    End If
'''    'Montar RecordSet
'''    PROC_COLUNASMATRIZ = 8
'''    PROC_LINHASMATRIZ = 0
'''    PROC_MontaMatriz
'''    grdProcedimento.Bookmark = Null
'''    grdProcedimento.ReBind
'''    SetarFoco grdProcedimento
'''  End Select
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, "SisMotel.frmUserVendaInc.tabDetalhes"
'''  AmpN
'''End Sub
'''
'''
'''
Private Sub txtPrestEspec_GotFocus()
  Selecionar_Conteudo txtPrestEspec
End Sub



Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'GR
  LimparCampoTexto txtDiaDaSemana
  LimparCampoTexto txtTurno
  LimparCampoTexto txtCaixa
  LimparCampoTexto txtSequencial
  LimparCampoMask mskData(0)
  '
  LimparCampoTexto txtPrestEspec
  LimparCampoTexto txtPrestador
  LimparCampoTexto txtEspecialidade
  LimparCampoTexto txtSala
  LimparCampoTexto txtPeriodo
  '
  LimparCampoTexto txtProntuario
  LimparCampoTexto txtDescricao
  LimparCampoTexto txtProntuarioFim
  LimparCampoMask mskDataNascFim
      
  'GRPRESTADOR
  LimparCampoTexto txtProcedimento
  LimparCampoTexto txtProcedimentoFim
  LimparCampoMask mskQuantidade
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserGRInc.LimparCampos]", _
            Err.Description
End Sub

Private Sub TratarEstadoGR()
  On Error GoTo trata
  'Propósito : Tratar estado da GR
  If IcEstadoGR = tpIcEstadoGR_Inic Then
    picTrava(1).Enabled = True
    picTrava(2).Enabled = False
    picTrava(3).Enabled = False
    grdProcedimento.Enabled = False
    '
    cmdImprimir.Enabled = False
    cmdExcluir.Enabled = False
    cmdPagamento.Enabled = False
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    blnPrimeiraVez = True
  ElseIf IcEstadoGR = tpIcEstadoGR_Proc Then
    picTrava(1).Enabled = False
    picTrava(2).Enabled = True
    picTrava(3).Enabled = False
    grdProcedimento.Enabled = True
    '
    cmdImprimir.Enabled = True
    cmdExcluir.Enabled = True
    If gsNivel = gsLaboratorio Then
      cmdPagamento.Enabled = False
    Else
      cmdPagamento.Enabled = True
    End If
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
  ElseIf IcEstadoGR = tpIcEstadoGR_Con Then
    picTrava(1).Enabled = False
    picTrava(2).Enabled = False
    picTrava(3).Enabled = False
    grdProcedimento.Enabled = True
    '
    cmdImprimir.Enabled = False
    cmdExcluir.Enabled = False
    If gsNivel = gsLaboratorio Then
      cmdPagamento.Enabled = False
    Else
      cmdPagamento.Enabled = True
    End If
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
  End If
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserGRInc.TratarEstadoGR]", _
            Err.Description
End Sub


Private Sub TratarCampos()
  On Error GoTo trata
  'GRPRESTADOR
  '
  txtTurno.Text = RetornaDescTurnoCorrente
  '
  TratarEstadoGR
  '
  If Status = tpStatus_Incluir Then
    'Trtar exclusão
    '
    txtTurno.Text = RetornaDescTurnoCorrente
    txtDiaDaSemana.Text = Retorna_DIADASEMANA_Descr(Date)
    txtCaixa.Text = Retorna_CAIXA_Nome(giFuncionarioId)
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Visible
  End If
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserGRInc.TratarCampos]", _
            Err.Description
End Sub


Private Sub txtPrestEspec_LostFocus()
  On Error GoTo trata
  Dim objUserPrestEspecCons As SisMed.frmUserPrestEspecCons
  Dim objGR     As busSisMed.clsGR
  Dim objRs     As ADODB.Recordset
  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub
  'If Me.ActiveControl.Name = "txtProntuario" Then Exit Sub
  'If Me.ActiveControl.Name = "txtProcedimento" Then Exit Sub
  
  
  Pintar_Controle txtPrestEspec, tpCorContr_Normal
  If Len(txtPrestEspec.Text) = 0 Then
    TratarErroPrevisto "Entre com o prestador ou especialidade."
    Pintar_Controle txtPrestEspec, tpCorContr_Erro
    SetarFoco txtPrestEspec
    Exit Sub
  End If
  Set objGR = New busSisMed.clsGR
  '
  Set objRs = objGR.CapturaPrestEspec(txtPrestEspec.Text, _
                                      Retorna_DIADASEMANA_Descr(Date))
  If objRs.EOF Then
    LimparCampoTexto txtPrestador
    LimparCampoTexto txtEspecialidade
    LimparCampoTexto txtSala
    LimparCampoTexto txtPeriodo
    TratarErroPrevisto "Prestador/Especialidade não cadastrado"
    Pintar_Controle txtPrestEspec, tpCorContr_Erro
    SetarFoco txtPrestEspec
    Exit Sub
  Else
    If objRs.RecordCount = 1 Then
      txtPrestador = objRs.Fields("PREST_PRESTADOR").Value & ""
      txtEspecialidade = objRs.Fields("ESPEC_ESPECIALIDADE").Value & ""
      txtSala = objRs.Fields("SALA_NUMERO").Value & ""
      txtPeriodo = objRs.Fields("PERIODO_PERIODO").Value & ""
    Else
      'Novo : apresentar tela para seleção do produto
      Set objUserPrestEspecCons = New SisMed.frmUserPrestEspecCons
      objUserPrestEspecCons.strPrestadorEspecialidade = txtPrestEspec.Text
      objUserPrestEspecCons.Show vbModal

      If objUserPrestEspecCons.strPrestadorEspecialidade = "" Then
        txtPrestador.Text = ""
        txtEspecialidade.Text = ""
        txtSala.Text = ""
        txtPeriodo.Text = ""
        TratarErroPrevisto "Selecione um prestador / especialidade"
        Pintar_Controle txtPrestEspec, tpCorContr_Erro
        SetarFoco txtPrestEspec
        Exit Sub
      Else
'''        If Len(mskQuantidade.ClipText) = 0 Then
'''          SetarFoco mskQuantidade
'''        ElseIf Len(mskCodigo.Text) = 0 Then
'''          SetarFoco txtPrestEspec
'''        Else
'''          SetarFoco txtDescricaoItem
'''        End If
      End If
      Set objUserPrestEspecCons = Nothing
    End If
  End If
  '
  objRs.Close
  Set objRs = Nothing
  Set objGR = Nothing
  'cmdOk.Default = True
  
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtProntuario_GotFocus()
  Selecionar_Conteudo txtProntuario
End Sub

Private Sub txtProntuario_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtProntuario_LostFocus()
  On Error GoTo trata
  Dim objUserProntuarioCons As SisMed.frmUserProntuarioCons
  Dim objGR     As busSisMed.clsGR
  Dim objRs     As ADODB.Recordset
  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub
  If Me.ActiveControl.Name = "txtPrestEspec" Then Exit Sub
  If Me.ActiveControl.Name = "txtProntuario" Then Exit Sub
  'If Me.ActiveControl.Name = "txtProcedimento" Then Exit Sub

  Pintar_Controle txtProntuario, tpCorContr_Normal
  If Len(txtProntuario.Text) = 0 Then
    TratarErroPrevisto "Entre com o prontuário."
    Pintar_Controle txtProntuario, tpCorContr_Erro
    SetarFoco txtProntuario
    Exit Sub
  End If
  Set objGR = New busSisMed.clsGR
  '
  Set objRs = objGR.CapturaProntuario(txtProntuario.Text, _
                                      "", _
                                      "")
  If objRs.EOF Then
    'Novo : apresentar tela para seleção do prontuário
    Set objUserProntuarioCons = New frmUserProntuarioCons
    objUserProntuarioCons.strNome = txtProntuario.Text
    objUserProntuarioCons.Show vbModal

    If objUserProntuarioCons.strNome = "" Then
      txtProntuario.Text = ""
      TratarErroPrevisto "Selecione um prontuário"
      Pintar_Controle txtProntuario, tpCorContr_Erro
      SetarFoco txtProntuario
      Exit Sub
    Else
      'Cadastrar GR
      CadastrarGR
    End If
    Set objUserProntuarioCons = Nothing
  Else
    If objRs.RecordCount = 1 Then
      txtProntuarioFim = objRs.Fields("NOME").Value & ""
      INCLUIR_VALOR_NO_MASK mskDataNascFim, objRs.Fields("DTNASCIMENTO").Value & "", TpMaskData
    Else
      'Novo : apresentar tela para seleção do prontuário
      Set objUserProntuarioCons = New frmUserProntuarioCons
      objUserProntuarioCons.strNome = txtProntuario.Text
      objUserProntuarioCons.Show vbModal

      If objUserProntuarioCons.strNome = "" Then
        txtProntuario.Text = ""
        txtProntuarioFim.Text = ""
        LimparCampoMask mskDataNascFim
        TratarErroPrevisto "Selecione um prontuário"
        Pintar_Controle txtProntuario, tpCorContr_Erro
        SetarFoco txtProntuario
        Exit Sub
      Else
        'Cadastrar GR
        CadastrarGR
      End If
      Set objUserProntuarioCons = Nothing
    End If
  End If
  '
  objRs.Close
  Set objRs = Nothing
  Set objGR = Nothing
  'cmdOk.Default = True
  
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Public Sub CadastrarGR()
  On Error GoTo trata
  '
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source

End Sub


Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  
  If Len(Trim(txtPrestEspec.Text & "")) = 0 Then
    SetarFoco txtPrestEspec
    Pintar_Controle txtPrestEspec, tpCorContr_Erro
    strMsg = strMsg & "Selecionar o prestador" & vbCrLf
    blnSetarFocoControle = False
  End If
  If Len(Trim(txtProntuarioFim.Text & "")) = 0 Then
    If blnSetarFocoControle = True Then
      SetarFoco txtProntuario
    End If
    Pintar_Controle txtProntuario, tpCorContr_Erro
    strMsg = strMsg & "Selecionar o prontuario" & vbCrLf
    blnSetarFocoControle = False
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserGRInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserGRInc.ValidaCampos]", _
            Err.Description
End Function

Private Function ValidaCamposProcedimento() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCamposProcedimento = False
  If Len(txtProcedimentoFim.Text) = 0 Then
    strMsg = strMsg & "Preencher o procedimento."
    Pintar_Controle txtProcedimento, tpCorContr_Erro
    SetarFoco txtProcedimento
  End If
  If mskValor.Enabled = True Then
    If Not Valida_Moeda(mskValor, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Pressncher o valor válido" & vbCrLf
    End If
  End If
  If Not Valida_Moeda(mskQuantidade, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Pressncher a quantidade válida" & vbCrLf
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserGRInc.ValidaCamposProcedimento]"
    ValidaCamposProcedimento = False
  Else
    ValidaCamposProcedimento = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserGRInc.ValidaCamposProcedimento]", _
            Err.Description
End Function

Private Sub cmdImprimir_Click()
  Dim objGR     As busSisMed.clsGR
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim strStatus As String
  Dim strNivel As String
  On Error GoTo trata
  'Imprimir GR
  Set objGR = New busSisMed.clsGR
  Set objRs = objGR.SelecionarGRPeloPkid(lngGRID)
  strStatus = ""
  If Not objRs.EOF Then
    strStatus = objRs.Fields("STATUS").Value & ""
    strNivel = objRs.Fields("NIVEL").Value & ""
  End If
  If strStatus <> "F" Then
    If strNivel <> gsLaboratorio Then
      MsgBox "Apenas poderá haver impressão de uma GR fechada ou lançada pelo Laboratório.", vbExclamation, TITULOSISTEMA
      SetarFoco grdProcedimento
      Exit Sub
    End If
  End If
  'Confirmação
  If MsgBox("Confirma impressão da GR " & txtSequencial.Text & " de " & txtProntuarioFim.Text & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdProcedimento
    Exit Sub
  End If
  
  IMP_COMP_GR lngGRID, gsNomeEmpresa, 1, False
  'Após impressão altera status para impressa
  Set objGR = New busSisMed.clsGR
  objGR.AlterarStatusGR lngGRID, _
                        "", _
                        "S"
  Set objGR = Nothing
  'Form_Load
  'SetarFoco grdProcedimento
  blnFechar = True
  blnRetorno = True
  Unload Me
  
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

