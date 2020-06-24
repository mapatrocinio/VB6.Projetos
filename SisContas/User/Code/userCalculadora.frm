VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmUserCalculadora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Calculo de troco"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4170
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4170
      Left            =   8520
      ScaleHeight     =   4170
      ScaleWidth      =   1860
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2145
         Left            =   90
         ScaleHeight     =   2085
         ScaleWidth      =   1605
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1800
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   150
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   3855
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Cálculo do troco"
      TabPicture(0)   =   "userCalculadora.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraProf"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraProf 
         Height          =   2895
         Left            =   120
         TabIndex        =   11
         Top             =   360
         Width           =   7935
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   2415
            Index           =   0
            Left            =   120
            ScaleHeight     =   2415
            ScaleWidth      =   7695
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   240
            Width           =   7695
            Begin VB.TextBox txtUnidade 
               BackColor       =   &H00C0C0C0&
               Height          =   285
               Left            =   1080
               Locked          =   -1  'True
               MaxLength       =   5
               TabIndex        =   0
               TabStop         =   0   'False
               Top             =   120
               Width           =   3495
            End
            Begin VB.Frame Frame3 
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
               Height          =   1815
               Left            =   0
               TabIndex        =   13
               Top             =   480
               Width           =   7695
               Begin VB.TextBox txtTroco 
                  Alignment       =   1  'Right Justify
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
                  Height          =   285
                  Left            =   4110
                  Locked          =   -1  'True
                  TabIndex        =   5
                  TabStop         =   0   'False
                  Text            =   "txtTroco"
                  ToolTipText     =   "DESCONTO EM %"
                  Top             =   1320
                  Width           =   1455
               End
               Begin VB.TextBox txtTotalaPagar 
                  Alignment       =   1  'Right Justify
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
                  Height          =   285
                  Left            =   1530
                  Locked          =   -1  'True
                  TabIndex        =   3
                  TabStop         =   0   'False
                  Text            =   "txtTotalaPagar"
                  Top             =   990
                  Width           =   1455
               End
               Begin VB.TextBox txtDesconto 
                  Alignment       =   1  'Right Justify
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
                  Height          =   285
                  Left            =   1530
                  Locked          =   -1  'True
                  TabIndex        =   2
                  TabStop         =   0   'False
                  Text            =   "txtDesconto"
                  ToolTipText     =   "DESCONTO EM %"
                  Top             =   630
                  Width           =   1455
               End
               Begin VB.TextBox txtTotalsDesc 
                  Alignment       =   1  'Right Justify
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
                  Height          =   285
                  Left            =   1530
                  Locked          =   -1  'True
                  TabIndex        =   1
                  TabStop         =   0   'False
                  Text            =   "txtTotalsDesc"
                  Top             =   270
                  Width           =   1455
               End
               Begin MSMask.MaskEdBox mskPgtoEspecie 
                  Height          =   255
                  Left            =   1530
                  TabIndex        =   4
                  Top             =   1320
                  Width           =   1575
                  _ExtentX        =   2778
                  _ExtentY        =   450
                  _Version        =   393216
                  Format          =   "#,##0.00;($#,##0.00)"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label53 
                  Caption         =   "Troco"
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
                  Left            =   3270
                  TabIndex        =   19
                  Top             =   1320
                  Width           =   975
               End
               Begin VB.Label Label38 
                  Caption         =   "Total a pagar"
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
                  TabIndex        =   18
                  Top             =   960
                  Width           =   1335
               End
               Begin VB.Label Label35 
                  Caption         =   "Desc."
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
                  Left            =   150
                  TabIndex        =   17
                  Top             =   600
                  Width           =   975
               End
               Begin VB.Label Label33 
                  Caption         =   "Total s/ desc"
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
                  TabIndex        =   16
                  Top             =   270
                  Width           =   1335
               End
               Begin VB.Label Label17 
                  Caption         =   "Pgto. Espécie"
                  Height          =   255
                  Left            =   150
                  TabIndex        =   15
                  Top             =   1320
                  Width           =   1455
               End
            End
            Begin VB.Label Label44 
               Caption         =   "Unidade"
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
               TabIndex        =   14
               Top             =   120
               Width           =   735
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserCalculadora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                 As tpStatus
Public lngSERVDESPID          As Long
Public lngLOCACAOID           As Long
Public bRetorno               As Boolean
Public bFechar                As Boolean
Public strNumeroSuiteApto     As String
Public intQuemChamou          As Integer
Private blnPrimeiraVez        As Boolean


Private Sub cmdCancelar_Click()
  '
  On Error GoTo trata
  '
  bRetorno = True
  bFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub



Private Sub cmdOk_Click()
  On Error GoTo trata
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração
    If Not ValidaCampos Then Exit Sub
    '
    If Status = tpStatus_Consultar Then
      'Código para alteração
      SetarFoco mskPgtoEspecie
      '
    End If
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg          As String
  Dim strSql          As String
  Dim vrEspecie       As Currency
  Dim vrPago          As Currency
  Dim vrTotLoc        As Currency
  Dim vrTotDescLoc    As Currency
  '
  If Not Valida_Moeda(mskPgtoEspecie, TpObrigatorio) Then
    strMsg = strMsg & "Preencher o valor válido" & vbCrLf
    Pintar_Controle mskPgtoEspecie, tpCorContr_Erro
  End If
  '
  If Len(strMsg) = 0 Then
    'Validar Valor
    
    vrEspecie = CCur(IIf(Not IsNumeric(mskPgtoEspecie.Text), 0, mskPgtoEspecie.Text))
    'Calcula Valor Pago
    vrPago = vrEspecie
    vrTotLoc = CCur(IIf(Not IsNumeric(txtTotalsDesc.Text), 0, txtTotalsDesc.Text))
    vrTotDescLoc = CCur(IIf(Not IsNumeric(txtDesconto.Text), 0, txtDesconto.Text))
 
    If vrPago < (vrTotLoc - vrTotDescLoc) Then
      'Valor do pagamento < que valor a pagar
      strMsg = "Valor pago não pode ser menor que valor a pagar" & vbCrLf
    ElseIf vrPago >= (vrTotLoc - vrTotDescLoc) Then
      'Valor do pagamento > que valor a pagar
      'strMsg = "Troco - R$ " & Format(vrPago - (vrTotLoc - vrTotDescLoc), "###,###,##0.00")
      txtTroco.Text = Format(vrPago - (vrTotLoc - vrTotDescLoc), "###,###,##0.00")
      
    End If
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserCalculadora.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    SetarFoco mskPgtoEspecie
    blnPrimeiraVez = False
    
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserCalculadora.Form_Activate]"
End Sub

Private Sub Form_Load()
On Error GoTo trata
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim objServicoDespertador     As busSisContas.clsServDesp
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 4545
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  txtUnidade.Text = strNumeroSuiteApto
  If tpStatus_Consultar Then
    INCLUIR_VALOR_NO_MASK mskPgtoEspecie, "", TpMaskMoeda
    txtTroco.Text = ""
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
  If Not bFechar Then Cancel = True
End Sub




Private Sub mskPgtoEspecie_GotFocus()
  Selecionar_Conteudo mskPgtoEspecie
End Sub

Private Sub mskPgtoEspecie_LostFocus()
  Pintar_Controle mskPgtoEspecie, tpCorContr_Normal
End Sub
