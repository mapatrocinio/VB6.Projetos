VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmIPIInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de IPI"
   ClientHeight    =   2565
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2565
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   2565
      Left            =   8520
      ScaleHeight     =   2565
      ScaleWidth      =   1860
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   30
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   330
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   2295
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   4048
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do IPI"
      TabPicture(0)   =   "userIPIInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picTrava(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox picTrava 
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   0
         Left            =   120
         ScaleHeight     =   1335
         ScaleWidth      =   7695
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   480
         Width           =   7695
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
            Height          =   1215
            Left            =   0
            TabIndex        =   7
            Top             =   0
            Width           =   7695
            Begin MSMask.MaskEdBox mskValor 
               Height          =   255
               Left            =   1200
               TabIndex        =   0
               Top             =   240
               Width           =   1695
               _ExtentX        =   2990
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label6 
               Caption         =   "IPI"
               Height          =   255
               Index           =   3
               Left            =   240
               TabIndex        =   8
               Top             =   240
               Width           =   855
            End
         End
      End
   End
End
Attribute VB_Name = "frmIPIInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngIPIID                   As Long
Public blnRetorno                 As Boolean
Public blnFechar                  As Boolean
Private blnPrimeiraVez            As Boolean



Private Sub cmdCancelar_Click()
  blnFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim objIPI                  As busSisMetal.clsIPI
  Dim objGral                 As busSisMetal.clsGeral

  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Unidade de estoque
    If Not ValidaCampos Then Exit Sub
    'Valida se unidade de estoque já cadastrada
    Set objGral = New busSisMetal.clsGeral
    strSql = "Select * From IPI WHERE IPI = " & Formata_Dados(mskValor.Text, tpDados_Moeda) & _
      " AND PKID <> " & Formata_Dados(lngIPIID, tpDados_Longo, tpNulo_NaoAceita)
    Set objRs = objGral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGral = Nothing
      Pintar_Controle mskValor, tpCorContr_Erro
      TratarErroPrevisto "IPI já cadastrado", "cmdOK_Click"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGral = Nothing
    '
    Set objIPI = New busSisMetal.clsIPI
    '
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      '
      objIPI.AlterarIPI lngIPIID, _
                        mskValor.Text

      blnRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Pega Informações para inserir
      '
      objIPI.InserirIPI mskValor.Text
      '
      blnRetorno = True
    End If
    Set objIPI = Nothing
    blnFechar = True
    Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg     As String
  Dim blnSetarFocoControle  As Boolean
  '
  If Not Valida_Moeda(mskValor, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Informar o percentual do IPI válido" & vbCrLf
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserIPIInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    tabDetalhes.Tab = 0
    SetarFoco mskValor
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserIPIInc.Form_Activate]"
End Sub




Private Sub Form_Load()
On Error GoTo trata
  Dim objRs     As ADODB.Recordset
  Dim strSql    As String
  Dim objIPI   As busSisMetal.clsIPI
  '
  blnFechar = False
  blnRetorno = False
  AmpS
  Me.Height = 2940
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  LimparCampoMask mskValor
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    '
  ElseIf Status = tpStatus_Alterar Then
    'Pega Dados do Banco de dados
    Set objIPI = New busSisMetal.clsIPI
    Set objRs = objIPI.ListarIPI(lngIPIID)
    '
    If Not objRs.EOF Then
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("IPI").Value, TpMaskMoeda
      '
    End If
    Set objIPI = Nothing
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
  If Not blnFechar Then Cancel = True
End Sub

Private Sub mskValor_GotFocus()
  Selecionar_Conteudo mskValor
End Sub

Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub

