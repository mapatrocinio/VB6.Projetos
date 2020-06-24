VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserSaldoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Saldo"
   ClientHeight    =   2985
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2985
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   2985
      Left            =   8250
      ScaleHeight     =   2985
      ScaleWidth      =   1860
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   1875
         Left            =   0
         ScaleHeight     =   1815
         ScaleWidth      =   1605
         TabIndex        =   5
         Top             =   960
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Default         =   -1  'True
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   2655
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   4683
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da despesa"
      TabPicture(0)   =   "userSaldoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
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
         Height          =   1935
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   7575
         Begin VB.TextBox txtDescricao 
            Height          =   285
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   1
            Text            =   "txtDescricao"
            Top             =   600
            Width           =   5655
         End
         Begin MSMask.MaskEdBox mskPercentual 
            Height          =   255
            Left            =   1560
            TabIndex        =   0
            Top             =   240
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            Caption         =   "Descrição"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Percentual"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   240
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmUserSaldoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                         As tpStatus
Public lngSALDOID                     As Long
Public bRetorno                       As Boolean
Public blnPrimeiraVez                 As Boolean
Public bFechar                        As Boolean
Public strTipo                        As String


Private Sub cmdCancelar_Click()
  bFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub


Private Sub cmdOK_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim objSaldo                As busSisContas.clsSaldo
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Despesa
    If Not ValidaCampos Then Exit Sub
    Set objSaldo = New busSisContas.clsSaldo
    'Valida se unidade de estoque já cadastrada
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objSaldo.AlterarSaldo mskPercentual.Text, _
                            txtDescricao.Text, _
                            lngSALDOID
      bRetorno = True
      bFechar = True
      Set objSaldo = Nothing
      Unload Me
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      objSaldo.IncluirSaldo mskPercentual.Text, _
                            txtDescricao.Text
      'Limpar campos
      LimparCampoMask mskPercentual
      LimparCampoTexto txtDescricao
      SetarFoco mskPercentual
      bRetorno = True
      
    End If
    Set objSaldo = Nothing
    'bFechar = True
    'Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg              As String
  '
  If Not Valida_Moeda(mskPercentual, TpObrigatorio) Then
    strMsg = strMsg & "Informar o Percentual válido" & vbCrLf
    Pintar_Controle mskPercentual, tpCorContr_Erro
  End If
  If strMsg = "" Then
    If CCur(mskPercentual.Text) > 100 Or CCur(mskPercentual.Text) <= 0 Then
      strMsg = strMsg & "Informar o Percentual entre 0,01 e 100,00" & vbCrLf
      Pintar_Controle mskPercentual, tpCorContr_Erro
    End If
  End If
  If Len(txtDescricao.Text) = 0 Then
    strMsg = strMsg & "Informar a descrição válida" & vbCrLf
    Pintar_Controle txtDescricao, tpCorContr_Erro
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserSaldoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Setar foco
    SetarFoco mskPercentual
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserSaldoInc.Form_Activate]"
End Sub



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim objSaldo        As busSisContas.clsSaldo
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 3360
  Me.Width = 10200
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    LimparCampoMask mskPercentual
    LimparCampoTexto txtDescricao
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objSaldo = New busSisContas.clsSaldo
    Set objRs = objSaldo.SelecionarSaldo(lngSALDOID)
    '
    If Not objRs.EOF Then
      INCLUIR_VALOR_NO_MASK mskPercentual, objRs.Fields("PERCENTUAL").Value, TpMaskMoeda
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
    End If
    Set objSaldo = Nothing
  End If
  '
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not bFechar Then Cancel = True
End Sub

Private Sub mskPercentual_GotFocus()
  Selecionar_Conteudo mskPercentual
End Sub

Private Sub mskPercentual_LostFocus()
  Pintar_Controle mskPercentual, tpCorContr_Normal
End Sub

Private Sub txtDescricao_GotFocus()
  Selecionar_Conteudo txtDescricao
End Sub

Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub

