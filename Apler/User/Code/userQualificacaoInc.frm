VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUserQualificacaoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclus�o da Qualifica��o"
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
         Height          =   2055
         Left            =   0
         ScaleHeight     =   1995
         ScaleWidth      =   1605
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   360
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
      TabCaption(0)   =   "&Dados da Qualifica��o"
      TabPicture(0)   =   "userQualificacaoInc.frx":0000
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
            Begin VB.TextBox txtDescricao 
               Height          =   285
               Left            =   1560
               MaxLength       =   100
               TabIndex        =   0
               Text            =   "txtDescricao"
               Top             =   240
               Width           =   5895
            End
            Begin VB.Label Label6 
               Caption         =   "Qualifica��o"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   8
               Top             =   240
               Width           =   1455
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserQualificacaoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngQUALIFICACAOID          As Long
Public bRetorno                   As Boolean
Public bFechar                    As Boolean
Private blnPrimeiraVez            As Boolean

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
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim objQualificacao         As busApler.clsQualificacao
  Dim objGer                  As busApler.clsGeral
  
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclus�o/Altera��o de Grupo card�pio
    If Not ValidaCampos Then Exit Sub
    'Valida se Grupo card�pio j� cadastrado
    Set objGer = New busApler.clsGeral
    strSql = "Select * From QUALIFICACAO WHERE DESCRICAO = " & Formata_Dados(txtDescricao.Text, tpDados_Texto, tpNulo_Aceita) & _
      " AND PKID <> " & Formata_Dados(lngQUALIFICACAOID, tpDados_Longo, tpNulo_NaoAceita)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGer = Nothing
      TratarErroPrevisto "Qualifica��o j� cadastrada", "cmdOK_Click"
      Pintar_Controle txtDescricao, tpCorContr_Erro
      SetarFoco txtDescricao
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGer = Nothing
    '
    Set objQualificacao = New busApler.clsQualificacao
    '
    If Status = tpStatus_Alterar Then
      'C�digo para altera��o
      '
      '
      objQualificacao.AlterarQualificacao lngQUALIFICACAOID, _
                                          txtDescricao.Text
                            
      bRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Pega Informa��es para inserir
      '
      objQualificacao.InserirQualificacao txtDescricao.Text
      '
      bRetorno = True
    End If
    Set objQualificacao = Nothing
    bFechar = True
    Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg     As String
  '
  If Len(txtDescricao.Text) = 0 Then
    strMsg = strMsg & "Informar a descri��o da Qualifica��o" & vbCrLf
    Pintar_Controle txtDescricao, tpCorContr_Erro
    SetarFoco txtDescricao
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserQualificacaoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    tabDetalhes.Tab = 0
    SetarFoco txtDescricao
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserQualificacaoInc.Form_Activate]"
End Sub

Private Sub txtDescricao_GotFocus()
  Selecionar_Conteudo txtDescricao
End Sub

Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs               As ADODB.Recordset
  Dim strSql              As String
  Dim objQualificacao     As busApler.clsQualificacao
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 2940
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclus�o, Inclui o Pedido
    txtDescricao.Text = ""
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objQualificacao = New busApler.clsQualificacao
    Set objRs = objQualificacao.ListarQualificacao(lngQUALIFICACAOID)
    '
    If Not objRs.EOF Then
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      '
    End If
    Set objQualificacao = Nothing
  End If
  
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub



Private Sub Form_Unload(Cancel As Integer)
  If Not bFechar Then Cancel = True
End Sub
