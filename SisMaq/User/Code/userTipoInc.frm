VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUserTipoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de tipo de equipamento"
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
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2055
         Left            =   0
         ScaleHeight     =   1995
         ScaleWidth      =   1605
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   360
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   2295
      Left            =   120
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   4048
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do Tipo"
      TabPicture(0)   =   "userTipoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picTrava(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox picTrava 
         BorderStyle     =   0  'None
         Height          =   1665
         Index           =   0
         Left            =   120
         ScaleHeight     =   1665
         ScaleWidth      =   7695
         TabIndex        =   7
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
            Height          =   1575
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   7695
            Begin VB.TextBox txtDescricao 
               Height          =   825
               Left            =   1560
               MaxLength       =   255
               TabIndex        =   1
               Text            =   "txtDescricao"
               Top             =   570
               Width           =   5895
            End
            Begin VB.TextBox txtTipo 
               Height          =   285
               Left            =   1560
               MaxLength       =   100
               TabIndex        =   0
               Text            =   "txtTipo"
               Top             =   240
               Width           =   5895
            End
            Begin VB.Label Label6 
               Caption         =   "Descrição"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   10
               Top             =   570
               Width           =   1455
            End
            Begin VB.Label Label6 
               Caption         =   "Tipo"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Width           =   1455
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserTipoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngTIPOID                  As Long
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

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim objTipo                 As busSisMaq.clsTipo
  Dim objGer                  As busSisMaq.clsGeral
  
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração de Grupo cardápio
    If Not ValidaCampos Then Exit Sub
    'Valida se Grupo cardápio já cadastrado
    Set objGer = New busSisMaq.clsGeral
    strSql = "Select * From TIPO WHERE TIPO = " & Formata_Dados(txtTipo.Text, tpDados_Texto, tpNulo_Aceita) & _
      " AND PKID <> " & Formata_Dados(lngTIPOID, tpDados_Longo, tpNulo_NaoAceita)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGer = Nothing
      TratarErroPrevisto "Tipo já cadastrada", "cmdOK_Click"
      Pintar_Controle txtTipo, tpCorContr_Erro
      SetarFoco txtTipo
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGer = Nothing
    '
    Set objTipo = New busSisMaq.clsTipo
    '
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      '
      objTipo.AlterarTipo lngTIPOID, _
                          txtTipo.Text, _
                          txtDescricao.Text
                            
      bRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Pega Informações para inserir
      '
      objTipo.InserirTipo txtTipo.Text, _
                          txtDescricao.Text
      '
      bRetorno = True
    End If
    Set objTipo = Nothing
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
  If Len(txtTipo.Text) = 0 Then
    strMsg = strMsg & "Informar a descrição da Tipo" & vbCrLf
    Pintar_Controle txtTipo, tpCorContr_Erro
    SetarFoco txtTipo
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserTipoInc.ValidaCampos]"
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
    SetarFoco txtTipo
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserTipoInc.Form_Activate]"
End Sub




Private Sub Form_Load()
On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim objTipo  As busSisMaq.clsTipo
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
  txtTipo.Text = ""
  txtDescricao.Text = ""
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objTipo = New busSisMaq.clsTipo
    Set objRs = objTipo.ListarTipo(lngTIPOID)
    '
    If Not objRs.EOF Then
      txtTipo.Text = objRs.Fields("TIPO").Value & ""
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      '
    End If
    Set objTipo = Nothing
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

Private Sub txtTipo_GotFocus()
  Selecionar_Conteudo txtTipo
End Sub

Private Sub txtTipo_LostFocus()
  Pintar_Controle txtTipo, tpCorContr_Normal
End Sub

