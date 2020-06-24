VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUserUsuarioInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclus�o de usu�rio"
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
      TabCaption(0)   =   "&Dados do usu�rio"
      TabPicture(0)   =   "userUsuarioInc.frx":0000
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
            Height          =   1545
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   7695
            Begin VB.ComboBox cboNivel 
               Height          =   315
               Left            =   1590
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   630
               Width           =   2775
            End
            Begin VB.TextBox txtUsuario 
               Height          =   285
               Left            =   1590
               MaxLength       =   30
               TabIndex        =   0
               Top             =   270
               Width           =   2745
            End
            Begin VB.Label Label6 
               Caption         =   "Usu�rio"
               Height          =   255
               Index           =   0
               Left            =   150
               TabIndex        =   10
               Top             =   270
               Width           =   1455
            End
            Begin VB.Label Label6 
               Caption         =   "N�vel"
               Height          =   255
               Index           =   3
               Left            =   180
               TabIndex        =   9
               Top             =   600
               Width           =   1455
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserUsuarioInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngCONTROLEACESSOID        As Long
Public bRetorno                   As Boolean
Public bFechar                    As Boolean
Public sTitulo                    As String
Public intQuemChamou              As Integer
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
  Dim objUsuario              As busSisContas.clsUsuario
  Dim objGer                  As busSisContas.clsGeral
  
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclus�o/Altera��o de Grupo card�pio
    If Not ValidaCampos Then Exit Sub
    'Valida se cart�o j� cadastrado
    Set objGer = New busSisContas.clsGeral
    strSql = "Select * From CONTROLEACESSO WHERE USUARIO = " & Formata_Dados(txtUsuario.Text, tpDados_Texto, tpNulo_Aceita) & _
      " AND PKID <> " & Formata_Dados(lngCONTROLEACESSOID, tpDados_Longo, tpNulo_NaoAceita)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGer = Nothing
      TratarErroPrevisto "Usuario j� cadastrado", "cmdOK_Click"
      Pintar_Controle txtUsuario, tpCorContr_Erro
      SetarFoco txtUsuario
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objUsuario = New busSisContas.clsUsuario
    '
    If Status = tpStatus_Alterar Then
      'C�digo para altera��o
      '
      '
      objUsuario.AlterarUsuario lngCONTROLEACESSOID, _
                                txtUsuario.Text, _
                                Left(cboNivel.Text, 3)
                            
      bRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Pega Informa��es para inserir
      '
      objUsuario.InserirUsuario txtUsuario.Text, _
                                Left(cboNivel.Text, 3)
      '
      bRetorno = True
    End If
    Set objUsuario = Nothing
    bFechar = True
    Unload Me
  End Select
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
  If Not Valida_String(txtUsuario, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Informar o nome do Usuario" & vbCrLf
  End If
  If Not Valida_String(cboNivel, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o n�vel do usu�rio" & vbCrLf
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserUsuarioInc.ValidaCampos]"
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
    SetarFoco txtUsuario
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserUsuarioInc.Form_Activate]"
End Sub

Private Sub txtUsuario_GotFocus()
  Selecionar_Conteudo txtUsuario
End Sub

Private Sub txtUsuario_LostFocus()
  Pintar_Controle txtUsuario, tpCorContr_Normal
End Sub

Private Sub cboNivel_LostFocus()
  Pintar_Controle cboNivel, tpCorContr_Normal
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim objUsuario    As busSisContas.clsUsuario
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
  cboNivel.Clear
  cboNivel.AddItem ""
  If gsNivel = gsAdmin Then _
    cboNivel.AddItem "ADMINISTRADOR"
  cboNivel.AddItem "FINANCEIRO"
  '
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclus�o, Inclui o Pedido
    txtUsuario.Text = ""
    cboNivel.ListIndex = -1
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Usuario de dados
    Set objUsuario = New busSisContas.clsUsuario
    Set objRs = objUsuario.ListarUsuario(lngCONTROLEACESSOID)
    '
    If Not objRs.EOF Then
      txtUsuario.Text = objRs.Fields("USUARIO").Value & ""
      cboNivel.Text = objRs.Fields("DESCNIVEL").Value & ""
      '
    End If
    Set objUsuario = Nothing
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

