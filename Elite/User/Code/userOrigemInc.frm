VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmOrigemInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de origem/destino"
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
         Height          =   2085
         Left            =   90
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   270
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
      TabCaption(0)   =   "&Dados de origem"
      TabPicture(0)   =   "userOrigemInc.frx":0000
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
            Height          =   1215
            Left            =   0
            TabIndex        =   8
            Top             =   0
            Width           =   7695
            Begin VB.ComboBox cboIndicador 
               Height          =   315
               Left            =   1950
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   570
               Width           =   3285
            End
            Begin VB.TextBox txtNome 
               Height          =   285
               Left            =   1950
               MaxLength       =   50
               TabIndex        =   0
               Text            =   "txtNome"
               Top             =   240
               Width           =   5625
            End
            Begin VB.Label Label6 
               Caption         =   "Indicador"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   10
               Top             =   600
               Width           =   1455
            End
            Begin VB.Label Label6 
               Caption         =   "Origem"
               Height          =   255
               Index           =   3
               Left            =   120
               TabIndex        =   9
               Top             =   240
               Width           =   1755
            End
         End
      End
   End
End
Attribute VB_Name = "frmOrigemInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngORIGEMID                As Long
Public bRetorno                   As Boolean
Public bFechar                    As Boolean
Public sTitulo                    As String
Public intQuemChamou              As Integer
Private blnPrimeiraVez            As Boolean



Private Sub cboIndicador_LostFocus()
  Pintar_Controle cboIndicador, tpCorContr_Normal
End Sub

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
  Dim objOrigem                As busElite.clsOrigem
  Dim clsGer                  As busElite.clsGeral
  
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Unidade de estoque
    If Not ValidaCampos Then Exit Sub
    'Valida se Documento de entrada já cadastrado
    Set clsGer = New busElite.clsGeral
    strSql = "Select * From ORIGEM WHERE NOME = " & Formata_Dados(txtNome.Text, tpDados_Texto, tpNulo_NaoAceita) & _
      " AND PKID <> " & Formata_Dados(lngORIGEMID, tpDados_Longo, tpNulo_NaoAceita)
    Set objRs = clsGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set clsGer = Nothing
      TratarErroPrevisto "Origem já cadastrada", "cmdOK_Click"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set clsGer = Nothing
    '
    Set objOrigem = New busElite.clsOrigem
    '
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      '
      objOrigem.AlterarOrigem lngORIGEMID, _
                              txtNome.Text, _
                              cboIndicador.Text
                            
      bRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Pega Informações para inserir
      '
      objOrigem.InserirOrigem txtNome.Text, _
                              cboIndicador.Text
      '
      bRetorno = True
    End If
    Set objOrigem = Nothing
    bFechar = True
    Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg     As String
  '
  If Len(txtNome.Text) = 0 Then
    strMsg = strMsg & "Informar a origem" & vbCrLf
    Pintar_Controle txtNome, tpCorContr_Erro
  End If
  If Len(cboIndicador.Text) = 0 Then
    strMsg = "Selecionar o indicador" & vbCrLf
    Pintar_Controle cboIndicador, tpCorContr_Erro
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserOrigemLis.ValidaCampos]"
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
    txtNome.SetFocus
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserOrigemLis.Form_Activate]"
End Sub

Private Sub txtNome_GotFocus()
  Selecionar_Conteudo txtNome
End Sub

Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
End Sub



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs     As ADODB.Recordset
  Dim strSql    As String
  Dim objOrigem As busElite.clsOrigem
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
  txtNome.Text = ""
  LimparCampoCombo cboIndicador
  'Combos
  cboIndicador.AddItem ""
  cboIndicador.AddItem "AMBOS"
  cboIndicador.AddItem "DESTINO"
  cboIndicador.AddItem "ORIGEM"
  
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    '
  ElseIf Status = tpStatus_Alterar Then
    'Pega Dados do Banco de dados
    Set objOrigem = New busElite.clsOrigem
    Set objRs = objOrigem.ListarOrigem(lngORIGEMID)
    '
    If Not objRs.EOF Then
      txtNome.Text = objRs.Fields("NOME").Value & ""
      If objRs.Fields("IC_ORIGEM_DEST").Value & "" <> "" Then
        cboIndicador.Text = objRs.Fields("IC_ORIGEM_DEST").Value & ""
      End If
      
      '
    End If
    Set objOrigem = Nothing
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
