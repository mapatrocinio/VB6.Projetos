VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmEmbalagemInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de unidade de medida"
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
         Height          =   2085
         Left            =   0
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   5
         Top             =   690
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
      TabCaption(0)   =   "&Dados da unidade de medida"
      TabPicture(0)   =   "userEmbalagemInc.frx":0000
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
         Begin VB.TextBox txtSigla 
            Height          =   285
            Left            =   1560
            MaxLength       =   10
            TabIndex        =   1
            Text            =   "txtSigla"
            Top             =   600
            Width           =   2295
         End
         Begin VB.TextBox txtNome 
            Height          =   285
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   0
            Text            =   "txtNome"
            Top             =   270
            Width           =   5655
         End
         Begin VB.Label Label9 
            Caption         =   "Nome"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   270
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Sigla"
            Height          =   255
            Left            =   120
            TabIndex        =   8
            Top             =   600
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmEmbalagemInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                         As tpStatus
Public lngEMBALAGEMID                 As Long
Public blnRetorno                     As Boolean
Public blnPrimeiraVez                 As Boolean
Public blnFechar                      As Boolean


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
  Dim objEmbalagem            As busSisMetal.clsEmbalagem
  Dim objRs                   As ADODB.Recordset
  Dim objGer                  As busSisMetal.clsGeral
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Despesa
    If Not ValidaCampos Then Exit Sub
    'Valida se embalagem já cadastrada
    Set objGer = New busSisMetal.clsGeral
    strSql = "Select * From EMBALAGEM WHERE (NOME = " & Formata_Dados(txtNome.Text, tpDados_Texto, tpNulo_Aceita) & _
      " OR SIGLA = " & Formata_Dados(txtSigla.Text, tpDados_Texto, tpNulo_Aceita) & ") " & _
      " AND PKID <> " & Formata_Dados(lngEMBALAGEMID, tpDados_Longo, tpNulo_NaoAceita)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGer = Nothing
      TratarErroPrevisto "Nome ou Sigla da unidade de medida já cadastrada", "cmdOK_Click"
      Pintar_Controle txtNome, tpCorContr_Erro
      Pintar_Controle txtSigla, tpCorContr_Erro
      SetarFoco txtNome
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGer = Nothing
    '
    Set objEmbalagem = New busSisMetal.clsEmbalagem
    'Valida se unidade de estoque já cadastrada
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objEmbalagem.AlterarEmbalagem txtSigla.Text, _
                                    txtNome.Text, _
                                    lngEMBALAGEMID
      blnRetorno = True
      blnFechar = True
      Set objEmbalagem = Nothing
      Unload Me
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      objEmbalagem.IncluirEmbalagem txtSigla.Text, _
                                    txtNome.Text
      'Limpar campos
      blnRetorno = True
      blnFechar = True
      Set objEmbalagem = Nothing
      Unload Me
      
    End If
    Set objEmbalagem = Nothing
    'blnFechar = True
    'Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg              As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  '
  If Not Valida_String(txtNome, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Informar o nome válido" & vbCrLf
    'Pintar_Controle txtNome, tpCorContr_Erro
  End If
  If Not Valida_String(txtSigla, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Informar a sigla válida" & vbCrLf
    'Pintar_Controle txtSigla, tpCorContr_Erro
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmEmbalagemInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Setar foco
    SetarFoco txtNome
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmEmbalagemInc.Form_Activate]"
End Sub



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim objEmbalagem    As busSisMetal.clsEmbalagem
  '
  blnFechar = False
  blnRetorno = False
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
    LimparCampoTexto txtSigla
    LimparCampoTexto txtNome
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objEmbalagem = New busSisMetal.clsEmbalagem
    Set objRs = objEmbalagem.SelecionarEmbalagem(lngEMBALAGEMID)
    '
    If Not objRs.EOF Then
      txtSigla.Text = objRs.Fields("SIGLA").Value & ""
      txtNome.Text = objRs.Fields("NOME").Value & ""
    End If
    Set objEmbalagem = Nothing
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
  If Not blnFechar Then Cancel = True
End Sub

Private Sub txtSigla_GotFocus()
  Selecionar_Conteudo txtSigla
End Sub

Private Sub txtSigla_LostFocus()
  Pintar_Controle txtSigla, tpCorContr_Normal
End Sub

Private Sub txtNome_GotFocus()
  Selecionar_Conteudo txtNome
End Sub

Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
End Sub

