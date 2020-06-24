VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLinhaPerfilInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Linha-Perfil"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   3375
      Left            =   8250
      ScaleHeight     =   3375
      ScaleWidth      =   1860
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   30
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   11
         Top             =   1140
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   3105
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   5477
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da linha-perfil"
      TabPicture(0)   =   "userLinhaPerfilInc.frx":0000
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
         Height          =   2475
         Left            =   240
         TabIndex        =   13
         Top             =   480
         Width           =   7575
         Begin VB.TextBox txtNomeProduto 
            Height          =   285
            Left            =   1560
            MaxLength       =   100
            TabIndex        =   3
            Text            =   "txtNomeProduto"
            Top             =   1290
            Width           =   5265
         End
         Begin VB.ComboBox cboLinha 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   210
            Width           =   5295
         End
         Begin VB.ComboBox cboVara 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   930
            Width           =   2805
         End
         Begin VB.TextBox txtCodigo 
            Height          =   285
            Left            =   1560
            MaxLength       =   30
            TabIndex        =   1
            Text            =   "txtCodigo"
            Top             =   600
            Width           =   2805
         End
         Begin MSMask.MaskEdBox mskPesoVara 
            Height          =   255
            Left            =   1560
            TabIndex        =   4
            Top             =   1620
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.0000;($#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskLargura 
            Height          =   255
            Left            =   5130
            TabIndex        =   5
            Top             =   1620
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.0000;($#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskAba 
            Height          =   255
            Left            =   1560
            TabIndex        =   6
            Top             =   1920
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.0000;($#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskEspessura 
            Height          =   255
            Left            =   5130
            TabIndex        =   7
            Top             =   1920
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.0000;($#,##0.0000)"
            PromptChar      =   "_"
         End
         Begin VB.Label Label7 
            Caption         =   "Espessura"
            Height          =   255
            Left            =   3690
            TabIndex        =   21
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label6 
            Caption         =   "Aba"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1920
            Width           =   1095
         End
         Begin VB.Label Label3 
            Caption         =   "Largura"
            Height          =   255
            Left            =   3690
            TabIndex        =   19
            Top             =   1620
            Width           =   1095
         End
         Begin VB.Label Label2 
            Caption         =   "Nome Produto"
            Height          =   255
            Left            =   120
            TabIndex        =   18
            Top             =   1290
            Width           =   1215
         End
         Begin VB.Label Label5 
            Caption         =   "Comprimento"
            Height          =   195
            Index           =   24
            Left            =   120
            TabIndex        =   17
            Top             =   960
            Width           =   1215
         End
         Begin VB.Label Label4 
            Caption         =   "Peso"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   1620
            Width           =   1095
         End
         Begin VB.Label Label1 
            Caption         =   "Código do perfil"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   600
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Nome da Linha"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   270
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmLinhaPerfilInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                         As tpStatus
Public lngLINHAID                     As Long
Public blnRetorno                     As Boolean
Public blnPrimeiraVez                 As Boolean
Public blnFechar                      As Boolean


Private Sub cboLinha_LostFocus()
  Pintar_Controle cboLinha, tpCorContr_Normal
End Sub

Private Sub cmdCancelar_Click()
  blnFechar = True
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
  Dim objLinhaPerfil          As busSisMetal.clsLinhaPerfil
  Dim objRs                   As ADODB.Recordset
  Dim objGer                  As busSisMetal.clsGeral
  Dim lngVARAID               As Long
  Dim lngTIPOLINHAID          As Long
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Despesa
    If Not ValidaCampos Then Exit Sub
    'Valida se linha já cadastrada
    Set objGer = New busSisMetal.clsGeral
    '
    'TIPO_LINHA
    lngTIPOLINHAID = 0
    strSql = "SELECT PKID FROM TIPO_LINHA WHERE NOME = " & Formata_Dados(cboLinha.Text, tpDados_Texto)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngTIPOLINHAID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
    strSql = "Select * From LINHA " & _
      " WHERE (TIPO_LINHAID = " & Formata_Dados(lngTIPOLINHAID, tpDados_Longo) & _
      " AND CODIGO = " & Formata_Dados(txtCodigo.Text, tpDados_Texto, tpNulo_Aceita) & ") " & _
      " AND PKID <> " & Formata_Dados(lngLINHAID, tpDados_Longo, tpNulo_NaoAceita)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGer = Nothing
      TratarErroPrevisto "Linha-perfil já cadastrado", "cmdOK_Click"
      Pintar_Controle cboLinha, tpCorContr_Erro
      Pintar_Controle txtCodigo, tpCorContr_Erro
      SetarFoco cboLinha
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    '
    'VARA
    lngVARAID = 0
    strSql = "SELECT PKID FROM VARA WHERE NOME = " & Formata_Dados(cboVara.Text, tpDados_Texto)
    Set objRs = objGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngVARAID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objGer = Nothing
    '
    Set objLinhaPerfil = New busSisMetal.clsLinhaPerfil
    'Valida se unidade de estoque já cadastrada
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objLinhaPerfil.AlterarLinha lngTIPOLINHAID, _
                            txtCodigo.Text, _
                            lngVARAID, _
                            IIf(Len(mskPesoVara.ClipText) = 0, "", mskPesoVara.Text), _
                            lngLINHAID, _
                            txtNomeProduto.Text, _
                            IIf(Len(mskLargura.ClipText) = 0, "", mskLargura.Text), _
                            IIf(Len(mskAba.ClipText) = 0, "", mskAba.Text), _
                            IIf(Len(mskEspessura.ClipText) = 0, "", mskEspessura.Text)
      blnRetorno = True
      blnFechar = True
      Set objLinhaPerfil = Nothing
      Unload Me
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      objLinhaPerfil.IncluirLinha lngTIPOLINHAID, _
                            txtCodigo.Text, _
                            lngVARAID, _
                            IIf(Len(mskPesoVara.ClipText) = 0, "", mskPesoVara.Text), _
                            txtNomeProduto.Text, _
                            IIf(Len(mskLargura.ClipText) = 0, "", mskLargura.Text), _
                            IIf(Len(mskAba.ClipText) = 0, "", mskAba.Text), _
                            IIf(Len(mskEspessura.ClipText) = 0, "", mskEspessura.Text)
      'Limpar campos
      blnRetorno = True
      blnFechar = True
      Set objLinhaPerfil = Nothing
      Unload Me
      
    End If
    Set objLinhaPerfil = Nothing
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
  If Not Valida_String(cboLinha, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar a linha" & vbCrLf
  End If
  If Not Valida_String(txtCodigo, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Informar o código do perfil válido" & vbCrLf
  End If
  If Not Valida_String(cboVara, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o comprimento" & vbCrLf
  End If
'''  If Not Valida_String(txtNomeProduto, TpObrigatorio, blnSetarFocoControle) Then
'''    strMsg = strMsg & "Informar o nome do produto válido" & vbCrLf
'''  End If
  If Not Valida_Moeda(mskPesoVara, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Informar o peso válido" & vbCrLf
  End If
  If Not Valida_Moeda(mskLargura, TpnaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Informar a largura válida" & vbCrLf
  End If
  If Not Valida_Moeda(mskAba, TpnaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Informar a aba válida" & vbCrLf
  End If
  If Not Valida_Moeda(mskEspessura, TpnaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Informar a espessura válida" & vbCrLf
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmLinhaPerfilInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Setar foco
    SetarFoco cboLinha
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmLinhaPerfilInc.Form_Activate]"
End Sub

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Linha
  LimparCampoCombo cboLinha
  LimparCampoTexto txtCodigo
  LimparCampoCombo cboVara
  LimparCampoTexto txtNomeProduto
  LimparCampoMask mskPesoVara
  LimparCampoMask mskLargura
  LimparCampoMask mskAba
  LimparCampoMask mskEspessura
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmLinhaPerfilInc.LimparCampos]", _
            Err.Description
End Sub


Private Sub Form_Load()
On Error GoTo trata
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim objLinhaPerfil  As busSisMetal.clsLinhaPerfil
  '
  blnFechar = False
  blnRetorno = False
  AmpS
  Me.Height = 3855
  Me.Width = 10200
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  LimparCampos
  '
  'Vara
  strSql = "Select NOME from VARA ORDER BY NOME"
  PreencheCombo cboVara, strSql, False, True
  'Linha
  strSql = "Select NOME from TIPO_LINHA ORDER BY NOME"
  PreencheCombo cboLinha, strSql, False, True
  '
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objLinhaPerfil = New busSisMetal.clsLinhaPerfil
    Set objRs = objLinhaPerfil.SelecionarLinha(lngLINHAID)
    '
    If Not objRs.EOF Then
      cboLinha.Text = objRs.Fields("NOME").Value & ""
      txtCodigo.Text = objRs.Fields("CODIGO").Value & ""
      cboVara.Text = objRs.Fields("NOME_VARA").Value & ""
      txtNomeProduto.Text = objRs.Fields("NOME_PRODUTO").Value & ""
      INCLUIR_VALOR_NO_MASK mskPesoVara, objRs.Fields("PESO_VARA").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskLargura, objRs.Fields("LARGURA").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskAba, objRs.Fields("ABA").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskEspessura, objRs.Fields("ESPESSURA").Value, TpMaskMoeda
    End If
    Set objLinhaPerfil = Nothing
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



Private Sub mskAba_GotFocus()
  Selecionar_Conteudo mskAba
End Sub

Private Sub mskAba_LostFocus()
  Pintar_Controle mskAba, tpCorContr_Normal
End Sub

Private Sub mskEspessura_GotFocus()
  Selecionar_Conteudo mskEspessura
End Sub

Private Sub mskEspessura_LostFocus()
  Pintar_Controle mskEspessura, tpCorContr_Normal
End Sub

Private Sub mskLargura_GotFocus()
  Selecionar_Conteudo mskLargura
End Sub

Private Sub mskLargura_LostFocus()
  Pintar_Controle mskLargura, tpCorContr_Normal
End Sub

Private Sub txtCodigo_GotFocus()
  Selecionar_Conteudo txtCodigo
End Sub

Private Sub txtCodigo_LostFocus()
  Pintar_Controle txtCodigo, tpCorContr_Normal
End Sub
Private Sub cboVara_LostFocus()
  Pintar_Controle cboVara, tpCorContr_Normal
End Sub

Private Sub mskPesoVara_GotFocus()
  Selecionar_Conteudo mskPesoVara
End Sub

Private Sub mskPesoVara_LostFocus()
  Pintar_Controle mskPesoVara, tpCorContr_Normal
End Sub



Private Sub txtNomeProduto_GotFocus()
  Selecionar_Conteudo txtNomeProduto
End Sub

Private Sub txtNomeProduto_LostFocus()
  Pintar_Controle txtNomeProduto, tpCorContr_Normal
End Sub
