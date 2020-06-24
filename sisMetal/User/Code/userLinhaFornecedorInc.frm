VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmLinhaFornecedorInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Linha do Fornecedor"
   ClientHeight    =   4890
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4890
      Left            =   8430
      ScaleHeight     =   4890
      ScaleWidth      =   1860
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2115
         Left            =   60
         ScaleHeight     =   2055
         ScaleWidth      =   1605
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2640
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   150
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4665
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   8229
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da linha"
      TabPicture(0)   =   "userLinhaFornecedorInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
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
         Left            =   90
         TabIndex        =   11
         Top             =   390
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2385
            Index           =   0
            Left            =   120
            ScaleHeight     =   2385
            ScaleWidth      =   7575
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtCodigo 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   1
               Text            =   "txtCodigo"
               Top             =   420
               Width           =   6075
            End
            Begin VB.TextBox txtLinhaFim 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   3690
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   3
               TabStop         =   0   'False
               Text            =   "txtLinhaFim"
               Top             =   780
               Width           =   3705
            End
            Begin VB.TextBox txtCodigoFim 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1320
               Locked          =   -1  'True
               MaxLength       =   100
               TabIndex        =   2
               TabStop         =   0   'False
               Text            =   "txtCodigoFim"
               Top             =   780
               Width           =   2355
            End
            Begin VB.TextBox txtCodigoLin 
               Height          =   285
               Left            =   1320
               MaxLength       =   30
               TabIndex        =   4
               Text            =   "txtCodigoLin"
               Top             =   1140
               Width           =   3165
            End
            Begin VB.TextBox txtFornecedor 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtFornecedor"
               Top             =   90
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskPeso 
               Height          =   255
               Left            =   1320
               TabIndex        =   5
               Top             =   1470
               Visible         =   0   'False
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.000;($#,##0.000)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Caption         =   "Nome da Linha/Código Perfil"
               Height          =   615
               Index           =   0
               Left            =   90
               TabIndex        =   16
               Top             =   450
               Width           =   1095
            End
            Begin VB.Label Label1 
               Caption         =   "Código"
               Height          =   255
               Index           =   1
               Left            =   60
               TabIndex        =   15
               Top             =   1140
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Fornecedor"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   14
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Peso"
               Height          =   195
               Index           =   21
               Left            =   60
               TabIndex        =   13
               Top             =   1485
               Visible         =   0   'False
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmLinhaFornecedorInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngFORNECEDORID          As Long
Public strDescrFornecedor       As String

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Valor Produto Convênio
  LimparCampoTexto txtFornecedor
  LimparCampoTexto txtCodigo
  LimparCampoTexto txtCodigoFim
  LimparCampoTexto txtLinhaFim
  LimparCampoTexto txtCodigoLin
  LimparCampoMask mskPeso
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserLinhaFornecedorInc.LimparCampos]", _
            Err.Description
End Sub


Private Sub cmdCancelar_Click()
  blnFechar = True
  blnRetorno = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdOK_Click()
  Dim objLinhaFornecedor        As busSisMetal.clsLinhaFornecedor
  Dim objGeral                  As busSisMetal.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngLINHAID                As Long
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisMetal.clsGeral
  Set objLinhaFornecedor = New busSisMetal.clsLinhaFornecedor
  '
  'LINHA
  lngLINHAID = 0
  strSql = "SELECT LINHA.PKID FROM LINHA " & _
    " INNER JOIN TIPO_LINHA ON TIPO_LINHA.PKID = LINHA.TIPO_LINHAID " & _
      " WHERE LINHA.CODIGO LIKE '%" & txtCodigoFim.Text & "%'" & _
      " AND TIPO_LINHA.NOME LIKE '%" & txtLinhaFim.Text & "%'"
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngLINHAID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'Validar se a linha já cadastrada para fornecedor
  strSql = "SELECT * FROM LINHA_FORNECEDOR " & _
    " WHERE LINHA_FORNECEDOR.LINHAID = " & Formata_Dados(lngLINHAID, tpDados_Longo) & _
    " AND LINHA_FORNECEDOR.FORNECEDORID = " & Formata_Dados(lngFORNECEDORID, tpDados_Longo) & _
    " AND LINHA_FORNECEDOR.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    TratarErroPrevisto "Linha já cadastrada para este fornecedor"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objLinhaFornecedor = Nothing
    cmdOk.Enabled = True
    SetarFoco txtCodigo
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  'Validar se código do perfil já cadastrado para o fornecedor
  strSql = "SELECT * FROM LINHA_FORNECEDOR " & _
    " WHERE LINHA_FORNECEDOR.CODIGO = " & Formata_Dados(txtCodigoLin.Text, tpDados_Texto) & _
    " AND LINHA_FORNECEDOR.FORNECEDORID = " & Formata_Dados(lngFORNECEDORID, tpDados_Longo) & _
    " AND LINHA_FORNECEDOR.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    TratarErroPrevisto "Código do fornecedor já cadastrado para este fornecedor"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objLinhaFornecedor = Nothing
    cmdOk.Enabled = True
    SetarFoco txtCodigoLin
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar LinhaFornecedor
    objLinhaFornecedor.AlterarLinhaFornecedor lngPKID, _
                                              txtCodigoLin.Text, _
                                              mskPeso.Text
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir LinhaFornecedor
    objLinhaFornecedor.InserirLinhaFornecedor lngLINHAID, _
                                              lngFORNECEDORID, _
                                              txtCodigoLin.Text, _
                                              mskPeso.Text
  End If
  Set objLinhaFornecedor = Nothing
  blnRetorno = True
  blnFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  cmdOk.Enabled = True
End Sub


Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  If txtCodigoFim.Text = "" Or txtLinhaFim.Text = "" Then
    strMsg = strMsg & "Selecionar a linha" & vbCrLf
    Pintar_Controle txtCodigo, tpCorContr_Erro
    SetarFoco txtCodigo
    blnSetarFocoControle = False
  End If
  If Not Valida_String(txtCodigoLin, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o código válido" & vbCrLf
  End If
  If Not Valida_Moeda(mskPeso, TpnaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o peso válido" & vbCrLf
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserLinhaFornecedorInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserLinhaFornecedorInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco txtCodigo
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserLinhaFornecedorInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objLinhaFornecedor      As busSisMetal.clsLinhaFornecedor
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 5370
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  txtFornecedor.Text = strDescrFornecedor
  '
  If Status = tpStatus_Incluir Then
    '
    tabDetalhes.TabEnabled(0) = True
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objLinhaFornecedor = New busSisMetal.clsLinhaFornecedor
    Set objRs = objLinhaFornecedor.SelecionarLinhaFornecedorPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      txtCodigoFim.Text = objRs.Fields("CODIGO_LINHA").Value & ""
      txtLinhaFim.Text = objRs.Fields("NOME_LINHA").Value & ""
      txtCodigoLin.Text = objRs.Fields("CODIGO").Value & ""
      INCLUIR_VALOR_NO_MASK mskPeso, objRs.Fields("PESO").Value & "", TpMaskMoeda
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objLinhaFornecedor = Nothing
    'Visible
    tabDetalhes.TabEnabled(0) = True
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

Private Sub mskPeso_GotFocus()
  Seleciona_Conteudo_Controle mskPeso
End Sub
Private Sub mskPeso_LostFocus()
  Pintar_Controle mskPeso, tpCorContr_Normal
End Sub

Private Sub txtCodigo_GotFocus()
  Seleciona_Conteudo_Controle txtCodigo
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtCodigo_LostFocus()
  On Error GoTo trata
  Dim objLinhaCons    As Form
  Dim objLinhaPerfil        As busSisMetal.clsLinhaPerfil
  Dim objRs           As ADODB.Recordset
  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub

  Pintar_Controle txtCodigo, tpCorContr_Normal
  If Len(txtCodigo.Text) = 0 Then
    If Len(txtCodigoFim.Text) <> 0 And Len(txtLinhaFim.Text) <> 0 Then
      Exit Sub
    Else
      TratarErroPrevisto "Entre com o código ou descrição da linha."
      Pintar_Controle txtCodigo, tpCorContr_Erro
      SetarFoco txtCodigo
      Exit Sub
    End If
  End If
  Set objLinhaPerfil = New busSisMetal.clsLinhaPerfil
  '
  Set objRs = objLinhaPerfil.CapturaItemLinha(txtCodigo.Text)
  If objRs.EOF Then
    LimparCampoTexto txtCodigoFim
    LimparCampoTexto txtLinhaFim
    TratarErroPrevisto "Descrição/Código da linha não cadastrado"
    Pintar_Controle txtCodigo, tpCorContr_Erro
    SetarFoco txtCodigo
    Exit Sub
  Else
    If objRs.RecordCount = 1 Then
      txtCodigoFim.Text = objRs.Fields("CODIGO").Value & ""
      txtLinhaFim.Text = objRs.Fields("NOME").Value & ""
    
    Else
      'Novo : apresentar tela para seleção da linha
      Set objLinhaCons = New frmLinhaCons
      objLinhaCons.intIcOrigemLn = 2
      objLinhaCons.strCodigoDescricao = txtCodigo.Text
      objLinhaCons.Show vbModal
    End If
    '
    objRs.Close
    Set objRs = Nothing
    Set objLinhaPerfil = Nothing
'''    cmdOk.Default = True
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub


Private Sub txtCodigoLin_GotFocus()
  Seleciona_Conteudo_Controle txtCodigoLin
End Sub
Private Sub txtCodigoLin_LostFocus()
  Pintar_Controle txtCodigoLin, tpCorContr_Normal
End Sub

