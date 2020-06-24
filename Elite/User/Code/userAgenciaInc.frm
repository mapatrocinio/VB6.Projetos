VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmAgenciaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Agência"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5610
      Left            =   8430
      ScaleHeight     =   5610
      ScaleWidth      =   1860
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4575
         Left            =   90
         ScaleHeight     =   4515
         ScaleWidth      =   1605
         TabIndex        =   30
         TabStop         =   0   'False
         Top             =   810
         Width           =   1665
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   2700
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   3540
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   1830
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   90
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5295
      Left            =   120
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userAgenciaInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&CNPJ"
      TabPicture(1)   =   "userAgenciaInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdCNPJ"
      Tab(1).ControlCount=   1
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
         Height          =   4755
         Left            =   120
         TabIndex        =   22
         Top             =   330
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   4515
            Index           =   0
            Left            =   120
            ScaleHeight     =   4515
            ScaleWidth      =   7575
            TabIndex        =   23
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtObservacao 
               Height          =   615
               Left            =   1320
               MaxLength       =   255
               MultiLine       =   -1  'True
               TabIndex        =   11
               Text            =   "userAgenciaInc.frx":0038
               Top             =   2700
               Width           =   6075
            End
            Begin VB.TextBox txtCidade 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   10
               Text            =   "txtCidade"
               Top             =   2370
               Width           =   6075
            End
            Begin VB.TextBox txtBairro 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   9
               Text            =   "txtBairro"
               Top             =   2040
               Width           =   6075
            End
            Begin VB.TextBox txtRua 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   4
               Text            =   "txtRua"
               Top             =   1050
               Width           =   6075
            End
            Begin VB.TextBox txtNumero 
               Height          =   285
               Left            =   1320
               MaxLength       =   15
               TabIndex        =   5
               Text            =   "txtNumero"
               Top             =   1380
               Width           =   2175
            End
            Begin VB.TextBox txtComplemento 
               Height          =   285
               Left            =   5220
               MaxLength       =   30
               TabIndex        =   6
               Text            =   "txtComplemento"
               Top             =   1380
               Width           =   2175
            End
            Begin VB.TextBox txtEstado 
               Height          =   285
               Left            =   1320
               MaxLength       =   2
               TabIndex        =   7
               Text            =   "txtEstado"
               Top             =   1710
               Width           =   435
            End
            Begin VB.TextBox txtTelefone3 
               Height          =   285
               Left            =   1320
               MaxLength       =   30
               TabIndex        =   3
               Text            =   "txtTelefone3"
               Top             =   720
               Width           =   2175
            End
            Begin VB.TextBox txtTelefone2 
               Height          =   285
               Left            =   5220
               MaxLength       =   30
               TabIndex        =   2
               Text            =   "txtTelefone2"
               Top             =   390
               Width           =   2175
            End
            Begin VB.TextBox txtTelefone1 
               Height          =   285
               Left            =   1320
               MaxLength       =   30
               TabIndex        =   1
               Text            =   "txtTelefone1"
               Top             =   390
               Width           =   2175
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1290
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   3360
               Width           =   2235
               Begin VB.OptionButton optStatus 
                  Caption         =   "Inativo"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   13
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optStatus 
                  Caption         =   "Ativo"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   12
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   825
               End
            End
            Begin VB.TextBox txtNome 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               Text            =   "txtNome"
               Top             =   75
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskCep 
               Height          =   255
               Left            =   5220
               TabIndex        =   8
               Top             =   1710
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##.###-###"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Observação"
               Height          =   195
               Index           =   32
               Left            =   60
               TabIndex        =   38
               Top             =   2745
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cep"
               Height          =   195
               Index           =   3
               Left            =   3960
               TabIndex        =   37
               Top             =   1710
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cidade"
               Height          =   195
               Index           =   16
               Left            =   60
               TabIndex        =   36
               Top             =   2415
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Bairro"
               Height          =   195
               Index           =   2
               Left            =   60
               TabIndex        =   35
               Top             =   2085
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Rua"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   34
               Top             =   1095
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Número"
               Height          =   195
               Index           =   7
               Left            =   60
               TabIndex        =   33
               Top             =   1380
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Complemento"
               Height          =   195
               Index           =   8
               Left            =   3960
               TabIndex        =   32
               Top             =   1380
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Estado"
               Height          =   195
               Index           =   9
               Left            =   60
               TabIndex        =   31
               Top             =   1710
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Telefone 3"
               Height          =   195
               Index           =   29
               Left            =   60
               TabIndex        =   29
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Telefone 2"
               Height          =   195
               Index           =   28
               Left            =   3960
               TabIndex        =   28
               Top             =   390
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Telefone 1"
               Height          =   195
               Index           =   27
               Left            =   60
               TabIndex        =   27
               Top             =   390
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Status"
               Height          =   315
               Index           =   5
               Left            =   60
               TabIndex        =   25
               Top             =   3390
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Nome"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   24
               Top             =   120
               Width           =   1215
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdCNPJ 
         Height          =   4545
         Left            =   -74880
         OleObjectBlob   =   "userAgenciaInc.frx":0046
         TabIndex        =   14
         Top             =   420
         Width           =   4425
      End
   End
End
Attribute VB_Name = "frmAgenciaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long

Private blnPrimeiraVez          As Boolean

Dim CNPJ_COLUNASMATRIZ         As Long
Dim CNPJ_LINHASMATRIZ          As Long

Private CNPJ_Matriz()          As String

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Agencia
  LimparCampoTexto txtNome
  LimparCampoTexto txtTelefone1
  LimparCampoTexto txtTelefone2
  LimparCampoTexto txtTelefone3
  LimparCampoTexto txtRua
  LimparCampoTexto txtNumero
  LimparCampoTexto txtComplemento
  LimparCampoTexto txtEstado
  LimparCampoMask mskCep
  LimparCampoTexto txtBairro
  LimparCampoTexto txtCidade
  LimparCampoTexto txtObservacao
  optStatus(0).Value = False
  optStatus(1).Value = False
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmAgenciaInc.LimparCampos]", _
            Err.Description
End Sub



Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 1
    If Not IsNumeric(grdCNPJ.Columns("PKID").Value & "") Then
      MsgBox "Selecione um CNPJ !", vbExclamation, TITULOSISTEMA
      SetarFoco grdCNPJ
      Exit Sub
    End If

    frmAgenciaCNPJInc.lngPKID = grdCNPJ.Columns("PKID").Value
    frmAgenciaCNPJInc.lngAGENCIAID = lngPKID
    frmAgenciaCNPJInc.strNomeAgencia = txtNome.Text
    frmAgenciaCNPJInc.Status = tpStatus_Alterar
    frmAgenciaCNPJInc.Show vbModal

    If frmAgenciaCNPJInc.blnRetorno Then
      CNPJ_MontaMatriz
      grdCNPJ.Bookmark = Null
      grdCNPJ.ReBind
      grdCNPJ.ApproxCount = CNPJ_LINHASMATRIZ
    End If
    SetarFoco grdCNPJ
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
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





Private Sub cmdExcluir_Click()
  On Error GoTo trata
  Dim objAgenciaCNPJ     As busElite.clsAgenciaCNPJ
  '
  Select Case tabDetalhes.Tab
  Case 1 'Exclusão de Agência CNPJ
    '
    If Len(Trim(grdCNPJ.Columns("PKID").Value & "")) = 0 Then
      MsgBox "Selecione um CNPJ.", vbExclamation, TITULOSISTEMA
      SetarFoco grdCNPJ
      Exit Sub
    End If
    '
    Set objAgenciaCNPJ = New busElite.clsAgenciaCNPJ
    '
    If MsgBox("Confirma exclusão do CNPJ " & grdCNPJ.Columns("CNPJ").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdCNPJ
      Exit Sub
    End If
    'OK
    objAgenciaCNPJ.ExcluirAgenciaCNPJ CLng(grdCNPJ.Columns("PKID").Value)
    '
    CNPJ_MontaMatriz
    grdCNPJ.Bookmark = Null
    grdCNPJ.ReBind
    grdCNPJ.ApproxCount = CNPJ_LINHASMATRIZ

    Set objAgenciaCNPJ = Nothing
    SetarFoco grdCNPJ
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub





Private Sub cmdIncluir_Click()
  On Error GoTo trata
  Dim objForm As Form
  '
  Select Case tabDetalhes.Tab
  Case 1
    frmAgenciaCNPJInc.Status = tpStatus_Incluir
    frmAgenciaCNPJInc.lngPKID = 0
    frmAgenciaCNPJInc.lngAGENCIAID = lngPKID
    frmAgenciaCNPJInc.strNomeAgencia = txtNome.Text
    frmAgenciaCNPJInc.Show vbModal

    If frmAgenciaCNPJInc.blnRetorno Then
      CNPJ_MontaMatriz
      grdCNPJ.Bookmark = Null
      grdCNPJ.ReBind
      grdCNPJ.ApproxCount = CNPJ_LINHASMATRIZ
    End If
    SetarFoco grdCNPJ
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdOK_Click()
  Dim objAgencia                As busElite.clsAgencia
  Dim objGeral                  As busElite.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim strStatus                 As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busElite.clsGeral
  Set objAgencia = New busElite.clsAgencia
  'Status
  If optStatus(0).Value Then
    strStatus = "A"
  Else
    strStatus = "I"
  End If

  'Validar se agência já cadastrada
  strSql = "SELECT * FROM AGENCIA " & _
    " WHERE AGENCIA.NOME = " & Formata_Dados(txtNome.Text, tpDados_Texto) & _
    " AND AGENCIA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
    " AND AGENCIA.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtNome, tpCorContr_Erro
    TratarErroPrevisto "Agência já cadastrada"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objAgencia = Nothing
    cmdOk.Enabled = True
    SetarFoco txtNome
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Agencia
    objAgencia.AlterarAgencia lngPKID, _
                              txtNome.Text, _
                              txtTelefone1.Text, _
                              txtTelefone2.Text, _
                              txtTelefone3.Text, _
                              txtRua.Text, _
                              txtNumero.Text, _
                              txtComplemento.Text, _
                              txtEstado.Text, _
                              IIf(mskCep.ClipText = "", "", mskCep.ClipText), _
                              txtBairro.Text, _
                              txtCidade.Text, _
                              txtObservacao.Text, _
                              strStatus
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Agencia
    objAgencia.InserirAgencia lngPKID, _
                              txtNome.Text, _
                              txtTelefone1.Text, _
                              txtTelefone2.Text, _
                              txtTelefone3.Text, _
                              txtRua.Text, _
                              txtNumero.Text, _
                              txtComplemento.Text, _
                              txtEstado.Text, _
                              IIf(mskCep.ClipText = "", "", mskCep.ClipText), _
                              txtBairro.Text, _
                              txtCidade.Text, _
                              txtObservacao.Text, _
                              strStatus
    blnRetorno = True
    'Selecionar plano cadastrado
    'entrar em modo de alteração
    'lngPKID = objRs.Fields("PKID")
    Status = tpStatus_Alterar
    'Reload na tela
    Form_Load
    'Acerta tabs
    'tabDetalhes.Tab = 2
    blnRetorno = True
  End If
  Set objAgencia = Nothing
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
  If Not Valida_String(txtNome, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o nome" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Option(optStatus, blnSetarFocoControle) Then
    strMsg = strMsg & "Slecionar o status" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(Trim(mskCep.ClipText)) > 0 Then
    If Len(Trim(mskCep.ClipText)) <> 8 Then
      strMsg = strMsg & "Informar o CEP válido" & vbCrLf
      Pintar_Controle mskCep, tpCorContr_Erro
      SetarFoco mskCep
      tabDetalhes.Tab = 0
      blnSetarFocoControle = False
    End If
  End If

  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmAgenciaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmAgenciaInc.ValidaCampos]", _
            Err.Description
End Function



Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco txtNome
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmAgenciaInc.Form_Activate]"
End Sub



Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objAgencia              As busElite.clsAgencia
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 6090
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , cmdIncluir, cmdAlterar
  '
  'Limpar Campos
  LimparCampos
  tabDetalhes_Click 1
  'TEMPORÁRIO
  tabDetalhes.TabVisible(1) = False
  'exluir depois
  If Status = tpStatus_Incluir Then
    '
    optStatus(0).Value = True
    'Visible
    optStatus(0).Visible = False
    optStatus(1).Visible = False
    Label5(5).Visible = False
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objAgencia = New busElite.clsAgencia
    Set objRs = objAgencia.SelecionarAgenciaPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      txtNome.Text = objRs.Fields("NOME").Value & ""
      txtTelefone1.Text = objRs.Fields("TELEFONE1").Value & ""
      txtTelefone2.Text = objRs.Fields("TELEFONE2").Value & ""
      txtTelefone3.Text = objRs.Fields("TELEFONE3").Value & ""
      txtRua.Text = objRs.Fields("ENDRUA").Value & ""
      txtNumero.Text = objRs.Fields("ENDNUMERO").Value & ""
      txtComplemento.Text = objRs.Fields("ENDCOMPLEMENTO").Value & ""
      txtEstado.Text = objRs.Fields("ENDESTADO").Value & ""
      INCLUIR_VALOR_NO_MASK mskCep, objRs.Fields("ENDCEP").Value, TpMaskSemMascara
      txtBairro.Text = objRs.Fields("ENDBAIRRO").Value & ""
      txtCidade.Text = objRs.Fields("ENDCIDADE").Value & ""
      txtObservacao.Text = objRs.Fields("OBSERVACAO").Value & ""
      txtNumero.Text = objRs.Fields("ENDNUMERO").Value & ""
      If objRs.Fields("STATUS").Value & "" = "A" Then
        optStatus(0).Value = True
        optStatus(1).Value = False
      ElseIf objRs.Fields("STATUS").Value & "" = "I" Then
        optStatus(0).Value = False
        optStatus(1).Value = True
      Else
        optStatus(0).Value = False
        optStatus(1).Value = False
      End If
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objAgencia = Nothing
    'Visible
    optStatus(0).Visible = True
    optStatus(1).Visible = True
    Label5(5).Visible = True
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = True
    tabDetalhes.TabVisible(1) = True
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


Private Sub txtBairro_GotFocus()
  Seleciona_Conteudo_Controle txtBairro
End Sub
Private Sub txtBairro_LostFocus()
  Pintar_Controle txtBairro, tpCorContr_Normal
End Sub

Private Sub txtCidade_GotFocus()
  Seleciona_Conteudo_Controle txtCidade
End Sub
Private Sub txtCidade_LostFocus()
  Pintar_Controle txtCidade, tpCorContr_Normal
End Sub
Private Sub txtObservacao_GotFocus()
  Seleciona_Conteudo_Controle txtObservacao
End Sub
Private Sub txtObservacao_LostFocus()
  Pintar_Controle txtObservacao, tpCorContr_Normal
End Sub

Private Sub txtComplemento_GotFocus()
  Seleciona_Conteudo_Controle txtComplemento
End Sub
Private Sub txtComplemento_LostFocus()
  Pintar_Controle txtComplemento, tpCorContr_Normal
End Sub

Private Sub txtEstado_GotFocus()
  Seleciona_Conteudo_Controle txtEstado
End Sub
Private Sub txtEstado_LostFocus()
  Pintar_Controle txtEstado, tpCorContr_Normal
End Sub

Private Sub txtNome_GotFocus()
  Seleciona_Conteudo_Controle txtNome
End Sub
Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
End Sub
Private Sub txtNumero_GotFocus()
  Seleciona_Conteudo_Controle txtNumero
End Sub
Private Sub txtNumero_LostFocus()
  Pintar_Controle txtNumero, tpCorContr_Normal
End Sub

Private Sub txtRua_GotFocus()
  Seleciona_Conteudo_Controle txtRua
End Sub
Private Sub txtRua_LostFocus()
  Pintar_Controle txtRua, tpCorContr_Normal
End Sub
Private Sub txtTelefone1_GotFocus()
  Seleciona_Conteudo_Controle txtTelefone1
End Sub
Private Sub txtTelefone1_LostFocus()
  Pintar_Controle txtTelefone1, tpCorContr_Normal
End Sub

Private Sub txtTelefone2_GotFocus()
  Seleciona_Conteudo_Controle txtTelefone2
End Sub
Private Sub txtTelefone2_LostFocus()
  Pintar_Controle txtTelefone2, tpCorContr_Normal
End Sub

Private Sub txtTelefone3_GotFocus()
  Seleciona_Conteudo_Controle txtTelefone3
End Sub
Private Sub txtTelefone3_LostFocus()
  Pintar_Controle txtTelefone3, tpCorContr_Normal
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    grdCNPJ.Enabled = False
    pictrava(0).Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco txtNome
  Case 1
    grdCNPJ.Enabled = True
    pictrava(0).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = True
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    'Montar RecordSet
    CNPJ_COLUNASMATRIZ = grdCNPJ.Columns.Count
    CNPJ_LINHASMATRIZ = 0
    CNPJ_MontaMatriz
    grdCNPJ.Bookmark = Null
    grdCNPJ.ReBind
    grdCNPJ.ApproxCount = CNPJ_LINHASMATRIZ
    '
    SetarFoco grdCNPJ
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "Elite.frmAgenciaInc.tabDetalhes"
  AmpN
End Sub


Private Sub grdCNPJ_UnboundReadDataEx( _
     ByVal RowBuf As TrueDBGrid60.RowBuffer, _
    StartLocation As Variant, ByVal Offset As Long, _
    ApproximatePosition As Long)
  ' UnboundReadData is fired by an unbound grid whenever
  ' it requires data for display. This event will fire
  ' when the grid is first shown, when Refresh or ReBind
  ' is used, when the grid is scrolled, and after a
  ' record in the grid is modified and the user commits
  ' the change by moving off of the current row. The
  ' grid fetches data in "chunks", and the number of rows
  ' the grid is asking for is given by RowBuf.RowCount.
  ' RowBuf is the row buffer where you place the data
  ' the bookmarks for the rows that the grid is
  ' requesting to display. It will also hold the number
  ' of rows that were successfully supplied to the grid.
  ' StartLocation is a vrtBookmark which, together with
  ' Offset, specifies the row for the programmer to start
  ' transferring data. A StartLocation of Null indicates
  ' a request for data from BOF or EOF.
  ' Offset specifies the relative position (from
  ' StartLocation) of the row for the programmer to start
  ' transferring data. A positive number indicates a
  ' forward relative position while a negative number
  ' indicates a backward relative position. Regardless
  ' of whether the rows to be read are before or after
  ' StartLocation, rows are always fetched going forward
  ' (this is why there is no ReadPriorRows parameter to
  ' the procedure).
  ' If you page down on the grid, for instance, the new
  ' top row of the grid will have an index greater than
  ' the StartLocation (Offset > 0). If you page up on
  ' the grid, the new index is less than that of
  ' StartLocation, so Offset < 0. If StartLocation is
  ' a vrtBookmark to row N, the grid always asks for row
  ' data in the following order:
  '   (N + Offset), (N + Offset + 1), (N + Offset + 2)...
  ' ApproximatePosition is a value you can set to indicate
  ' the ordinal position of (StartLocation + Offset).
  ' Setting this variable will enhance the ability of the
  ' grid to display its vertical scroll bar accurately.
  ' If the exact ordinal position of the new location is
  ' not known, you can set it to a reasonable,
  ' approximate value, or just ignore this parameter.

  On Error GoTo trata
  '
  Dim intColIndex      As Integer
  Dim intJ             As Integer
  Dim intRowsFetched   As Integer
  Dim intI             As Long
  Dim lngNewPosition   As Long
  Dim vrtBookmark      As Variant
  '
  intRowsFetched = 0
  For intI = 0 To RowBuf.RowCount - 1
    ' Get the vrtBookmark of the next available row
    vrtBookmark = GetRelativeBookmarkGeral(StartLocation, _
               Offset + intI, CNPJ_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, CNPJ_COLUNASMATRIZ, CNPJ_LINHASMATRIZ, CNPJ_Matriz)
    Next intJ

    ' Set the vrtBookmark for the row
    RowBuf.Bookmark(intI) = vrtBookmark

    ' Increment the count of fetched rows
    intRowsFetched = intRowsFetched + 1
  Next intI

' Tell the grid how many rows were fetched
  RowBuf.RowCount = intRowsFetched

' Set the approximate scroll bar position. Only
' nonnegative values of IndexFromBookmark() are valid.
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, CNPJ_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmAgenciaInc.grdCNPJ_UnboundReadDataEx]"
End Sub

Public Sub CNPJ_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busElite.clsGeral
  '
  On Error GoTo trata

  Set clsGer = New busElite.clsGeral
  '
  strSql = "SELECT AGENCIACNPJ.PKID, dbo.formataCNPJ(AGENCIACNPJ.CNPJ) " & _
          "FROM AGENCIACNPJ " & _
          "WHERE AGENCIACNPJ.AGENCIAID = " & lngPKID & _
          " ORDER BY AGENCIACNPJ.CNPJ"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    CNPJ_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim CNPJ_Matriz(0 To CNPJ_COLUNASMATRIZ - 1, 0 To CNPJ_LINHASMATRIZ - 1)
  Else
    ReDim CNPJ_Matriz(0 To CNPJ_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To CNPJ_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To CNPJ_COLUNASMATRIZ - 1    'varre as colunas
          CNPJ_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

