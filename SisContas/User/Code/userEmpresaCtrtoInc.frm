VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmUserEmpresaCtrtoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de empresa para contrato"
   ClientHeight    =   4875
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4875
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4875
      Left            =   8430
      ScaleHeight     =   4875
      ScaleWidth      =   1860
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2025
         Left            =   90
         ScaleHeight     =   1965
         ScaleWidth      =   1605
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2700
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
      Height          =   4635
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   8176
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userEmpresaCtrtoInc.frx":0000
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
         Left            =   150
         TabIndex        =   11
         Top             =   450
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2385
            Index           =   0
            Left            =   90
            ScaleHeight     =   2385
            ScaleWidth      =   7575
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   180
            Width           =   7575
            Begin VB.TextBox txtDescricao 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   2
               Text            =   "txtDescricao"
               Top             =   720
               Width           =   6075
            End
            Begin VB.ComboBox cboEmpresa 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   390
               Width           =   6075
            End
            Begin VB.TextBox txtContrato 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtContrato"
               Top             =   90
               Width           =   6075
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1290
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   1320
               Width           =   2235
               Begin VB.OptionButton optStatus 
                  Caption         =   "Inativo"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   5
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optStatus 
                  Caption         =   "Ativo"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   4
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   825
               End
            End
            Begin MSMask.MaskEdBox mskValor 
               Height          =   255
               Left            =   1320
               TabIndex        =   3
               Top             =   1020
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label6 
               Caption         =   "Descrição"
               Height          =   255
               Index           =   26
               Left            =   90
               TabIndex        =   18
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Contrato"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   17
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Valor contrato"
               Height          =   195
               Index           =   21
               Left            =   60
               TabIndex        =   16
               Top             =   1035
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Status"
               Height          =   315
               Index           =   5
               Left            =   90
               TabIndex        =   14
               Top             =   1350
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Empresa"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   13
               Top             =   450
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserEmpresaCtrtoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngCONTRATOID            As Long
Public strDescrContrato         As String

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Valor EmpresaCtrto
  LimparCampoTexto txtContrato
  LimparCampoCombo cboEmpresa
  LimparCampoTexto txtDescricao
  LimparCampoMask mskValor
  optStatus(0).Value = False
  optStatus(1).Value = False
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserEmpresaCtrtoInc.LimparCampos]", _
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
  Dim objEmpresaCtrto           As busSisContas.clsEmpresaCtrto
  Dim objGeral                  As busSisContas.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim strStatus                 As String
  Dim lngEMPRESAID              As Long
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisContas.clsGeral
  Set objEmpresaCtrto = New busSisContas.clsEmpresaCtrto
  'Status
  If optStatus(0).Value Then
    strStatus = "A"
  Else
    strStatus = "I"
  End If
  'Obtem Tipo Empresa
  Set objGeral = New busSisContas.clsGeral
  strSql = "SELECT PKID FROM EMPRESA WHERE NOME = " & Formata_Dados(cboEmpresa.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  lngEMPRESAID = 0
  If Not objRs.EOF Then
    lngEMPRESAID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Validar se empresa já associada ao contrato
  strSql = "SELECT * FROM EMPRESACTRTO " & _
    " WHERE EMPRESAID = " & Formata_Dados(lngEMPRESAID, tpDados_Longo) & _
    " AND DESCRICAO = " & Formata_Dados(txtDescricao, tpDados_Texto) & _
    " AND PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtDescricao, tpCorContr_Erro
    TratarErroPrevisto "Descrição já cadatrada"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objEmpresaCtrto = Nothing
    cmdOk.Enabled = True
    SetarFoco cboEmpresa
    tabDetalhes.Tab = 1
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar EmpresaCtrto
    objEmpresaCtrto.AlterarEmpresaCtrto lngPKID, _
                                        lngEMPRESAID, _
                                        txtDescricao.Text, _
                                        mskValor.Text, _
                                        strStatus
    '
    blnRetorno = True
    blnFechar = True
    Unload Me
  ElseIf Status = tpStatus_Incluir Then
    'Inserir EmpresaCtrto
    objEmpresaCtrto.InserirEmpresaCtrto lngCONTRATOID, _
                                        lngEMPRESAID, _
                                        txtDescricao.Text, _
                                        mskValor.Text
    'Selecionar plano cadastrado
    Set objRs = objEmpresaCtrto.SelecionarEmpresaCtrtoPelaDescricao(lngCONTRATOID, _
                                                                    txtDescricao.Text)
    If Not objRs.EOF Then
      'Captura dados para entrar em modo de alteração
      lngPKID = objRs.Fields("PKID")
      Status = tpStatus_Alterar
      'Reload na tela
      Form_Load
      'Acerta tabs
      tabDetalhes.Tab = 0
      blnRetorno = True
    Else
      blnRetorno = True
      blnFechar = True
      Unload Me
    End If
  End If
  Set objEmpresaCtrto = Nothing
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
  If Not Valida_String(cboEmpresa, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar a empresa" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(txtDescricao, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher uma descrição" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskValor, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o valor para o contrato" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Option(optStatus, blnSetarFocoControle) Then
    strMsg = strMsg & "Slecionar o status" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserEmpresaCtrtoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserEmpresaCtrtoInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco cboEmpresa
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserEmpresaCtrtoInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objEmpresaCtrto           As busSisContas.clsEmpresaCtrto
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 5355
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  txtContrato.Text = strDescrContrato
  strSql = "SELECT NOME FROM EMPRESA " & _
      " WHERE PARCEIROID = " & Formata_Dados(glParceiroId, tpDados_Longo) & _
      " ORDER BY NOME"
  PreencheCombo cboEmpresa, strSql, False, True
  '
  tabDetalhes_Click 1
  If Status = tpStatus_Incluir Then
    '
    optStatus(0).Value = True
    'Visible
    optStatus(0).Visible = False
    optStatus(1).Visible = False
    Label5(5).Visible = False
    '
    tabDetalhes.TabEnabled(0) = True
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objEmpresaCtrto = New busSisContas.clsEmpresaCtrto
    Set objRs = objEmpresaCtrto.SelecionarEmpresaCtrtoPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      INCLUIR_VALOR_NO_COMBO objRs.Fields("NOME_EMPRESA").Value & "", cboEmpresa
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR").Value & "", TpMaskMoeda
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
    Set objEmpresaCtrto = Nothing
    'Visible
    optStatus(0).Visible = True
    optStatus(1).Visible = True
    Label5(5).Visible = True
    '
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


Private Sub mskValor_GotFocus()
  Seleciona_Conteudo_Controle mskValor
End Sub
Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    pictrava(0).Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    '
    SetarFoco cboEmpresa
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "SisContas.frmUserEmpresaCtrtoInc.tabDetalhes"
  AmpN
End Sub


Private Sub cboEmpresa_LostFocus()
  Pintar_Controle cboEmpresa, tpCorContr_Normal
End Sub

Private Sub txtDescricao_GotFocus()
  Seleciona_Conteudo_Controle txtDescricao
End Sub
Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub
