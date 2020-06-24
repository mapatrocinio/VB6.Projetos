VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmUserBeneficiarioInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de beneficiário para associado/convênio"
   ClientHeight    =   3720
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3720
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   3720
      Left            =   8430
      ScaleHeight     =   3720
      ScaleWidth      =   1860
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   90
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   1500
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   150
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   3465
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   6112
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userBeneficiarioInc.frx":0000
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
         Height          =   2895
         Left            =   120
         TabIndex        =   14
         Top             =   420
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2595
            Index           =   0
            Left            =   90
            ScaleHeight     =   2595
            ScaleWidth      =   7575
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   180
            Width           =   7575
            Begin VB.ComboBox cboGrauParentesco 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   1680
               Width           =   6105
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   5190
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   1050
               Width           =   2235
               Begin VB.OptionButton optSexo 
                  Caption         =   "Masculino"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   4
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1065
               End
               Begin VB.OptionButton optSexo 
                  Caption         =   "Feminino"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   5
                  Top             =   0
                  Width           =   1095
               End
            End
            Begin VB.TextBox txtIdentidade 
               Height          =   285
               Left            =   4830
               MaxLength       =   20
               TabIndex        =   7
               Text            =   "txtIdentidade"
               Top             =   1320
               Width           =   2565
            End
            Begin VB.TextBox txtNome 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   2
               Text            =   "txtNome"
               Top             =   750
               Width           =   6075
            End
            Begin VB.TextBox txtConvenio 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   1
               TabStop         =   0   'False
               Text            =   "txtConvenio"
               Top             =   420
               Width           =   6075
            End
            Begin VB.TextBox txtAssociado 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtAssociado"
               Top             =   90
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskDtNascimento 
               Height          =   255
               Left            =   1320
               TabIndex        =   6
               Top             =   1380
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskCpf 
               Height          =   255
               Left            =   1320
               TabIndex        =   3
               Top             =   1080
               Width           =   1635
               _ExtentX        =   2884
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   14
               Mask            =   "###.###.###-##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Grau Parentesco"
               Height          =   195
               Index           =   33
               Left            =   60
               TabIndex        =   24
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Sexo"
               Height          =   315
               Index           =   6
               Left            =   3870
               TabIndex        =   23
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label7 
               Caption         =   "Dt. Nascimento"
               Height          =   255
               Index           =   1
               Left            =   60
               TabIndex        =   22
               Top             =   1380
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "CPF"
               Height          =   195
               Index           =   5
               Left            =   60
               TabIndex        =   21
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Identidade"
               Height          =   195
               Index           =   39
               Left            =   3870
               TabIndex        =   20
               Top             =   1365
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Nome"
               Height          =   195
               Index           =   4
               Left            =   60
               TabIndex        =   18
               Top             =   795
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Convênio"
               Height          =   195
               Index           =   3
               Left            =   60
               TabIndex        =   17
               Top             =   465
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Associado"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   16
               Top             =   135
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserBeneficiarioInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngTABCONVASSOCID        As Long
Public strNomeAssociado         As String
Public strNomeConvenio          As String

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Valor Beneficiario
  LimparCampoTexto txtAssociado
  LimparCampoTexto txtConvenio
  LimparCampoTexto txtNome
  LimparCampoMask mskCpf
  optSexo(0).Value = False
  optSexo(1).Value = False
  LimparCampoMask mskDtNascimento
  LimparCampoTexto txtIdentidade
  LimparCampoCombo cboGrauParentesco
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserBeneficiarioInc.LimparCampos]", _
            Err.Description
End Sub


Private Sub cboGrauParentesco_LostFocus()
  Pintar_Controle cboGrauParentesco, tpCorContr_Normal
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
  Dim objBeneficiario           As busApler.clsBeneficiario
  Dim objGeral                  As busApler.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngGRAUPARENTESCOID       As Long
  Dim strSexo                   As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  'Status
  If optSexo(0).Value Then
    strSexo = "M"
  ElseIf optSexo(1).Value Then
    strSexo = "F"
  Else
    strSexo = ""
  End If
  Set objGeral = New busApler.clsGeral
  Set objBeneficiario = New busApler.clsBeneficiario
  'GRAU DE PARENTESCO
  lngGRAUPARENTESCOID = 0
  strSql = "SELECT PKID FROM GRAUPARENTESCO WHERE DESCRICAO = " & Formata_Dados(cboGrauParentesco.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngGRAUPARENTESCOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'Validar se valor plano já cadastrado
  strSql = "SELECT * FROM BENEFICIARIO " & _
    " WHERE BENEFICIARIO.TABCONVASSOCID = " & Formata_Dados(lngTABCONVASSOCID, tpDados_Longo) & _
    " AND BENEFICIARIO.NOME = " & Formata_Dados(txtNome.Text, tpDados_Texto) & _
    " AND BENEFICIARIO.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtNome, tpCorContr_Erro
    TratarErroPrevisto "Beneficiário já associado ao associado/convênio"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objBeneficiario = Nothing
    cmdOk.Enabled = True
    SetarFoco txtNome
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Beneficiario
    objBeneficiario.AlterarBeneficiario lngPKID, _
                                        lngGRAUPARENTESCOID, _
                                        txtNome.Text, _
                                        IIf(mskCpf.ClipText = "", "", mskCpf.ClipText), _
                                        txtIdentidade.Text, _
                                        IIf(mskDtNascimento.ClipText = "", "", mskDtNascimento.Text), _
                                        strSexo
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Beneficiario
    objBeneficiario.InserirBeneficiario lngTABCONVASSOCID, _
                                        lngGRAUPARENTESCOID, _
                                        txtNome.Text, _
                                        IIf(mskCpf.ClipText = "", "", mskCpf.ClipText), _
                                        txtIdentidade.Text, _
                                        IIf(mskDtNascimento.ClipText = "", "", mskDtNascimento.Text), _
                                        strSexo
  End If
  Set objBeneficiario = Nothing
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
  If Not Valida_String(txtNome, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o nome" & vbCrLf
    tabDetalhes.Tab = 0
  End If
'''  If Len(Trim(mskCpf.ClipText)) = 0 Then
'''    strMsg = strMsg & "Informar o CPF" & vbCrLf
'''    Pintar_Controle mskCpf, tpCorContr_Erro
'''    SetarFoco mskCpf
'''    tabDetalhes.Tab = 0
'''    blnSetarFocoControle = False
'''  End If
  If Len(Trim(mskCpf.ClipText)) > 0 Then
    If Not TestaCPF(mskCpf.ClipText) Then
      strMsg = strMsg & "Informar o CPF válido" & vbCrLf
      Pintar_Controle mskCpf, tpCorContr_Erro
      SetarFoco mskCpf
      tabDetalhes.Tab = 0
      blnSetarFocoControle = False
    End If
  End If
  If Not Valida_Option(optSexo, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o sexo" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskDtNascimento, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de nascimento válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
'  If Not Valida_String(cboGrauParentesco, TpObrigatorio, blnSetarFocoControle) Then
'    strMsg = strMsg & "Slecionar o grau de parentesco" & vbCrLf
'    tabDetalhes.Tab = 0
'  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserBeneficiarioInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserBeneficiarioInc.ValidaCampos]", _
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
  TratarErro Err.Number, Err.Description, "[frmUserBeneficiarioInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objBeneficiario         As busApler.clsBeneficiario
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 4200
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  'Grau Parentesco
  strSql = "Select DESCRICAO from GRAUPARENTESCO ORDER BY DESCRICAO"
  PreencheCombo cboGrauParentesco, strSql, False, True
  '
  txtAssociado.Text = strNomeAssociado
  txtConvenio.Text = strNomeConvenio
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objBeneficiario = New busApler.clsBeneficiario
    Set objRs = objBeneficiario.SelecionarBeneficiarioPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      txtNome.Text = objRs.Fields("NOME").Value & ""
      INCLUIR_VALOR_NO_MASK mskCpf, objRs.Fields("CPF").Value, TpMaskSemMascara
      If objRs.Fields("SEXO").Value & "" = "M" Then
        optSexo(0).Value = True
        optSexo(1).Value = False
      ElseIf objRs.Fields("SEXO").Value & "" = "F" Then
        optSexo(0).Value = False
        optSexo(1).Value = True
      Else
        optSexo(0).Value = False
        optSexo(1).Value = False
      End If
      INCLUIR_VALOR_NO_MASK mskDtNascimento, objRs.Fields("DATANASCIMENTO").Value, TpMaskData
      txtIdentidade.Text = objRs.Fields("IDENTIDADE").Value & ""
      If objRs.Fields("DESCR_GRAUPARENTESCO").Value & "" <> "" Then
        cboGrauParentesco.Text = objRs.Fields("DESCR_GRAUPARENTESCO").Value & ""
      End If
      
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objBeneficiario = Nothing
    '
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

Private Sub mskCpf_GotFocus()
  Seleciona_Conteudo_Controle mskCpf
End Sub
Private Sub mskCpf_LostFocus()
  Pintar_Controle mskCpf, tpCorContr_Normal
End Sub

Private Sub mskDtNascimento_GotFocus()
  Seleciona_Conteudo_Controle mskDtNascimento
End Sub
Private Sub mskDtNascimento_LostFocus()
  Pintar_Controle mskDtNascimento, tpCorContr_Normal
End Sub

Private Sub txtAssociado_GotFocus()
  Seleciona_Conteudo_Controle txtAssociado
End Sub

Private Sub txtConvenio_GotFocus()
  Seleciona_Conteudo_Controle txtConvenio
End Sub

Private Sub txtIdentidade_GotFocus()
  Seleciona_Conteudo_Controle txtIdentidade
End Sub
Private Sub txtIdentidade_LostFocus()
  Pintar_Controle txtIdentidade, tpCorContr_Normal
End Sub

Private Sub txtNome_GotFocus()
  Seleciona_Conteudo_Controle txtNome
End Sub
Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
End Sub

