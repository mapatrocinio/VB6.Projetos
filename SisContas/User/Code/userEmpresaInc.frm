VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmUserEmpresaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de empresa"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5265
      Left            =   8520
      ScaleHeight     =   5265
      ScaleWidth      =   1860
      TabIndex        =   37
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2115
         Left            =   90
         ScaleHeight     =   2055
         ScaleWidth      =   1605
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   2700
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   36
            Top             =   1005
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   35
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5055
      Left            =   120
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   8916
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da empresa"
      TabPicture(0)   =   "userEmpresaInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraProf"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Endereço"
      TabPicture(1)   =   "userEmpresaInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin VB.Frame Frame1 
         Height          =   4575
         Left            =   -74865
         TabIndex        =   53
         Top             =   360
         Width           =   7935
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   4380
            Index           =   1
            Left            =   45
            ScaleHeight     =   4380
            ScaleWidth      =   7695
            TabIndex        =   54
            TabStop         =   0   'False
            Top             =   135
            Width           =   7695
            Begin VB.CheckBox chkEndCobIgualEndCorr 
               Caption         =   "Endereço de Cobrança igual endereço de correspondência ?"
               Height          =   255
               Left            =   1350
               TabIndex        =   26
               Top             =   2280
               Width           =   5685
            End
            Begin VB.TextBox txtPaisCob 
               Height          =   285
               Left            =   4095
               MaxLength       =   30
               TabIndex        =   34
               Text            =   "txtPaisCob"
               Top             =   3840
               Width           =   3390
            End
            Begin VB.TextBox txtEstadoCob 
               Height          =   285
               Left            =   1335
               MaxLength       =   2
               TabIndex        =   33
               Text            =   "txtEstado"
               Top             =   3840
               Width           =   510
            End
            Begin VB.TextBox txtCidadeCob 
               Height          =   285
               Left            =   1335
               MaxLength       =   50
               TabIndex        =   32
               Text            =   "txtCidadeCob"
               Top             =   3525
               Width           =   6135
            End
            Begin VB.TextBox txtBairroCob 
               Height          =   285
               Left            =   4080
               MaxLength       =   50
               TabIndex        =   31
               Text            =   "txtBairroCob"
               Top             =   3210
               Width           =   3390
            End
            Begin VB.TextBox txtCepCob 
               Height          =   285
               Left            =   1335
               MaxLength       =   8
               TabIndex        =   30
               Text            =   "txtCepCo"
               Top             =   3210
               Width           =   1455
            End
            Begin VB.TextBox txtComplementoCob 
               Height          =   285
               Left            =   4080
               MaxLength       =   50
               TabIndex        =   29
               Text            =   "txtComplementoCob"
               Top             =   2895
               Width           =   3390
            End
            Begin VB.TextBox txtNumeroCob 
               Height          =   285
               Left            =   1335
               MaxLength       =   10
               TabIndex        =   28
               Text            =   "txtNumeroC"
               Top             =   2895
               Width           =   1455
            End
            Begin VB.TextBox txtRuaCob 
               Height          =   285
               Left            =   1335
               MaxLength       =   50
               TabIndex        =   27
               Text            =   "txtRuaCob"
               Top             =   2580
               Width           =   6135
            End
            Begin VB.TextBox txtPais 
               Height          =   285
               Left            =   4095
               MaxLength       =   30
               TabIndex        =   25
               Text            =   "txtPais"
               Top             =   1620
               Width           =   3390
            End
            Begin VB.TextBox txtEstado 
               Height          =   285
               Left            =   1335
               MaxLength       =   2
               TabIndex        =   24
               Text            =   "txtEstado"
               Top             =   1620
               Width           =   510
            End
            Begin VB.TextBox txtCidade 
               Height          =   285
               Left            =   1335
               MaxLength       =   50
               TabIndex        =   23
               Text            =   "txtCidade"
               Top             =   1305
               Width           =   6135
            End
            Begin VB.TextBox txtBairro 
               Height          =   285
               Left            =   4080
               MaxLength       =   50
               TabIndex        =   22
               Text            =   "txtBairro"
               Top             =   990
               Width           =   3390
            End
            Begin VB.TextBox txtCep 
               Height          =   285
               Left            =   1335
               MaxLength       =   8
               TabIndex        =   21
               Text            =   "txtCep"
               Top             =   990
               Width           =   1455
            End
            Begin VB.TextBox txtComplemento 
               Height          =   285
               Left            =   4080
               MaxLength       =   50
               TabIndex        =   20
               Text            =   "txtComplemento"
               Top             =   675
               Width           =   3390
            End
            Begin VB.TextBox txtNumero 
               Height          =   285
               Left            =   1335
               MaxLength       =   10
               TabIndex        =   19
               Text            =   "txtNumero"
               Top             =   675
               Width           =   1455
            End
            Begin VB.TextBox txtRua 
               Height          =   285
               Left            =   1335
               MaxLength       =   50
               TabIndex        =   18
               Text            =   "txtRua"
               Top             =   360
               Width           =   6135
            End
            Begin VB.Label Label6 
               Caption         =   "País"
               Height          =   255
               Index           =   25
               Left            =   2880
               TabIndex        =   72
               Top             =   3840
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Estado"
               Height          =   255
               Index           =   24
               Left            =   135
               TabIndex        =   71
               Top             =   3840
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Cidade"
               Height          =   255
               Index           =   23
               Left            =   135
               TabIndex        =   70
               Top             =   3525
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Bairro"
               Height          =   255
               Index           =   22
               Left            =   2880
               TabIndex        =   69
               Top             =   3210
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Cep"
               Height          =   255
               Index           =   21
               Left            =   150
               TabIndex        =   68
               Top             =   3210
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Complemento"
               Height          =   255
               Index           =   20
               Left            =   2880
               TabIndex        =   67
               Top             =   2895
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Número"
               Height          =   255
               Index           =   19
               Left            =   135
               TabIndex        =   66
               Top             =   2895
               Width           =   1215
            End
            Begin VB.Label Label2 
               Caption         =   "Endereço de Cobrança"
               Height          =   195
               Left            =   135
               TabIndex        =   65
               Top             =   1935
               Width           =   2310
            End
            Begin VB.Line Line2 
               X1              =   135
               X2              =   7470
               Y1              =   2160
               Y2              =   2160
            End
            Begin VB.Label Label6 
               Caption         =   "Rua"
               Height          =   255
               Index           =   18
               Left            =   135
               TabIndex        =   64
               Top             =   2580
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "País"
               Height          =   255
               Index           =   17
               Left            =   2880
               TabIndex        =   63
               Top             =   1620
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Estado"
               Height          =   255
               Index           =   16
               Left            =   135
               TabIndex        =   62
               Top             =   1620
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Cidade"
               Height          =   255
               Index           =   15
               Left            =   135
               TabIndex        =   61
               Top             =   1305
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Bairro"
               Height          =   255
               Index           =   14
               Left            =   2880
               TabIndex        =   60
               Top             =   990
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Cep"
               Height          =   255
               Index           =   13
               Left            =   135
               TabIndex        =   59
               Top             =   990
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Complemento"
               Height          =   255
               Index           =   12
               Left            =   2880
               TabIndex        =   58
               Top             =   675
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Número"
               Height          =   255
               Index           =   11
               Left            =   135
               TabIndex        =   57
               Top             =   675
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "Endereço de Correspondencia"
               Height          =   195
               Left            =   135
               TabIndex        =   56
               Top             =   45
               Width           =   2310
            End
            Begin VB.Line Line1 
               X1              =   135
               X2              =   7470
               Y1              =   270
               Y2              =   270
            End
            Begin VB.Label Label6 
               Caption         =   "Rua"
               Height          =   255
               Index           =   10
               Left            =   135
               TabIndex        =   55
               Top             =   360
               Width           =   1215
            End
         End
      End
      Begin VB.Frame fraProf 
         Height          =   4215
         Left            =   120
         TabIndex        =   40
         Top             =   360
         Width           =   7935
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   3930
            Index           =   0
            Left            =   120
            ScaleHeight     =   3930
            ScaleWidth      =   7695
            TabIndex        =   41
            TabStop         =   0   'False
            Top             =   240
            Width           =   7695
            Begin VB.TextBox txtNroContrato 
               Height          =   285
               Left            =   1275
               MaxLength       =   20
               TabIndex        =   2
               Text            =   "txtNroContrato"
               Top             =   720
               Width           =   2175
            End
            Begin VB.TextBox txtNomeFantasia 
               Height          =   285
               Left            =   1260
               MaxLength       =   100
               TabIndex        =   1
               Text            =   "txtNomeFantasia"
               Top             =   390
               Width           =   6135
            End
            Begin VB.OptionButton optCredito 
               Caption         =   "Bloqueado"
               Height          =   255
               Index           =   1
               Left            =   2430
               TabIndex        =   17
               Top             =   3570
               Width           =   1065
            End
            Begin VB.OptionButton optCredito 
               Caption         =   "Liberado"
               Height          =   255
               Index           =   0
               Left            =   1290
               TabIndex        =   16
               Top             =   3570
               Width           =   1065
            End
            Begin VB.ComboBox cboTipo 
               Height          =   315
               Left            =   1290
               Style           =   2  'Dropdown List
               TabIndex        =   5
               Top             =   1350
               Width           =   2625
            End
            Begin VB.TextBox txtInscrMunicipal 
               Height          =   285
               Left            =   5250
               MaxLength       =   30
               TabIndex        =   12
               Text            =   "txtInscrMunicipal"
               Top             =   2340
               Width           =   2175
            End
            Begin VB.TextBox txtInscrEstadual 
               Height          =   285
               Left            =   1290
               MaxLength       =   30
               TabIndex        =   11
               Text            =   "txtInscrEstadual"
               Top             =   2340
               Width           =   2175
            End
            Begin VB.TextBox txtCNPJ 
               Height          =   285
               Left            =   5250
               MaxLength       =   30
               TabIndex        =   10
               Text            =   "txtCNPJ"
               Top             =   2010
               Width           =   2175
            End
            Begin VB.TextBox txtTelefone2 
               Height          =   285
               Left            =   1290
               MaxLength       =   20
               TabIndex        =   9
               Text            =   "txtTelefone2"
               Top             =   2025
               Width           =   2175
            End
            Begin VB.TextBox txtTelefone1 
               Height          =   285
               Left            =   5250
               MaxLength       =   20
               TabIndex        =   8
               Text            =   "txtTelefone1"
               Top             =   1680
               Width           =   2175
            End
            Begin VB.TextBox txtContato 
               Height          =   285
               Left            =   1290
               MaxLength       =   50
               TabIndex        =   13
               Text            =   "txtContato"
               Top             =   2640
               Width           =   6135
            End
            Begin VB.TextBox txtNome 
               Height          =   285
               Left            =   1260
               MaxLength       =   100
               TabIndex        =   0
               Text            =   "txtNome"
               Top             =   60
               Width           =   6135
            End
            Begin VB.TextBox txtObservacao 
               Height          =   285
               Left            =   1290
               MaxLength       =   100
               MultiLine       =   -1  'True
               TabIndex        =   14
               Text            =   "userEmpresaInc.frx":0038
               Top             =   2940
               Width           =   6135
            End
            Begin VB.TextBox txtTelefone 
               Height          =   285
               Left            =   1290
               MaxLength       =   20
               TabIndex        =   7
               Text            =   "txtTelefone"
               Top             =   1710
               Width           =   2175
            End
            Begin MSMask.MaskEdBox mskPercentual 
               Height          =   255
               Left            =   5250
               TabIndex        =   6
               Top             =   1350
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0;-#,##0"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDescDiaria 
               Height          =   255
               Left            =   1290
               TabIndex        =   15
               Top             =   3270
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   6
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskData 
               Height          =   255
               Index           =   0
               Left            =   1290
               TabIndex        =   3
               Top             =   1050
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskData 
               Height          =   255
               Index           =   1
               Left            =   5250
               TabIndex        =   4
               Top             =   1050
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label4 
               Caption         =   "Dt. Fim Contr."
               Height          =   255
               Left            =   4050
               TabIndex        =   78
               Top             =   1050
               Width           =   1275
            End
            Begin VB.Label Label12 
               Caption         =   "Dt. Início Contr."
               Height          =   255
               Left            =   60
               TabIndex        =   77
               Top             =   1050
               Width           =   1275
            End
            Begin VB.Label Label6 
               Caption         =   "Nro. Contrato"
               Height          =   255
               Index           =   27
               Left            =   60
               TabIndex        =   76
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Nome Fantasia"
               Height          =   255
               Index           =   26
               Left            =   90
               TabIndex        =   75
               Top             =   390
               Width           =   1215
            End
            Begin VB.Label Label3 
               Caption         =   "Crédito"
               Height          =   255
               Left            =   120
               TabIndex        =   74
               Top             =   3570
               Width           =   1095
            End
            Begin VB.Label Label17 
               Caption         =   "Desc. Diária (%)"
               Height          =   255
               Left            =   120
               TabIndex        =   73
               Top             =   3240
               Width           =   1455
            End
            Begin VB.Label lblPercentual 
               Caption         =   "Percentual"
               Height          =   255
               Left            =   4035
               TabIndex        =   52
               Top             =   1350
               Width           =   1455
            End
            Begin VB.Label Label6 
               Caption         =   "Tipo"
               Height          =   255
               Index           =   9
               Left            =   75
               TabIndex        =   51
               Top             =   1395
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Inscr. Municipal"
               Height          =   255
               Index           =   8
               Left            =   4035
               TabIndex        =   50
               Top             =   2340
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Inscr. Estadual"
               Height          =   255
               Index           =   7
               Left            =   120
               TabIndex        =   49
               Top             =   2295
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "CNPJ"
               Height          =   255
               Index           =   6
               Left            =   4050
               TabIndex        =   48
               Top             =   2025
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Telefone 2"
               Height          =   255
               Index           =   5
               Left            =   75
               TabIndex        =   47
               Top             =   2025
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Telefone 1"
               Height          =   255
               Index           =   4
               Left            =   4050
               TabIndex        =   46
               Top             =   1680
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Nome"
               Height          =   255
               Index           =   3
               Left            =   90
               TabIndex        =   45
               Top             =   60
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Observação"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   44
               Top             =   2925
               Width           =   1095
            End
            Begin VB.Label Label6 
               Caption         =   "Contato"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   43
               Top             =   2610
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Telefone"
               Height          =   255
               Index           =   1
               Left            =   75
               TabIndex        =   42
               Top             =   1710
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserEmpresaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngEMPRESAID               As Long
Public bRetorno                   As Boolean
Public bFechar                    As Boolean
Public sTitulo                    As String
Public intQuemChamou              As Integer
Private blnPrimeiraVez            As Boolean

Private Sub cboTipo_click()
  On Error Resume Next
  If cboTipo.Text = "Agência de Turismo" Then
    lblPercentual.Enabled = True
    mskPercentual.Enabled = True
  Else
    LimparCampoMask mskPercentual
    lblPercentual.Enabled = False
    mskPercentual.Enabled = False
  End If
End Sub

Private Sub cboTipo_LostFocus()
  Pintar_Controle cboTipo, tpCorContr_Normal
End Sub


Private Sub chkEndCobIgualEndCorr_Click()
  On Error GoTo trata
  If chkEndCobIgualEndCorr.Value = 0 Then
    'Selecionado
    HabDesEndCob True
  Else
    'Não selecionado
    HabDesEndCob False
    'Limpar os campos
    LimparCampoTexto txtRuaCob
    LimparCampoTexto txtNumeroCob
    LimparCampoTexto txtComplementoCob
    LimparCampoTexto txtCepCob
    LimparCampoTexto txtBairroCob
    LimparCampoTexto txtCidadeCob
    LimparCampoTexto txtEstadoCob
    LimparCampoTexto txtPaisCob
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source & ".[frmUserEmpresaInc.chkEndCobIgualEndCorr_Click]"
End Sub
Private Sub HabDesEndCob(blnHabDes As Boolean)
  On Error GoTo trata
  Label6(18).Enabled = blnHabDes
  txtRuaCob.Enabled = blnHabDes
  Label6(19).Enabled = blnHabDes
  txtNumeroCob.Enabled = blnHabDes
  Label6(20).Enabled = blnHabDes
  txtComplementoCob.Enabled = blnHabDes
  Label6(21).Enabled = blnHabDes
  txtCepCob.Enabled = blnHabDes
  Label6(22).Enabled = blnHabDes
  txtBairroCob.Enabled = blnHabDes
  Label6(23).Enabled = blnHabDes
  txtCidadeCob.Enabled = blnHabDes
  Label6(24).Enabled = blnHabDes
  txtEstadoCob.Enabled = blnHabDes
  Label6(25).Enabled = blnHabDes
  txtPaisCob.Enabled = blnHabDes
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserEmpresaInc.HabDesEndCob]", _
            Err.Description
End Sub

Private Sub cmdCancelar_Click()
  On Error GoTo trata
  bFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source & ".[frmUserEmpresaInc.cmdCancelar_Click]"
  Unload Me
End Sub

Private Sub cmdOK_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim objEmpresa              As busSisContas.clsEmpresa
  Dim objGeral                As busSisContas.clsGeral
  Dim strTipoEmpresaId        As String
  Dim strCredito              As String
  '
  Select Case tabDetalhes.Tab
  Case 0, 1 'Inclusão/Alteração de Empresa
    If Not ValidaCamposEmpresa Then Exit Sub
    'Valida se CPF do cliente já cadastrado
    Set objEmpresa = New busSisContas.clsEmpresa
    Set objRs = objEmpresa.ListarEmpresaPeloNome(txtNome.Text, _
                                                 lngEMPRESAID, _
                                                 glParceiroId)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objEmpresa = Nothing
      TratarErroPrevisto "Nome já cadastrado", "cmdOK_Click"
      Pintar_Controle txtNome, tpCorContr_Erro
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    'Obtem Tipo Empresa
    Set objGeral = New busSisContas.clsGeral
    strSql = "SELECT PKID FROM TIPOEMPRESA WHERE DESCRICAO = " & Formata_Dados(cboTipo.Text, tpDados_Texto, tpNulo_NaoAceita)
    Set objRs = objGeral.ExecutarSQL(strSql)
    strTipoEmpresaId = ""
    If Not objRs.EOF Then
      strTipoEmpresaId = objRs.Fields("PKID").Value
    End If
    objRs.Close
    If optCredito(0).Value = True Then
      strCredito = "L"
    ElseIf optCredito(1).Value = True Then
      strCredito = "B"
    Else
      strCredito = ""
    End If
    Set objRs = Nothing
    Set objGeral = Nothing
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      '
      objEmpresa.AlterarEmpresa lngEMPRESAID, txtNome.Text, _
                                txtTelefone.Text, _
                                txtTelefone1.Text, _
                                txtTelefone2.Text, _
                                txtContato.Text, _
                                txtObservacao.Text, _
                                strTipoEmpresaId, _
                                mskPercentual.ClipText, _
                                txtRua.Text, _
                                txtNumero.Text, _
                                txtComplemento.Text, _
                                txtCep.Text, _
                                txtBairro.Text, _
                                txtCidade.Text, _
                                txtEstado.Text, _
                                txtPais.Text, _
                                txtRuaCob.Text, _
                                txtNumeroCob.Text, _
                                txtComplementoCob.Text, _
                                txtCepCob.Text, _
                                txtBairroCob.Text, txtCidadeCob.Text, txtEstadoCob.Text, txtPaisCob.Text, _
                                txtCNPJ.Text, txtInscrEstadual.Text, txtInscrMunicipal.Text, mskDescDiaria.Text, strCredito, _
                                txtNomeFantasia.Text, txtNroContrato.Text, IIf(mskData(0).ClipText = "", "", mskData(0).Text), IIf(mskData(1).ClipText = "", "", mskData(1).Text), chkEndCobIgualEndCorr.Value
                                
                                
      bRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Pega Informações para inserir
      '
      objEmpresa.InserirEmpresa txtNome.Text, _
                                txtTelefone.Text, _
                                txtTelefone1.Text, _
                                txtTelefone2.Text, _
                                txtContato.Text, _
                                txtObservacao.Text, _
                                strTipoEmpresaId, _
                                mskPercentual.ClipText, _
                                txtRua.Text, _
                                txtNumero.Text, _
                                txtComplemento.Text, _
                                txtCep.Text, _
                                txtBairro.Text, _
                                txtCidade.Text, _
                                txtEstado.Text, _
                                txtPais.Text, _
                                txtRuaCob.Text, _
                                txtNumeroCob.Text, _
                                txtComplementoCob.Text, _
                                txtCepCob.Text, _
                                txtBairroCob.Text, txtCidadeCob.Text, txtEstadoCob.Text, txtPaisCob.Text, _
                                txtCNPJ.Text, txtInscrEstadual.Text, txtInscrMunicipal.Text, mskDescDiaria.Text, strCredito, _
                                txtNomeFantasia.Text, txtNroContrato.Text, IIf(mskData(0).ClipText = "", "", mskData(0).Text), IIf(mskData(1).ClipText = "", "", mskData(1).Text), chkEndCobIgualEndCorr.Value, glParceiroId
      '
      Status = tpStatus_Alterar
      '
      tabDetalhes.TabEnabled(1) = True
      '
      'tabDetalhes.Tab = 1
      'cmdIncluir_Click
      bRetorno = True
      'Exit Sub
    End If
    Set objEmpresa = Nothing

    bFechar = True
    Unload Me

  End Select
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source & ".[frmUserEmpresaInc.cmdOk_Click]"
End Sub

Private Function ValidaCamposEmpresa() As Boolean
  On Error GoTo trata
  Dim strMsg        As String
  Dim blnSetarFoco  As Boolean
  '
  blnSetarFoco = True
  If Not Valida_String(txtNome, TpObrigatorio, blnSetarFoco) Then
    strMsg = strMsg & "Informar o Nome" & vbCrLf
  End If
  If Not Valida_String(txtNomeFantasia, TpObrigatorio, blnSetarFoco) Then
    'strMsg = strMsg & "Informar o Nome Fantasia" & vbCrLf
  End If
  If Not Valida_Data(mskData(0), TpNaoObrigatorio, blnSetarFoco) Then
    strMsg = strMsg & "Informar a data de início do contrato válida" & vbCrLf
  End If
  If Not Valida_Data(mskData(1), TpNaoObrigatorio, blnSetarFoco) Then
    strMsg = strMsg & "Informar a data de fim do contrato válida" & vbCrLf
  End If
  If Not Valida_String(cboTipo, TpObrigatorio, blnSetarFoco) Then
    strMsg = strMsg & "Selecionar o Tipo" & vbCrLf
  End If
  If Not Valida_Moeda(mskPercentual, IIf(cboTipo.Text = "Agência de Turismo", TpObrigatorio, TpNaoObrigatorio), blnSetarFoco) Then
    strMsg = strMsg & "Informar o Percentual válido" & vbCrLf
  End If
  If Not Valida_Moeda(mskDescDiaria, TpObrigatorio, blnSetarFoco) Then
    strMsg = strMsg & "Informar o Desconto sobre a diária válido" & vbCrLf
  End If
  If Len(strMsg) = 0 Then
    If CLng(mskDescDiaria.Text) < 0 Or CLng(mskDescDiaria.Text) > 100 Then
      strMsg = strMsg & "Informar o Desconto sobre a diária na faixa de 0 a 100%" & vbCrLf
      blnSetarFoco = False
      SetarFoco mskDescDiaria
      Pintar_Controle mskDescDiaria, tpCorContr_Erro
    End If
  End If
  If Len(strMsg) = 0 Then
    If cboTipo.Text = "Agência de Turismo" Then
      If CLng(mskPercentual.Text) < 0 Or CLng(mskPercentual.Text) > 100 Then
        strMsg = strMsg & "Informar o Percentual na faixa de 0 a 100%" & vbCrLf
        blnSetarFoco = False
        SetarFoco mskPercentual
        Pintar_Controle mskPercentual, tpCorContr_Erro
      End If
    End If
  End If
  If Len(strMsg) = 0 Then
    If optCredito(0).Value = False And optCredito(1).Value = False Then
      strMsg = strMsg & "Selecionar crédito liberado ou bloqueado" & vbCrLf
      If blnSetarFoco = True Then SetarFoco optCredito(0)
      blnSetarFoco = False
    End If
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserEmpresaInc.ValidaCamposEmpresa]"
    ValidaCamposEmpresa = False
  Else
    ValidaCamposEmpresa = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserEmpresaInc.ValidaCamposEmpresa]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    tabDetalhes.Tab = 0
    SetarFoco txtNome
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source & ".[frmUserEmpresaInc.Form_Activate]"
End Sub




Private Sub mskData_GotFocus(Index As Integer)
  Selecionar_Conteudo mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub

Private Sub mskPercentual_GotFocus()
  Selecionar_Conteudo mskPercentual
End Sub

Private Sub mskPercentual_LostFocus()
  Pintar_Controle mskPercentual, tpCorContr_Normal
End Sub

Private Sub mskDescDiaria_GotFocus()
  Selecionar_Conteudo mskDescDiaria
End Sub

Private Sub mskDescDiaria_LostFocus()
  Pintar_Controle mskDescDiaria, tpCorContr_Normal
End Sub



Private Sub txtBairro_GotFocus()
  Selecionar_Conteudo txtBairro
End Sub

Private Sub txtBairroCob_GotFocus()
  Selecionar_Conteudo txtBairroCob
End Sub

Private Sub txtCep_GotFocus()
  Selecionar_Conteudo txtCep
End Sub

Private Sub txtCepCob_GotFocus()
  Selecionar_Conteudo txtCepCob
End Sub

Private Sub txtCidade_GotFocus()
  Selecionar_Conteudo txtCidade
End Sub

Private Sub txtCidadeCob_GotFocus()
  Selecionar_Conteudo txtCidadeCob
End Sub

Private Sub txtCNPJ_GotFocus()
  Selecionar_Conteudo txtCNPJ
End Sub

Private Sub txtComplemento_GotFocus()
  Selecionar_Conteudo txtComplemento
End Sub

Private Sub txtComplementoCob_GotFocus()
  Selecionar_Conteudo txtComplementoCob
End Sub

Private Sub txtContato_GotFocus()
  Selecionar_Conteudo txtContato
End Sub

Private Sub txtEstado_GotFocus()
  Selecionar_Conteudo txtEstado
End Sub

Private Sub txtEstadoCob_GotFocus()
  Selecionar_Conteudo txtEstadoCob
End Sub

Private Sub txtInscrEstadual_GotFocus()
  Selecionar_Conteudo txtInscrEstadual
End Sub

Private Sub txtInscrMunicipal_GotFocus()
  Selecionar_Conteudo txtInscrMunicipal
End Sub

Private Sub txtNome_GotFocus()
  Selecionar_Conteudo txtNome
End Sub

Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
End Sub

Private Sub txtNomeFantasia_GotFocus()
  Selecionar_Conteudo txtNomeFantasia
End Sub

Private Sub txtNomeFantasia_LostFocus()
  Pintar_Controle txtNomeFantasia, tpCorContr_Normal
End Sub


Private Sub Form_Load()
On Error GoTo trata
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim objEmpresa      As busSisContas.clsEmpresa
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 5640
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  strSql = "SELECT DESCRICAO FROM TIPOEMPRESA ORDER BY DESCRICAO"
  PreencheCombo cboTipo, strSql, False, True
  '
  tabDetalhes_Click 0
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    cboTipo_click
    LimparCampoTexto txtNome
    LimparCampoTexto txtNomeFantasia
    LimparCampoTexto txtNroContrato
    LimparCampoMask mskData(0)
    LimparCampoMask mskData(1)
    LimparCampoTexto txtTelefone
    LimparCampoTexto txtTelefone1
    LimparCampoTexto txtTelefone2
    LimparCampoTexto txtCNPJ
    LimparCampoTexto txtInscrEstadual
    LimparCampoTexto txtInscrMunicipal
    LimparCampoMask mskPercentual
    LimparCampoTexto txtContato
    LimparCampoTexto txtObservacao
    LimparCampoMask mskDescDiaria
    '
    LimparCampoTexto txtRua
    LimparCampoTexto txtNumero
    LimparCampoTexto txtComplemento
    LimparCampoTexto txtCep
    LimparCampoTexto txtBairro
    LimparCampoTexto txtCidade
    LimparCampoTexto txtEstado
    LimparCampoTexto txtPais
    LimparCampoTexto txtRuaCob
    LimparCampoTexto txtNumeroCob
    LimparCampoTexto txtComplementoCob
    LimparCampoTexto txtCepCob
    LimparCampoTexto txtBairroCob
    LimparCampoTexto txtCidadeCob
    LimparCampoTexto txtEstadoCob
    LimparCampoTexto txtPaisCob
    optCredito(0).Value = False
    optCredito(1).Value = False
    '
    'tabDetalhes.TabEnabled(1) = False
    'picTrava(0).Enabled = True
    'picTrava(1).Enabled = True
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objEmpresa = New busSisContas.clsEmpresa
    Set objRs = objEmpresa.ListarEmpresa(lngEMPRESAID)
    '
    If Not objRs.EOF Then
      If objRs.Fields("ENDCOBIGUALENDCORR").Value Then
        chkEndCobIgualEndCorr.Value = 1
      Else
        chkEndCobIgualEndCorr.Value = 0
      End If
      txtNome.Text = objRs.Fields("NOME").Value & ""
      If objRs.Fields("DESCRTIPOEMPRESA").Value & "" <> "" Then
        cboTipo = objRs.Fields("DESCRTIPOEMPRESA").Value & ""
      End If
      If objRs.Fields("CREDITO").Value & "" = "L" Then
        optCredito(0).Value = True
      ElseIf objRs.Fields("CREDITO").Value & "" = "B" Then
        optCredito(1).Value = True
      End If
      cboTipo_click
    
      txtNomeFantasia.Text = objRs.Fields("NOMEFANTASIA").Value & ""
      txtNroContrato.Text = objRs.Fields("NROCONTRATO").Value & ""
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DTINICIOCONTRATO").Value, TpMaskData
      INCLUIR_VALOR_NO_MASK mskData(1), objRs.Fields("DTFIMCONTRATO").Value, TpMaskData
      '
      INCLUIR_VALOR_NO_MASK mskPercentual, objRs.Fields("PERCENTUALAG").Value, TpMaskLongo
      txtTelefone.Text = objRs.Fields("TEL").Value & ""
      txtTelefone1.Text = objRs.Fields("TEL1").Value & ""
      txtTelefone2.Text = objRs.Fields("TEL2").Value & ""
      txtCNPJ.Text = objRs.Fields("CGC").Value & ""
      txtInscrEstadual.Text = objRs.Fields("INSCRESTADUAL").Value & ""
      txtInscrMunicipal.Text = objRs.Fields("INSCRMUNICIPAL").Value & ""
      txtContato.Text = objRs.Fields("CONTATO").Value & ""
      txtObservacao.Text = objRs.Fields("OBSERVACAO").Value & ""
      INCLUIR_VALOR_NO_MASK mskDescDiaria, objRs.Fields("PERCDESCDIARIA").Value, TpMaskMoeda
      '
      txtRua.Text = objRs.Fields("ENDRUA").Value & ""
      txtNumero.Text = objRs.Fields("ENDNUMERO").Value & ""
      txtComplemento.Text = objRs.Fields("ENDCOMPLEMENTO").Value & ""
      txtCep.Text = objRs.Fields("ENDCEP").Value & ""
      txtBairro.Text = objRs.Fields("ENDBAIRRO").Value & ""
      txtCidade.Text = objRs.Fields("ENDCIDADE").Value & ""
      txtEstado.Text = objRs.Fields("ENDESTADO").Value & ""
      txtPais.Text = objRs.Fields("ENDPAIS").Value & ""
      txtRuaCob.Text = objRs.Fields("COBRUA").Value & ""
      txtNumeroCob.Text = objRs.Fields("COBNUMERO").Value & ""
      txtComplementoCob.Text = objRs.Fields("COBCOMPLEMENTO").Value & ""
      txtCepCob.Text = objRs.Fields("COBCEP").Value & ""
      txtBairroCob.Text = objRs.Fields("COBBAIRRO").Value & ""
      txtCidadeCob.Text = objRs.Fields("COBCIDADE").Value & ""
      txtEstadoCob.Text = objRs.Fields("COBESTADO").Value & ""
      txtPaisCob.Text = objRs.Fields("COBPAIS").Value & ""
      '
      tabDetalhes.TabEnabled(1) = True
    End If
    'picTrava(0).Enabled = True
    'picTrava(1).Enabled = True
    'tabDetalhes.TabEnabled(0) = True
    'tabDetalhes.TabEnabled(1) = True
    Set objEmpresa = Nothing
  End If
  chkEndCobIgualEndCorr_Click
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source & ".[frmUserEmpresaInc.Form_Load]"
  AmpN
  Unload Me
End Sub



Private Sub Form_Unload(Cancel As Integer)
  If Not bFechar Then Cancel = True
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    picTrava(0).Enabled = True
    picTrava(1).Enabled = False
    '
    'cmdCancelar.Enabled = True
    'cmdExcluir.Enabled = False
    'cmdIncluir.Enabled = False
    'cmdAlterar.Enabled = False
    If Status = tpStatus_Consultar Then
      cmdOk.Enabled = False
    Else
      cmdOk.Enabled = True
    End If
    SetarFoco txtNome
  Case 1
    picTrava(0).Enabled = False
    picTrava(1).Enabled = True
    '
    SetarFoco txtRua
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source & ".[frmUserEmpresaInc.tabDetalhes_Click]"
  AmpN
End Sub

Private Sub txtNroContrato_GotFocus()
  Selecionar_Conteudo txtNroContrato
End Sub

Private Sub txtNroContrato_LostFocus()
  Pintar_Controle txtNroContrato, tpCorContr_Normal
End Sub

Private Sub txtNumero_GotFocus()
  Selecionar_Conteudo txtNumero
End Sub

Private Sub txtNumeroCob_GotFocus()
  Selecionar_Conteudo txtNumeroCob
End Sub

Private Sub txtObservacao_GotFocus()
  Selecionar_Conteudo txtObservacao
End Sub

Private Sub txtPais_GotFocus()
  Selecionar_Conteudo txtPais
End Sub

Private Sub txtPaisCob_GotFocus()
  Selecionar_Conteudo txtPaisCob
End Sub

Private Sub txtRua_GotFocus()
  Selecionar_Conteudo txtRua
End Sub

Private Sub txtRuaCob_GotFocus()
  Selecionar_Conteudo txtRuaCob
End Sub

Private Sub txtTelefone_GotFocus()
  Selecionar_Conteudo txtTelefone
End Sub
Private Sub txtTelefone1_GotFocus()
  Selecionar_Conteudo txtTelefone1
End Sub
Private Sub txtTelefone2_GotFocus()
  Selecionar_Conteudo txtTelefone2
End Sub

