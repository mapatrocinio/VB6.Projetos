VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserEmpresaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Empresa"
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
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4575
         Left            =   90
         ScaleHeight     =   4515
         ScaleWidth      =   1605
         TabIndex        =   55
         TabStop         =   0   'False
         Top             =   810
         Width           =   1665
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   25
            Top             =   2700
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   26
            Top             =   3540
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   1830
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   90
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5295
      Left            =   120
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userEmpresaInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Endere�o"
      TabPicture(1)   =   "userEmpresaInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Contrato"
      TabPicture(2)   =   "userEmpresaInc.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdContrato"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame2 
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
         Left            =   -74880
         TabIndex        =   33
         Top             =   360
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   4515
            Index           =   1
            Left            =   120
            ScaleHeight     =   4515
            ScaleWidth      =   7575
            TabIndex        =   34
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtEstado 
               Height          =   285
               Left            =   1320
               MaxLength       =   2
               TabIndex        =   18
               Text            =   "txtEstado"
               Top             =   750
               Width           =   435
            End
            Begin VB.TextBox txtComplemento 
               Height          =   285
               Left            =   5220
               MaxLength       =   30
               TabIndex        =   17
               Text            =   "txtComplemento"
               Top             =   420
               Width           =   2175
            End
            Begin VB.TextBox txtNumero 
               Height          =   285
               Left            =   1320
               MaxLength       =   15
               TabIndex        =   16
               Text            =   "txtNumero"
               Top             =   420
               Width           =   2175
            End
            Begin VB.TextBox txtRua 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   15
               Text            =   "txtRua"
               Top             =   90
               Width           =   6075
            End
            Begin VB.TextBox txtBairro 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   20
               Text            =   "txtBairro"
               Top             =   1080
               Width           =   6075
            End
            Begin VB.TextBox txtCidade 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   21
               Text            =   "txtCidade"
               Top             =   1410
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskCep 
               Height          =   255
               Left            =   5220
               TabIndex        =   19
               Top             =   750
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
               Caption         =   "Estado"
               Height          =   195
               Index           =   9
               Left            =   60
               TabIndex        =   54
               Top             =   750
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Complemento"
               Height          =   195
               Index           =   8
               Left            =   3960
               TabIndex        =   53
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "N�mero"
               Height          =   195
               Index           =   7
               Left            =   60
               TabIndex        =   52
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Rua"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   51
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Bairro"
               Height          =   195
               Index           =   2
               Left            =   60
               TabIndex        =   50
               Top             =   1125
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cidade"
               Height          =   195
               Index           =   16
               Left            =   60
               TabIndex        =   49
               Top             =   1455
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cep"
               Height          =   195
               Index           =   3
               Left            =   3960
               TabIndex        =   48
               Top             =   750
               Width           =   1215
            End
         End
      End
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
         TabIndex        =   29
         Top             =   360
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   4515
            Index           =   0
            Left            =   120
            ScaleHeight     =   4515
            ScaleWidth      =   7575
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtTelefoneContato 
               Height          =   285
               Left            =   1320
               MaxLength       =   10
               TabIndex        =   12
               Text            =   "txtTelefoneContato"
               Top             =   3060
               Width           =   2175
            End
            Begin VB.TextBox txtContato 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   11
               Text            =   "txtContato"
               Top             =   2730
               Width           =   6075
            End
            Begin VB.TextBox txtEmail 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   10
               Text            =   "txtEmail"
               Top             =   2400
               Width           =   6075
            End
            Begin VB.TextBox txtFax 
               Height          =   285
               Left            =   5220
               MaxLength       =   30
               TabIndex        =   9
               Text            =   "txtFax"
               Top             =   2070
               Width           =   2175
            End
            Begin VB.TextBox txtTelefone3 
               Height          =   285
               Left            =   1320
               MaxLength       =   30
               TabIndex        =   8
               Text            =   "txtTelefone3"
               Top             =   2070
               Width           =   2175
            End
            Begin VB.TextBox txtTelefone2 
               Height          =   285
               Left            =   5220
               MaxLength       =   30
               TabIndex        =   7
               Text            =   "txtTelefone2"
               Top             =   1740
               Width           =   2175
            End
            Begin VB.TextBox txtTelefone1 
               Height          =   285
               Left            =   1320
               MaxLength       =   30
               TabIndex        =   6
               Text            =   "txtTelefone1"
               Top             =   1740
               Width           =   2175
            End
            Begin VB.TextBox txtInscrMunicipal 
               Height          =   285
               Left            =   5220
               MaxLength       =   30
               TabIndex        =   5
               Text            =   "txtInscrMunicipal"
               Top             =   1410
               Width           =   2175
            End
            Begin VB.TextBox txtInscrEstadual 
               Height          =   285
               Left            =   1320
               MaxLength       =   30
               TabIndex        =   4
               Text            =   "txtInscrEstadual"
               Top             =   1410
               Width           =   2175
            End
            Begin VB.ComboBox cboTipoEmpresa 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   720
               Width           =   6105
            End
            Begin VB.TextBox txtNomeFantasia 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   1
               Text            =   "txtNomeFantasia"
               Top             =   390
               Width           =   6075
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   5190
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   35
               TabStop         =   0   'False
               Top             =   3090
               Width           =   2235
               Begin VB.OptionButton optStatus 
                  Caption         =   "Inativo"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   14
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optStatus 
                  Caption         =   "Ativo"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   13
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
            Begin MSMask.MaskEdBox mskCnpj 
               Height          =   255
               Left            =   1320
               TabIndex        =   3
               Top             =   1080
               Width           =   2205
               _ExtentX        =   3889
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   18
               Mask            =   "##.###.###/####-##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Tel. Contato"
               Height          =   195
               Index           =   33
               Left            =   60
               TabIndex        =   47
               Top             =   3060
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Contato"
               Height          =   195
               Index           =   32
               Left            =   60
               TabIndex        =   46
               Top             =   2760
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "E-mail"
               Height          =   195
               Index           =   31
               Left            =   60
               TabIndex        =   45
               Top             =   2400
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Fax"
               Height          =   195
               Index           =   30
               Left            =   3960
               TabIndex        =   44
               Top             =   2070
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Telefone 3"
               Height          =   195
               Index           =   29
               Left            =   60
               TabIndex        =   43
               Top             =   2070
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Telefone 2"
               Height          =   195
               Index           =   28
               Left            =   3960
               TabIndex        =   42
               Top             =   1740
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Telefone 1"
               Height          =   195
               Index           =   27
               Left            =   60
               TabIndex        =   41
               Top             =   1740
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Inscr. Municipal"
               Height          =   195
               Index           =   26
               Left            =   3960
               TabIndex        =   40
               Top             =   1410
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Inscr. Estadual"
               Height          =   195
               Index           =   25
               Left            =   60
               TabIndex        =   39
               Top             =   1410
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cnpj"
               Height          =   195
               Index           =   4
               Left            =   60
               TabIndex        =   38
               Top             =   1080
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Tipo Empresa"
               Height          =   195
               Index           =   24
               Left            =   60
               TabIndex        =   37
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Nome Fantasia"
               Height          =   195
               Index           =   21
               Left            =   60
               TabIndex        =   36
               Top             =   405
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Status"
               Height          =   315
               Index           =   5
               Left            =   3960
               TabIndex        =   32
               Top             =   3120
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Nome"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   31
               Top             =   120
               Width           =   1215
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdContrato 
         Height          =   4545
         Left            =   -74880
         OleObjectBlob   =   "userEmpresaInc.frx":0054
         TabIndex        =   56
         Top             =   390
         Width           =   7965
      End
   End
End
Attribute VB_Name = "frmUserEmpresaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Public strNumeroAptoPrinc       As String

Public lngPKID                  As Long

Private blnPrimeiraVez          As Boolean

Dim CONTRATO_COLUNASMATRIZ         As Long
Dim CONTRATO_LINHASMATRIZ          As Long

Private CONTRATO_Matriz()          As String

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Empresa
  LimparCampoTexto txtNome
  LimparCampoTexto txtNomeFantasia
  LimparCampoCombo cboTipoEmpresa
  LimparCampoMask mskCnpj
  LimparCampoTexto txtInscrEstadual
  LimparCampoTexto txtInscrMunicipal
  LimparCampoTexto txtTelefone1
  LimparCampoTexto txtTelefone2
  LimparCampoTexto txtTelefone3
  LimparCampoTexto txtFax
  LimparCampoTexto txtEmail
  LimparCampoTexto txtContato
  LimparCampoTexto txtTelefoneContato
  LimparCampoTexto txtRua
  LimparCampoTexto txtNumero
  LimparCampoTexto txtComplemento
  LimparCampoTexto txtEstado
  LimparCampoMask mskCep
  LimparCampoTexto txtBairro
  LimparCampoTexto txtCidade
  optStatus(0).Value = False
  optStatus(1).Value = False
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserEmpresaInc.LimparCampos]", _
            Err.Description
End Sub


Private Sub cboTipoEmpresa_LostFocus()
  Pintar_Controle cboTipoEmpresa, tpCorContr_Normal
End Sub


Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 2
    'Contrato
    If Not IsNumeric(grdContrato.Columns("PKID").Value & "") Then
      MsgBox "Selecione um contrato!", vbExclamation, TITULOSISTEMA
      SetarFoco grdContrato
      Exit Sub
    End If

    frmUserContratoInc.lngPKID = grdContrato.Columns("PKID").Value
    frmUserContratoInc.lngEMPRESAID = lngPKID
    frmUserContratoInc.strDescrEmpresa = txtNome.Text
    frmUserContratoInc.Status = tpStatus_Alterar
    frmUserContratoInc.Show vbModal

    If frmUserContratoInc.blnRetorno Then
      CONTRATO_MontaMatriz
      grdContrato.Bookmark = Null
      grdContrato.ReBind
      grdContrato.ApproxCount = CONTRATO_LINHASMATRIZ
    End If
    SetarFoco grdContrato
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
  Dim objContrato     As busSisLoc.clsContrato
  Dim objGeral        As busSisLoc.clsGeral
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  '
  Select Case tabDetalhes.Tab
  Case 2 'Exclus�o de Contrato
    '
    If Len(Trim(grdContrato.Columns("PKID").Value & "")) = 0 Then
      MsgBox "Selecione um contrato da empresa.", vbExclamation, TITULOSISTEMA
      SetarFoco grdContrato
      Exit Sub
    End If
    '
    '
    Set objGeral = New busSisLoc.clsGeral
    'OBRA
    strSql = "SELECT * FROM OBRA WHERE CONTRATOID = " & Formata_Dados(grdContrato.Columns("PKID").Value, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      TratarErroPrevisto "N�o � poss�vel excluir o contrato, pois existem obras associadas a ele.", "[cmdExcluir_Click]"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    'NF
    strSql = "SELECT * FROM NF WHERE CONTRATOID = " & Formata_Dados(grdContrato.Columns("PKID").Value, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      TratarErroPrevisto "N�o � poss�vel excluir o contrato, pois existem NF associadas a ele.", "[cmdExcluir_Click]"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    'DEVOLUCAO
    strSql = "SELECT * FROM DEVOLUCAO WHERE CONTRATOID = " & Formata_Dados(grdContrato.Columns("PKID").Value, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      TratarErroPrevisto "N�o � poss�vel excluir o contrato, pois existem devolu��es associadas a ele.", "[cmdExcluir_Click]"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    'DEVOLUCAO
    strSql = "SELECT * FROM BM WHERE CONTRATOID = " & Formata_Dados(grdContrato.Columns("PKID").Value, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      TratarErroPrevisto "N�o � poss�vel excluir o contrato, pois existem BMs associadas a ele.", "[cmdExcluir_Click]"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    'OK
    '
    If MsgBox("Confirma exclus�o do contrato " & grdContrato.Columns("N�mero").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
      SetarFoco grdContrato
      Exit Sub
    End If
    'OK
    Set objContrato = New busSisLoc.clsContrato
    objContrato.ExcluirContrato CLng(grdContrato.Columns("PKID").Value)
    '
    CONTRATO_MontaMatriz
    grdContrato.Bookmark = Null
    grdContrato.ReBind
    grdContrato.ApproxCount = CONTRATO_LINHASMATRIZ

    Set objContrato = Nothing
    SetarFoco grdContrato
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
  Case 2 'Contrato
    frmUserContratoInc.Status = tpStatus_Incluir
    frmUserContratoInc.lngPKID = 0
    frmUserContratoInc.lngEMPRESAID = lngPKID
    frmUserContratoInc.strDescrEmpresa = txtNome.Text
    frmUserContratoInc.Show vbModal

    If frmUserContratoInc.blnRetorno Then
      CONTRATO_MontaMatriz
      grdContrato.Bookmark = Null
      grdContrato.ReBind
      grdContrato.ApproxCount = CONTRATO_LINHASMATRIZ
    End If
    SetarFoco grdContrato
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdOK_Click()
  Dim objEmpresa               As busSisLoc.clsEmpresa
  Dim objGeral                  As busSisLoc.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngTIPOEMPRESAID         As Long
  Dim strStatus                 As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisLoc.clsGeral
  Set objEmpresa = New busSisLoc.clsEmpresa
  'TIPO EMPRESA
  lngTIPOEMPRESAID = 0
  strSql = "SELECT PKID FROM TIPOEMPRESA WHERE DESCRICAO = " & Formata_Dados(cboTipoEmpresa.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngTIPOEMPRESAID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Status
  If optStatus(0).Value Then
    strStatus = "A"
  Else
    strStatus = "I"
  End If

  'Validar se funcion�rio j� cadastrado
  strSql = "SELECT * FROM EMPRESA " & _
    " WHERE EMPRESA.NOME = " & Formata_Dados(txtNome.Text, tpDados_Texto) & _
    " AND EMPRESA.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtNome, tpCorContr_Erro
    TratarErroPrevisto "Empresa j� cadastrada"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objEmpresa = Nothing
    cmdOk.Enabled = True
    SetarFoco txtNome
    tabDetalhes.Tab = 1
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Empresa
    objEmpresa.AlterarEmpresa lngPKID, _
                                lngTIPOEMPRESAID, _
                                txtNome.Text, _
                                txtNomeFantasia.Text, _
                                mskCnpj.ClipText, _
                                txtInscrEstadual.Text, _
                                txtInscrMunicipal.Text, _
                                txtTelefone1.Text, _
                                txtTelefone2.Text, _
                                txtTelefone3.Text, _
                                txtFax.Text, _
                                txtEmail.Text, _
                                txtContato.Text, _
                                txtTelefoneContato.Text, _
                                txtRua.Text, _
                                txtNumero.Text, _
                                txtComplemento.Text, _
                                txtEstado.Text, _
                                IIf(mskCep.ClipText = "", "", mskCep.ClipText), _
                                txtBairro.Text, _
                                txtCidade.Text, strStatus
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Empresa
    objEmpresa.InserirEmpresa lngPKID, _
                                lngTIPOEMPRESAID, _
                                txtNome.Text, _
                                txtNomeFantasia.Text, _
                                mskCnpj.ClipText, _
                                txtInscrEstadual.Text, _
                                txtInscrMunicipal.Text, _
                                txtTelefone1.Text, _
                                txtTelefone2.Text, _
                                txtTelefone3.Text, _
                                txtFax.Text, _
                                txtEmail.Text, _
                                txtContato.Text, _
                                txtTelefoneContato.Text, _
                                txtRua.Text, _
                                txtNumero.Text, _
                                txtComplemento.Text, _
                                txtEstado.Text, _
                                IIf(mskCep.ClipText = "", "", mskCep.ClipText), _
                                txtBairro.Text, _
                                txtCidade.Text, strStatus
    blnRetorno = True
    'Selecionar plano cadastrado
    'entrar em modo de altera��o
    'lngPKID = objRs.Fields("PKID")
    Status = tpStatus_Alterar
    'Reload na tela
    Form_Load
    'Acerta tabs
    tabDetalhes.Tab = 2
    blnRetorno = True
  End If
  Set objEmpresa = Nothing
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
  If Not Valida_String(txtNomeFantasia, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o nome fantasia" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(cboTipoEmpresa, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o tipo da empresa" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(mskCnpj, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o cnpj" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Option(optStatus, blnSetarFocoControle) Then
    strMsg = strMsg & "Slecionar o status" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(Trim(mskCep.ClipText)) > 0 Then
    If Len(Trim(mskCep.ClipText)) <> 8 Then
      strMsg = strMsg & "Informar o CEP v�lido" & vbCrLf
      Pintar_Controle mskCep, tpCorContr_Erro
      SetarFoco mskCep
      tabDetalhes.Tab = 1
      blnSetarFocoControle = False
    End If
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserEmpresaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserEmpresaInc.ValidaCampos]", _
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
  TratarErro Err.Number, Err.Description, "[frmUserEmpresaInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objEmpresa             As busSisLoc.clsEmpresa
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
  'Tipo de Conv�nio
  strSql = "Select DESCRICAO from TIPOEMPRESA ORDER BY DESCRICAO"
  PreencheCombo cboTipoEmpresa, strSql, False, True
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
    tabDetalhes.TabEnabled(1) = True
    tabDetalhes.TabEnabled(2) = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objEmpresa = New busSisLoc.clsEmpresa
    Set objRs = objEmpresa.SelecionarEmpresaPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      txtNome.Text = objRs.Fields("NOME").Value & ""
      txtNomeFantasia.Text = objRs.Fields("NOMEFANTASIA").Value & ""
      txtNumero.Text = objRs.Fields("NOMEFANTASIA").Value & ""
      cboTipoEmpresa.Text = objRs.Fields("DESC_TIPOEMPRESA").Value & ""
      INCLUIR_VALOR_NO_MASK mskCnpj, objRs.Fields("CNPJ").Value, TpMaskSemMascara
      txtInscrEstadual.Text = objRs.Fields("INSCRICAOESTADUAL").Value & ""
      txtInscrMunicipal.Text = objRs.Fields("INSCRICAOMUNICIPAL").Value & ""
      txtTelefone1.Text = objRs.Fields("TELEFONE1").Value & ""
      txtTelefone2.Text = objRs.Fields("TELEFONE2").Value & ""
      txtTelefone3.Text = objRs.Fields("TELEFONE3").Value & ""
      txtFax.Text = objRs.Fields("FAX").Value & ""
      txtEmail.Text = objRs.Fields("EMAIL").Value & ""
      txtContato.Text = objRs.Fields("CONTATO").Value & ""
      txtTelefoneContato.Text = objRs.Fields("TELEFONECONTATO").Value & ""
      txtRua.Text = objRs.Fields("ENDRUA").Value & ""
      txtNumero.Text = objRs.Fields("ENDNUMERO").Value & ""
      txtComplemento.Text = objRs.Fields("ENDCOMPL").Value & ""
      txtEstado.Text = objRs.Fields("ENDESTADO").Value & ""
      INCLUIR_VALOR_NO_MASK mskCep, objRs.Fields("ENDCEP").Value, TpMaskSemMascara
      txtBairro.Text = objRs.Fields("ENDBAIRRO").Value & ""
      txtCidade.Text = objRs.Fields("ENDCIDADE").Value & ""
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
    Set objEmpresa = Nothing
    'Visible
    optStatus(0).Visible = True
    optStatus(1).Visible = True
    Label5(5).Visible = True
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = True
    tabDetalhes.TabEnabled(2) = True
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

Private Sub mskCep_GotFocus()
  Seleciona_Conteudo_Controle mskCep
End Sub
Private Sub mskCep_LostFocus()
  Pintar_Controle mskCep, tpCorContr_Normal
End Sub

Private Sub mskCnpj_GotFocus()
  Seleciona_Conteudo_Controle mskCnpj
End Sub
Private Sub mskCnpj_LostFocus()
  Pintar_Controle mskCnpj, tpCorContr_Normal
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

Private Sub txtComplemento_GotFocus()
  Seleciona_Conteudo_Controle txtComplemento
End Sub
Private Sub txtComplemento_LostFocus()
  Pintar_Controle txtComplemento, tpCorContr_Normal
End Sub

Private Sub txtContato_GotFocus()
  Seleciona_Conteudo_Controle txtContato
End Sub
Private Sub txtContato_LostFocus()
  Pintar_Controle txtContato, tpCorContr_Normal
End Sub
Private Sub txtEmail_GotFocus()
  Seleciona_Conteudo_Controle txtEmail
End Sub
Private Sub txtEmail_LostFocus()
  Pintar_Controle txtEmail, tpCorContr_Normal
End Sub

Private Sub txtEstado_GotFocus()
  Seleciona_Conteudo_Controle txtEstado
End Sub
Private Sub txtEstado_LostFocus()
  Pintar_Controle txtEstado, tpCorContr_Normal
End Sub

Private Sub txtFax_GotFocus()
  Seleciona_Conteudo_Controle txtFax
End Sub
Private Sub txtFax_LostFocus()
  Pintar_Controle txtFax, tpCorContr_Normal
End Sub
Private Sub txtInscrEstadual_GotFocus()
  Seleciona_Conteudo_Controle txtInscrEstadual
End Sub
Private Sub txtInscrEstadual_LostFocus()
  Pintar_Controle txtInscrEstadual, tpCorContr_Normal
End Sub

Private Sub txtInscrMunicipal_GotFocus()
  Seleciona_Conteudo_Controle txtInscrMunicipal
End Sub
Private Sub txtInscrMunicipal_LostFocus()
  Pintar_Controle txtInscrMunicipal, tpCorContr_Normal
End Sub

Private Sub txtNome_GotFocus()
  Seleciona_Conteudo_Controle txtNome
End Sub
Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
End Sub

Private Sub txtNomeFantasia_GotFocus()
  Seleciona_Conteudo_Controle txtNomeFantasia
End Sub
Private Sub txtNomeFantasia_LostFocus()
  Pintar_Controle txtNomeFantasia, tpCorContr_Normal
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

Private Sub txtTelefoneContato_GotFocus()
  Seleciona_Conteudo_Controle txtTelefoneContato
End Sub
Private Sub txtTelefoneContato_LostFocus()
  Pintar_Controle txtTelefoneContato, tpCorContr_Normal
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    grdContrato.Enabled = False
    pictrava(0).Enabled = True
    pictrava(1).Enabled = False
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco txtNome
  Case 1
    grdContrato.Enabled = False
    pictrava(0).Enabled = False
    pictrava(1).Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco txtRua
  Case 2
    'Contrato
    grdContrato.Enabled = True
    pictrava(0).Enabled = False
    pictrava(1).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = True
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    'Montar RecordSet
    CONTRATO_COLUNASMATRIZ = grdContrato.Columns.Count
    CONTRATO_LINHASMATRIZ = 0
    CONTRATO_MontaMatriz
    grdContrato.Bookmark = Null
    grdContrato.ReBind
    grdContrato.ApproxCount = CONTRATO_LINHASMATRIZ
    '
    SetarFoco grdContrato
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "SisLoc.frmUserEmpresaInc.tabDetalhes"
  AmpN
End Sub


Private Sub grdContrato_UnboundReadDataEx( _
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
               Offset + intI, CONTRATO_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, CONTRATO_COLUNASMATRIZ, CONTRATO_LINHASMATRIZ, CONTRATO_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, CONTRATO_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserEmpresaInc.grdGeral_UnboundReadDataEx]"
End Sub

Public Sub CONTRATO_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisLoc.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisLoc.clsGeral
  '
  strSql = "SELECT CONTRATO.PKID, CONTRATO.NUMERO, CONTRATO.DATAINICIO, CONTRATO.DATAFIM, PESSOA.NOME " & _
          "FROM CONTRATO " & _
          " LEFT JOIN PESSOA ON PESSOA.PKID = CONTRATO.PESSOAID " & _
          "WHERE CONTRATO.EMPRESAID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
          " ORDER BY CONTRATO.ANO DESC, CONTRATO.SEQUENCIAL DESC"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    CONTRATO_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim CONTRATO_Matriz(0 To CONTRATO_COLUNASMATRIZ - 1, 0 To CONTRATO_LINHASMATRIZ - 1)
  Else
    ReDim CONTRATO_Matriz(0 To CONTRATO_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se j� houver algum item
    For intI = 0 To CONTRATO_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda n�o se atingiu fim do recordset
        For intJ = 0 To CONTRATO_COLUNASMATRIZ - 1  'varre as colunas
          CONTRATO_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'pr�xima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
