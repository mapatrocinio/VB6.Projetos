VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserBMInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de BM"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4425
      Left            =   8520
      ScaleHeight     =   4425
      ScaleWidth      =   1860
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2055
         Left            =   30
         ScaleHeight     =   1995
         ScaleWidth      =   1605
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   2220
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4155
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   7329
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do BM"
      TabPicture(0)   =   "userBMInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picTrava(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox picTrava 
         BorderStyle     =   0  'None
         Height          =   3585
         Index           =   0
         Left            =   120
         ScaleHeight     =   3585
         ScaleWidth      =   7695
         TabIndex        =   11
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
            Height          =   3465
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   7695
            Begin VB.ComboBox cboMedicao 
               BackColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   5100
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   540
               Width           =   2475
            End
            Begin VB.TextBox txtNumero 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1440
               MaxLength       =   100
               TabIndex        =   1
               TabStop         =   0   'False
               Text            =   "txtNumero"
               Top             =   540
               Width           =   2295
            End
            Begin VB.TextBox txtEmpresa 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1440
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtEmpresa"
               Top             =   210
               Width           =   6135
            End
            Begin VB.ComboBox cboContrato 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   900
               Width           =   6135
            End
            Begin MSMask.MaskEdBox mskInicio 
               Height          =   255
               Left            =   1440
               TabIndex        =   4
               Top             =   1260
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskTermino 
               Height          =   255
               Left            =   6300
               TabIndex        =   5
               Top             =   1260
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Início"
               Height          =   195
               Index           =   6
               Left            =   180
               TabIndex        =   18
               Top             =   1260
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Término"
               Height          =   195
               Index           =   7
               Left            =   5070
               TabIndex        =   17
               Top             =   1260
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Medição"
               Height          =   195
               Index           =   2
               Left            =   3840
               TabIndex        =   16
               Top             =   555
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Número"
               Height          =   195
               Index           =   0
               Left            =   180
               TabIndex        =   15
               Top             =   555
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Empresa"
               Height          =   195
               Index           =   1
               Left            =   180
               TabIndex        =   14
               Top             =   225
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Contrato"
               Height          =   195
               Index           =   24
               Left            =   180
               TabIndex        =   13
               Top             =   930
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserBMInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngBMID                    As Long
Public bRetorno                   As Boolean
Public bFechar                    As Boolean
Private blnPrimeiraVez            As Boolean


Private Sub cboContrato_LostFocus()
  Pintar_Controle cboContrato, tpCorContr_Normal
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
  Dim objBM                   As busSisLoc.clsBM
  Dim objGeral                As busSisLoc.clsGeral
  Dim lngCONTRATOID           As Long
  Dim strNumero               As String
  Dim strMedicao              As String
  Dim strDataEmissa           As String
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração de Grupo cardápio
    If Not ValidaCampos Then Exit Sub
    'Valida se Grupo cardápio já cadastrado
    '
    Set objGeral = New busSisLoc.clsGeral
    'TIPO EMPRESA
    lngCONTRATOID = 0
    strSql = "SELECT PKID FROM CONTRATO WHERE NUMERO = " & Formata_Dados(cboContrato.Text, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngCONTRATOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    
    Set objBM = New busSisLoc.clsBM
    '
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      If cboMedicao.Text = "INICIAL" Then
        strMedicao = "I"
      ElseIf cboMedicao.Text = "INTERMEDIÁRIO" Then
        strMedicao = "M"
      ElseIf cboMedicao.Text = "FINAL" Then
        strMedicao = "F"
      Else
        strMedicao = ""
      End If
      objBM.AlterarBM lngBMID, _
                      lngCONTRATOID, _
                      strMedicao, _
                      mskInicio.Text, _
                      mskTermino.Text

      bRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      '
      strDataEmissa = Format(Now, "DD/MM/YYYY")
      strNumero = RetornaGravaCampoSequencialBM("SEQUENCIAL", cboContrato.Text) & ""
      strMedicao = RetornaMedicaoBM(lngCONTRATOID)
      '
      objBM.InserirBM lngCONTRATOID, _
                      strNumero, _
                      strMedicao, _
                      strDataEmissa, _
                      mskInicio.Text, _
                      mskTermino.Text

      bRetorno = True
    End If
    Set objBM = Nothing
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
  ValidaCampos = False
  If Not Valida_String(cboContrato, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Slecionar o contrato" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskInicio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de início" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskTermino, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de término" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserBMInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserBMInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    tabDetalhes.Tab = 0
    SetarFoco cboContrato
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserBMInc.Form_Activate]"
End Sub
Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Empresa
  LimparCampoTexto txtEmpresa
  LimparCampoTexto txtNumero
  LimparCampoCombo cboMedicao
  LimparCampoCombo cboContrato
  LimparCampoMask mskInicio
  LimparCampoMask mskTermino
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserEmpresaInc.LimparCampos]", _
            Err.Description
End Sub

Private Sub Form_Load()
On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim objBM         As busSisLoc.clsBM
  Dim strMedicao    As String
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 4905
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  'Limpar campos
  LimparCampos
  '
  'Contrato
  strSql = "SELECT NUMERO FROM CONTRATO ORDER BY NUMERO"
  PreencheCombo cboContrato, strSql, False, True
  'Medição
  cboMedicao.AddItem ""
  cboMedicao.AddItem "INICIAL"
  cboMedicao.AddItem "INTERMEDIÁRIO"
  cboMedicao.AddItem "FINAL"
  '
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objBM = New busSisLoc.clsBM
    Set objRs = objBM.SelecionarBM(lngBMID)
    '
    If Not objRs.EOF Then
      Select Case objRs.Fields("MEDICAO").Value & ""
      Case "I": strMedicao = "INICIAL"
      Case "M": strMedicao = "INTERMEDIÁRIO"
      Case "F": strMedicao = "FINAL"
      Case Else: strMedicao = ""
      End Select
      '
      txtEmpresa.Text = objRs.Fields("NOME_EMPRESA").Value & ""
      txtNumero.Text = Format(objRs.Fields("NUMERO").Value & "", "000")
      INCLUIR_VALOR_NO_COMBO strMedicao, cboMedicao
      INCLUIR_VALOR_NO_COMBO objRs.Fields("NUMERO_CONTRATO").Value & "", cboContrato
      INCLUIR_VALOR_NO_MASK mskInicio, objRs.Fields("DATAINICIO").Value & "", TpMaskData
      INCLUIR_VALOR_NO_MASK mskTermino, objRs.Fields("DATATERMINO").Value & "", TpMaskData
      '
    End If
    Set objBM = Nothing
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

Private Sub mskInicio_GotFocus()
  Selecionar_Conteudo mskInicio
End Sub

Private Sub mskInicio_LostFocus()
  Pintar_Controle mskInicio, tpCorContr_Normal
End Sub

Private Sub mskTermino_GotFocus()
  Selecionar_Conteudo mskTermino
End Sub

Private Sub mskTermino_LostFocus()
  Pintar_Controle mskTermino, tpCorContr_Normal
End Sub

