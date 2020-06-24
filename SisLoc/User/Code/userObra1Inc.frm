VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserObra1Inc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de contrato de empresa"
   ClientHeight    =   4125
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4125
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4125
      Left            =   8430
      ScaleHeight     =   4125
      ScaleWidth      =   1860
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   90
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   1830
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   150
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   3795
      Left            =   120
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   6694
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userObra1Inc.frx":0000
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
         Height          =   3165
         Left            =   120
         TabIndex        =   11
         Top             =   420
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2925
            Index           =   0
            Left            =   120
            ScaleHeight     =   2925
            ScaleWidth      =   7575
            TabIndex        =   12
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtAno 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   2460
               MaxLength       =   100
               TabIndex        =   3
               Text            =   "txtAno"
               Top             =   780
               Width           =   615
            End
            Begin VB.TextBox txtSequencial 
               BackColor       =   &H00FFFFFF&
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   2
               Text            =   "txtSequencial"
               Top             =   780
               Width           =   1125
            End
            Begin VB.ComboBox cboFuncionario 
               BackColor       =   &H00E0E0E0&
               Height          =   315
               Left            =   1320
               Locked          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   420
               Width           =   6105
            End
            Begin VB.TextBox txtEmpresa 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtEmpresa"
               Top             =   90
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskInicio 
               Height          =   255
               Left            =   1320
               TabIndex        =   4
               Top             =   1110
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
            Begin MSMask.MaskEdBox mskFim 
               Height          =   255
               Left            =   5820
               TabIndex        =   5
               Top             =   1110
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
               Caption         =   "Número"
               Height          =   195
               Index           =   5
               Left            =   60
               TabIndex        =   17
               Top             =   765
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Funcionário"
               Height          =   195
               Index           =   24
               Left            =   60
               TabIndex        =   16
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Fim"
               Height          =   195
               Index           =   7
               Left            =   4560
               TabIndex        =   15
               Top             =   1110
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Início"
               Height          =   195
               Index           =   6
               Left            =   60
               TabIndex        =   14
               Top             =   1110
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Empresa"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   13
               Top             =   105
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserObra1Inc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngEMPRESAID            As Long
Public strDescrEmpresa         As String

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Contrato
  LimparCampoTexto txtEmpresa
  LimparCampoCombo cboFuncionario
  LimparCampoTexto txtSequencial
  LimparCampoTexto txtAno
  LimparCampoMask mskInicio
  LimparCampoMask mskFim
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserContratoInc.LimparCampos]", _
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

Private Sub cmdOk_Click()
  Dim objContrato               As busSisLoc.clsContrato
  Dim objGeral                  As busSisLoc.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngFUNCIONARIOID          As Long
  Dim strNumero                 As String
  Dim strSequencial             As String
  Dim lngAno                    As Long
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisLoc.clsGeral
  Set objContrato = New busSisLoc.clsContrato
  'FUNCIONARIO
  lngFUNCIONARIOID = 0
  strSql = "SELECT PKID FROM PESSOA WHERE NOME = " & Formata_Dados(cboFuncionario.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngFUNCIONARIOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing

  'Validar se contrato já cadastrado
  strSql = "SELECT * FROM CONTRATO " & _
    " WHERE CONTRATO.SEQUENCIAL = " & Formata_Dados(txtSequencial.Text, tpDados_Longo) & _
    " AND CONTRATO.ANO = " & Formata_Dados(txtAno.Text, tpDados_Longo) & _
    " AND CONTRATO.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtSequencial, tpCorContr_Erro
    TratarErroPrevisto "Contrato já cadastrado"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objContrato = Nothing
    cmdOk.Enabled = True
    SetarFoco txtSequencial
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Contrato
    objContrato.AlterarContrato lngPKID, _
                                mskInicio.Text, _
                                mskFim.Text, _
                                lngFUNCIONARIOID & ""
    '
  ElseIf Status = tpStatus_Incluir Then
    'Obter dados do contrato
    'lngAno = Right(mskInicio.Text, 4)
    'strSequencial = RetornaGravaCampoSequencialCtrto("SEQUENCIAL", lngAno) & ""
    lngAno = txtAno.Text
    strSequencial = txtSequencial.Text
    strNumero = "RF" & Format(strSequencial, "0000") & "/" & lngAno

    'Inserir Contrato
    objContrato.InserirContrato strNumero, _
                                strSequencial, _
                                lngAno & "", _
                                mskInicio.Text, _
                                mskFim.Text, _
                                lngEMPRESAID & "", _
                                lngFUNCIONARIOID & ""
  End If
  Set objContrato = Nothing
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
  If Not Valida_Moeda(txtSequencial, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o sequencial válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(txtAno, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o ano válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  
  If strMsg = "" Then
    If Len(txtAno.Text) <> 4 Then
      SetarFoco txtAno
      blnSetarFocoControle = False
      strMsg = strMsg & "Preencher o ano válido" & vbCrLf
      tabDetalhes.Tab = 0
    End If
  End If
  If Not Valida_Data(mskInicio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de início" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskFim, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de fim válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserContratoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserContratoInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    If Status = tpStatus_Alterar Then
      SetarFoco mskInicio
    Else
      SetarFoco txtSequencial
    End If
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserContratoInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objContrato           As busSisLoc.clsContrato
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 4605
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  'Fucnionário
  strSql = "Select NOME from PESSOA ORDER BY NOME"
  PreencheCombo cboFuncionario, strSql, False, True
  
  txtEmpresa.Text = strDescrEmpresa
  INCLUIR_VALOR_NO_COMBO gsNomeUsuCompleto & "", cboFuncionario
  '
  txtAno.Enabled = True
  txtSequencial.Enabled = True
  '
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    txtAno.Enabled = False
    txtSequencial.Enabled = False
    Set objContrato = New busSisLoc.clsContrato
    Set objRs = objContrato.SelecionarContratoPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      
      INCLUIR_VALOR_NO_COMBO objRs.Fields("FUNCIONARIO").Value & "", cboFuncionario
      txtSequencial.Text = objRs.Fields("SEQUENCIAL").Value & ""
      txtAno.Text = objRs.Fields("ANO").Value & ""
      INCLUIR_VALOR_NO_MASK mskInicio, objRs.Fields("DATAINICIO").Value & "", TpMaskData
      INCLUIR_VALOR_NO_MASK mskFim, objRs.Fields("DATAFIM").Value & "", TpMaskData
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objContrato = Nothing
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

Private Sub mskFim_GotFocus()
  Seleciona_Conteudo_Controle mskFim
End Sub
Private Sub mskFim_LostFocus()
  Pintar_Controle mskFim, tpCorContr_Normal
End Sub

Private Sub mskInicio_GotFocus()
  Seleciona_Conteudo_Controle mskInicio
End Sub
Private Sub mskInicio_LostFocus()
  Pintar_Controle mskInicio, tpCorContr_Normal
End Sub

Private Sub txtAno_GotFocus()
  Seleciona_Conteudo_Controle txtAno
End Sub
Private Sub txtAno_LostFocus()
  Pintar_Controle txtAno, tpCorContr_Normal
End Sub

Private Sub txtSequencial_GotFocus()
  Seleciona_Conteudo_Controle txtSequencial
End Sub
Private Sub txtSequencial_LostFocus()
  Pintar_Controle txtSequencial, tpCorContr_Normal
End Sub
