VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserEntradaAtendInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de entrada para atendente"
   ClientHeight    =   2880
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2880
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   2880
      Left            =   8430
      ScaleHeight     =   2880
      ScaleWidth      =   1860
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2025
         Left            =   90
         ScaleHeight     =   1965
         ScaleWidth      =   1605
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   690
         Width           =   1665
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   960
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   2595
      Left            =   120
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   4577
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userEntradaAtendInc.frx":0000
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
         Height          =   2175
         Left            =   120
         TabIndex        =   8
         Top             =   330
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   1935
            Index           =   0
            Left            =   120
            ScaleHeight     =   1935
            ScaleWidth      =   7575
            TabIndex        =   9
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.ComboBox cboAtendente 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   810
               Width           =   6165
            End
            Begin VB.TextBox txtData 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1320
               TabIndex        =   1
               Text            =   "txtData"
               Top             =   480
               Width           =   1815
            End
            Begin VB.TextBox txtTurno 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1320
               TabIndex        =   0
               Text            =   "txtTurno"
               Top             =   150
               Width           =   6165
            End
            Begin MSMask.MaskEdBox mskValor 
               Height          =   285
               Left            =   1320
               TabIndex        =   3
               Top             =   1170
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label6 
               Caption         =   "Atendente"
               Height          =   255
               Index           =   1
               Left            =   60
               TabIndex        =   14
               Top             =   810
               Width           =   1455
            End
            Begin VB.Label Label5 
               Caption         =   "Data"
               Height          =   195
               Index           =   2
               Left            =   60
               TabIndex        =   13
               Top             =   480
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Turno"
               Height          =   285
               Index           =   0
               Left            =   60
               TabIndex        =   12
               Top             =   150
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Valor"
               Height          =   225
               Index           =   1
               Left            =   60
               TabIndex        =   11
               Top             =   1140
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserEntradaAtendInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Public lngTURNOATENDEPESQ       As Long

Public lngPKID                  As Long
Public curSaldo                 As Currency
Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'EntradaAtend
  LimparCampoTexto txtTurno
  LimparCampoTexto txtData
  LimparCampoMask mskValor
  LimparCampoCombo cboAtendente
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserEntradaAtendInc.LimparCampos]", _
            Err.Description
End Sub


Private Sub cboAtendente_LostFocus()
  Pintar_Controle cboAtendente, tpCorContr_Normal
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
  Dim objEntradaAtend           As busSisMaq.clsEntradaAtend
  Dim objCaixaAtend             As busSisMaq.clsCaixaAtend
  Dim objGeral                  As busSisMaq.clsGeral
  Dim objRS                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngTURNOID                As Long
  Dim lngATENDENTEID            As Long
  Dim lngCAIXAATENDID           As Long
  Dim strData                   As String
  Dim strMsg                    As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objEntradaAtend = New busSisMaq.clsEntradaAtend
  'DONO
  lngTURNOID = RetornaCodTurnoCorrente
  Set objGeral = New busSisMaq.clsGeral
  'ATENDENTE
  lngATENDENTEID = 0
  strSql = "SELECT PESSOA.PKID FROM PESSOA WHERE PESSOA.NOME = " & Formata_Dados(cboAtendente.Text, tpDados_Texto)
  Set objRS = objGeral.ExecutarSQL(strSql)
  If Not objRS.EOF Then
    lngATENDENTEID = objRS.Fields("PKID").Value
  End If
  objRS.Close
  Set objRS = Nothing
  '
  'Verifica se inclirá turno do atendente
  lngCAIXAATENDID = RetornaCodTurnoCorrenteAtend(lngATENDENTEID, lngTURNOATENDEPESQ)
  If lngCAIXAATENDID = -1 Then
    'há mais de um caiatend aberto
    strMsg = "Há mais de um turno aberto para este atendente. Contacte o administrador do sistema."
    TratarErroPrevisto strMsg, "cmdConfirmar_Click"
    cmdOk.Enabled = True
    SetarFoco cboAtendente
    Exit Sub
  End If
  '
  
  Set objGeral = Nothing
  'Pede liberação do atendente
  frmUserLoginLibera.lngFUNCIONARIOID = lngATENDENTEID
  frmUserLoginLibera.Show vbModal
  If Len(Trim(gsNomeUsuLib)) = 0 Then
    strMsg = "É necessário confirmação do atendente para executar esta ação."
    TratarErroPrevisto strMsg, "cmdConfirmar_Click"
    cmdOk.Enabled = True
    SetarFoco cboAtendente
    Exit Sub
  End If
  
  
  If Status = tpStatus_Alterar Then
    'Alterar EntradaAtend
    objEntradaAtend.AlterarEntradaAtend lngPKID, _
                                        mskValor.ClipText
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir EntradaAtend
    strData = Format(Now, "DD/MM/YYYY hh:mm")
    'Verifica se existe caixa para o atendente
    If lngCAIXAATENDID = 0 Then
      'Não existe caixa atend cadastrado, cadastra
      Set objCaixaAtend = New busSisMaq.clsCaixaAtend
      objCaixaAtend.InserirCaixaAtend lngCAIXAATENDID, _
                                      lngATENDENTEID, _
                                      RetornaCodTurnoCorrente
      Set objCaixaAtend = Nothing
    End If
    objEntradaAtend.InserirEntradaAtend lngTURNOID, _
                                        lngCAIXAATENDID, _
                                        mskValor.ClipText, _
                                        strData
    
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  End If
  Set objEntradaAtend = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  cmdOk.Enabled = True
End Sub


Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  Dim curValor              As Currency
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  If Not Valida_String(cboAtendente, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o Atendente" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskValor, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o valor válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) = 0 Then
    'valida saldo
    curValor = CCur(mskValor.Text)
    If curValor > curSaldo Then
      strMsg = strMsg & "Saldo insuficiente no caixa. máximo permitido é " & Format(curSaldo, "###,##0.00") & vbCrLf
    End If
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserEntradaAtendInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserEntradaAtendInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco cboAtendente
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserEntradaAtendInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRS                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objEntradaAtend         As busSisMaq.clsEntradaAtend
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 3360
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  'Atendente
  strSql = "Select PESSOA.NOME " & _
        " FROM PESSOA " & _
        " INNER JOIN ATENDENTE ON PESSOA.PKID = ATENDENTE.PESSOAID " & _
        " ORDER BY PESSOA.NOME"
  PreencheCombo cboAtendente, strSql, False, True
  '
  If Status = tpStatus_Incluir Then
    '
    cboAtendente.Enabled = True
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    cboAtendente.Enabled = False
    '
    Set objEntradaAtend = New busSisMaq.clsEntradaAtend
    Set objRS = objEntradaAtend.SelecionarEntradaAtendPeloPkid(lngPKID)
    '
    If Not objRS.EOF Then
      txtTurno.Text = RetornaDescTurnoCorrente(objRS.Fields("TURNOENTRADAID").Value)
      txtData.Text = Format(objRS.Fields("DATA").Value, "DD/MM/YYYY hh:mm")
      INCLUIR_VALOR_NO_MASK mskValor, objRS.Fields("VALOR").Value, TpMaskMoeda
      cboAtendente.Text = objRS.Fields("DESC_ATENDENTE").Value & ""
      
    End If
    objRS.Close
    Set objRS = Nothing
    '
    Set objEntradaAtend = Nothing
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
