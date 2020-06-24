VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserZerarContaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Zerar Conta"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5145
      Left            =   8250
      ScaleHeight     =   5145
      ScaleWidth      =   1860
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   1875
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   1605
         TabIndex        =   6
         Top             =   3000
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Default         =   -1  'True
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4815
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da Conta"
      TabPicture(0)   =   "userZerarContaInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0C0&
         Caption         =   "Informações cadastrais"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3855
         Left            =   120
         TabIndex        =   8
         Top             =   480
         Width           =   7335
         Begin VB.ComboBox cboConta 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   1440
            Width           =   3855
         End
         Begin MSMask.MaskEdBox mskValor 
            Height          =   255
            Left            =   1560
            TabIndex        =   1
            Top             =   1080
            Width           =   1695
            _ExtentX        =   2990
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskData 
            Height          =   255
            Index           =   0
            Left            =   1560
            TabIndex        =   0
            Top             =   720
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label lblDe 
            Caption         =   "Conta"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblData 
            Caption         =   "Data"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label lblValor 
            Caption         =   "Valor"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   1080
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmUserZerarContaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                         As tpStatus
Public lngMOVIMENTACAOID                   As Long
Public bRetorno                       As Boolean
Public blnPrimeiraVez                 As Boolean
Public bFechar                        As Boolean
Public strStatus                      As String


Private Sub cboConta_LostFocus()
  Pintar_Controle cboConta, tpCorContr_Normal
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
  Dim objConta                As busSisContas.clsConta
  Dim objGeral                As busSisContas.clsGeral
  Dim lngCONTAID              As Long
  Dim curSaldo                As Currency
  Dim strDocumento            As String
  Dim objMovimentacao         As busSisContas.clsMovimentacao
  '
  Set objMovimentacao = New busSisContas.clsMovimentacao
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Despesa
    If Not ValidaCampos Then Exit Sub
    'Valida se unidade de estoque já cadastrada
    Set objGeral = New busSisContas.clsGeral
    strSql = "Select PKID From CONTA WHERE DESCRICAO = " & Formata_Dados(cboConta.Text, tpDados_Texto, tpNulo_NaoAceita)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      TratarErroPrevisto "Conta não cadastrada", "cmdOK_Click"
      Exit Sub
      
    Else
      lngCONTAID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    'Set objConta = New busSisContas.clsConta
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      Set objConta = New busSisContas.clsConta
      curSaldo = objConta.RetornarSaldo(Format(Now, "DD/MM/YYYY"), _
                                        Format(Now, "DD/MM/YYYY"), _
                                        lngCONTAID)
      Set objConta = Nothing
      curSaldo = curSaldo * -1
      'Código para inclusão
      '
      strDocumento = RetornaGravaSequencial("SEQUENCIALMOV")
      objMovimentacao.IncluirMovimentacao IIf(curSaldo <= 0, "D", "C"), _
                                          Format(Now, "DD/MM/YYYY"), _
                                          strDocumento, _
                                          IIf(curSaldo <= 0, lngCONTAID & "", ""), _
                                          IIf(curSaldo <= 0, "", lngCONTAID & ""), _
                                          Format(IIf(curSaldo < 0, curSaldo * -1, curSaldo), "###,##0.00"), _
                                          "Zerar Conta"
      bRetorno = True
      bFechar = True
      MsgBox "Conta zerada!", vbExclamation, TITULOSISTEMA
      'Unload Me
    End If
    'Set objMovimentacao = Nothing
    'bFechar = True
    'Unload Me
  End Select
  Set objMovimentacao = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg              As String
  Dim objSubGrupoDespesa  As busSisContas.clsSubGrupoDespesa
  Dim strTipoDespesa      As String
  '
  If Not Valida_Data(mskData(0), TpObrigatorio) Then
    strMsg = strMsg & "Informar a data da movimentacao válida" & vbCrLf
    Pintar_Controle mskData(0), tpCorContr_Erro
  End If
  If Not Valida_Moeda(mskValor, TpObrigatorio) Then
    strMsg = strMsg & "Informar o valor da movimentação válido" & vbCrLf
    Pintar_Controle mskValor, tpCorContr_Erro
  End If
  If Len(cboConta.Text) = 0 Then
    strMsg = strMsg & "Selecionar uma conta" & vbCrLf
    Pintar_Controle cboConta, tpCorContr_Erro
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserZerarContaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    SetarFoco cboConta
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserZerarContaInc.Form_Activate]"
End Sub



Private Sub Form_Load()
On Error GoTo trata
  '
  Dim strSql As String
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 5520
  Me.Width = 10200
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'lblData.Enabled = False
  lblValor.Enabled = False
  'mskData(0).Enabled = False
  mskValor.Enabled = False
  '
  mskData(0).Text = Format(Now(), "DD/MM/YYYY")
  mskValor.Text = 0
  strSql = "SELECT DESCRICAO FROM CONTA ORDER BY DESCRICAO;"
  PreencheCombo cboConta, strSql, False, True
  '
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not bFechar Then Cancel = True
End Sub

Private Sub mskData_GotFocus(Index As Integer)
  Selecionar_Conteudo mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub

Private Sub mskValor_GotFocus()
  Selecionar_Conteudo mskValor
End Sub

Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub
