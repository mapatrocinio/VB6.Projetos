VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserRelBalancoCad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de saldo"
   ClientHeight    =   4185
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10110
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4185
   ScaleWidth      =   10110
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4185
      Left            =   8250
      ScaleHeight     =   4185
      ScaleWidth      =   1860
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   120
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   5
         Top             =   1890
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   3
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   3855
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do Saldo"
      TabPicture(0)   =   "userRelBalancoCad.frx":0000
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
         Height          =   3135
         Left            =   120
         TabIndex        =   7
         Top             =   360
         Width           =   7335
         Begin VB.Frame Frame5 
            Caption         =   "Sexo"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   840
            TabIndex        =   8
            Top             =   3480
            Width           =   2295
         End
         Begin MSMask.MaskEdBox mskData 
            Height          =   255
            Left            =   1470
            TabIndex        =   0
            Top             =   360
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   450
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox mskSaldo 
            Height          =   255
            Left            =   1470
            TabIndex        =   1
            Top             =   720
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   450
            _Version        =   393216
            Format          =   "#,##0.00;($#,##0.00)"
            PromptChar      =   "_"
         End
         Begin VB.Label lblCheque 
            Caption         =   "Saldo"
            Height          =   255
            Left            =   120
            TabIndex        =   10
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Data"
            Height          =   255
            Left            =   120
            TabIndex        =   9
            Top             =   360
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmUserRelBalancoCad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                         As tpStatus
Public lngRELBALANCOID                As Long
Public blnRetorno                     As Boolean
Public blnPrimeiraVez                 As Boolean
Public blnFechar                      As Boolean


Private Sub cmdCancelar_Click()
  blnFechar = True
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim objRelBalanco           As busSisMed.clsRelBalanco
  Dim objGeral                As busSisMed.clsGeral
  '
  Dim curSaldoAnterior      As Currency
  Dim datDataSaldoAnterior  As Date
  Dim curReceita            As Currency
  Dim curPrestador          As Currency
  Dim curDespesa            As Currency
  Dim curSaldo              As Currency
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração
    If Not ValidaCampos Then Exit Sub
    'Valida se grupo da despesa já cadastrada
    Set objGeral = New busSisMed.clsGeral
    strSql = "Select PKID From RELBALANCO WHERE DATA = " & Formata_Dados(mskData.Text, tpDados_DataHora) & _
      " AND PKID <> " & lngRELBALANCOID

    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      TratarErroPrevisto "Data referente ao saldo já cadastrada", "cmdOK_Click"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objRelBalanco = New busSisMed.clsRelBalanco
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objRelBalanco.AlterarRelBalanco lngRELBALANCOID, _
                                      mskData.Text, _
                                      mskSaldo.Text
      Set objRelBalanco = Nothing
      blnRetorno = True
      blnFechar = True
      Unload Me
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      'Obter valores
      objRelBalanco.SelecionarSaldoBalanco curSaldoAnterior, _
                                           datDataSaldoAnterior, _
                                           curReceita, _
                                           curPrestador, _
                                           curDespesa, _
                                           mskData.Text, _
                                           mskData.Text
      '
      'Calculo do saldo
      curSaldo = curSaldoAnterior + curReceita - curDespesa
      INCLUIR_VALOR_NO_MASK mskSaldo, curSaldo, TpMaskMoeda
      objRelBalanco.IncluirRelBalanco mskData.Text, _
                                      mskSaldo.Text
      Set objRelBalanco = Nothing
      blnRetorno = True
      blnFechar = True
      Unload Me
    End If
    Set objRelBalanco = Nothing
    'blnFechar = True
    'Unload Me
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
  If Not Valida_Data(mskData, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Status <> tpStatus_Incluir Then
    If Not Valida_Moeda(mskSaldo, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Preencher o saldo válido" & vbCrLf
      tabDetalhes.Tab = 0
    End If
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserRelBalancoCad.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserRelBalancoCad.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco mskData
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserRelBalancoCad.Form_Activate]"
End Sub



Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs               As ADODB.Recordset
  Dim strSql              As String
  Dim objRelBalanco       As busSisMed.clsRelBalanco
  '
  blnFechar = False
  blnRetorno = False
  AmpS
  Me.Height = 4560
  Me.Width = 10200
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    LimparCampoMask mskData
    LimparCampoMask mskSaldo
    '
    mskSaldo.Enabled = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objRelBalanco = New busSisMed.clsRelBalanco
    Set objRs = objRelBalanco.SelecionarRelBalanco(lngRELBALANCOID)
    '
    If Not objRs.EOF Then
      INCLUIR_VALOR_NO_MASK mskData, objRs.Fields("DATA").Value & "", TpMaskData
      INCLUIR_VALOR_NO_MASK mskSaldo, objRs.Fields("SALDO").Value & "", TpMaskMoeda
    End If
    Set objRelBalanco = Nothing
    mskData.Enabled = False
  End If
  '
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub
Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub

Private Sub mskData_GotFocus()
  Selecionar_Conteudo mskData
End Sub

Private Sub mskData_LostFocus()
  Pintar_Controle mskData, tpCorContr_Normal
End Sub

Private Sub mskSaldo_GotFocus()
  Selecionar_Conteudo mskSaldo
End Sub

Private Sub mskSaldo_LostFocus()
  Pintar_Controle mskSaldo, tpCorContr_Normal
End Sub
