VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmVeiculoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de veículo"
   ClientHeight    =   4620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4620
      Left            =   8520
      ScaleHeight     =   4620
      ScaleWidth      =   1860
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   60
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   2400
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4365
      Left            =   120
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   7699
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do veículo"
      TabPicture(0)   =   "userVeiculoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "picTrava(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.PictureBox picTrava 
         BorderStyle     =   0  'None
         Height          =   3735
         Index           =   0
         Left            =   120
         ScaleHeight     =   3735
         ScaleWidth      =   7695
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   510
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
            Height          =   3585
            Left            =   0
            TabIndex        =   10
            Top             =   0
            Width           =   7695
            Begin VB.TextBox txtObservacao 
               Height          =   915
               Left            =   1200
               MaxLength       =   255
               MultiLine       =   -1  'True
               TabIndex        =   3
               Text            =   "userVeiculoInc.frx":001C
               Top             =   1260
               Width           =   6075
            End
            Begin VB.ComboBox cboModelo 
               Height          =   315
               Left            =   1200
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   240
               Width           =   3855
            End
            Begin MSMask.MaskEdBox mskPlaca 
               Height          =   255
               Left            =   1200
               TabIndex        =   1
               Top             =   600
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   8
               Mask            =   "???-####"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskAno 
               Height          =   255
               Left            =   1200
               TabIndex        =   2
               Top             =   930
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Observação"
               Height          =   195
               Index           =   32
               Left            =   240
               TabIndex        =   14
               Top             =   1305
               Width           =   915
            End
            Begin VB.Label Label1 
               Caption         =   "Ano"
               Height          =   255
               Left            =   240
               TabIndex        =   13
               Top             =   900
               Width           =   675
            End
            Begin VB.Label Label10 
               Caption         =   "Placa"
               Height          =   255
               Left            =   240
               TabIndex        =   12
               Top             =   600
               Width           =   675
            End
            Begin VB.Label Label11 
               Caption         =   "Modelo"
               Height          =   255
               Left            =   240
               TabIndex        =   11
               Top             =   300
               Width           =   1095
            End
         End
      End
   End
End
Attribute VB_Name = "frmVeiculoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngVEICULOID               As Long
Public blnRetorno                 As Boolean
Public blnFechar                  As Boolean
Private blnPrimeiraVez            As Boolean



Private Sub cboModelo_LostFocus()
  Pintar_Controle cboModelo, tpCorContr_Normal
End Sub

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
  Dim objVeiculo              As busElite.clsVeiculo
  Dim objGeral                As busElite.clsGeral
  Dim lngMODELOID             As Long
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Unidade de estoque
    If Not ValidaCampos Then Exit Sub
    '
    Set objGeral = New busElite.clsGeral
    'MODELOID
    lngMODELOID = 0
    strSql = "SELECT PKID FROM MODELO WHERE NOME = " & Formata_Dados(Right(cboModelo.Text, Len(cboModelo.Text) - InStr(1, cboModelo.Text, "/", vbTextCompare)), tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngMODELOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    If lngMODELOID = 0 Then
      Set objGeral = Nothing
      TratarErroPrevisto "Selecionar um modelo", "cmdOK_Click"
      Pintar_Controle cboModelo, tpCorContr_Erro
      SetarFoco cboModelo
      Exit Sub
    End If
    'Valida se veículo já cadastrado
    strSql = "Select * From VEICULO WHERE PLACA = " & Formata_Dados(mskPlaca.Text, tpDados_Texto, tpNulo_NaoAceita) & _
      " AND PKID <> " & Formata_Dados(lngVEICULOID, tpDados_Longo, tpNulo_NaoAceita)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      TratarErroPrevisto "Veículo já cadastrado", "cmdOK_Click"
      Pintar_Controle mskPlaca, tpCorContr_Erro
      SetarFoco mskPlaca
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objVeiculo = New busElite.clsVeiculo
    '
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      '
      objVeiculo.AlterarVeiculo lngVEICULOID, _
                                mskPlaca.Text, _
                                lngMODELOID, _
                                mskAno.Text, _
                                txtObservacao.Text


      blnRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Pega Informações para inserir
      '
      objVeiculo.InserirVeiculo lngVEICULOID, _
                                mskPlaca.Text, _
                                lngMODELOID, _
                                mskAno.Text, _
                                txtObservacao.Text
      '
      blnRetorno = True
    End If
    Set objVeiculo = Nothing
    blnFechar = True
    Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg     As String
  '
  If Len(cboModelo.Text) = 0 Then
    strMsg = strMsg & "Selecionar o modelo" & vbCrLf
    Pintar_Controle cboModelo, tpCorContr_Erro
    SetarFoco cboModelo
  End If
  If Len(mskPlaca.ClipText) = 0 Then
    strMsg = strMsg & "Informar a placa do veículo" & vbCrLf
    Pintar_Controle mskPlaca, tpCorContr_Erro
  Else
    If Len(mskPlaca.ClipText) <> 7 Then
      strMsg = strMsg & "Informar a placa do veículo válida" & vbCrLf
      Pintar_Controle mskPlaca, tpCorContr_Erro
    End If
  End If
  If Len(mskAno.ClipText) = 0 Then
    strMsg = strMsg & "Informar o ano do veículo" & vbCrLf
    Pintar_Controle mskAno, tpCorContr_Erro
  Else
    If Len(mskAno.ClipText) <> 4 Then
      strMsg = strMsg & "Informar o ano do veículo válido" & vbCrLf
      Pintar_Controle mskAno, tpCorContr_Erro
    End If
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserVeiculoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    tabDetalhes.Tab = 0
    cboModelo.SetFocus
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserModeloInc.Form_Activate]"
End Sub

Private Sub mskAno_GotFocus()
  Selecionar_Conteudo mskAno
End Sub

Private Sub mskAno_LostFocus()
  Pintar_Controle mskAno, tpCorContr_Normal
End Sub

Private Sub mskPlaca_KeyPress(KeyAscii As Integer)
  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
End Sub
Private Sub mskPlaca_GotFocus()
  Selecionar_Conteudo mskPlaca
End Sub
Private Sub mskPlaca_LostFocus()
  mskPlaca.Text = UCase(mskPlaca.Text)
  Pintar_Controle mskPlaca, tpCorContr_Normal
End Sub
Private Sub txtObservacao_GotFocus()
  Selecionar_Conteudo txtObservacao
End Sub

Private Sub txtObservacao_LostFocus()
  Pintar_Controle txtObservacao, tpCorContr_Normal
End Sub



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs         As ADODB.Recordset
  Dim strSql        As String
  Dim objVeiculo    As busElite.clsVeiculo
  '
  blnFechar = False
  blnRetorno = False
  AmpS
  Me.Height = 5100
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  LimparCampoCombo cboModelo
  LimparCampoMask mskPlaca
  LimparCampoMask mskAno
  txtObservacao.Text = ""
  '
  'MODELO
  strSql = "SELECT MARCA.NOME + '/' + MODELO.NOME AS MARCA_MODELO " & _
    "FROM MARCA " & _
    " INNER JOIN MODELO ON MARCA.PKID = MODELO.MARCAID " & _
    " ORDER BY MARCA.NOME, MODELO.NOME "
  PreencheCombo cboModelo, strSql, False, True
  '
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Then
    'Pega Dados do Banco de dados
    Set objVeiculo = New busElite.clsVeiculo
    Set objRs = objVeiculo.ListarVeiculo(lngVEICULOID)
    '
    If Not objRs.EOF Then
      If objRs.Fields("MARCA_MODELO").Value & "" <> "" Then
        cboModelo.Text = objRs.Fields("MARCA_MODELO").Value & ""
      End If
      INCLUIR_VALOR_NO_MASK mskPlaca, objRs.Fields("PLACA").Value & "", TpMaskOutros
      INCLUIR_VALOR_NO_MASK mskAno, objRs.Fields("ANO").Value & "", TpMaskOutros
      txtObservacao.Text = objRs.Fields("OBSERVACAO").Value & ""
      '
    End If
    Set objVeiculo = Nothing
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
