VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserContaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Contas"
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2115
         Left            =   90
         ScaleHeight     =   2055
         ScaleWidth      =   1605
         TabIndex        =   7
         Top             =   2820
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
      Height          =   4815
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da conta"
      TabPicture(0)   =   "userContaInc.frx":0000
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
         Height          =   3375
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   7335
         Begin VB.ComboBox cboTipo 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   2040
            Width           =   3255
         End
         Begin VB.TextBox txtDescricao 
            Height          =   525
            Left            =   1560
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   2
            Text            =   "userContaInc.frx":001C
            Top             =   1440
            Width           =   5655
         End
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
            TabIndex        =   10
            Top             =   3480
            Width           =   2295
         End
         Begin MSMask.MaskEdBox mskValorSaldo 
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
         Begin VB.Label Label1 
            Caption         =   "Tipo"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label9 
            Caption         =   "Descrição"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Data do saldo"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Valor do saldo"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   1080
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmUserContaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                         As tpStatus
Public lngCONTAID                     As Long
Public bRetorno                       As Boolean
Public blnPrimeiraVez                 As Boolean
Public bFechar                        As Boolean
Public strTipo                        As String



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
  Dim objConta                As busApler.clsConta
  Dim lngTIPOCONTAID          As Long
  Dim objGeral                As busApler.clsGeral
  Dim objRs                   As ADODB.Recordset
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Despesa
    If Not ValidaCampos Then Exit Sub
    'Pegar TIPOCONTAID
    Set objGeral = New busApler.clsGeral
    strSql = "SELECT * FROM TIPOCONTA WHERE DESCRICAO = " & Formata_Dados(cboTipo.Text, tpDados_Texto, tpNulo_Aceita)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngTIPOCONTAID = objRs.Fields("PKID").Value
    Else
      lngTIPOCONTAID = 0
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    'Valida se unidade de estoque já cadastrada
    Set objConta = New busApler.clsConta
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objConta.AlterarConta IIf(Len(mskValorSaldo.ClipText) = 0, "", mskValorSaldo.Text), _
                            IIf(Len(mskData(0).ClipText) = 0, "", mskData(0).Text), _
                            txtDescricao.Text, _
                            lngCONTAID, _
                            lngTIPOCONTAID
      Set objConta = Nothing
      bRetorno = True
      bFechar = True
      Unload Me
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      objConta.IncluirConta IIf(Len(mskValorSaldo.ClipText) = 0, "", mskValorSaldo.Text), _
                            IIf(Len(mskData(0).ClipText) = 0, "", mskData(0).Text), _
                            txtDescricao.Text, _
                            lngTIPOCONTAID, _
                            glParceiroId
      '
      'Limpar campos
      LimparCampoMask mskData(0)
      LimparCampoMask mskValorSaldo
      LimparCampoTexto txtDescricao
      mskData(0).Text = Format(Now, "DD/MM/YYYY")
      cboTipo.ListIndex = -1
      SetarFoco mskData(0)
      Set objConta = Nothing
      bRetorno = True
    End If
    'Set objConta = Nothing
    'bFechar = True
    'Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg              As String
  Dim objGeral            As busApler.clsGeral
  Dim strSql              As String
  Dim objRs               As ADODB.Recordset
  '
  
  If Not Valida_Data(mskData(0), TpObrigatorio) Then
    strMsg = strMsg & "Informar a data do saldo válida" & vbCrLf
    Pintar_Controle mskData(0), tpCorContr_Erro
  End If
  If Not Valida_Moeda(mskValorSaldo, TpObrigatorio) Then
    strMsg = strMsg & "Informar o valor do saldo válido" & vbCrLf
    Pintar_Controle mskValorSaldo, tpCorContr_Erro
  End If
  If Len(txtDescricao.Text) = 0 Then
    strMsg = strMsg & "Informar a descrição da conta válida" & vbCrLf
    Pintar_Controle txtDescricao, tpCorContr_Erro
  End If
  If strMsg = "" Then
    If Len(Trim(mskData(0).ClipText)) = 0 Or Len(Trim(mskValorSaldo.ClipText)) = 0 Then
      strMsg = strMsg & "Preencher o valor e a data do saldo" & vbCrLf
      If Len(Trim(mskData(0).ClipText)) = 0 Then
        Pintar_Controle mskData(0), tpCorContr_Erro
      End If
      If Len(Trim(mskValorSaldo.ClipText)) = 0 Then
        Pintar_Controle mskValorSaldo, tpCorContr_Erro
      End If
    End If
  End If
  '
  If strMsg = "" Then
    'Digitou Grupo e subGrupo, Preencher
    Set objGeral = New busApler.clsGeral
    strSql = "SELECT * FROM CONTA WHERE " & _
      " PKID <> " & lngCONTAID & _
      " AND DESCRICAO = " & Formata_Dados(txtDescricao.Text, tpDados_Texto, tpNulo_NaoAceita)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      strMsg = strMsg & "Descrição da conta já cadastrada" & vbCrLf
      Pintar_Controle txtDescricao, tpCorContr_Erro
    End If
    Set objGeral = Nothing
    objRs.Close
    Set objRs = Nothing
  End If
'  If Len(cboTipo.Text) = 0 Then
'    strMsg = strMsg & "Selecione um Tipo de Conta" & vbCrLf
'    Pintar_Controle cboTipo, tpCorContr_Erro
'  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserContaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    SetarFoco mskData(0)
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserContaInc.Form_Activate]"
End Sub



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim objConta        As busApler.clsConta
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
  strSql = "SELECT DESCRICAO FROM TIPOCONTA ORDER BY DESCRICAO;"
  PreencheCombo cboTipo, strSql, False, True
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    LimparCampoMask mskData(0)
    LimparCampoMask mskValorSaldo
    LimparCampoTexto txtDescricao
    '
    mskData(0).Text = Format(Now, "DD/MM/YYYY")
  ElseIf Status = tpStatus_Alterar Then
    'Pega Dados do Banco de dados
    Set objConta = New busApler.clsConta
    Set objRs = objConta.SelecionarConta(lngCONTAID)
    '
    If Not objRs.EOF Then
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DTSALDO").Value, TpMaskData
      INCLUIR_VALOR_NO_MASK mskValorSaldo, objRs.Fields("VRSALDO").Value, TpMaskMoeda
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      INCLUIR_VALOR_NO_COMBO objRs.Fields("DESC_TIPOCONTA").Value & "", cboTipo
    End If
    Set objConta = Nothing
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
  If Not bFechar Then Cancel = True
End Sub

Private Sub mskData_GotFocus(Index As Integer)
  Selecionar_Conteudo mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub


Private Sub mskValorSaldo_GotFocus()
  Selecionar_Conteudo mskValorSaldo
End Sub

Private Sub mskValorSaldo_LostFocus()
  Pintar_Controle mskValorSaldo, tpCorContr_Normal
End Sub

Private Sub txtDescricao_GotFocus()
  Selecionar_Conteudo txtDescricao
End Sub

Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub

Private Sub cboTipo_LostFocus()
  Pintar_Controle cboTipo, tpCorContr_Normal
End Sub



