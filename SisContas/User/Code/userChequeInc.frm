VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmUserChequeInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de cheques"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5610
      Left            =   8520
      ScaleHeight     =   5610
      ScaleWidth      =   1860
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   30
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   16
         Top             =   3300
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   14
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5295
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do cheque"
      TabPicture(0)   =   "userChequeInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraProf"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraProf 
         Height          =   4695
         Left            =   120
         TabIndex        =   18
         Top             =   360
         Width           =   7935
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   4335
            Index           =   0
            Left            =   120
            ScaleHeight     =   4335
            ScaleWidth      =   7695
            TabIndex        =   19
            Top             =   240
            Width           =   7695
            Begin VB.CommandButton cmdBuscarCheque 
               Caption         =   "&Z"
               Height          =   800
               Left            =   3330
               Style           =   1  'Graphical
               TabIndex        =   1
               TabStop         =   0   'False
               Top             =   0
               Width           =   800
            End
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
               Height          =   3495
               Left            =   0
               TabIndex        =   20
               Top             =   720
               Width           =   7695
               Begin VB.TextBox txtConta 
                  DataField       =   "CONTA"
                  DataSource      =   "Data1"
                  Height          =   288
                  Left            =   5640
                  MaxLength       =   20
                  TabIndex        =   5
                  Text            =   "txtConta"
                  Top             =   600
                  Width           =   1815
               End
               Begin VB.TextBox txtCheque 
                  DataField       =   "CHEQUE"
                  DataSource      =   "Data1"
                  Height          =   288
                  Left            =   1320
                  MaxLength       =   20
                  TabIndex        =   6
                  Text            =   "txtCheque"
                  Top             =   960
                  Width           =   1815
               End
               Begin VB.TextBox txtMotivoDevol 
                  BackColor       =   &H00E0E0E0&
                  Height          =   288
                  Left            =   2040
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   12
                  TabStop         =   0   'False
                  Text            =   "txtMotivoDevol"
                  Top             =   2400
                  Width           =   5415
               End
               Begin VB.TextBox txtNrDevol 
                  Height          =   288
                  Left            =   1320
                  MaxLength       =   4
                  TabIndex        =   11
                  Text            =   "txtNrDevol"
                  Top             =   2400
                  Width           =   735
               End
               Begin VB.TextBox txtBanco 
                  BackColor       =   &H00E0E0E0&
                  Height          =   288
                  Left            =   2040
                  Locked          =   -1  'True
                  MaxLength       =   20
                  TabIndex        =   3
                  TabStop         =   0   'False
                  Text            =   "txtBanco"
                  Top             =   240
                  Width           =   5415
               End
               Begin VB.TextBox txtAgencia 
                  DataField       =   "AGENCIA"
                  DataSource      =   "Data1"
                  Height          =   285
                  Left            =   1320
                  MaxLength       =   15
                  TabIndex        =   4
                  Text            =   "txtAgenc"
                  Top             =   600
                  Width           =   1815
               End
               Begin MSMask.MaskEdBox mskData 
                  Height          =   255
                  Index           =   0
                  Left            =   1320
                  TabIndex        =   8
                  Top             =   1320
                  Width           =   1215
                  _ExtentX        =   2143
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
                  Left            =   1320
                  TabIndex        =   10
                  Top             =   2040
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  _Version        =   393216
                  AutoTab         =   -1  'True
                  MaxLength       =   10
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskValor 
                  Height          =   255
                  Left            =   5640
                  TabIndex        =   7
                  Top             =   960
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   450
                  _Version        =   393216
                  AutoTab         =   -1  'True
                  Format          =   "#,##0.00;($#,##0.00)"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskNrBanco 
                  Height          =   255
                  Left            =   1320
                  TabIndex        =   2
                  Top             =   240
                  Width           =   735
                  _ExtentX        =   1296
                  _ExtentY        =   450
                  _Version        =   393216
                  AutoTab         =   -1  'True
                  MaxLength       =   3
                  Format          =   "000"
                  Mask            =   "###"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskData 
                  Height          =   255
                  Index           =   2
                  Left            =   1320
                  TabIndex        =   9
                  Top             =   1680
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   450
                  _Version        =   393216
                  AutoTab         =   -1  'True
                  MaxLength       =   10
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin VB.Label label 
                  Caption         =   "Dt. Recup."
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   30
                  Top             =   1680
                  Width           =   1215
               End
               Begin VB.Label Label9 
                  Caption         =   "Banco"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   29
                  Top             =   240
                  Width           =   1215
               End
               Begin VB.Label Label3 
                  Caption         =   "Valor"
                  Height          =   255
                  Left            =   4440
                  TabIndex        =   28
                  Top             =   960
                  Width           =   855
               End
               Begin VB.Label label 
                  Caption         =   "Dt. Devolução"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   27
                  Top             =   2040
                  Width           =   1095
               End
               Begin VB.Label Label1 
                  Caption         =   "Conta"
                  Height          =   255
                  Left            =   4440
                  TabIndex        =   26
                  Top             =   600
                  Width           =   615
               End
               Begin VB.Label Label4 
                  Caption         =   "Nr. Cheque"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   25
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.Label label 
                  Caption         =   "Mot. Devolução"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   24
                  Top             =   2400
                  Width           =   1215
               End
               Begin VB.Label Label7 
                  Caption         =   "Agência"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   23
                  Top             =   600
                  Width           =   1215
               End
               Begin VB.Label Label2 
                  Caption         =   "Data Receb."
                  Height          =   255
                  Left            =   120
                  TabIndex        =   22
                  Top             =   1320
                  Width           =   1095
               End
            End
            Begin MSMask.MaskEdBox mskCPF 
               Height          =   255
               Left            =   1320
               TabIndex        =   0
               Top             =   120
               Width           =   1455
               _ExtentX        =   2566
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               Enabled         =   0   'False
               MaxLength       =   12
               Mask            =   "#########/##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label44 
               Caption         =   "CPF"
               Enabled         =   0   'False
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   120
               TabIndex        =   21
               Top             =   120
               Width           =   735
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserChequeInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                 As tpStatus
Public lngCLIENTEID           As Long
Public lngCHEQUEID            As Long
Public strCPF                 As String
Public strStatus              As String
Public bRetorno               As Boolean
Public bFechar                As Boolean
Public sTitulo                As String
Public intQuemChamou          As Integer
Private blnPrimeiraVez        As Boolean


Private Sub cmdBuscarCheque_Click()
  BuscaChequeEmLoc
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

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim clsChq                  As busSisContas.clsCheque
  Dim BANCOID                 As Long
  Dim MOTIVODEVOLID           As Long
  
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração de Cliente
    If Not ValidaCampos(BANCOID, MOTIVODEVOLID) Then Exit Sub
    'Valida se o cheque já é cadastrado
    Set clsChq = New busSisContas.clsCheque
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      '
      clsChq.AlterarCheque lngCLIENTEID, _
                           lngCHEQUEID, _
                           BANCOID, _
                           MOTIVODEVOLID, _
                           txtConta.Text, _
                           txtCheque.Text, _
                           txtAgencia.Text, _
                           mskValor.Text, _
                           mskData(1).Text, _
                           mskData(0).Text, _
                           mskData(2).Text, _
                           IIf(Len(mskData(2).ClipText) > 0, "R", strStatus)
                            
    ElseIf Status = tpStatus_Incluir Then
      'Pega Informações para inserir
      '
      clsChq.InserirCheque lngCLIENTEID, _
                           BANCOID, _
                           MOTIVODEVOLID, _
                           txtConta.Text, _
                           txtCheque.Text, _
                           txtAgencia.Text, _
                           mskValor.Text, _
                           mskData(1).Text, _
                           mskData(0).Text, _
                           mskData(2).Text, _
                           IIf(Len(mskData(2).ClipText) > 0, "R", strStatus)
      '
    End If
    Set clsChq = Nothing

  End Select
  bRetorno = True
  bFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos(ByRef BANCOID As Long, _
                              ByRef MOTIVODEVOLID As Long) As Boolean
  Dim strMsg        As String
  Dim strSql        As String
  Dim objRs         As ADODB.Recordset
  Dim clsGer        As busSisContas.clsGeral
  
  Set clsGer = New busSisContas.clsGeral
  '
  If Len(mskNrBanco.ClipText) <> 3 Then
    strMsg = strMsg & "Informar o número do banco com 3 dígitos" & vbCrLf
    Pintar_Controle mskNrBanco, tpCorContr_Erro
  End If
  If Not IsNumeric(mskNrBanco.Text) Then
    strMsg = strMsg & "Informar o número do banco válido" & vbCrLf
    Pintar_Controle mskNrBanco, tpCorContr_Erro
  End If
  If Len(Trim(txtAgencia.Text)) = 0 Then
    strMsg = strMsg & "Informar o número da agência válido" & vbCrLf
    Pintar_Controle txtAgencia, tpCorContr_Erro
  End If
  If Len(Trim(txtConta.Text)) = 0 Then
    strMsg = strMsg & "Informar o número da conta válido" & vbCrLf
    Pintar_Controle txtConta, tpCorContr_Erro
  End If
  If Len(Trim(txtCheque.Text)) = 0 Then
    strMsg = strMsg & "Informar o número do cheque válido" & vbCrLf
    Pintar_Controle txtCheque, tpCorContr_Erro
  End If
  If Not Valida_Moeda(mskValor, TpObrigatorio) Then
    strMsg = strMsg & "Informar o valor do cheque válido" & vbCrLf
    Pintar_Controle mskValor, tpCorContr_Erro
  End If
  If Not Valida_Data(mskData(0), TpObrigatorio) Then
    strMsg = strMsg & "Informar a data de recebimento válida" & vbCrLf
    Pintar_Controle mskData(0), tpCorContr_Erro
  End If
  If Not Valida_Data(mskData(2), TpNaoObrigatorio) Then
    strMsg = strMsg & "Informar a data de recuperacao válida" & vbCrLf
    Pintar_Controle mskData(2), tpCorContr_Erro
  End If
  If strStatus = "D" Then
    If Not Valida_Data(mskData(1), TpObrigatorio) Then
      strMsg = strMsg & "Informar a data de devolução válida" & vbCrLf
      Pintar_Controle mskData(1), tpCorContr_Erro
    End If
    If Not IsNumeric(txtNrDevol.Text) Then
      strMsg = strMsg & "Informar o número do motivo da devolução válido" & vbCrLf
      Pintar_Controle txtNrDevol, tpCorContr_Erro
    End If
  End If
  If Len(strMsg) = 0 Then 'Não houve erro, validações avançadas
    'Valida banco
    strSql = "Select PKID FROM BANCO WHERE NUMERO = " & Formata_Dados(mskNrBanco.Text, tpDados_Texto, tpNulo_Aceita)
    Set objRs = clsGer.ExecutarSQL(strSql)
    If objRs.EOF Then
      strMsg = strMsg & "Banco não cadastrado" & vbCrLf
      Pintar_Controle mskNrBanco, tpCorContr_Erro
    Else
      BANCOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  If Len(strMsg) = 0 Then 'Não houve erro, validações avançadas
    'Valida Motivo devolução
    If strStatus = "D" Then
      strSql = "Select PKID FROM MOTIVODEVOL WHERE CODMOTIVO = " & txtNrDevol.Text
      Set objRs = clsGer.ExecutarSQL(strSql)
      If objRs.EOF Then
        strMsg = strMsg & "Motivo de devolução não cadastrado" & vbCrLf
        Pintar_Controle txtNrDevol, tpCorContr_Erro
      Else
        MOTIVODEVOLID = objRs.Fields("PKID").Value
      End If
      objRs.Close
      Set objRs = Nothing
    End If
  End If
  If Len(strMsg) = 0 Then 'Não houve erro, validações avançadas
    'vERIFICA DUPLICIDADE
    strSql = "Select PKID FROM CHEQUE WHERE CHEQUE = " & Formata_Dados(txtCheque.Text, tpDados_Texto, tpNulo_Aceita) & _
      " AND BANCOID = " & BANCOID & _
      " AND CONTA = " & Formata_Dados(txtConta.Text, tpDados_Texto, tpNulo_Aceita) & " " & _
      " AND AGENCIA = " & Formata_Dados(txtAgencia.Text, tpDados_Texto, tpNulo_Aceita) & " " & _
      " AND PKID <> " & lngCHEQUEID
    Set objRs = clsGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      'if MsgBox
      strMsg = strMsg & "Cheque já cadastrado" & vbCrLf
      Pintar_Controle txtCheque, tpCorContr_Erro
    
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  '
  Set clsGer = Nothing
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserChequeInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

'Propósito: Buscar os Dados do Cheque cadastrados em Locação
Public Sub BuscaChequeEmLoc()
  On Error GoTo trata
  
  Dim clsGer                  As busSisContas.clsGeral
  
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim bMontarGrid             As Boolean
  '
  Set clsGer = New busSisContas.clsGeral
  '
  bMontarGrid = False
  '
  strSql = "Select * From CONTACORRENTE WHERE " & _
      "CPF = " & Formata_Dados(mskCPF.ClipText, tpDados_Texto) & _
      " And STATUSCC = " & Formata_Dados("CH", tpDados_Texto)
  Set objRs = clsGer.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then
    bMontarGrid = True
  End If
  '
  objRs.Close
  Set objRs = Nothing
  '
  If bMontarGrid Then
    frmUserPlanilhaChqsDevolLis.QuemChamou = 2
    frmUserPlanilhaChqsDevolLis.Show vbModal
  Else
    TratarErroPrevisto "Não há lançamento de cheques para este CPF."
  End If
  '
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserChequeInc.BuscaChequeEmLoc]"
End Sub


Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    If Status = tpStatus_Incluir Then
      tabDetalhes.Tab = 0
      BuscaChequeEmLoc
    Else
      tabDetalhes.Tab = 0
    End If
    blnPrimeiraVez = False
    SetarFoco mskCPF
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[CHQDEV.Form_Activate]"
End Sub


Private Sub mskCPF_GotFocus()
  Selecionar_Conteudo mskCPF
End Sub

Private Sub mskCPF_LostFocus()
  Pintar_Controle mskCPF, tpCorContr_Normal
  If Status = tpStatus_Incluir Or Status = tpStatus_Alterar Then
    If Not TestaCPF(mskCPF.Text) Then
      MsgBox "O número do CPF digitado é inválido !", vbExclamation, TITULOSISTEMA
      Exit Sub
    End If
    '
    tabDetalhes.Tab = 0
    BuscaChequeEmLoc
  End If
End Sub

Private Sub Form_Load()
On Error GoTo trata
  Dim objRs     As ADODB.Recordset
  Dim strSql    As String
  Dim clsChq    As busSisContas.clsCheque
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 5985
  Me.Width = 10470
  Me.Caption = Me.Caption & " - " & sTitulo
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  LerFigurasAvulsas cmdBuscarCheque, "Filtrar.ico", "FiltrarDown.ico", "Pesquisar Cheques"
  
  '
  tabDetalhes_Click 0
  '
  If strStatus = "C" Then
    label(0).Visible = False
    label(1).Visible = False
    mskData(1).Visible = False
    txtNrDevol.Visible = False
    txtMotivoDevol.Visible = False
  Else
    label(0).Visible = True
    label(1).Visible = True
    mskData(1).Visible = True
    txtNrDevol.Visible = True
    txtMotivoDevol.Visible = True
  End If
  mskCPF.Text = strCPF
  '
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    mskNrBanco.Text = "___"
    txtBanco.Text = ""
    txtAgencia.Text = ""
    txtConta.Text = ""
    txtCheque.Text = ""
    txtNrDevol.Text = ""
    txtMotivoDevol.Text = ""
    
    '
  ElseIf Status = tpStatus_Alterar Then
    'Pega Dados do Banco de dados
    Set clsChq = New busSisContas.clsCheque
    Set objRs = clsChq.ListarCheque(lngCHEQUEID)
    '
    If Not objRs.EOF Then
      mskNrBanco.Text = objRs.Fields("NUMERO").Value & ""
      txtBanco.Text = objRs.Fields("NOME").Value & ""
      txtAgencia.Text = objRs.Fields("AGENCIA").Value & ""
      txtConta.Text = objRs.Fields("CONTA").Value & ""
      txtCheque.Text = objRs.Fields("CHEQUE").Value & ""
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DTRECEBIMENTO").Value, TpMaskData
      If strStatus = "D" Then
        INCLUIR_VALOR_NO_MASK mskData(1), objRs.Fields("DTDEVOLUCAO").Value, TpMaskData
        txtNrDevol.Text = objRs.Fields("CODMOTIVO").Value & ""
        txtMotivoDevol.Text = objRs.Fields("DESCMOTIVO").Value & ""
      End If
    End If
    
    Set clsChq = Nothing
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
  If Not bFechar Then Cancel = True
End Sub



Private Sub mskData_GotFocus(Index As Integer)
  Seleciona_Conteudo_Controle mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub

Private Sub mskNrBanco_GotFocus()
  Seleciona_Conteudo_Controle mskNrBanco
End Sub

Private Sub mskNrBanco_LostFocus()
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  Dim clsGer As busSisContas.clsGeral
  '
  Pintar_Controle mskNrBanco, tpCorContr_Normal
  If Len(Trim(mskNrBanco.Text)) <> 3 Or Not IsNumeric(mskNrBanco.Text) Then Exit Sub
  '
  Set clsGer = New busSisContas.clsGeral
  '
  'Valida banco
  strSql = "Select NOME FROM BANCO WHERE NUMERO = '" & mskNrBanco.Text & "'"
  Set objRs = clsGer.ExecutarSQL(strSql)
  If objRs.EOF Then
    txtBanco.Text = ""
  Else
    txtBanco.Text = objRs.Fields("NOME").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set clsGer = Nothing
End Sub

Private Sub mskValor_GotFocus()
  Seleciona_Conteudo_Controle mskValor
End Sub

Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    'dados principais da venda
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
  Case 1
    'Inclusão de Iten do Pedido
    cmdCancelar.Enabled = True
    cmdOk.Enabled = False
    '
  Case 2
    'Vizualização dos Itens do pedido
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    'Montar RecordSet
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "SisContas.frmUserChequeInc.tabDetalhes"
  AmpN
End Sub



Private Sub txtAgencia_GotFocus()
  Seleciona_Conteudo_Controle txtAgencia
End Sub

Private Sub txtAgencia_LostFocus()
  Pintar_Controle txtAgencia, tpCorContr_Normal
End Sub

Private Sub txtCheque_GotFocus()
  Seleciona_Conteudo_Controle txtCheque
End Sub

Private Sub txtCheque_LostFocus()
  Pintar_Controle txtCheque, tpCorContr_Normal
End Sub

Private Sub txtConta_GotFocus()
  Seleciona_Conteudo_Controle txtConta
End Sub

Private Sub txtConta_LostFocus()
  Pintar_Controle txtConta, tpCorContr_Normal
End Sub

Private Sub txtNrDevol_GotFocus()
  Seleciona_Conteudo_Controle txtNrDevol
End Sub

Private Sub txtNrDevol_LostFocus()
  Dim strSql As String
  Dim objRs As ADODB.Recordset
  Dim clsGer As busSisContas.clsGeral
  '
  Pintar_Controle txtNrDevol, tpCorContr_Normal
  If IsNumeric(txtNrDevol.Text) Then 'Não houve erro, validações avançadas
    'Valida Motivo devolução
    If strStatus = "D" Then
      Set clsGer = New busSisContas.clsGeral
      strSql = "Select PKID, DESCMOTIVO FROM MOTIVODEVOL WHERE CODMOTIVO = " & txtNrDevol.Text
      Set objRs = clsGer.ExecutarSQL(strSql)
      If objRs.EOF Then
        txtMotivoDevol.Text = ""
      Else
        txtMotivoDevol.Text = objRs.Fields("DESCMOTIVO").Value & ""
      End If
      objRs.Close
      Set objRs = Nothing
      '
      Set clsGer = Nothing
    End If
  End If
End Sub

