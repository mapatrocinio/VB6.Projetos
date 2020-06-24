VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserMovimentacaoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Movimentação"
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
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   1875
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   1605
         TabIndex        =   9
         Top             =   3000
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Default         =   -1  'True
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4815
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   8493
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da Movimentação"
      TabPicture(0)   =   "userMovimentacaoInc.frx":0000
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
         TabIndex        =   11
         Top             =   480
         Width           =   7335
         Begin VB.TextBox txtDocumento 
            Height          =   285
            Left            =   1560
            MaxLength       =   50
            TabIndex        =   5
            Top             =   2760
            Width           =   3855
         End
         Begin VB.TextBox txtDescricao 
            Height          =   525
            Left            =   1560
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   2
            Text            =   "userMovimentacaoInc.frx":001C
            Top             =   1440
            Width           =   5655
         End
         Begin VB.ComboBox cboDe 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   2040
            Width           =   3855
         End
         Begin VB.ComboBox cboPara 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   2400
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
         Begin VB.Label lblDescrDoc 
            Caption         =   "Obs.: Deixe o campo DOCUMENTO em branco, caso queira que ele seja um sequencial automático"
            ForeColor       =   &H000000FF&
            Height          =   615
            Left            =   1560
            TabIndex        =   18
            Top             =   3120
            Width           =   4095
         End
         Begin VB.Label lblDocumento 
            Caption         =   "Documento"
            Height          =   255
            Left            =   120
            TabIndex        =   17
            Top             =   2760
            Width           =   1095
         End
         Begin VB.Label lblPara 
            Caption         =   "Para"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label Label9 
            Caption         =   "Descrição"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label lblDe 
            Caption         =   "De"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   2040
            Width           =   1215
         End
         Begin VB.Label Label3 
            Caption         =   "Data"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   720
            Width           =   1335
         End
         Begin VB.Label Label4 
            Caption         =   "Valor"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1080
            Width           =   1095
         End
      End
   End
End
Attribute VB_Name = "frmUserMovimentacaoInc"
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


Private Sub cboDe_LostFocus()
  Pintar_Controle cboDe, tpCorContr_Normal
End Sub

Private Sub cboPara_LostFocus()
  Pintar_Controle cboPara, tpCorContr_Normal
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
  Dim objMovimentacao         As busSisContas.clsMovimentacao
  Dim clsGer                  As busSisContas.clsGeral
  Dim lngCONTADEBITOID        As Long
  Dim lngCONTACREDITOID       As Long
  Dim strDocumento            As String
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Despesa
    If Not ValidaCampos Then Exit Sub
    'Valida se unidade de estoque já cadastrada
    Set clsGer = New busSisContas.clsGeral
    strSql = "Select PKID From CONTA WHERE DESCRICAO = " & Formata_Dados(cboDe.Text, tpDados_Texto, tpNulo_NaoAceita)
    Set objRs = clsGer.ExecutarSQL(strSql)
    If objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set clsGer = Nothing
      TratarErroPrevisto "Conta não cadastrada", "cmdOK_Click"
      Exit Sub
      
    Else
      lngCONTADEBITOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    If strStatus = "M" Then
      'Movimentacao
      strSql = "Select PKID From CONTA WHERE DESCRICAO = " & Formata_Dados(cboPara.Text, tpDados_Texto, tpNulo_NaoAceita)
      Set objRs = clsGer.ExecutarSQL(strSql)
      If objRs.EOF Then
        objRs.Close
        Set objRs = Nothing
        Set clsGer = Nothing
        TratarErroPrevisto "Conta não cadastrada", "cmdOK_Click"
        Exit Sub
        
      Else
        lngCONTACREDITOID = objRs.Fields("PKID").Value
      End If
      objRs.Close
      Set objRs = Nothing
    End If
    Set clsGer = Nothing
    Set objMovimentacao = New busSisContas.clsMovimentacao
    If strStatus = "C" Then
      lngCONTACREDITOID = lngCONTADEBITOID
      lngCONTADEBITOID = 0
    End If
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objMovimentacao.AlterarMovimentacao IIf(Len(mskData(0).ClipText) = 0, "", mskData(0).Text), _
                                          txtDocumento.Text, _
                                          IIf(lngCONTADEBITOID = 0, "", lngCONTADEBITOID & ""), _
                                          IIf(lngCONTACREDITOID = 0, "", lngCONTACREDITOID & ""), _
                                          mskValor.Text, _
                                          txtDescricao.Text, _
                                          lngMOVIMENTACAOID
      Set objMovimentacao = Nothing
      bRetorno = True
      bFechar = True
      Unload Me
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      If Len(Trim(txtDocumento.Text)) = 0 And strStatus = "M" Then
        strDocumento = RetornaGravaSequencial("SEQUENCIALMOV")
      Else
        strDocumento = txtDocumento.Text
      End If
      objMovimentacao.IncluirMovimentacao strStatus, _
                                          IIf(Len(mskData(0).ClipText) = 0, "", mskData(0).Text), _
                                          strDocumento, _
                                          IIf(lngCONTADEBITOID = 0, "", lngCONTADEBITOID & ""), _
                                          IIf(lngCONTACREDITOID = 0, "", lngCONTACREDITOID & ""), _
                                          mskValor.Text, _
                                          txtDescricao.Text
      'Limpar campos
      'LimparCampoMask mskData(0)
      LimparCampoMask mskValor
      LimparCampoTexto txtDescricao
      LimparCampoTexto txtDocumento
      'cboDe.ListIndex = -1
      cboPara.ListIndex = -1
      SetarFoco mskValor
      Set objMovimentacao = Nothing
      bRetorno = True
    End If
    'Set objMovimentacao = Nothing
    'bFechar = True
    'Unload Me
  End Select
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
'''  If Len(txtDescricao.Text) = 0 Then
'''    strMsg = strMsg & "Informar a descrição da movimentação válida" & vbCrLf
'''    Pintar_Controle txtDescricao, tpCorContr_Erro
'''  End If
  If strStatus = "M" Then
    'Movimentação
    If Len(cboDe.Text) = 0 Then
      strMsg = strMsg & "Selecionar De" & vbCrLf
      Pintar_Controle cboDe, tpCorContr_Erro
    End If
    If Len(cboPara.Text) = 0 Then
      strMsg = strMsg & "Selecionar Para" & vbCrLf
      Pintar_Controle cboPara, tpCorContr_Erro
    End If
''    If Len(txtDocumento.Text) = 0 Then
''      strMsg = strMsg & "Informar o documento válido" & vbCrLf
''      Pintar_Controle txtDocumento, tpCorContr_Erro
''    End If
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserMovimentacaoInc.ValidaCampos]"
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
  TratarErro Err.Number, Err.Description, "[frmUserMovimentacaoInc.Form_Activate]"
End Sub



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim objMovimentacao As busSisContas.clsMovimentacao
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 5520
  Me.Width = 10200
  CenterForm Me
  blnPrimeiraVez = True
  '
  If strStatus = "M" Then
    Me.Caption = "Cadastro de Movimentações"
    tabDetalhes.TabCaption(0) = "&Dados da Movimentação"
  ElseIf strStatus = "D" Then
    Me.Caption = "Cadastro de Ajustes - Débito"
    tabDetalhes.TabCaption(0) = "&Dados do Ajuste"
    cboPara.Visible = False
    lblPara.Visible = False
    lblDocumento.Visible = False
    txtDocumento.Visible = False
    lblDe.Caption = "Conta Debitada"
  Else
    Me.Caption = "Cadastro de Ajustes - Crédito"
    tabDetalhes.TabCaption(0) = "&Dados do Ajuste"
    cboPara.Visible = False
    lblPara.Visible = False
    lblDocumento.Visible = False
    txtDocumento.Visible = False
    lblDe.Caption = "Conta Creditada"
  End If
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  strSql = "SELECT DESCRICAO FROM CONTA ORDER BY DESCRICAO;"
  PreencheCombo cboDe, strSql, False, True
  PreencheCombo cboPara, strSql, False, True
  If Status = tpStatus_Incluir Then
    If strStatus = "D" Or strStatus = "C" Then
      lblDescrDoc.Visible = False
    End If
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    LimparCampoMask mskData(0)
    LimparCampoMask mskValor
    LimparCampoTexto txtDescricao
    LimparCampoTexto txtDocumento
    cboDe.ListIndex = -1
    cboPara.ListIndex = -1
    '
  ElseIf Status = tpStatus_Alterar Then
    lblDescrDoc.Visible = False
    'Pega Dados do Banco de dados
    Set objMovimentacao = New busSisContas.clsMovimentacao
    Set objRs = objMovimentacao.SelecionarMovimentacao(lngMOVIMENTACAOID)
    '
    If Not objRs.EOF Then
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DATA").Value, TpMaskData
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR").Value, TpMaskMoeda
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      txtDocumento.Text = objRs.Fields("DOCUMENTO").Value & ""
      
      If objRs.Fields("DESCRICAOCONTADEBITO").Value & "" <> "" Then
        cboDe.Text = objRs.Fields("DESCRICAOCONTADEBITO").Value & ""
      End If
      If objRs.Fields("DESCRICAOCONTACREDITO").Value & "" <> "" Then
        If strStatus = "C" Then
          cboDe.Text = objRs.Fields("DESCRICAOCONTACREDITO").Value & ""
        Else
          cboPara.Text = objRs.Fields("DESCRICAOCONTACREDITO").Value & ""
        End If
      End If
    End If
    Set objMovimentacao = Nothing
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

Private Sub mskValor_GotFocus()
  Selecionar_Conteudo mskValor
End Sub

Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub


Private Sub txtDescricao_GotFocus()
  Selecionar_Conteudo txtDescricao
End Sub

Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub

Private Sub txtDocumento_GotFocus()
  Selecionar_Conteudo txtDocumento
End Sub

Private Sub txtDocumento_LostFocus()
  Pintar_Controle txtDocumento, tpCorContr_Normal
End Sub

