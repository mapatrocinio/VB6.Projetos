VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmServicoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Serviço"
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
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2055
         Left            =   90
         ScaleHeight     =   1995
         ScaleWidth      =   1605
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3360
         Width           =   1665
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   990
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5295
      Left            =   120
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userServicoInc.frx":0000
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
         Height          =   4755
         Left            =   120
         TabIndex        =   19
         Top             =   360
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   4515
            Index           =   0
            Left            =   120
            ScaleHeight     =   4515
            ScaleWidth      =   7575
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.TextBox txtVoo 
               Height          =   285
               Left            =   1320
               MaxLength       =   10
               TabIndex        =   10
               Text            =   "txtVoo"
               Top             =   2670
               Width           =   2475
            End
            Begin VB.ComboBox cboDestino 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   3
               Top             =   1050
               Width           =   6105
            End
            Begin VB.ComboBox cboOrigem 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   720
               Width           =   6105
            End
            Begin VB.ComboBox cboAgencia 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   390
               Width           =   6105
            End
            Begin VB.TextBox txtObservacao 
               Height          =   615
               Left            =   1320
               MaxLength       =   255
               MultiLine       =   -1  'True
               TabIndex        =   12
               Text            =   "userServicoInc.frx":001C
               Top             =   3000
               Width           =   6075
            End
            Begin VB.TextBox txtSolicitante 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   4
               Text            =   "txtSolicitante"
               Top             =   1380
               Width           =   6075
            End
            Begin VB.TextBox txtPassageiro 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   5
               Text            =   "txtPassageiro"
               Top             =   1710
               Width           =   6075
            End
            Begin VB.TextBox txtReserva 
               Height          =   285
               Left            =   1320
               MaxLength       =   30
               TabIndex        =   6
               Text            =   "txtReserva"
               Top             =   2040
               Width           =   2475
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1320
               ScaleHeight     =   285
               ScaleWidth      =   2535
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   3690
               Width           =   2535
               Begin VB.OptionButton optStatus 
                  Caption         =   "Finalizado"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   14
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optStatus 
                  Caption         =   "Inicial"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   13
                  TabStop         =   0   'False
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   825
               End
            End
            Begin MSMask.MaskEdBox mskDtHora 
               Height          =   255
               Left            =   1320
               TabIndex        =   0
               Top             =   90
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   16
               Mask            =   "##/##/#### ##:##"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskQtdPassageiro 
               Height          =   255
               Left            =   1320
               TabIndex        =   8
               Top             =   2370
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   2
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskQtdCriancas 
               Height          =   255
               Left            =   5220
               TabIndex        =   9
               Top             =   2340
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   2
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskTerminal 
               Height          =   255
               Left            =   5220
               TabIndex        =   11
               Top             =   2670
               Width           =   855
               _ExtentX        =   1508
               _ExtentY        =   450
               _Version        =   393216
               MaxLength       =   2
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskValor 
               Height          =   255
               Left            =   5220
               TabIndex        =   7
               Top             =   2040
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label3 
               Caption         =   "Valor"
               Height          =   255
               Left            =   3990
               TabIndex        =   36
               Top             =   2040
               Width           =   1455
            End
            Begin VB.Label Label2 
               Caption         =   "Terminal"
               Height          =   255
               Left            =   3990
               TabIndex        =   35
               Top             =   2670
               Width           =   1455
            End
            Begin VB.Label Label5 
               Caption         =   "Voo"
               Height          =   195
               Index           =   3
               Left            =   90
               TabIndex        =   34
               Top             =   2670
               Width           =   1215
            End
            Begin VB.Label Label1 
               Caption         =   "Qt. Crianças"
               Height          =   255
               Left            =   3990
               TabIndex        =   33
               Top             =   2340
               Width           =   1455
            End
            Begin VB.Label Label17 
               Caption         =   "Qt. Passageiros"
               Height          =   255
               Left            =   90
               TabIndex        =   32
               Top             =   2370
               Width           =   1455
            End
            Begin VB.Label Label6 
               Caption         =   "Destino"
               Height          =   255
               Index           =   2
               Left            =   90
               TabIndex        =   31
               Top             =   1080
               Width           =   1455
            End
            Begin VB.Label Label6 
               Caption         =   "Origem"
               Height          =   255
               Index           =   0
               Left            =   90
               TabIndex        =   30
               Top             =   750
               Width           =   1455
            End
            Begin VB.Label Label6 
               Caption         =   "Agência"
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   29
               Top             =   420
               Width           =   1455
            End
            Begin VB.Label Label5 
               Caption         =   "Observação"
               Height          =   195
               Index           =   32
               Left            =   90
               TabIndex        =   28
               Top             =   2985
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Solicitante"
               Height          =   195
               Index           =   2
               Left            =   90
               TabIndex        =   27
               Top             =   1425
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Passageiro"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   26
               Top             =   1755
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Reserva"
               Height          =   195
               Index           =   7
               Left            =   90
               TabIndex        =   25
               Top             =   2040
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Status"
               Height          =   315
               Index           =   5
               Left            =   90
               TabIndex        =   22
               Top             =   3720
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Data/hora"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   21
               Top             =   120
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmServicoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Servico
  LimparCampoMask mskDtHora
  LimparCampoCombo cboAgencia
  LimparCampoCombo cboOrigem
  LimparCampoCombo cboDestino
  LimparCampoTexto txtSolicitante
  LimparCampoTexto txtPassageiro
  LimparCampoTexto txtReserva
  LimparCampoMask mskValor
  LimparCampoMask mskQtdPassageiro
  LimparCampoMask mskQtdCriancas
  LimparCampoTexto txtVoo
  LimparCampoMask mskTerminal
  LimparCampoTexto txtObservacao
  LimparCampoOption optStatus
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmServicoInc.LimparCampos]", _
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


Private Sub cmdOK_Click()
  Dim objServico                As busElite.clsServico
  Dim objGeral                  As busElite.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim strStatus                 As String
  Dim lngAGENCIACNPJID          As Long
  Dim lngORIGEMID               As Long
  Dim lngDESTINOID              As Long
  Dim strCNPJ                   As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busElite.clsGeral
  Set objServico = New busElite.clsServico
  'Status
  If optStatus(0).Value Then
    strStatus = "I"
  Else
    strStatus = "F"
  End If
  'AGENCIACNPJID
  strCNPJ = ""
  strCNPJ = RetornaCPFSemMascara(Left(Right(cboAgencia.Text, 19), 18))
  lngAGENCIACNPJID = 0
  strSql = "SELECT PKID FROM AGENCIACNPJ WHERE AGENCIACNPJ.CNPJ = " & Formata_Dados(strCNPJ, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngAGENCIACNPJID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'ORIGEMID
  lngORIGEMID = 0
  strSql = "SELECT PKID FROM ORIGEM " & _
      " WHERE ORIGEM.IC_ORIGEM IN ('O','A') " & _
    " AND ORIGEM.NOME = " & Formata_Dados(cboOrigem.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngORIGEMID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'DESTINOID
  lngDESTINOID = 0
  strSql = "SELECT PKID FROM ORIGEM DESTINO " & _
      " WHERE DESTINO.IC_ORIGEM IN ('D','A') " & _
      " AND DESTINO.NOME = " & Formata_Dados(cboDestino.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngDESTINOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Valida se obterve os campos com sucesso
  If lngAGENCIACNPJID = 0 Then
    Set objGeral = Nothing
    TratarErroPrevisto "Selecionar uma agência", "cmdOK_Click"
    Pintar_Controle cboAgencia, tpCorContr_Erro
    SetarFoco cboAgencia
    Exit Sub
  End If
  If lngORIGEMID = 0 Then
    Set objGeral = Nothing
    TratarErroPrevisto "Selecionar uma origem", "cmdOK_Click"
    Pintar_Controle cboOrigem, tpCorContr_Erro
    SetarFoco cboOrigem
    Exit Sub
  End If
  If lngDESTINOID = 0 Then
    Set objGeral = Nothing
    TratarErroPrevisto "Selecionar um destino", "cmdOK_Click"
    Pintar_Controle cboDestino, tpCorContr_Erro
    SetarFoco cboDestino
    Exit Sub
  End If
  'Validar se serviço já cadastrado
  strSql = "SELECT * FROM SERVICO " & _
    " WHERE SERVICO.DATAHORA = " & Formata_Dados(mskDtHora.Text, tpDados_DataHora) & _
    " AND SERVICO.SOLICITANTE = " & Formata_Dados(txtSolicitante.Text, tpDados_Texto) & _
    " AND SERVICO.AGENCIACNPJID = " & Formata_Dados(lngAGENCIACNPJID, tpDados_Longo) & _
    " AND SERVICO.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle mskDtHora, tpCorContr_Erro
    TratarErroPrevisto "Data/Solicitante/Agência já cadastrada"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objServico = Nothing
    cmdOk.Enabled = True
    SetarFoco mskDtHora
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Servico
    objServico.AlterarServico lngPKID, _
                              mskDtHora.Text, _
                              lngAGENCIACNPJID, _
                              lngORIGEMID, _
                              lngDESTINOID, _
                              txtSolicitante.Text, _
                              txtPassageiro.Text, _
                              txtReserva.Text, _
                              IIf(mskQtdPassageiro.ClipText = "", "", mskQtdPassageiro.Text), _
                              IIf(mskQtdCriancas.ClipText = "", "", mskQtdCriancas.Text), _
                              txtVoo.Text, _
                              IIf(mskTerminal.ClipText = "", "", mskTerminal.Text), _
                              IIf(mskValor.ClipText = "", "", mskValor.Text), _
                              txtObservacao.Text, _
                              strStatus
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Servico
    objServico.InserirServico lngPKID, _
                              mskDtHora.Text, _
                              lngAGENCIACNPJID, _
                              lngORIGEMID, _
                              lngDESTINOID, _
                              txtSolicitante.Text, _
                              txtPassageiro.Text, _
                              txtReserva.Text, _
                              IIf(mskQtdPassageiro.ClipText = "", "", mskQtdPassageiro.Text), _
                              IIf(mskQtdCriancas.ClipText = "", "", mskQtdCriancas.Text), _
                              txtVoo.Text, _
                              IIf(mskTerminal.ClipText = "", "", mskTerminal.Text), _
                              IIf(mskValor.ClipText = "", "", mskValor.Text), _
                              txtObservacao.Text
    blnRetorno = True
    'Selecionar plano cadastrado
    'entrar em modo de alteração
    'lngPKID = objRs.Fields("PKID")
    Status = tpStatus_Alterar
    'Reload na tela
    Form_Load
    'Acerta tabs
    'tabDetalhes.Tab = 2
    blnRetorno = True
  End If
  Set objServico = Nothing
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
  If Not Valida_Data(mskDtHora, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(cboAgencia, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar a agência" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(cboOrigem, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar a origem" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(cboDestino, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o destino" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(txtSolicitante, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o solicitante" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(txtPassageiro, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o passageiro" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskValor, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencha o valor válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskQtdPassageiro, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a quantidade de passageiros válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskQtdCriancas, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a quantidade de crianças válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskTerminal, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número do terminal válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Option(optStatus, blnSetarFocoControle) Then
    strMsg = strMsg & "Slecionar o status" & vbCrLf
    tabDetalhes.Tab = 0
  End If

  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmServicoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmServicoInc.ValidaCampos]", _
            Err.Description
End Function



Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco mskDtHora
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmServicoInc.Form_Activate]"
End Sub



Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objServico              As busElite.clsServico
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
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  tabDetalhes_Click 0
  '
  'Combos
  'AGENCIACNPJ
  strSql = "Select AGENCIA.NOME + ' (' + dbo.formataCNPJ(AGENCIACNPJ.CNPJ) + ')' FROM AGENCIA  " & _
      " INNER JOIN AGENCIACNPJ ON AGENCIA.PKID = AGENCIACNPJ.AGENCIAID " & _
      "ORDER BY AGENCIA.NOME, AGENCIACNPJ.CNPJ"
  PreencheCombo cboAgencia, strSql, False, True
  'ORIGEM
  strSql = "SELECT ORIGEM.NOME FROM ORIGEM " & _
      " WHERE ORIGEM.IC_ORIGEM IN ('O','A') " & _
      "ORDER BY ORIGEM.NOME"
  PreencheCombo cboOrigem, strSql, False, True
  'DESTINO
  strSql = "SELECT DESTINO.NOME FROM ORIGEM DESTINO " & _
      " WHERE DESTINO.IC_ORIGEM IN ('D','A') " & _
      "ORDER BY DESTINO.NOME"
  PreencheCombo cboDestino, strSql, False, True
  'Desabilita status
  Picture1.Enabled = False
  If Status = tpStatus_Incluir Then
    '
    optStatus(0).Value = True
    'Visible
    optStatus(0).Visible = False
    optStatus(1).Visible = False
    Label5(5).Visible = False
    '
    tabDetalhes.TabEnabled(0) = True
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objServico = New busElite.clsServico
    Set objRs = objServico.SelecionarServicoPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      INCLUIR_VALOR_NO_MASK mskDtHora, objRs.Fields("DATAHORA").Value, TpMaskData
      cboAgencia.Text = objRs.Fields("DESC_AGENCIA").Value
      cboOrigem.Text = objRs.Fields("DESC_ORIGEM").Value
      cboDestino.Text = objRs.Fields("DESC_DESTINO").Value
      txtSolicitante = objRs.Fields("SOLICITANTE").Value & ""
      txtPassageiro = objRs.Fields("PASSAGEIRO").Value & ""
      txtReserva = objRs.Fields("RESERVA").Value & ""
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskQtdPassageiro, objRs.Fields("QTDPASSAGEIRO").Value, TpMaskLongo
      INCLUIR_VALOR_NO_MASK mskQtdCriancas, objRs.Fields("QTDCRIANCAS").Value, TpMaskLongo
      txtVoo = objRs.Fields("VOO").Value & ""
      INCLUIR_VALOR_NO_MASK mskTerminal, objRs.Fields("TERMINAL").Value, TpMaskLongo
      txtObservacao = objRs.Fields("OBSERVACAO").Value & ""
      If objRs.Fields("STATUS").Value & "" = "I" Then
        optStatus(0).Value = True
        optStatus(1).Value = False
      ElseIf objRs.Fields("STATUS").Value & "" = "F" Then
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
    Set objServico = Nothing
    'Visible
    optStatus(0).Visible = True
    optStatus(1).Visible = True
    Label5(5).Visible = True
    '
    tabDetalhes.TabEnabled(0) = True
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


Private Sub mskDtHora_GotFocus()
  Seleciona_Conteudo_Controle mskDtHora
End Sub
Private Sub mskDtHora_LostFocus()
  Pintar_Controle mskDtHora, tpCorContr_Normal
End Sub



Private Sub mskQtdCriancas_GotFocus()
  Seleciona_Conteudo_Controle mskQtdCriancas
End Sub
Private Sub mskQtdCriancas_LostFocus()
  Pintar_Controle mskQtdCriancas, tpCorContr_Normal
End Sub

Private Sub mskQtdPassageiro_GotFocus()
  Seleciona_Conteudo_Controle mskQtdPassageiro
End Sub
Private Sub mskQtdPassageiro_LostFocus()
  Pintar_Controle mskQtdPassageiro, tpCorContr_Normal
End Sub

Private Sub mskTerminal_GotFocus()
  Seleciona_Conteudo_Controle mskTerminal
End Sub
Private Sub mskTerminal_LostFocus()
  Pintar_Controle mskTerminal, tpCorContr_Normal
End Sub

Private Sub mskValor_GotFocus()
  Seleciona_Conteudo_Controle mskValor
End Sub
Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub


Private Sub txtObservacao_GotFocus()
  Seleciona_Conteudo_Controle txtObservacao
End Sub
Private Sub txtObservacao_LostFocus()
  Pintar_Controle txtObservacao, tpCorContr_Normal
End Sub


Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    pictrava(0).Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    '
    SetarFoco mskDtHora
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "Elite.frmServicoInc.tabDetalhes"
  AmpN
End Sub


Private Sub txtPassageiro_GotFocus()
  Seleciona_Conteudo_Controle txtPassageiro
End Sub
Private Sub txtPassageiro_LostFocus()
  Pintar_Controle txtPassageiro, tpCorContr_Normal
End Sub


Private Sub txtReserva_GotFocus()
  Seleciona_Conteudo_Controle txtReserva
End Sub
Private Sub txtReserva_LostFocus()
  Pintar_Controle txtReserva, tpCorContr_Normal
End Sub

Private Sub txtSolicitante_GotFocus()
  Seleciona_Conteudo_Controle txtSolicitante
End Sub
Private Sub txtSolicitante_LostFocus()
  Pintar_Controle txtSolicitante, tpCorContr_Normal
End Sub

Private Sub txtVoo_GotFocus()
  Seleciona_Conteudo_Controle txtVoo
End Sub
Private Sub txtVoo_LostFocus()
  Pintar_Controle txtVoo, tpCorContr_Normal
End Sub



Private Sub cboAgencia_LostFocus()
  Pintar_Controle cboAgencia, tpCorContr_Normal
End Sub

Private Sub cboDestino_LostFocus()
  Pintar_Controle cboDestino, tpCorContr_Normal
End Sub

Private Sub cboOrigem_LostFocus()
  Pintar_Controle cboOrigem, tpCorContr_Normal
End Sub


