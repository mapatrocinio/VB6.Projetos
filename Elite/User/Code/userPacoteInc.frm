VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmPacoteInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Pacote"
   ClientHeight    =   5610
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12180
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5610
   ScaleWidth      =   12180
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5610
      Left            =   10320
      ScaleHeight     =   5610
      ScaleWidth      =   1860
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   3765
         Left            =   90
         ScaleHeight     =   3705
         ScaleWidth      =   1605
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   1620
         Width           =   1665
         Begin VB.CommandButton cmdAlterarPlaca 
            Caption         =   "&YHQ-3995"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   9
            ToolTipText     =   "Alterar Placa"
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   1020
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   1890
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5295
      Left            =   120
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   10095
      _ExtentX        =   17806
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userPacoteInc.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Serviços"
      TabPicture(1)   =   "userPacoteInc.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "grdServico"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Histórico"
      TabPicture(2)   =   "userPacoteInc.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdHistorico"
      Tab(2).ControlCount=   1
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
         Left            =   -74880
         TabIndex        =   15
         Top             =   360
         Width           =   9855
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   4515
            Index           =   0
            Left            =   120
            ScaleHeight     =   4515
            ScaleWidth      =   9645
            TabIndex        =   16
            TabStop         =   0   'False
            Top             =   150
            Width           =   9645
            Begin VB.ComboBox cboMotorista 
               Height          =   315
               Left            =   1530
               Style           =   2  'Dropdown List
               TabIndex        =   2
               Top             =   690
               Width           =   5955
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1530
               ScaleHeight     =   285
               ScaleWidth      =   4785
               TabIndex        =   19
               TabStop         =   0   'False
               Top             =   1350
               Width           =   4785
               Begin VB.OptionButton optStatus 
                  Caption         =   "Pago"
                  Height          =   315
                  Index           =   2
                  Left            =   2040
                  TabIndex        =   6
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optStatus 
                  Caption         =   "Concluído"
                  Height          =   315
                  Index           =   1
                  Left            =   870
                  TabIndex        =   5
                  TabStop         =   0   'False
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optStatus 
                  Caption         =   "Inicial"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   4
                  TabStop         =   0   'False
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   825
               End
            End
            Begin MSMask.MaskEdBox mskDtHoraInicio 
               Height          =   255
               Left            =   1530
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
            Begin MSMask.MaskEdBox mskValor 
               Height          =   255
               Left            =   1530
               TabIndex        =   3
               Top             =   1050
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskDtHoraTermino 
               Height          =   255
               Left            =   1530
               TabIndex        =   1
               Top             =   390
               Width           =   1875
               _ExtentX        =   3307
               _ExtentY        =   450
               _Version        =   393216
               AutoTab         =   -1  'True
               MaxLength       =   16
               Mask            =   "##/##/#### ##:##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Data/hora Término"
               Height          =   195
               Index           =   1
               Left            =   90
               TabIndex        =   23
               Top             =   420
               Width           =   1425
            End
            Begin VB.Label Label3 
               Caption         =   "Valor"
               Height          =   255
               Left            =   120
               TabIndex        =   22
               Top             =   1050
               Width           =   1455
            End
            Begin VB.Label Label6 
               Caption         =   "Motorista"
               Height          =   255
               Index           =   1
               Left            =   90
               TabIndex        =   21
               Top             =   720
               Width           =   1455
            End
            Begin VB.Label Label5 
               Caption         =   "Status"
               Height          =   315
               Index           =   5
               Left            =   90
               TabIndex        =   18
               Top             =   1440
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Data/hora Início"
               Height          =   195
               Index           =   0
               Left            =   90
               TabIndex        =   17
               Top             =   120
               Width           =   1215
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdServico 
         Height          =   4410
         Left            =   60
         OleObjectBlob   =   "userPacoteInc.frx":0054
         TabIndex        =   7
         Top             =   390
         Width           =   9930
      End
      Begin TrueDBGrid60.TDBGrid grdHistorico 
         Height          =   4410
         Left            =   -74940
         OleObjectBlob   =   "userPacoteInc.frx":4AE8
         TabIndex        =   8
         Top             =   390
         Width           =   9870
      End
   End
End
Attribute VB_Name = "frmPacoteInc"
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

Dim SERV_COLUNASMATRIZ         As Long
Dim SERV_LINHASMATRIZ          As Long

Private SERV_Matriz()          As String

Dim HIST_COLUNASMATRIZ         As Long
Dim HIST_LINHASMATRIZ          As Long

Private HIST_Matriz()          As String

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Pacote
  LimparCampoMask mskDtHoraInicio
  LimparCampoMask mskDtHoraTermino
  LimparCampoCombo cboMotorista
  LimparCampoMask mskValor
  LimparCampoOption optStatus
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmPacoteInc.LimparCampos]", _
            Err.Description
End Sub

Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 1
    frmServicoAssoc.strCaption = mskDtHoraInicio.Text & " a " & mskDtHoraTermino.Text & " - " & cboMotorista.Text
    frmServicoAssoc.lngPACOTEID = lngPKID
    frmServicoAssoc.Show vbModal
    If frmServicoAssoc.blnRetorno Then
      
      'Montar RecordSet
      SERV_COLUNASMATRIZ = grdServico.Columns.Count
      SERV_LINHASMATRIZ = 0
      SERV_MontaMatriz
      grdServico.Bookmark = Null
      grdServico.ReBind
      grdServico.ApproxCount = SERV_LINHASMATRIZ
      '
      SetarFoco grdServico
    End If
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdAlterarPlaca_Click()
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 1
    If Not IsNumeric(grdServico.Columns("PACOTESERVICOID").Value & "") Then
      MsgBox "Selecione um serviço!", vbExclamation, TITULOSISTEMA
      SetarFoco grdServico
      Exit Sub
    End If
    frmServicoPlacaInc.strCaption = mskDtHoraInicio.Text & " a " & mskDtHoraTermino.Text & " - " & cboMotorista.Text
    frmServicoPlacaInc.strPlaca = grdServico.Columns("Placa").Value
    frmServicoPlacaInc.lngPACOTESERVICOID = grdServico.Columns("PACOTESERVICOID").Value
    frmServicoPlacaInc.Show vbModal
    If frmServicoPlacaInc.blnRetorno Then
      
      'Montar RecordSet
      SERV_COLUNASMATRIZ = grdServico.Columns.Count
      SERV_LINHASMATRIZ = 0
      SERV_MontaMatriz
      grdServico.Bookmark = Null
      grdServico.ReBind
      grdServico.ApproxCount = SERV_LINHASMATRIZ
      '
      SetarFoco grdServico
    End If
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
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
  Dim objPacote                 As busElite.clsPacote
  Dim objGeral                  As busElite.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim strStatus                 As String
  Dim lngMOTORISTAID            As Long
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busElite.clsGeral
  Set objPacote = New busElite.clsPacote
  'Status
  If optStatus(0).Value Then
    strStatus = "I"
  ElseIf optStatus(1).Value Then
    strStatus = "C"
  Else
    strStatus = "P"
  End If
  'MOTORISTAID
  lngMOTORISTAID = 0
  strSql = "SELECT PKID FROM PESSOA " & _
    " INNER JOIN MOTORISTA ON PESSOA.PKID = MOTORISTA.PESSOAID " & _
    " WHERE PESSOA.NOME = " & Formata_Dados(cboMotorista.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngMOTORISTAID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  If lngMOTORISTAID = 0 Then
    Set objGeral = Nothing
    TratarErroPrevisto "Selecionar o motorista", "cmdOK_Click"
    Pintar_Controle cboMotorista, tpCorContr_Erro
    SetarFoco cboMotorista
    Exit Sub
  End If
'''  'Validar se pacote já cadastrado
'''  strSql = "SELECT * FROM SERVICO " & _
'''    " WHERE SERVICO.DATAHORA = " & Formata_Dados(mskDtHoraInicio.Text, tpDados_DataHora) & _
'''    " AND SERVICO.SOLICITANTE = " & Formata_Dados(txtSolicitante.Text, tpDados_Texto) & _
'''    " AND SERVICO.AGENCIACNPJID = " & Formata_Dados(lngMOTORISTAID, tpDados_Longo) & _
'''    " AND SERVICO.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
'''  Set objRs = objGeral.ExecutarSQL(strSql)
'''  If Not objRs.EOF Then
'''    Pintar_Controle mskDtHoraInicio, tpCorContr_Erro
'''    TratarErroPrevisto "Data/Solicitante/Agência já cadastrada"
'''    objRs.Close
'''    Set objRs = Nothing
'''    Set objGeral = Nothing
'''    Set objPacote = Nothing
'''    cmdOk.Enabled = True
'''    SetarFoco mskDtHoraInicio
'''    tabDetalhes.Tab = 0
'''    Exit Sub
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Pacote
    objPacote.AlterarPacote lngPKID, _
                            mskDtHoraInicio.Text, _
                            mskDtHoraTermino.Text, _
                            lngMOTORISTAID, _
                            IIf(mskValor.ClipText = "", "", mskValor.Text), _
                            strStatus
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Pacote
    objPacote.InserirPacote lngPKID, _
                            mskDtHoraInicio.Text, _
                            mskDtHoraTermino.Text, _
                            lngMOTORISTAID, _
                            IIf(mskValor.ClipText = "", "", mskValor.Text)
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
    'Novo abre a tela para associação de serviços
    tabDetalhes.Tab = 1
    cmdAlterar_Click
  End If
  Set objPacote = Nothing
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
  If Not Valida_Data(mskDtHoraInicio, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de início válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskDtHoraTermino, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de término válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(cboMotorista, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o motorista" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskValor, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencha o valor válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Option(optStatus, blnSetarFocoControle) Then
    strMsg = strMsg & "Slecionar o status" & vbCrLf
    tabDetalhes.Tab = 0
  End If

  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmPacoteInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmPacoteInc.ValidaCampos]", _
            Err.Description
End Function



Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco mskDtHoraInicio
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmPacoteInc.Form_Activate]"
End Sub



Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objPacote              As busElite.clsPacote
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 6090
  Me.Width = 12270
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, , , , cmdAlterar
  '
  'Limpar Campos
  LimparCampos
  tabDetalhes_Click 0
  tabDetalhes.TabVisible(1) = False
  '
  'MOTORISTA
  strSql = "SELECT PESSOA.NOME FROM MOTORISTA " & _
      " INNER JOIN PESSOA ON PESSOA.PKID = MOTORISTA.PESSOAID " & _
      "ORDER BY PESSOA.NOME"
  PreencheCombo cboMotorista, strSql, False, True
  'Picture1 desabilita botões de status, não pode haver alteração do status
  'Alteração apenas via função
  Picture1.Enabled = False
  'Desabilita data final do pacote
  mskDtHoraTermino.Enabled = False
  Label5(1).Enabled = False
  If Status = tpStatus_Incluir Then
    '
    optStatus(0).Value = True
    'Visible
    optStatus(0).Visible = False
    optStatus(1).Visible = False
    optStatus(2).Visible = False
    Label5(5).Visible = False
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = False
    tabDetalhes.TabEnabled(2) = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objPacote = New busElite.clsPacote
    Set objRs = objPacote.SelecionarPacotePeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      INCLUIR_VALOR_NO_MASK mskDtHoraInicio, objRs.Fields("DATAINICIO").Value, TpMaskData
      INCLUIR_VALOR_NO_MASK mskDtHoraTermino, objRs.Fields("DATATERMINO").Value, TpMaskData
      cboMotorista.Text = objRs.Fields("DESC_MOTORISTA").Value
      INCLUIR_VALOR_NO_MASK mskValor, objRs.Fields("VALOR").Value, TpMaskMoeda
      If objRs.Fields("STATUS").Value & "" = "I" Then
        optStatus(0).Value = True
        optStatus(1).Value = False
        optStatus(2).Value = False
      ElseIf objRs.Fields("STATUS").Value & "" = "C" Then
        optStatus(0).Value = False
        optStatus(1).Value = True
        optStatus(2).Value = False
      Else
        optStatus(0).Value = False
        optStatus(1).Value = False
        optStatus(2).Value = True
      End If
      
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objPacote = Nothing
    'Visible
    optStatus(0).Visible = True
    optStatus(1).Visible = True
    optStatus(2).Visible = True
    Label5(5).Visible = True
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = True
    tabDetalhes.TabVisible(1) = True
    tabDetalhes.TabEnabled(2) = True
    tabDetalhes.TabVisible(2) = True
    
    '
    
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




Private Sub mskDtHoraInicio_GotFocus()
  Seleciona_Conteudo_Controle mskDtHoraInicio
End Sub
Private Sub mskDtHoraInicio_LostFocus()
  Pintar_Controle mskDtHoraInicio, tpCorContr_Normal
End Sub



Private Sub mskDtHoraTermino_GotFocus()
  Seleciona_Conteudo_Controle mskDtHoraTermino
End Sub
Private Sub mskDtHoraTermino_LostFocus()
  Pintar_Controle mskDtHoraTermino, tpCorContr_Normal
End Sub

Private Sub mskValor_GotFocus()
  Seleciona_Conteudo_Controle mskValor
End Sub
Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub


Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    If Status = tpStatus_Consultar Then
      grdServico.Enabled = False
      grdHistorico.Enabled = False
      pictrava(0).Enabled = False
      '
      cmdAlterar.Enabled = False
      cmdAlterarPlaca.Enabled = False
      cmdOk.Enabled = False
      
    Else
      grdServico.Enabled = False
      grdHistorico.Enabled = False
      pictrava(0).Enabled = True
      '
      cmdAlterar.Enabled = False
      cmdAlterarPlaca.Enabled = False
      cmdOk.Enabled = True
    
    End If
    cmdCancelar.Enabled = True
    '
    SetarFoco mskDtHoraInicio
  Case 1
    If Status = tpStatus_Consultar Then
      grdServico.Enabled = True
      grdHistorico.Enabled = False
      pictrava(0).Enabled = False
      '
      cmdAlterar.Enabled = False
      cmdAlterarPlaca.Enabled = False
      cmdOk.Enabled = False
    Else
      grdServico.Enabled = True
      grdHistorico.Enabled = False
      pictrava(0).Enabled = False
      '
      cmdAlterar.Enabled = True
      cmdAlterarPlaca.Enabled = True
      cmdOk.Enabled = False
    
    End If
    cmdCancelar.Enabled = True
    'Montar RecordSet
    SERV_COLUNASMATRIZ = grdServico.Columns.Count
    SERV_LINHASMATRIZ = 0
    SERV_MontaMatriz
    grdServico.Bookmark = Null
    grdServico.ReBind
    grdServico.ApproxCount = SERV_LINHASMATRIZ
    '
    SetarFoco grdServico
  Case 2
    If Status = tpStatus_Consultar Then
      grdServico.Enabled = False
      grdHistorico.Enabled = True
      pictrava(0).Enabled = False
      '
      cmdAlterar.Enabled = False
      cmdAlterarPlaca.Enabled = False
      cmdOk.Enabled = False
    Else
      grdServico.Enabled = False
      grdHistorico.Enabled = True
      pictrava(0).Enabled = False
      '
      cmdAlterar.Enabled = False
      cmdAlterarPlaca.Enabled = False
      cmdOk.Enabled = False
    
    End If
    cmdCancelar.Enabled = True
    'Montar RecordSet
    HIST_COLUNASMATRIZ = grdHistorico.Columns.Count
    HIST_LINHASMATRIZ = 0
    HIST_MontaMatriz
    grdHistorico.Bookmark = Null
    grdHistorico.ReBind
    grdHistorico.ApproxCount = HIST_LINHASMATRIZ
    '
    SetarFoco grdHistorico
    'MODELO.NOME + ' (' + VEICULO.PLACA + ')'  AS DESC_VEICULO,
    'VEICULO.PLACA + ' - ' + MARCA.NOME,
  'strSql = "SELECT PACOTE.PKID, PACOTE.DATAINICIO, PACOTE.DATATERMINO, PESSOA.NOME, PACOTE.VALOR, " & _
        " PACOTE.STATUS " & _
        "FROM PACOTE LEFT JOIN MOTORISTA ON MOTORISTA.PESSOAID = PACOTE.MOTORISTAID " & _
        " LEFT JOIN PESSOA ON PESSOA.PKID = MOTORISTA.PESSOAID " & _
        " LEFT JOIN VEICULO ON VEICULO.PKID = PACOTE.VEICULOID " & _
        " LEFT JOIN MODELO ON MODELO.PKID = VEICULO.MODELOID " & _
        " LEFT JOIN MARCA ON MARCA.PKID = MODELO.MARCAID " & _
        " ORDER BY PACOTE.DATAINICIO DESC, PACOTE.DATATERMINO DESC;"
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "Elite.frmPacoteInc.tabDetalhes"
  AmpN
End Sub


Private Sub grdServico_UnboundReadDataEx( _
     ByVal RowBuf As TrueDBGrid60.RowBuffer, _
    StartLocation As Variant, ByVal Offset As Long, _
    ApproximatePosition As Long)
  ' UnboundReadData is fired by an unbound grid whenever
  ' it requires data for display. This event will fire
  ' when the grid is first shown, when Refresh or ReBind
  ' is used, when the grid is scrolled, and after a
  ' record in the grid is modified and the user commits
  ' the change by moving off of the current row. The
  ' grid fetches data in "chunks", and the number of rows
  ' the grid is asking for is given by RowBuf.RowCount.
  ' RowBuf is the row buffer where you place the data
  ' the bookmarks for the rows that the grid is
  ' requesting to display. It will also hold the number
  ' of rows that were successfully supplied to the grid.
  ' StartLocation is a vrtBookmark which, together with
  ' Offset, specifies the row for the programmer to start
  ' transferring data. A StartLocation of Null indicates
  ' a request for data from BOF or EOF.
  ' Offset specifies the relative position (from
  ' StartLocation) of the row for the programmer to start
  ' transferring data. A positive number indicates a
  ' forward relative position while a negative number
  ' indicates a backward relative position. Regardless
  ' of whether the rows to be read are before or after
  ' StartLocation, rows are always fetched going forward
  ' (this is why there is no ReadPriorRows parameter to
  ' the procedure).
  ' If you page down on the grid, for instance, the new
  ' top row of the grid will have an index greater than
  ' the StartLocation (Offset > 0). If you page up on
  ' the grid, the new index is less than that of
  ' StartLocation, so Offset < 0. If StartLocation is
  ' a vrtBookmark to row N, the grid always asks for row
  ' data in the following order:
  '   (N + Offset), (N + Offset + 1), (N + Offset + 2)...
  ' ApproximatePosition is a value you can set to indicate
  ' the ordinal position of (StartLocation + Offset).
  ' Setting this variable will enhance the ability of the
  ' grid to display its vertical scroll bar accurately.
  ' If the exact ordinal position of the new location is
  ' not known, you can set it to a reasonable,
  ' approximate value, or just ignore this parameter.

  On Error GoTo trata
  '
  Dim intColIndex      As Integer
  Dim intJ             As Integer
  Dim intRowsFetched   As Integer
  Dim intI             As Long
  Dim lngNewPosition   As Long
  Dim vrtBookmark      As Variant
  '
  intRowsFetched = 0
  For intI = 0 To RowBuf.RowCount - 1
    ' Get the vrtBookmark of the next available row
    vrtBookmark = GetRelativeBookmarkGeral(StartLocation, _
               Offset + intI, SERV_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, SERV_COLUNASMATRIZ, SERV_LINHASMATRIZ, SERV_Matriz)
    Next intJ

    ' Set the vrtBookmark for the row
    RowBuf.Bookmark(intI) = vrtBookmark

    ' Increment the count of fetched rows
    intRowsFetched = intRowsFetched + 1
  Next intI

' Tell the grid how many rows were fetched
  RowBuf.RowCount = intRowsFetched

' Set the approximate scroll bar position. Only
' nonnegative values of IndexFromBookmark() are valid.
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, SERV_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmPacoteInc.grdServico_UnboundReadDataEx]"
End Sub

Private Sub grdHistorico_UnboundReadDataEx( _
     ByVal RowBuf As TrueDBGrid60.RowBuffer, _
    StartLocation As Variant, ByVal Offset As Long, _
    ApproximatePosition As Long)
  ' UnboundReadData is fired by an unbound grid whenever
  ' it requires data for display. This event will fire
  ' when the grid is first shown, when Refresh or ReBind
  ' is used, when the grid is scrolled, and after a
  ' record in the grid is modified and the user commits
  ' the change by moving off of the current row. The
  ' grid fetches data in "chunks", and the number of rows
  ' the grid is asking for is given by RowBuf.RowCount.
  ' RowBuf is the row buffer where you place the data
  ' the bookmarks for the rows that the grid is
  ' requesting to display. It will also hold the number
  ' of rows that were successfully supplied to the grid.
  ' StartLocation is a vrtBookmark which, together with
  ' Offset, specifies the row for the programmer to start
  ' transferring data. A StartLocation of Null indicates
  ' a request for data from BOF or EOF.
  ' Offset specifies the relative position (from
  ' StartLocation) of the row for the programmer to start
  ' transferring data. A positive number indicates a
  ' forward relative position while a negative number
  ' indicates a backward relative position. Regardless
  ' of whether the rows to be read are before or after
  ' StartLocation, rows are always fetched going forward
  ' (this is why there is no ReadPriorRows parameter to
  ' the procedure).
  ' If you page down on the grid, for instance, the new
  ' top row of the grid will have an index greater than
  ' the StartLocation (Offset > 0). If you page up on
  ' the grid, the new index is less than that of
  ' StartLocation, so Offset < 0. If StartLocation is
  ' a vrtBookmark to row N, the grid always asks for row
  ' data in the following order:
  '   (N + Offset), (N + Offset + 1), (N + Offset + 2)...
  ' ApproximatePosition is a value you can set to indicate
  ' the ordinal position of (StartLocation + Offset).
  ' Setting this variable will enhance the ability of the
  ' grid to display its vertical scroll bar accurately.
  ' If the exact ordinal position of the new location is
  ' not known, you can set it to a reasonable,
  ' approximate value, or just ignore this parameter.

  On Error GoTo trata
  '
  Dim intColIndex      As Integer
  Dim intJ             As Integer
  Dim intRowsFetched   As Integer
  Dim intI             As Long
  Dim lngNewPosition   As Long
  Dim vrtBookmark      As Variant
  '
  intRowsFetched = 0
  For intI = 0 To RowBuf.RowCount - 1
    ' Get the vrtBookmark of the next available row
    vrtBookmark = GetRelativeBookmarkGeral(StartLocation, _
               Offset + intI, HIST_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, HIST_COLUNASMATRIZ, HIST_LINHASMATRIZ, HIST_Matriz)
    Next intJ

    ' Set the vrtBookmark for the row
    RowBuf.Bookmark(intI) = vrtBookmark

    ' Increment the count of fetched rows
    intRowsFetched = intRowsFetched + 1
  Next intI

' Tell the grid how many rows were fetched
  RowBuf.RowCount = intRowsFetched

' Set the approximate scroll bar position. Only
' nonnegative values of IndexFromBookmark() are valid.
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, HIST_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmPacoteInc.grdHistorico_UnboundReadDataEx]"
End Sub

Public Sub SERV_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGeral    As busElite.clsGeral
  '
  On Error GoTo trata

  Set objGeral = New busElite.clsGeral
  '
  strSql = "SELECT PACOTESERVICO.PKID, SERVICO.DATAHORA, " & _
           " AGENCIA.NOME + ' (' + dbo.formataCNPJ(AGENCIACNPJ.CNPJ) + ')', " & _
           " VEICULO.PLACA, SERVICO.SOLICITANTE, SERVICO.STATUS " & _
           "FROM PACOTESERVICO " & _
           " INNER JOIN SERVICO ON SERVICO.PKID = PACOTESERVICO.SERVICOID " & _
           " INNER JOIN VEICULO ON VEICULO.PKID = PACOTESERVICO.VEICULOID " & _
           " INNER JOIN AGENCIACNPJ ON AGENCIACNPJ.PKID = SERVICO.AGENCIACNPJID " & _
           " INNER JOIN AGENCIA ON AGENCIA.PKID = AGENCIACNPJ.AGENCIAID " & _
           " WHERE PACOTESERVICO.PACOTEID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
           " AND PACOTESERVICO.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
           " ORDER BY SERVICO.DATAHORA DESC, AGENCIA.NOME, AGENCIACNPJ.CNPJ"

  '
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    SERV_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim SERV_Matriz(0 To SERV_COLUNASMATRIZ - 1, 0 To SERV_LINHASMATRIZ - 1)
  Else
    ReDim SERV_Matriz(0 To SERV_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To SERV_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To SERV_COLUNASMATRIZ - 1    'varre as colunas
          SERV_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set objGeral = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cboMotorista_LostFocus()
  Pintar_Controle cboMotorista, tpCorContr_Normal
End Sub

Public Sub HIST_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGeral    As busElite.clsGeral
  '
  On Error GoTo trata

  Set objGeral = New busElite.clsGeral
  '
  strSql = "SELECT HISTORICOSERVICO.PKID, HISTORICOSERVICO.DATAHORA, " & _
           " HISTORICOSERVICO.OBSERVACAO, " & _
           " SERVICO.PASSAGEIRO " & _
           "FROM HISTORICOSERVICO " & _
           " INNER JOIN PACOTESERVICO ON PACOTESERVICO.PKID = HISTORICOSERVICO.PACOTESERVICOID " & _
           " INNER JOIN SERVICO ON SERVICO.PKID = PACOTESERVICO.SERVICOID " & _
           " WHERE PACOTESERVICO.PACOTEID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
           " AND PACOTESERVICO.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
           " ORDER BY HISTORICOSERVICO.DATAHORA DESC"

  '
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    HIST_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim HIST_Matriz(0 To HIST_COLUNASMATRIZ - 1, 0 To HIST_LINHASMATRIZ - 1)
  Else
    ReDim HIST_Matriz(0 To HIST_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To HIST_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To HIST_COLUNASMATRIZ - 1    'varre as colunas
          HIST_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set objGeral = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub





