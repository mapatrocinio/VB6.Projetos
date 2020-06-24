VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmFichaClienteInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Cliente"
   ClientHeight    =   6090
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10290
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6090
   ScaleWidth      =   10290
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6090
      Left            =   8430
      ScaleHeight     =   6090
      ScaleWidth      =   1860
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2085
         Left            =   90
         ScaleHeight     =   2025
         ScaleWidth      =   1605
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3930
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   150
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   150
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5895
      Left            =   150
      TabIndex        =   25
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   10398
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do Cliente"
      TabPicture(0)   =   "userFichaClienteInc.frx":0000
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
         Height          =   5385
         Left            =   150
         TabIndex        =   26
         Top             =   420
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   5115
            Index           =   2
            Left            =   120
            ScaleHeight     =   5115
            ScaleWidth      =   7575
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   180
            Width           =   7575
            Begin VB.TextBox txtObservacao 
               Height          =   585
               Left            =   1320
               MaxLength       =   255
               MultiLine       =   -1  'True
               TabIndex        =   20
               Text            =   "userFichaClienteInc.frx":001C
               Top             =   4170
               Width           =   6075
            End
            Begin VB.TextBox txtEmail 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   19
               Text            =   "txtEmail"
               Top             =   3840
               Width           =   6075
            End
            Begin VB.OptionButton optSexo 
               Caption         =   "Feminino"
               Height          =   195
               Index           =   1
               Left            =   5880
               TabIndex        =   6
               Top             =   1065
               Width           =   1395
            End
            Begin VB.TextBox txtNumeroDoc 
               Height          =   285
               Left            =   5220
               MaxLength       =   50
               TabIndex        =   1
               Text            =   "txtNumeroDoc"
               Top             =   90
               Width           =   2175
            End
            Begin VB.ComboBox cboTipoDoc 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   0
               Top             =   90
               Width           =   2025
            End
            Begin VB.TextBox txtNome 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   3
               Text            =   "txtNome"
               Top             =   765
               Width           =   6075
            End
            Begin VB.TextBox txtEndereco 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   7
               Text            =   "txtEndereco"
               Top             =   1365
               Width           =   6075
            End
            Begin VB.TextBox txtNumero 
               Height          =   285
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   8
               Text            =   "txtNumero"
               Top             =   1725
               Width           =   1995
            End
            Begin VB.TextBox txtComplemento 
               Height          =   285
               Left            =   4680
               MaxLength       =   30
               TabIndex        =   9
               Text            =   "txtComplemento"
               Top             =   1725
               Width           =   2715
            End
            Begin VB.TextBox txtBairro 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   10
               Text            =   "txtBairro"
               Top             =   2025
               Width           =   6075
            End
            Begin VB.TextBox txtCidade 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   11
               Text            =   "txtCidade"
               Top             =   2325
               Width           =   6075
            End
            Begin VB.TextBox txtEstado 
               Height          =   285
               Left            =   1320
               MaxLength       =   2
               TabIndex        =   12
               Text            =   "txtEstado"
               Top             =   2625
               Width           =   495
            End
            Begin VB.TextBox txtCep 
               Height          =   285
               Left            =   5220
               MaxLength       =   20
               TabIndex        =   13
               Text            =   "txtCep"
               Top             =   2625
               Width           =   2175
            End
            Begin VB.TextBox txtPais 
               Height          =   285
               Left            =   1320
               MaxLength       =   50
               TabIndex        =   14
               Text            =   "txtPais"
               Top             =   2925
               Width           =   6075
            End
            Begin VB.TextBox txtTel1 
               Height          =   285
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   15
               Text            =   "txtTel1"
               Top             =   3225
               Width           =   2175
            End
            Begin VB.TextBox txtTel2 
               Height          =   285
               Left            =   5220
               MaxLength       =   20
               TabIndex        =   16
               Text            =   "txtTel2"
               Top             =   3225
               Width           =   2175
            End
            Begin VB.TextBox txtTel3 
               Height          =   285
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   17
               Text            =   "txtTel3"
               Top             =   3525
               Width           =   2175
            End
            Begin VB.TextBox txtSobrenome 
               Height          =   285
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   2
               Text            =   "txtSobrenome"
               Top             =   435
               Width           =   2610
            End
            Begin MSMask.MaskEdBox mskNascimento 
               Height          =   255
               Left            =   1320
               TabIndex        =   4
               Top             =   1065
               Width           =   1335
               _ExtentX        =   2355
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               AutoTab         =   -1  'True
               MaxLength       =   10
               Mask            =   "##/##/####"
               PromptChar      =   "_"
            End
            Begin VB.OptionButton optSexo 
               Caption         =   "Masculino"
               Height          =   195
               Index           =   0
               Left            =   4620
               TabIndex        =   5
               Top             =   1065
               Value           =   -1  'True
               Width           =   1395
            End
            Begin MSMask.MaskEdBox mskPercDesconto 
               Height          =   255
               Left            =   5220
               TabIndex        =   18
               Top             =   3540
               Width           =   1575
               _ExtentX        =   2778
               _ExtentY        =   450
               _Version        =   393216
               Format          =   "#,##0.00##;($#,##0.00##)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Observação"
               Height          =   195
               Index           =   18
               Left            =   90
               TabIndex        =   47
               Top             =   4170
               Width           =   975
            End
            Begin VB.Label Label5 
               Caption         =   "E-mail"
               Height          =   195
               Index           =   17
               Left            =   90
               TabIndex        =   46
               Top             =   3840
               Width           =   975
            End
            Begin VB.Label Label14 
               Caption         =   "% Desc."
               Height          =   255
               Left            =   4140
               TabIndex        =   45
               Top             =   3510
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Tipo Doc."
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   44
               Top             =   90
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Número"
               Height          =   195
               Index           =   1
               Left            =   4080
               TabIndex        =   43
               Top             =   90
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Nome"
               Height          =   195
               Index           =   2
               Left            =   60
               TabIndex        =   42
               Top             =   765
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Nascimento"
               Height          =   195
               Index           =   3
               Left            =   60
               TabIndex        =   41
               Top             =   1065
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Sexo"
               Height          =   195
               Index           =   4
               Left            =   4080
               TabIndex        =   40
               Top             =   1395
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Endereço"
               Height          =   195
               Index           =   5
               Left            =   60
               TabIndex        =   39
               Top             =   1365
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Número"
               Height          =   195
               Index           =   6
               Left            =   60
               TabIndex        =   38
               Top             =   1725
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Complemento"
               Height          =   195
               Index           =   7
               Left            =   3420
               TabIndex        =   37
               Top             =   1725
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Bairro"
               Height          =   195
               Index           =   8
               Left            =   60
               TabIndex        =   36
               Top             =   2025
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cidade"
               Height          =   195
               Index           =   9
               Left            =   60
               TabIndex        =   35
               Top             =   2325
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Estado"
               Height          =   195
               Index           =   10
               Left            =   60
               TabIndex        =   34
               Top             =   2625
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Cep"
               Height          =   195
               Index           =   11
               Left            =   4140
               TabIndex        =   33
               Top             =   2625
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "País"
               Height          =   195
               Index           =   12
               Left            =   60
               TabIndex        =   32
               Top             =   2925
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Tel 1"
               Height          =   195
               Index           =   13
               Left            =   60
               TabIndex        =   31
               Top             =   3225
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Tel 2"
               Height          =   195
               Index           =   14
               Left            =   4140
               TabIndex        =   30
               Top             =   3225
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Tel 3"
               Height          =   195
               Index           =   15
               Left            =   60
               TabIndex        =   29
               Top             =   3525
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Sobrenome"
               Height          =   195
               Index           =   16
               Left            =   45
               TabIndex        =   28
               Top             =   435
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmFichaClienteInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Public intOrigem                As Integer
'intOrigem = 0 cadastro
'intOrigem = 1 pedido

Public lngFICHACLIENTEID        As Long

Private blnPrimeiraVez          As Boolean 'Propósito: Preencher lista no combo

Public Function RetornaApartamentoId(ByVal strApartamento As String) As String
  On Error GoTo trata
  Dim strSql      As String
  Dim objRs       As ADODB.Recordset
  Dim strRetorno  As Long
  Dim objGeral    As busSisMetal.clsGeral
  '
  Set objGeral = New busSisMetal.clsGeral
  strSql = "SELECT APARTAMENTO.PKID FROM APARTAMENTO WHERE NUMERO = '" & strApartamento & "'  AND APARTAMENTO.INTERDITADO = False AND APARTAMENTO.EXCLUIDO = False"
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then
    If IsNumeric(objRs.Fields("PKID").Value) Then
      strRetorno = objRs.Fields("PKID").Value
    Else
      strRetorno = 0
    End If
  Else
    strRetorno = 0
  End If
  '
  objRs.Close
  Set objRs = Nothing
  '
  RetornaApartamentoId = strRetorno
  '
  Exit Function
trata:
  TratarErro Err.Number, Err.Description, "[frmUserFichaClienteInc.RetornaApartamentoId]"
End Function



Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Cliente
  LimparCampoCombo cboTipoDoc
  LimparCampoTexto txtNumeroDoc
  LimparCampoTexto txtSobrenome
  LimparCampoTexto txtNome
  LimparCampoMask mskNascimento
  optSexo(0).Value = False
  optSexo(1).Value = False
  LimparCampoTexto txtEndereco
  LimparCampoTexto txtNumero
  LimparCampoTexto txtComplemento
  LimparCampoTexto txtBairro
  LimparCampoTexto txtCidade
  LimparCampoTexto txtEstado
  LimparCampoTexto txtCep
  LimparCampoTexto txtPais
  LimparCampoTexto txtTel1
  LimparCampoTexto txtTel2
  LimparCampoTexto txtTel3
  LimparCampoMask mskPercDesconto
  LimparCampoTexto txtEmail
  LimparCampoTexto txtObservacao
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserFichaClienteInc.LimparCampos]", _
            Err.Description
End Sub

Private Sub cboTipoDoc_Click()
  On Error GoTo trata
  If Len(Trim(txtNome.Text)) = 0 Or Len(Trim(txtSobrenome.Text)) = 0 Or Len(Trim(mskNascimento.ClipText)) = 0 Then
    If Len(Trim(cboTipoDoc.Text)) <> 0 _
        And Len(Trim(txtNumeroDoc.Text)) <> 0 Then
      If Status = tpStatus_Incluir Then
        VerificaPessoaJaCadastrada
      End If
    End If
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source & ".[frmUserFichaClienteInc.cboTipoDoc_Click]"
End Sub
Private Sub cboTipoDoc_LostFocus()
  Pintar_Controle cboTipoDoc, tpCorContr_Normal
End Sub

Private Sub cmdAlterar_Click()
  On Error GoTo trata
'''  If Len(Trim(grdResponsabilidade.Columns("PKID").Value & "")) = 0 Then
'''    MsgBox "Selecione uma Responsabilidade.", vbExclamation, TITULOSISTEMA
'''    SetarFoco grdResponsabilidade
'''    Exit Sub
'''  End If
'''  frmUserTabRespInc.Status = tpStatus_Alterar
'''  frmUserTabRespInc.intQuemChamou = 2
'''  frmUserTabRespInc.lngTABRESPID = grdResponsabilidade.Columns("PKID").Value
'''  frmUserTabRespInc.lngRESPLOCID = lngLOCACAOID
'''  frmUserTabRespInc.Show vbModal
'''  If frmUserTabRespInc.bRetorno Then
'''    RESP_COLUNASMATRIZ = grdResponsabilidade.Columns.Count
'''    RESP_LINHASMATRIZ = 0
'''    RESP_MontaMatriz
'''    grdResponsabilidade.Bookmark = Null
'''    grdResponsabilidade.ReBind
'''    grdResponsabilidade.ApproxCount = RESP_LINHASMATRIZ
'''    '
'''  End If
'''  SetarFoco grdResponsabilidade
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdCancelar_Click()
  blnFechar = True
  blnRetorno = True
  If intOrigem = 1 Then
    'LimparCampoTexto frmPedidoItemVendaInc.txtCodClieFornFim
    'LimparCampoTexto frmPedidoItemVendaInc.txtNomeClieFornFim
  End If
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdExcluir_Click()
'''  Dim objTabResp      As busSisMetal.clsTabResp
'''  '
'''  On Error GoTo trata
'''  If Len(Trim(grdResponsabilidade.Columns("PKID").Value & "")) = 0 Then
'''    MsgBox "Selecione uma Responsabilidade.", vbExclamation, TITULOSISTEMA
'''    SetarFoco grdResponsabilidade
'''    Exit Sub
'''  End If
'''  '
'''  Set objTabResp = New busSisMetal.clsTabResp
'''  '
'''  If MsgBox("Confirma exclusão da Responsabilidade " & grdResponsabilidade.Columns("Responsabilidade").Value & " ?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''    SetarFoco grdResponsabilidade
'''    Exit Sub
'''  End If
'''  '
'''  'OK
'''  objTabResp.ExcluirTabResp 2, _
'''                            CLng(grdResponsabilidade.Columns("PKID").Value)
'''
'''  RESP_COLUNASMATRIZ = grdResponsabilidade.Columns.Count
'''  RESP_LINHASMATRIZ = 0
'''  RESP_MontaMatriz
'''  grdResponsabilidade.Bookmark = Null
'''  grdResponsabilidade.ReBind
'''  grdResponsabilidade.ApproxCount = RESP_LINHASMATRIZ
'''  '
'''  Set objTabResp = Nothing
'''  SetarFoco grdResponsabilidade
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdOk_Click()
  Dim objFichaCliente           As busSisMetal.clsFichaCliente
  Dim objGeral                  As busSisMetal.clsGeral
  Dim lngTIPODOCUMENTOID        As Long
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim strSexo                   As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCliente Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisMetal.clsGeral
  Set objFichaCliente = New busSisMetal.clsFichaCliente
  '
  lngTIPODOCUMENTOID = 0
  'Tipo Documento
  strSql = "SELECT PKID FROM TIPODOCUMENTO WHERE DESCRICAO = " & Formata_Dados(cboTipoDoc.Text, tpDados_Texto, tpNulo_Aceita)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngTIPODOCUMENTOID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  'Sexo
  If optSexo(0).Value Then
    strSexo = "M"
  Else
    strSexo = "F"
  End If

  'Validar se ficha cliente já cadastrado
  'Tipo/documento e número
  Set objRs = objFichaCliente.SelecionarFichaCliente(lngTIPODOCUMENTOID, _
                                                     txtNumeroDoc.Text, _
                                                     lngFICHACLIENTEID)
  If Not objRs.EOF Then
    objRs.Close
    Set objRs = Nothing
    Set objFichaCliente = Nothing
    TratarErroPrevisto "Documento/Número já cadastrado para outro cliente", "cmdOK_Click"
    Pintar_Controle txtNumeroDoc, tpCorContr_Erro
    Pintar_Controle cboTipoDoc, tpCorContr_Erro
    SetarFoco cboTipoDoc
    cmdOk.Enabled = True
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  'Nome/Sobrenome/data nascimento e número
  Set objRs = objFichaCliente.SelecionarFichaClientePeloNome(txtNome.Text, _
                                                            txtSobrenome.Text, _
                                                            mskNascimento.Text, _
                                                            lngFICHACLIENTEID)
  If Not objRs.EOF Then
    objRs.Close
    Set objRs = Nothing
    Set objFichaCliente = Nothing
    TratarErroPrevisto "Nome/Sobrenome/Data de nascimento já cadastrados para outro cliente", "cmdOK_Click"
    Pintar_Controle txtNome, tpCorContr_Erro
    Pintar_Controle txtSobrenome, tpCorContr_Erro
    Pintar_Controle mskNascimento, tpCorContr_Erro
    SetarFoco txtNome
    cmdOk.Enabled = True
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  'Inserir Ficha Cliente
  lngFICHACLIENTEID = objFichaCliente.CadastrarFichaCliente(lngFICHACLIENTEID, _
                                                            lngTIPODOCUMENTOID, _
                                                            txtSobrenome.Text, _
                                                            txtNome.Text, _
                                                            txtEndereco.Text, _
                                                            txtNumero.Text, _
                                                            txtComplemento.Text, _
                                                            txtBairro.Text, _
                                                            txtCidade.Text, _
                                                            txtEstado.Text, _
                                                            txtCep.Text, _
                                                            txtPais.Text, _
                                                            txtTel1.Text, _
                                                            txtTel2.Text, _
                                                            txtTel3.Text, _
                                                            mskNascimento.Text, _
                                                            strSexo, _
                                                            txtNumeroDoc, _
                                                            mskPercDesconto.Text, _
                                                            txtEmail.Text, _
                                                            txtObservacao.Text)
  Set objFichaCliente = Nothing
  If intOrigem = 1 Then
    frmPedidoItemVendaInc.txtCodClieFornFim.Text = lngFICHACLIENTEID
    frmPedidoItemVendaInc.txtNomeClieFornFim = txtNome.Text
  End If
  
  blnRetorno = True
  blnFechar = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  cmdOk.Enabled = True
End Sub


'Propósito: Retornar o PKID do Grupo da Locação
Public Function RetornarGrupoId(ByVal lngLOCACAOID As Long) As Long
  Dim strSql      As String
  Dim objRs       As ADODB.Recordset
  Dim lngRetorno  As Long
  Dim objGeral    As busSisMetal.clsGeral
  '
  Set objGeral = New busSisMetal.clsGeral
  On Error GoTo trata
  strSql = "SELECT GRUPO.PKID FROM (GRUPO INNER JOIN APARTAMENTO ON GRUPO.PKID = APARTAMENTO.GRUPOID) INNER JOIN LOCACAO ON APARTAMENTO.PKID = LOCACAO.APARTAMENTOID WHERE Locacao.pkid = " & lngLOCACAOID
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If objRs.EOF Then
    lngRetorno = 0
  ElseIf Not IsNumeric(objRs.Fields("PKID").Value) Then
    lngRetorno = 0
  Else
    lngRetorno = objRs.Fields("PKID").Value
  End If
  '
  objRs.Close
  Set objRs = Nothing
  '
  Set objGeral = Nothing
  RetornarGrupoId = lngRetorno
  Exit Function
trata:
  TratarErro Err.Number, Err.Description, "[frmUserFichaClienteInc.RetornarGrupoId]"
End Function

Private Sub mskNascimento_GotFocus()
  Seleciona_Conteudo_Controle mskNascimento
End Sub
Private Sub mskNascimento_LostFocus()
  On Error GoTo trata
  Pintar_Controle mskNascimento, tpCorContr_Normal
  If Len(Trim(cboTipoDoc.Text)) = 0 Or Len(Trim(txtNumeroDoc.Text)) = 0 Then
    If Len(Trim(txtNome.Text)) <> 0 And _
        Len(Trim(txtSobrenome.Text)) <> 0 And _
        Len(Trim(mskNascimento.ClipText)) <> 0 Then
      If Status = tpStatus_Incluir Then
        VerificaPessoaJaCadastradaNome
      End If
    End If
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source & ".[frmUserFichaClienteInc.mskNascimento_LostFocus]"
End Sub

Private Sub mskPercDesconto_GotFocus()
  Seleciona_Conteudo_Controle mskPercDesconto
End Sub

Private Sub mskPercDesconto_LostFocus()
  Pintar_Controle mskPercDesconto, tpCorContr_Normal
End Sub


Private Sub txtEmail_GotFocus()
  Seleciona_Conteudo_Controle txtEmail
End Sub
Private Sub txtEmail_LostFocus()
  Pintar_Controle txtEmail, tpCorContr_Normal
End Sub

Private Sub txtEstado_GotFocus()
  Seleciona_Conteudo_Controle txtEstado
End Sub
Private Sub txtEstado_LostFocus()
  Pintar_Controle txtEstado, tpCorContr_Normal
End Sub

Private Function ValidaCliente() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim strMsgCPF             As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCliente = False
  If Not Valida_String(cboTipoDoc, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o Tipo de Documento" & vbCrLf
  End If
  If Not Valida_String(txtNumeroDoc, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número do documento" & vbCrLf
  End If
  If UCase(cboTipoDoc.Text) = "CPF" Then
    If Not TestaCPF(txtNumeroDoc.Text, strMsgCPF) Then
      'Não informou o cpf
      strMsg = strMsg & strMsgCPF & vbCrLf
      Pintar_Controle txtNumeroDoc, tpCorContr_Erro
      SetarFoco txtNumeroDoc
      blnSetarFocoControle = False
    End If
  End If
  If Not Valida_String(txtSobrenome, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o sobrenome do cliente" & vbCrLf
  End If
  If Not Valida_String(txtNome, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o nome do cliente" & vbCrLf
  End If
  If optSexo(0).Value = False And optSexo(1).Value = False Then
    strMsg = strMsg & "Selecionar o sexo" & vbCrLf
    'SetarFoco optSexo(0)
    'blnSetarFocoControle = False
  End If
  If Not Valida_Data(mskNascimento, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de nascimento válida" & vbCrLf
  End If
  '
  If Not Valida_Moeda(mskPercDesconto, TpnaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o percentual de desconto válido" & vbCrLf
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserFichaClienteInc.ValidaCliente]"
    ValidaCliente = False
  Else
    ValidaCliente = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserFichaClienteInc.ValidaCliente]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  
  If blnPrimeiraVez Then
    SetarFoco cboTipoDoc
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserFichaClienteInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objFichaClie            As busSisMetal.clsFichaCliente
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 6570
  Me.Width = 10380
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  'Tipo Documento
  strSql = "Select DESCRICAO from TIPODOCUMENTO ORDER BY DESCRICAO"
  PreencheCombo cboTipoDoc, strSql, False, True
  'tabDetalhes_Click 0
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objFichaClie = New busSisMetal.clsFichaCliente
    Set objRs = objFichaClie.SelecionarFichaClientePeloPkid(lngFICHACLIENTEID)
    '
    If Not objRs.EOF Then
      txtNumeroDoc.Text = objRs.Fields("NRODOCUMENTO").Value & ""
      If objRs.Fields("DESC_TIPODOCUMENTO").Value & "" <> "" Then
        cboTipoDoc.Text = objRs.Fields("DESC_TIPODOCUMENTO").Value & ""
      End If
      txtSobrenome.Text = objRs.Fields("SOBRENOME").Value & ""
      txtNome.Text = objRs.Fields("NOME").Value & ""
      txtEndereco.Text = objRs.Fields("ENDERECO").Value & ""
      txtNumero.Text = objRs.Fields("NUMERO").Value & ""
      txtComplemento.Text = objRs.Fields("COMPLEMENTO").Value & ""
      txtBairro.Text = objRs.Fields("BAIRRO").Value & ""
      txtCidade.Text = objRs.Fields("CIDADE").Value & ""
      txtEstado.Text = objRs.Fields("ESTADO").Value & ""
      txtCep.Text = objRs.Fields("CEP").Value & ""
      txtPais.Text = objRs.Fields("PAIS").Value & ""
      txtTel1.Text = objRs.Fields("TEL1").Value & ""
      txtTel2.Text = objRs.Fields("TEL2").Value & ""
      txtTel3.Text = objRs.Fields("TEL3").Value & ""
      INCLUIR_VALOR_NO_MASK mskNascimento, objRs.Fields("DTNASCIMENTO").Value & "", TpMaskData
      If objRs.Fields("SEXO").Value & "" = "M" Then
        optSexo(0).Value = True
      Else
        optSexo(1).Value = True
      End If
      INCLUIR_VALOR_NO_MASK mskPercDesconto, objRs.Fields("PERCDESC").Value & "", TpMaskMoeda
      txtEmail.Text = objRs.Fields("EMAIL").Value & ""
      txtObservacao.Text = objRs.Fields("OBSERVACAO").Value & ""
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objFichaClie = Nothing
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

Private Sub txtNome_GotFocus()
  Seleciona_Conteudo_Controle txtNome
End Sub
Private Sub txtNome_LostFocus()
  On Error GoTo trata
  Pintar_Controle txtNome, tpCorContr_Normal
  If Len(Trim(cboTipoDoc.Text)) = 0 Or Len(Trim(txtNumeroDoc.Text)) = 0 Then
    If Len(Trim(txtNome.Text)) <> 0 And _
        Len(Trim(txtSobrenome.Text)) <> 0 And _
        Len(Trim(mskNascimento.ClipText)) <> 0 Then
      If Status = tpStatus_Incluir Then
        VerificaPessoaJaCadastradaNome
      End If
    End If
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source & ".[frmUserFichaClienteInc.txtSobrenome_LostFocus]"
  
End Sub


Private Sub txtNumeroDoc_GotFocus()
  Seleciona_Conteudo_Controle txtNumeroDoc
End Sub
Public Sub VerificaPessoaJaCadastradaNome()
  On Error GoTo trata
  Dim objFichaCliente       As busSisMetal.clsFichaCliente
  Dim objGeral              As busSisMetal.clsGeral
  Dim objRs                 As ADODB.Recordset
  Dim strSql                As String
  '
  Set objGeral = New busSisMetal.clsGeral
  Set objFichaCliente = New busSisMetal.clsFichaCliente
  'validação básica de data
  If Not Valida_Data(mskNascimento, TpObrigatorio, False) Then
    Exit Sub
  End If
  Set objRs = objFichaCliente.SelecionarFichaClientePeloNome(txtNome.Text, _
                                                            txtSobrenome.Text, _
                                                            mskNascimento.Text, _
                                                            lngFICHACLIENTEID)
  '
  If Not objRs.EOF Then
    'MsgBox "Cliente já cadastrado para este Documento. Os dados serão carregados na tela.", vbExclamation, TITULOSISTEMA
    txtSobrenome.Text = objRs.Fields("SOBRENOME").Value & ""
    txtNome.Text = objRs.Fields("NOME").Value & ""
    txtNumeroDoc.Text = objRs.Fields("NRODOCUMENTO").Value & ""
    If objRs.Fields("DESC_TIPODOCUMENTO").Value & "" <> "" Then
      cboTipoDoc.Text = objRs.Fields("DESC_TIPODOCUMENTO").Value & ""
    End If
    
    txtEndereco.Text = objRs.Fields("ENDERECO").Value & ""
    txtNumero.Text = objRs.Fields("NUMERO").Value & ""
    txtComplemento.Text = objRs.Fields("COMPLEMENTO").Value & ""
    txtBairro.Text = objRs.Fields("BAIRRO").Value & ""
    txtCidade.Text = objRs.Fields("CIDADE").Value & ""
    txtEstado.Text = objRs.Fields("ESTADO").Value & ""
    txtCep.Text = objRs.Fields("CEP").Value & ""
    txtPais.Text = objRs.Fields("PAIS").Value & ""
    txtTel1.Text = objRs.Fields("TEL1").Value & ""
    txtTel2.Text = objRs.Fields("TEL2").Value & ""
    txtTel3.Text = objRs.Fields("TEL3").Value & ""
    INCLUIR_VALOR_NO_MASK mskNascimento, objRs.Fields("DTNASCIMENTO").Value & "", TpMaskData
    If objRs.Fields("SEXO").Value & "" = "M" Then
      optSexo(0).Value = True
    Else
      optSexo(1).Value = True
    End If
    txtEmail.Text = objRs.Fields("EMAIL").Value & ""
    txtObservacao.Text = objRs.Fields("OBSERVACAO").Value & ""
    'Altera para alteração
    Status = tpStatus_Alterar
    lngFICHACLIENTEID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  
  Set objGeral = Nothing
  Set objFichaCliente = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserFichaClienteInc.VerificaPessoaJaCadastradaNome]", _
            Err.Description
End Sub

Public Sub VerificaPessoaJaCadastrada()
  On Error GoTo trata
  Dim objFichaCliente       As busSisMetal.clsFichaCliente
  Dim objGeral              As busSisMetal.clsGeral
  Dim objRs                 As ADODB.Recordset
  Dim lngTIPODOCUMENTOID    As Long
  Dim strSql                As String
  '
  Set objGeral = New busSisMetal.clsGeral
  Set objFichaCliente = New busSisMetal.clsFichaCliente
  'Selecionar Ficha Cliente Id
  lngTIPODOCUMENTOID = 0
  strSql = "SELECT * FROM TIPODOCUMENTO WHERE DESCRICAO = " & Formata_Dados(cboTipoDoc.Text, tpDados_Texto, tpNulo_Aceita)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    If IsNumeric(objRs.Fields("PKID").Value) Then
      lngTIPODOCUMENTOID = objRs.Fields("PKID").Value
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set objRs = objFichaCliente.SelecionarFichaCliente(lngTIPODOCUMENTOID, _
                                                     txtNumeroDoc.Text, _
                                                     lngFICHACLIENTEID)
  '
  If Not objRs.EOF Then
    'MsgBox "Cliente já cadastrado para este Documento. Os dados serão carregados na tela.", vbExclamation, TITULOSISTEMA
    txtSobrenome.Text = objRs.Fields("SOBRENOME").Value & ""
    txtNome.Text = objRs.Fields("NOME").Value & ""
    txtEndereco.Text = objRs.Fields("ENDERECO").Value & ""
    txtNumero.Text = objRs.Fields("NUMERO").Value & ""
    txtComplemento.Text = objRs.Fields("COMPLEMENTO").Value & ""
    txtBairro.Text = objRs.Fields("BAIRRO").Value & ""
    txtCidade.Text = objRs.Fields("CIDADE").Value & ""
    txtEstado.Text = objRs.Fields("ESTADO").Value & ""
    txtCep.Text = objRs.Fields("CEP").Value & ""
    txtPais.Text = objRs.Fields("PAIS").Value & ""
    txtTel1.Text = objRs.Fields("TEL1").Value & ""
    txtTel2.Text = objRs.Fields("TEL2").Value & ""
    txtTel3.Text = objRs.Fields("TEL3").Value & ""
    INCLUIR_VALOR_NO_MASK mskNascimento, objRs.Fields("DTNASCIMENTO").Value & "", TpMaskData
    If objRs.Fields("SEXO").Value & "" = "M" Then
      optSexo(0).Value = True
    Else
      optSexo(1).Value = True
    End If
    txtEmail.Text = objRs.Fields("EMAIL").Value & ""
    txtObservacao.Text = objRs.Fields("OBSERVACAO").Value & ""
    'Altera para alteração
    Status = tpStatus_Alterar
    lngFICHACLIENTEID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  
  Set objGeral = Nothing
  Set objFichaCliente = Nothing
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserFichaClienteInc.VerificaPessoaJaCadastrada]", _
            Err.Description
End Sub
Private Sub txtNumeroDoc_LostFocus()
  On Error GoTo trata
  Pintar_Controle txtNumeroDoc, tpCorContr_Normal
  Dim strMsg      As String
  If Len(Trim(txtNome.Text)) = 0 Or Len(Trim(txtSobrenome.Text)) = 0 Or Len(Trim(mskNascimento.ClipText)) = 0 Then
    If Len(Trim(cboTipoDoc.Text)) <> 0 _
        And Len(Trim(txtNumeroDoc.Text)) <> 0 Then
      If Status = tpStatus_Incluir Or Status = tpStatus_Alterar Then
        If cboTipoDoc.Text = "CPF" Then
          If Not TestaCPF(txtNumeroDoc.Text, strMsg) Then
            MsgBox strMsg, vbExclamation, TITULOSISTEMA
            Exit Sub
          End If
        End If
        '
        tabDetalhes.Tab = 0
      End If
      If Status = tpStatus_Incluir Then
        VerificaPessoaJaCadastrada
      End If
    End If
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source & ".[frmUserFichaClienteInc.txtNumeroDoc_LostFocus]"
End Sub
Private Sub txtObservacao_GotFocus()
  Seleciona_Conteudo_Controle txtObservacao
End Sub
Private Sub txtObservacao_LostFocus()
  Pintar_Controle txtObservacao, tpCorContr_Normal
End Sub

Private Sub txtSobrenome_GotFocus()
  Seleciona_Conteudo_Controle txtSobrenome
End Sub
Private Sub txtSobrenome_LostFocus()
  On Error GoTo trata
  Pintar_Controle txtSobrenome, tpCorContr_Normal
  If Len(Trim(cboTipoDoc.Text)) = 0 Or Len(Trim(txtNumeroDoc.Text)) = 0 Then
    If Len(Trim(txtNome.Text)) <> 0 And _
        Len(Trim(txtSobrenome.Text)) <> 0 And _
        Len(Trim(mskNascimento.ClipText)) <> 0 Then
      If Status = tpStatus_Incluir Then
        VerificaPessoaJaCadastradaNome
      End If
    End If
  End If
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source & ".[frmUserFichaClienteInc.txtSobrenome_LostFocus]"
End Sub

