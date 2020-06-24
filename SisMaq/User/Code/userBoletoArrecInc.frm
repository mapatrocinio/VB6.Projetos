VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserBoletoArrecInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de boleto para arrecadador"
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10470
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4410
   ScaleWidth      =   10470
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   4410
      Left            =   8610
      ScaleHeight     =   4410
      ScaleWidth      =   1860
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2025
         Left            =   90
         ScaleHeight     =   1965
         ScaleWidth      =   1605
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1665
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   960
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4155
      Left            =   120
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   120
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   7329
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userBoletoArrecInc.frx":0000
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
         Height          =   3705
         Left            =   120
         TabIndex        =   21
         Top             =   330
         Width           =   8115
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   3435
            Index           =   0
            Left            =   120
            ScaleHeight     =   3435
            ScaleWidth      =   7875
            TabIndex        =   22
            TabStop         =   0   'False
            Top             =   150
            Width           =   7875
            Begin VB.TextBox txtDataDevol 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1320
               TabIndex        =   3
               Text            =   "txtDataDevol"
               Top             =   1020
               Width           =   1815
            End
            Begin VB.TextBox txtTurnoDevol 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1320
               TabIndex        =   2
               Text            =   "txtTurnoDevol"
               Top             =   690
               Width           =   6405
            End
            Begin VB.ComboBox cboArrecadador 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   4
               Top             =   1350
               Width           =   6435
            End
            Begin VB.TextBox txtData 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1320
               TabIndex        =   1
               Text            =   "txtData"
               Top             =   360
               Width           =   1815
            End
            Begin VB.TextBox txtTurno 
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1320
               TabIndex        =   0
               Text            =   "txtTurno"
               Top             =   30
               Width           =   6405
            End
            Begin MSMask.MaskEdBox mskInicio 
               Height          =   285
               Left            =   1320
               TabIndex        =   5
               Top             =   1710
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskTermino 
               Height          =   285
               Left            =   1320
               TabIndex        =   6
               Top             =   2040
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskNumero1 
               Height          =   285
               Left            =   1320
               TabIndex        =   7
               Top             =   2580
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskNumero2 
               Height          =   285
               Left            =   2610
               TabIndex        =   8
               Top             =   2580
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskNumero3 
               Height          =   285
               Left            =   3900
               TabIndex        =   9
               Top             =   2580
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskNumero4 
               Height          =   285
               Left            =   5190
               TabIndex        =   10
               Top             =   2580
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskNumero5 
               Height          =   285
               Left            =   6480
               TabIndex        =   11
               Top             =   2580
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskNumero6 
               Height          =   285
               Left            =   1320
               TabIndex        =   12
               Top             =   2910
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskNumero7 
               Height          =   285
               Left            =   2610
               TabIndex        =   13
               Top             =   2910
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskNumero8 
               Height          =   285
               Left            =   3900
               TabIndex        =   14
               Top             =   2910
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskNumero9 
               Height          =   285
               Left            =   5190
               TabIndex        =   15
               Top             =   2910
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskNumero10 
               Height          =   285
               Left            =   6480
               TabIndex        =   16
               Top             =   2910
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   503
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0;($#,##0)"
               PromptChar      =   "_"
            End
            Begin VB.Line Line1 
               X1              =   60
               X2              =   7680
               Y1              =   2430
               Y2              =   2430
            End
            Begin VB.Label Label5 
               Caption         =   "Número"
               Height          =   225
               Index           =   6
               Left            =   60
               TabIndex        =   31
               Top             =   2550
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Data Devolução"
               Height          =   195
               Index           =   5
               Left            =   60
               TabIndex        =   30
               Top             =   1020
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Turno Devolução"
               Height          =   285
               Index           =   4
               Left            =   60
               TabIndex        =   29
               Top             =   690
               Width           =   1305
            End
            Begin VB.Label Label5 
               Caption         =   "Término"
               Height          =   225
               Index           =   3
               Left            =   60
               TabIndex        =   28
               Top             =   2010
               Width           =   1215
            End
            Begin VB.Label Label6 
               Caption         =   "Arrecadador"
               Height          =   255
               Index           =   1
               Left            =   60
               TabIndex        =   27
               Top             =   1350
               Width           =   1455
            End
            Begin VB.Label Label5 
               Caption         =   "Data"
               Height          =   195
               Index           =   2
               Left            =   60
               TabIndex        =   26
               Top             =   360
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Turno Entrada"
               Height          =   285
               Index           =   0
               Left            =   60
               TabIndex        =   25
               Top             =   30
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Início"
               Height          =   225
               Index           =   1
               Left            =   60
               TabIndex        =   24
               Top             =   1680
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserBoletoArrecInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngTURNOARRECEPESQ             As Long

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'BoletoArrec
  LimparCampoTexto txtTurno
  LimparCampoTexto txtData
  LimparCampoTexto txtTurnoDevol
  LimparCampoTexto txtDataDevol
  '
  LimparCampoCombo cboArrecadador
  LimparCampoMask mskInicio
  LimparCampoMask mskTermino
  '
  LimparCampoMask mskNumero1
  LimparCampoMask mskNumero2
  LimparCampoMask mskNumero3
  LimparCampoMask mskNumero4
  LimparCampoMask mskNumero5
  LimparCampoMask mskNumero6
  LimparCampoMask mskNumero7
  LimparCampoMask mskNumero8
  LimparCampoMask mskNumero9
  LimparCampoMask mskNumero10
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserBoletoArrecInc.LimparCampos]", _
            Err.Description
End Sub


Private Sub cboArrecadador_LostFocus()
  Pintar_Controle cboArrecadador, tpCorContr_Normal
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
  Dim objBoletoArrec            As busSisMaq.clsBoletoArrec
  Dim objCaixaArrec             As busSisMaq.clsCaixaArrec
  Dim objGeral                  As busSisMaq.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngTURNOID                As Long
  Dim lngARRECADADORID          As Long
  Dim lngCAIXAARRECID           As Long
  Dim strData                   As String
  Dim strMsg                    As String
  Dim intI                      As Integer
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objBoletoArrec = New busSisMaq.clsBoletoArrec
  lngTURNOID = RetornaCodTurnoCorrente
  Set objGeral = New busSisMaq.clsGeral
  'ARRECADADOR
  lngARRECADADORID = 0
  strSql = "SELECT PESSOA.PKID FROM PESSOA WHERE PESSOA.NOME = " & Formata_Dados(cboArrecadador.Text, tpDados_Texto)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    lngARRECADADORID = objRs.Fields("PKID").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  'Verifica se inclirá turno do arrecadador
  lngCAIXAARRECID = RetornaCodTurnoCorrenteArrec(lngARRECADADORID, lngTURNOARRECEPESQ)
  If lngCAIXAARRECID = -1 Then
    MsgBox "Há mais de um turno aberto para o arrecadador, entre em contato com o administrador"
    cmdOk.Enabled = True
    Exit Sub
    
  End If
  '
  Set objGeral = Nothing
  'Pede liberação do arrecadador
  frmUserLoginLibera.lngFUNCIONARIOID = lngARRECADADORID
  frmUserLoginLibera.Show vbModal
  If Len(Trim(gsNomeUsuLib)) = 0 Then
    strMsg = "É necessário confirmação do arrecadador para executar esta ação."
    TratarErroPrevisto strMsg, "cmdConfirmar_Click"
    cmdOk.Enabled = True
    SetarFoco cboArrecadador
    Exit Sub
  End If
  
  
  If Status = tpStatus_Alterar Then
'''    'Alterar BoletoArrec
'''    objBoletoArrec.AlterarBoletoArrec lngPKID, _
'''                                        mskInicio.ClipText
'''    blnRetorno = True
'''    blnFechar = True
'''    Unload Me
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir EntradaArrec
    strData = Format(Now, "DD/MM/YYYY hh:mm")
    'Verifica se existe caixa para o atendente
    If lngCAIXAARRECID = 0 Then
      'Não existe caixa arrec cadastrado, cadastra
      Set objCaixaArrec = New busSisMaq.clsCaixaArrec
      objCaixaArrec.InserirCaixaArrec lngCAIXAARRECID, _
                                      lngARRECADADORID, _
                                      RetornaCodTurnoCorrente
      Set objCaixaArrec = Nothing
    End If
    
    
    'Inserir BoletoArrec
    'Verifica Se entrará pela data de início/término
    If mskInicio.ClipText <> "" And mskTermino.ClipText <> "" Then
      For intI = mskInicio.ClipText To mskTermino.ClipText
        objBoletoArrec.InserirBoletoArrec lngTURNOID, _
                                          lngCAIXAARRECID, _
                                          intI & "", _
                                          strData, _
                                          "I"
      Next intI
    End If
    If mskNumero1.ClipText <> "" Then
      objBoletoArrec.InserirBoletoArrec lngTURNOID, _
                                        lngCAIXAARRECID, _
                                        mskNumero1.ClipText, _
                                        strData, _
                                        "I"
    End If
    If mskNumero2.ClipText <> "" Then
      objBoletoArrec.InserirBoletoArrec lngTURNOID, _
                                        lngCAIXAARRECID, _
                                        mskNumero2.ClipText, _
                                        strData, _
                                        "I"
    End If
    If mskNumero3.ClipText <> "" Then
      objBoletoArrec.InserirBoletoArrec lngTURNOID, _
                                        lngCAIXAARRECID, _
                                        mskNumero3.ClipText, _
                                        strData, _
                                        "I"
    End If
    If mskNumero4.ClipText <> "" Then
      objBoletoArrec.InserirBoletoArrec lngTURNOID, _
                                        lngCAIXAARRECID, _
                                        mskNumero4.ClipText, _
                                        strData, _
                                        "I"
    End If
    If mskNumero5.ClipText <> "" Then
      objBoletoArrec.InserirBoletoArrec lngTURNOID, _
                                        lngCAIXAARRECID, _
                                        mskNumero5.ClipText, _
                                        strData, _
                                        "I"
    End If
    If mskNumero6.ClipText <> "" Then
      objBoletoArrec.InserirBoletoArrec lngTURNOID, _
                                        lngCAIXAARRECID, _
                                        mskNumero6.ClipText, _
                                        strData, _
                                        "I"
    End If
    If mskNumero7.ClipText <> "" Then
      objBoletoArrec.InserirBoletoArrec lngTURNOID, _
                                        lngCAIXAARRECID, _
                                        mskNumero7.ClipText, _
                                        strData, _
                                        "I"
    End If
    If mskNumero8.ClipText <> "" Then
      objBoletoArrec.InserirBoletoArrec lngTURNOID, _
                                        lngCAIXAARRECID, _
                                        mskNumero8.ClipText, _
                                        strData, _
                                        "I"
    End If
    If mskNumero9.ClipText <> "" Then
      objBoletoArrec.InserirBoletoArrec lngTURNOID, _
                                        lngCAIXAARRECID, _
                                        mskNumero9.ClipText, _
                                        strData, _
                                        "I"
    End If
    If mskNumero10.ClipText <> "" Then
      objBoletoArrec.InserirBoletoArrec lngTURNOID, _
                                        lngCAIXAARRECID, _
                                        mskNumero10.ClipText, _
                                        strData, _
                                        "I"
    End If
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  End If
  Set objBoletoArrec = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  cmdOk.Enabled = True
End Sub


Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  Dim lngInicio             As Long
  Dim lngTermino            As Long
  Dim strSql                As String
  Dim strWhere              As String
  Dim objRs                 As ADODB.Recordset
  Dim objGeral              As busSisMaq.clsGeral
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  If Not Valida_String(cboArrecadador, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o Arrecadador" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskInicio, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o início válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskTermino, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o término válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  '
  If Not Valida_Moeda(mskNumero1, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número 1 válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskNumero2, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número 2 válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskNumero3, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número 3 válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskNumero4, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número 4 válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskNumero5, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número 5 válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskNumero6, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número 6 válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskNumero7, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número 7 válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskNumero8, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número 8 válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskNumero9, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número 9 válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskNumero10, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número 10 válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  '
  If Len(strMsg) = 0 Then
    'Verifica se lançou ao menos 1 valor
    If mskInicio.ClipText = "" And mskTermino.ClipText = "" _
        And mskNumero1.ClipText = "" And mskNumero2.ClipText = "" _
        And mskNumero3.ClipText = "" And mskNumero4.ClipText = "" _
        And mskNumero5.ClipText = "" And mskNumero6.ClipText = "" _
        And mskNumero7.ClipText = "" And mskNumero8.ClipText = "" _
        And mskNumero9.ClipText = "" And mskNumero10.ClipText = "" Then
      strMsg = strMsg & "Preencher ao menos 1 boleto" & vbCrLf
      tabDetalhes.Tab = 0
    End If
  End If
  '
  If Len(strMsg) = 0 Then
    'Outras validações
    If (mskInicio.ClipText <> "" Or mskTermino.ClipText <> "") And _
      (mskInicio.ClipText = "" Or mskTermino.ClipText = "") Then
      strMsg = strMsg & "Preencher o início e término" & vbCrLf
      tabDetalhes.Tab = 0
    End If
  End If
  If Len(strMsg) = 0 Then
    If mskInicio.ClipText <> "" And mskTermino.ClipText <> "" Then
      'Preencheu início e término
      lngInicio = CCur(mskInicio.ClipText)
      lngTermino = CCur(mskTermino.ClipText)
      '
      If lngTermino < lngInicio Then
        strMsg = strMsg & "Preencher o término maior ou igual ao início" & vbCrLf
        tabDetalhes.Tab = 0
      ElseIf (lngTermino - lngInicio + 1) > 20 Then
        strMsg = strMsg & "Preencher uma faixa sequencial não superior a 20 Boletos" & vbCrLf
        tabDetalhes.Tab = 0
      End If
    End If
  End If
  If Len(strMsg) = 0 Then
    'Verificar se os boletos já foram cadastrados anteriormente
    Set objGeral = New busSisMaq.clsGeral
    strSql = "SELECT BOLETOARREC.PKID, BOLETOARREC.STATUS " & _
          " FROM BOLETOARREC "
    strWhere = " WHERE "
    If mskInicio.ClipText <> "" And mskTermino.ClipText <> "" Then
      strWhere = strWhere & " (BOLETOARREC.NUMERO >= " & Formata_Dados(mskInicio.ClipText, tpDados_Longo) & _
      " AND BOLETOARREC.NUMERO <= " & Formata_Dados(mskTermino.ClipText, tpDados_Longo) & ") "
    End If
    If mskNumero1.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOARREC.NUMERO = " & Formata_Dados(mskNumero1.ClipText, tpDados_Longo)
    End If
    If mskNumero2.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOARREC.NUMERO = " & Formata_Dados(mskNumero2.ClipText, tpDados_Longo)
    End If
    If mskNumero3.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOARREC.NUMERO = " & Formata_Dados(mskNumero3.ClipText, tpDados_Longo)
    End If
    If mskNumero4.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOARREC.NUMERO = " & Formata_Dados(mskNumero4.ClipText, tpDados_Longo)
    End If
    If mskNumero5.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOARREC.NUMERO = " & Formata_Dados(mskNumero5.ClipText, tpDados_Longo)
    End If
    If mskNumero6.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOARREC.NUMERO = " & Formata_Dados(mskNumero6.ClipText, tpDados_Longo)
    End If
    If mskNumero7.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOARREC.NUMERO = " & Formata_Dados(mskNumero7.ClipText, tpDados_Longo)
    End If
    If mskNumero8.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOARREC.NUMERO = " & Formata_Dados(mskNumero8.ClipText, tpDados_Longo)
    End If
    If mskNumero9.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOARREC.NUMERO = " & Formata_Dados(mskNumero9.ClipText, tpDados_Longo)
    End If
    If mskNumero10.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOARREC.NUMERO = " & Formata_Dados(mskNumero10.ClipText, tpDados_Longo)
    End If
    strSql = strSql & strWhere
    Set objRs = objGeral.ExecutarSQL(strSql)
    objRs.Filter = "STATUS <> 'C'"
    If Not objRs.EOF Then
      strMsg = strMsg & "Há Boletos cadastrados com este número para arrecadadores" & vbCrLf
      tabDetalhes.Tab = 0
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
  End If
  If Len(strMsg) = 0 Then
    'Verificar se os boletos já foram cadastrados anteriormente para atendente
    Set objGeral = New busSisMaq.clsGeral
    strSql = "SELECT BOLETOATEND.PKID, BOLETOATEND.STATUS " & _
          " FROM BOLETOATEND "
    strWhere = " WHERE "
    If mskInicio.ClipText <> "" And mskTermino.ClipText <> "" Then
      strWhere = strWhere & " (BOLETOATEND.NUMERO >= " & Formata_Dados(mskInicio.ClipText, tpDados_Longo) & _
      " AND BOLETOATEND.NUMERO <= " & Formata_Dados(mskTermino.ClipText, tpDados_Longo) & ") "
    End If
    If mskNumero1.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOATEND.NUMERO = " & Formata_Dados(mskNumero1.ClipText, tpDados_Longo)
    End If
    If mskNumero2.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOATEND.NUMERO = " & Formata_Dados(mskNumero2.ClipText, tpDados_Longo)
    End If
    If mskNumero3.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOATEND.NUMERO = " & Formata_Dados(mskNumero3.ClipText, tpDados_Longo)
    End If
    If mskNumero4.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOATEND.NUMERO = " & Formata_Dados(mskNumero4.ClipText, tpDados_Longo)
    End If
    If mskNumero5.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOATEND.NUMERO = " & Formata_Dados(mskNumero5.ClipText, tpDados_Longo)
    End If
    If mskNumero6.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOATEND.NUMERO = " & Formata_Dados(mskNumero6.ClipText, tpDados_Longo)
    End If
    If mskNumero7.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOATEND.NUMERO = " & Formata_Dados(mskNumero7.ClipText, tpDados_Longo)
    End If
    If mskNumero8.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOATEND.NUMERO = " & Formata_Dados(mskNumero8.ClipText, tpDados_Longo)
    End If
    If mskNumero9.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOATEND.NUMERO = " & Formata_Dados(mskNumero9.ClipText, tpDados_Longo)
    End If
    If mskNumero10.ClipText <> "" Then
      If strWhere <> " WHERE " Then strWhere = strWhere & " OR "
      strWhere = strWhere & " BOLETOATEND.NUMERO = " & Formata_Dados(mskNumero10.ClipText, tpDados_Longo)
    End If
    strSql = strSql & strWhere
    Set objRs = objGeral.ExecutarSQL(strSql)
    objRs.Filter = "STATUS <> 'C'"
    If Not objRs.EOF Then
      strMsg = strMsg & "Há Boletos cadastrados com este número para atendente" & vbCrLf
      tabDetalhes.Tab = 0
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserBoletoArrecInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserBoletoArrecInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco cboArrecadador
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserBoletoArrecInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objBoletoArrec         As busSisMaq.clsBoletoArrec
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 4890
  Me.Width = 10560
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  'Arrecadador
  strSql = "Select PESSOA.NOME " & _
        " FROM PESSOA " & _
        " INNER JOIN ARRECADADOR ON PESSOA.PKID = ARRECADADOR.PESSOAID " & _
        " ORDER BY PESSOA.NOME"
  PreencheCombo cboArrecadador, strSql, False, True
  '
  If Status = tpStatus_Incluir Then
    '
    cboArrecadador.Enabled = True
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    cboArrecadador.Enabled = False
    '
'''    Set objBoletoArrec = New busSisMaq.clsBoletoArrec
'''    Set objRs = objBoletoArrec.SelecionarBoletoArrecPeloPkid(lngPKID)
'''    '
'''    If Not objRs.EOF Then
'''      txtTurno.Text = RetornaDescTurnoCorrente(objRs.Fields("TURNOENTRADAID").Value)
'''      txtData.Text = Format(objRs.Fields("DATA").Value, "DD/MM/YYYY hh:mm")
'''      INCLUIR_VALOR_NO_MASK mskInicio, objRs.Fields("VALOR").Value, TpMaskMoeda
'''      cboArrecadador.Text = objRs.Fields("DESC_ARRECADADOR").Value & ""
'''
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    '
'''    Set objBoletoArrec = Nothing
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

Private Sub mskInicio_GotFocus()
  Seleciona_Conteudo_Controle mskInicio
End Sub
Private Sub mskInicio_LostFocus()
  Pintar_Controle mskInicio, tpCorContr_Normal
End Sub

Private Sub mskNumero1_GotFocus()
  Seleciona_Conteudo_Controle mskNumero1
End Sub
Private Sub mskNumero1_LostFocus()
  Pintar_Controle mskNumero1, tpCorContr_Normal
End Sub
Private Sub mskNumero2_GotFocus()
  Seleciona_Conteudo_Controle mskNumero2
End Sub
Private Sub mskNumero2_LostFocus()
  Pintar_Controle mskNumero2, tpCorContr_Normal
End Sub
Private Sub mskNumero3_GotFocus()
  Seleciona_Conteudo_Controle mskNumero3
End Sub
Private Sub mskNumero3_LostFocus()
  Pintar_Controle mskNumero3, tpCorContr_Normal
End Sub
Private Sub mskNumero4_GotFocus()
  Seleciona_Conteudo_Controle mskNumero4
End Sub
Private Sub mskNumero4_LostFocus()
  Pintar_Controle mskNumero4, tpCorContr_Normal
End Sub
Private Sub mskNumero5_GotFocus()
  Seleciona_Conteudo_Controle mskNumero5
End Sub
Private Sub mskNumero5_LostFocus()
  Pintar_Controle mskNumero5, tpCorContr_Normal
End Sub
Private Sub mskNumero6_GotFocus()
  Seleciona_Conteudo_Controle mskNumero6
End Sub
Private Sub mskNumero6_LostFocus()
  Pintar_Controle mskNumero6, tpCorContr_Normal
End Sub
Private Sub mskNumero7_GotFocus()
  Seleciona_Conteudo_Controle mskNumero7
End Sub
Private Sub mskNumero7_LostFocus()
  Pintar_Controle mskNumero7, tpCorContr_Normal
End Sub
Private Sub mskNumero8_GotFocus()
  Seleciona_Conteudo_Controle mskNumero8
End Sub
Private Sub mskNumero8_LostFocus()
  Pintar_Controle mskNumero8, tpCorContr_Normal
End Sub
Private Sub mskNumero9_GotFocus()
  Seleciona_Conteudo_Controle mskNumero9
End Sub
Private Sub mskNumero9_LostFocus()
  Pintar_Controle mskNumero9, tpCorContr_Normal
End Sub
Private Sub mskNumero10_GotFocus()
  Seleciona_Conteudo_Controle mskNumero10
End Sub
Private Sub mskNumero10_LostFocus()
  Pintar_Controle mskNumero10, tpCorContr_Normal
End Sub


Private Sub mskTermino_GotFocus()
  Seleciona_Conteudo_Controle mskTermino
End Sub
Private Sub mskTermino_LostFocus()
  Pintar_Controle mskTermino, tpCorContr_Normal
End Sub

