VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUserProcReceitaInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de tipos de procedimento"
   ClientHeight    =   5970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10935
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5970
   ScaleWidth      =   10935
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   5970
      Left            =   9075
      ScaleHeight     =   5970
      ScaleWidth      =   1860
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2055
         Left            =   90
         ScaleHeight     =   1995
         ScaleWidth      =   1605
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   3720
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   6
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5685
      Left            =   120
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   120
      Width           =   8805
      _ExtentX        =   15531
      _ExtentY        =   10028
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userProcReceitaInc.frx":0000
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
         Height          =   5235
         Left            =   120
         TabIndex        =   10
         Top             =   360
         Width           =   8595
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   4965
            Index           =   0
            Left            =   120
            ScaleHeight     =   4965
            ScaleWidth      =   8385
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   180
            Width           =   8385
            Begin VB.PictureBox Picture5 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1290
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   15
               TabStop         =   0   'False
               Top             =   4500
               Width           =   2235
               Begin VB.OptionButton optAtivo 
                  Caption         =   "Não"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   4
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optAtivo 
                  Caption         =   "Sim"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   3
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   1065
               End
            End
            Begin VB.TextBox txtDescricao 
               Height          =   3705
               Left            =   1320
               MaxLength       =   4000
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   2
               Text            =   "userProcReceitaInc.frx":001C
               Top             =   750
               Width           =   7005
            End
            Begin VB.TextBox txtTipo 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   1
               Text            =   "txtTipo"
               Top             =   420
               Width           =   7005
            End
            Begin VB.TextBox txtProcedimento 
               BackColor       =   &H00E0E0E0&
               Enabled         =   0   'False
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               TabStop         =   0   'False
               Text            =   "txtProcedimento"
               Top             =   90
               Width           =   7005
            End
            Begin VB.Label Label5 
               Caption         =   "Ativo"
               Height          =   285
               Index           =   14
               Left            =   360
               TabIndex        =   16
               Top             =   4530
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Descrição"
               Height          =   195
               Index           =   2
               Left            =   60
               TabIndex        =   14
               Top             =   780
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Procedimento"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   13
               Top             =   135
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Tipo"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   12
               Top             =   450
               Width           =   1215
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserProcReceitaInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean

Public lngPKID                  As Long
Public lngPROCEDIMENTOID           As Long
Public strNomeProcedimento         As String

Private blnPrimeiraVez          As Boolean

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Valor ProcReceita
  LimparCampoTexto txtProcedimento
  LimparCampoTexto txtTipo
  LimparCampoTexto txtDescricao
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserProcReceitaInc.LimparCampos]", _
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

Private Sub cmdOk_Click()
  Dim objProcReceita            As busSisMed.clsProcReceita
  Dim objGeral                  As busSisMed.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim strAtivo                  As String
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objProcReceita = New busSisMed.clsProcReceita
  '
  'Validar se Procedimento já esta associado ao prestador
  strSql = "SELECT * FROM PROCEDIMENTORECEITA " & _
    " WHERE PROCEDIMENTORECEITA.PROCEDIMENTOID = " & Formata_Dados(lngPROCEDIMENTOID, tpDados_Longo) & _
    " AND PROCEDIMENTORECEITA.TIPO = " & Formata_Dados(txtTipo.Text, tpDados_Texto) & _
    " AND PROCEDIMENTORECEITA.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objGeral = New busSisMed.clsGeral
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    Pintar_Controle txtTipo, tpCorContr_Erro
    TratarErroPrevisto "Tipo  de procedimento já cadastrado"
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    Set objProcReceita = Nothing
    cmdOk.Enabled = True
    SetarFoco txtTipo
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  If optAtivo(0).Value Then
    strAtivo = "A"
  ElseIf optAtivo(1).Value Then
    strAtivo = "I"
  Else
    strAtivo = "I"
  End If
  If Status = tpStatus_Alterar Then
    'Alterar ProcReceita
    objProcReceita.AlterarProcReceita lngPKID, _
                                      txtTipo.Text, _
                                      txtDescricao.Text, _
                                      strAtivo
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir ProcReceita
    objProcReceita.InserirProcReceita lngPROCEDIMENTOID, _
                                      txtTipo.Text, _
                                      txtDescricao.Text, _
                                      strAtivo
  End If
  Set objProcReceita = Nothing
  blnRetorno = True
  blnFechar = True
  Unload Me
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
  If Not Valida_String(txtTipo, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o tipo" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(txtDescricao, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a descrição" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Status = tpStatus_Alterar Then
    If Not Valida_Option(optAtivo, blnSetarFocoControle) Then
      strMsg = strMsg & "Selecionar procedimento está ativo" & vbCrLf
      tabDetalhes.Tab = 0
    End If
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserProcReceitaInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserProcReceitaInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco txtTipo
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserProcReceitaInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objProcReceita            As busSisMed.clsProcReceita
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 6450
  Me.Width = 11025
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  'Limpar Campos
  LimparCampos
  '
  txtProcedimento.Text = strNomeProcedimento
  If Status = tpStatus_Incluir Then
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objProcReceita = New busSisMed.clsProcReceita
    Set objRs = objProcReceita.SelecionarProcReceitaPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      txtTipo.Text = objRs.Fields("TIPO").Value & ""
      txtDescricao.Text = objRs.Fields("DESCRICAO").Value & ""
      If objRs.Fields("STATUS").Value & "" = "A" Then
        optAtivo(0).Value = True
        optAtivo(1).Value = False
      ElseIf objRs.Fields("STATUS").Value & "" = "I" Then
        optAtivo(0).Value = False
        optAtivo(1).Value = True
      Else
        optAtivo(0).Value = False
        optAtivo(1).Value = False
      End If
      
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objProcReceita = Nothing
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


Private Sub txtDescricao_GotFocus()
  Seleciona_Conteudo_Controle txtDescricao
End Sub
Private Sub txtDescricao_LostFocus()
  Pintar_Controle txtDescricao, tpCorContr_Normal
End Sub

Private Sub txtTipo_GotFocus()
  Seleciona_Conteudo_Controle txtTipo
End Sub
Private Sub txtTipo_LostFocus()
  Pintar_Controle txtTipo, tpCorContr_Normal
End Sub

