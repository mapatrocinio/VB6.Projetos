VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmUserLivroInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Livro"
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
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   1875
         Left            =   120
         ScaleHeight     =   1815
         ScaleWidth      =   1605
         TabIndex        =   7
         Top             =   2160
         Width           =   1665
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   5
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Default         =   -1  'True
            Height          =   735
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   4
            Top             =   120
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   3855
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   6800
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do livro"
      TabPicture(0)   =   "userLivroInc.frx":0000
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
         TabIndex        =   9
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
            TabIndex        =   10
            Top             =   3480
            Width           =   2295
         End
         Begin VB.ComboBox cboBanco 
            Height          =   315
            Left            =   1560
            Style           =   2  'Dropdown List
            TabIndex        =   1
            Top             =   720
            Width           =   2535
         End
         Begin VB.TextBox txtAgencia 
            Height          =   285
            Left            =   1560
            MaxLength       =   5
            TabIndex        =   2
            Top             =   1080
            Width           =   975
         End
         Begin VB.TextBox txtConta 
            Height          =   285
            Left            =   1560
            MaxLength       =   20
            TabIndex        =   3
            Top             =   1440
            Width           =   2535
         End
         Begin VB.TextBox txtLivro 
            Height          =   285
            Left            =   1560
            MaxLength       =   15
            TabIndex        =   0
            Top             =   360
            Width           =   1935
         End
         Begin VB.Label lblLivro 
            Caption         =   "Banco"
            Height          =   255
            Left            =   120
            TabIndex        =   14
            Top             =   720
            Width           =   1215
         End
         Begin VB.Label lblCheque 
            Caption         =   "Agência"
            Height          =   255
            Left            =   120
            TabIndex        =   13
            Top             =   1080
            Width           =   1215
         End
         Begin VB.Label Label1 
            Caption         =   "Conta"
            Height          =   255
            Left            =   120
            TabIndex        =   12
            Top             =   1440
            Width           =   1215
         End
         Begin VB.Label Label2 
            Caption         =   "Número Livro"
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   360
            Width           =   1215
         End
      End
   End
End
Attribute VB_Name = "frmUserLivroInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                         As tpStatus
Public lngLIVROID                     As Long
Public bRetorno                       As Boolean
Public blnPrimeiraVez                 As Boolean
Public bFechar                        As Boolean

Private Sub cboBanco_LostFocus()
  Pintar_Controle cboBanco, tpCorContr_Normal
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
  Dim objLivro                As busSisMed.clsLivro
  Dim clsGer                  As busSisMed.clsGeral
  Dim lngBANCOID              As Long
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração da Despesa
    If Not ValidaCampos Then Exit Sub
    'Valida se unidade de estoque já cadastrada
    Set clsGer = New busSisMed.clsGeral
    strSql = "Select PKID From LIVRO WHERE NUMEROLIVRO = " & Formata_Dados(txtLivro.Text, tpDados_Texto, tpNulo_NaoAceita) & _
      " AND PKID <> " & lngLIVROID
    Set objRs = clsGer.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set clsGer = Nothing
      TratarErroPrevisto "Livro Já cadastrado", "cmdOK_Click"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    
    strSql = "Select PKID From BANCO WHERE NOME = " & Formata_Dados(cboBanco.Text, tpDados_Texto, tpNulo_NaoAceita)
    Set objRs = clsGer.ExecutarSQL(strSql)
    If objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set clsGer = Nothing
      TratarErroPrevisto "Banco não cadastrado", "cmdOK_Click"
      Exit Sub
    Else
      lngBANCOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    Set clsGer = Nothing

    Set objLivro = New busSisMed.clsLivro
    
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      objLivro.AlterarLivro lngLIVROID, _
                            lngBANCOID, _
                            txtConta.Text, _
                            txtAgencia.Text, _
                            txtLivro.Text
      bRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Código para inclusão
      '
      objLivro.IncluirLivro lngBANCOID, _
                            txtConta.Text, _
                            txtAgencia.Text, _
                            txtLivro.Text
      bRetorno = True
    End If
    Set objLivro = Nothing
    bFechar = True
    Unload Me
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCampos() As Boolean
  Dim strMsg     As String
  '
  If Not IsNumeric(txtLivro.Text) Then
    strMsg = strMsg & "Informar o número do livro válido" & vbCrLf
    Pintar_Controle txtLivro, tpCorContr_Erro
  End If
  If Len(cboBanco.Text) = 0 Then
    strMsg = strMsg & "Selecionar o banco" & vbCrLf
    Pintar_Controle cboBanco, tpCorContr_Erro
  End If
  If Not IsNumeric(txtAgencia.Text) Or Len(txtAgencia.Text) < 4 Then
    strMsg = strMsg & "Informar o número da agência válido" & vbCrLf
    Pintar_Controle txtAgencia, tpCorContr_Erro
  End If
  If Not IsNumeric(txtConta.Text) Or Len(txtConta.Text) < 4 Then
    strMsg = strMsg & "Informar o número da conta válido" & vbCrLf
    Pintar_Controle txtConta, tpCorContr_Erro
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserLivroInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco txtLivro
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserLivroInc.Form_Activate]"
End Sub



Private Sub Form_Load()
On Error GoTo trata
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim objLivro        As busSisMed.clsLivro
  '
  bFechar = False
  bRetorno = False
  AmpS
  Me.Height = 4560
  Me.Width = 10200
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  '
  strSql = "SELECT NOME FROM BANCO ORDER BY NOME;"
  PreencheCombo cboBanco, strSql, False, True
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    LimparCampoTexto txtLivro
    LimparCampoTexto txtAgencia
    LimparCampoTexto txtConta
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set objLivro = New busSisMed.clsLivro
    Set objRs = objLivro.SelecionarLivro(lngLIVROID)
    '
    If Not objRs.EOF Then
      txtLivro.Text = objRs.Fields("NUMEROLIVRO").Value & ""
      cboBanco.Text = objRs.Fields("NOME").Value & ""
      txtAgencia.Text = objRs.Fields("AGENCIA").Value & ""
      txtConta.Text = objRs.Fields("CONTA").Value & ""
    End If
    Set objLivro = Nothing
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

Private Sub txtAgencia_GotFocus()
  Selecionar_Conteudo txtAgencia
End Sub

Private Sub txtAgencia_LostFocus()
  Pintar_Controle txtAgencia, tpCorContr_Normal
End Sub

Private Sub txtConta_GotFocus()
  Selecionar_Conteudo txtConta
End Sub

Private Sub txtConta_LostFocus()
  Pintar_Controle txtConta, tpCorContr_Normal
End Sub

Private Sub txtLivro_GotFocus()
  Selecionar_Conteudo txtLivro
End Sub

Private Sub txtLivro_LostFocus()
  Pintar_Controle txtLivro, tpCorContr_Normal
End Sub

