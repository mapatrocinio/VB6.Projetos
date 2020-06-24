VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Begin VB.Form frmUserAtendimentoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atendimento"
   ClientHeight    =   8520
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8520
   ScaleWidth      =   11910
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   8520
      Left            =   10050
      ScaleHeight     =   8520
      ScaleWidth      =   1860
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2055
         Left            =   90
         ScaleHeight     =   1995
         ScaleWidth      =   1605
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   6390
         Width           =   1665
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   1
            Top             =   150
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   2
            Top             =   990
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   8325
      Left            =   60
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   120
      Width           =   9855
      _ExtentX        =   17383
      _ExtentY        =   14684
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userAtendimentoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Histórico Receita"
      TabPicture(1)   =   "userAtendimentoInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "pictrava(2)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Histórico de &Receitas"
      TabPicture(2)   =   "userAtendimentoInc.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5(11)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "grdAtendimento"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "cmdReceitaScanner1"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).ControlCount=   3
      Begin VB.CommandButton cmdReceitaScanner1 
         Caption         =   "&A"
         Height          =   855
         Left            =   -66150
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   390
         Width           =   900
      End
      Begin VB.PictureBox pictrava 
         BorderStyle     =   0  'None
         Height          =   7845
         Index           =   2
         Left            =   -74880
         ScaleHeight     =   7845
         ScaleWidth      =   9675
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   390
         Width           =   9675
         Begin VB.TextBox txtHistoricoReceita 
            Height          =   7815
            Left            =   1380
            Locked          =   -1  'True
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   25
            Text            =   "userAtendimentoInc.frx":0054
            Top             =   0
            Width           =   8175
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Histórico de Recceitas"
            Height          =   435
            Index           =   0
            Left            =   60
            TabIndex        =   26
            Top             =   30
            Width           =   1215
         End
      End
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
         Height          =   7815
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   9585
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2505
            Index           =   3
            Left            =   120
            ScaleHeight     =   2505
            ScaleWidth      =   9345
            TabIndex        =   29
            TabStop         =   0   'False
            Top             =   5070
            Width           =   9345
            Begin VB.TextBox txtSelecionouScanner 
               BackColor       =   &H80000000&
               Height          =   285
               Left            =   6210
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   37
               Text            =   "txtSelecionouScanner"
               Top             =   1740
               Visible         =   0   'False
               Width           =   2835
            End
            Begin VB.TextBox txtArquivo 
               BackColor       =   &H80000000&
               Height          =   285
               Left            =   6210
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   36
               Text            =   "txtArquivo"
               Top             =   1350
               Visible         =   0   'False
               Width           =   2835
            End
            Begin VB.TextBox txtCaminho 
               BackColor       =   &H80000000&
               Height          =   285
               Left            =   6210
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   35
               Text            =   "txtCaminho"
               Top             =   960
               Visible         =   0   'False
               Width           =   2835
            End
            Begin VB.PictureBox Picture1 
               Height          =   5085
               Left            =   1380
               MousePointer    =   2  'Cross
               ScaleHeight     =   5025
               ScaleWidth      =   4755
               TabIndex        =   34
               Top             =   30
               Width           =   4815
               Begin VB.Image Image1 
                  Height          =   5025
                  Left            =   0
                  Stretch         =   -1  'True
                  Top             =   0
                  Width           =   4755
               End
            End
            Begin VB.CommandButton cmdReceitaScanner 
               Caption         =   "&A"
               Height          =   855
               Left            =   6240
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   30
               Width           =   900
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               Caption         =   "Clique na imagem ao lado para ampliá-la"
               ForeColor       =   &H000000FF&
               Height          =   675
               Index           =   10
               Left            =   90
               TabIndex        =   33
               Top             =   1230
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Receita Scanner"
               Height          =   195
               Index           =   8
               Left            =   0
               TabIndex        =   31
               Top             =   0
               Width           =   1215
            End
         End
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2505
            Index           =   1
            Left            =   120
            ScaleHeight     =   2505
            ScaleWidth      =   9345
            TabIndex        =   21
            TabStop         =   0   'False
            Top             =   2550
            Width           =   9345
            Begin VB.CommandButton cmdProcedimento 
               Caption         =   "&A"
               Height          =   855
               Left            =   8370
               Style           =   1  'Graphical
               TabIndex        =   28
               Top             =   60
               Width           =   900
            End
            Begin VB.TextBox txtReceita 
               Height          =   5025
               Left            =   1380
               MultiLine       =   -1  'True
               ScrollBars      =   2  'Vertical
               TabIndex        =   23
               Text            =   "userAtendimentoInc.frx":006A
               Top             =   60
               Width           =   6915
            End
            Begin VB.Label Label5 
               Alignment       =   2  'Center
               Caption         =   "Ctrl + ENTER = Pular linha"
               ForeColor       =   &H000000FF&
               Height          =   675
               Index           =   9
               Left            =   90
               TabIndex        =   30
               Top             =   990
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Recceita"
               Height          =   195
               Index           =   13
               Left            =   60
               TabIndex        =   22
               Top             =   60
               Width           =   1215
            End
         End
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   2385
            Index           =   0
            Left            =   120
            ScaleHeight     =   2385
            ScaleWidth      =   9345
            TabIndex        =   6
            TabStop         =   0   'False
            Top             =   150
            Width           =   9345
            Begin VB.TextBox txtAtendente 
               BackColor       =   &H80000000&
               Height          =   285
               Left            =   1380
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   20
               Text            =   "txtAtendente"
               Top             =   2040
               Width           =   7815
            End
            Begin VB.TextBox txtSala 
               BackColor       =   &H80000000&
               Height          =   285
               Left            =   1380
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   19
               Text            =   "txtSala"
               Top             =   1710
               Width           =   795
            End
            Begin VB.TextBox txtPrestador 
               BackColor       =   &H80000000&
               Height          =   285
               Left            =   1380
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   18
               Text            =   "txtPrestador"
               Top             =   1380
               Width           =   7815
            End
            Begin VB.TextBox txtEspecialidade 
               BackColor       =   &H80000000&
               Height          =   285
               Left            =   1380
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   17
               Text            =   "txtEspecialidade"
               Top             =   1050
               Width           =   7815
            End
            Begin VB.TextBox txtProntuario 
               BackColor       =   &H80000000&
               Height          =   285
               Left            =   1380
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   16
               Text            =   "txtProntuario"
               Top             =   720
               Width           =   7815
            End
            Begin VB.TextBox txtSequencial 
               BackColor       =   &H80000000&
               Height          =   285
               Left            =   1380
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   15
               Text            =   "txtSequencial"
               Top             =   390
               Width           =   1605
            End
            Begin VB.TextBox txtHora 
               BackColor       =   &H80000000&
               Height          =   285
               Left            =   1380
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   0
               Text            =   "txtHora"
               Top             =   60
               Width           =   1605
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Atendente"
               Height          =   195
               Index           =   7
               Left            =   60
               TabIndex        =   14
               Top             =   2040
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Sala"
               Height          =   195
               Index           =   6
               Left            =   60
               TabIndex        =   13
               Top             =   1710
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Prestador"
               Height          =   195
               Index           =   5
               Left            =   60
               TabIndex        =   12
               Top             =   1380
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Especialidade"
               Height          =   195
               Index           =   4
               Left            =   60
               TabIndex        =   11
               Top             =   1050
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Prontuário"
               Height          =   195
               Index           =   3
               Left            =   60
               TabIndex        =   10
               Top             =   720
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Sequencial"
               Height          =   195
               Index           =   2
               Left            =   60
               TabIndex        =   9
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Hora"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   8
               Top             =   90
               Width           =   1215
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdAtendimento 
         Height          =   7425
         Left            =   -74910
         OleObjectBlob   =   "userAtendimentoInc.frx":0077
         TabIndex        =   27
         Top             =   390
         Width           =   8715
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "St assume S = Scanner, N = Não Scanner e A = Importação Automática"
         ForeColor       =   &H000000FF&
         Height          =   435
         Index           =   11
         Left            =   -74880
         TabIndex        =   39
         Top             =   7800
         Width           =   3135
      End
   End
End
Attribute VB_Name = "frmUserAtendimentoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                   As tpStatus
Public blnRetorno               As Boolean
Public blnFechar                As Boolean
Public lngGRID                  As Long
Public lngPKID                  As Long
Public lngPRONTUARIOID          As Long
Public strNomeProntuario        As String
Public strStatus                As String
Public blnTrabComScaner         As Boolean

Private blnPrimeiraVez          As Boolean

Public strHora                  As String
Public strSequencial            As String
Public strProntuario            As String
Public strEspecialidade         As String
Public strPrestador             As String
Public strSala                  As String
Public strAtendente             As String



Dim ATEND_COLUNASMATRIZ         As Long
Dim ATEND_LINHASMATRIZ          As Long
Private ATEND_Matriz()          As String

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Atendimento
  LimparCampoTexto txtHora
  LimparCampoTexto txtSequencial
  LimparCampoTexto txtProntuario
  LimparCampoTexto txtEspecialidade
  LimparCampoTexto txtPrestador
  LimparCampoTexto txtSala
  LimparCampoTexto txtAtendente
  LimparCampoTexto txtReceita
  LimparCampoTexto txtHistoricoReceita
  '
  Image1.Picture = LoadPicture("")
  Image1.Refresh
  LimparCampoTexto txtArquivo
  LimparCampoTexto txtCaminho
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserAtendimentoInc.LimparCampos]", _
            Err.Description
End Sub

Private Sub TratarConfiguracao()
  Dim sMask As String
  Dim intTopSuperior As Integer
  Dim intTopInferior As Integer
  Dim intHeightMaior As Integer
  Dim intHeightMenor As Integer
  '
  On Error GoTo trata
  '
  intTopSuperior = 2550
  intTopInferior = 5100
  intHeightMaior = 5205
  intHeightMenor = 2505
  '
  If blnTrabComScaner Then
    'Trabalha com scaner
    pictrava(3).Top = intTopSuperior
    pictrava(3).Height = intHeightMaior
    pictrava(1).Top = intTopInferior
    pictrava(1).Height = intHeightMenor
    pictrava(3).Visible = True
    pictrava(1).Visible = False
    pictrava(3).Enabled = True
    pictrava(1).Enabled = False
  Else
    'Não trabalha com scaner
    pictrava(1).Top = intTopSuperior
    pictrava(1).Height = intHeightMaior
    pictrava(3).Top = intTopInferior
    pictrava(3).Height = intHeightMenor
    pictrava(1).Visible = True
    pictrava(3).Visible = False
    pictrava(1).Enabled = True
    pictrava(3).Enabled = False
  End If
  '
  txtHora.Text = strHora & ""
  txtSequencial.Text = strSequencial & ""
  txtProntuario.Text = strProntuario & ""
  txtEspecialidade.Text = strEspecialidade & ""
  txtPrestador.Text = strPrestador & ""
  txtSala.Text = strSala & ""
  txtAtendente.Text = strAtendente & ""
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserAtendimentoInc.TratarConfiguracao]", _
            Err.Description
End Sub


'''Private Sub cboPredio_LostFocus()
'''  Pintar_Controle cboPredio, tpCorContr_Normal
'''End Sub
'''
'''Private Sub cmdAlterar_Click()
'''  On Error GoTo trata
'''  Select Case tabDetalhes.Tab
'''  Case 1
'''    If Not IsNumeric(grdAtendimento.Columns("PKID").Value & "") Then
'''      MsgBox "Selecione um atendimento !", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendimento
'''      Exit Sub
'''    End If
'''
'''    frmUserAtendeInc.lngGRID = grdAtendimento.Columns("PKID").Value
'''    frmUserAtendeInc.lngSALAID = lngGRID
'''    frmUserAtendeInc.strDescrAtendimento = cboPredio.Text & " - " & txtNumero.Text
'''    frmUserAtendeInc.Status = tpStatus_Alterar
'''    frmUserAtendeInc.Show vbModal
'''
'''    If frmUserAtendeInc.blnRetorno Then
'''      CarregaHistoricoReceita
'''      grdAtendimento.Bookmark = Null
'''      grdAtendimento.ReBind
'''      grdAtendimento.ApproxCount = ATEND_LINHASMATRIZ
'''    End If
'''    SetarFoco grdAtendimento
'''  End Select
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub
'''
Private Sub cmdCancelar_Click()
  On Error GoTo trata
  blnFechar = True
  blnRetorno = True
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub
'''
'''
'''
'''
'''Private Sub cmdExcluir_Click()
'''  On Error GoTo trata
'''  Dim objAtende     As busSisMed.clsAtende
'''  '
'''  Select Case tabDetalhes.Tab
'''  Case 1 'Exclusão de grade de atendimento
'''    '
'''    If Len(Trim(grdAtendimento.Columns("PKID").Value & "")) = 0 Then
'''      MsgBox "Selecione um atendimento.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdAtendimento
'''      Exit Sub
'''    End If
'''    '
'''    Set objAtende = New busSisMed.clsAtende
'''    '
'''    If MsgBox("Confirma exclusão do ítem da grade de atendimento " & grdAtendimento.Columns("Dia da Semana").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdAtendimento
'''      Exit Sub
'''    End If
'''    'OK
'''    objAtende.ExcluirAtende CLng(grdAtendimento.Columns("PKID").Value)
'''    '
'''    CarregaHistoricoReceita
'''    grdAtendimento.Bookmark = Null
'''    grdAtendimento.ReBind
'''    grdAtendimento.ApproxCount = ATEND_LINHASMATRIZ
'''
'''    Set objAtende = Nothing
'''    SetarFoco grdAtendimento
'''  End Select
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub
'''
'''
'''
'''
'''
'''Private Sub cmdIncluir_Click()
'''  On Error GoTo trata
'''  Dim objForm As Form
'''  '
'''  Select Case tabDetalhes.Tab
'''  Case 1
'''    frmUserAtendeInc.Status = tpStatus_Incluir
'''    frmUserAtendeInc.lngGRID = 0
'''    frmUserAtendeInc.lngSALAID = lngGRID
'''    frmUserAtendeInc.strDescrAtendimento = cboPredio.Text & " - " & txtNumero.Text
'''    frmUserAtendeInc.Show vbModal
'''
'''    If frmUserAtendeInc.blnRetorno Then
'''      CarregaHistoricoReceita
'''      grdAtendimento.Bookmark = Null
'''      grdAtendimento.ReBind
'''      grdAtendimento.ApproxCount = ATEND_LINHASMATRIZ
'''    End If
'''    SetarFoco grdAtendimento
'''  End Select
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub




Private Sub cmdOk_Click()
  Dim objAtendimento            As busSisMed.clsAtendimento
  Dim objCC                     As busSisMed.clsContaCorrente
  Dim objGR                     As busSisMed.clsGR
  Dim objGeral                  As busSisMed.clsGeral
  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim strPathArquivo            As String
  Dim strNomeArquivo            As String
  Dim strData                   As String
  Dim datData                   As Date
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  datData = Now
  strData = Format(datData, "DD/MM/YYYY hh:mm")
'''  Set objGeral = New busSisMed.clsGeral
  Set objAtendimento = New busSisMed.clsAtendimento
  Set objGR = New busSisMed.clsGR
  
'''  'PRÉDIO
'''  lngPREDIOID = 0
'''  strSql = "SELECT PKID FROM PREDIO WHERE NOME = " & Formata_Dados(cboPredio.Text, tpDados_Texto)
'''  Set objRs = objGeral.ExecutarSQL(strSql)
'''  If Not objRs.EOF Then
'''    lngPREDIOID = objRs.Fields("PKID").Value
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
'''  Set objGeral = Nothing

  If blnTrabComScaner Then
    If txtSelecionouScanner.Text = "S" Then
      TratarImagemScanner datData, _
                          txtCaminho.Text, _
                          gsPathLocalBackup, _
                          txtArquivo.Text, _
                          lngPRONTUARIOID, _
                          lngGRID, _
                          strPathArquivo, _
                          strNomeArquivo
    End If
    
  End If
  If Status = tpStatus_Alterar Then
    'Alterar Atendimento
    objAtendimento.AlterarAtendimento lngPKID, _
                                      strPathArquivo, _
                                      strNomeArquivo, _
                                      txtReceita.Text
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Atendimento
    objAtendimento.InserirAtendimento lngGRID, _
                                      strData, _
                                      IIf(blnTrabComScaner, "S", "N"), _
                                      strPathArquivo, _
                                      strNomeArquivo, _
                                      txtReceita.Text, _
                                      "", _
                                      0
    If strStatus = "L" Then
      'Altera status da GR para Atendida
      objGR.AlterarStatusGR lngGRID, _
                            "P", _
                            ""
    Else
      'Altera status da GR para Atendida
      objGR.AlterarStatusGR lngGRID, _
                            "A", _
                            ""
    End If
    Set objCC = New busSisMed.clsContaCorrente
    
    
    objCC.InserirFinanceiro lngGRID, _
                            tpIcTipoGR_Prest, _
                            gsNomeUsu
    Set objCC = Nothing
    blnRetorno = True
    blnFechar = True
    Unload Me
  End If
  Set objAtendimento = Nothing
  Set objGR = Nothing
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
  If blnTrabComScaner Then
    'Trabalha com scaner
    If gsNivel = gsArquivista Then
      'Obrigatório apenas para arquivista
      If txtCaminho.Text = "" Then
        strMsg = strMsg & "Nova receita não selecionada do Scanner" & vbCrLf
        tabDetalhes.Tab = 0
      End If
      If Len(strMsg) = 0 Then
        If txtArquivo.Text = "" Then
          strMsg = strMsg & "Nova receita não selecionada do Scanner" & vbCrLf
          tabDetalhes.Tab = 0
        End If
      End If
      If Len(strMsg) = 0 Then
        If txtSelecionouScanner.Text <> "S" Then
          strMsg = strMsg & "Nova receita não selecionada do Scanner" & vbCrLf
          tabDetalhes.Tab = 0
        End If
      End If
    End If
  Else
    'Não Trabalha com scaner
    If Not Valida_String(txtReceita, TpObrigatorio, blnSetarFocoControle) Then
      strMsg = strMsg & "Preencher a descriçao da receita" & vbCrLf
      tabDetalhes.Tab = 0
    End If
    
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserAtendimentoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserAtendimentoInc.ValidaCampos]", _
            Err.Description


End Function

Private Sub cmdProcedimento_Click()
  On Error GoTo trata
  frmUserAtendimentoRecCons.QuemChamou = 0
  frmUserAtendimentoRecCons.lngGRID = lngGRID
  frmUserAtendimentoRecCons.Show vbModal
  txtReceita.Text = frmUserAtendimentoRecCons.strRetorno
  SetarFoco txtReceita
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub


Private Sub cmdReceitaScanner_Click()
  On Error GoTo trata
  'frmUserAtendimentoScannerCons.QuemChamou = 0
  'frmUserAtendimentoScannerCons.lngGRID = lngGRID
  frmUserAtendimentoScannerCons.Show vbModal
  If frmUserAtendimentoScannerCons.strCaminhoFinal <> "" Then
    txtCaminho.Text = frmUserAtendimentoScannerCons.strCaminhoFinal
    txtArquivo.Text = frmUserAtendimentoScannerCons.strArquivoFinal
    txtSelecionouScanner.Text = "S"
    Image1.Picture = LoadPicture(frmUserAtendimentoScannerCons.strCaminhoFinal & frmUserAtendimentoScannerCons.strArquivoFinal)
    Image1.Refresh
  End If
  '
  SetarFoco cmdReceitaScanner
  Exit Sub
trata:
  'Tratamento de erro de leitura de imagem (File not found)
  If Err.Number = 53 Then
    Image1.Picture = LoadPicture(gsIconsPath & "Excluir.ico")
    Image1.Refresh
    txtCaminho.Text = gsIconsPath
    txtArquivo.Text = "Excluir.ico"
    txtSelecionouScanner.Text = "N"
    Resume Next
  End If
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdReceitaScanner1_Click()
  '
  On Error GoTo trata
  '
  If grdAtendimento.Columns(1).Value & "" = "" Then
    AmpN
    TratarErroPrevisto "Selecione uma linha da tabela para exibir a imagem do scanner", "[frmUserAtendimentoInc.grdAtendimento_Click]"
    Exit Sub
  End If
  If grdAtendimento.Columns(2).Value & "" <> "N" Then
    frmUserScannerCons.Image1.Picture = LoadPicture(grdAtendimento.Columns(3).Value & "")
    frmUserScannerCons.Image1.Refresh
    frmUserScannerCons.Show vbModal
  Else
    frmUserScannerRecCons.strDescricao = grdAtendimento.Columns(0).Value & ""
    frmUserScannerRecCons.Show vbModal
  End If
  Exit Sub
trata:
  'Tratamento de erro de leitura de imagem (File not found)
  If Err.Number = 53 Then
    frmUserScannerCons.Image1.Picture = LoadPicture(gsIconsPath & "Excluir.ico")
    frmUserScannerCons.Image1.Refresh
    Resume Next
  End If
  TratarErro Err.Number, Err.Description, "frmUserAtendimentoInc.grdAtendimento_Click"
  AmpN

End Sub

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    If blnTrabComScaner Then
      SetarFoco cmdReceitaScanner
    Else
      SetarFoco txtReceita
    End If
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAtendimentoInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objGeral                As busSisMed.clsGeral
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objAtendimento          As busSisMed.clsAtendimento
  '
  blnFechar = False
  blnRetorno = False
  '
  AmpS
  Me.Height = 9000
  Me.Width = 12000
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar
  LerFigurasAvulsas cmdProcedimento, "PEN04.ICO", "PEN05.ICO", "Procedimentos"
  LerFigurasAvulsas cmdReceitaScanner, "SCANER.BMP", "SCANERDOWN.BMP", "Scanner"
  LerFigurasAvulsas cmdReceitaScanner1, "SCANER.BMP", "SCANERDOWN.BMP", "Scanner"
  '
  'Limpar Campos
  LimparCampos
  'Verifica se está em evento de inclusão ou alteração da GR
  
  Set objGeral = New busSisMed.clsGeral
  strSql = "Select ATENDIMENTO.PKID, ATENDIMENTO.INDSCANER from ATENDIMENTO " & _
      " WHERE ATENDIMENTO.GRID = " & Formata_Dados(lngGRID, tpDados_Longo)
  '
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then
    'Evento de alteração
    lngPKID = objRs.Fields("PKID").Value
    Status = tpStatus_Alterar
    blnTrabComScaner = IIf(objRs.Fields("INDSCANER").Value & "" = "S", True, False)
  Else
    'Evento de inclusão
    lngPKID = 0
    Status = tpStatus_Incluir
    blnTrabComScaner = gbTrabComScaner
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strSql = "Select PRONTUARIO.PKID, PRONTUARIO.NOME from GR " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
      " WHERE GR.PKID = " & Formata_Dados(lngGRID, tpDados_Longo)
  '
  Set objRs = objGeral.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then
    lngPRONTUARIOID = objRs.Fields("PKID").Value
    strNomeProntuario = objRs.Fields("NOME").Value & ""
  Else
    lngPRONTUARIOID = 0
    strNomeProntuario = ""
  End If
  objRs.Close
  Set objRs = Nothing
  '
  Set objGeral = Nothing
  '
  txtSelecionouScanner.Text = "N"
  '
  TratarConfiguracao
  tabDetalhes_Click 0
  tabDetalhes.TabVisible(1) = False
  '
  If Status = tpStatus_Incluir Then
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = True
    tabDetalhes.TabEnabled(2) = True
    '
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objAtendimento = New busSisMed.clsAtendimento
    Set objRs = objAtendimento.SelecionarAtendimentoPeloPkid(lngPKID)
    '
    If Not objRs.EOF Then
      txtReceita.Text = objRs.Fields("DESCRICAO").Value & ""
      'txtReceitaScaner.Text = objRs.Fields("PATHARQUIVO").Value & objRs.Fields("NOMEARQUIVO").Value & ""
      Image1.Picture = LoadPicture(objRs.Fields("PATHARQUIVO").Value & objRs.Fields("NOMEARQUIVO").Value & "")
      Image1.Refresh
      txtCaminho.Text = objRs.Fields("PATHARQUIVO").Value & ""
      txtArquivo.Text = objRs.Fields("NOMEARQUIVO").Value & ""
    End If
    objRs.Close
    Set objRs = Nothing
    '
    Set objAtendimento = Nothing
    '
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = True
    tabDetalhes.TabEnabled(2) = True
    '
  End If
  '
  AmpN
  Exit Sub
trata:
  'Tratamento de erro de leitura de imagem (File not found)
  If Err.Number = 53 Then
    Image1.Picture = LoadPicture(gsIconsPath & "Excluir.ico")
    Image1.Refresh
    txtCaminho.Text = gsIconsPath
    txtArquivo.Text = "Excluir.ico"
    Resume Next
  End If
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
  Resume Next
  Unload Me
End Sub




Private Sub Image1_Click()
  '
  On Error GoTo trata
  AmpS
  '
  If Image1.Picture = 0 Then
    AmpN
    TratarErroPrevisto "Imagem não selecionada", "[frmUserAtendimentoInc.Image1_Click]"
    Exit Sub
  End If
  frmUserScannerCons.Image1.Picture = LoadPicture(txtCaminho.Text & txtArquivo.Text)
  frmUserScannerCons.Show vbModal
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, "frmUserAtendimentoInc.tabDetalhes"
End Sub


'''
'''Private Sub Form_Unload(Cancel As Integer)
'''  If Not blnFechar Then Cancel = True
'''End Sub
'''
'''
'''Private Sub txtAndar_GotFocus()
'''  Seleciona_Conteudo_Controle txtAndar
'''End Sub
'''Private Sub txtAndar_LostFocus()
'''  Pintar_Controle txtAndar, tpCorContr_Normal
'''End Sub
'''Private Sub txtNumero_GotFocus()
'''  Seleciona_Conteudo_Controle txtNumero
'''End Sub
'''Private Sub txtNumero_LostFocus()
'''  Pintar_Controle txtNumero, tpCorContr_Normal
'''End Sub
'''
Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    pictrava(0).Enabled = True
    TratarConfiguracao
    pictrava(2).Enabled = False
    grdAtendimento.Enabled = False
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    '
    SetarFoco txtReceita
  Case 1
    pictrava(0).Enabled = False
    pictrava(1).Enabled = False
    pictrava(3).Enabled = False
    pictrava(2).Enabled = True
    grdAtendimento.Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    '
    CarregaHistoricoReceita
    '
    SetarFoco txtHistoricoReceita
  Case 2
    pictrava(0).Enabled = False
    pictrava(1).Enabled = False
    pictrava(3).Enabled = False
    pictrava(2).Enabled = False
    grdAtendimento.Enabled = True
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    '
    'Montar RecordSet
    ATEND_COLUNASMATRIZ = grdAtendimento.Columns.Count
    ATEND_LINHASMATRIZ = 0
    ATEND_Monta_Matriz
    grdAtendimento.Bookmark = Null
    grdAtendimento.ReBind
    grdAtendimento.ApproxCount = ATEND_LINHASMATRIZ
    '
    SetarFoco grdAtendimento
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "frmUserAtendimentoInc.tabDetalhes"
  AmpN
End Sub


Private Sub grdAtendimento_UnboundReadDataEx( _
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
               Offset + intI, ATEND_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, ATEND_COLUNASMATRIZ, ATEND_LINHASMATRIZ, ATEND_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, ATEND_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserAtendimentoInc.grdAtendimento_UnboundReadDataEx]"
End Sub

Public Sub ATEND_Monta_Matriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  '
  On Error GoTo trata

  Set clsGer = New busSisMed.clsGeral
  '
  'strSql = "SELECT ATENDIMENTO.DATA, ATENDIMENTO.PATHARQUIVO + ATENDIMENTO.NOMEARQUIVO, ESPECIALIDADE.ESPECIALIDADE, PRESTADOR.NOME " & _
          "FROM ATENDIMENTO " & _
          " INNER JOIN GR ON GR.PKID = ATENDIMENTO.GRID " & _
          " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
          " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
            " INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
            " INNER JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
          " INNER JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = GR.ESPECIALIDADEID " & _
          "WHERE PRONTUARIO.PKID = " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo) & _
          " AND ATENDIMENTO.INDSCANER = " & Formata_Dados("S", tpDados_Texto) & _
          " ORDER BY ATENDIMENTO.DATA "

  strSql = "SELECT ATENDIMENTO.DESCRICAO, ATENDIMENTO.DATA, ATENDIMENTO.INDSCANER, ATENDIMENTO.PATHARQUIVO + ATENDIMENTO.NOMEARQUIVO, ESPECIALIDADE.ESPECIALIDADE, PRESTADOR.NOME " & _
          "FROM ATENDIMENTO " & _
          " LEFT JOIN GR ON GR.PKID = ATENDIMENTO.GRID " & _
          " LEFT JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
          " LEFT JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
            " LEFT JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
            " LEFT JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
          " LEFT JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = GR.ESPECIALIDADEID " & _
          " LEFT JOIN PRONTUARIO PRONTUARIO_ATENDE ON PRONTUARIO_ATENDE.PKID = ATENDIMENTO.PRONTUARIOID " & _
          "WHERE (PRONTUARIO.PKID = " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo) & _
          " OR PRONTUARIO_ATENDE.PKID = " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo) & ")" & _
          " -- AND ATENDIMENTO.INDSCANER = " & Formata_Dados("S", tpDados_Texto) & _
          " ORDER BY ATENDIMENTO.DATA "

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    ATEND_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim ATEND_Matriz(0 To ATEND_COLUNASMATRIZ - 1, 0 To ATEND_LINHASMATRIZ - 1)
  Else
    ReDim ATEND_Matriz(0 To ATEND_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To ATEND_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To ATEND_COLUNASMATRIZ - 1  'varre as colunas
          ATEND_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

'''Private Sub txtTelefone_GotFocus()
'''  Seleciona_Conteudo_Controle txtTelefone
'''End Sub
'''Private Sub txtTelefone_LostFocus()
'''  Pintar_Controle txtTelefone, tpCorContr_Normal
'''End Sub


Public Sub CarregaHistoricoReceita()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objGer    As busSisMed.clsGeral
  '
  On Error GoTo trata

  Set objGer = New busSisMed.clsGeral
  '
  strSql = "SELECT ATENDIMENTO.DATA, ATENDIMENTO.DESCRICAO, ESPECIALIDADE.ESPECIALIDADE, PRESTADOR.NOME " & _
          "FROM ATENDIMENTO " & _
          " INNER JOIN GR ON GR.PKID = ATENDIMENTO.GRID " & _
          " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = GR.PRONTUARIOID " & _
          " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
            " INNER JOIN SALA ON SALA.PKID = ATENDE.SALAID " & _
            " INNER JOIN PRONTUARIO AS PRESTADOR ON PRESTADOR.PKID = ATENDE.PRONTUARIOID " & _
          " INNER JOIN ESPECIALIDADE ON ESPECIALIDADE.PKID = GR.ESPECIALIDADEID " & _
          "WHERE PRONTUARIO.PKID = " & Formata_Dados(lngPRONTUARIOID, tpDados_Longo) & _
          " AND ATENDIMENTO.INDSCANER = " & Formata_Dados("N", tpDados_Texto) & _
          " ORDER BY ATENDIMENTO.DATA "
  '
  Set objRs = objGer.ExecutarSQL(strSql)
  '
  txtHistoricoReceita.Text = ""
  Do While Not objRs.EOF
    txtHistoricoReceita.Text = txtHistoricoReceita.Text & "Atendimento: " & Format(objRs.Fields("DATA").Value & "", "DD/MM/YYYY hh:mm") & vbCrLf & _
      "Prestador: " & objRs.Fields("NOME").Value & vbCrLf & _
      "Especialidade: " & objRs.Fields("ESPECIALIDADE").Value & vbCrLf & vbCrLf & _
      "Histórico: " & objRs.Fields("DESCRICAO").Value & "" & vbCrLf & vbCrLf
    objRs.MoveNext
  Loop
  objRs.Close
  Set objRs = Nothing
  Set objGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub



