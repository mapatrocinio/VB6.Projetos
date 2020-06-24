VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserSerieInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Série"
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
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4575
         Left            =   90
         ScaleHeight     =   4515
         ScaleWidth      =   1605
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   810
         Width           =   1665
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   2700
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   855
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3540
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   1830
            Width           =   1335
         End
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   960
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   90
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   5295
      Left            =   120
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   9340
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados cadastrais"
      TabPicture(0)   =   "userSerieInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Máquina"
      TabPicture(1)   =   "userSerieInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdMaquina"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Histórico Série"
      TabPicture(2)   =   "userSerieInc.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdHistSerie"
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
         Left            =   120
         TabIndex        =   14
         Top             =   330
         Width           =   7815
         Begin VB.PictureBox pictrava 
            BorderStyle     =   0  'None
            Height          =   4515
            Index           =   0
            Left            =   120
            ScaleHeight     =   4515
            ScaleWidth      =   7575
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   150
            Width           =   7575
            Begin VB.ComboBox cboDono 
               Height          =   315
               Left            =   1320
               Style           =   2  'Dropdown List
               TabIndex        =   1
               Top             =   420
               Width           =   6105
            End
            Begin VB.PictureBox Picture1 
               BorderStyle     =   0  'None
               Height          =   285
               Left            =   1290
               ScaleHeight     =   285
               ScaleWidth      =   2235
               TabIndex        =   18
               TabStop         =   0   'False
               Top             =   1350
               Width           =   2235
               Begin VB.OptionButton optStatus 
                  Caption         =   "Inativo"
                  Height          =   315
                  Index           =   1
                  Left            =   1290
                  TabIndex        =   5
                  Top             =   0
                  Width           =   1095
               End
               Begin VB.OptionButton optStatus 
                  Caption         =   "Ativo"
                  Height          =   315
                  Index           =   0
                  Left            =   0
                  TabIndex        =   4
                  Top             =   0
                  Value           =   -1  'True
                  Width           =   825
               End
            End
            Begin VB.TextBox txtNumero 
               Height          =   285
               Left            =   1320
               MaxLength       =   100
               TabIndex        =   0
               Text            =   "txtNumero"
               Top             =   75
               Width           =   6075
            End
            Begin MSMask.MaskEdBox mskPercDono 
               Height          =   255
               Left            =   1320
               TabIndex        =   2
               Top             =   780
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0.00;($#,##0.00)"
               PromptChar      =   "_"
            End
            Begin MSMask.MaskEdBox mskCoeficiente 
               Height          =   255
               Left            =   1320
               TabIndex        =   3
               Top             =   1080
               Width           =   1245
               _ExtentX        =   2196
               _ExtentY        =   450
               _Version        =   393216
               BackColor       =   16777215
               Format          =   "#,##0.0000;($#,##0.0000)"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Caption         =   "Coeficiente"
               Height          =   285
               Index           =   2
               Left            =   60
               TabIndex        =   22
               Top             =   1050
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Perc. Dono"
               Height          =   285
               Index           =   1
               Left            =   60
               TabIndex        =   21
               Top             =   750
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Dono"
               Height          =   285
               Index           =   24
               Left            =   60
               TabIndex        =   19
               Top             =   420
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Status"
               Height          =   315
               Index           =   5
               Left            =   60
               TabIndex        =   17
               Top             =   1380
               Width           =   1215
            End
            Begin VB.Label Label5 
               Caption         =   "Número"
               Height          =   195
               Index           =   0
               Left            =   60
               TabIndex        =   16
               Top             =   120
               Width           =   1215
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdMaquina 
         Height          =   4545
         Left            =   -74910
         OleObjectBlob   =   "userSerieInc.frx":0054
         TabIndex        =   6
         Top             =   390
         Width           =   7965
      End
      Begin TrueDBGrid60.TDBGrid grdHistSerie 
         Height          =   4740
         Left            =   -74910
         OleObjectBlob   =   "userSerieInc.frx":5D48
         TabIndex        =   23
         Top             =   390
         Width           =   7950
      End
   End
End
Attribute VB_Name = "frmUserSerieInc"
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

Dim MAQ_COLUNASMATRIZ           As Long
Dim MAQ_LINHASMATRIZ            As Long
Private MAQ_Matriz()            As String

Dim HIST_COLUNASMATRIZ           As Long
Dim HIST_LINHASMATRIZ            As Long
Private HIST_Matriz()            As String

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'Serie
  LimparCampoTexto txtNumero
  LimparCampoCombo cboDono
  LimparCampoMask mskPercDono
  LimparCampoMask mskCoeficiente
  optStatus(0).Value = False
  optStatus(1).Value = False
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserSerieInc.LimparCampos]", _
            Err.Description
End Sub


Private Sub cboDono_LostFocus()
  Pintar_Controle cboDono, tpCorContr_Normal
End Sub


Private Sub cmdAlterar_Click()
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 1
    'Máquina
    If Not IsNumeric(grdMaquina.Columns("PKID").Value & "") Then
      MsgBox "Selecione uma máquina!", vbExclamation, TITULOSISTEMA
      SetarFoco grdMaquina
      Exit Sub
    End If

    frmUserMaquinaInc.lngPKID = grdMaquina.Columns("PKID").Value
    frmUserMaquinaInc.lngSERIEID = lngPKID
    frmUserMaquinaInc.strDescSerie = txtNumero.Text
    frmUserMaquinaInc.Status = tpStatus_Alterar
    frmUserMaquinaInc.Show vbModal

    If frmUserMaquinaInc.blnRetorno Then
      MAQ_MontaMatriz
      grdMaquina.Bookmark = Null
      grdMaquina.ReBind
      grdMaquina.ApproxCount = MAQ_LINHASMATRIZ
    End If
    SetarFoco grdMaquina
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




Private Sub cmdExcluir_Click()
'''  On Error GoTo trata
'''  Dim objContrato     As busSisMaq.clsContrato
'''  Dim objGeral        As busSisMaq.clsGeral
'''  Dim objRs           As ADODB.Recordset
'''  Dim strSql          As String
'''  '
'''  Select Case tabDetalhes.Tab
'''  Case 2 'Exclusão de Contrato
'''    '
'''    If Len(Trim(grdMaquina.Columns("PKID").Value & "")) = 0 Then
'''      MsgBox "Selecione um contrato da empresa.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdMaquina
'''      Exit Sub
'''    End If
'''    '
'''    '
'''    Set objGeral = New busSisMaq.clsGeral
'''    'OBRA
'''    strSql = "SELECT * FROM OBRA WHERE CONTRATOID = " & Formata_Dados(grdMaquina.Columns("PKID").Value, tpDados_Longo)
'''    Set objRs = objGeral.ExecutarSQL(strSql)
'''    If Not objRs.EOF Then
'''      objRs.Close
'''      Set objRs = Nothing
'''      Set objGeral = Nothing
'''      TratarErroPrevisto "Não é possível excluir o contrato, pois existem obras associadas a ele.", "[cmdExcluir_Click]"
'''      Exit Sub
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    'NF
'''    strSql = "SELECT * FROM NF WHERE CONTRATOID = " & Formata_Dados(grdMaquina.Columns("PKID").Value, tpDados_Longo)
'''    Set objRs = objGeral.ExecutarSQL(strSql)
'''    If Not objRs.EOF Then
'''      objRs.Close
'''      Set objRs = Nothing
'''      Set objGeral = Nothing
'''      TratarErroPrevisto "Não é possível excluir o contrato, pois existem NF associadas a ele.", "[cmdExcluir_Click]"
'''      Exit Sub
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    'DEVOLUCAO
'''    strSql = "SELECT * FROM DEVOLUCAO WHERE CONTRATOID = " & Formata_Dados(grdMaquina.Columns("PKID").Value, tpDados_Longo)
'''    Set objRs = objGeral.ExecutarSQL(strSql)
'''    If Not objRs.EOF Then
'''      objRs.Close
'''      Set objRs = Nothing
'''      Set objGeral = Nothing
'''      TratarErroPrevisto "Não é possível excluir o contrato, pois existem devoluções associadas a ele.", "[cmdExcluir_Click]"
'''      Exit Sub
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    'DEVOLUCAO
'''    strSql = "SELECT * FROM BM WHERE CONTRATOID = " & Formata_Dados(grdMaquina.Columns("PKID").Value, tpDados_Longo)
'''    Set objRs = objGeral.ExecutarSQL(strSql)
'''    If Not objRs.EOF Then
'''      objRs.Close
'''      Set objRs = Nothing
'''      Set objGeral = Nothing
'''      TratarErroPrevisto "Não é possível excluir o contrato, pois existem BMs associadas a ele.", "[cmdExcluir_Click]"
'''      Exit Sub
'''    End If
'''    objRs.Close
'''    Set objRs = Nothing
'''    Set objGeral = Nothing
'''    'OK
'''    '
'''    If MsgBox("Confirma exclusão do contrato " & grdMaquina.Columns("Número").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      SetarFoco grdMaquina
'''      Exit Sub
'''    End If
'''    'OK
'''    Set objContrato = New busSisMaq.clsContrato
'''    objContrato.ExcluirContrato CLng(grdMaquina.Columns("PKID").Value)
'''    '
'''    MAQ_MontaMatriz
'''    grdMaquina.Bookmark = Null
'''    grdMaquina.ReBind
'''    grdMaquina.ApproxCount = MAQ_LINHASMATRIZ
'''
'''    Set objContrato = Nothing
'''    SetarFoco grdMaquina
'''  End Select
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
End Sub





Private Sub cmdIncluir_Click()
  On Error GoTo trata
  Dim objForm As Form
  '
  Select Case tabDetalhes.Tab
  Case 1 'Máquina
    frmUserMaquinaInc.Status = tpStatus_Incluir
    frmUserMaquinaInc.lngPKID = 0
    frmUserMaquinaInc.lngSERIEID = lngPKID
    frmUserMaquinaInc.strDescSerie = txtNumero.Text
    frmUserMaquinaInc.Show vbModal

    If frmUserMaquinaInc.blnRetorno Then
      MAQ_MontaMatriz
      grdMaquina.Bookmark = Null
      grdMaquina.ReBind
      grdMaquina.ApproxCount = MAQ_LINHASMATRIZ
    End If
    SetarFoco grdMaquina
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdOk_Click()
  Dim objSerie                  As busSisMaq.clsSerie
  Dim objGeral                  As busSisMaq.clsGeral
  Dim objRS                     As ADODB.Recordset
  Dim strSql                    As String
  Dim lngDONOID                 As Long
  Dim strStatus                 As String
  Dim strPercCasa               As String
  Dim curPercCasa               As Currency
  '
  On Error GoTo trata
  cmdOk.Enabled = False
  If Not ValidaCampos Then
    cmdOk.Enabled = True
    Exit Sub
  End If
  Set objGeral = New busSisMaq.clsGeral
  Set objSerie = New busSisMaq.clsSerie
  'DONO
  lngDONOID = 0
  strSql = "SELECT PKID FROM PESSOA WHERE PESSOA.NOME = " & Formata_Dados(cboDono.Text, tpDados_Texto)
  Set objRS = objGeral.ExecutarSQL(strSql)
  If Not objRS.EOF Then
    lngDONOID = objRS.Fields("PKID").Value
  End If
  objRS.Close
  Set objRS = Nothing
  'Status
  If optStatus(0).Value Then
    strStatus = "A"
  Else
    strStatus = "I"
  End If
  'PercCasa
  curPercCasa = 100 - CCur(mskPercDono.Text)
  strPercCasa = Format(curPercCasa, "###,##0.00")
  'Validar se funcionário já cadastrado
  strSql = "SELECT * FROM SERIE " & _
    " WHERE SERIE.NUMERO = " & Formata_Dados(txtNumero.Text, tpDados_Texto) & _
    " AND SERIE.PKID <> " & Formata_Dados(lngPKID, tpDados_Longo)
  Set objRS = objGeral.ExecutarSQL(strSql)
  If Not objRS.EOF Then
    Pintar_Controle txtNumero, tpCorContr_Erro
    TratarErroPrevisto "Serie já cadastrada"
    objRS.Close
    Set objRS = Nothing
    Set objGeral = Nothing
    Set objSerie = Nothing
    cmdOk.Enabled = True
    SetarFoco txtNumero
    tabDetalhes.Tab = 0
    Exit Sub
  End If
  objRS.Close
  Set objRS = Nothing
  Set objGeral = Nothing
  If Status = tpStatus_Alterar Then
    'Alterar Serie
    objSerie.AlterarSerie lngPKID, _
                          lngDONOID, _
                          txtNumero.Text, _
                          mskPercDono.ClipText, _
                          strPercCasa, _
                          mskCoeficiente.ClipText, _
                          strStatus, _
                          gsNomeUsu
    blnRetorno = True
    blnFechar = True
    Unload Me
    '
  ElseIf Status = tpStatus_Incluir Then
    'Inserir Serie
    objSerie.InserirSerie lngPKID, _
                          lngDONOID, _
                          txtNumero.Text, _
                          mskPercDono.ClipText, _
                          strPercCasa, _
                          mskCoeficiente.ClipText, _
                          strStatus
    blnRetorno = True
    'Selecionar plano cadastrado
    'entrar em modo de alteração
    'lngPKID = objRs.Fields("PKID")
    Status = tpStatus_Alterar
    'Reload na tela
    Form_Load
    'Acerta tabs
    tabDetalhes.Tab = 1
    blnRetorno = True
  End If
  Set objSerie = Nothing
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
  If Not Valida_String(txtNumero, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(cboDono, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o dono" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskPercDono, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o percentual do dono válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(mskCoeficiente, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o coeficiente válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Option(optStatus, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar o status" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(strMsg) = 0 Then
    'Não houve erro válida percentual
    If CCur(mskPercDono.Text) > 100 Then
      strMsg = strMsg & "Preencher o percentual do dono válido" & vbCrLf
      tabDetalhes.Tab = 0
    End If
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserSerieInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserSerieInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    SetarFoco txtNumero
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserSerieInc.Form_Activate]"
End Sub


Private Sub Form_Load()
  On Error GoTo trata
  Dim objRS                   As ADODB.Recordset
  Dim strSql                  As String
  Dim objSerie             As busSisMaq.clsSerie
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
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , cmdIncluir, cmdAlterar
  '
  'Limpar Campos
  LimparCampos
  'Tipo de Convênio
  strSql = "Select PESSOA.NOME FROM PESSOA " & _
      " INNER JOIN DONO ON PESSOA.PKID = DONO.PESSOAID " & _
      " ORDER BY PESSOA.NOME"
  PreencheCombo cboDono, strSql, False, True
  tabDetalhes_Click 1
  If Status = tpStatus_Incluir Then
    '
    optStatus(0).Value = True
    'Visible
    optStatus(0).Visible = False
    optStatus(1).Visible = False
    Label5(5).Visible = False
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = False
    tabDetalhes.TabEnabled(2) = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    Set objSerie = New busSisMaq.clsSerie
    Set objRS = objSerie.SelecionarSeriePeloPkid(lngPKID)
    '
    If Not objRS.EOF Then
      txtNumero.Text = objRS.Fields("NUMERO").Value & ""
      cboDono.Text = objRS.Fields("DESC_DONO").Value & ""
      INCLUIR_VALOR_NO_MASK mskPercDono, objRS.Fields("PERCDONO").Value, TpMaskMoeda
      INCLUIR_VALOR_NO_MASK mskCoeficiente, objRS.Fields("COEFICIENTE").Value, TpMaskMoeda
      If objRS.Fields("STATUS").Value & "" = "A" Then
        optStatus(0).Value = True
        optStatus(1).Value = False
      ElseIf objRS.Fields("STATUS").Value & "" = "I" Then
        optStatus(0).Value = False
        optStatus(1).Value = True
      Else
        optStatus(0).Value = False
        optStatus(1).Value = False
      End If
    End If
    objRS.Close
    Set objRS = Nothing
    '
    Set objSerie = Nothing
    'Visible
    optStatus(0).Visible = True
    optStatus(1).Visible = True
    Label5(5).Visible = True
    '
    tabDetalhes.TabEnabled(0) = True
    tabDetalhes.TabEnabled(1) = True
    If gsNivel <> gsAdmin And gsNivel <> gsDiretor Then
      tabDetalhes.TabEnabled(2) = False
    Else
      tabDetalhes.TabEnabled(2) = True
    End If
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


Private Sub grdHistSerie_UnboundReadDataEx( _
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
  TratarErro Err.Number, Err.Description, "[frmUserSerieInc.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub mskCoeficiente_GotFocus()
  Seleciona_Conteudo_Controle mskCoeficiente
End Sub
Private Sub mskCoeficiente_LostFocus()
  Pintar_Controle mskCoeficiente, tpCorContr_Normal
End Sub

Private Sub mskPercDono_GotFocus()
  Seleciona_Conteudo_Controle mskPercDono
End Sub
Private Sub mskPercDono_LostFocus()
  Pintar_Controle mskPercDono, tpCorContr_Normal
End Sub

Private Sub txtNumero_GotFocus()
  Seleciona_Conteudo_Controle txtNumero
End Sub
Private Sub txtNumero_LostFocus()
  Pintar_Controle txtNumero, tpCorContr_Normal
End Sub


Private Sub tabDetalhes_Click(PreviousTab As Integer)
  '
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    grdMaquina.Enabled = False
    grdHistSerie.Enabled = False
    pictrava(0).Enabled = True
    '
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    '
    SetarFoco txtNumero
  Case 1
    'Máquina
    grdMaquina.Enabled = True
    grdHistSerie.Enabled = False
    pictrava(0).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = True
    cmdIncluir.Enabled = True
    cmdAlterar.Enabled = True
    'Montar RecordSet
    MAQ_COLUNASMATRIZ = grdMaquina.Columns.Count
    MAQ_LINHASMATRIZ = 0
    MAQ_MontaMatriz
    grdMaquina.Bookmark = Null
    grdMaquina.ReBind
    grdMaquina.ApproxCount = MAQ_LINHASMATRIZ
    '
    SetarFoco grdMaquina
  Case 2
    'Histórico
    grdMaquina.Enabled = False
    grdHistSerie.Enabled = True
    pictrava(0).Enabled = False
    '
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    'Montar RecordSet
    HIST_COLUNASMATRIZ = grdHistSerie.Columns.Count
    HIST_LINHASMATRIZ = 0
    HIST_MontaMatriz
    grdHistSerie.Bookmark = Null
    grdHistSerie.ReBind
    grdHistSerie.ApproxCount = HIST_LINHASMATRIZ
    '
    SetarFoco grdHistSerie
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "SisMaq.frmUserSerieInc.tabDetalhes"
  AmpN
End Sub


Private Sub grdMaquina_UnboundReadDataEx( _
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
               Offset + intI, MAQ_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, MAQ_COLUNASMATRIZ, MAQ_LINHASMATRIZ, MAQ_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, MAQ_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserSerieInc.grdGeral_UnboundReadDataEx]"
End Sub


Public Sub HIST_MontaMatriz()
  Dim strSql    As String
  Dim objRS     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMaq.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMaq.clsGeral
  '
  strSql = "SELECT SERIEHIST.SERIEID, SERIEHIST.NUMERO, SERIEHIST.USUARIO, SERIEHIST.DATA, SERIEHIST.PERCDONO, SERIEHIST.PERCCASA, SERIEHIST.COEFICIENTE "
  strSql = strSql & " FROM SERIEHIST " & _
            " WHERE SERIEHIST.SERIEID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
            " ORDER BY SERIEHIST.PKID DESC"
  '
  Set objRS = clsGer.ExecutarSQL(strSql)
  If Not objRS.EOF Then
    HIST_LINHASMATRIZ = objRS.RecordCount
  End If
  If Not objRS.EOF Then
    ReDim HIST_Matriz(0 To HIST_COLUNASMATRIZ - 1, 0 To HIST_LINHASMATRIZ - 1)
  Else
    ReDim HIST_Matriz(0 To HIST_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRS.EOF Then   'se já houver algum item
    For intI = 0 To HIST_LINHASMATRIZ - 1  'varre as linhas
      If Not objRS.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To HIST_COLUNASMATRIZ - 1  'varre as colunas
          HIST_Matriz(intJ, intI) = objRS(intJ) & ""
        Next
        objRS.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Public Sub MAQ_MontaMatriz()
  Dim strSql    As String
  Dim objRS     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMaq.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMaq.clsGeral
  '
  strSql = "SELECT EQUIPAMENTO.PKID, EQUIPAMENTO.NUMERO, MAQUINA.INICIO, EQUIPAMENTO.STATUS, EQUIPAMENTO.COEFICIENTE, TIPO.TIPO " & _
          "FROM EQUIPAMENTO INNER JOIN MAQUINA ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
          "         AND MAQUINA.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
          " INNER JOIN TIPO ON TIPO.PKID = MAQUINA.TIPOID " & _
          "WHERE EQUIPAMENTO.SERIEID = " & Formata_Dados(lngPKID, tpDados_Longo) & _
          " ORDER BY EQUIPAMENTO.NUMERO"

  '
  Set objRS = clsGer.ExecutarSQL(strSql)
  If Not objRS.EOF Then
    MAQ_LINHASMATRIZ = objRS.RecordCount
  End If
  If Not objRS.EOF Then
    ReDim MAQ_Matriz(0 To MAQ_COLUNASMATRIZ - 1, 0 To MAQ_LINHASMATRIZ - 1)
  Else
    ReDim MAQ_Matriz(0 To MAQ_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRS.EOF Then   'se já houver algum item
    For intI = 0 To MAQ_LINHASMATRIZ - 1  'varre as linhas
      If Not objRS.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To MAQ_COLUNASMATRIZ - 1  'varre as colunas
          MAQ_Matriz(intJ, intI) = objRS(intJ) & ""
        Next
        objRS.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub
