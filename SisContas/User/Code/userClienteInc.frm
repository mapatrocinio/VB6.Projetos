VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmUserClienteInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de clientes/cheques"
   ClientHeight    =   6645
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6645
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   6645
      Left            =   8520
      ScaleHeight     =   6645
      ScaleWidth      =   1860
      TabIndex        =   24
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   5505
         Left            =   0
         ScaleHeight     =   5445
         ScaleWidth      =   1605
         TabIndex        =   25
         TabStop         =   0   'False
         Top             =   990
         Width           =   1665
         Begin VB.CommandButton cmdIncluir 
            Caption         =   "&V"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdTransferir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Z"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   2730
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   22
            Top             =   4440
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   21
            Top             =   3570
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   6375
      Left            =   120
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   11245
      _Version        =   393216
      Style           =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados do cliente"
      TabPicture(0)   =   "userClienteInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraProf"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cheques &Compensados"
      TabPicture(1)   =   "userClienteInc.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "grdChqComp"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Cheques Devolvidos/&Recuperado"
      TabPicture(2)   =   "userClienteInc.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdChqDevol"
      Tab(2).ControlCount=   1
      Begin VB.Frame fraProf 
         Height          =   5895
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   7935
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   5535
            Index           =   0
            Left            =   120
            ScaleHeight     =   5535
            ScaleWidth      =   7695
            TabIndex        =   28
            TabStop         =   0   'False
            Top             =   240
            Width           =   7695
            Begin VB.Frame Frame1 
               Caption         =   "Incluir automaticamente após a incluão/alteração do cliente"
               Height          =   495
               Left            =   2880
               TabIndex        =   23
               Top             =   0
               Width           =   4815
               Begin VB.OptionButton optCheque 
                  Caption         =   "Sair"
                  Height          =   195
                  Index           =   2
                  Left            =   3360
                  TabIndex        =   3
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   975
               End
               Begin VB.OptionButton optCheque 
                  Caption         =   "Chq. Devolvido"
                  Height          =   195
                  Index           =   1
                  Left            =   1800
                  TabIndex        =   2
                  TabStop         =   0   'False
                  Top             =   240
                  Width           =   1695
               End
               Begin VB.OptionButton optCheque 
                  Caption         =   "Chq. Compensado"
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   1
                  TabStop         =   0   'False
                  Top             =   240
                  Value           =   -1  'True
                  Width           =   1695
               End
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
               Height          =   5055
               Left            =   0
               TabIndex        =   29
               Top             =   480
               Width           =   7695
               Begin VB.TextBox txtEstado 
                  Height          =   285
                  Left            =   6480
                  MaxLength       =   2
                  MultiLine       =   -1  'True
                  TabIndex        =   11
                  Text            =   "userClienteInc.frx":0054
                  Top             =   2040
                  Width           =   975
               End
               Begin VB.TextBox txtCidade 
                  Height          =   285
                  Left            =   1320
                  MaxLength       =   100
                  MultiLine       =   -1  'True
                  TabIndex        =   10
                  Text            =   "userClienteInc.frx":0060
                  Top             =   2040
                  Width           =   4455
               End
               Begin VB.TextBox txtBairro 
                  Height          =   285
                  Left            =   1320
                  MaxLength       =   100
                  MultiLine       =   -1  'True
                  TabIndex        =   9
                  Text            =   "userClienteInc.frx":006C
                  Top             =   1680
                  Width           =   6135
               End
               Begin VB.TextBox txtTelefone3 
                  Height          =   285
                  Left            =   1320
                  MaxLength       =   20
                  TabIndex        =   14
                  Text            =   "txtTelefone3"
                  Top             =   2760
                  Width           =   2175
               End
               Begin VB.TextBox txtTelefone2 
                  Height          =   285
                  Left            =   5280
                  MaxLength       =   20
                  TabIndex        =   13
                  Text            =   "txtTelefone2"
                  Top             =   2400
                  Width           =   2175
               End
               Begin VB.TextBox txtTelefone1 
                  Height          =   285
                  Left            =   1320
                  MaxLength       =   20
                  TabIndex        =   12
                  Text            =   "txtTelefone1"
                  Top             =   2400
                  Width           =   2175
               End
               Begin VB.TextBox txtEndereco 
                  Height          =   285
                  Left            =   1320
                  MaxLength       =   100
                  MultiLine       =   -1  'True
                  TabIndex        =   8
                  Text            =   "userClienteInc.frx":0078
                  Top             =   1320
                  Width           =   6135
               End
               Begin VB.TextBox txtNome 
                  Height          =   285
                  Left            =   1320
                  MaxLength       =   100
                  TabIndex        =   4
                  Text            =   "txtNome"
                  Top             =   240
                  Width           =   6135
               End
               Begin VB.TextBox txtVeiculo 
                  Height          =   285
                  Left            =   1320
                  MaxLength       =   50
                  TabIndex        =   7
                  Text            =   "txtVeiculo"
                  Top             =   960
                  Width           =   6135
               End
               Begin MSMask.MaskEdBox mskPlaca 
                  Height          =   255
                  Left            =   6480
                  TabIndex        =   6
                  Top             =   600
                  Width           =   975
                  _ExtentX        =   1720
                  _ExtentY        =   450
                  _Version        =   393216
                  MaxLength       =   8
                  Format          =   "000"
                  Mask            =   "???-####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mskData 
                  Height          =   255
                  Index           =   0
                  Left            =   1320
                  TabIndex        =   5
                  Top             =   600
                  Width           =   1335
                  _ExtentX        =   2355
                  _ExtentY        =   450
                  _Version        =   393216
                  AutoTab         =   -1  'True
                  MaxLength       =   10
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin VB.Label Estado 
                  Caption         =   "Estado"
                  Height          =   255
                  Index           =   8
                  Left            =   5880
                  TabIndex        =   41
                  Top             =   2040
                  Width           =   1095
               End
               Begin VB.Label Label6 
                  Caption         =   "Cidade"
                  Height          =   255
                  Index           =   7
                  Left            =   120
                  TabIndex        =   40
                  Top             =   2040
                  Width           =   1095
               End
               Begin VB.Label Label6 
                  Caption         =   "Bairro"
                  Height          =   255
                  Index           =   6
                  Left            =   120
                  TabIndex        =   39
                  Top             =   1680
                  Width           =   1095
               End
               Begin VB.Label Label6 
                  Caption         =   "Telefone 3"
                  Height          =   255
                  Index           =   5
                  Left            =   120
                  TabIndex        =   38
                  Top             =   2760
                  Width           =   1215
               End
               Begin VB.Label Label6 
                  Caption         =   "Telefone 2"
                  Height          =   255
                  Index           =   4
                  Left            =   4080
                  TabIndex        =   37
                  Top             =   2400
                  Width           =   1215
               End
               Begin VB.Label Label6 
                  Caption         =   "Telefone 1"
                  Height          =   255
                  Index           =   1
                  Left            =   120
                  TabIndex        =   36
                  Top             =   2400
                  Width           =   1215
               End
               Begin VB.Label Label2 
                  Caption         =   "Data Nasc."
                  Height          =   255
                  Left            =   120
                  TabIndex        =   35
                  Top             =   600
                  Width           =   1095
               End
               Begin VB.Label Label10 
                  Caption         =   "Placa"
                  Height          =   255
                  Left            =   5880
                  TabIndex        =   34
                  Top             =   600
                  Width           =   615
               End
               Begin VB.Label Label6 
                  Caption         =   "Veículo"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   33
                  Top             =   960
                  Width           =   1215
               End
               Begin VB.Label Label6 
                  Caption         =   "Endereco"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   32
                  Top             =   1320
                  Width           =   1095
               End
               Begin VB.Label Label6 
                  Caption         =   "Nome"
                  Height          =   255
                  Index           =   3
                  Left            =   120
                  TabIndex        =   31
                  Top             =   240
                  Width           =   1215
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
               MaxLength       =   12
               Mask            =   "#########/##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label44 
               Caption         =   "CPF"
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
               TabIndex        =   30
               Top             =   120
               Width           =   735
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdChqComp 
         Height          =   5595
         Left            =   -74880
         OleObjectBlob   =   "userClienteInc.frx":0086
         TabIndex        =   15
         Top             =   480
         Width           =   7935
      End
      Begin TrueDBGrid60.TDBGrid grdChqDevol 
         Height          =   5595
         Left            =   -74880
         OleObjectBlob   =   "userClienteInc.frx":5832
         TabIndex        =   16
         Top             =   480
         Width           =   7935
      End
   End
End
Attribute VB_Name = "frmUserClienteInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public Status                     As tpStatus
Public lngCLIENTEID               As Long
Public bRetorno                   As Boolean
Public bFechar                    As Boolean
Public sTitulo                    As String
Public intQuemChamou              As Integer
Private blnPrimeiraVez            As Boolean

'Variáveis para Grids
Dim CHQCOMP_COLUNASMATRIZ         As Long
Dim CHQCOMP_LINHASMATRIZ          As Long
Private CHQCOMP_Matriz()          As String

Dim CHQDEV_COLUNASMATRIZ          As Long
Dim CHQDEV_LINHASMATRIZ           As Long
Private CHQDEV_Matriz()           As String


Public Sub CHQDEV_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisContas.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisContas.clsGeral
  '
  strSql = "SELECT CHEQUE.PKID, BANCO.NOME, CHEQUE.AGENCIA, CHEQUE.CONTA, CHEQUE.CHEQUE, CHEQUE.VALOR, CHEQUE.DTRECEBIMENTO, CHEQUE.DTRECUPERACAO, CASE CHEQUE.STATUS WHEN 'D' THEN 'DEV' ELSE 'REC' END, CHEQUE.DTDEVOLUCAO, CONVERT(VARCHAR(20), MOTIVODEVOL.CODMOTIVO) + ' - ' + MOTIVODEVOL.DESCMOTIVO "
  strSql = strSql & " FROM (CHEQUE LEFT JOIN BANCO ON BANCO.PKID =  CHEQUE.BANCOID) " & _
          "LEFT JOIN MOTIVODEVOL ON MOTIVODEVOL.PKID = CHEQUE.MOTIVODEVOLID " & _
          "WHERE (CHEQUE.STATUS = " & Formata_Dados("D", tpDados_Texto) & " OR CHEQUE.STATUS = " & Formata_Dados("R", tpDados_Texto) & " ) " & _
          " AND CHEQUE.CLIENTEID = " & Formata_Dados(lngCLIENTEID, tpDados_Longo)
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    CHQDEV_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim CHQDEV_Matriz(0 To CHQDEV_COLUNASMATRIZ - 1, 0 To CHQDEV_LINHASMATRIZ - 1)
  Else
    ReDim CHQDEV_Matriz(0 To CHQDEV_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To CHQDEV_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To CHQDEV_COLUNASMATRIZ - 1  'varre as colunas
          CHQDEV_Matriz(intJ, intI) = objRs(intJ) & ""
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

Public Sub CHQCOMP_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisContas.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisContas.clsGeral
  '
  strSql = "SELECT CHEQUE.PKID, BANCO.NOME, CHEQUE.AGENCIA, CHEQUE.CONTA, CHEQUE.CHEQUE, CHEQUE.VALOR, CHEQUE.DTRECEBIMENTO " & _
          "FROM CHEQUE INNER JOIN BANCO ON BANCO.PKID =  CHEQUE.BANCOID " & _
          "WHERE CHEQUE.STATUS = " & Formata_Dados("C", tpDados_Texto) & _
          " AND CHEQUE.CLIENTEID = " & lngCLIENTEID

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    CHQCOMP_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim CHQCOMP_Matriz(0 To CHQCOMP_COLUNASMATRIZ - 1, 0 To CHQCOMP_LINHASMATRIZ - 1)
  Else
    ReDim CHQCOMP_Matriz(0 To CHQCOMP_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To CHQCOMP_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To CHQCOMP_COLUNASMATRIZ - 1  'varre as colunas
          CHQCOMP_Matriz(intJ, intI) = objRs(intJ) & ""
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



Private Sub cmdAlterar_Click()
  On Error GoTo trata
  '
  Select Case tabDetalhes.Tab
  Case 1 'Inclusão de cheques compensados
    If Not IsNumeric(grdChqComp.Columns("PKID").Value) Then
      TratarErroPrevisto "Selecionar o cheque compensado a ser alterado.", "frmuserClienteInc.cmdAlterar_Click"
      SetarFoco grdChqComp
      Exit Sub
    End If
    frmUserChequeInc.lngCHEQUEID = grdChqComp.Columns("PKID").Value
    frmUserChequeInc.lngCLIENTEID = lngCLIENTEID
    frmUserChequeInc.strStatus = "C"
    frmUserChequeInc.sTitulo = "COMPENSADOS"
    frmUserChequeInc.strCPF = mskCPF.Text
    frmUserChequeInc.Status = tpStatus.tpStatus_Alterar
    frmUserChequeInc.Show vbModal
    If frmUserChequeInc.bRetorno Then
      'Montar RecordSet
      CHQCOMP_COLUNASMATRIZ = 7
      CHQCOMP_LINHASMATRIZ = 0
      CHQCOMP_MontaMatriz
      grdChqComp.Bookmark = Null
      grdChqComp.ReBind
    End If
    SetarFoco grdChqComp
  Case 2 'Inclusão de Cheques devolvidos
  '
    If Not IsNumeric(grdChqDevol.Columns("PKID").Value) Then
      TratarErroPrevisto "Selecionar o cheque compensado a ser alterado.", "frmuserClienteInc.cmdAlterar_Click"
      SetarFoco grdChqDevol
      Exit Sub
    End If
    frmUserChequeInc.lngCHEQUEID = grdChqDevol.Columns("PKID").Value
    frmUserChequeInc.lngCLIENTEID = lngCLIENTEID
    frmUserChequeInc.strStatus = "D"
    frmUserChequeInc.sTitulo = "DEVOLVIDOS"
    frmUserChequeInc.strCPF = mskCPF.Text
    frmUserChequeInc.Status = tpStatus.tpStatus_Alterar
    frmUserChequeInc.Show vbModal
    If frmUserChequeInc.bRetorno Then
      'Montar RecordSet
      CHQDEV_COLUNASMATRIZ = 9
      CHQDEV_LINHASMATRIZ = 0
      CHQDEV_MontaMatriz
      grdChqDevol.Bookmark = Null
      grdChqDevol.ReBind
    End If
    SetarFoco grdChqDevol
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdExcluir_Click()

  Dim objRs                     As ADODB.Recordset
  Dim strSql                    As String
  Dim intI                      As Integer
  Dim strMsg                    As String
  Dim lngCHEQUEID               As Long
  Dim strCHEQUE                 As String
  Dim clsChq                    As busSisContas.clsCheque
  '
  On Error GoTo trata
  '
  Select Case tabDetalhes.Tab
  Case 0 'Exclusão de Cliente
  Case 1, 2 'Exclusão de Cheque Compensados/Devolvidos
    If tabDetalhes.Tab = 1 Then 'Chqs Compensados
      If grdChqComp.Columns("PKID").Value & "" = "" Then
        MsgBox "Selecione um cheque compensado para exclui-lo.", vbExclamation, TITULOSISTEMA
        SetarFoco grdChqComp
        Exit Sub
      End If
      lngCHEQUEID = grdChqComp.Columns("PKID").Value
      strCHEQUE = grdChqComp.Columns("Nro. Chq.").Value
    Else 'Chqs Devolvidos
      If grdChqDevol.Columns("PKID").Value & "" = "" Then
        MsgBox "Selecione um cheque devolvido para exclui-lo.", vbExclamation, TITULOSISTEMA
        SetarFoco grdChqDevol
        Exit Sub
      End If
      lngCHEQUEID = grdChqDevol.Columns("PKID").Value
      strCHEQUE = grdChqDevol.Columns("Nro. Chq.").Value
    End If
    '
    Set clsChq = New busSisContas.clsCheque
    If MsgBox("Deseja excluir o cheque nro. " & strCHEQUE & " ?", vbYesNo, TITULOSISTEMA) = vbYes Then
      clsChq.ExcluirCHEQUE lngCHEQUEID
    End If
    Set clsChq = Nothing
    If tabDetalhes.Tab = 1 Then 'Chqs Compensados
      'Montar RecordSet
      CHQCOMP_COLUNASMATRIZ = 7
      CHQCOMP_LINHASMATRIZ = 0
      CHQCOMP_MontaMatriz
      grdChqComp.Bookmark = Null
      grdChqComp.ReBind
      '
      SetarFoco grdChqComp
      '
    Else 'Chqs Devolvidos
      'Montar RecordSet
      CHQDEV_COLUNASMATRIZ = 9
      CHQDEV_LINHASMATRIZ = 0
      CHQDEV_MontaMatriz
      grdChqDevol.Bookmark = Null
      grdChqDevol.ReBind
      '
      SetarFoco grdChqDevol
    End If
    '
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserClienteInc.cmdExcluir_Click]"
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




Private Sub cmdIncluir_Click()
  On Error GoTo trata
  '
  Select Case tabDetalhes.Tab
  Case 1 'Inclusão de cheques compensados
    frmUserChequeInc.lngCHEQUEID = 0
    frmUserChequeInc.lngCLIENTEID = lngCLIENTEID
    frmUserChequeInc.strStatus = "C"
    frmUserChequeInc.sTitulo = "COMPENSADOS"
    frmUserChequeInc.strCPF = mskCPF.Text
    frmUserChequeInc.Status = tpStatus.tpStatus_Incluir
    frmUserChequeInc.Show vbModal
    If frmUserChequeInc.bRetorno Then
      'Montar RecordSet
      CHQCOMP_COLUNASMATRIZ = 7
      CHQCOMP_LINHASMATRIZ = 0
      CHQCOMP_MontaMatriz
      grdChqComp.Bookmark = Null
      grdChqComp.ReBind
    End If
    SetarFoco grdChqComp
  Case 2 'Inclusão de Cheques devolvidos
    frmUserChequeInc.lngCHEQUEID = 0
    frmUserChequeInc.lngCLIENTEID = lngCLIENTEID
    frmUserChequeInc.strStatus = "D"
    frmUserChequeInc.sTitulo = "DEVOLVIDOS"
    frmUserChequeInc.strCPF = mskCPF.Text
    frmUserChequeInc.Status = tpStatus.tpStatus_Incluir
    frmUserChequeInc.Show vbModal
    If frmUserChequeInc.bRetorno Then
      'Montar RecordSet
      CHQDEV_COLUNASMATRIZ = 9
      CHQDEV_LINHASMATRIZ = 0
      CHQDEV_MontaMatriz
      grdChqDevol.Bookmark = Null
      grdChqDevol.ReBind
    End If
    SetarFoco grdChqDevol
  '
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdOk_Click()
  On Error GoTo trata
  Dim strSql                  As String
  Dim strMsgErro              As String
  Dim objRs                   As ADODB.Recordset
  Dim clsChq                  As busSisContas.clsCheque
  
  '
  Select Case tabDetalhes.Tab
  Case 0 'Inclusão/Alteração de Cliente
    If Not ValidaCamposCliente Then Exit Sub
    'Valida se CPF do cliente já cadastrado
    Set clsChq = New busSisContas.clsCheque
    Set objRs = clsChq.ListarClientePorCPF(mskCPF.Text, lngCLIENTEID)
    If Not objRs.EOF Then
      objRs.Close
      Set objRs = Nothing
      Set clsChq = Nothing
      TratarErroPrevisto "CPF já cadastrado", "cmdOK_Click"
      Exit Sub
    End If
    objRs.Close
    Set objRs = Nothing
    If Status = tpStatus_Alterar Then
      'Código para alteração
      '
      '
      clsChq.AlterarCliente lngCLIENTEID, _
                            mskCPF.Text, _
                            mskPlaca.Text, _
                            txtVeiculo.Text, _
                            txtNome.Text, _
                            txtTelefone1.Text, _
                            txtTelefone2.Text, _
                            txtTelefone3.Text, _
                            txtEndereco.Text, _
                            txtCidade.Text, _
                            txtBairro.Text, _
                            mskData(0).Text, _
                            txtEstado.Text
      bRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Pega Informações para inserir
      '
      clsChq.InserirCliente lngCLIENTEID, _
                            mskCPF.Text, _
                            mskPlaca.Text, _
                            txtVeiculo.Text, _
                            txtNome.Text, _
                            txtTelefone1.Text, _
                            txtTelefone2.Text, _
                            txtTelefone3.Text, _
                            txtEndereco.Text, _
                            txtCidade.Text, _
                            txtBairro.Text, _
                            mskData(0).Text, _
                            txtEstado.Text
      '
      Status = tpStatus_Alterar
      '
      tabDetalhes.TabEnabled(1) = True
      tabDetalhes.TabEnabled(2) = True
      '
      bRetorno = True
    End If
    Set clsChq = Nothing
    
    If optCheque(0).Value Then 'Ir p/ Inclusão de cheque compensado
      tabDetalhes.Tab = 1
      cmdIncluir_Click
    ElseIf optCheque(1).Value Then 'Ir p/ Inclusão de cheque devolvido
      tabDetalhes.Tab = 2
      cmdIncluir_Click
    Else 'Fechar cadastro
      bFechar = True
      Unload Me
    End If

  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Function ValidaCamposCliente() As Boolean
  Dim strMsg     As String
  '
  If Not TestaCPF(mskCPF.Text) Then
    strMsg = strMsg & "Informar o CPF válido" & vbCrLf
    Pintar_Controle mskCPF, tpCorContr_Erro
  End If
  If Len(txtNome.Text) = 0 Then
    strMsg = strMsg & "Informar o Nome" & vbCrLf
    Pintar_Controle txtNome, tpCorContr_Erro
  End If
  If Not Valida_Data(mskData(0), TpNaoObrigatorio) Then
    strMsg = strMsg & "Informar a data de nascimento válida" & vbCrLf
    Pintar_Controle mskData(0), tpCorContr_Erro
  End If
  '
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserChequeInc.ValidaCamposCliente]"
    ValidaCamposCliente = False
  Else
    ValidaCamposCliente = True
  End If
End Function


Private Sub cmdTransferir_Click()
  On Error GoTo trata
  '
  Select Case tabDetalhes.Tab
  Case 1 'Inclusão de cheques compensados
    If Not IsNumeric(grdChqComp.Columns("PKID").Value) Then
      TratarErroPrevisto "Selecionar o cheque compensado a ser transferido.", "frmuserClienteInc.cmdAlterar_Click"
      SetarFoco grdChqComp
      Exit Sub
    End If
    frmUserChequeInc.lngCHEQUEID = grdChqComp.Columns("PKID").Value
    frmUserChequeInc.lngCLIENTEID = lngCLIENTEID
    frmUserChequeInc.strStatus = "D"
    frmUserChequeInc.sTitulo = "DEVOLVIDOS"
    frmUserChequeInc.strCPF = mskCPF.Text
    frmUserChequeInc.Status = tpStatus.tpStatus_Alterar
    frmUserChequeInc.Show vbModal
    If frmUserChequeInc.bRetorno Then
      'Montar RecordSet
      CHQCOMP_COLUNASMATRIZ = 7
      CHQCOMP_LINHASMATRIZ = 0
      CHQCOMP_MontaMatriz
      grdChqComp.Bookmark = Null
      grdChqComp.ReBind
    End If
    SetarFoco grdChqComp
  Case 2 'Inclusão de Cheques devolvidos
  '
    If Not IsNumeric(grdChqDevol.Columns("PKID").Value) Then
      TratarErroPrevisto "Selecionar o cheque compensado a ser alterado.", "frmuserClienteInc.cmdAlterar_Click"
      SetarFoco grdChqDevol
      Exit Sub
    End If
    frmUserChequeInc.lngCHEQUEID = grdChqDevol.Columns("PKID").Value
    frmUserChequeInc.lngCLIENTEID = lngCLIENTEID
    frmUserChequeInc.strStatus = "C"
    frmUserChequeInc.sTitulo = "COMPENSADO"
    frmUserChequeInc.strCPF = mskCPF.Text
    frmUserChequeInc.Status = tpStatus.tpStatus_Alterar
    frmUserChequeInc.Show vbModal
    If frmUserChequeInc.bRetorno Then
      'Montar RecordSet
      CHQDEV_COLUNASMATRIZ = 9
      CHQDEV_LINHASMATRIZ = 0
      CHQDEV_MontaMatriz
      grdChqDevol.Bookmark = Null
      grdChqDevol.ReBind
    End If
    SetarFoco grdChqDevol
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    If Status = tpStatus_Incluir Then
      tabDetalhes.Tab = 0
      SetarFoco mskCPF
    ElseIf Status = tpStatus_Alterar Then
      tabDetalhes.Tab = 0
      SetarFoco mskCPF
    Else 'Status de COnsulta
      tabDetalhes.Tab = 0
      'mskCPF.SetFocus
      frmUserFiltroCheqCons.QuemChamou = 0
      frmUserFiltroCheqCons.Show vbModal
      blnPrimeiraVez = False
      If lngCLIENTEID = 0 Then 'Caso não tenha capturado o Cliente pelo CPF, Saiu
        If Status <> tpStatus_Incluir Then
          bFechar = True
          Unload Me
        Else
          Form_Load
        End If
        
      Else
        Form_Load
      End If
      SetarFoco mskCPF
    End If
    blnPrimeiraVez = False
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserChequeInc.Form_Activate]"
End Sub

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
  strSql = "Select * From Locacao WHERE CPF = '" & mskCPF.Text & "' And PGTOCHEQUE > 0"
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
    frmUserPlanilhaChqsDevolLis.QuemChamou = 1
    frmUserPlanilhaChqsDevolLis.Show vbModal
  End If
  '
  Set clsGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserChequeInc.BuscaChequeEmLoc]"
End Sub



Private Sub grdChqComp_UnboundReadDataEx( _
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
               Offset + intI, CHQCOMP_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, CHQCOMP_COLUNASMATRIZ, CHQCOMP_LINHASMATRIZ, CHQCOMP_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, CHQCOMP_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserChequeInc.grdChqComp_UnboundReadDataEx]"
End Sub





Private Sub grdChqDevol_UnboundReadDataEx( _
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
               Offset + intI, CHQDEV_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, CHQDEV_COLUNASMATRIZ, CHQDEV_LINHASMATRIZ, CHQDEV_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, CHQDEV_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserChequeInc.grdChqDevol_UnboundReadDataEx]"
End Sub

Private Sub mskCPF_Change()
  If Not blnPrimeiraVez Then
    If Len(mskCPF.ClipText) = 11 Then
      If Not TestaCPF(mskCPF.Text) Then
        MsgBox "O número do CPF digitado é inválido !", vbExclamation, TITULOSISTEMA
        Exit Sub
      End If
      '
      tabDetalhes.Tab = 0
      BuscaChequeEmLoc
    End If
  End If
End Sub

Private Sub mskCPF_GotFocus()
  Selecionar_Conteudo mskCPF
End Sub

Private Sub mskCPF_LostFocus()
  Pintar_Controle mskCPF, tpCorContr_Normal
End Sub

Private Sub mskData_GotFocus(Index As Integer)
  Selecionar_Conteudo mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub

Private Sub mskPlaca_GotFocus()
  Selecionar_Conteudo mskPlaca
End Sub

Private Sub txtBairro_GotFocus()
  Selecionar_Conteudo txtBairro
End Sub

Private Sub txtCidade_GotFocus()
  Selecionar_Conteudo txtCidade
End Sub

Private Sub txtEndereco_GotFocus()
  Selecionar_Conteudo txtEndereco
End Sub

Private Sub txtEstado_GotFocus()
  Selecionar_Conteudo txtEstado
End Sub

Private Sub txtNome_GotFocus()
  Selecionar_Conteudo txtNome
End Sub

Private Sub txtNome_LostFocus()
  Pintar_Controle txtNome, tpCorContr_Normal
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
  Me.Height = 7020
  Me.Width = 10470
  CenterForm Me
  blnPrimeiraVez = True
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , cmdIncluir, cmdAlterar
  LerFigurasAvulsas cmdTransferir, "Transferencia.ico", "TransferenciaDown.ico", "Transferencia do status do cheque compensado/Devolvido e vice versa"
  '
  tabDetalhes_Click 0
  If Status = tpStatus_Incluir Then
    'Caso esteja em um evento de Inclusão, Inclui o Pedido
    txtNome.Text = ""
    txtVeiculo.Text = ""
    txtEndereco.Text = ""
    txtBairro.Text = ""
    txtCidade.Text = ""
    txtEstado.Text = ""
    txtTelefone1.Text = ""
    txtTelefone2.Text = ""
    txtTelefone3.Text = ""
    '
    tabDetalhes.TabEnabled(1) = False
    tabDetalhes.TabEnabled(2) = False
    picTrava(0).Enabled = True
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Pega Dados do Banco de dados
    Set clsChq = New busSisContas.clsCheque
    Set objRs = clsChq.ListarCliente(lngCLIENTEID)
    '
    If Not objRs.EOF Then
      mskCPF.Text = objRs.Fields("CPF").Value & ""
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DTNASCIMENTO").Value, TpMaskData
      mskPlaca.Text = IIf(IsNull(objRs.Fields("PLACA").Value), "___-____", objRs.Fields("PLACA").Value)
      txtNome.Text = objRs.Fields("NOME").Value & ""
      txtVeiculo.Text = objRs.Fields("VEICULO").Value & ""
      txtEndereco.Text = objRs.Fields("ENDERECO").Value & ""
      txtBairro.Text = objRs.Fields("BAIRRO").Value & ""
      txtCidade.Text = objRs.Fields("CIDADE").Value & ""
      txtEstado.Text = objRs.Fields("ESTADO").Value & ""
      txtTelefone1.Text = objRs.Fields("TEL1").Value & ""
      txtTelefone2.Text = objRs.Fields("TEL2").Value & ""
      txtTelefone3.Text = objRs.Fields("TEL3").Value & ""
      '
      tabDetalhes.TabEnabled(1) = True
      tabDetalhes.TabEnabled(2) = True
      
    Else 'Status Consultar
      txtNome.Text = ""
      txtVeiculo.Text = ""
      txtEndereco.Text = ""
      txtBairro.Text = ""
      txtCidade.Text = ""
      txtEstado.Text = ""
      txtTelefone1.Text = ""
      txtTelefone2.Text = ""
      txtTelefone3.Text = ""
      '
      tabDetalhes.TabEnabled(1) = False
      tabDetalhes.TabEnabled(2) = False
    
    End If
    If Status = tpStatus_Consultar Then
      picTrava(0).Enabled = False
      If gsNivel = gsRecepcao Then
        tabDetalhes.TabEnabled(0) = False
        tabDetalhes.TabEnabled(1) = False
        tabDetalhes.Tab = 2
        SetarFoco grdChqDevol
      Else
        tabDetalhes.TabEnabled(0) = True
        tabDetalhes.TabEnabled(1) = True
      End If
    Else
      picTrava(0).Enabled = True
      tabDetalhes.TabEnabled(0) = True
      tabDetalhes.TabEnabled(1) = True
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



Private Sub tabDetalhes_Click(PreviousTab As Integer)
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0
    fraProf.Enabled = True
    grdChqComp.Enabled = False
    grdChqDevol.Enabled = False
    '
    cmdCancelar.Enabled = True
    cmdExcluir.Enabled = False
    cmdIncluir.Enabled = False
    cmdAlterar.Enabled = False
    cmdTransferir.Enabled = False
    If Status = tpStatus_Consultar Then
      cmdOk.Enabled = False
    Else
      cmdOk.Enabled = True
    End If
    SetarFoco mskCPF
  Case 1
    fraProf.Enabled = False
    grdChqComp.Enabled = True
    grdChqDevol.Enabled = False
    '
    If Status = tpStatus_Consultar Then
      cmdExcluir.Enabled = False
      cmdIncluir.Enabled = False
      cmdAlterar.Enabled = False
      cmdTransferir.Enabled = False
    Else
      cmdExcluir.Enabled = True
      cmdIncluir.Enabled = True
      cmdAlterar.Enabled = True
      cmdTransferir.Enabled = True
    End If
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    '
    'Montar RecordSet
    CHQCOMP_COLUNASMATRIZ = grdChqComp.Columns.Count
    CHQCOMP_LINHASMATRIZ = 0
    CHQCOMP_MontaMatriz
    grdChqComp.Bookmark = Null
    grdChqComp.ReBind
    SetarFoco grdChqComp
  Case 2
    fraProf.Enabled = False
    grdChqComp.Enabled = False
    grdChqDevol.Enabled = True
    '
    If Status = tpStatus_Consultar Then
      cmdExcluir.Enabled = False
      cmdIncluir.Enabled = False
      cmdAlterar.Enabled = False
      cmdTransferir.Enabled = False
    Else
      cmdExcluir.Enabled = True
      cmdIncluir.Enabled = True
      cmdAlterar.Enabled = True
      cmdTransferir.Enabled = True
    End If
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
    'Montar RecordSet
    'Montar RecordSet
    CHQDEV_COLUNASMATRIZ = grdChqDevol.Columns.Count
    CHQDEV_LINHASMATRIZ = 0
    CHQDEV_MontaMatriz
    grdChqDevol.Bookmark = Null
    grdChqDevol.ReBind
    SetarFoco grdChqDevol
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "SisContas.frmUserChequeInc.tabDetalhes"
  AmpN
End Sub


Private Sub txtTelefone1_GotFocus()
  Selecionar_Conteudo txtTelefone1
End Sub
Private Sub txtTelefone2_GotFocus()
  Selecionar_Conteudo txtTelefone2
End Sub
Private Sub txtTelefone3_GotFocus()
  Selecionar_Conteudo txtTelefone3
End Sub

Private Sub txtVeiculo_GotFocus()
  Selecionar_Conteudo txtVeiculo
End Sub

