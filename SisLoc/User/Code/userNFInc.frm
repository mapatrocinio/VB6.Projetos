VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserNFInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Inclusão de NFSR"
   ClientHeight    =   7590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10380
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   10380
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   7590
      Left            =   8520
      ScaleHeight     =   7590
      ScaleWidth      =   1860
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   4725
         Left            =   90
         ScaleHeight     =   4665
         ScaleWidth      =   1605
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   2790
         Width           =   1665
         Begin VB.CommandButton cmdPagamento 
            Caption         =   "&Z"
            Enabled         =   0   'False
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   1860
            Width           =   1335
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&X"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   16
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   17
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdCancelar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   3600
            Width           =   1335
         End
         Begin VB.CommandButton cmdOk 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   120
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   2730
            Width           =   1335
         End
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   7395
      Left            =   120
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   120
      Width           =   8205
      _ExtentX        =   14473
      _ExtentY        =   13044
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "&Dados da NFSR"
      TabPicture(0)   =   "userNFInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraProf"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      Begin VB.Frame fraProf 
         Height          =   6945
         Left            =   120
         TabIndex        =   24
         Top             =   330
         Width           =   7935
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   6615
            Index           =   0
            Left            =   120
            ScaleHeight     =   6615
            ScaleWidth      =   7785
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   240
            Width           =   7785
            Begin VB.Frame Frame4 
               Height          =   2895
               Left            =   0
               TabIndex        =   40
               Top             =   3450
               Width           =   7695
               Begin TrueDBGrid60.TDBGrid grdPeca 
                  Height          =   2655
                  Left            =   60
                  OleObjectBlob   =   "userNFInc.frx":001C
                  TabIndex        =   15
                  Top             =   150
                  Width           =   7545
               End
            End
            Begin VB.Frame Frame2 
               Height          =   885
               Left            =   0
               TabIndex        =   35
               Top             =   -90
               Width           =   7695
               Begin VB.PictureBox picTrava 
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   645
                  Index           =   3
                  Left            =   90
                  ScaleHeight     =   645
                  ScaleWidth      =   7425
                  TabIndex        =   36
                  TabStop         =   0   'False
                  Top             =   150
                  Width           =   7425
                  Begin VB.TextBox txtNumero 
                     BackColor       =   &H00E0E0E0&
                     Height          =   285
                     Left            =   6000
                     Locked          =   -1  'True
                     TabIndex        =   2
                     TabStop         =   0   'False
                     Text            =   "txtNumero"
                     Top             =   300
                     Width           =   1335
                  End
                  Begin VB.TextBox txtCaixa 
                     BackColor       =   &H00E0E0E0&
                     Height          =   288
                     Left            =   1200
                     Locked          =   -1  'True
                     TabIndex        =   0
                     TabStop         =   0   'False
                     Text            =   "txtCaixa"
                     Top             =   0
                     Width           =   6135
                  End
                  Begin VB.PictureBox Picture2 
                     BorderStyle     =   0  'None
                     Enabled         =   0   'False
                     Height          =   255
                     Left            =   0
                     ScaleHeight     =   255
                     ScaleWidth      =   3045
                     TabIndex        =   37
                     TabStop         =   0   'False
                     Top             =   300
                     Width           =   3045
                     Begin MSMask.MaskEdBox mskData 
                        Height          =   255
                        Index           =   0
                        Left            =   1200
                        TabIndex        =   1
                        TabStop         =   0   'False
                        Top             =   0
                        Width           =   1695
                        _ExtentX        =   2990
                        _ExtentY        =   450
                        _Version        =   393216
                        BackColor       =   14737632
                        AutoTab         =   -1  'True
                        MaxLength       =   16
                        Mask            =   "##/##/#### ##:##"
                        PromptChar      =   "_"
                     End
                     Begin VB.Label Label2 
                        Caption         =   "Data"
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
                        Left            =   0
                        TabIndex        =   38
                        Top             =   0
                        Width           =   615
                     End
                  End
                  Begin VB.Label Label4 
                     Caption         =   "Número"
                     Height          =   255
                     Left            =   4800
                     TabIndex        =   47
                     Top             =   300
                     Width           =   1065
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Funcionário"
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
                     Index           =   5
                     Left            =   0
                     TabIndex        =   39
                     Top             =   0
                     Width           =   1215
                  End
               End
            End
            Begin VB.Frame Frame1 
               Height          =   1155
               Left            =   0
               TabIndex        =   27
               Top             =   2310
               Width           =   7695
               Begin VB.PictureBox picTrava 
                  BorderStyle     =   0  'None
                  Height          =   975
                  Index           =   2
                  Left            =   60
                  ScaleHeight     =   975
                  ScaleWidth      =   7575
                  TabIndex        =   29
                  TabStop         =   0   'False
                  Top             =   150
                  Width           =   7575
                  Begin VB.TextBox txtPeca 
                     Height          =   285
                     Left            =   1260
                     MaxLength       =   100
                     TabIndex        =   9
                     Top             =   0
                     Width           =   6135
                  End
                  Begin VB.TextBox txtPecaFim 
                     BackColor       =   &H00E0E0E0&
                     Height          =   285
                     Left            =   1260
                     Locked          =   -1  'True
                     MaxLength       =   100
                     TabIndex        =   10
                     TabStop         =   0   'False
                     Top             =   300
                     Width           =   6135
                  End
                  Begin MSMask.MaskEdBox mskQuantidade 
                     Height          =   255
                     Left            =   1260
                     TabIndex        =   11
                     Top             =   600
                     Width           =   885
                     _ExtentX        =   1561
                     _ExtentY        =   450
                     _Version        =   393216
                     Format          =   "#,##0;($#,##0)"
                     PromptChar      =   "_"
                  End
                  Begin MSMask.MaskEdBox mskAltura 
                     Height          =   255
                     Left            =   4890
                     TabIndex        =   13
                     Top             =   600
                     Width           =   945
                     _ExtentX        =   1667
                     _ExtentY        =   450
                     _Version        =   393216
                     BackColor       =   -2147483644
                     Format          =   "#,##0.00;($#,##0.00)"
                     PromptChar      =   "_"
                  End
                  Begin MSMask.MaskEdBox mskLargura 
                     Height          =   255
                     Left            =   6480
                     TabIndex        =   14
                     Top             =   600
                     Width           =   945
                     _ExtentX        =   1667
                     _ExtentY        =   450
                     _Version        =   393216
                     BackColor       =   -2147483644
                     Format          =   "#,##0.00;($#,##0.00)"
                     PromptChar      =   "_"
                  End
                  Begin MSMask.MaskEdBox mskValor 
                     Height          =   255
                     Left            =   2670
                     TabIndex        =   12
                     Top             =   600
                     Width           =   1395
                     _ExtentX        =   2461
                     _ExtentY        =   450
                     _Version        =   393216
                     BackColor       =   -2147483644
                     Format          =   "#,##0.00;($#,##0.00)"
                     PromptChar      =   "_"
                  End
                  Begin VB.Label Label7 
                     Caption         =   "Valor"
                     Height          =   225
                     Left            =   2190
                     TabIndex        =   50
                     Top             =   600
                     Width           =   465
                  End
                  Begin VB.Label lblAltura 
                     Caption         =   "Altura"
                     Height          =   225
                     Left            =   4410
                     TabIndex        =   49
                     Top             =   600
                     Width           =   465
                  End
                  Begin VB.Label lblLargura 
                     Caption         =   "Largura"
                     Height          =   225
                     Left            =   5850
                     TabIndex        =   48
                     Top             =   600
                     Width           =   615
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Código da Peça"
                     Height          =   255
                     Index           =   0
                     Left            =   30
                     TabIndex        =   32
                     Top             =   30
                     Width           =   1245
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Descr. da Peça"
                     Height          =   255
                     Index           =   3
                     Left            =   30
                     TabIndex        =   31
                     Top             =   300
                     Width           =   1185
                  End
                  Begin VB.Label Label1 
                     Caption         =   "Quantidade"
                     Height          =   255
                     Index           =   1
                     Left            =   30
                     TabIndex        =   30
                     Top             =   600
                     Width           =   1095
                  End
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
               Height          =   1515
               Left            =   0
               TabIndex        =   26
               Top             =   810
               Width           =   7695
               Begin VB.PictureBox picTrava 
                  BackColor       =   &H00C0C0C0&
                  BorderStyle     =   0  'None
                  Height          =   1245
                  Index           =   1
                  Left            =   120
                  ScaleHeight     =   1245
                  ScaleWidth      =   7455
                  TabIndex        =   33
                  TabStop         =   0   'False
                  Top             =   180
                  Width           =   7455
                  Begin VB.TextBox txtSequencial 
                     BackColor       =   &H00FFFFFF&
                     Height          =   288
                     Left            =   1230
                     TabIndex        =   5
                     Text            =   "txtSequencial"
                     Top             =   270
                     Width           =   1335
                  End
                  Begin VB.ComboBox cboObra 
                     Height          =   315
                     Left            =   1230
                     Style           =   2  'Dropdown List
                     TabIndex        =   6
                     Top             =   570
                     Width           =   6135
                  End
                  Begin VB.TextBox txtEmpresaFim 
                     BackColor       =   &H00E0E0E0&
                     Height          =   288
                     Left            =   3210
                     Locked          =   -1  'True
                     TabIndex        =   8
                     TabStop         =   0   'False
                     Text            =   "txtEmpresaFim"
                     Top             =   900
                     Width           =   4155
                  End
                  Begin VB.TextBox txtContratoFim 
                     BackColor       =   &H00E0E0E0&
                     Height          =   288
                     Left            =   1230
                     Locked          =   -1  'True
                     TabIndex        =   7
                     TabStop         =   0   'False
                     Text            =   "txtContratoFim"
                     Top             =   900
                     Width           =   1965
                  End
                  Begin MSMask.MaskEdBox mskDtSaida 
                     Height          =   255
                     Left            =   1230
                     TabIndex        =   3
                     Top             =   0
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
                  Begin MSMask.MaskEdBox mskDtIniCob 
                     Height          =   255
                     Left            =   6030
                     TabIndex        =   4
                     Top             =   0
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
                  Begin VB.Label Label44 
                     Caption         =   "Número"
                     Height          =   255
                     Left            =   30
                     TabIndex        =   46
                     Top             =   270
                     Width           =   1065
                  End
                  Begin VB.Label Label5 
                     Caption         =   "Obra"
                     Height          =   195
                     Index           =   24
                     Left            =   30
                     TabIndex        =   45
                     Top             =   600
                     Width           =   1095
                  End
                  Begin VB.Label Label3 
                     Caption         =   "Data ini. cobrança"
                     Height          =   225
                     Left            =   4590
                     TabIndex        =   44
                     Top             =   0
                     Width           =   1425
                  End
                  Begin VB.Label Label5 
                     Caption         =   "Data saída"
                     Height          =   225
                     Index           =   0
                     Left            =   30
                     TabIndex        =   43
                     Top             =   0
                     Width           =   855
                  End
                  Begin VB.Label Label6 
                     Caption         =   "Contrato/Empresa"
                     Height          =   255
                     Index           =   4
                     Left            =   30
                     TabIndex        =   34
                     Top             =   900
                     Width           =   1215
                  End
               End
            End
            Begin ComctlLib.StatusBar StatusBar1 
               Height          =   255
               Left            =   5970
               TabIndex        =   28
               Top             =   6330
               Width           =   1740
               _ExtentX        =   3069
               _ExtentY        =   450
               SimpleText      =   ""
               _Version        =   327682
               BeginProperty Panels {0713E89E-850A-101B-AFC0-4210102A8DA7} 
                  NumPanels       =   5
                  BeginProperty Panel1 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
                     Style           =   6
                     Alignment       =   1
                     Bevel           =   2
                     Object.Width           =   1940
                     MinWidth        =   1940
                     TextSave        =   "20/11/2010"
                     Key             =   ""
                     Object.Tag             =   ""
                  EndProperty
                  BeginProperty Panel2 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
                     Style           =   5
                     Alignment       =   1
                     Bevel           =   2
                     Object.Width           =   1235
                     MinWidth        =   1235
                     TextSave        =   "12:40"
                     Key             =   ""
                     Object.Tag             =   ""
                  EndProperty
                  BeginProperty Panel3 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
                     Style           =   1
                     Alignment       =   1
                     Bevel           =   2
                     Enabled         =   0   'False
                     Object.Visible         =   0   'False
                     Object.Width           =   1235
                     MinWidth        =   1235
                     TextSave        =   "CAPS"
                     Key             =   ""
                     Object.Tag             =   ""
                  EndProperty
                  BeginProperty Panel4 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
                     Style           =   2
                     Alignment       =   1
                     Bevel           =   2
                     Enabled         =   0   'False
                     Object.Visible         =   0   'False
                     Object.Width           =   1235
                     MinWidth        =   1235
                     TextSave        =   "NUM"
                     Key             =   ""
                     Object.Tag             =   ""
                  EndProperty
                  BeginProperty Panel5 {0713E89F-850A-101B-AFC0-4210102A8DA7} 
                     Style           =   3
                     Alignment       =   1
                     AutoSize        =   2
                     Bevel           =   2
                     Enabled         =   0   'False
                     Object.Visible         =   0   'False
                     Object.Width           =   1244
                     MinWidth        =   1235
                     TextSave        =   "INS"
                     Key             =   ""
                     Object.Tag             =   ""
                  EndProperty
               EndProperty
            End
            Begin VB.Label lblCor 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFF80&
               Caption         =   "Movimento após o fechamento"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   1
               Left            =   990
               TabIndex        =   42
               Top             =   6360
               Width           =   2625
            End
            Begin VB.Label lblCor 
               Alignment       =   2  'Center
               Caption         =   "Status :"
               BeginProperty Font 
                  Name            =   "Small Fonts"
                  Size            =   6.75
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   165
               Index           =   0
               Left            =   0
               TabIndex        =   41
               Top             =   6360
               Width           =   765
            End
         End
      End
   End
End
Attribute VB_Name = "frmUserNFInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Enum tpIcEstadoNF
  tpIcEstadoNF_Inic = 0
  tpIcEstadoNF_Proc = 1
  tpIcEstadoNF_Con = 2
End Enum

Public Status                 As tpStatus
Public IcEstadoNF             As tpIcEstadoNF
Public lngNFID                As Long
Public blnRetorno             As Boolean
Public blnFechar              As Boolean
Private blnPrimeiraVez        As Boolean
Private strStatusNF           As String
Private strEntrada            As String
'''
'Variáveis para Grid Peça
Dim PECA_COLUNASMATRIZ         As Long
Dim PECA_LINHASMATRIZ          As Long
Private PECA_Matriz()          As String




Public Sub PECA_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim objGer    As busSisLoc.clsGeral
  '
  On Error GoTo trata

  Set objGer = New busSisLoc.clsGeral
  '
  strSql = "SELECT ITEMNF.QUANTIDADE,ISNULL(ITEMNF.QUANTIDADE,0) * ISNULL(ESTOQUE.PESO,0) AS PESO, (ISNULL(ITEMNF.QUANTIDADE,0) * ISNULL(ITEMNF.VALORUNITARIO,0)), ESTOQUE.DESCRICAO + CASE ISNULL(ITEMNF.LARGURA, 0) WHEN 0 THEN '' ELSE ' (' + CONVERT(VARCHAR, ITEMNF.LARGURA) + ')' END + CASE ISNULL(ITEMNF.ALTURA, 0) WHEN 0 THEN '' ELSE ' X (' + CONVERT(VARCHAR, ITEMNF.ALTURA) + ')' END, ITEMNF.PKID, ITEMNF.NFID, ITEMNF.ESTOQUEID " & _
          "FROM ITEMNF INNER JOIN  ESTOQUE ON ITEMNF.ESTOQUEID =  ESTOQUE.PKID " & _
          "WHERE ITEMNF.NFID = " & lngNFID & " " & _
          "ORDER BY ITEMNF.PKID DESC"
  '
  Set objRs = objGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    PECA_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim PECA_Matriz(0 To PECA_COLUNASMATRIZ - 1, 0 To PECA_LINHASMATRIZ - 1)
  Else
    ReDim PECA_Matriz(0 To PECA_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To PECA_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To PECA_COLUNASMATRIZ - 1  'varre as colunas
          PECA_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set objGer = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cboObra_Click()
  On Error GoTo trata
  Dim objObra As busSisLoc.clsObra
  Dim objRs As ADODB.Recordset
  'Alterna para status de alteração/inclusão
  If cboObra.Text = "" Then
    txtContratoFim.Text = ""
    txtEmpresaFim.Text = ""
    '
    Exit Sub
  End If
  Set objObra = New busSisLoc.clsObra
  Set objRs = objObra.CapturaObra(cboObra.Text)
  If objRs.EOF Then
    TratarErroPrevisto "Obra " & cboObra.Text & " não cadastrada!"
    objRs.Close
    Set objRs = Nothing
    Set objObra = Nothing
  Else
    txtContratoFim.Text = objRs.Fields("NUMERO").Value & ""
    txtEmpresaFim.Text = objRs.Fields("EMPRESA_NOME").Value & ""
  End If
  '
  objRs.Close
  Set objRs = Nothing
  Set objObra = Nothing
  'cmdOK_Click
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub cboObra_LostFocus()
  Pintar_Controle cboObra, tpCorContr_Normal
End Sub

Private Sub cmdExcluir_Click()
  Dim objItemNF           As busSisLoc.clsItemNF
  Dim objGer              As busSisLoc.clsGeral
  Dim objRs               As ADODB.Recordset
  Dim strSql              As String
  '
  On Error GoTo trata
  If Len(Trim(grdPeca.Columns("ITEMNFID").Value & "")) = 0 Then
    MsgBox "Selecione uma peça para exclusão da NFSR.", vbExclamation, TITULOSISTEMA
    SetarFoco grdPeca
    Exit Sub
  End If
  '
  Set objGer = New busSisLoc.clsGeral
  '
  If MsgBox("Confirma exclusão da peça " & grdPeca.Columns("Peça").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdPeca
    Exit Sub
  End If
  'OK
  Set objItemNF = New busSisLoc.clsItemNF
  
  objItemNF.AlterarEstoquePeloRetItemNF CLng(grdPeca.Columns("ESTOQUEID").Value), _
                                        grdPeca.Columns("Quantidade").Text
  objItemNF.ExcluirITEMNF CLng(grdPeca.Columns("ITEMNFID").Value)
  'Verifica movimento após o fechamento
  VerificaMovAposFecha lngNFID
  '
  'Montar RecordSet
  PECA_COLUNASMATRIZ = grdPeca.Columns.Count
  PECA_LINHASMATRIZ = 0
  PECA_MontaMatriz
  grdPeca.Bookmark = Null
  grdPeca.ReBind
  SetarFoco txtPeca
  Form_Load
  Set objItemNF = Nothing
  SetarFoco grdPeca
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub



Private Sub cmdPagamento_Click()
'''  Dim objUserContaCorrente  As SisLoc.frmUserContaCorrente
'''  Dim strSql                As String
'''  Dim objRs                 As ADODB.Recordset
'''  Dim objGeral              As busSisLoc.clsGeral
'''  Dim lngTotal              As Long
'''  Dim objNF                 As busSisLoc.clsNF
'''  On Error GoTo trata
'''  'CAPTURA TOTAL
'''  lngTotal = 0
'''  Set objGeral = New busSisLoc.clsGeral
'''  strSql = "SELECT COUNT(PKID) AS TOTAL " & _
'''    " FROM NFPROCEDIMENTO " & _
'''    " WHERE NFPROCEDIMENTO.NFID = " & Formata_Dados(lngNFID, tpDados_Longo)
'''  Set objRs = objGeral.ExecutarSQL(strSql)
'''  If Not objRs.EOF Then
'''    If IsNumeric(objRs.Fields("TOTAL").Value) Then
'''      lngTotal = objRs.Fields("TOTAL").Value
'''    End If
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
'''  Set objGeral = Nothing
'''  If lngTotal = 0 Then
'''    MsgBox "Não há itens lançados nesta NF.", vbExclamation, TITULOSISTEMA
'''    SetarFoco txtPeca
'''    Exit Sub
'''  End If
'''  If strCortesiaNF = "S" And strStatusNF & "" = "I" Then
'''    'Cortesia com status Inicial
'''    MsgBox "Atenção. esta é uma NF de cortesia para funcionário, por isso não há pagamento para ela. Seu status será alterado para [fchado].", vbExclamation, TITULOSISTEMA
'''    Set objNF = New busSisLoc.clsNF
'''    objNF.AlterarStatusNF lngNFID, _
'''                          "F", _
'''                          ""
'''    Set objNF = Nothing
'''    '
'''    Form_Load
'''  Else
'''    '
'''    Set objUserContaCorrente = New frmUserContaCorrente
'''    objUserContaCorrente.lngCCID = 0
'''    objUserContaCorrente.lngNFID = lngNFID
'''    objUserContaCorrente.intGrupo = 0
'''    objUserContaCorrente.strFuncionarioNome = gsNomeUsuCompleto
'''    If Status = tpStatus_Consultar Then
'''      objUserContaCorrente.Status = tpStatus_Consultar
'''    Else
'''      objUserContaCorrente.Status = tpStatus_Incluir
'''    End If
'''    objUserContaCorrente.strStatusLanc = "RC"
'''    objUserContaCorrente.strNivelAcesso = ""
'''    objUserContaCorrente.Show vbModal
'''    If objUserContaCorrente.blnRetorno = True Then
'''      Form_Load
'''    End If
'''    Set objUserContaCorrente = Nothing
'''  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserNFInc.cmdExcluir_Click]"
End Sub

'''
'''Private Sub cmdExcluir_Click()
'''  Dim objRs                     As ADODB.Recordset
'''  Dim strSql                    As String
'''  Dim intI                      As Integer
'''  Dim strMsg                    As String
'''  Dim strMsgErro                As String
'''  '
'''  Dim strQtd                    As String
'''  Dim strValor                  As String
'''  Dim strItem                   As String
'''  Dim strDesc                   As String
'''  Dim clsPed                    As busSisLoc.clsPedido
'''  Dim clsLoc                    As busSisLoc.clsLocacao
'''  Dim clsvend                   As busSisLoc.clsVenda
'''  Dim lngQtdRestanteProd        As Long
'''  Dim curValorUnitarioProd      As Currency
'''  '
'''  On Error GoTo trata
'''  '
'''  Select Case tabDetalhes.Tab
'''  Case 0 'Inclusão/Alteração de Venda
'''  Case 2
'''    If grdPeca.Columns(1).Text = "" Then
'''      MsgBox "Selecione um item da venda para exclui-lo/extorna-lo.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdPeca
'''      Exit Sub
'''    End If
'''    Set clsvend = New busSisLoc.clsVenda
'''    Set clsPed = New busSisLoc.clsPedido
'''    Set clsLoc = New busSisLoc.clsLocacao
'''    'VERIFICAR SE ESSA FUNCIONALIDADE IRÁ ENTRAR DEPOIS
''''''    If Not clsPed.TemPermExcluirPedido(gsNivel, lngNFID, strMsgErro) Then
''''''      MsgBox strMsgErro, vbExclamation, TITULOSISTEMA
''''''      Set clsPed = Nothing
''''''      Set clsLoc = Nothing
''''''      Exit Sub
''''''    End If
'''    If MsgBox("Deseja excluir/estornar o item do cardápio nro. " & grdPeca.Columns(2).Text & " desta venda ?", vbYesNo, TITULOSISTEMA) = vbYes Then
''''''      If intQuemChamou = 1 Then 'Chamada da Exclusão
''''''        '----------------------------
''''''        '----------------------------
''''''        'Pede Senha Superior (Diretor, Gerente ou Administrador
''''''        'Apenas na alteração, na Inclusão não
''''''        If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
''''''          'Só pede senha superior se quem estiver logado não for superior
''''''          frmUserLoginSup.Show vbModal
''''''
''''''          If Len(Trim(gsNomeUsuLib)) = 0 Then
''''''            Set clsvend = Nothing
''''''            Set clsLoc = Nothing
''''''            strMsg = "Para efetuar a exclusão/estorno de itens da venda é necessário a confirmação com senha superior."
''''''            TratarErroPrevisto strMsg, "frmUserVendaInc.cmdExcluir_Click"
''''''            SetarFoco grdPeca
''''''            Exit Sub
''''''          End If
''''''          '
''''''          'Capturou Nome do Usuário, continua processo de Sangria
''''''        Else
''''''          gsNomeUsuLib = gsNomeUsu
''''''        End If
''''''      Else
''''''        gsNomeUsuLib = gsNomeUsu
''''''      End If
'''
'''      'Independente da Chamada ou se motel trabalha ou não com estorno
'''      'Verificar se produto que está sendo estornado já não foi estornado ou se não é o próprio estorno
'''      If Not clsvend.VerificaEstorno(CLng(grdPeca.Columns("TAB_VENDAITEMID").Text), strMsgErro, lngQtdRestanteProd, curValorUnitarioProd) Then
'''        Set clsvend = Nothing
'''        Set clsLoc = Nothing
'''        Set clsPed = Nothing
'''        TratarErroPrevisto strMsgErro, "userVendaInc.cmdExcluir_Click"
'''        SetarFoco grdPeca
'''        Exit Sub
'''      End If
'''      '--------------------------------
'''      '--------------------------------
'''      frmUserTextoExc.QuemChamou = 4 'Chamada de Venda
'''      frmUserTextoExc.QuantidadeASerEstornada = lngQtdRestanteProd
'''      frmUserTextoExc.Show vbModal
'''      If Len(strMotivo & "") = 0 Then
'''        SetarFoco grdPeca
'''        Exit Sub
'''      End If
'''      If grdPeca.Columns(1).Text <> "" Then
'''
'''
'''        'caso trabalhe com estorno de mercadoria,
'''        'pede a quantidade a ser estornada,
'''        'estorna do estoque
'''        'Guarda os Valores para impressão
'''
'''        If gbTrabComEstorno Then
'''          clsvend.EstornaNFPROCEDIMENTO CLng(grdPeca.Columns("TAB_VENDAITEMID").Text), lngNFID, CInt(grdPeca.Columns(0).Text), lngQtdRestanteProdReal
'''        End If
'''        'Carrega variáveis para impressão
'''        strQtd = CStr(Format(lngQtdRestanteProdReal, "###,##0"))
'''        strValor = CStr(Format(curValorUnitarioProd * lngQtdRestanteProdReal, "###,##0.00"))
'''        strItem = grdPeca.Columns(0).Text
'''        strDesc = grdPeca.Columns(2).Text
'''        strDesc = strDesc & " - " & grdPeca.Columns(4).Text
'''        '
'''        clsPed.RetornaProdutoEstoque grdPeca.Columns(2).Text, CStr(IIf(gbTrabComEstorno, lngQtdRestanteProdReal, lngQtdRestanteProd)), 0
'''        '
'''        If Not gbTrabComEstorno Then 'Caso não trabalhe com estorno, deleta
'''          clsvend.ExcluirNFPROCEDIMENTO grdPeca.Columns("TAB_VENDAITEMID").Text
'''          'clsPed.ExcluirPedidoSeVazio lngPEDIDOID
'''        End If
''''''        If intQuemChamou = 1 Then 'Só imprime na alteração
''''''          'Imprimir Exc Venda
''''''          IMP_COMPROV_CANC_VENDA lngNFID, gsNomeEmpresa, 1, strQtd, strValor, strItem, strDesc, strMotivo
''''''        End If
'''        '
'''        blnAlterouItem = True
'''        Set clsPed = Nothing
'''        Set clsLoc = Nothing
'''        Set clsvend = Nothing
'''      End If
'''    End If
'''    'Montar RecordSet
'''    PECA_COLUNASMATRIZ = 8
'''    PECA_LINHASMATRIZ = 0
'''    PECA_MontaMatriz
'''    grdPeca.Bookmark = Null
'''    grdPeca.ReBind
'''    '
'''    SetarFoco grdPeca
'''  End Select
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, "[frmUserVendaInc.cmdExcluir_Click]"
'''End Sub



Private Sub grdPeca_UnboundReadDataEx( _
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
               Offset + intI, PECA_LINHASMATRIZ)

    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For

    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, PECA_COLUNASMATRIZ, PECA_LINHASMATRIZ, PECA_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, PECA_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserNFInc.grdPeca_UnboundReadDataEx]"
End Sub

Private Sub cmdCancelar_Click()
'''  'Rotina para Impressão do Pedido
'''  '------------
'''  'Verificar antes se tem algo a ser impresso
'''  Dim objRs       As ADODB.Recordset
'''  Dim objRsVda    As ADODB.Recordset
'''  Dim strSql      As String
'''  Dim objVenda     As busSisLoc.clsVenda
'''  Dim strCobranca As String
'''  Dim objCC       As busSisLoc.clsContaCorrente
'''  Dim objForm     As SisMotel.frmUserLocContaCorrente
'''  Dim lngCCId     As Long
'''  Dim curVrVenda  As Currency
  '
  On Error GoTo trata
  '
  blnFechar = True
  blnRetorno = True
'''  Set objVenda = New busSisLoc.clsVenda
'''  '
'''  Set objRs = objVenda.ListarNFPROCEDIMENTO(lngNFID)
'''  '
'''  If Not objRs.EOF Then
'''    'Existem Itens de pedido associado ao pedido
'''    'Então imprimir pedido
'''    '
'''    If (intQuemChamou = 0) Or (intQuemChamou = 1 And blnAlterouItem = True) Then 'Chamada original de inclusão
'''      'Antes da impressão
'''      'Novo Preparar para inclusão de pagamento para despesa
'''      If optCobranca(0).Value Then
'''        strCobranca = "S"
'''      Else
'''        strCobranca = "N"
'''      End If
'''      If strCobranca = "S" Then
'''        If chkPgto.Value = False Then
'''          'Pagamento em apenas uma forma de pagamento
'''          'Assumir pagamento em dinheiro
'''          curVrVenda = 0
'''          Set objRsVda = objVenda.ListarVenda(lngNFID)
'''          If Not objRsVda.EOF Then
'''            curVrVenda = objRsVda.Fields("VR_TOT_VENDA").Value
'''          End If
'''          Set objRsVda = Nothing
'''          Set objCC = New busSisLoc.clsContaCorrente
'''          lngCCId = objCC.InserirCC(lngNFID, _
'''                            RetornaCodTurnoCorrente, _
'''                            Format(Now, "DD/MM/YYYY hh:mm"), _
'''                            Format(curVrVenda, "##0.00"), _
'''                            "C", _
'''                            "ES", _
'''                            "VD", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            gsNomeUsu, _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "", _
'''                            "0", _
'''                            "", _
'''                            "")
'''          Set objCC = Nothing
'''        Else
'''          Set objForm = New frmUserLocContaCorrente
'''          objForm.lngLOCDESPVDAEXTID = lngNFID
'''          objForm.intGrupo = 0
'''          objForm.strNumeroAptoPrinc = ""
'''          objForm.Status = tpStatus_Incluir
'''          objForm.strStatusLanc = "VD"
'''          objForm.Show vbModal
'''          Set objForm = Nothing
'''        End If
'''      End If
'''      '
'''      IMP_COMP_VENDA lngNFID, gsNomeEmpresa
'''      '----- Imprimir Impressora Fiscal
'''
'''      If optCobranca(0).Value Then '= "S" Then  'Apenas Imprime Venda Cobradas
'''        If gbTrabComImpFiscal Then
'''          If blnImprimirCupomFiscal = True And intQuemChamou = 0 Then 'imprime cupom fiscal apenas na inclusão da venda
'''            IMP_CUPOM_FISCAL_VENDA lngNFID, gsNomeEmpresa
'''          End If
'''        End If
'''      End If
'''    End If
'''  Else
'''    If MsgBox("Atenção: Não foram lançados itens nesta venda. Caso confirme, a venda não será impressa. Deseja continuar ?", vbYesNo, TITULOSISTEMA) = vbNo Then
'''      objRs.Close
'''      Set objRs = Nothing
'''      blnImprimirCupomFiscal = True
'''      blnFechar = False
'''      Exit Sub
'''    End If
'''  End If
'''  '
'''  objRs.Close
'''  Set objRs = Nothing
'''  '
'''  Set objVenda = Nothing
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  Unload Me
End Sub

Private Sub cmdOK_Click()
  'NF
  Dim objNF                       As busSisLoc.clsNF
  Dim objGeral                    As busSisLoc.clsGeral
  Dim objRs                       As ADODB.Recordset
  Dim strSql                      As String
  Dim lngCONTRATOID     As Long
  Dim lngAno            As Long
  Dim strSequencial     As String
  Dim strNumero         As String
  Dim lngOBRAID         As Long
  '
  'ITENS NF
  Dim objUserEstoqueCons      As SisLoc.frmUserEstoqueCons
  Dim objItemNF               As busSisLoc.clsItemNF
  Dim lngESTOQUEID            As Long
  Dim curVALORESTOQUE         As Currency
  Dim curALTURA               As Currency
  Dim curLARGURA              As Currency
  Dim curALTURALAN            As Currency
  Dim curLARGURALAN           As Currency
  Dim curVALORLAN             As Currency
  Dim strUnidade              As String
  Dim curVALORFINAL           As Currency
  
  On Error GoTo trata
  If IcEstadoNF = tpIcEstadoNF_Inic Then
    'cmdOk.Enabled = False
    If Not ValidaCampos Then
      cmdOk.Enabled = True
      Exit Sub
    End If
    Set objGeral = New busSisLoc.clsGeral
    Set objNF = New busSisLoc.clsNF
    'CONTRATO/OBRA
    lngCONTRATOID = 0
    lngOBRAID = 0
    strSql = "SELECT CONTRATO.PKID AS CONTRATOID, OBRA.PKID AS OBRAID FROM OBRA " & _
          " INNER JOIN CONTRATO ON CONTRATO.PKID = OBRA.CONTRATOID " & _
          " WHERE NUMERO = " & Formata_Dados(txtContratoFim.Text, tpDados_Texto) & _
          " AND OBRA.DESCRICAO = " & Formata_Dados(cboObra.Text, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngCONTRATOID = objRs.Fields("CONTRATOID").Value
      lngOBRAID = objRs.Fields("OBRAID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    If lngCONTRATOID = 0 Then
      Pintar_Controle cboObra, tpCorContr_Erro
      TratarErroPrevisto "Contrato/Obra não cadastrado"
      Set objGeral = Nothing
      Set objNF = Nothing
      cmdOk.Enabled = True
      SetarFoco cboObra
      Exit Sub
    End If
    '
    If Status = tpStatus_Alterar Then
      'Alterar NF
      objNF.AlterarNF lngNFID, _
                      lngCONTRATOID, _
                      mskDtSaida.Text, _
                      mskDtIniCob.Text, _
                      "", _
                      lngOBRAID
      'Verifica MOV
      VerificaMovAposFecha lngNFID
      '
    ElseIf Status = tpStatus_Incluir Then
      'Inserir NF
      lngAno = Right(mskDtSaida.Text, 4)
      'NÚMEOR/DATA SAÍDA
      strSql = "SELECT * FROM NF " & _
            " WHERE NF.SEQUENCIAL = " & Formata_Dados(txtSequencial.Text, tpDados_Longo) & _
            " AND NF.ANO = " & Formata_Dados(lngAno, tpDados_Longo) & _
            " AND NF.PKID <> " & Formata_Dados(lngNFID, tpDados_Longo)
      Set objRs = objGeral.ExecutarSQL(strSql)
      If Not objRs.EOF Then
        objRs.Close
        Set objRs = Nothing
        Set objGeral = Nothing
        Pintar_Controle cboObra, tpCorContr_Erro
        TratarErroPrevisto "Número da NF já cadastrado para o ano informado"
        Set objGeral = Nothing
        Set objNF = Nothing
        cmdOk.Enabled = True
        SetarFoco cboObra
        Exit Sub
      End If
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      'Sequencial
      'strSequencial = RetornaGravaCampoSequencialNF("SEQUENCIAL", lngAno) & ""
      strSequencial = txtSequencial.Text
      strNumero = Format(strSequencial, "0000") & "/" & lngAno
      objNF.InserirNF lngNFID, _
                      lngCONTRATOID, _
                      strNumero, _
                      strSequencial & "", _
                      lngAno & "", _
                      Format(Now, "DD/MM/YYYY hh:mm"), _
                      mskDtSaida.Text, _
                      mskDtIniCob.Text, _
                      "", _
                      lngOBRAID
  
    End If
    'Verificação
    If Status = tpStatus_Alterar Then
      'Selecionar contrato pelo nome
      Status = tpStatus_Alterar
      IcEstadoNF = tpIcEstadoNF_Proc
      'Reload na tela
      Form_Load
      'Acerta tabs
      blnRetorno = True
    ElseIf Status = tpStatus_Incluir Then
      'Selecionar contrato pelo nome
      Status = tpStatus_Alterar
      IcEstadoNF = tpIcEstadoNF_Proc
      'Reload na tela
      Form_Load
      'Acerta tabs
      blnRetorno = True
    End If
    'cmdOk.Enabled = True
    SetarFoco txtPeca
    Set objNF = Nothing
  ElseIf IcEstadoNF = tpIcEstadoNF_Proc Then
    'Ítens da NF
    If Not ValidaCamposEstoque Then
      Exit Sub
    End If
  
    Set objGeral = New busSisLoc.clsGeral
    '
    'ESTOQUE
    lngESTOQUEID = 0
    strSql = "SELECT ESTOQUE.PKID, ESTOQUE.LARGURA, ESTOQUE.ALTURA, ESTOQUE.VALOR, UNIDADE.UNIDADE FROM ESTOQUE INNER JOIN UNIDADE ON UNIDADE.PKID = ESTOQUE.UNIDADEID " & _
      "WHERE ESTOQUE.DESCRICAO = " & Formata_Dados(txtPecaFim.Text, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngESTOQUEID = objRs.Fields("PKID").Value
      curVALORESTOQUE = IIf(IsNull(objRs.Fields("VALOR").Value), 0, objRs.Fields("VALOR").Value)
      strUnidade = objRs.Fields("UNIDADE").Value & ""
      curALTURA = IIf(Not IsNumeric(objRs.Fields("ALTURA").Value), 0, objRs.Fields("ALTURA").Value)
      curLARGURA = IIf(Not IsNumeric(objRs.Fields("LARGURA").Value), 0, objRs.Fields("LARGURA").Value)
    End If
    objRs.Close
    Set objRs = Nothing
    '
    If lngESTOQUEID = 0 Then
      TratarErroPrevisto "Peça não cadastrada."
      Pintar_Controle txtPeca, tpCorContr_Erro
      SetarFoco txtPeca
      Exit Sub
    End If
    '
    Set objGeral = Nothing
    'Calculo do Valor a ser cobrado depende da unidade
    If Not ValidaCamposEstoqueDim(strUnidade) Then
      Exit Sub
    End If
    '
    curALTURALAN = 0
    curLARGURALAN = 0
    curVALORLAN = 0
    If mskValor.ClipText <> "" Then
      curVALORLAN = CCur(mskValor.ClipText)
      curVALORESTOQUE = curVALORLAN
    End If
    If mskAltura.ClipText <> "" Then
      curALTURALAN = CCur(mskAltura.ClipText)
    End If
    If mskLargura.ClipText <> "" Then
      curLARGURALAN = CCur(mskLargura.ClipText)
    End If
    'Verifica valor estoque
    
    Select Case strUnidade
    Case RectpIcUnidade.tpIcUnidade_M2
      'M2
      If curLARGURALAN <> 0 And curALTURALAN <> 0 Then
        curVALORFINAL = (curLARGURALAN * curALTURALAN / 10000) * curVALORESTOQUE
      Else
        curVALORFINAL = (curLARGURA * curALTURA / 10000) * curVALORESTOQUE
      End If
      '
    Case RectpIcUnidade.tpIcUnidade_MLINEAR
      'Mlinear
      If curLARGURALAN <> 0 Then
        curVALORFINAL = (curLARGURALAN / 100) * curVALORESTOQUE
      Else
        curVALORFINAL = (curLARGURA / 100) * curVALORESTOQUE
      End If
      '
    Case RectpIcUnidade.tpIcUnidade_UNID
      'Unidade
      curVALORFINAL = curVALORESTOQUE
      '
    End Select
    'Inclusão de proceidmentos
    Set objItemNF = New busSisLoc.clsItemNF
    '
    'Baixa estoque
    objItemNF.AlterarEstoquePeloItemNF lngESTOQUEID, _
                                       mskQuantidade.Text
    
    objItemNF.InserirITEMNF lngNFID, _
                            lngESTOQUEID, _
                            mskQuantidade.Text, _
                            Format(curVALORFINAL, "###,##0.0000") & "", _
                            Format(curVALORESTOQUE, "###,##0.0000") & "", _
                            mskLargura.Text, _
                            mskAltura.Text, _
                            mskValor.Text
    '
    Set objItemNF = Nothing
    VerificaMovAposFecha lngNFID
    'cmdOk.Default = True
    'Novo procedimento
    txtPeca.Text = ""
    txtPecaFim.Text = ""
    LimparCampoMask mskQuantidade
    LimparCampoMask mskValor
    LimparCampoMask mskLargura
    LimparCampoMask mskAltura
    '
    'Montar RecordSet
    PECA_COLUNASMATRIZ = grdPeca.Columns.Count
    PECA_LINHASMATRIZ = 0
    PECA_MontaMatriz
    grdPeca.Bookmark = Null
    grdPeca.ReBind
    SetarFoco txtPeca
  
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  cmdOk.Enabled = True
End Sub






'''
'''Private Sub mskQuantidade_KeyUp(KeyCode As Integer, Shift As Integer)
'''  On Error GoTo trata
'''  Dim objUserEstoqueCons      As SisLoc.frmUserEstoqueCons
'''  Dim objItemNF               As busSisLoc.clsItemNF
'''  Dim objGeral                As busSisLoc.clsGeral
'''  Dim objRs                   As ADODB.Recordset
'''  Dim strSql                  As String
'''  Dim lngESTOQUEID            As Long
'''  Dim curVALORESTOQUE         As Currency
'''  Dim curALTURA               As Currency
'''  Dim curLARGURA              As Currency
'''  Dim strUnidade              As String
'''  '
'''  Dim curVALORFINAL           As Currency
'''  '
'''  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub
'''  If KeyCode = 13 Then Exit Sub
'''  If Len(mskQuantidade.Text) < 2 Then Exit Sub
'''  '
'''  Pintar_Controle txtPeca, tpCorContr_Normal
'''  If Not ValidaCamposEstoque Then
'''    Exit Sub
'''  End If
'''
'''  Set objGeral = New busSisLoc.clsGeral
'''  '
'''  'ESTOQUE
'''  lngESTOQUEID = 0
'''  strSql = "SELECT ESTOQUE.PKID, ESTOQUE.LARGURA, ESTOQUE.ALTURA, ESTOQUE.VALOR, UNIDADE.UNIDADE FROM ESTOQUE INNER JOIN UNIDADE ON UNIDADE.PKID = ESTOQUE.UNIDADEID " & _
'''    "WHERE ESTOQUE.DESCRICAO = " & Formata_Dados(txtPecaFim.Text, tpDados_Texto)
'''  Set objRs = objGeral.ExecutarSQL(strSql)
'''  If Not objRs.EOF Then
'''    lngESTOQUEID = objRs.Fields("PKID").Value
'''    curVALORESTOQUE = IIf(IsNull(objRs.Fields("VALOR").Value), 0, objRs.Fields("VALOR").Value)
'''    strUnidade = objRs.Fields("UNIDADE").Value & ""
'''    curALTURA = IIf(Not IsNumeric(objRs.Fields("ALTURA").Value), 0, objRs.Fields("ALTURA").Value)
'''    curLARGURA = IIf(Not IsNumeric(objRs.Fields("LARGURA").Value), 0, objRs.Fields("LARGURA").Value)
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
'''  '
'''  If lngESTOQUEID = 0 Then
'''    TratarErroPrevisto "Peça não cadastrada."
'''    Pintar_Controle txtPeca, tpCorContr_Erro
'''    SetarFoco txtPeca
'''    Exit Sub
'''  End If
'''  '
'''  Set objGeral = Nothing
'''  'Calculo do Valor a ser cobrado depende da unidade
'''  Select Case strUnidade
'''  Case RectpIcUnidade.tpIcUnidade_M2
'''    'M2
'''    If curLARGURA = 0 Then
'''      curVALORFINAL = (curALTURA) * curVALORESTOQUE / 10000
'''    Else
'''      curVALORFINAL = (curLARGURA * curALTURA / 10000) * curVALORESTOQUE
'''    End If
'''    '
'''  Case RectpIcUnidade.tpIcUnidade_MLINEAR
'''    'Mlinear
'''    curVALORFINAL = (curLARGURA) * curVALORESTOQUE
'''    '
'''  Case RectpIcUnidade.tpIcUnidade_UNID
'''    'Unidade
'''    curVALORFINAL = curVALORESTOQUE
'''    '
'''  End Select
'''  'Inclusão de proceidmentos
'''  Set objItemNF = New busSisLoc.clsItemNF
'''  '
'''  'Baixa estoque
'''  objItemNF.AlterarEstoquePeloItemNF lngESTOQUEID, _
'''                                     mskQuantidade.Text
'''
'''  objItemNF.InserirITEMNF lngNFID, _
'''                          lngESTOQUEID, _
'''                          mskQuantidade.Text, _
'''                          Format(curVALORFINAL, "###,##0.00") & "", _
'''                          Format(curVALORESTOQUE, "###,##0.00") & ""

Private Sub mskValor_GotFocus()
  Selecionar_Conteudo mskValor
End Sub
Private Sub mskValor_LostFocus()
  Pintar_Controle mskValor, tpCorContr_Normal
End Sub

Private Sub mskAltura_GotFocus()
  Selecionar_Conteudo mskAltura
End Sub
Private Sub mskAltura_LostFocus()
  Pintar_Controle mskAltura, tpCorContr_Normal
End Sub

Private Sub mskDtIniCob_GotFocus()
  Selecionar_Conteudo mskDtIniCob
End Sub
Private Sub mskDtIniCob_LostFocus()
  Pintar_Controle mskDtIniCob, tpCorContr_Normal
End Sub
Private Sub mskDtSaida_GotFocus()
  Selecionar_Conteudo mskDtSaida
End Sub
Private Sub mskDtSaida_LostFocus()
  Pintar_Controle mskDtSaida, tpCorContr_Normal
End Sub
Private Sub mskLargura_GotFocus()
  Selecionar_Conteudo mskLargura
End Sub
Private Sub mskLargura_LostFocus()
  Pintar_Controle mskLargura, tpCorContr_Normal
End Sub

'''  '
'''  Set objItemNF = Nothing
'''  VerificaMovAposFecha lngNFID
'''  'cmdOk.Default = True
'''  'Novo procedimento
'''  txtPeca.Text = ""
'''  txtPecaFim.Text = ""
'''  LimparCampoMask mskQuantidade
'''  '
'''  'Montar RecordSet
'''  PECA_COLUNASMATRIZ = grdPeca.Columns.Count
'''  PECA_LINHASMATRIZ = 0
'''  PECA_MontaMatriz
'''  grdPeca.Bookmark = Null
'''  grdPeca.ReBind
'''  SetarFoco txtPeca
'''  '
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub
Private Sub mskQuantidade_GotFocus()
  Selecionar_Conteudo mskQuantidade
End Sub
Private Sub mskQuantidade_LostFocus()
  Pintar_Controle mskQuantidade, tpCorContr_Normal
End Sub


Private Sub txtPeca_GotFocus()
  Selecionar_Conteudo txtPeca
End Sub


Private Sub Form_Activate()
  On Error GoTo trata
  If blnPrimeiraVez Then
    'Seta foco no grid
    'Montar RecordSet
    PECA_COLUNASMATRIZ = grdPeca.Columns.Count
    PECA_LINHASMATRIZ = 0
    PECA_MontaMatriz
    grdPeca.Bookmark = Null
    grdPeca.ReBind
    '
    tabDetalhes.Tab = 0
    blnPrimeiraVez = False
    SetarFoco mskDtSaida
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserNFInc.Form_Activate]"
End Sub

Private Sub txtPeca_KeyPress(KeyAscii As Integer)
  On Error GoTo trata
  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub txtPeca_LostFocus()
  On Error GoTo trata
  Dim objUserEstoqueCons As SisLoc.frmUserEstoqueCons
  Dim objEstoque              As busSisLoc.clsEstoque
  Dim objGeral                As busSisLoc.clsGeral
  Dim objRs                   As ADODB.Recordset
  Dim strSql                  As String
  Dim lngESTOQUEID            As Long
  '
  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub
  If Me.ActiveControl.Name = "grdPeca" Then Exit Sub
  If Me.ActiveControl.Name = "cmdPagamento" Then Exit Sub
  If Me.ActiveControl.Name = "cmdExcluir" Then Exit Sub
  If Me.ActiveControl.Name = "cmdImprimir" Then Exit Sub
  '
  Pintar_Controle txtPeca, tpCorContr_Normal
  If Len(txtPeca.Text) = 0 Then
    TratarErroPrevisto "Entre com a descrição da peça."
    Pintar_Controle txtPeca, tpCorContr_Erro
    SetarFoco txtPeca
    Exit Sub
  End If
  Set objEstoque = New busSisLoc.clsEstoque
  Set objGeral = New busSisLoc.clsGeral
  '
  Set objRs = objEstoque.CapturaEstoquePeloCodigo(txtPeca.Text)
  If objRs.EOF Then
    'Novo : apresentar tela para seleção do contrato
    Set objUserEstoqueCons = New SisLoc.frmUserEstoqueCons
    objUserEstoqueCons.strCodigoEstoque = txtPeca.Text
    'objUserEstoqueCons.lngESTOQUEID = lngESTOQUEID
    objUserEstoqueCons.QuemChamou = 1
    objUserEstoqueCons.Show vbModal

    If txtPecaFim.Text = "" Then
      txtPeca.Text = ""
      txtPecaFim.Text = ""
      TratarErroPrevisto "Selecionar uma peça"
      Pintar_Controle txtPeca, tpCorContr_Erro
      SetarFoco txtPeca
      Exit Sub
    Else
      'SetarFoco mskQuantidade
    End If
    Set objUserEstoqueCons = Nothing
  Else
    If objRs.RecordCount = 1 Then
      txtPecaFim = objRs.Fields("DESCRICAO").Value & ""
    Else
      'Novo : apresentar tela para seleção do contrato
      Set objUserEstoqueCons = New frmUserEstoqueCons
      objUserEstoqueCons.strCodigoEstoque = txtPeca.Text
      objUserEstoqueCons.QuemChamou = 1
      objUserEstoqueCons.Show vbModal

      If txtPecaFim.Text = "" Then
        txtPeca.Text = ""
        txtPecaFim.Text = ""
        TratarErroPrevisto "Preencher o código da peça"
        Pintar_Controle txtPeca, tpCorContr_Erro
        SetarFoco txtPeca
        Exit Sub
      Else
        'SetarFoco mskQuantidade
        'Tratar Valor
      End If
      Set objUserEstoqueCons = Nothing
    End If
  End If
  '
  objRs.Close
  Set objRs = Nothing
  Set objEstoque = Nothing
  'cmdOk.Default = True

  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Load()
On Error GoTo trata
  Dim objRs     As ADODB.Recordset
  Dim strSql    As String
  Dim objNF As busSisLoc.clsNF
  '
  blnPrimeiraVez = True
  blnFechar = False
  blnRetorno = False
  AmpS
  Me.Height = 8070
  Me.Width = 10470
  CenterForm Me
  '
  LerFiguras Me, tpBmp_Vazio, cmdOk, , cmdCancelar, cmdExcluir, , , , , cmdImprimir
  LerFigurasAvulsas cmdPagamento, "InfFinanc.ico", "InfFinancDown.ico", "Recebimento"
  '
  'Limpar Campos
  LimparCampos
  'Tratar campos
  TratarCampos
  '
  'Obra
  strSql = "Select DESCRICAO from OBRA ORDER BY DESCRICAO"
  PreencheCombo cboObra, strSql, False, True
  If Status = tpStatus_Incluir Then
    '
    lblCor(0).Visible = False
    lblCor(1).Visible = False
    txtSequencial.Locked = False
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    '-----------------------------
    'NF
    '------------------------------
    lblCor(0).Visible = True
    lblCor(1).Visible = True
    txtSequencial.Locked = True
    '
    Set objNF = New busSisLoc.clsNF
    Set objRs = objNF.SelecionarNFPeloPkid(lngNFID)
    '
    If Not objRs.EOF Then
      'NF
      strStatusNF = objRs.Fields("STATUS").Value & ""
      'Cabeçalho
      txtSequencial.Text = objRs.Fields("SEQUENCIAL").Value & ""
      txtNumero.Text = objRs.Fields("NUMERO").Value & ""
      INCLUIR_VALOR_NO_MASK mskData(0), objRs.Fields("DATA").Value, TpMaskData
      txtCaixa.Text = gsNomeUsuCompleto
      'NF
      INCLUIR_VALOR_NO_MASK mskDtSaida, objRs.Fields("DTSAIDA").Value, TpMaskData
      INCLUIR_VALOR_NO_MASK mskDtIniCob, objRs.Fields("DTINICIOCOB").Value, TpMaskData
      'txtContrato.Text = objRs.Fields("CONTRATO_NUMERO").Value & ""
      'txtNrRF.Text = objRs.Fields("NUMERORF").Value & ""
      txtContratoFim.Text = objRs.Fields("CONTRATO_NUMERO").Value & ""
      txtEmpresaFim.Text = objRs.Fields("NOME_EMPRESA").Value & ""
      cboObra.Text = objRs.Fields("DESC_OBRA").Value & ""
      '
      TratarStatus objRs.Fields("STATUS").Value & "", _
                   lblCor(1)
    End If
    objRs.Close
    Set objRs = Nothing
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

Private Sub LimparCampos()
  Dim sMask As String

  On Error GoTo trata
  'NF
  LimparCampoTexto txtCaixa
  LimparCampoTexto txtSequencial
  LimparCampoTexto txtNumero
  LimparCampoMask mskData(0)
  '
  'LimparCampoTexto txtContrato
  LimparCampoTexto txtContratoFim
  LimparCampoTexto txtEmpresaFim
  LimparCampoCombo cboObra

  'NFPRESTADOR
  LimparCampoTexto txtPeca
  LimparCampoTexto txtPecaFim
  LimparCampoMask mskQuantidade
  LimparCampoMask mskValor
  LimparCampoMask mskAltura
  LimparCampoMask mskLargura
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserNFInc.LimparCampos]", _
            Err.Description
End Sub

Private Sub TratarEstadoNF()
  On Error GoTo trata
  'Propósito : Tratar estado da NF
  If IcEstadoNF = tpIcEstadoNF_Inic Then
    picTrava(1).Enabled = True
    picTrava(2).Enabled = False
    picTrava(3).Enabled = False
    grdPeca.Enabled = False
    '
    cmdImprimir.Enabled = False
    cmdExcluir.Enabled = False
    cmdPagamento.Enabled = False
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
    blnPrimeiraVez = True
  ElseIf IcEstadoNF = tpIcEstadoNF_Proc Then
    picTrava(1).Enabled = False
    picTrava(2).Enabled = True
    picTrava(3).Enabled = False
    grdPeca.Enabled = True
    '
    cmdImprimir.Enabled = False
    cmdExcluir.Enabled = True
    cmdPagamento.Enabled = False
    cmdOk.Enabled = True
    cmdCancelar.Enabled = True
  ElseIf IcEstadoNF = tpIcEstadoNF_Con Then
    picTrava(1).Enabled = False
    picTrava(2).Enabled = False
    picTrava(3).Enabled = False
    grdPeca.Enabled = True
    '
    cmdImprimir.Enabled = False
    cmdExcluir.Enabled = False
    cmdPagamento.Enabled = False
    cmdOk.Enabled = False
    cmdCancelar.Enabled = True
  End If
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserNFInc.TratarEstadoNF]", _
            Err.Description
End Sub


Private Sub TratarCampos()
  On Error GoTo trata
  'NFPRESTADOR
  '
  TratarEstadoNF
  '
  If Status = tpStatus_Incluir Then
    'Trtar exclusão
    '
    txtCaixa.Text = gsNomeUsuCompleto
  ElseIf Status = tpStatus_Alterar Or Status = tpStatus_Consultar Then
    'Visible
  End If
  '
  Exit Sub
trata:
  Err.Raise Err.Number, _
            "[frmUserNFInc.TratarCampos]", _
            Err.Description
End Sub

'''Private Sub txtContrato_GotFocus()
'''  Selecionar_Conteudo txtContrato
'''End Sub
'''
'''Private Sub txtContrato_KeyPress(KeyAscii As Integer)
'''  On Error GoTo trata
'''  KeyAscii = TRANSFORMA_MAIUSCULA(KeyAscii)
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub
'''
'''Private Sub txtContrato_LostFocus()
'''  On Error GoTo trata
'''  Dim objUserContratoCons As SisLoc.frmUserContratoCons
'''  Dim objContrato     As busSisLoc.clsContrato
'''  Dim objRs     As ADODB.Recordset
'''  If Me.ActiveControl.Name = "cmdCancelar" Then Exit Sub
'''  If Me.ActiveControl.Name = "txtPrestEspec" Then Exit Sub
'''  If Me.ActiveControl.Name = "txtContrato" Then Exit Sub
'''  If Me.ActiveControl.Name = "txtNrRF" Then Exit Sub
'''  If Me.ActiveControl.Name = "mskDtIniCob" Then Exit Sub
'''  If Me.ActiveControl.Name = "mskDtSaida" Then Exit Sub
'''  If Me.ActiveControl.Name = "cboObra" Then Exit Sub
'''
'''  Pintar_Controle txtContrato, tpCorContr_Normal
'''  If Len(txtContrato.Text) = 0 Then
'''    TratarErroPrevisto "Entre com o contrato."
'''    Pintar_Controle txtContrato, tpCorContr_Erro
'''    SetarFoco txtContrato
'''    Exit Sub
'''  End If
'''  If Len(txtNrRF.Text) = 0 Then
'''    TratarErroPrevisto "Entre com o número Rio Formas."
'''    Pintar_Controle txtNrRF, tpCorContr_Erro
'''    SetarFoco txtNrRF
'''    Exit Sub
'''  End If
'''  Set objContrato = New busSisLoc.clsContrato
'''  '
'''  Set objRs = objContrato.CapturaContrato(txtContrato.Text, _
'''                                          "")
'''  If objRs.EOF Then
'''    'Novo : apresentar tela para seleção do contrato
'''    Set objUserContratoCons = New frmUserContratoCons
'''    objUserContratoCons.strContrato = txtContrato.Text
'''    objUserContratoCons.Show vbModal
'''
'''    If objUserContratoCons.strContrato = "" Then
'''      txtContrato.Text = ""
'''      TratarErroPrevisto "Selecione um contrato"
'''      Pintar_Controle txtContrato, tpCorContr_Erro
'''      SetarFoco txtContrato
'''      Exit Sub
'''    Else
'''      'Cadastrar NF
'''      CadastrarNF
'''    End If
'''    Set objUserContratoCons = Nothing
'''  Else
'''    If objRs.RecordCount = 1 Then
'''      txtContratoFim = objRs.Fields("NUMERO").Value & ""
'''      txtEmpresaFim = objRs.Fields("NOME").Value & ""
'''    Else
'''      'Novo : apresentar tela para seleção do contrato
'''      Set objUserContratoCons = New frmUserContratoCons
'''      objUserContratoCons.strContrato = txtContrato.Text
'''      objUserContratoCons.Show vbModal
'''
'''      If objUserContratoCons.strContrato = "" Then
'''        txtContrato.Text = ""
'''        txtContratoFim.Text = ""
'''        txtEmpresaFim.Text = ""
'''        TratarErroPrevisto "Selecione um contrato"
'''        Pintar_Controle txtContrato, tpCorContr_Erro
'''        SetarFoco txtContrato
'''        Exit Sub
'''      Else
'''        'Cadastrar NF
'''        CadastrarNF
'''      End If
'''      Set objUserContratoCons = Nothing
'''    End If
'''  End If
'''  '
'''  objRs.Close
'''  Set objRs = Nothing
'''  Set objContrato = Nothing
'''  'cmdOk.Default = True
'''
'''  Exit Sub
'''trata:
'''  TratarErro Err.Number, Err.Description, Err.Source
'''End Sub

Public Sub CadastrarNF()
  On Error GoTo trata
  '
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source

End Sub


Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  '
  If Not Valida_Data(mskDtSaida, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de início válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Data(mskDtIniCob, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data de término válida" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_Moeda(txtSequencial, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher o número da NFSR válido" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Not Valida_String(cboObra, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Selecionar a obra" & vbCrLf
    tabDetalhes.Tab = 0
  End If
  If Len(Trim(txtContratoFim.Text & "")) = 0 Then
    If blnSetarFocoControle = True Then
      SetarFoco cboObra
    End If
    Pintar_Controle cboObra, tpCorContr_Erro
    strMsg = strMsg & "Selecionar a obra/contrato" & vbCrLf
    blnSetarFocoControle = False
  End If

  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserNFInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserNFInc.ValidaCampos]", _
            Err.Description
End Function

Private Function ValidaCamposEstoqueDim(strUnidade As String) As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCamposEstoqueDim = False
  If Not Valida_Moeda(mskValor, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Pressncher o valor válido" & vbCrLf
  End If
  If Not Valida_Moeda(mskAltura, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Pressncher a Altura válida" & vbCrLf
  End If
  If Not Valida_Moeda(mskLargura, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Pressncher a Largura válida" & vbCrLf
  End If
  If Len(strMsg) = 0 Then
    If mskLargura.ClipText <> "" Or mskAltura.ClipText <> "" Then
      'Entrou com a Largura e/ou altura da peça
      Select Case strUnidade
      Case RectpIcUnidade.tpIcUnidade_M2
        'M2
        If mskLargura.ClipText = "" Or mskAltura.ClipText = "" Then
          'Unidade M2 a largura e altura são obrigatórios
          strMsg = strMsg & "Para o tipo metro quadrado a largura e a altura são obrigatórias" & vbCrLf
          SetarFoco txtPeca
        End If
        '
      Case RectpIcUnidade.tpIcUnidade_MLINEAR
        'Mlinear
        If mskLargura.ClipText = "" Then
          'Unidade MLinear a largura é obrigatório
          strMsg = strMsg & "Para o tipo metro Linear a largura é obrigatória" & vbCrLf
          SetarFoco txtPeca
        End If
        If mskAltura.ClipText <> "" Then
          'Unidade MLinear a largura é obrigatório
          strMsg = strMsg & "Para o tipo metro Linear a altura não pode ser cadastrada" & vbCrLf
          SetarFoco txtPeca
        End If
        '
      Case RectpIcUnidade.tpIcUnidade_UNID
        'Unidade não aceita largura ou altura
        strMsg = strMsg & "Para o tipo unidade não aceita digitar largura e/ou altura" & vbCrLf
        SetarFoco txtPeca
        '
      End Select
      
      
      
    End If
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserNFInc.ValidaCamposEstoqueDim]"
    ValidaCamposEstoqueDim = False
  Else
    ValidaCamposEstoqueDim = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserNFInc.ValidaCamposEstoqueDim]", _
            Err.Description
End Function

Private Function ValidaCamposEstoque() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCamposEstoque = False
  If Len(txtPecaFim.Text) = 0 Then
    strMsg = strMsg & "Preencher a descrição da peça."
    Pintar_Controle txtPeca, tpCorContr_Erro
    SetarFoco txtPeca
  End If
  If Not Valida_Moeda(mskQuantidade, TpObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Pressncher a quantidade válida" & vbCrLf
  End If
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserNFInc.ValidaCamposEstoque]"
    ValidaCamposEstoque = False
  Else
    ValidaCamposEstoque = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserNFInc.ValidaCamposEstoque]", _
            Err.Description
End Function

Private Sub txtSequencial_GotFocus()
  Selecionar_Conteudo txtSequencial
End Sub
Private Sub txtSequencial_LostFocus()
  Pintar_Controle txtSequencial, tpCorContr_Normal
End Sub
