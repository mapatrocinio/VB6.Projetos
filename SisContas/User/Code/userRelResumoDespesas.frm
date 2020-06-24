VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmUserRelResumoDespesas 
   Caption         =   "Demonstrativo de Resumo Geral de Despesas"
   ClientHeight    =   3840
   ClientLeft      =   450
   ClientTop       =   1920
   ClientWidth     =   6105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   3840
   ScaleWidth      =   6105
   Begin VB.Frame Frame1 
      Height          =   975
      Left            =   120
      TabIndex        =   8
      Top             =   120
      Width           =   5775
      Begin VB.OptionButton optSai2 
         Caption         =   "Impressora"
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
         TabIndex        =   1
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton optSai1 
         Caption         =   "V�deo"
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
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   765
      Left            =   0
      ScaleHeight     =   765
      ScaleWidth      =   6105
      TabIndex        =   6
      Top             =   3075
      Width           =   6105
      Begin VB.CommandButton cmdRelatorio 
         Default         =   -1  'True
         Height          =   735
         Left            =   1800
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   0
         Width           =   1215
      End
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Height          =   735
         Left            =   3000
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   0
         Width           =   1215
      End
   End
   Begin VB.Frame Frame2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   120
      TabIndex        =   4
      Top             =   1200
      Width           =   5775
      Begin MSMask.MaskEdBox mskData 
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskData 
         Height          =   255
         Index           =   1
         Left            =   3120
         TabIndex        =   3
         Top             =   360
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   450
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label1 
         Caption         =   "At�"
         Height          =   255
         Left            =   2640
         TabIndex        =   7
         Tag             =   "lblIdCliente"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCliente 
         Caption         =   "Per�odo : "
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Tag             =   "lblIdCliente"
         Top             =   360
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   3360
      Top             =   5760
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmUserRelResumoDespesas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdRelatorio_Click()
  Dim objGeral        As busSisContas.clsGeral
  Dim objRs           As ADODB.Recordset
  Dim strSql          As String
  Dim curVrCartTxAdm  As Currency
  Dim curVrChqDevol   As Currency
  Dim curVrPenhor     As Currency
  On Error GoTo TratErro
  AmpS
  
  If Not Valida_Data(mskData(0), TpObrigatorio) Then
    AmpN
    TratarErroPrevisto "Data Inicial Inv�lida", "[cmdRelatorio_Click]"
    SetarFoco mskData(0)
    Pintar_Controle mskData(0), tpCorContr_Erro
    Exit Sub
  ElseIf Not Valida_Data(mskData(1), TpObrigatorio) Then
    AmpN
    TratarErroPrevisto "Data Final Inv�lida", "[cmdRelatorio_Click]"
    SetarFoco mskData(1)
    Pintar_Controle mskData(1), tpCorContr_Erro
    Exit Sub
  End If
  '
  If optSai1.Value Then
    Report1.Destination = 0 'Video
  ElseIf optSai2.Value Then
    Report1.Destination = 1   'Impressora
  End If
'''  Set objGeral = New busSisContas.clsGeral
'''  'Captura total de taxa de cart�es
'''  strSql = "SELECT Sum(vw_cons_t_cred.VrCalcDescCartTxAdm) As VALOR " & _
'''    "FROM TURNO LEFT JOIN vw_cons_t_cred ON TURNO.PKID = vw_cons_t_cred.PKID " & _
'''    "WHERE TURNO.DATA >= " & Formata_Dados(mskData(0).Text & " 00:00", tpDados_DataHora, tpNulo_Aceita) & _
'''    " AND TURNO.DATA <= " & Formata_Dados(mskData(1).Text & " 23:59", tpDados_DataHora, tpNulo_Aceita)
'''  Set objRs = objGeral.ExecutarSQL(strSql)
'''  If objRs.EOF Then
    curVrCartTxAdm = 0
'''  ElseIf Not IsNumeric(objRs.Fields("VALOR").Value) Then
'''    curVrCartTxAdm = 0
'''  Else
'''    curVrCartTxAdm = objRs.Fields("VALOR").Value
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
  '
  'Captura total de cheques devolvidos
'''  strSql = "SELECT Sum(CHEQUE.VALOR) As VALORCHEQUE " & _
'''    "FROM CHEQUE " & _
'''    "WHERE DTDEVOLUCAO >= " & Formata_Dados(mskData(0).Text & " 00:00", tpDados_DataHora, tpNulo_Aceita) & _
'''    " AND DTDEVOLUCAO <= " & Formata_Dados(mskData(1).Text & " 23:59", tpDados_DataHora, tpNulo_Aceita) & _
'''    " AND CHEQUE.STATUS = " & Formata_Dados("D", tpDados_Texto)
'''  Set objRs = objGeral.ExecutarSQL(strSql)
'''  If objRs.EOF Then
    curVrChqDevol = 0
'''  ElseIf Not IsNumeric(objRs.Fields("VALORCHEQUE").Value) Then
'''    curVrChqDevol = 0
'''  Else
'''    curVrChqDevol = objRs.Fields("VALORCHEQUE").Value
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
  'Captura total de penhores
'''  strSql = "SELECT Sum(vw_cons_t_penhor.PgtoPenhor) As VALOR " & _
'''    "FROM TURNO LEFT JOIN vw_cons_t_penhor ON TURNO.PKID = vw_cons_t_penhor.PKID " & _
'''    "WHERE TURNO.DATA >= " & Formata_Dados(mskData(0).Text & " 00:00", tpDados_DataHora, tpNulo_Aceita) & _
'''    " AND TURNO.DATA <= " & Formata_Dados(mskData(1).Text & " 23:59", tpDados_DataHora, tpNulo_Aceita)
'''  Set objRs = objGeral.ExecutarSQL(strSql)
'''  If objRs.EOF Then
'''    curVrPenhor = 0
'''  ElseIf Not IsNumeric(objRs.Fields("VALOR").Value) Then
    curVrPenhor = 0
'''  Else
'''    curVrPenhor = objRs.Fields("VALOR").Value
'''  End If
'''  objRs.Close
'''  Set objRs = Nothing
'''  '
  Report1.CopiesToPrinter = 1
  Report1.WindowState = crptMaximized
  '
  Report1.Formulas(0) = "DataIni = Date(" & Right(mskData(0).Text, 4) & ", " & Mid(mskData(0).Text, 4, 2) & ", " & Left(mskData(0).Text, 2) & ")"
  Report1.Formulas(1) = "DataFim = Date(" & Right(mskData(1).Text, 4) & ", " & Mid(mskData(1).Text, 4, 2) & ", " & Left(mskData(1).Text, 2) & ")"
  Report1.Formulas(2) = "TotTaxaCart = " & Formata_Dados(curVrCartTxAdm, tpDados_Moeda, tpNulo_Aceita)
  Report1.Formulas(3) = "TotCheqsDevol = " & Formata_Dados(curVrChqDevol, tpDados_Moeda, tpNulo_Aceita)
  Report1.Formulas(4) = "TotPenhor = " & Formata_Dados(curVrPenhor, tpDados_Moeda, tpNulo_Aceita)
  Report1.Formulas(5) = "ParceiroId = " & glParceiroId
  
  Report1.Action = 1
  '
  AmpN
  Set objGeral = Nothing
  Exit Sub
  
TratErro:
  AmpN
  TratarErro Err.Number, Err.Description, "[frmUserRelContas.cmdRelatorio_Click]"
End Sub

Private Sub Form_Activate()
  mskData(0).SetFocus
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  AmpS
  CenterForm Me
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar, , , , , , cmdRelatorio
  '
  Report1.Connect = ConnectRpt
  '
  Report1.ReportFileName = gsReportPath & "ResumoDespesas.rpt"
  mskData(0).Text = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
  mskData(1).Text = Format(Now, "DD/MM/YYYY")
  '
  AmpN
  Exit Sub
trata:
  AmpN
  TratarErro Err.Number, Err.Description, "[frmUserRelContas.Form_Load]"
End Sub

Private Sub mskData_GotFocus(Index As Integer)
  Selecionar_Conteudo mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub
