VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmUserRelFechaCaixa 
   Caption         =   "Relat�rio de Fechamento do Caixa"
   ClientHeight    =   4485
   ClientLeft      =   450
   ClientTop       =   1920
   ClientWidth     =   6105
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   4485
   ScaleWidth      =   6105
   Begin VB.Frame Frame1 
      Height          =   615
      Left            =   120
      TabIndex        =   11
      Top             =   120
      Width           =   5775
      Begin VB.OptionButton optSai1 
         Caption         =   "V�deo"
         Height          =   255
         Left            =   120
         TabIndex        =   0
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optSai2 
         Caption         =   "Impressora"
         Height          =   255
         Left            =   1440
         TabIndex        =   1
         Top             =   240
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
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   6105
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   3510
      Width           =   6105
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   880
         Left            =   3150
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdRelatorio 
         Caption         =   "ENTER"
         Default         =   -1  'True
         Height          =   880
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   60
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
      Height          =   2535
      Left            =   120
      TabIndex        =   7
      Top             =   840
      Width           =   5775
      Begin VB.ComboBox cboPeriodo 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   630
         Width           =   1665
      End
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
         Caption         =   "Per�odo"
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   630
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "At�"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   10
         Tag             =   "lblIdCliente"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCliente 
         Caption         =   "Per�odo : "
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   8
         Tag             =   "lblIdCliente"
         Top             =   360
         Width           =   855
      End
   End
   Begin Crystal.CrystalReport Report1 
      Left            =   5730
      Top             =   3150
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      PrintFileLinesPerPage=   60
   End
End
Attribute VB_Name = "frmUserRelFechaCaixa"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancelar_Click()
  Unload Me
End Sub

Private Sub cmdRelatorio_Click()
  On Error GoTo TratErro
  Dim lngSeq    As Long
  Dim objGeral  As busSisMaq.clsGeral
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim curDespesa      As Currency
  Dim curEntradaCaixa As Currency
  Dim datFinal        As Date
  Dim datFinalReal    As Date
  Dim lngPERIODOID    As Long
  AmpS
  
  If Not IsDate(mskData(0).Text) Then
    AmpN
    MsgBox "Data Inicial Inv�lida !", vbOKOnly, TITULOSISTEMA
    SetarFoco mskData(0)
    Pintar_Controle mskData(0), tpCorContr_Erro
    Exit Sub
  ElseIf Not IsDate(mskData(1).Text) Then
    AmpN
    MsgBox "Data Final Inv�lida !", vbOKOnly, TITULOSISTEMA
    SetarFoco mskData(1)
    Pintar_Controle mskData(1), tpCorContr_Erro
    Exit Sub
  End If
  'Tratar inser��o em tabela tempor�ria
  Set objGeral = New busSisMaq.clsGeral
  '
  'Tratar Per�odo
  lngPERIODOID = 0
  If cboPeriodo.Text <> "" Then
    strSql = "Select PKID FROM PERIODO WHERE PERIODO = " & Formata_Dados(cboPeriodo.Text, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngPERIODOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  '
  lngSeq = objGeral.ExecutarSQLRetInteger("SP_REL_MOVIMENTO_MAQUINA", Array( _
                                            mp("@PESSOAID", adInteger, 4, giFuncionarioId), _
                                            mp("@PERIODOID", adInteger, 4, lngPERIODOID), _
                                            mp("@DATAINICHR", adVarChar, 30, mskData(0).Text), _
                                            mp("@DATAFIMCHR", adVarChar, 30, mskData(1).Text)))
  'DESPESA
  datFinal = CDate(Right(mskData(1).Text, 4) & "/" & Mid(mskData(1).Text, 4, 2) & "/" & Left(mskData(1).Text, 2))
  datFinalReal = DateAdd("d", 1, datFinal)
  '
  strSql = "SELECT ISNULL(SUM(DESPESA.VR_PAGAR), 0) AS TOTAL "
  strSql = strSql & " FROM DESPESA " & _
          " INNER JOIN TURNO ON TURNO.PKID = DESPESA.TURNOID " & _
          " WHERE TURNO.DATA >= " & Formata_Dados(mskData(0).Text, tpDados_DataHora) & _
          " AND TURNO.DATA < " & Formata_Dados(Format(datFinalReal, "DD/MM/YYYY"), tpDados_DataHora)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    curDespesa = objRs.Fields("TOTAL").Value
  End If
  objRs.Close
  Set objRs = Nothing
  '
  strSql = "SELECT ISNULL(SUM(ENTRADA.VALOR),0) AS TOTAL " & _
            "FROM ENTRADA " & _
            " INNER JOIN TURNO ON ENTRADA.TURNOID = TURNO.PKID " & _
            " WHERE TURNO.DATA >= " & Formata_Dados(mskData(0).Text, tpDados_DataHora) & _
            " AND TURNO.DATA < " & Formata_Dados(Format(datFinalReal, "DD/MM/YYYY"), tpDados_DataHora)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    curEntradaCaixa = objRs.Fields("TOTAL").Value
  End If
  objRs.Close
  Set objRs = Nothing
  
  Set objGeral = Nothing
  '
  '
  If optSai1.Value Then
    Report1.Destination = 0 'Video
  ElseIf optSai2.Value Then
    Report1.Destination = 1   'Impressora
  End If
  Report1.CopiesToPrinter = 1
  Report1.WindowState = crptMaximized
  '
  Report1.Formulas(0) = "DataIni = Date(" & Right(mskData(0).Text, 4) & ", " & Mid(mskData(0).Text, 4, 2) & ", " & Left(mskData(0).Text, 2) & ")"
  Report1.Formulas(1) = "DataFim = Date(" & Right(mskData(1).Text, 4) & ", " & Mid(mskData(1).Text, 4, 2) & ", " & Left(mskData(1).Text, 2) & ")"
  Report1.Formulas(2) = "PessoaId = " & giFuncionarioId
  Report1.Formulas(3) = "Vr_Despesa = " & Formata_Dados(curDespesa, tpDados_Moeda)
  Report1.Formulas(4) = "Vr_Fundo = " & Formata_Dados(curEntradaCaixa, tpDados_Moeda)
  '
  Report1.Action = 1
  '
  AmpN
  Exit Sub
  
TratErro:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Private Sub Form_Activate()
  SetarFoco mskData(0)
End Sub

Private Sub Form_Load()
  On Error GoTo RotErro
  AmpS
  CenterForm Me
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdCancelar, , , , , , cmdRelatorio
  '
  Report1.Connect = ConnectRpt
  Report1.ReportFileName = gsReportPath & "FechamentoCaixa.rpt"
  '
  mskData(0).Text = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
  mskData(1).Text = Format(Now, "DD/MM/YYYY")
  '
  '
  strSql = "SELECT PERIODO.PERIODO FROM PERIODO " & _
      " ORDER BY PERIODO.PERIODO;"
  PreencheCombo cboPeriodo, strSql, False, True
  '
  AmpN
  Exit Sub
RotErro:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Private Sub mskData_GotFocus(Index As Integer)
  Seleciona_Conteudo_Controle mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub
