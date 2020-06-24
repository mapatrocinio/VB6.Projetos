VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmUserRelMedArrecadador 
   Caption         =   "Relatório de Medições de Arrecadador"
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
      TabIndex        =   13
      Top             =   120
      Width           =   5775
      Begin VB.OptionButton optSai1 
         Caption         =   "Vídeo"
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
      TabIndex        =   11
      TabStop         =   0   'False
      Top             =   3510
      Width           =   6105
      Begin VB.CommandButton cmdCancelar 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   880
         Left            =   3150
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdRelatorio 
         Caption         =   "ENTER"
         Default         =   -1  'True
         Height          =   880
         Left            =   1950
         Style           =   1  'Graphical
         TabIndex        =   7
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
      TabIndex        =   9
      Top             =   840
      Width           =   5775
      Begin VB.ComboBox cboArrecadador 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   1020
         Width           =   3165
      End
      Begin VB.ComboBox cboMaquina 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   660
         Width           =   3165
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
      Begin MSMask.MaskEdBox mskBoleto 
         Height          =   255
         Left            =   1200
         TabIndex        =   6
         Top             =   1380
         Width           =   1695
         _ExtentX        =   2990
         _ExtentY        =   450
         _Version        =   393216
         Format          =   "#,##0;($#,##0)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Boleto"
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1380
         Width           =   1095
      End
      Begin VB.Label Label1 
         Caption         =   "Arrecadador"
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   15
         Top             =   1020
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Máquina"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   14
         Top             =   660
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Até"
         Height          =   255
         Index           =   0
         Left            =   2640
         TabIndex        =   12
         Tag             =   "lblIdCliente"
         Top             =   360
         Width           =   855
      End
      Begin VB.Label lblCliente 
         Caption         =   "Período : "
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   10
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
Attribute VB_Name = "frmUserRelMedArrecadador"
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
  Dim datFinal        As Date
  Dim datFinalReal    As Date
  Dim lngMAQUINAID      As Long
  Dim lngBOLETOID       As Long
  Dim lngARRECADADOR    As Long
  AmpS
  
  If Not IsDate(mskData(0).Text) Then
    AmpN
    MsgBox "Data Inicial Inválida !", vbOKOnly, TITULOSISTEMA
    SetarFoco mskData(0)
    Pintar_Controle mskData(0), tpCorContr_Erro
    Exit Sub
  ElseIf Not IsDate(mskData(1).Text) Then
    AmpN
    MsgBox "Data Final Inválida !", vbOKOnly, TITULOSISTEMA
    SetarFoco mskData(1)
    Pintar_Controle mskData(1), tpCorContr_Erro
    Exit Sub
  ElseIf Not Valida_Moeda(mskBoleto, TpNaoObrigatorio, True, False, True) Then
    AmpN
    MsgBox "Número do Boleto inválido !", vbOKOnly, TITULOSISTEMA
    Exit Sub
  End If
  '
  Set objGeral = New busSisMaq.clsGeral
  '
  lngMAQUINAID = 0
  If cboMaquina.Text <> "" Then
    strSql = "SELECT MAQUINA.PKID FROM EQUIPAMENTO " & _
        " INNER JOIN MAQUINA ON EQUIPAMENTO.PKID = MAQUINA.EQUIPAMENTOID " & _
        " WHERE EQUIPAMENTO.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
        " AND EQUIPAMENTO.NUMERO = " & Formata_Dados(cboMaquina.Text, tpDados_Texto) & _
        " ORDER BY EQUIPAMENTO.NUMERO;"
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngMAQUINAID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  '
  lngARRECADADOR = 0
  If cboArrecadador.Text <> "" Then
    strSql = "SELECT FUNCIONARIO.PESSOAID PKID FROM PESSOA " & _
        " INNER JOIN ARRECADADOR ON PESSOA.PKID = ARRECADADOR.PESSOAID " & _
        " INNER JOIN FUNCIONARIO ON PESSOA.PKID = FUNCIONARIO.PESSOAID " & _
        " WHERE FUNCIONARIO.INDEXCLUIDO = " & Formata_Dados("N", tpDados_Texto) & _
        " AND FUNCIONARIO.USUARIO = " & Formata_Dados(cboArrecadador.Text, tpDados_Texto)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngARRECADADOR = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
  End If
  '
  lngBOLETOID = 0
  If mskBoleto.Text <> "" Then
    strSql = "SELECT BOLETOARREC.PKID FROM BOLETOARREC " & _
        " WHERE BOLETOARREC.STATUS <> " & Formata_Dados("C", tpDados_Texto) & _
        " AND BOLETOARREC.NUMERO = " & Formata_Dados(mskBoleto.Text, tpDados_Longo)
    Set objRs = objGeral.ExecutarSQL(strSql)
    If Not objRs.EOF Then
      lngBOLETOID = objRs.Fields("PKID").Value
    End If
    objRs.Close
    Set objRs = Nothing
    '
  End If
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
  If lngMAQUINAID = 0 Then
    Report1.Formulas(2) = "MAQUINAID = True = true"
  Else
    Report1.Formulas(2) = "MAQUINAID = {MAQUINA.PKID} = " & lngMAQUINAID
  End If
  If lngARRECADADOR = 0 Then
    Report1.Formulas(3) = "ARRECADADORID = True = true"
  Else
    Report1.Formulas(3) = "ARRECADADORID = {FUNCIONARIO.PESSOAID} = " & lngARRECADADOR
  End If
  If lngBOLETOID = 0 Then
    Report1.Formulas(4) = "BOLETOID = True = true"
  Else
    Report1.Formulas(4) = "BOLETOID = {BOLETOARREC.PKID} = " & lngBOLETOID
  End If
  '
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
  Report1.ReportFileName = gsReportPath & "MedicaoArrecadador.rpt"
  '
  mskData(0).Text = "01/" & Format(Month(Now), "00") & "/" & Format(Year(Now), "0000")
  mskData(1).Text = Format(Now, "DD/MM/YYYY")
  '
  strSql = "SELECT EQUIPAMENTO.NUMERO FROM EQUIPAMENTO " & _
      " WHERE EQUIPAMENTO.STATUS = " & Formata_Dados("A", tpDados_Texto) & _
      " ORDER BY EQUIPAMENTO.NUMERO;"
  PreencheCombo cboMaquina, strSql, False, True
  '
  strSql = "SELECT FUNCIONARIO.USUARIO FROM PESSOA " & _
      " INNER JOIN ARRECADADOR ON PESSOA.PKID = ARRECADADOR.PESSOAID " & _
      " INNER JOIN FUNCIONARIO ON PESSOA.PKID = FUNCIONARIO.PESSOAID " & _
      " WHERE FUNCIONARIO.INDEXCLUIDO = " & Formata_Dados("N", tpDados_Texto) & _
      " ORDER BY FUNCIONARIO.USUARIO;"
  PreencheCombo cboArrecadador, strSql, False, True
  '
  AmpN
  Exit Sub
RotErro:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub


Private Sub mskBoleto_GotFocus()
  Seleciona_Conteudo_Controle mskBoleto
End Sub

Private Sub mskBoleto_LostFocus()
  Pintar_Controle mskBoleto, tpCorContr_Normal
End Sub

Private Sub mskData_GotFocus(Index As Integer)
  Seleciona_Conteudo_Controle mskData(Index)
End Sub

Private Sub mskData_LostFocus(Index As Integer)
  Pintar_Controle mskData(Index), tpCorContr_Normal
End Sub

