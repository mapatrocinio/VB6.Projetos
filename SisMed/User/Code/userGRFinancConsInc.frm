VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmUserGRFinancCons 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar GR"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   6645
      Left            =   90
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   120
      Width           =   11715
      _ExtentX        =   20664
      _ExtentY        =   11721
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Consultar GR"
      TabPicture(0)   =   "userGRFinancConsInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label6(3)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label7(1)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label6(0)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label6(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Report1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "mskDtInicio"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "grdGR"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtProntuario"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "cmdNormal(0)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtPrestador"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtFuncionario"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "SSPanel1"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Picture1"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).ControlCount=   14
      Begin VB.PictureBox Picture1 
         BorderStyle     =   0  'None
         Height          =   285
         Left            =   4020
         ScaleHeight     =   285
         ScaleWidth      =   3885
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   420
         Width           =   3885
         Begin VB.OptionButton optGR 
            Caption         =   "&Todas"
            Height          =   315
            Index           =   0
            Left            =   0
            TabIndex        =   1
            Top             =   0
            Value           =   -1  'True
            Width           =   1065
         End
         Begin VB.OptionButton optGR 
            Caption         =   "&Não paga a prestadores"
            Height          =   315
            Index           =   1
            Left            =   1290
            TabIndex        =   2
            Top             =   0
            Width           =   2505
         End
      End
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   1005
         Left            =   3540
         ScaleHeight     =   945
         ScaleWidth      =   8025
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   5610
         Width           =   8085
         Begin VB.CommandButton cmdImprimirTodas 
            Caption         =   "&V"
            Height          =   880
            Left            =   30
            Style           =   1  'Graphical
            TabIndex        =   8
            Top             =   30
            Width           =   1335
         End
         Begin VB.CommandButton cmdExcluir 
            Caption         =   "&X"
            Height          =   880
            Left            =   1380
            Style           =   1  'Graphical
            TabIndex        =   9
            Top             =   30
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   6660
            Style           =   1  'Graphical
            TabIndex        =   13
            Top             =   30
            Width           =   1335
         End
         Begin VB.CommandButton cmdAlterar 
            Caption         =   "&Z"
            Height          =   880
            Left            =   4020
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   30
            Width           =   1335
         End
         Begin VB.CommandButton cmdConsultar 
            Caption         =   "ENTER"
            Default         =   -1  'True
            Height          =   880
            Left            =   5340
            Style           =   1  'Graphical
            TabIndex        =   12
            Top             =   30
            Width           =   1335
         End
         Begin VB.CommandButton cmdImprimir 
            Caption         =   "&Y"
            Height          =   880
            Left            =   2700
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   30
            Width           =   1335
         End
      End
      Begin VB.TextBox txtFuncionario 
         Height          =   285
         Left            =   870
         TabIndex        =   5
         Text            =   "txtFuncionario"
         Top             =   1320
         Width           =   7035
      End
      Begin VB.TextBox txtPrestador 
         Height          =   285
         Left            =   870
         TabIndex        =   4
         Text            =   "txtPrestador"
         Top             =   1020
         Width           =   7035
      End
      Begin VB.CommandButton cmdNormal 
         Caption         =   "&Consultar"
         Height          =   255
         Index           =   0
         Left            =   8010
         TabIndex        =   6
         Top             =   1350
         Width           =   1215
      End
      Begin VB.TextBox txtProntuario 
         Height          =   285
         Left            =   870
         TabIndex        =   3
         Text            =   "txtProntuario"
         Top             =   720
         Width           =   7035
      End
      Begin TrueDBGrid60.TDBGrid grdGR 
         Height          =   3960
         Left            =   90
         OleObjectBlob   =   "userGRFinancConsInc.frx":001C
         TabIndex        =   7
         Top             =   1650
         Width           =   11520
      End
      Begin MSMask.MaskEdBox mskDtInicio 
         Height          =   255
         Left            =   870
         TabIndex        =   0
         Top             =   450
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         AutoTab         =   -1  'True
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin Crystal.CrystalReport Report1 
         Left            =   90
         Top             =   5790
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label7 
         Caption         =   "GR´s"
         Height          =   255
         Index           =   0
         Left            =   2730
         TabIndex        =   21
         Top             =   450
         Width           =   645
      End
      Begin VB.Label Label6 
         Caption         =   "Func."
         Height          =   255
         Index           =   1
         Left            =   90
         TabIndex        =   18
         Top             =   1320
         Width           =   795
      End
      Begin VB.Label Label6 
         Caption         =   "Prestador"
         Height          =   255
         Index           =   0
         Left            =   90
         TabIndex        =   17
         Top             =   1020
         Width           =   795
      End
      Begin VB.Label Label7 
         Caption         =   "Data"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   16
         Top             =   450
         Width           =   645
      End
      Begin VB.Label Label6 
         Caption         =   "Prontuário"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   15
         Top             =   720
         Width           =   795
      End
   End
End
Attribute VB_Name = "frmUserGRFinancCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnRetorno           As Boolean
Dim blnFechar               As Boolean
Public strGR                As String
'
Public strDataIni           As String
Public strDataFim           As String
Public strProntuario        As String

'Variáveis para Grid
'
Dim GR_COLUNASMATRIZ        As Long
Dim GR_LINHASMATRIZ         As Long
Private GR_Matriz()         As String
'

Public Sub GR_MontaMatriz(strDataIni As String, _
                          strDataFim As String, _
                          strGR As String, _
                          Optional strProntuario As String, _
                          Optional strPrestador As String, _
                          Optional strFucnionario As String)
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  'strGR P=PAGA A PRESTADORES T=TOTAL
  On Error GoTo trata
  
  Set clsGer = New busSisMed.clsGeral
  '
  strSql = "SELECT MAX(GRPGTO.GRPAGAMENTOID), MAX(GR.STATUS), GR.PKID, MAX(PRONTUARIO.NOME) AS NOME, MAX(PACIENTE.NOME) AS NOME_PACIENTE, MAX(FUNCIONARIO.NOME) AS FUNC, MAX(GR.SEQUENCIAL) AS SEQUENCIAL, MAX(GR.SENHA) AS SENHA, MAX(GR.DATA) AS DATA, SUM(GRPROCEDIMENTO.VALOR) AS VALOR " & _
      " From GR " & _
      " INNER JOIN PRONTUARIO AS FUNCIONARIO ON GR.FUNCIONARIOID = FUNCIONARIO.PKID " & _
      " LEFT JOIN GRPROCEDIMENTO ON GR.PKID = GRPROCEDIMENTO.GRID " & _
      " INNER JOIN PRONTUARIO PACIENTE ON GR.PRONTUARIOID = PACIENTE.PKID " & _
      " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = ATENDE.PRONTUARIOID " & _
      " INNER JOIN TURNO ON TURNO.PKID = GR.TURNOID " & _
      " LEFT JOIN PRESTADORPROCEDIMENTO ON PRESTADORPROCEDIMENTO.PROCEDIMENTOID = GRPROCEDIMENTO.PROCEDIMENTOID " & _
      "           AND PRESTADORPROCEDIMENTO.PRONTUARIOID = ATENDE.PRONTUARIOID " & _
      " LEFT JOIN GRPGTO ON GR.PKID = GRPGTO.GRID " & _
      " WHERE "
  If strDataIni & "" <> "" Then
    strSql = strSql & " TURNO.DATA >= " & Formata_Dados(strDataIni, tpDados_DataHora) & _
        " AND TURNO.DATA < " & Formata_Dados(strDataFim, tpDados_DataHora) & _
        " AND GR.STATUS = " & Formata_Dados("F", tpDados_Texto)
  End If
  If strProntuario & "" <> "" Then
    If strDataIni & "" <> "" Then
      strSql = strSql & " AND PACIENTE.NOME LIKE " & Formata_Dados(strProntuario & "%", tpDados_Texto)
    Else
      strSql = strSql & " PACIENTE.NOME LIKE " & Formata_Dados(strProntuario & "%", tpDados_Texto)
    End If
  End If
  If strPrestador & "" <> "" Then
    If strDataIni & "" <> "" Or strProntuario & "" <> "" Then
      strSql = strSql & " AND PRONTUARIO.NOME LIKE " & Formata_Dados(strPrestador & "%", tpDados_Texto)
    Else
      strSql = strSql & " PRONTUARIO.NOME LIKE " & Formata_Dados(strPrestador & "%", tpDados_Texto)
    End If
  End If
  If strFucnionario & "" <> "" Then
    If strDataIni & "" <> "" Or strProntuario & "" <> "" Or strPrestador & "" <> "" Then
      strSql = strSql & " AND FUNCIONARIO.NOME LIKE " & Formata_Dados(strFucnionario & "%", tpDados_Texto)
    Else
      strSql = strSql & " FUNCIONARIO.NOME LIKE " & Formata_Dados(strFucnionario & "%", tpDados_Texto)
    End If
  End If
  If strGR & "" = "P" Then
    If strDataIni & "" <> "" Or strProntuario & "" <> "" Or strPrestador & "" <> "" Or strFucnionario & "" <> "" Then
      strSql = strSql & " AND GR.PKID NOT IN (SELECT TOP 1 GRID FROM GRPGTO WHERE GRPGTO.GRID = GR.PKID)"
    Else
      strSql = strSql & " GR.PKID NOT IN (SELECT TOP 1 GRID FROM GRPGTO WHERE GRPGTO.GRID = GR.PKID)"
    End If
  End If
  'strGR = IIf(optGR(1).Value, "P", "T")
  strSql = strSql & " GROUP BY GR.PKID "
  strSql = strSql & " ORDER BY GR.SENHA, GR.SEQUENCIAL"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    GR_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim GR_Matriz(0 To GR_COLUNASMATRIZ - 1, 0 To GR_LINHASMATRIZ - 1)
  Else
    ReDim GR_Matriz(0 To GR_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To GR_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To GR_COLUNASMATRIZ - 1  'varre as colunas
          GR_Matriz(intJ, intI) = objRs(intJ) & ""
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
  Dim objUserGRInc As SisMed.frmUserGRInc
  Dim strMsg As String
  On Error GoTo trata
  'Itens da GR
  If RetornaCodTurnoCorrente = 0 Then
    MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGR
    Exit Sub
  End If
  If Len(Trim(grdGR.Columns("GRID").Value & "")) = 0 Then
    MsgBox "Selecione uma GR para alterar seus ítens.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGR
    Exit Sub
  End If
'  If Trim(grdGR.Columns("Status").Value & "") = "F" Then
    'Pedir senha superior para alterar uma GR já fechada
    '----------------------------
    '----------------------------
    'Pede Senha Superior (Diretor, Gerente ou Administrador
    If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
      'Só pede senha superior se quem estiver logado não for superior
      gsNomeUsuLib = ""
      gsNivelUsuLib = ""
      frmUserLoginSup.Show vbModal
      
      If Len(Trim(gsNomeUsuLib)) = 0 Then
        strMsg = "É necessário a confirmação com senha superior para alterar uma GR."
        TratarErroPrevisto strMsg, "cmdConfirmar_Click"
        SetarFoco grdGR
        Exit Sub
      Else
        'Capturou Nome do Usuário, continua com processo
      End If
    End If
    '--------------------------------
    '--------------------------------
' End If
  Set objUserGRInc = New SisMed.frmUserGRInc
  objUserGRInc.Status = tpStatus_Alterar
  objUserGRInc.IcEstadoGR = tpIcEstadoGR_Proc
  objUserGRInc.lngGRID = grdGR.Columns("GRID").Value
  objUserGRInc.Show vbModal
  Set objUserGRInc = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdConsultar_Click()
  Dim objUserGRInc As SisMed.frmUserGRInc
  On Error GoTo trata
  'consultar GR
  If RetornaCodTurnoCorrente = 0 Then
    MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGR
    Exit Sub
  End If
  If Len(Trim(grdGR.Columns("GRID").Value & "")) = 0 Then
    MsgBox "Selecione uma GR para alterá-la.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGR
    Exit Sub
  End If
  Set objUserGRInc = New SisMed.frmUserGRInc
  objUserGRInc.Status = tpStatus_Consultar
  objUserGRInc.IcEstadoGR = tpIcEstadoGR_Con
  objUserGRInc.lngGRID = grdGR.Columns("GRID").Value
  objUserGRInc.Show vbModal
  Set objUserGRInc = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdExcluir_Click()
  On Error GoTo trata
  Dim objGeral    As busSisMed.clsGeral
  Dim strSql      As String
  Dim objRs       As ADODB.Recordset
  Dim objGR       As busSisMed.clsGR
  Dim datInicio             As Date
  Dim strGR                 As String
  Dim strMsg                As String
  'Cancelamento da GR
  If RetornaCodTurnoCorrente = 0 Then
    MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGR
    Exit Sub
  End If
  If Len(Trim(grdGR.Columns("GRID").Value & "")) = 0 Then
    MsgBox "Selecione uma GR para excluí-la.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGR
    Exit Sub
  End If
'''  If Trim(RetornaDescAtende(grdGR.Columns("Atendente").Value & "")) <> gsNomeUsuCompleto Then
'''    If Mid(grdGR.Columns("Atendente").Value & "", 2, 3) <> gsLaboratorio Then
'''      MsgBox "Apenas o atendente que lançou a GR pode excluí-la.", vbExclamation, TITULOSISTEMA
'''      SetarFoco grdGR
'''      Exit Sub
'''    End If
'''  End If
  'Verifica GR já cancelada
  Set objGeral = New busSisMed.clsGeral
  strSql = "SELECT GR.PKID, GR.STATUS FROM GR " & _
      " WHERE GR.PKID = " & Formata_Dados(grdGR.Columns("GRID").Value, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    If objRs.Fields("STATUS").Value & "" = "C" Then
      'GR CANCELADA
      objRs.Close
      Set objRs = Nothing
      Set objGeral = Nothing
      MsgBox "GR já cancelada.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGR
      Exit Sub
    End If
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGeral = Nothing
  'If Trim(grdGR.Columns("Status").Value & "") <> "F" Then
  '  MsgBox "Apenas pode de excluida uma GR fechada.", vbExclamation, TITULOSISTEMA
  '  SetarFoco grdGR
  '  Exit Sub
  'End If
  'If Trim(grdGR.Columns("Status").Value & "") = "F" Then
    'Pedir senha superior para alterar uma GR já fechada
    '----------------------------
    '----------------------------
    'Pede Senha Superior (Diretor, Gerente ou Administrador
    gsNomeUsuLib = ""
    gsNivelUsuLib = ""
    If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
      'Só pede senha superior se quem estiver logado não for superior
      frmUserLoginSup.Show vbModal

      If Len(Trim(gsNomeUsuLib)) = 0 Then
        strMsg = "É necessário a confirmação com senha superior para cancelar uma GR."
        TratarErroPrevisto strMsg, "cmdConfirmar_Click"
        SetarFoco grdGR
        Exit Sub
      Else
        'Capturou Nome do Usuário, continua com processo
      End If
    End If
    '--------------------------------
    '--------------------------------
  'End If
  'Confirmação
  If MsgBox("Confirma cancelamento da GR " & grdGR.Columns("Seq.").Value & " de " & grdGR.Columns("Prontuário").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdGR
    Exit Sub
  End If
  
  Set objGR = New busSisMed.clsGR
  objGR.AlterarStatusGR grdGR.Columns("GRID").Value, _
                        "C", _
                        "", _
                        RetornaCodTurnoCorrente
  Set objGR = Nothing
  IMP_COMP_CANC_GR grdGR.Columns("GRID").Value, gsNomeEmpresa, 1
  If Not ValidaCampos Then
    Exit Sub
  End If
  '
  strDataIni = ""
  strDataFim = ""
  If mskDtInicio.Text <> "__/__/____" Then
    datInicio = CDate(mskDtInicio.Text)
    strDataIni = mskDtInicio.Text & " 00:00"
    strDataFim = Format(DateAdd("d", 1, datInicio), "DD/MM/YYYY 00:00")
  End If
  strGR = IIf(optGR(1).Value, "P", "T")
  '
  GR_COLUNASMATRIZ = grdGR.Columns.Count
  GR_LINHASMATRIZ = 0
  GR_MontaMatriz strDataIni, _
                 strDataFim, _
                 strGR, _
                 txtProntuario.Text, _
                 txtPrestador.Text, _
                 txtFuncionario.Text
                 
  grdGR.Bookmark = Null
  grdGR.ReBind
  grdGR.ApproxCount = GR_LINHASMATRIZ

  grdGR.SetFocus
  Exit Sub
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserGRFinancCons.ValidaCampos]", _
            Err.Description
End Sub

Private Sub cmdFechar_Click()
  '
  blnFechar = True
  Unload Me
End Sub



Private Function ValidaCampos() As Boolean
  On Error GoTo trata
  Dim strMsg                As String
  Dim blnSetarFocoControle  As Boolean
  '
  blnSetarFocoControle = True
  ValidaCampos = False
  If Not Valida_Data(mskDtInicio, TpNaoObrigatorio, blnSetarFocoControle) Then
    strMsg = strMsg & "Preencher a data válida" & vbCrLf
  End If
  If mskDtInicio.Text = "__/__/____" And txtProntuario.Text = "" And txtPrestador.Text = "" And txtFuncionario.Text = "" Then
    strMsg = strMsg & "Preencher a data, prontuário ou prestador" & vbCrLf
    SetarFoco mskDtInicio
  End If
  
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserGRFinancCons.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserGRFinancCons.ValidaCampos]", _
            Err.Description
End Function



Private Sub cmdImprimir_Click()
  Dim objGR As busSisMed.clsGR
  On Error GoTo trata
  'Imprimir GR
  If RetornaCodTurnoCorrente = 0 Then
    MsgBox "Não há turno aberto. favor abrir o turno antes de iniciar a GR.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGR
    Exit Sub
  End If
  If Len(Trim(grdGR.Columns("GRID").Value & "")) = 0 Then
    MsgBox "Selecione uma GR para imprimí-la.", vbExclamation, TITULOSISTEMA
    SetarFoco grdGR
    Exit Sub
  End If
  If Trim(grdGR.Columns("STATUS").Value & "") <> "F" Then
'    If Trim(RetornaNivelAtende(grdGR.Columns("Atendente").Value & "")) <> gsLaboratorio Then
      MsgBox "Não pode haver impressão de uma GR que não esteja fechada ou seja lançada pelo Laboratório.", vbExclamation, TITULOSISTEMA
      SetarFoco grdGR
      Exit Sub
'    End If
  End If
'''  If IsNumeric(Trim(grdGR.Columns("GRPAGAMENTOID").Value)) Then
'''    MsgBox "Apenas poderá ser impressa uma GR ainda não lançada no financeiro.", vbExclamation, TITULOSISTEMA
'''    SetarFoco grdGR
'''    Exit Sub
'''  End If
  'Confirmação
  If MsgBox("Confirma impressão da GR " & grdGR.Columns("Seq.").Value & " de " & grdGR.Columns("Prontuário").Value & "?", vbYesNo, TITULOSISTEMA) = vbNo Then
    SetarFoco grdGR
    Exit Sub
  End If
  
  IMP_COMP_GR grdGR.Columns("GRID").Value, gsNomeEmpresa, 1, True
  'Após impressão altera status para impressa
  Set objGR = New busSisMed.clsGR
  'objGR.AlterarStatusGR grdGR.Columns("ID").Value, _
                        "", _
                        "S"
  Set objGR = Nothing
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub cmdImprimirTodas_Click()
  On Error GoTo TratErro
  AmpS
  
'''  If Not IsDate("01/" & mskData(0).Text) Then
'''    AmpN
'''    MsgBox "Data Inicial Inválida !", vbOKOnly, TITULOSISTEMA
'''    SetarFoco mskData(0)
'''    Pintar_Controle mskData(0), tpCorContr_Erro
'''    Exit Sub
'''  ElseIf Not IsDate("01/" & mskData(1).Text) Then
'''    AmpN
'''    MsgBox "Data Final Inválida !", vbOKOnly, TITULOSISTEMA
'''    SetarFoco mskData(1)
'''    Pintar_Controle mskData(1), tpCorContr_Erro
'''    Exit Sub
'''  End If
  '
  'If optSai1.Value Then
    Report1.Destination = 0 'Video
  'ElseIf optSai2.Value Then
  '  Report1.Destination = 1   'Impressora
  'End If
  'If chkTotal.Value = 0 Then
  '  Report1.ReportFileName = gsReportPath & "ReceitaComparativoMensal.rpt"
  'Else
  '  Report1.ReportFileName = gsReportPath & "ReceitaComparativoMensalTotal.rpt"
  'End If
  Report1.CopiesToPrinter = 1
  Report1.WindowState = crptMaximized
  '
  'Report1.Formulas(0) = "DataIni = Date(" & Right(mskData(0).Text, 4) & ", " & Left(mskData(0).Text, 2) & ", 01)"
  'Report1.Formulas(1) = "DataFim = Date(" & Right(mskData(1).Text, 4) & ", " & Left(mskData(1).Text, 2) & ", " & Retorna_ultimo_dia_do_mes(Left(mskData(1).Text, 2), Right(mskData(1).Text, 4)) & ")"
  'Report1.Formulas(2) = "Especialidade = '" & IIf(cboEspecialidade.Text = "<TODOS>", "*", cboEspecialidade.Text) & "'"
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

Private Sub cmdNormal_Click(Index As Integer)
  Dim objUserProntuarioInc  As SisMed.frmUserProntuarioInc
  Dim datInicio             As Date
  Dim strGR                 As String
  '
  If Index = 0 Then
    If Not ValidaCampos Then
      Exit Sub
    End If
    strDataIni = ""
    strDataFim = ""
    If mskDtInicio.Text <> "__/__/____" Then
      datInicio = CDate(mskDtInicio.Text)
      strDataIni = mskDtInicio.Text & " 00:00"
      strDataFim = Format(DateAdd("d", 1, datInicio), "DD/MM/YYYY 00:00")
    End If
    strGR = IIf(optGR(1).Value, "P", "T")
    '
    GR_COLUNASMATRIZ = grdGR.Columns.Count
    GR_LINHASMATRIZ = 0
    GR_MontaMatriz strDataIni, _
                   strDataFim, _
                   strGR, _
                   txtProntuario.Text, _
                   txtPrestador.Text, _
                   txtFuncionario.Text
                   
    grdGR.Bookmark = Null
    grdGR.ReBind
    grdGR.ApproxCount = GR_LINHASMATRIZ
  
    grdGR.SetFocus
  End If
End Sub


Private Sub grdGR_UnboundReadDataEx( _
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
               Offset + intI, GR_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, GR_COLUNASMATRIZ, GR_LINHASMATRIZ, GR_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, GR_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserGRFinancCons.grdGR_UnboundReadDataEx]"
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim strSql      As String
  Dim objRs       As ADODB.Recordset
  Dim objGeral    As busSisMed.clsGeral
  Dim strGR       As String
  '
  blnFechar = False 'Não Pode Fechar pelo X
  blnRetorno = False
  AmpS
  Me.Height = 7320
  Me.Width = 11985
  CenterForm Me
  Report1.Connect = ConnectRpt
  Report1.ReportFileName = gsReportPath & "GRConsulta.rpt"
  'Limpar Campos
  LimparCampoMask mskDtInicio
  LimparCampoTexto txtProntuario
  LimparCampoTexto txtPrestador
  LimparCampoTexto txtFuncionario
  optGR(0).Value = True
  '
  If gsNivel = gsAdmin Then
    cmdAlterar.Enabled = True
  Else
    cmdAlterar.Enabled = False
  End If
  'Tratar Campos
  INCLUIR_VALOR_NO_MASK mskDtInicio, Format(Now, "DD/MM/YYYY"), TpMaskData
  txtProntuario.Text = strProntuario
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar, cmdExcluir, pbtnAlterar:=cmdAlterar, pbtnImprimir:=cmdImprimir
  LerFigurasAvulsas cmdConsultar, "FILTRAR.ICO", "filtrarDown.ico", "Consultar GR"
  LerFigurasAvulsas cmdImprimirTodas, "Impressora.ico", "ImpressoraDown.ico", "Imprimir Todas as GR´s"
  'Obter campos
  Set objGeral = New busSisMed.clsGeral
  '
  strDataIni = Format(Now, "DD/MM/YYYY 00:00")
  strDataFim = Format(DateAdd("d", 1, Now), "DD/MM/YYYY 00:00")
  strGR = IIf(optGR(1).Value, "P", "T")
  '
  GR_COLUNASMATRIZ = grdGR.Columns.Count
  GR_LINHASMATRIZ = 0
  GR_MontaMatriz strDataIni, strDataFim, strGR
  grdGR.ApproxCount = GR_LINHASMATRIZ
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub

Private Sub mskDtInicio_GotFocus()
  Seleciona_Conteudo_Controle mskDtInicio
End Sub
Private Sub mskDtInicio_LostFocus()
  Pintar_Controle mskDtInicio, tpCorContr_Normal
End Sub

Private Sub txtFuncionario_GotFocus()
  Seleciona_Conteudo_Controle txtFuncionario
End Sub
Private Sub txtFuncionario_LostFocus()
  Pintar_Controle txtFuncionario, tpCorContr_Normal
End Sub

Private Sub txtPrestador_GotFocus()
  Seleciona_Conteudo_Controle txtPrestador
End Sub
Private Sub txtPrestador_LostFocus()
  Pintar_Controle txtPrestador, tpCorContr_Normal
End Sub

Private Sub txtProntuario_GotFocus()
  Seleciona_Conteudo_Controle txtProntuario
End Sub
Private Sub txtProntuario_LostFocus()
  Pintar_Controle txtProntuario, tpCorContr_Normal
End Sub

