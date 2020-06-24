VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "msmask32.ocx"
Begin VB.Form frmUserFiltroCheqCons 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Aplicar filtro em cliente para vizualização de cheques"
   ClientHeight    =   2535
   ClientLeft      =   30
   ClientTop       =   330
   ClientWidth     =   6705
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2535
   ScaleWidth      =   6705
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton optFiltro 
      Caption         =   "Alterar após seleção"
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   5
      Top             =   360
      Width           =   2415
   End
   Begin VB.OptionButton optFiltro 
      Caption         =   "Consultar após seleção"
      Height          =   255
      Index           =   0
      Left            =   1200
      TabIndex        =   4
      Top             =   120
      Value           =   -1  'True
      Width           =   2415
   End
   Begin VB.CommandButton cmdCancelar 
      Cancel          =   -1  'True
      Caption         =   "ESC"
      Height          =   880
      Left            =   3240
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1560
      Width           =   1215
   End
   Begin VB.CommandButton cmdConfirmar 
      Caption         =   "ENTER"
      Default         =   -1  'True
      Height          =   880
      Left            =   2040
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1560
      Width           =   1215
   End
   Begin MSMask.MaskEdBox mskCPF 
      Height          =   255
      Left            =   1200
      TabIndex        =   0
      Top             =   720
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   450
      _Version        =   393216
      AutoTab         =   -1  'True
      MaxLength       =   12
      Mask            =   "#########/##"
      PromptChar      =   "_"
   End
   Begin VB.Label Label14 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "C.P.F. :"
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
      Index           =   0
      Left            =   -360
      TabIndex        =   1
      Top             =   720
      Width           =   1455
   End
End
Attribute VB_Name = "frmUserFiltroCheqCons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public strCPF     As String
Public strMsg     As String
Public QuemChamou As Integer
Public GravarCPF  As Boolean
Public sNumeroAptoPrinc As String

'Assume 1 -  chamada de locação
Public lngLOCACAOID As Long

Option Explicit


Private Sub cmdCancelar_Click()
  strMsg = ""
  Unload Me
End Sub

Private Sub cmdConfirmar_Click()
  On Error GoTo trata
  Dim clsChq              As busSisContas.clsCheque
  Dim clsGer              As busSisContas.clsGeral
  Dim objRs               As ADODB.Recordset
  Dim possuichqdevolvido  As Boolean
  Dim strSql              As String
  Dim DtAtualMenosNDias   As Date
  
  '
  If Len(mskCPF.ClipText) <> 0 Then
    If Not TestaCPF(mskCPF.Text) Then
      MsgBox "Nro. do CPF Inválido", vbOKOnly, TITULOSISTEMA
      SetarFoco mskCPF
      Exit Sub
    End If
  Else
    MsgBox "Informe o nro. do CPF", vbOKOnly, TITULOSISTEMA
    SetarFoco mskCPF
    Exit Sub
  End If
  '
  If QuemChamou <> 1 Then
    Set clsChq = New busSisContas.clsCheque
    '
    Set objRs = clsChq.ListarClientePorCPF(mskCPF.Text, 0)
    If objRs.EOF Then
      MsgBox "Nro. do CPF não cadastrado", vbOKOnly, TITULOSISTEMA
      SetarFoco mskCPF
      Exit Sub
'''      If (MsgBox(", deseja incluir o cliente para este CPF agora?", vbYesNo, TITULOSISTEMA) = vbNo) Then
'''        objRs.Close
'''        Set objRs = Nothing
'''        Set clsChq = Nothing
'''        Exit Sub
'''      End If
'''      'Captura dados do Cliente
'''      frmUserClienteInc.lngCLIENTEID = 0
'''      frmUserClienteInc.mskCPF.Text = mskCPF.Text
'''      frmUserClienteInc.Status = tpStatus_Incluir
    Else
      'Captura dados do Cliente
      frmUserClienteInc.lngCLIENTEID = objRs.Fields("PKID").Value
      If optFiltro(0).Value Then 'Consulta
        frmUserClienteInc.Status = tpStatus_Consultar
      Else
        frmUserClienteInc.Status = tpStatus_Alterar
      End If
    End If
    '
    objRs.Close
    Set objRs = Nothing
    Set clsChq = Nothing
  Else
    Set clsGer = New busSisContas.clsGeral
    possuichqdevolvido = False
    strMsg = ""
    strSql = "Select CHEQUE.* from CLIENTE INNER JOIN CHEQUE ON CLIENTE.PKID = CHEQUE.CLIENTEID WHERE CLIENTE.CPF  = '" & mskCPF.Text & "' AND STATUS = 'D'"
    Set objRs = clsGer.ExecutarSQL(strSql)
    '
    GravarCPF = False
    If Not objRs.EOF Then
      possuichqdevolvido = True
      GravarCPF = True
      strMsg = "CLIENTE COM CHEQUE DEVOLVIDO"
      If gbPedirSenhaSupLibChqReceb Then
        'PEDIR SENHA SUPERIOR
        strMsg = strMsg & vbCrLf & vbCrLf & "Será pedido senha superior no recebimento."
        INCLUI_LOG_UNIDADE MODOALTERAR, lngLOCACAOID, "Cheque-cliente com problemas", "Unidade " & sNumeroAptoPrinc & " - CPF Nr. " & mskCPF.Text, "", "", "", ""
      Else
        strMsg = strMsg & vbCrLf & vbCrLf & "Contacte o gerente, pois o cliente não poderá efetuar o pagamento com cheque."
        INCLUI_LOG_UNIDADE MODOALTERAR, lngLOCACAOID, "Cheque-cliente com problemas", "Unidade " & sNumeroAptoPrinc & " - CPF Nr. " & mskCPF.Text, "", "", "", ""
      End If
    End If
    objRs.Close
    Set objRs = Nothing
    'TRATA CHEQUES BONS
    If Len(strMsg) = 0 Then
      If Not possuichqdevolvido Then
        If gbTrabComChequesBons Then
          'SELECIONA CHEQUES COMPENSADOS (BONS)
          'CALCULA DATA ATUAL - N DIAS DE COMPENSAÇÃO
          DtAtualMenosNDias = DateAdd("d", giQtdDiasParaCompensar * (-1), Now)
          strSql = "Select COUNT(*) AS QTDCHQCOMP from CLIENTE INNER JOIN CHEQUE ON CLIENTE.PKID = CHEQUE.CLIENTEID WHERE CLIENTE.CPF  = '" & mskCPF.Text & "' AND STATUS = 'C' AND DTRECEBIMENTO <= " & Formata_Dados(Format(DtAtualMenosNDias, "DD/MM/YYYY"), tpDados_DataHora, tpNulo_NaoAceita)
          Set objRs = clsGer.ExecutarSQL(strSql)
          '
          If Not objRs.EOF Then
            If IsNumeric(objRs.Fields("QTDCHQCOMP").Value) Then
              If objRs.Fields("QTDCHQCOMP").Value >= giQtdChequesBons Then
                strMsg = "CLIENTE ESPECIAL" & vbCrLf & vbCrLf & "Não precisa consultar o cheque"
                'gravar log
                INCLUI_LOG_UNIDADE MODOALTERAR, lngLOCACAOID, "Cheque-cliente especial", "Unidade " & sNumeroAptoPrinc & " - CPF Nr. " & mskCPF.Text, "", "", "", ""
                GravarCPF = True
              Else
                If MsgBox("CONSULTAR CHEQUE" & vbCrLf & vbCrLf & "Você já fez a consulta?", vbYesNo, TITULOSISTEMA) = vbYes Then
                  'gravar log
                  INCLUI_LOG_UNIDADE MODOALTERAR, lngLOCACAOID, "Consultar cheque", "Unidade " & sNumeroAptoPrinc & " - CPF Nr. " & mskCPF.Text, "", "", "", ""
                  GravarCPF = True
                Else
                  'strMsg = "Haverá a necessidade de consulta do cheque. Volte ao recebimento após ter feito a consulta." & vbCrLf & vbCrLf & "O sistema retornará a tela de entrada sem efetuar o recebimento." & vbCrLf
                End If
              End If
            End If
          End If
          objRs.Close
          Set objRs = Nothing
        End If
      End If
    End If
    '
    'Verifica se irá gravar CPF
    If GravarCPF Then
      strCPF = mskCPF.Text
    End If
    Set clsGer = Nothing
  End If
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub Form_Activate()
  SetarFoco mskCPF
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  AmpS
  strCPF = ""
  CenterForm Me
  LerFiguras Me, tpBmp_Vazio, cmdConfirmar, cmdCancelar
  If QuemChamou <> 1 Then
    If gsNivel = gsDiretor Or gsNivel = gsGerente Or gsNivel = gsAdmin Then
      optFiltro(0).Visible = True
      optFiltro(1).Visible = True
      optFiltro(1).Value = True
    Else
      optFiltro(0).Visible = False
      optFiltro(1).Visible = False
    End If
  Else
    optFiltro(0).Visible = False
    optFiltro(1).Visible = False
  End If
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub


Private Sub mskCPF_Change()
  On Error GoTo trata
  If Len(mskCPF.ClipText) = 11 Then
    cmdConfirmar_Click
  End If
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Private Sub mskCPF_LostFocus()
  Seleciona_Conteudo_Controle mskCPF
End Sub

Private Sub mskCPF_GotFocus()
  Seleciona_Conteudo_Controle mskCPF
End Sub

