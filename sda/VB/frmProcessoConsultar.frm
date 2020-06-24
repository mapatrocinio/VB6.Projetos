VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Begin VB.Form frmProcessoConsultar 
   Caption         =   "Consultar Processo"
   ClientHeight    =   4050
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9795
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4050
   ScaleMode       =   0  'User
   ScaleWidth      =   14195.65
   Begin VB.CommandButton Command1 
      Caption         =   "Fechar"
      Height          =   375
      Left            =   8430
      TabIndex        =   9
      Top             =   3480
      Width           =   735
   End
   Begin MSRDC.MSRDC rdcPessoa 
      Height          =   330
      Left            =   360
      Top             =   3840
      Visible         =   0   'False
      Width           =   2355
      _ExtentX        =   4154
      _ExtentY        =   582
      _Version        =   393216
      Options         =   0
      CursorDriver    =   0
      BOFAction       =   0
      EOFAction       =   0
      RecordsetType   =   1
      LockType        =   3
      QueryType       =   0
      Prompt          =   3
      Appearance      =   1
      QueryTimeout    =   30
      RowsetSize      =   100
      LoginTimeout    =   15
      KeysetSize      =   0
      MaxRows         =   0
      ErrorThreshold  =   -1
      BatchSize       =   15
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Enabled         =   -1  'True
      ReadOnly        =   0   'False
      Appearance      =   -1  'True
      DataSourceName  =   "sda"
      RecordSource    =   ""
      UserName        =   ""
      Password        =   ""
      Connect         =   ""
      LogMessages     =   ""
      Caption         =   "rdcPessoa"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Caption         =   "Processo"
      Height          =   1725
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8025
      Begin VB.TextBox txtDtaInicio 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   6
         Top             =   720
         Width           =   1950
      End
      Begin VB.TextBox TxtDtaFim 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   5
         Top             =   1140
         Width           =   1950
      End
      Begin VB.TextBox txtNumProcesso 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1950
         TabIndex        =   3
         Top             =   300
         Width           =   1950
      End
      Begin MSDBGrid.DBGrid dbgPessoa 
         Bindings        =   "frmProcessoConsultar.frx":0000
         Height          =   1875
         Left            =   0
         OleObjectBlob   =   "frmProcessoConsultar.frx":0018
         TabIndex        =   7
         Top             =   1680
         Width           =   9255
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Número do Processo:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Fim:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   2
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Início:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   1
         Top             =   780
         Width           =   1695
      End
   End
   Begin MSDBGrid.DBGrid DBGrid1 
      Bindings        =   "frmProcessoConsultar.frx":10F1
      Height          =   1845
      Left            =   120
      OleObjectBlob   =   "frmProcessoConsultar.frx":1109
      TabIndex        =   8
      Top             =   2040
      Width           =   8055
   End
End
Attribute VB_Name = "frmProcessoConsultar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Command1_Click()
Unload frmProcessoConsultar

End Sub

Private Sub Form_Load()
Dim i As Integer
Dim Rs As ADODB.Recordset
Dim dta_inicio As String
Dim dta_fim As String
Dim NumProcesso As String
Dim Sql As String

On Error GoTo TrataErro
Set Rs = New ADODB.Recordset

frmProcessoConsultar.Width = 9500
frmProcessoConsultar.Height = 4455

'********************************************************
'Para carregar os dados para esse processo
'********************************************************
Sql = "select sig_orgao_processo, seq_processo, ano_processo," & _
      " dta_inicio, dta_fim from processo_auditoria "
      
Set Rs = ExecutaSqlRs(Sql)

If Not Rs.EOF Then
    NumProcesso = "PA" & "-" & Rs("sig_orgao_processo") & "-" & Right("000" & Rs("seq_processo"), 3) & "/" & Rs("ano_processo")
    If Not IsNull(Rs("dta_inicio")) Then
        dta_inicio = Rs("dta_inicio")
    End If
    If Not IsNull(Rs("dta_fim")) Then
        dta_fim = Rs("dta_fim")
    End If
    frmProcessoConsultar.txtNumProcesso.Text = NumProcesso
    frmProcessoConsultar.txtDtaInicio = dta_inicio
    frmProcessoConsultar.TxtDtaFim = dta_fim
    
End If
Rs.Close
Set cn = Nothing
'********************************************************
'Para carregar os auditores cadastrados para esse processo
'********************************************************
rdcPessoa.DataSourceName = "sda"
rdcPessoa.Sql = "select * from pessoa_inmetro"
rdcPessoa.Refresh
Exit Sub
TrataErro:
    MsgBox Err.Description, vbCritical, "Erro no aplicativo"
    Set cn = Nothing
    Set Rs = Nothing
End Sub
