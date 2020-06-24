VERSION 5.00
Object = "{00028C01-0000-0000-0000-000000000046}#1.0#0"; "DBGRID32.OCX"
Object = "{F6125AB1-8AB1-11CE-A77F-08002B2F4E98}#2.0#0"; "MSRDC20.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmSaConsultar 
   Caption         =   "Consultar SA"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9375
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6360
   ScaleMode       =   0  'User
   ScaleWidth      =   13586.96
   Begin TabDlg.SSTab TabSA 
      Height          =   6075
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   9345
      _ExtentX        =   16484
      _ExtentY        =   10716
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Dados da SA"
      TabPicture(0)   =   "frmSaConsultar.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Ítens da SA"
      TabPicture(1)   =   "frmSaConsultar.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "dbgPessoa"
      Tab(1).Control(1)=   "rdcItens"
      Tab(1).Control(2)=   "SelArea"
      Tab(1).Control(3)=   "SelAssunto"
      Tab(1).Control(4)=   "Label4"
      Tab(1).Control(5)=   "Label6"
      Tab(1).Control(6)=   "Shape1"
      Tab(1).ControlCount=   7
      Begin MSDBGrid.DBGrid dbgPessoa 
         Bindings        =   "frmSaConsultar.frx":0038
         Height          =   1965
         Left            =   -74550
         OleObjectBlob   =   "frmSaConsultar.frx":004F
         TabIndex        =   11
         Top             =   1740
         Width           =   8505
      End
      Begin MSRDC.MSRDC rdcItens 
         Height          =   330
         Left            =   -74190
         Top             =   3240
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
         Caption         =   "rdcItens"
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
      Begin VB.ComboBox SelArea 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73800
         TabIndex        =   13
         Top             =   1110
         Width           =   1965
      End
      Begin VB.ComboBox SelAssunto 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   -73800
         TabIndex        =   12
         Top             =   660
         Width           =   1965
      End
      Begin VB.Frame Frame1 
         Height          =   3255
         Left            =   180
         TabIndex        =   1
         Top             =   540
         Width           =   8955
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
            Height          =   330
            Left            =   1950
            TabIndex        =   6
            Top             =   300
            Width           =   1950
         End
         Begin VB.TextBox txtDtaSa 
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1950
            TabIndex        =   5
            Top             =   1080
            Width           =   1950
         End
         Begin VB.ComboBox SelSa 
            Appearance      =   0  'Flat
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   330
            Left            =   1950
            TabIndex        =   4
            Top             =   690
            Width           =   1965
         End
         Begin VB.OptionButton OpTipoSa 
            Caption         =   "Normal"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   0
            Left            =   2040
            TabIndex        =   3
            Top             =   1530
            Width           =   795
         End
         Begin VB.OptionButton OpTipoSa 
            Caption         =   "Complementar"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Index           =   1
            Left            =   2970
            TabIndex        =   2
            Top             =   1530
            Width           =   1395
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Número da SA:"
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
            TabIndex        =   10
            Top             =   765
            Width           =   1695
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Data:"
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
            TabIndex        =   9
            Top             =   1155
            Width           =   1695
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
            TabIndex        =   8
            Top             =   375
            Width           =   1695
         End
         Begin VB.Label Label5 
            Alignment       =   1  'Right Justify
            Caption         =   "Tipo de Auditoria:"
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
            TabIndex        =   7
            Top             =   1500
            Width           =   1695
         End
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Assunto: "
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
         Left            =   -74820
         TabIndex        =   15
         Top             =   690
         Width           =   915
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Área: "
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
         Left            =   -74850
         TabIndex        =   14
         Top             =   1170
         Width           =   885
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H8000000F&
         BackStyle       =   1  'Opaque
         BorderStyle     =   0  'Transparent
         Height          =   5685
         Left            =   -75000
         Top             =   330
         Width           =   9345
      End
   End
End
Attribute VB_Name = "frmSAConsultar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim NumProcesso As String
Dim sig_orgao_processo As String
Dim seq_processo As Integer
Dim ano_processo As Integer
Dim seq_sa As Integer
Dim seq_sa_complementar As String
Dim seq_assunto As Variant
Dim seq_area As Variant
Dim Rs As ADODB.Recordset
Dim Sql As String


Private Sub Command1_Click()
Unload frmSAConsultar

End Sub
Private Sub Form_Load()

On Error GoTo TrataErro
Set Rs = New ADODB.Recordset


frmSAConsultar.Width = 9500
frmSAConsultar.Height = 4455

'********************************************************
'Para carregar os dados para esse processo
'********************************************************
Sql = "select sig_orgao_processo, seq_processo, ano_processo" & _
      "  from processo_auditoria "
      
Set Rs = ExecutaSqlRs(Sql)

If Not Rs.EOF Then
    NumProcesso = "PA" & "-" & Rs("sig_orgao_processo") & "-" & Right("000" & Rs("seq_processo"), 3) & "/" & Rs("ano_processo")
    sig_orgao_processo = Rs("sig_orgao_processo")
    seq_processo = Rs("seq_processo")
    ano_processo = Rs("ano_processo")
    frmSAConsultar.txtNumProcesso.Text = NumProcesso
End If
Rs.Close
'********************************************************
'Para pegar os dados da(s) SA(s) desse processo
'********************************************************


Sql = "select seq_sa, seq_sa_complementar from sa where ano_processo = " & ano_processo & " and sig_orgao_processo='" & sig_orgao_processo & _
              "' and seq_Processo=" & seq_processo & " and ind_desativacao='A'"
Set Rs = ExecutaSqlRs(Sql)

Call CarregaSA(Rs, frmSAConsultar, sig_orgao_processo, seq_processo, ano_processo)
Rs.Close
'********************************************************
'Para pegar os dados do(s) assunto
'********************************************************
Sql = " select seq_assunto,descr_assunto,ind_status from " & _
        " assunto where ind_status='S'"
Set Rs = ExecutaSqlRs(Sql)

Call CarregaAssunto(Rs, frmSAConsultar)
Rs.Close

'********************************************************
'Para pegar os dados da(s) áreas
'********************************************************
Sql = " select seq_area,descr_area,ind_status   " & _
        " From area_auditoria " & _
        " where ind_status='S'"
Set Rs = ExecutaSqlRs(Sql)
Call CarregaArea(Rs, frmSAConsultar)
Rs.Close



Set cn = Nothing

Exit Sub
TrataErro:
    MsgBox Err.Description, vbCritical, "Erro no aplicativo"
    Set cn = Nothing
    Set Rs = Nothing
End Sub

Private Sub SelArea_Click()
    If SelAssunto.ListIndex = -1 Then
        seq_assunto = ""
    Else
        seq_assunto = SelAssunto.ItemData(SelAssunto.ListIndex)
    End If
    If SelArea.ListIndex = -1 Then
        seq_area = ""
    Else
        seq_area = SelArea.ItemData(SelArea.ListIndex)
    End If
    Call SelecionaItens(sig_orgao_processo, seq_processo, ano_processo, seq_sa, seq_sa_complementar, seq_assunto, seq_area)
End Sub

Private Sub SelAssunto_Click()
Dim seq_assunto
Dim seq_area
    If SelAssunto.ListIndex = -1 Then
        seq_assunto = ""
    Else
        seq_assunto = SelAssunto.ItemData(SelAssunto.ListIndex)
    End If
    If SelArea.ListIndex = -1 Then
        seq_area = ""
    Else
        seq_area = SelArea.ItemData(SelArea.ListIndex)
    End If
    Call SelecionaItens(sig_orgao_processo, seq_processo, ano_processo, seq_sa, seq_sa_complementar, seq_assunto, seq_area)
End Sub

Private Sub SelSa_Click()
Dim CodSa As String
Dim dta_sa As String
Dim ind_complementar
Dim Vet() As String
    
On Error GoTo TrataErro
    If SelSa.ListIndex = -1 Then
        Exit Sub
    End If
    Set Rs = New ADODB.Recordset
    CodSa = SelSa.List(SelSa.ListIndex)
    If CodSa <> "" Then
        ReDim Vet(2)
        If InStr(CodSa, "-") > 0 Then
            Vet = Split(CodSa, "-")
            seq_sa = Vet(0)
            seq_sa_complementar = Vet(1)
        Else
            seq_sa = CodSa
            seq_sa_complementar = "0"
        End If
            
        Sql = "select dta_sa, ind_complementar from sa " & _
                " where sig_orgao_processo='" & sig_orgao_processo & "'" & _
                " and seq_processo =" & seq_processo & " and ano_processo=" & _
                ano_processo & " and seq_sa = " & seq_sa & _
                " and seq_sa_complementar='" & seq_sa_complementar & "'"
            Set Rs = ExecutaSqlRs(Sql)
            If Not Rs.EOF Then
                dta_sa = Format(Rs("dta_sa"), "dd/mm/yyyy")
                ind_complementar = Rs("ind_complementar")
            End If
            Rs.Close
            frmSAConsultar.txtDtaSa = dta_sa
            If ind_complementar = "S" Then
                frmSAConsultar.OpTipoSa(1) = Checked
            Else
                frmSAConsultar.OpTipoSa(0) = Checked
            End If
            
            Call SelecionaItens(sig_orgao_processo, seq_processo, ano_processo, seq_sa, seq_sa_complementar, seq_assunto, seq_area)
            
    End If
    Exit Sub

TrataErro:
    MsgBox Err.Description, vbExclamation

End Sub

