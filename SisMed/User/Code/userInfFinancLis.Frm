VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frmUserInfFinancLis 
   Caption         =   "Informações Financeiras do Turno"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11265
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   6165
   ScaleWidth      =   11265
   Begin VB.PictureBox Picture1 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   990
      Left            =   0
      ScaleHeight     =   990
      ScaleWidth      =   11265
      TabIndex        =   39
      TabStop         =   0   'False
      Top             =   5175
      Width           =   11265
      Begin VB.CommandButton cmdSangria 
         Caption         =   "ENTER"
         Default         =   -1  'True
         Height          =   880
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   37
         Top             =   60
         Width           =   1215
      End
      Begin VB.CommandButton cmdFechar 
         Cancel          =   -1  'True
         Caption         =   "ESC"
         Height          =   880
         Left            =   9840
         Style           =   1  'Graphical
         TabIndex        =   38
         Top             =   60
         Width           =   1215
      End
   End
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   4980
      Left            =   120
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   8784
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      TabsPerRow      =   8
      TabHeight       =   529
      BackColor       =   12632256
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "&Faturamento"
      TabPicture(0)   =   "userInfFinancLis.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "fraAluno"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "&Saldos"
      TabPicture(1)   =   "userInfFinancLis.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame2"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "&Movimentações"
      TabPicture(2)   =   "userInfFinancLis.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "grdSangria"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Cartão de &Crédito"
      TabPicture(3)   =   "userInfFinancLis.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "grdGeral"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Cartão de &Débito"
      TabPicture(4)   =   "userInfFinancLis.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "grdCartaoDeb"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "&Garçom"
      TabPicture(5)   =   "userInfFinancLis.frx":008C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "grdGarcom"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "D&espesas"
      TabPicture(6)   =   "userInfFinancLis.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "grdDespesa"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "Extra fora &Unidade"
      TabPicture(7)   =   "userInfFinancLis.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "grdExtra"
      Tab(7).ControlCount=   1
      Begin VB.Frame Frame2 
         Enabled         =   0   'False
         Height          =   4335
         Left            =   -74880
         TabIndex        =   67
         Top             =   330
         Width           =   10575
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   3495
            Index           =   1
            Left            =   120
            ScaleHeight     =   3495
            ScaleWidth      =   8535
            TabIndex        =   68
            TabStop         =   0   'False
            Top             =   240
            Width           =   8535
            Begin VB.TextBox txtSaldoFatura 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Index           =   1
               Left            =   5160
               Locked          =   -1  'True
               TabIndex        =   33
               Text            =   "txtSaldoFatura"
               Top             =   810
               Width           =   1695
            End
            Begin VB.TextBox txtSaldocartaodebs 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   32
               Text            =   "txtSaldocartaodeb"
               Top             =   840
               Width           =   1695
            End
            Begin VB.TextBox txtSaldoTotalMov 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Index           =   1
               Left            =   5160
               Locked          =   -1  'True
               TabIndex        =   34
               Text            =   "txtSaldoTotalMov"
               Top             =   1140
               Width           =   1695
            End
            Begin VB.TextBox txtSaldoDinheiro 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Index           =   1
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   28
               Text            =   "txtSaldoDinheiro"
               Top             =   120
               Width           =   1695
            End
            Begin VB.TextBox txtSaldoCheque 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Index           =   1
               Left            =   5160
               Locked          =   -1  'True
               TabIndex        =   29
               Text            =   "txtSaldoCheque"
               Top             =   120
               Width           =   1695
            End
            Begin VB.TextBox txtSaldocartao 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Index           =   1
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   30
               Text            =   "txtSaldocartao"
               Top             =   480
               Width           =   1695
            End
            Begin VB.TextBox txtSaldoPenhor 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00E0E0E0&
               Height          =   285
               Index           =   1
               Left            =   5160
               Locked          =   -1  'True
               TabIndex        =   31
               Text            =   "txtSaldoPenhor"
               Top             =   480
               Width           =   1695
            End
            Begin VB.Label Label3 
               Caption         =   "Saldo em fatura"
               Height          =   255
               Index           =   3
               Left            =   3600
               TabIndex        =   84
               Top             =   810
               Width           =   1575
            End
            Begin VB.Label Label3 
               Caption         =   "Saldo em cartão déb"
               Height          =   255
               Index           =   5
               Left            =   0
               TabIndex        =   80
               Top             =   840
               Width           =   1575
            End
            Begin VB.Label Label16 
               Caption         =   "Saldo total c/ Movimentação"
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
               Index           =   1
               Left            =   2400
               TabIndex        =   73
               Top             =   1140
               Width           =   2655
            End
            Begin VB.Label Label1 
               Caption         =   "Saldo em dinheiro"
               Height          =   255
               Index           =   1
               Left            =   30
               TabIndex        =   72
               Top             =   120
               Width           =   1335
            End
            Begin VB.Label Label2 
               Caption         =   "Saldo em cheque"
               Height          =   255
               Index           =   1
               Left            =   3600
               TabIndex        =   71
               Top             =   120
               Width           =   1455
            End
            Begin VB.Label Label3 
               Caption         =   "Saldo em cartão créd"
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   70
               Top             =   480
               Width           =   1575
            End
            Begin VB.Label Label5 
               Caption         =   "Saldo em Penhor"
               Height          =   255
               Index           =   1
               Left            =   3600
               TabIndex        =   69
               Top             =   480
               Width           =   1455
            End
         End
      End
      Begin VB.Frame fraAluno 
         Enabled         =   0   'False
         Height          =   4455
         Left            =   90
         TabIndex        =   41
         Top             =   360
         Width           =   10695
         Begin VB.PictureBox picTrava 
            BorderStyle     =   0  'None
            Height          =   4245
            Index           =   0
            Left            =   120
            ScaleHeight     =   4245
            ScaleWidth      =   10455
            TabIndex        =   42
            TabStop         =   0   'False
            Top             =   150
            Width           =   10455
            Begin VB.TextBox txtFuncionario 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   1560
               Locked          =   -1  'True
               TabIndex        =   0
               Text            =   "txtFuncionario"
               Top             =   0
               Width           =   3735
            End
            Begin VB.TextBox txtTurno 
               BackColor       =   &H00E0E0E0&
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   285
               Left            =   5280
               Locked          =   -1  'True
               TabIndex        =   1
               Text            =   "txtTurno"
               Top             =   0
               Width           =   4935
            End
            Begin VB.Frame Frame3 
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
               Height          =   1815
               Left            =   0
               TabIndex        =   43
               Top             =   240
               Width           =   10335
               Begin VB.TextBox txtSaldoFatura 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Height          =   285
                  Index           =   0
                  Left            =   1560
                  Locked          =   -1  'True
                  TabIndex        =   11
                  Text            =   "txtSaldoFatura"
                  Top             =   1140
                  Width           =   1695
               End
               Begin VB.TextBox txtSaldoDespDiverChqResg 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Left            =   8520
                  Locked          =   -1  'True
                  TabIndex        =   10
                  Text            =   "txtSaldoDespDiverChqResg"
                  Top             =   840
                  Width           =   1695
               End
               Begin VB.TextBox txtSaldocartaodeb 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Height          =   285
                  Left            =   8520
                  Locked          =   -1  'True
                  TabIndex        =   4
                  Text            =   "txtSaldocartaoDeb"
                  Top             =   150
                  Width           =   1695
               End
               Begin VB.TextBox txtSaldoTroco 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Left            =   5160
                  Locked          =   -1  'True
                  TabIndex        =   13
                  Text            =   "txtSaldoTroco"
                  Top             =   1470
                  Width           =   1695
               End
               Begin VB.TextBox txtSaldoDespDiverDin 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Left            =   8520
                  Locked          =   -1  'True
                  TabIndex        =   7
                  Text            =   "txtSaldoDespDiverDin"
                  Top             =   480
                  Width           =   1695
               End
               Begin VB.TextBox txtSaldoGorjeta 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Left            =   1560
                  Locked          =   -1  'True
                  TabIndex        =   12
                  Text            =   "txtSaldoGorjeta"
                  Top             =   1440
                  Width           =   1695
               End
               Begin VB.TextBox txtSaldoPenhor 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Height          =   285
                  Index           =   0
                  Left            =   5160
                  Locked          =   -1  'True
                  TabIndex        =   9
                  Text            =   "txtSaldoPenhor"
                  Top             =   810
                  Width           =   1695
               End
               Begin VB.TextBox txtSaldoTotal 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Left            =   8520
                  Locked          =   -1  'True
                  TabIndex        =   14
                  Text            =   "txtSaldoTotal"
                  Top             =   1470
                  Width           =   1695
               End
               Begin VB.TextBox txtSaldocartao 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Height          =   285
                  Index           =   0
                  Left            =   1560
                  Locked          =   -1  'True
                  TabIndex        =   8
                  Text            =   "txtSaldocartao"
                  Top             =   810
                  Width           =   1695
               End
               Begin VB.TextBox txtSaldoCheque 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Height          =   285
                  Index           =   0
                  Left            =   5160
                  Locked          =   -1  'True
                  TabIndex        =   6
                  Text            =   "txtSaldoCheque"
                  Top             =   480
                  Width           =   1695
               End
               Begin VB.TextBox txtSaldoDinheiro 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Height          =   285
                  Index           =   0
                  Left            =   1560
                  Locked          =   -1  'True
                  TabIndex        =   5
                  Text            =   "txtSaldoDinheiro"
                  Top             =   480
                  Width           =   1695
               End
               Begin VB.TextBox txtVrInicial 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Height          =   285
                  Left            =   5160
                  Locked          =   -1  'True
                  TabIndex        =   3
                  Text            =   "txtVrInicial"
                  Top             =   150
                  Width           =   1695
               End
               Begin VB.PictureBox Picture2 
                  BorderStyle     =   0  'None
                  Enabled         =   0   'False
                  Height          =   255
                  Left            =   120
                  ScaleHeight     =   255
                  ScaleWidth      =   3255
                  TabIndex        =   45
                  TabStop         =   0   'False
                  Top             =   150
                  Width           =   3255
                  Begin MSMask.MaskEdBox mskData 
                     Height          =   255
                     Index           =   0
                     Left            =   1440
                     TabIndex        =   2
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
                  Begin VB.Label Da 
                     Caption         =   "Dt./Hr.  Turno"
                     Height          =   255
                     Left            =   0
                     TabIndex        =   46
                     Top             =   0
                     Width           =   1335
                  End
               End
               Begin VB.Frame Frame5 
                  Caption         =   "Sexo"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   615
                  Left            =   840
                  TabIndex        =   44
                  Top             =   3480
                  Width           =   2295
               End
               Begin VB.Label Label3 
                  Caption         =   "Fatura"
                  Height          =   255
                  Index           =   2
                  Left            =   120
                  TabIndex        =   83
                  Top             =   1140
                  Width           =   1455
               End
               Begin VB.Label Label22 
                  Caption         =   "Desp. Div. Chq. Resg."
                  Height          =   255
                  Left            =   6960
                  TabIndex        =   82
                  Top             =   840
                  Width           =   1605
               End
               Begin VB.Label Label3 
                  Caption         =   "Cartão de Débito"
                  Height          =   255
                  Index           =   4
                  Left            =   6960
                  TabIndex        =   77
                  Top             =   150
                  Width           =   1455
               End
               Begin VB.Label Label9 
                  Caption         =   "Troco"
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
                  Left            =   3600
                  TabIndex        =   56
                  Top             =   1470
                  Width           =   1455
               End
               Begin VB.Label Label8 
                  Caption         =   "Desp. Div. Din."
                  Height          =   255
                  Left            =   6960
                  TabIndex        =   55
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.Label Label7 
                  Caption         =   "Gorjetas"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   54
                  Top             =   1440
                  Width           =   1455
               End
               Begin VB.Label Label5 
                  Caption         =   "Penhor"
                  Height          =   255
                  Index           =   0
                  Left            =   3600
                  TabIndex        =   53
                  Top             =   810
                  Width           =   1455
               End
               Begin VB.Label Label4 
                  Caption         =   "Total"
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
                  Left            =   6960
                  TabIndex        =   52
                  Top             =   1470
                  Width           =   1455
               End
               Begin VB.Label Label3 
                  Caption         =   "Cartão de Crédito"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   51
                  Top             =   810
                  Width           =   1455
               End
               Begin VB.Label Label2 
                  Caption         =   "Cheque"
                  Height          =   255
                  Index           =   0
                  Left            =   3600
                  TabIndex        =   50
                  Top             =   480
                  Width           =   1455
               End
               Begin VB.Label Label1 
                  Caption         =   "Dinheiro"
                  Height          =   255
                  Index           =   0
                  Left            =   120
                  TabIndex        =   49
                  Top             =   480
                  Width           =   1335
               End
               Begin VB.Label Label6 
                  Caption         =   "Vr. inicial caixa"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   48
                  Top             =   150
                  Width           =   1215
               End
            End
            Begin VB.Frame Frame1 
               Height          =   2115
               Left            =   0
               TabIndex        =   57
               Top             =   2040
               Width           =   10335
               Begin VB.TextBox txtDepFat 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Height          =   285
                  Left            =   8520
                  Locked          =   -1  'True
                  TabIndex        =   26
                  Text            =   "txtDepFat"
                  Top             =   1170
                  Width           =   1695
               End
               Begin VB.TextBox txtRetFat 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Left            =   8520
                  Locked          =   -1  'True
                  TabIndex        =   20
                  Text            =   "txtRetFat"
                  Top             =   510
                  Width           =   1695
               End
               Begin VB.TextBox txtDepCarDeb 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Height          =   285
                  Left            =   5160
                  Locked          =   -1  'True
                  TabIndex        =   25
                  Text            =   "txtDepCarDeb"
                  Top             =   1170
                  Width           =   1695
               End
               Begin VB.TextBox txtRetCarDeb 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Left            =   5160
                  Locked          =   -1  'True
                  TabIndex        =   19
                  Text            =   "txtRetCarDeb"
                  Top             =   510
                  Width           =   1695
               End
               Begin VB.TextBox txtDepPen 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Height          =   285
                  Left            =   1560
                  Locked          =   -1  'True
                  TabIndex        =   24
                  Text            =   "txtDepPen"
                  Top             =   1170
                  Width           =   1695
               End
               Begin VB.TextBox txtRetPen 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Left            =   1560
                  Locked          =   -1  'True
                  TabIndex        =   18
                  Text            =   "txtRetPen"
                  Top             =   510
                  Width           =   1695
               End
               Begin VB.TextBox txtSaldoTotalMov 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   285
                  Index           =   0
                  Left            =   8520
                  Locked          =   -1  'True
                  TabIndex        =   27
                  Text            =   "txtSaldoTotalMov"
                  Top             =   1680
                  Width           =   1695
               End
               Begin VB.TextBox txtDepCar 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Height          =   285
                  Left            =   8520
                  Locked          =   -1  'True
                  TabIndex        =   23
                  Text            =   "txtDepCar"
                  Top             =   840
                  Width           =   1695
               End
               Begin VB.TextBox txtDepChq 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Height          =   285
                  Left            =   5160
                  Locked          =   -1  'True
                  TabIndex        =   22
                  Text            =   "txtDepChq"
                  Top             =   840
                  Width           =   1695
               End
               Begin VB.TextBox txtDepDin 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  Height          =   285
                  Left            =   1560
                  Locked          =   -1  'True
                  TabIndex        =   21
                  Text            =   "txtDepDin"
                  Top             =   840
                  Width           =   1695
               End
               Begin VB.TextBox txtRetCar 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Left            =   8520
                  Locked          =   -1  'True
                  TabIndex        =   17
                  Text            =   "txtRetCar"
                  Top             =   180
                  Width           =   1695
               End
               Begin VB.TextBox txtRetChq 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Left            =   5160
                  Locked          =   -1  'True
                  TabIndex        =   16
                  Text            =   "txtRetChq"
                  Top             =   180
                  Width           =   1695
               End
               Begin VB.TextBox txtRetDin 
                  Alignment       =   1  'Right Justify
                  BackColor       =   &H00E0E0E0&
                  ForeColor       =   &H000000FF&
                  Height          =   285
                  Left            =   1560
                  Locked          =   -1  'True
                  TabIndex        =   15
                  Text            =   "txtRetDin"
                  Top             =   180
                  Width           =   1695
               End
               Begin VB.Label Label24 
                  Caption         =   "Deposito Fatura"
                  Height          =   255
                  Left            =   6960
                  TabIndex        =   86
                  Top             =   1170
                  Width           =   1335
               End
               Begin VB.Label Label23 
                  Caption         =   "Retirada Fatura"
                  Height          =   255
                  Left            =   6960
                  TabIndex        =   85
                  Top             =   510
                  Width           =   1215
               End
               Begin VB.Label Label20 
                  Caption         =   "Deposito Cart Déb"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   79
                  Top             =   1170
                  Width           =   1515
               End
               Begin VB.Label Label19 
                  Caption         =   "Retirada Cartão Déb"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   78
                  Top             =   510
                  Width           =   1515
               End
               Begin VB.Label Label18 
                  Caption         =   "Deposito Penhor"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   66
                  Top             =   1170
                  Width           =   1335
               End
               Begin VB.Label Label17 
                  Caption         =   "Retirada Penhor"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   65
                  Top             =   510
                  Width           =   1215
               End
               Begin VB.Label Label16 
                  Caption         =   "Total com Movimentação"
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
                  Left            =   5760
                  TabIndex        =   64
                  Top             =   1680
                  Width           =   2655
               End
               Begin VB.Label Label15 
                  Caption         =   "Deposito Cartão Créd"
                  Height          =   255
                  Left            =   6960
                  TabIndex        =   63
                  Top             =   840
                  Width           =   1515
               End
               Begin VB.Label Label14 
                  Caption         =   "Deposito Cheque"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   62
                  Top             =   840
                  Width           =   1455
               End
               Begin VB.Label Label13 
                  Caption         =   "Deposito Dinheiro"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   61
                  Top             =   840
                  Width           =   1455
               End
               Begin VB.Label Label12 
                  Caption         =   "Retirada Cartão Créd"
                  Height          =   255
                  Left            =   6960
                  TabIndex        =   60
                  Top             =   180
                  Width           =   1515
               End
               Begin VB.Label Label11 
                  Caption         =   "Retirada Cheque"
                  Height          =   255
                  Left            =   3600
                  TabIndex        =   59
                  Top             =   180
                  Width           =   1455
               End
               Begin VB.Label Label10 
                  Caption         =   "Retirada Dinheiro"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   58
                  Top             =   180
                  Width           =   1455
               End
            End
            Begin VB.Label Label21 
               Caption         =   "Turno Corrente"
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
               TabIndex        =   47
               Top             =   0
               Width           =   1455
            End
         End
      End
      Begin TrueDBGrid60.TDBGrid grdGeral 
         Height          =   3975
         Left            =   -74880
         OleObjectBlob   =   "userInfFinancLis.frx":00E0
         TabIndex        =   36
         Top             =   420
         Width           =   10335
      End
      Begin TrueDBGrid60.TDBGrid grdSangria 
         Height          =   3975
         Left            =   -74880
         OleObjectBlob   =   "userInfFinancLis.frx":2AFA
         TabIndex        =   35
         Top             =   420
         Width           =   10695
      End
      Begin TrueDBGrid60.TDBGrid grdGarcom 
         Height          =   3975
         Left            =   -74880
         OleObjectBlob   =   "userInfFinancLis.frx":9EAA
         TabIndex        =   74
         Top             =   420
         Width           =   10335
      End
      Begin TrueDBGrid60.TDBGrid grdCartaoDeb 
         Height          =   3975
         Left            =   -74880
         OleObjectBlob   =   "userInfFinancLis.frx":CEA1
         TabIndex        =   75
         Top             =   420
         Width           =   10335
      End
      Begin TrueDBGrid60.TDBGrid grdDespesa 
         Height          =   3975
         Left            =   -74880
         OleObjectBlob   =   "userInfFinancLis.frx":F8BF
         TabIndex        =   76
         Top             =   420
         Width           =   10335
      End
      Begin TrueDBGrid60.TDBGrid grdExtra 
         Height          =   3975
         Left            =   -74880
         OleObjectBlob   =   "userInfFinancLis.frx":12E67
         TabIndex        =   81
         Top             =   420
         Width           =   10335
      End
   End
End
Attribute VB_Name = "frmUserInfFinancLis"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim GER_COLUNASMATRIZ            As Long
Dim GER_LINHASMATRIZ             As Long
Private GER_Matriz()             As String

Dim RESCART_COLUNASMATRIZ        As Long
Dim RESCART_LINHASMATRIZ         As Long
Private RESCART_Matriz()         As String

Dim RESCARTDEB_COLUNASMATRIZ     As Long
Dim RESCARTDEB_LINHASMATRIZ      As Long
Private RESCARTDEB_Matriz()      As String

Dim GARCOM_COLUNASMATRIZ         As Long
Dim GARCOM_LINHASMATRIZ          As Long
Private GARCOM_Matriz()          As String

Dim DESP_COLUNASMATRIZ         As Long
Dim DESP_LINHASMATRIZ          As Long
Private DESP_Matriz()          As String

Dim EXTR_COLUNASMATRIZ         As Long
Dim EXTR_LINHASMATRIZ          As Long
Private EXTR_Matriz()          As String

Option Explicit

Private Sub cmdFechar_Click()
On Error GoTo trata
  AmpS
  Unload Me
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Private Sub cmdSangria_Click()
  On Error GoTo trata
  Dim vrCaixaInicial          As Currency
  Dim vrCheque                As Currency
  Dim vrEspecie               As Currency
  Dim vrCartao                As Currency
  Dim vrCartaoDeb             As Currency
  Dim vrPenhor                As Currency
  Dim vrFatura                As Currency
  Dim vrGorjeta               As Currency
  Dim vrSaldoDespDiveDin      As Currency
  Dim vrSaldoDespDiveChqResg  As Currency
  Dim vrTroco                 As Currency
  '
  Dim vrDepDin      As Currency
  Dim vrDepChq      As Currency
  Dim vrDepCar      As Currency
  Dim vrDepCarDeb   As Currency
  Dim vrDepPen      As Currency
  Dim vrDepFat      As Currency
  Dim vrRetDin      As Currency
  Dim vrRetChq      As Currency
  Dim vrRetCar      As Currency
  Dim vrRetCarDeb   As Currency
  Dim vrRetPen      As Currency
  Dim vrRetFat      As Currency
  '
  Dim vrTotComMov As Currency
  '
  Dim vrTotComMovDin      As Currency
  Dim vrTotComMovChq      As Currency
  Dim vrTotComMovCar      As Currency
  Dim vrTotComMovCarDeb   As Currency
  Dim vrTotComMovPen      As Currency
  Dim vrTotComMovFat      As Currency
  
  MovFinanc vrCaixaInicial, _
            vrCheque, _
            vrEspecie, _
            vrCartao, _
            vrCartaoDeb, _
            vrPenhor, _
            vrFatura, _
            vrGorjeta, _
            vrSaldoDespDiveDin, _
            vrSaldoDespDiveChqResg, _
            vrTroco, _
            vrDepDin, _
            vrDepChq, _
            vrDepCar, _
            vrDepCarDeb, _
            vrDepPen, _
            vrDepFat, _
            vrRetDin, _
            vrRetChq, _
            vrRetCar, _
            vrRetCarDeb, _
            vrRetPen, _
            vrRetFat, _
            vrTotComMov
  '
  vrTotComMovDin = vrEspecie - vrSaldoDespDiveDin - vrTroco
  vrTotComMovChq = vrCheque - vrSaldoDespDiveChqResg
  vrTotComMovCar = vrCartao
  vrTotComMovCarDeb = vrCartaoDeb
  vrTotComMovPen = vrPenhor
  vrTotComMovFat = vrFatura
  '
  vrTotComMovDin = vrTotComMovDin + vrDepDin - vrRetDin
  vrTotComMovChq = vrTotComMovChq + vrDepChq - vrRetChq
  vrTotComMovCar = vrTotComMovCar + vrDepCar - vrRetCar
  vrTotComMovCarDeb = vrTotComMovCarDeb + vrDepCarDeb - vrRetCarDeb
  vrTotComMovPen = vrTotComMovPen + vrDepPen - vrRetPen
  vrTotComMovFat = vrTotComMovFat + vrDepFat - vrRetFat
  '
  'dá Prosseguimento com o sistema
  '
'''  If gbPedirSenhaSupSangria Then
'''    '----------------------------
'''    '----------------------------
'''    'Pede Senha Superior (Diretor, Gerente ou Administrador
'''    If Not (gsNivel = "DIR" Or gsNivel = "GER" Or gsNivel = "ADM") Then
'''      'Só pede senha superior se quem estiver logado não for superior
'''      frmUserLoginSup.Show vbModal
'''
'''      If Len(Trim(gsNomeUsuLib)) = 0 Then
'''        MsgBox "É necessário a Confirmação com senha superior para acessar a Sangria.", vbExclamation, TITULOSISTEMA
'''        Exit Sub
'''      End If
'''      '
'''      'Capturou Nome do Usuário, continua processo de Sangria
'''    Else 'Senha Superior
'''      gsNomeUsuLib = gsNomeUsu
'''      '
'''    End If
'''
'''    '--------------------------------
'''    '--------------------------------
'''  End If
  '
  frmUserMovCaixaInc.vrTotCaixa = vrTotComMov
  frmUserMovCaixaInc.vrTotCaixaDin = vrTotComMovDin
  frmUserMovCaixaInc.vrTotCaixaChq = vrTotComMovChq
  frmUserMovCaixaInc.vrTotCaixaCar = vrTotComMovCar
  frmUserMovCaixaInc.vrTotCaixaCarDeb = vrTotComMovCarDeb
  frmUserMovCaixaInc.vrTotCaixaPen = vrTotComMovPen
  frmUserMovCaixaInc.vrTotCaixaFat = vrTotComMovFat
  
  frmUserMovCaixaInc.Show vbModal
  '
  Unload Me
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Public Sub CarregaInfFinanc()
  On Error GoTo trata
  Dim vrCaixaInicial  As Currency
  Dim vrCheque        As Currency
  Dim vrEspecie       As Currency
  Dim vrCartao        As Currency
  Dim vrCartaoDeb     As Currency
  Dim vrPenhor        As Currency
  Dim vrFatura        As Currency
  Dim vrGorjeta       As Currency
  Dim vrSaldoDespDiveDin      As Currency
  Dim vrSaldoDespDiveChqResg  As Currency
  Dim vrTroco                 As Currency
  '
  Dim vrDepDin          As Currency
  Dim vrDepChq          As Currency
  Dim vrDepCar          As Currency
  Dim vrDepCarDeb       As Currency
  Dim vrDepPen          As Currency
  Dim vrDepFat          As Currency
  Dim vrRetDin          As Currency
  Dim vrRetChq          As Currency
  Dim vrRetCar          As Currency
  Dim vrRetCarDeb       As Currency
  Dim vrRetPen          As Currency
  Dim vrRetFat          As Currency
  '
  Dim vrTotComMov As Currency
  
  MovFinanc vrCaixaInicial, _
            vrCheque, _
            vrEspecie, _
            vrCartao, _
            vrCartaoDeb, _
            vrPenhor, _
            vrFatura, _
            vrGorjeta, _
            vrSaldoDespDiveDin, _
            vrSaldoDespDiveChqResg, _
            vrTroco, _
            vrDepDin, _
            vrDepChq, _
            vrDepCar, _
            vrDepCarDeb, _
            vrDepPen, _
            vrDepFat, _
            vrRetDin, _
            vrRetChq, _
            vrRetCar, _
            vrRetCarDeb, _
            vrRetPen, _
            vrRetFat, _
            vrTotComMov
  txtVrInicial.Text = Format(vrCaixaInicial, "#,##0.00")
  txtSaldoCheque(0).Text = Format(vrCheque, "#,##0.00")
  txtSaldocartao(0).Text = Format(vrCartao, "#,##0.00")
  txtSaldocartaodeb.Text = Format(vrCartaoDeb, "#,##0.00")
  txtSaldoPenhor(0).Text = Format(vrPenhor, "#,##0.00")
  txtSaldoFatura(0).Text = Format(vrFatura, "#,##0.00")
  txtSaldoGorjeta.Text = Format(vrGorjeta, "#,##0.00")
  txtSaldoTroco.Text = Format(vrTroco, "#,##0.00")
  'txtSaldoDinheiro(0).Text = Format(vrEspecie + vrVendaCob + vrExtraForaDaUnidade - vrSaldoDespDive - vrTroco, "#,##0.00")
  txtSaldoDinheiro(0).Text = Format(vrEspecie - vrTroco, "#,##0.00")
  txtSaldoDespDiverDin.Text = Format(vrSaldoDespDiveDin, "#,##0.00")
  txtSaldoDespDiverChqResg.Text = Format(vrSaldoDespDiveChqResg, "#,##0.00")
  txtSaldoTotal.Text = Format(vrEspecie + vrCheque + vrCartao + vrCartaoDeb + vrPenhor + vrFatura - vrSaldoDespDiveDin - vrSaldoDespDiveChqResg - vrTroco, "#,##0.00")
  'Joga nos campos os valores já somados
  txtDepDin.Text = Format(vrDepDin, "#,##0.00")
  txtDepChq.Text = Format(vrDepChq, "#,##0.00")
  txtDepCar.Text = Format(vrDepCar, "#,##0.00")
  txtDepCarDeb.Text = Format(vrDepCarDeb, "#,##0.00")
  txtDepPen.Text = Format(vrDepPen, "#,##0.00")
  txtDepFat.Text = Format(vrDepFat, "#,##0.00")
  txtRetDin.Text = Format(vrRetDin, "#,##0.00")
  txtRetChq.Text = Format(vrRetChq, "#,##0.00")
  txtRetCar.Text = Format(vrRetCar, "#,##0.00")
  txtRetCarDeb.Text = Format(vrRetCarDeb, "#,##0.00")
  txtRetPen.Text = Format(vrRetPen, "#,##0.00")
  txtRetFat.Text = Format(vrRetFat, "#,##0.00")
  '
  txtSaldoTotalMov(0).Text = Format(vrTotComMov, "#,##0.00")
  '
  'Saldo Totais
  txtSaldoCheque(1).Text = Format(vrCheque + vrDepChq - vrRetChq - vrSaldoDespDiveChqResg, "#,##0.00")
  txtSaldocartao(1).Text = Format(vrCartao + vrDepCar - vrRetCar, "#,##0.00")
  txtSaldocartaodebs.Text = Format(vrCartaoDeb + vrDepCarDeb - vrRetCarDeb, "#,##0.00")
  txtSaldocartaodeb.Text = Format(vrCartaoDeb + vrDepCarDeb - vrRetCarDeb, "#,##0.00")
  txtSaldoPenhor(1).Text = Format(vrPenhor + vrDepPen - vrRetPen, "#,##0.00")
  txtSaldoFatura(1).Text = Format(vrFatura + vrDepFat - vrRetFat, "#,##0.00")
  txtSaldoDinheiro(1).Text = Format(vrEspecie - vrSaldoDespDiveDin - vrTroco + vrDepDin - vrRetDin, "#,##0.00")


  txtSaldoTotalMov(1).Text = Format(vrTotComMov, "#,##0.00")
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
End Sub
Private Sub Form_Load()
  On Error GoTo trata
  Dim sSql As String
  AmpS
  Me.Width = 11385
  Me.Height = 6570
  CenterForm Me
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar
  LerFigurasAvulsas cmdSangria, "Sangria.ico", "SangriaDown.ico", "Movimentações financeiras do turno (Sangria)"
  '
  CarregaInfFinanc
  '
  tabDetalhes.TabVisible(5) = False
  tabDetalhes.TabVisible(6) = False
  tabDetalhes.TabVisible(7) = False
  tabDetalhes_Click 0
  '
  txtTurno.Text = RetornaDescTurnoCorrente
  txtFuncionario.Text = gsNomeUsuCompleto
  mskData(0).Text = DataHoraAtualFormatada
  '
  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Private Sub grdExtra_UnboundReadDataEx( _
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
               Offset + intI, EXTR_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, EXTR_COLUNASMATRIZ, EXTR_LINHASMATRIZ, EXTR_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, EXTR_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
 TratarErro Err.Number, Err.Description, "[frmUserInfFinancLis.grdExtra_UnboundReadDataEx]"
End Sub

Private Sub grdDespesa_UnboundReadDataEx( _
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
               Offset + intI, DESP_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, DESP_COLUNASMATRIZ, DESP_LINHASMATRIZ, DESP_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, DESP_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
 TratarErro Err.Number, Err.Description, "[frmUserInfFinancLis.grdDespesa_UnboundReadDataEx]"
End Sub

Private Sub grdGarcom_UnboundReadDataEx( _
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
               Offset + intI, GARCOM_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, GARCOM_COLUNASMATRIZ, GARCOM_LINHASMATRIZ, GARCOM_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, GARCOM_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
 TratarErro Err.Number, Err.Description, "[frmUserInfFinancLis.grdGarcom_UnboundReadDataEx]"
End Sub

Private Sub grdCartaoDeb_UnboundReadDataEx( _
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
               Offset + intI, RESCARTDEB_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, RESCARTDEB_COLUNASMATRIZ, RESCARTDEB_LINHASMATRIZ, RESCARTDEB_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, RESCARTDEB_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserInfFinancLis.grdCartaoDeb_UnboundReadDataEx]"
End Sub

Private Sub grdGeral_UnboundReadDataEx( _
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
               Offset + intI, RESCART_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, RESCART_COLUNASMATRIZ, RESCART_LINHASMATRIZ, RESCART_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, RESCART_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserInfFinancLis.grdGeral_UnboundReadDataEx]"
End Sub

Private Sub grdSangria_UnboundReadDataEx( _
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
               Offset + intI, GER_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, GER_COLUNASMATRIZ, GER_LINHASMATRIZ, GER_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, GER_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserInfFinancLis.grdSangria_UnboundReadDataEx]"
End Sub


Public Sub MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMed.clsGeral
  '
  strSql = "SELECT RESPONSAVEL, DATA, VRRETDIN, VRRETPEN, VRRETFAT, VRRETCAR, VRRETCARDEB, VRRETCHQ, VRDEPDIN, VRDEPPEN, VRDEPFAT, VRDEPCAR, VRDEPCARDEB, VRDEPCHQ FROM  SANGRIA " & _
    "WHERE TURNOID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo) & ";"
  
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    GER_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim GER_Matriz(0 To GER_COLUNASMATRIZ - 1, 0 To GER_LINHASMATRIZ - 1)
  Else
    ReDim GER_Matriz(0 To GER_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To GER_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To GER_COLUNASMATRIZ - 1  'varre as colunas
          GER_Matriz(intJ, intI) = objRs(intJ) & ""
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

Public Sub MontaMatrizExtra()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMed.clsGeral
  '
  strSql = "SELECT EXTRAUNIDADE.NRO, TAB_EXTRAUNIDADE.DESCRICAO, TAB_EXTRAUNIDADE.VALOR AS VALOR FROM  EXTRAUNIDADE " & _
    "INNER JOIN TAB_EXTRAUNIDADE ON EXTRAUNIDADE.PKID = TAB_EXTRAUNIDADE.EXTRAUNIDADEID " & _
    "WHERE TURNOID = " & RetornaCodTurnoCorrente & " order by NRO DESC;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    EXTR_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim EXTR_Matriz(0 To EXTR_COLUNASMATRIZ - 1, 0 To EXTR_LINHASMATRIZ - 1)
  Else
    ReDim EXTR_Matriz(0 To EXTR_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To EXTR_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To EXTR_COLUNASMATRIZ - 1  'varre as colunas
          EXTR_Matriz(intJ, intI) = objRs(intJ) & ""
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

Public Sub MontaMatrizDespesa()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMed.clsGeral
  '
  strSql = "SELECT DESPESA.SEQUENCIAL, DESPESA.DESCRICAO, DESPESA.VALE, DESPESA.VR_PAGO AS VALOR FROM  DESPESA " & _
    "WHERE TIPO = 'T' AND TURNOID = " & RetornaCodTurnoCorrente & " order by sequencial DESC;"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    DESP_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim DESP_Matriz(0 To DESP_COLUNASMATRIZ - 1, 0 To DESP_LINHASMATRIZ - 1)
  Else
    ReDim DESP_Matriz(0 To DESP_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To DESP_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To DESP_COLUNASMATRIZ - 1  'varre as colunas
          DESP_Matriz(intJ, intI) = objRs(intJ) & ""
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

Public Sub MontaMatrizGarcom()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMed.clsGeral
  '
  strSql = "SELECT case when GARCOM.NOME is null then 'Garçom não informado' else GARCOM.NOME end AS NOME, SUM(CONTACORRENTE.VRTROCO) AS TROCO, SUM(CONTACORRENTE.VRGORJETA) AS TROCO FROM  CONTACORRENTE LEFT JOIN GARCOM ON (GARCOM.PKID = CONTACORRENTE.GARCOMID) " & _
    "WHERE TURNOCCID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo) & " AND (VRGORJETA <> 0 OR VRTROCO <> 0) group by nome;"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    GARCOM_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim GARCOM_Matriz(0 To GARCOM_COLUNASMATRIZ - 1, 0 To GARCOM_LINHASMATRIZ - 1)
  Else
    ReDim GARCOM_Matriz(0 To GARCOM_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To GARCOM_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To GARCOM_COLUNASMATRIZ - 1  'varre as colunas
          GARCOM_Matriz(intJ, intI) = objRs(intJ) & ""
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

Private Sub tabDetalhes_Click(PreviousTab As Integer)
  On Error GoTo trata
  Select Case tabDetalhes.Tab
  Case 0, 1
    grdGarcom.Enabled = False
    grdGeral.Enabled = False
    grdCartaoDeb.Enabled = False
    grdSangria.Enabled = False
    grdDespesa.Enabled = False
    grdExtra.Enabled = False
  Case 2
    grdGarcom.Enabled = False
    grdGeral.Enabled = False
    grdCartaoDeb.Enabled = False
    grdSangria.Enabled = True
    grdDespesa.Enabled = False
    grdExtra.Enabled = False
    'Montar RecordSet
    GER_COLUNASMATRIZ = grdSangria.Columns.Count
    GER_LINHASMATRIZ = 0
    MontaMatriz
    grdSangria.Bookmark = Null
    grdSangria.ReBind
    grdSangria.ApproxCount = GER_LINHASMATRIZ
    SetarFoco grdSangria
  Case 3
    grdGarcom.Enabled = False
    grdGeral.Enabled = True
    grdCartaoDeb.Enabled = False
    grdSangria.Enabled = False
    grdDespesa.Enabled = False
    grdExtra.Enabled = False
    'Montar RecordSet
    RESCART_COLUNASMATRIZ = grdGeral.Columns.Count
    RESCART_LINHASMATRIZ = 0
    MontaMatrizResumoCartao
    grdGeral.Bookmark = Null
    grdGeral.ReBind
    grdGeral.ApproxCount = RESCART_LINHASMATRIZ
    SetarFoco grdGeral
  Case 4
    grdGarcom.Enabled = False
    grdGeral.Enabled = False
    grdCartaoDeb.Enabled = True
    grdSangria.Enabled = False
    grdDespesa.Enabled = False
    grdExtra.Enabled = False
    'Montar RecordSet
    RESCARTDEB_COLUNASMATRIZ = grdCartaoDeb.Columns.Count
    RESCARTDEB_LINHASMATRIZ = 0
    MontaMatrizResumoCartaoDeb
    grdCartaoDeb.Bookmark = Null
    grdCartaoDeb.ReBind
    grdCartaoDeb.ApproxCount = RESCARTDEB_LINHASMATRIZ
    SetarFoco grdCartaoDeb
  Case 5
    grdGarcom.Enabled = True
    grdGeral.Enabled = False
    grdCartaoDeb.Enabled = False
    grdSangria.Enabled = False
    grdDespesa.Enabled = False
    grdExtra.Enabled = False
    'Montar RecordSet
    GARCOM_COLUNASMATRIZ = grdGarcom.Columns.Count
    GARCOM_LINHASMATRIZ = 0
    MontaMatrizGarcom
    grdGarcom.Bookmark = Null
    grdGarcom.ReBind
    grdGarcom.ApproxCount = GARCOM_LINHASMATRIZ
    SetarFoco grdGarcom
  Case 6
    grdGarcom.Enabled = True
    grdGeral.Enabled = False
    grdCartaoDeb.Enabled = False
    grdSangria.Enabled = False
    grdDespesa.Enabled = False
    grdExtra.Enabled = False
    'Montar RecordSet
    DESP_COLUNASMATRIZ = grdDespesa.Columns.Count
    DESP_LINHASMATRIZ = 0
    MontaMatrizDespesa
    grdDespesa.Bookmark = Null
    grdDespesa.ReBind
    grdDespesa.ApproxCount = DESP_LINHASMATRIZ
    SetarFoco grdDespesa
  Case 7
    grdGarcom.Enabled = False
    grdGeral.Enabled = False
    grdCartaoDeb.Enabled = False
    grdSangria.Enabled = False
    grdDespesa.Enabled = False
    grdExtra.Enabled = True
    'Montar RecordSet
    EXTR_COLUNASMATRIZ = grdExtra.Columns.Count
    EXTR_LINHASMATRIZ = 0
    MontaMatrizExtra
    grdExtra.Bookmark = Null
    grdExtra.ReBind
    grdExtra.ApproxCount = EXTR_LINHASMATRIZ
    SetarFoco grdExtra
  End Select
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "frmUserInfFinancLis.tabDetalhes_Click"
  AmpN
End Sub

Public Sub MontaMatrizResumoCartaoDeb()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMed.clsGeral
  '
  strSql = "SELECT CARTAODEBITO.NOME AS NOME, SUM(CONTACORRENTE.VALOR) AS VALOR FROM CONTACORRENTE INNER JOIN CARTAODEBITO ON CARTAODEBITO.PKID = CONTACORRENTE.CARTAODEBITOID " & _
    "WHERE CONTACORRENTE.TURNOCCID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo) & _
    " AND CONTACORRENTE.STATUSCC = " & Formata_Dados("CD", tpDados_Texto) & _
    " group by nome"

  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    RESCARTDEB_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim RESCARTDEB_Matriz(0 To RESCARTDEB_COLUNASMATRIZ - 1, 0 To RESCARTDEB_LINHASMATRIZ - 1)
  Else
    ReDim RESCARTDEB_Matriz(0 To RESCARTDEB_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To RESCARTDEB_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To RESCARTDEB_COLUNASMATRIZ - 1  'varre as colunas
          RESCARTDEB_Matriz(intJ, intI) = objRs(intJ) & ""
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



Public Sub MontaMatrizResumoCartao()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMed.clsGeral
  '
  strSql = "SELECT CARTAO.NOME AS NOME, SUM(CONTACORRENTE.VALOR) AS VALOR FROM  CONTACORRENTE INNER JOIN CARTAO ON (CARTAO.PKID = CONTACORRENTE.CARTAOID) " & _
    "WHERE CONTACORRENTE.TURNOCCID = " & Formata_Dados(RetornaCodTurnoCorrente, tpDados_Longo) & _
    " AND CONTACORRENTE.STATUSCC = " & Formata_Dados("CC", tpDados_Texto) & _
    " group by nome"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    RESCART_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim RESCART_Matriz(0 To RESCART_COLUNASMATRIZ - 1, 0 To RESCART_LINHASMATRIZ - 1)
  Else
    ReDim RESCART_Matriz(0 To RESCART_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To RESCART_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To RESCART_COLUNASMATRIZ - 1  'varre as colunas
          RESCART_Matriz(intJ, intI) = objRs(intJ) & ""
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


