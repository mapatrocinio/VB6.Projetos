VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{00028CDA-0000-0000-0000-000000000046}#6.0#0"; "tdbg6.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Begin VB.Form frmUserGRPgtoInc 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Associação de GR"
   ClientHeight    =   8445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11520
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   11520
   StartUpPosition =   3  'Windows Default
   Begin TabDlg.SSTab tabDetalhes 
      Height          =   8265
      Left            =   150
      TabIndex        =   53
      TabStop         =   0   'False
      Top             =   120
      Width           =   9405
      _ExtentX        =   16589
      _ExtentY        =   14579
      _Version        =   393216
      Style           =   1
      Tabs            =   1
      TabsPerRow      =   1
      TabHeight       =   520
      TabCaption(0)   =   "Associar GR"
      TabPicture(0)   =   "userGRPgtoInc.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label6(3)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label7(1)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label5(1)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label7(0)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label5(2)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label5(3)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label6(0)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label5(8)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label5(9)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label5(10)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "Label5(11)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "Label5(13)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label5(0)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label5(4)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "Label5(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label5(23)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label5(6)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label5(7)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label5(12)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label7(2)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label5(14)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label5(15)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "mskDonoUltraConsConvenio"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "mskTecRXConsConvenio"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "mskDonoRXConsConvenio"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "mskPrestConsConvenio"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "mskDonoUltraConvenio"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "mskTecRXConvenio"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "mskDonoRXConvenio"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "mskPrestConvenio"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "mskAuxiliar"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "mskDtAReceb"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "mskCasa"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "mskTotalARec"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "mskTotalRec"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "mskTotalDonoUltra"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "mskTotalTecRX"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "mskTotalDonoRX"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "mskTotalPrest"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "mskDonoUltraCartaoARec"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "mskTecRXCartaoARec"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "mskDonoRXCartaoARec"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "mskPrestCartaoARec"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "mskDonoUltraCartao"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "mskDonoUltraEspecie"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "mskDonoUltraConsCartao"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "mskTecRXCartao"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "mskTecRXEspecie"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "mskTecRXConsCartao"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "mskDonoRXCartao"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "mskDonoRXEspecie"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "mskDonoRXConsCartao"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "mskPrestCartao"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "mskPrestEspecie"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "mskPrestConsCartao"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "mskDonoUltraConsEspecie"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "mskPrestConsEspecie"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "mskDonoRXConsEspecie"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "mskTecRXConsEspecie"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "mskTotAPagar"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "mskTotPagas"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "mskHoraTermino"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "mskDtTermino"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "mskHoraInicio"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "mskDtInicio"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "mskTotal"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).Control(67)=   "grdGRAssoc"
      Tab(0).Control(67).Enabled=   0   'False
      Tab(0).Control(68)=   "grdGR"
      Tab(0).Control(68).Enabled=   0   'False
      Tab(0).Control(69)=   "cmdCadastraItem(0)"
      Tab(0).Control(69).Enabled=   0   'False
      Tab(0).Control(70)=   "cmdCadastraItem(1)"
      Tab(0).Control(70).Enabled=   0   'False
      Tab(0).Control(71)=   "cmdCadastraItem(2)"
      Tab(0).Control(71).Enabled=   0   'False
      Tab(0).Control(72)=   "cmdCadastraItem(3)"
      Tab(0).Control(72).Enabled=   0   'False
      Tab(0).Control(73)=   "txtPrestador"
      Tab(0).Control(73).Enabled=   0   'False
      Tab(0).Control(74)=   "txtAcCartao"
      Tab(0).Control(74).Enabled=   0   'False
      Tab(0).ControlCount=   75
      Begin VB.TextBox txtAcCartao 
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Left            =   7650
         TabIndex        =   64
         Text            =   "txtAcCartao"
         Top             =   600
         Width           =   1185
      End
      Begin VB.TextBox txtPrestador 
         BackColor       =   &H8000000B&
         Enabled         =   0   'False
         Height          =   285
         Left            =   870
         TabIndex        =   4
         Text            =   "txtPrestador"
         Top             =   600
         Width           =   5565
      End
      Begin VB.CommandButton cmdCadastraItem 
         Caption         =   "<<"
         Height          =   375
         Index           =   3
         Left            =   8970
         TabIndex        =   10
         Top             =   1980
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastraItem 
         Caption         =   "<"
         Height          =   375
         Index           =   2
         Left            =   8970
         TabIndex        =   9
         Top             =   1620
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastraItem 
         Caption         =   ">>"
         Height          =   375
         Index           =   1
         Left            =   8970
         TabIndex        =   8
         Top             =   1260
         Width           =   375
      End
      Begin VB.CommandButton cmdCadastraItem 
         Caption         =   ">"
         Height          =   375
         Index           =   0
         Left            =   8970
         TabIndex        =   7
         Top             =   900
         Width           =   375
      End
      Begin TrueDBGrid60.TDBGrid grdGR 
         Height          =   2460
         Left            =   90
         OleObjectBlob   =   "userGRPgtoInc.frx":001C
         TabIndex        =   5
         Top             =   900
         Width           =   8760
      End
      Begin TrueDBGrid60.TDBGrid grdGRAssoc 
         Height          =   2700
         Left            =   90
         OleObjectBlob   =   "userGRPgtoInc.frx":4A2F
         TabIndex        =   6
         Top             =   3360
         Width           =   8790
      End
      Begin MSMask.MaskEdBox mskTotal 
         Height          =   255
         Left            =   7080
         TabIndex        =   41
         Top             =   7740
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDtInicio 
         Height          =   255
         Left            =   870
         TabIndex        =   0
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   -2147483637
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHoraInicio 
         Height          =   255
         Left            =   2250
         TabIndex        =   1
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   -2147483637
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   8
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDtTermino 
         Height          =   255
         Left            =   3270
         TabIndex        =   2
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   -2147483637
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskHoraTermino 
         Height          =   255
         Left            =   4650
         TabIndex        =   3
         Top             =   360
         Width           =   855
         _ExtentX        =   1508
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   -2147483637
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   8
         Mask            =   "##:##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTotPagas 
         Height          =   255
         Left            =   6660
         TabIndex        =   59
         Top             =   330
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0;($#,##0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTotAPagar 
         Height          =   255
         Left            =   8190
         TabIndex        =   61
         Top             =   330
         Width           =   615
         _ExtentX        =   1085
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0;($#,##0)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTecRXConsEspecie 
         Height          =   255
         Left            =   3240
         TabIndex        =   13
         Top             =   6240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDonoRXConsEspecie 
         Height          =   255
         Left            =   2280
         TabIndex        =   12
         Top             =   6240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPrestConsEspecie 
         Height          =   255
         Left            =   1320
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   6240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDonoUltraConsEspecie 
         Height          =   255
         Left            =   4200
         TabIndex        =   14
         Top             =   6240
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPrestConsCartao 
         Height          =   255
         Left            =   1320
         TabIndex        =   15
         Top             =   6480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPrestEspecie 
         Height          =   255
         Left            =   1320
         TabIndex        =   23
         Top             =   6960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPrestCartao 
         Height          =   255
         Left            =   1320
         TabIndex        =   27
         Top             =   7200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDonoRXConsCartao 
         Height          =   255
         Left            =   2280
         TabIndex        =   16
         Top             =   6480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDonoRXEspecie 
         Height          =   255
         Left            =   2280
         TabIndex        =   24
         Top             =   6960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDonoRXCartao 
         Height          =   255
         Left            =   2280
         TabIndex        =   28
         Top             =   7200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTecRXConsCartao 
         Height          =   255
         Left            =   3240
         TabIndex        =   17
         Top             =   6480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTecRXEspecie 
         Height          =   255
         Left            =   3240
         TabIndex        =   25
         Top             =   6960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTecRXCartao 
         Height          =   255
         Left            =   3240
         TabIndex        =   29
         Top             =   7200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDonoUltraConsCartao 
         Height          =   255
         Left            =   4200
         TabIndex        =   18
         Top             =   6480
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDonoUltraEspecie 
         Height          =   255
         Left            =   4200
         TabIndex        =   26
         Top             =   6960
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDonoUltraCartao 
         Height          =   255
         Left            =   4200
         TabIndex        =   30
         Top             =   7200
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPrestCartaoARec 
         Height          =   255
         Left            =   1320
         TabIndex        =   42
         Top             =   7980
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDonoRXCartaoARec 
         Height          =   255
         Left            =   2280
         TabIndex        =   43
         Top             =   7980
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTecRXCartaoARec 
         Height          =   255
         Left            =   3240
         TabIndex        =   44
         Top             =   7980
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDonoUltraCartaoARec 
         Height          =   255
         Left            =   4200
         TabIndex        =   45
         Top             =   7980
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTotalPrest 
         Height          =   255
         Left            =   1320
         TabIndex        =   35
         Top             =   7740
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTotalDonoRX 
         Height          =   255
         Left            =   2280
         TabIndex        =   36
         Top             =   7740
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTotalTecRX 
         Height          =   255
         Left            =   3270
         TabIndex        =   37
         Top             =   7740
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTotalDonoUltra 
         Height          =   255
         Left            =   4200
         TabIndex        =   38
         Top             =   7740
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTotalRec 
         Height          =   255
         Left            =   5160
         TabIndex        =   39
         Top             =   7740
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTotalARec 
         Height          =   255
         Left            =   5160
         TabIndex        =   46
         Top             =   7980
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16776960
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskCasa 
         Height          =   255
         Left            =   6120
         TabIndex        =   40
         Top             =   7740
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDtAReceb 
         Height          =   255
         Left            =   6720
         TabIndex        =   47
         Top             =   7980
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   -2147483637
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskAuxiliar 
         Height          =   255
         Left            =   5610
         TabIndex        =   79
         Top             =   6960
         Visible         =   0   'False
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPrestConvenio 
         Height          =   255
         Left            =   1320
         TabIndex        =   31
         Top             =   7440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDonoRXConvenio 
         Height          =   255
         Left            =   2280
         TabIndex        =   32
         Top             =   7440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTecRXConvenio 
         Height          =   255
         Left            =   3240
         TabIndex        =   33
         Top             =   7440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDonoUltraConvenio 
         Height          =   255
         Left            =   4200
         TabIndex        =   34
         Top             =   7440
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskPrestConsConvenio 
         Height          =   255
         Left            =   1320
         TabIndex        =   19
         Top             =   6720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDonoRXConsConvenio 
         Height          =   255
         Left            =   2280
         TabIndex        =   20
         Top             =   6720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskTecRXConsConvenio 
         Height          =   255
         Left            =   3240
         TabIndex        =   21
         Top             =   6720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox mskDonoUltraConsConvenio 
         Height          =   255
         Left            =   4200
         TabIndex        =   22
         Top             =   6720
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   450
         _Version        =   393216
         BackColor       =   16777152
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         Format          =   "#,##0.00;($#,##0.00)"
         PromptChar      =   "_"
      End
      Begin VB.Label Label5 
         Caption         =   "Cons. Convênio"
         Enabled         =   0   'False
         Height          =   195
         Index           =   15
         Left            =   60
         TabIndex        =   81
         Top             =   6720
         Width           =   1125
      End
      Begin VB.Label Label5 
         Caption         =   "Convênio"
         Enabled         =   0   'False
         Height          =   195
         Index           =   14
         Left            =   60
         TabIndex        =   80
         Top             =   7440
         Width           =   1215
      End
      Begin VB.Label Label7 
         Caption         =   "Data"
         Height          =   255
         Index           =   2
         Left            =   6150
         TabIndex        =   78
         Top             =   7980
         Width           =   645
      End
      Begin VB.Label Label5 
         Caption         =   "Total Geral"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   7080
         TabIndex        =   77
         Top             =   7530
         Width           =   1035
      End
      Begin VB.Label Label5 
         Caption         =   "Casa"
         Enabled         =   0   'False
         Height          =   195
         Index           =   7
         Left            =   6150
         TabIndex        =   76
         Top             =   7530
         Visible         =   0   'False
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "Total"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   5190
         TabIndex        =   75
         Top             =   7530
         Width           =   525
      End
      Begin VB.Label Label5 
         Caption         =   "Cartão (à pagar)"
         Enabled         =   0   'False
         Height          =   195
         Index           =   23
         Left            =   60
         TabIndex        =   74
         Top             =   7980
         Width           =   1185
      End
      Begin VB.Label Label5 
         Caption         =   "Cartão"
         Enabled         =   0   'False
         Height          =   195
         Index           =   5
         Left            =   60
         TabIndex        =   73
         Top             =   7200
         Width           =   1215
      End
      Begin VB.Label Label5 
         Caption         =   "Espécie"
         Enabled         =   0   'False
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   72
         Top             =   6960
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "Cons. Cartão"
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   71
         Top             =   6480
         Width           =   1125
      End
      Begin VB.Label Label5 
         Caption         =   "Aparelho Ultra"
         Enabled         =   0   'False
         Height          =   195
         Index           =   13
         Left            =   4200
         TabIndex        =   70
         Top             =   6030
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "Técnico"
         Enabled         =   0   'False
         Height          =   195
         Index           =   11
         Left            =   3240
         TabIndex        =   69
         Top             =   6030
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "Aparelho"
         Enabled         =   0   'False
         Height          =   195
         Index           =   10
         Left            =   2280
         TabIndex        =   68
         Top             =   6030
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "Prestador"
         Enabled         =   0   'False
         Height          =   195
         Index           =   9
         Left            =   1320
         TabIndex        =   67
         Top             =   6030
         Width           =   1005
      End
      Begin VB.Label Label5 
         Caption         =   "Cons. Espécie"
         Enabled         =   0   'False
         Height          =   195
         Index           =   8
         Left            =   60
         TabIndex        =   66
         Top             =   6240
         Width           =   1155
      End
      Begin VB.Label Label6 
         Caption         =   "Cartão?"
         Height          =   255
         Index           =   0
         Left            =   6660
         TabIndex        =   63
         Top             =   630
         Width           =   855
      End
      Begin VB.Label Label5 
         Caption         =   "A pagar"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   7350
         TabIndex        =   62
         Top             =   345
         Width           =   765
      End
      Begin VB.Label Label5 
         Caption         =   "Total Pagas"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   5550
         TabIndex        =   60
         Top             =   345
         Width           =   1065
      End
      Begin VB.Label Label7 
         Caption         =   "a"
         Height          =   255
         Index           =   0
         Left            =   3120
         TabIndex        =   58
         Top             =   360
         Width           =   105
      End
      Begin VB.Label Label5 
         Caption         =   "Total"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   57
         Top             =   7740
         Width           =   525
      End
      Begin VB.Label Label7 
         Caption         =   "Período"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   56
         Top             =   360
         Width           =   645
      End
      Begin VB.Label Label6 
         Caption         =   "Prestador"
         Height          =   255
         Index           =   3
         Left            =   90
         TabIndex        =   55
         Top             =   600
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "* Aperte a tecla <CTRL> OU <SHIFT> + Botão direito do mouse para selecionar mais de um item do grid."
         ForeColor       =   &H000000FF&
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   54
         Top             =   5820
         Width           =   8715
         WordWrap        =   -1  'True
      End
   End
   Begin VB.PictureBox picBotoes 
      Align           =   4  'Align Right
      BorderStyle     =   0  'None
      Height          =   8445
      Left            =   9660
      ScaleHeight     =   8445
      ScaleWidth      =   1860
      TabIndex        =   51
      TabStop         =   0   'False
      Top             =   0
      Width           =   1860
      Begin VB.PictureBox SSPanel1 
         BackColor       =   &H00C0C0C0&
         Height          =   2985
         Left            =   120
         ScaleHeight     =   2925
         ScaleWidth      =   1545
         TabIndex        =   52
         TabStop         =   0   'False
         Top             =   5400
         Width           =   1605
         Begin VB.CommandButton cmdRelatorioARec 
            Caption         =   "À PAGAR"
            Height          =   880
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   49
            Top             =   990
            Width           =   1335
         End
         Begin VB.CommandButton cmdRelatorio 
            Caption         =   "PAGO"
            Height          =   880
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   120
            Width           =   1335
         End
         Begin VB.CommandButton cmdFechar 
            Cancel          =   -1  'True
            Caption         =   "ESC"
            Height          =   880
            Left            =   90
            Style           =   1  'Graphical
            TabIndex        =   50
            Top             =   1860
            Width           =   1335
         End
      End
      Begin Crystal.CrystalReport Report1 
         Left            =   750
         Top             =   210
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin Crystal.CrystalReport Report2 
         Left            =   750
         Top             =   630
         _ExtentX        =   741
         _ExtentY        =   741
         _Version        =   348160
         WindowControlBox=   -1  'True
         WindowMaxButton =   -1  'True
         WindowMinButton =   -1  'True
         PrintFileLinesPerPage=   60
      End
      Begin VB.Label Label1 
         Caption         =   "* As unidades serão incluidas/excluidas automaticamente após ser pressionado os botões >, >>, < ou <<."
         ForeColor       =   &H000000FF&
         Height          =   2355
         Index           =   0
         Left            =   90
         TabIndex        =   65
         Top             =   1110
         Width           =   1695
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmUserGRPgtoInc"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public blnRetorno           As Boolean
Dim blnFechar               As Boolean
Public strGR                As String
Public strStatusGR          As String
Public icTipoGR             As tpIcTipoGR
Public Status               As tpStatus
'
Public strDataIni           As String
Public strHoraIni           As String
Public strDataFim           As String
Public strHoraFim           As String
Public strPrestador         As String

Public lngGRID              As Long
Public lngGRPAGAMENTOID     As Long
'Variáveis para Grid
'
Dim GR_COLUNASMATRIZ        As Long
Dim GR_LINHASMATRIZ         As Long
Private GR_Matriz()         As String
'
Dim GRASSOC_COLUNASMATRIZ   As Long
Dim GRASSOC_LINHASMATRIZ    As Long
Private GRASSOC_Matriz()    As String

Private Sub cmdCadastraItem_Click(Index As Integer)
  TratarAssociacao Index + 1
  SetarFoco grdGR
End Sub



Public Sub GRASSOC_MontaMatriz()
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMed.clsGeral
  '
  strSql = "SELECT GRPGTO.PKID, GR.PKID, MAX(PRONTUARIO.NOME) AS NOME, MAX(GR.SEQUENCIAL) AS SEQUENCIAL, MAX(GR.SENHA) AS SENHA, MAX(GR.DATA) AS DATA, MAX(GRPAGAMENTO.DATAINICIO) AS DATAPGTO, SUM(GRPROCEDIMENTO.VALOR) AS VALOR " & _
      " FROM GRPAGAMENTO INNER JOIN GRPGTO ON GRPAGAMENTO.PKID = GRPGTO.GRPAGAMENTOID " & _
      " INNER JOIN GR ON GR.PKID = GRPGTO.GRID " & _
      " INNER JOIN GRPROCEDIMENTO ON GR.PKID = GRPROCEDIMENTO.GRID " & _
      " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = ATENDE.PRONTUARIOID " & _
      " WHERE GRPAGAMENTO.PKID = " & Formata_Dados(lngGRPAGAMENTOID, tpDados_Longo)

  strSql = strSql & " GROUP BY GR.PKID, GRPGTO.PKID " & _
      " ORDER BY PRONTUARIO.NOME, GR.SEQUENCIAL, GR.DATA"
  '
  Set objRs = clsGer.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    GRASSOC_LINHASMATRIZ = objRs.RecordCount
  End If
  If Not objRs.EOF Then
    ReDim GRASSOC_Matriz(0 To GRASSOC_COLUNASMATRIZ - 1, 0 To GRASSOC_LINHASMATRIZ - 1)
  Else
    ReDim GRASSOC_Matriz(0 To GRASSOC_COLUNASMATRIZ - 1, 0 To 0)
  End If
  '
  If Not objRs.EOF Then   'se já houver algum item
    For intI = 0 To GRASSOC_LINHASMATRIZ - 1  'varre as linhas
      If Not objRs.EOF Then 'enquanto ainda não se atingiu fim do recordset
        For intJ = 0 To GRASSOC_COLUNASMATRIZ - 1  'varre as colunas
          GRASSOC_Matriz(intJ, intI) = objRs(intJ) & ""
        Next
        objRs.MoveNext
      End If
    Next  'próxima linha matriz
  End If
  Set clsGer = Nothing
  'Cálculo dos totais
  Calculo_totais_GR
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Public Sub Calculo_totais_GR()
  On Error GoTo trata
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim objGer    As busSisMed.clsGeral
  'Limpar campos
  'LimparCampoMask mskCasa
  'LimparCampoMask mskPrestador
  'LimparCampoMask mskTotal
  'LimparCampoMask mskTotPagas
  'LimparCampoMask mskTotAPagar
  '
  'Tratar campos
  Set objGer = New busSisMed.clsGeral
  '
  'strSql = "SELECT  sum(vw_cons_t_Financ.PgtoTotal) as PgtoTotal, " & _
      " sum(vw_cons_t_Financ.FINALCASACONSESPECIE) as FINALCASACONSESPECIE, sum(vw_cons_t_Financ.FINALPRESTCONSESPECIE) as FINALPRESTCONSESPECIE, sum(vw_cons_t_Financ.FINALDONORXCONSESPECIE) as FINALDONORXCONSESPECIE, sum(vw_cons_t_Financ.FINALTECRXCONSESPECIE) as FINALTECRXCONSESPECIE, " & _
      " sum(vw_cons_t_Financ.FINALDONOULTRACONSESPECIE) as FINALDONOULTRACONSESPECIE, sum(vw_cons_t_Financ.FINALCASACONSCARTAO) as FINALCASACONSCARTAO, sum(vw_cons_t_Financ.FINALPRESTCONSCARTAO) as FINALPRESTCONSCARTAO, " & _
      " sum(vw_cons_t_Financ.FINALDONORXCONSCARTAO) as FINALDONORXCONSCARTAO, sum(vw_cons_t_Financ.FINALTECRXCONSCARTAO) as FINALTECRXCONSCARTAO, sum(vw_cons_t_Financ.FINALDONOULTRACONSCARTAO) as FINALDONOULTRACONSCARTAO, " & _
      " sum(vw_cons_t_Financ.FINALCASAESPECIE) as FINALCASAESPECIE, sum(vw_cons_t_Financ.FINALCASACARTAO) as FINALCASACARTAO, sum(vw_cons_t_Financ.FINALPRESTESPECIE) as FINALPRESTESPECIE, sum(vw_cons_t_Financ.FINALPRESTCARTAONAOACEITA) as FINALPRESTCARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALPRESTCARTAOACEITAPGFUTURO) as FINALPRESTCARTAOACEITAPGFUTURO, sum(vw_cons_t_Financ.FINALDONORXESPECIE) as FINALDONORXESPECIE, sum(vw_cons_t_Financ.FINALDONORXCARTAONAOACEITA) as FINALDONORXCARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALDONORXCARTAOACEITAPGFUTURO) as FINALDONORXCARTAOACEITAPGFUTURO, sum(vw_cons_t_Financ.FINALDONOULTRAESPECIE) as FINALDONOULTRAESPECIE, sum(vw_cons_t_Financ.FINALDONOULTRACARTAONAOACEITA) as FINALDONOULTRACARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALDONOULTRACARTAOACEITAPGFUTURO) as FINALDONOULTRACARTAOACEITAPGFUTURO, sum(vw_cons_t_Financ.FINALTECRXESPECIE) as FINALTECRXESPECIE, sum(vw_cons_t_Financ.FINALTECRXCARTAONAOACEITA) as FINALTECRXCARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALTECRXCARTAOACEITAPGFUTURO) as FINALTECRXCARTAOACEITAPGFUTURO " & _
      " FROM GRPAGAMENTO INNER JOIN GRPGTO ON GRPAGAMENTO.PKID = GRPGTO.GRPAGAMENTOID " & _
      " INNER JOIN GR ON GR.PKID = GRPGTO.GRID " & _
      " INNER JOIN vw_cons_t_Financ ON GR.PKID = vw_cons_t_Financ.GRID " & _
      " INNER JOIN GRPROCEDIMENTO ON GR.PKID = GRPROCEDIMENTO.GRID " & _
      " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      " INNER JOIN PRESTADORPROCEDIMENTO ON PRESTADORPROCEDIMENTO.PROCEDIMENTOID = GRPROCEDIMENTO.PROCEDIMENTOID " & _
      "                                     AND PRESTADORPROCEDIMENTO.PRONTUARIOID = ATENDE.PRONTUARIOID " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = ATENDE.PRONTUARIOID " & _
      " WHERE GRPAGAMENTO.PKID = " & Formata_Dados(lngGRPAGAMENTOID, tpDados_Longo) & _
      " GROUP BY GRPAGAMENTO.PKID"
  strSql = "SELECT  sum(vw_cons_t_Financ.PgtoTotal) as PgtoTotal, " & _
      " sum(vw_cons_t_Financ.FINALCASACONSESPECIE) as FINALCASACONSESPECIE, sum(vw_cons_t_Financ.FINALPRESTCONSESPECIE) as FINALPRESTCONSESPECIE, sum(vw_cons_t_Financ.FINALDONORXCONSESPECIE) as FINALDONORXCONSESPECIE, sum(vw_cons_t_Financ.FINALTECRXCONSESPECIE) as FINALTECRXCONSESPECIE, " & _
      " sum(vw_cons_t_Financ.FINALDONOULTRACONSESPECIE) as FINALDONOULTRACONSESPECIE, sum(vw_cons_t_Financ.FINALCASACONSCARTAO) as FINALCASACONSCARTAO, sum(vw_cons_t_Financ.FINALPRESTCONSCARTAO) as FINALPRESTCONSCARTAO, " & _
      " sum(vw_cons_t_Financ.FINALDONORXCONSCARTAO) as FINALDONORXCONSCARTAO, sum(vw_cons_t_Financ.FINALTECRXCONSCARTAO) as FINALTECRXCONSCARTAO, sum(vw_cons_t_Financ.FINALDONOULTRACONSCARTAO) as FINALDONOULTRACONSCARTAO, " & _
      " sum(vw_cons_t_Financ.FINALCASACONSCONVENIO) as FINALCASACONSCONVENIO, sum(vw_cons_t_Financ.FINALPRESTCONSCONVENIO) as FINALPRESTCONSCONVENIO, " & _
      " sum(vw_cons_t_Financ.FINALDONORXCONSCONVENIO) as FINALDONORXCONSCONVENIO, sum(vw_cons_t_Financ.FINALTECRXCONSCONVENIO) as FINALTECRXCONSCONVENIO, sum(vw_cons_t_Financ.FINALDONOULTRACONSCONVENIO) as FINALDONOULTRACONSCONVENIO, " & _
      " sum(vw_cons_t_Financ.FINALCASAESPECIE) as FINALCASAESPECIE, sum(vw_cons_t_Financ.FINALCASACARTAO) as FINALCASACARTAO, sum(vw_cons_t_Financ.FINALCASACONVENIO) as FINALCASACONVENIO, sum(vw_cons_t_Financ.FINALPRESTESPECIE) as FINALPRESTESPECIE, sum(vw_cons_t_Financ.FINALPRESTCONVENIO) as FINALPRESTCONVENIO, sum(vw_cons_t_Financ.FINALPRESTCARTAONAOACEITA) as FINALPRESTCARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALPRESTCARTAOACEITAPGFUTURO) as FINALPRESTCARTAOACEITAPGFUTURO, sum(vw_cons_t_Financ.FINALDONORXESPECIE) as FINALDONORXESPECIE, sum(vw_cons_t_Financ.FINALDONORXCONVENIO) as FINALDONORXCONVENIO, sum(vw_cons_t_Financ.FINALDONORXCARTAONAOACEITA) as FINALDONORXCARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALDONORXCARTAOACEITAPGFUTURO) as FINALDONORXCARTAOACEITAPGFUTURO, sum(vw_cons_t_Financ.FINALDONOULTRAESPECIE) as FINALDONOULTRAESPECIE, sum(vw_cons_t_Financ.FINALDONOULTRACONVENIO) as FINALDONOULTRACONVENIO, sum(vw_cons_t_Financ.FINALDONOULTRACARTAONAOACEITA) as FINALDONOULTRACARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALDONOULTRACARTAOACEITAPGFUTURO) as FINALDONOULTRACARTAOACEITAPGFUTURO, sum(vw_cons_t_Financ.FINALTECRXESPECIE) as FINALTECRXESPECIE, sum(vw_cons_t_Financ.FINALTECRXCONVENIO) as FINALTECRXCONVENIO, sum(vw_cons_t_Financ.FINALTECRXCARTAONAOACEITA) as FINALTECRXCARTAONAOACEITA, " & _
      " sum(vw_cons_t_Financ.FINALTECRXCARTAOACEITAPGFUTURO) as FINALTECRXCARTAOACEITAPGFUTURO " & _
      " FROM GRPAGAMENTO INNER JOIN GRPGTO ON GRPAGAMENTO.PKID = GRPGTO.GRPAGAMENTOID " & _
      " INNER JOIN GR ON GR.PKID = GRPGTO.GRID " & _
      " INNER JOIN vw_cons_t_Financ ON GR.PKID = vw_cons_t_Financ.GRID " & _
      " WHERE GRPAGAMENTO.PKID = " & Formata_Dados(lngGRPAGAMENTOID, tpDados_Longo) & _
      " GROUP BY GRPAGAMENTO.PKID"

  '
  Set objRs = objGer.ExecutarSQL(strSql)
  '
  If Not objRs.EOF Then   'se já houver algum item
    'INCLUIR_VALOR_NO_MASK mskCasa, objRs.Fields("VALOR_CASA").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskTotal, objRs.Fields("PgtoTotal").Value, TpMaskMoeda
    'Consulta Espécie
    INCLUIR_VALOR_NO_MASK mskPrestConsEspecie, objRs.Fields("FINALPRESTCONSESPECIE").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskDonoRXConsEspecie, objRs.Fields("FINALDONORXCONSESPECIE").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskTecRXConsEspecie, objRs.Fields("FINALTECRXCONSESPECIE").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskDonoUltraConsEspecie, objRs.Fields("FINALDONOULTRACONSESPECIE").Value, TpMaskMoeda
    'Consulta Cartão
    INCLUIR_VALOR_NO_MASK mskPrestConsCartao, objRs.Fields("FINALPRESTCONSCARTAO").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskDonoRXConsCartao, objRs.Fields("FINALDONORXCONSCARTAO").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskTecRXConsCartao, objRs.Fields("FINALTECRXCONSCARTAO").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskDonoUltraConsCartao, objRs.Fields("FINALDONOULTRACONSCARTAO").Value, TpMaskMoeda
    'Consulta Convênio
    INCLUIR_VALOR_NO_MASK mskPrestConsConvenio, objRs.Fields("FINALPRESTCONSCONVENIO").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskDonoRXConsConvenio, objRs.Fields("FINALDONORXCONSCONVENIO").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskTecRXConsConvenio, objRs.Fields("FINALTECRXCONSCONVENIO").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskDonoUltraConsConvenio, objRs.Fields("FINALDONOULTRACONSCONVENIO").Value, TpMaskMoeda
    'Prestador espécie
    INCLUIR_VALOR_NO_MASK mskPrestEspecie, objRs.Fields("FINALPRESTESPECIE").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskPrestConvenio, objRs.Fields("FINALPRESTCONVENIO").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskPrestCartao, objRs.Fields("FINALPRESTCARTAONAOACEITA").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskPrestCartaoARec, objRs.Fields("FINALPRESTCARTAOACEITAPGFUTURO").Value, TpMaskMoeda
    'Prestador Dono RX
    INCLUIR_VALOR_NO_MASK mskDonoRXEspecie, objRs.Fields("FINALDONORXESPECIE").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskDonoRXConvenio, objRs.Fields("FINALDONORXCONVENIO").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskDonoRXCartao, objRs.Fields("FINALDONORXCARTAONAOACEITA").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskDonoRXCartaoARec, objRs.Fields("FINALDONORXCARTAOACEITAPGFUTURO").Value, TpMaskMoeda
    'Prestador Técnico RX
    INCLUIR_VALOR_NO_MASK mskTecRXEspecie, objRs.Fields("FINALTECRXESPECIE").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskTecRXConvenio, objRs.Fields("FINALTECRXCONVENIO").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskTecRXCartao, objRs.Fields("FINALTECRXCARTAONAOACEITA").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskTecRXCartaoARec, objRs.Fields("FINALTECRXCARTAOACEITAPGFUTURO").Value, TpMaskMoeda
    'Prestador Dono Ultra
    INCLUIR_VALOR_NO_MASK mskDonoUltraEspecie, objRs.Fields("FINALDONOULTRAESPECIE").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskDonoUltraConvenio, objRs.Fields("FINALDONOULTRACONVENIO").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskDonoUltraCartao, objRs.Fields("FINALDONOULTRACARTAONAOACEITA").Value, TpMaskMoeda
    INCLUIR_VALOR_NO_MASK mskDonoUltraCartaoARec, objRs.Fields("FINALDONOULTRACARTAOACEITAPGFUTURO").Value, TpMaskMoeda
    
    'TOTAIS Prestador
    INCLUIR_VALOR_NO_MASK mskTotalPrest, objRs.Fields("FINALPRESTCONSESPECIE").Value + _
                                         objRs.Fields("FINALPRESTCONSCARTAO").Value + _
                                         objRs.Fields("FINALPRESTCONSCONVENIO").Value + _
                                         objRs.Fields("FINALPRESTESPECIE").Value + _
                                         objRs.Fields("FINALPRESTCONVENIO").Value + _
                                         objRs.Fields("FINALPRESTCARTAONAOACEITA").Value, TpMaskMoeda
    'TOTAIS Dono RX
    INCLUIR_VALOR_NO_MASK mskTotalDonoRX, objRs.Fields("FINALDONORXCONSESPECIE").Value + _
                                         objRs.Fields("FINALDONORXCONSCARTAO").Value + _
                                         objRs.Fields("FINALDONORXCONSCONVENIO").Value + _
                                         objRs.Fields("FINALDONORXESPECIE").Value + _
                                         objRs.Fields("FINALDONORXCONVENIO").Value + _
                                         objRs.Fields("FINALDONORXCARTAONAOACEITA").Value, TpMaskMoeda
    'TOTAIS Tecnico RX
    INCLUIR_VALOR_NO_MASK mskTotalTecRX, objRs.Fields("FINALTECRXCONSESPECIE").Value + _
                                         objRs.Fields("FINALTECRXCONSCARTAO").Value + _
                                         objRs.Fields("FINALTECRXCONSCONVENIO").Value + _
                                         objRs.Fields("FINALTECRXESPECIE").Value + _
                                         objRs.Fields("FINALTECRXCONVENIO").Value + _
                                         objRs.Fields("FINALTECRXCARTAONAOACEITA").Value, TpMaskMoeda
    'TOTAIS Dono Ultra
    INCLUIR_VALOR_NO_MASK mskTotalDonoUltra, objRs.Fields("FINALDONOULTRACONSESPECIE").Value + _
                                         objRs.Fields("FINALDONOULTRACONSCARTAO").Value + _
                                         objRs.Fields("FINALDONOULTRACONSCONVENIO").Value + _
                                         objRs.Fields("FINALDONOULTRAESPECIE").Value + _
                                         objRs.Fields("FINALDONOULTRACONVENIO").Value + _
                                         objRs.Fields("FINALDONOULTRACARTAONAOACEITA").Value, TpMaskMoeda
    
    'TOTAIS Receber
    INCLUIR_VALOR_NO_MASK mskTotalRec, CCur(mskPrestConsEspecie.FormattedText) + _
                                         CCur(mskPrestConsCartao.FormattedText) + _
                                         CCur(mskPrestConsConvenio.FormattedText) + _
                                         CCur(mskPrestEspecie.FormattedText) + _
                                         CCur(mskPrestCartao.FormattedText) + _
                                         CCur(mskPrestConvenio.FormattedText) + _
                                         CCur(mskDonoRXConsEspecie.FormattedText) + _
                                         CCur(mskDonoRXConsCartao.FormattedText) + _
                                         CCur(mskDonoRXConsConvenio.FormattedText) + _
                                         CCur(mskDonoRXEspecie.FormattedText) + _
                                         CCur(mskDonoRXCartao.FormattedText) + _
                                         CCur(mskDonoRXConvenio.FormattedText) + _
                                         CCur(mskTecRXConsEspecie.FormattedText) + _
                                         CCur(mskTecRXConsCartao.FormattedText) + _
                                         CCur(mskTecRXConsConvenio.FormattedText) + _
                                         CCur(mskTecRXEspecie.FormattedText) + _
                                         CCur(mskTecRXCartao.FormattedText) + _
                                         CCur(mskTecRXConvenio.FormattedText) + _
                                         CCur(mskDonoUltraConsEspecie.FormattedText) + _
                                         CCur(mskDonoUltraConsCartao.FormattedText) + _
                                         CCur(mskDonoUltraConsConvenio.FormattedText) + _
                                         CCur(mskDonoUltraEspecie.FormattedText) + _
                                         CCur(mskDonoUltraCartao.FormattedText) + _
                                         CCur(mskDonoUltraConvenio.FormattedText), TpMaskMoeda
    
    'TOTAIS A Receber (Futuramente)
    INCLUIR_VALOR_NO_MASK mskTotalARec, CCur(mskPrestCartaoARec.FormattedText) + _
                                        CCur(mskDonoRXCartaoARec.FormattedText) + _
                                        CCur(mskTecRXCartaoARec.FormattedText) + _
                                        CCur(mskDonoUltraCartaoARec.FormattedText), TpMaskMoeda
    'Casa
    INCLUIR_VALOR_NO_MASK mskCasa, objRs.Fields("FINALCASACONSESPECIE").Value + _
                                   objRs.Fields("FINALCASACONSCARTAO").Value + _
                                   objRs.Fields("FINALCASACONSCONVENIO").Value + _
                                   objRs.Fields("FINALCASAESPECIE").Value + _
                                   objRs.Fields("FINALCASACARTAO").Value + _
                                   objRs.Fields("FINALCASACONVENIO").Value, TpMaskMoeda
    'Verifica status
    'Select Case icTipoGR
    'Case tpIcTipoGR_DonoRX: INCLUIR_VALOR_NO_MASK mskPrestador, objRs.Fields("VALOR_DONO_RX").Value, TpMaskMoeda
    'Case tpIcTipoGR_DonoUltra: INCLUIR_VALOR_NO_MASK mskPrestador, objRs.Fields("VALOR_DONO_ULTRA").Value, TpMaskMoeda
    'Case tpIcTipoGR_Prest: INCLUIR_VALOR_NO_MASK mskPrestador, objRs.Fields("VALOR_PREST").Value, TpMaskMoeda
    'Case tpIcTipoGR_TecRX: INCLUIR_VALOR_NO_MASK mskPrestador, objRs.Fields("VALOR_TEC_RX").Value, TpMaskMoeda
    'Case Else: INCLUIR_VALOR_NO_MASK mskPrestador, "0", TpMaskMoeda
    'End Select
    
  End If
  objRs.Close
  Set objRs = Nothing
  Set objGer = Nothing
  '
  INCLUIR_VALOR_NO_MASK mskTotPagas, GRASSOC_LINHASMATRIZ, TpMaskLongo
  INCLUIR_VALOR_NO_MASK mskTotAPagar, GR_LINHASMATRIZ, TpMaskLongo
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
End Sub

Public Sub GR_MontaMatriz(strDataIni As String, _
                          strDataFim As String)
  Dim strSql    As String
  Dim objRs     As ADODB.Recordset
  Dim intI      As Integer
  Dim intJ      As Integer
  Dim clsGer    As busSisMed.clsGeral
  '
  On Error GoTo trata
  
  Set clsGer = New busSisMed.clsGeral
  '
  strSql = "SELECT GR.PKID, MAX(PRONTUARIO.NOME) AS NOME, MAX(GR.SEQUENCIAL) AS SEQUENCIAL, MAX(GR.SENHA) AS SENHA, MAX(GR.DATA) AS DATA, SUM(GRPROCEDIMENTO.VALOR) AS VALOR " & _
      " From GR " & _
      " INNER JOIN GRPROCEDIMENTO ON GR.PKID = GRPROCEDIMENTO.GRID " & _
      " INNER JOIN ATENDE ON ATENDE.PKID = GR.ATENDEID " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = ATENDE.PRONTUARIOID " & _
      " INNER JOIN TURNO ON TURNO.PKID = GR.TURNOID " & _
      " INNER JOIN PRESTADORPROCEDIMENTO ON PRESTADORPROCEDIMENTO.PROCEDIMENTOID = GRPROCEDIMENTO.PROCEDIMENTOID " & _
      "           AND PRESTADORPROCEDIMENTO.PRONTUARIOID = ATENDE.PRONTUARIOID " & _
      " WHERE GR.PKID NOT IN (SELECT GRPGTO.GRID FROM GRPGTO INNER JOIN GRPAGAMENTO ON GRPAGAMENTO.PKID = GRPGTO.GRPAGAMENTOID " & _
                                      " WHERE GR.PKID = GRPGTO.GRID "
  If icTipoGR = tpIcTipoGR_CancPont Or icTipoGR = tpIcTipoGR_CancAut Then
    'strSql = strSql & " AND GRPAGAMENTO.STATUS = " & Formata_Dados(tpIcTipoGR_Prest, tpDados_Texto) & ") "
    strSql = strSql & " ) "
  Else
    strSql = strSql & " AND GRPAGAMENTO.STATUS = " & Formata_Dados(strStatusGR, tpDados_Texto) & ") "
  End If
  strSql = strSql & " AND TURNO.DATA >= " & Formata_Dados(strDataIni, tpDados_DataHora) & _
      " AND TURNO.DATA < " & Formata_Dados(strDataFim, tpDados_DataHora) & _
      " AND GR.STATUS = " & Formata_Dados("F", tpDados_Texto)
  If strPrestador & "" <> "" Then
    If icTipoGR = tpIcTipoGR_Prest _
      Or icTipoGR = tpIcTipoGR_CancPont _
      Or icTipoGR = tpIcTipoGR_CancAut Then
      strSql = strSql & " AND PRONTUARIO.NOME = " & Formata_Dados(strPrestador, tpDados_Texto)
    End If
  End If
  strSql = strSql & " GROUP BY GR.PKID "
  If icTipoGR = tpIcTipoGR_DonoRX Then
    strSql = strSql & " HAVING MAX(PRESTADORPROCEDIMENTO.PERCRX) > 0 "
  ElseIf icTipoGR = tpIcTipoGR_TecRX Then
    strSql = strSql & " HAVING MAX(PRESTADORPROCEDIMENTO.PERCTECRX) > 0 "
  ElseIf icTipoGR = tpIcTipoGR_DonoUltra Then
    strSql = strSql & " HAVING MAX(PRESTADORPROCEDIMENTO.PERCULTRA) > 0 "
  End If
  strSql = strSql & " ORDER BY GR.SENHA, GR.SEQUENCIAL"
  'strSql = strSql & " ORDER BY PRONTUARIO.NOME, GR.SEQUENCIAL, GR.DATA"
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
  ValidaCampos = True
  If Len(strMsg) <> 0 Then
    TratarErroPrevisto strMsg, "[frmUserGRPgtoInc.ValidaCampos]"
    ValidaCampos = False
  Else
    ValidaCampos = True
  End If
  Exit Function
trata:
  Err.Raise Err.Number, _
            Err.Source & ".[frmUserProntuarioInc.ValidaCampos]", _
            Err.Description
End Function

Private Sub cmdRelatorio_Click()
  On Error GoTo TratErro
  Dim strDataIniFormula     As String
  Dim strDataFimFormula     As String
  Dim datData               As Date
  Dim strPrestadorFormula   As String
  Dim strSql                As String
  Dim objGeral              As busSisMed.clsGeral
  Dim objRs                 As ADODB.Recordset
  Dim strDonoRX             As String
  Dim strTecRX              As String
  Dim strDonoUltra          As String

  If CCur(mskTotalPrest.FormattedText) = 0 And _
     CCur(mskTotalDonoRX.FormattedText) = 0 And _
     CCur(mskTotalTecRX.FormattedText) = 0 And _
     CCur(mskTotalDonoUltra.FormattedText) = 0 Then
    AmpN
    MsgBox "Não há valores a serem pagos para esta(s) GR(s) selecionada(s). !", vbOKOnly, TITULOSISTEMA
    SetarFoco cmdRelatorio
    Exit Sub
  End If
  If icTipoGR = tpIcTipoGR_CancAut Or icTipoGR = tpIcTipoGR_CancPont Then
    IMP_COMP_CANC_GR lngGRPAGAMENTOID, gsNomeEmpresa, 1
  Else
    'Obter nome dos prestadores
    Set objGeral = New busSisMed.clsGeral
    strSql = "  SELECT  PRONTUARIO.NOME AS NOME_DONO_ULTRA, " & _
      " DONORX.NOME_DONO_RX, TECRX.NOME_TEC_RX " & _
      " From PRESTADOR " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID " & _
      " Left Join " & _
      " (SELECT PRONTUARIO.NOME AS NOME_DONO_RX " & _
      " From PRESTADOR " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID " & _
      " WHERE PRESTADOR.FUNCAOID IN (SELECT PKID FROM FUNCAO " & _
      " WHERE UPPER(FUNCAO) = 'DONO RX')) AS DONORX ON 1=1 " & _
      " Left Join " & _
      " (SELECT " & _
      " PRONTUARIO.NOME AS NOME_TEC_RX " & _
      " From PRESTADOR " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID " & _
      " WHERE PRESTADOR.FUNCAOID IN (SELECT PKID FROM FUNCAO " & _
      " WHERE UPPER(FUNCAO) = 'TÉCNICO RX')) AS TECRX ON 1=1 " & _
      " WHERE PRESTADOR.FUNCAOID IN (SELECT PKID FROM FUNCAO " & _
      " WHERE UPPER(FUNCAO) = 'DONO ULTRASON') "
    Set objRs = objGeral.ExecutarSQL(strSql)
    If objRs.EOF Then
      strDonoRX = "Dono de RX não cadastrado"
      strTecRX = "Técnico de RX não cadastrado"
      strDonoUltra = "Dono de Ultrason não cadastrado"
    Else
      strDonoRX = objRs.Fields("NOME_DONO_RX").Value & ""
      strTecRX = objRs.Fields("NOME_TEC_RX").Value & ""
      strDonoUltra = objRs.Fields("NOME_DONO_ULTRA").Value & ""
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    '
    'REPORT 1
    Report1.Destination = 0 'Video
    Report1.CopiesToPrinter = 1
    Report1.WindowState = crptMaximized
    Report1.Formulas(3) = "MesAno = '" & IIf(mskDtInicio.Text <> mskDtTermino.Text, "realizado no período de  ", "realizado no dia  ") & mskDtInicio.Text & IIf(mskDtInicio.Text <> mskDtTermino.Text, " a " & mskDtTermino.Text & " " & mskHoraTermino, "") & "'"
    'PRESTADOR ESPÉCIE
    If Not (CCur(mskPrestConsEspecie.FormattedText) = 0 _
        And CCur(mskPrestEspecie.FormattedText) = 0) Then
      INCLUIR_VALOR_NO_MASK mskAuxiliar, CCur(mskPrestConsEspecie.FormattedText) + CCur(mskPrestEspecie.FormattedText), TpMaskMoeda
      Report1.Formulas(0) = "Prestador = '" & strPrestador & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskAuxiliar.Text = "", 0, Replace(Replace(mskAuxiliar.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskAuxiliar.Text, "Reais", "Real") & ")'"
      Report1.Formulas(4) = "Tipo = 'espécie'"
      '
      Report1.Action = 1
    End If
    'PRESTADOR CARTÃO
    If Not (CCur(mskPrestConsCartao.FormattedText) = 0 _
        And CCur(mskPrestCartao.FormattedText) = 0) Then
      INCLUIR_VALOR_NO_MASK mskAuxiliar, CCur(mskPrestConsCartao.FormattedText) + CCur(mskPrestCartao.FormattedText), TpMaskMoeda
      Report1.Formulas(0) = "Prestador = '" & strPrestador & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskAuxiliar.Text = "", 0, Replace(Replace(mskAuxiliar.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskAuxiliar.Text, "Reais", "Real") & ")'"
      Report1.Formulas(4) = "Tipo = 'cartão'"
      '
      Report1.Action = 1
    End If
    '
    'PRESTADOR CONVÊNIO
    If Not (CCur(mskPrestConsConvenio.FormattedText) = 0 _
        And CCur(mskPrestConvenio.FormattedText) = 0) Then
      INCLUIR_VALOR_NO_MASK mskAuxiliar, CCur(mskPrestConsConvenio.FormattedText) + CCur(mskPrestConvenio.FormattedText), TpMaskMoeda
      Report1.Formulas(0) = "Prestador = '" & strPrestador & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskAuxiliar.Text = "", 0, Replace(Replace(mskAuxiliar.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskAuxiliar.Text, "Reais", "Real") & ")'"
      Report1.Formulas(4) = "Tipo = 'convênio'"
      '
      Report1.Action = 1
    End If
    '
    'DONO DE RX ESPÉCIE
    If Not (CCur(mskDonoRXConsEspecie.FormattedText) = 0 _
        And CCur(mskDonoRXEspecie.FormattedText) = 0) Then
      INCLUIR_VALOR_NO_MASK mskAuxiliar, CCur(mskDonoRXConsEspecie.FormattedText) + CCur(mskDonoRXEspecie.FormattedText), TpMaskMoeda
      Report1.Formulas(0) = "Prestador = '" & strDonoRX & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskAuxiliar.Text = "", 0, Replace(Replace(mskAuxiliar.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskAuxiliar.Text, "Reais", "Real") & ")'"
      Report1.Formulas(4) = "Tipo = 'espécie'"
      '
      Report1.Action = 1
    End If
    'DONO DE RX CARTÃO
    If Not (CCur(mskDonoRXConsCartao.FormattedText) = 0 _
        And CCur(mskDonoRXCartao.FormattedText) = 0) Then
      INCLUIR_VALOR_NO_MASK mskAuxiliar, CCur(mskDonoRXConsCartao.FormattedText) + CCur(mskDonoRXCartao.FormattedText), TpMaskMoeda
      Report1.Formulas(0) = "Prestador = '" & strDonoRX & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskAuxiliar.Text = "", 0, Replace(Replace(mskAuxiliar.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskAuxiliar.Text, "Reais", "Real") & ")'"
      Report1.Formulas(4) = "Tipo = 'cartão'"
      '
      Report1.Action = 1
    End If
    'DONO DE RX CONVÊNIO
    If Not (CCur(mskDonoRXConsConvenio.FormattedText) = 0 _
        And CCur(mskDonoRXConvenio.FormattedText) = 0) Then
      INCLUIR_VALOR_NO_MASK mskAuxiliar, CCur(mskDonoRXConsConvenio.FormattedText) + CCur(mskDonoRXConvenio.FormattedText), TpMaskMoeda
      Report1.Formulas(0) = "Prestador = '" & strDonoRX & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskAuxiliar.Text = "", 0, Replace(Replace(mskAuxiliar.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskAuxiliar.Text, "Reais", "Real") & ")'"
      Report1.Formulas(4) = "Tipo = 'convênio'"
      '
      Report1.Action = 1
    End If
    'TECNICO DE RX ESPÉCIE
    If Not (CCur(mskTecRXConsEspecie.FormattedText) = 0 _
        And CCur(mskTecRXEspecie.FormattedText) = 0) Then
      INCLUIR_VALOR_NO_MASK mskAuxiliar, CCur(mskTecRXConsEspecie.FormattedText) + CCur(mskTecRXEspecie.FormattedText), TpMaskMoeda
      Report1.Formulas(0) = "Prestador = '" & strTecRX & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskAuxiliar.Text = "", 0, Replace(Replace(mskAuxiliar.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskAuxiliar.Text, "Reais", "Real") & ")'"
      Report1.Formulas(4) = "Tipo = 'espécie'"
      '
      Report1.Action = 1
    End If
    'TECNICO DE RX CARTÃO
    If Not (CCur(mskTecRXConsCartao.FormattedText) = 0 _
        And CCur(mskTecRXCartao.FormattedText) = 0) Then
      INCLUIR_VALOR_NO_MASK mskAuxiliar, CCur(mskTecRXConsCartao.FormattedText) + CCur(mskTecRXCartao.FormattedText), TpMaskMoeda
      Report1.Formulas(0) = "Prestador = '" & strTecRX & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskAuxiliar.Text = "", 0, Replace(Replace(mskAuxiliar.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskAuxiliar.Text, "Reais", "Real") & ")'"
      Report1.Formulas(4) = "Tipo = 'cartão'"
      '
      Report1.Action = 1
    End If
    'TECNICO DE RX CONVÊNIO
    If Not (CCur(mskTecRXConsConvenio.FormattedText) = 0 _
        And CCur(mskTecRXConvenio.FormattedText) = 0) Then
      INCLUIR_VALOR_NO_MASK mskAuxiliar, CCur(mskTecRXConsConvenio.FormattedText) + CCur(mskTecRXConvenio.FormattedText), TpMaskMoeda
      Report1.Formulas(0) = "Prestador = '" & strTecRX & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskAuxiliar.Text = "", 0, Replace(Replace(mskAuxiliar.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskAuxiliar.Text, "Reais", "Real") & ")'"
      Report1.Formulas(4) = "Tipo = 'convênio'"
      '
      Report1.Action = 1
    End If
    'DONO DE ULTRASON ESPÉCIE
    If Not (CCur(mskDonoUltraConsEspecie.FormattedText) = 0 _
        And CCur(mskDonoUltraEspecie.FormattedText) = 0) Then
      INCLUIR_VALOR_NO_MASK mskAuxiliar, CCur(mskDonoUltraConsEspecie.FormattedText) + CCur(mskDonoUltraEspecie.FormattedText), TpMaskMoeda
      Report1.Formulas(0) = "Prestador = '" & strDonoUltra & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskAuxiliar.Text = "", 0, Replace(Replace(mskAuxiliar.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskAuxiliar.Text, "Reais", "Real") & ")'"
      Report1.Formulas(4) = "Tipo = 'espécie'"
      '
      Report1.Action = 1
    End If
    'DONO DE ULTRASON CARTÃO
    If Not (CCur(mskDonoUltraConsCartao.FormattedText) = 0 _
        And CCur(mskDonoUltraCartao.FormattedText) = 0) Then
      INCLUIR_VALOR_NO_MASK mskAuxiliar, CCur(mskDonoUltraConsCartao.FormattedText) + CCur(mskDonoUltraCartao.FormattedText), TpMaskMoeda
      Report1.Formulas(0) = "Prestador = '" & strDonoUltra & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskAuxiliar.Text = "", 0, Replace(Replace(mskAuxiliar.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskAuxiliar.Text, "Reais", "Real") & ")'"
      Report1.Formulas(4) = "Tipo = 'cartão'"
      '
      Report1.Action = 1
    End If
    'DONO DE ULTRASON CARTÃO
    If Not (CCur(mskDonoUltraConsConvenio.FormattedText) = 0 _
        And CCur(mskDonoUltraConvenio.FormattedText) = 0) Then
      INCLUIR_VALOR_NO_MASK mskAuxiliar, CCur(mskDonoUltraConsConvenio.FormattedText) + CCur(mskDonoUltraConvenio.FormattedText), TpMaskMoeda
      Report1.Formulas(0) = "Prestador = '" & strDonoUltra & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskAuxiliar.Text = "", 0, Replace(Replace(mskAuxiliar.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskAuxiliar.Text, "Reais", "Real") & ")'"
      Report1.Formulas(4) = "Tipo = 'convênio'"
      '
      Report1.Action = 1
    End If
    
    'REPORT 2
    '
    datData = CDate(mskDtInicio.Text & " " & mskHoraInicio.Text)
    strDataIniFormula = mskDtInicio.Text & " " & mskHoraInicio.Text
    strDataFimFormula = mskDtTermino.Text & " " & mskHoraTermino.Text
    'If cboPrestador.Text = "" Then
    '  strPrestadorFormula = "Prestador = True = true"
    'Else
      strPrestadorFormula = "Prestador = {PRONTUARIO.NOME} = '" & strPrestador & "'"
    'End If
    '
    Report2.Destination = 0 'Video
    Report2.CopiesToPrinter = 1
    Report2.WindowState = crptMaximized
    '
    Report2.Formulas(0) = "GRPAGAMENTOID = " & lngGRPAGAMENTOID
    Report2.Formulas(1) = "Prestador = '" & strPrestador & "'"
    Report2.Formulas(2) = "icTipoGR = " & icTipoGR
    'Report2.Formulas(0) = "Prestador = '" & strPrestador & "'"
    'Report2.Formulas(1) = "DataBaseIni = Date(" & Mid(strDataIniFormula, 7, 4) & ", " & Mid(strDataIniFormula, 4, 2) & ", " & Left(strDataIniFormula, 2) & ")"
    'Report2.Formulas(2) = "DataBaseFim = Date(" & Mid(strDataFimFormula, 7, 4) & ", " & Mid(strDataFimFormula, 4, 2) & ", " & Left(strDataFimFormula, 2) & ")"
    'Report2.Formulas(3) = "GRPGTOSTATUS = '" & strStatusGR & "'"
    'Report2.Formulas(4) = strPrestadorFormula
    '
    Report2.Action = 1
  End If
  '
  AmpN
  Exit Sub
  
TratErro:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Private Sub cmdRelatorioARec_Click()
  On Error GoTo TratErro
  Dim strDataIniFormula     As String
  Dim strDataFimFormula     As String
  Dim datData               As Date
  Dim datDataARec           As Date
  Dim strPrestadorFormula   As String
  Dim strSql                As String
  Dim objGeral              As busSisMed.clsGeral
  Dim objRs                 As ADODB.Recordset
  Dim strDonoRX             As String
  Dim strTecRX              As String
  Dim strDonoUltra          As String

  If CCur(mskPrestCartaoARec.FormattedText) = 0 And _
     CCur(mskDonoRXCartaoARec.FormattedText) = 0 And _
     CCur(mskTecRXCartaoARec.FormattedText) = 0 And _
     CCur(mskDonoUltraCartaoARec.FormattedText) = 0 Then
    AmpN
    MsgBox "Não há valores a pagar para esta(s) GR(s) selecionada(s). !", vbOKOnly, TITULOSISTEMA
    SetarFoco cmdRelatorio
    Exit Sub
  End If
  datDataARec = CDate(mskDtAReceb.Text)
  If Now < datDataARec Then
    AmpN
    MsgBox "GRs selecionadas só poderão ser impressas após " & mskDtAReceb.Text & "!", vbOKOnly, TITULOSISTEMA
    SetarFoco cmdRelatorio
    Exit Sub
  End If
  

  If icTipoGR = tpIcTipoGR_CancAut Or icTipoGR = tpIcTipoGR_CancPont Then
    IMP_COMP_CANC_GR lngGRPAGAMENTOID, gsNomeEmpresa, 1
  Else
    'Obter nome dos prestadores
    Set objGeral = New busSisMed.clsGeral
    strSql = "  SELECT  PRONTUARIO.NOME AS NOME_DONO_ULTRA, " & _
      " DONORX.NOME_DONO_RX, TECRX.NOME_TEC_RX " & _
      " From PRESTADOR " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID " & _
      " Left Join " & _
      " (SELECT PRONTUARIO.NOME AS NOME_DONO_RX " & _
      " From PRESTADOR " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID " & _
      " WHERE PRESTADOR.FUNCAOID IN (SELECT PKID FROM FUNCAO " & _
      " WHERE UPPER(FUNCAO) = 'DONO RX')) AS DONORX ON 1=1 " & _
      " Left Join " & _
      " (SELECT " & _
      " PRONTUARIO.NOME AS NOME_TEC_RX " & _
      " From PRESTADOR " & _
      " INNER JOIN PRONTUARIO ON PRONTUARIO.PKID = PRESTADOR.PRONTUARIOID " & _
      " WHERE PRESTADOR.FUNCAOID IN (SELECT PKID FROM FUNCAO " & _
      " WHERE UPPER(FUNCAO) = 'TÉCNICO RX')) AS TECRX ON 1=1 " & _
      " WHERE PRESTADOR.FUNCAOID IN (SELECT PKID FROM FUNCAO " & _
      " WHERE UPPER(FUNCAO) = 'DONO ULTRASON') "
    Set objRs = objGeral.ExecutarSQL(strSql)
    If objRs.EOF Then
      strDonoRX = "Dono de RX não cadastrado"
      strTecRX = "Técnico de RX não cadastrado"
      strDonoUltra = "Dono de Ultrason não cadastrado"
    Else
      strDonoRX = objRs.Fields("NOME_DONO_RX").Value & ""
      strTecRX = objRs.Fields("NOME_TEC_RX").Value & ""
      strDonoUltra = objRs.Fields("NOME_DONO_ULTRA").Value & ""
    End If
    objRs.Close
    Set objRs = Nothing
    Set objGeral = Nothing
    '
    'REPORT 1
    Report1.Destination = 0 'Video
    Report1.CopiesToPrinter = 1
    Report1.WindowState = crptMaximized
    'PRESTADOR
    If CCur(mskPrestCartaoARec.FormattedText) <> 0 Then
      Report1.Formulas(0) = "Prestador = '" & strPrestador & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskPrestCartaoARec.Text = "", 0, Replace(Replace(mskPrestCartaoARec.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskPrestCartaoARec.Text, "Reais", "Real") & ")'"
      '
    End If
    
    Report1.Formulas(3) = "MesAno = '" & IIf(mskDtInicio.Text <> mskDtTermino.Text, "realizado no período de  ", "realizado no dia  ") & mskDtInicio.Text & IIf(mskDtInicio.Text <> mskDtTermino.Text, " a " & mskDtTermino.Text & " " & mskHoraTermino, "") & "'"
    '
    If CCur(mskPrestCartaoARec.FormattedText) <> 0 Then
      Report1.Action = 1
    End If
    '
    'DONO DE RX
    If CCur(mskDonoRXCartaoARec.FormattedText) <> 0 Then
      Report1.Formulas(0) = "Prestador = '" & strDonoRX & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskDonoRXCartaoARec.Text = "", 0, Replace(Replace(mskDonoRXCartaoARec.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskDonoRXCartaoARec.Text, "Reais", "Real") & ")'"
      '
      Report1.Action = 1
    End If
    'TECNICO DE RX
    If CCur(mskTecRXCartaoARec.FormattedText) <> 0 Then
      Report1.Formulas(0) = "Prestador = '" & strTecRX & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskTecRXCartaoARec.Text = "", 0, Replace(Replace(mskTecRXCartaoARec.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskTecRXCartaoARec.Text, "Reais", "Real") & ")'"
      '
      Report1.Action = 1
    End If
    'DONO DE ULTRASON
    If CCur(mskDonoUltraCartaoARec.FormattedText) <> 0 Then
      Report1.Formulas(0) = "Prestador = '" & strDonoUltra & "'"
      Report1.Formulas(1) = "Valor = " & IIf(mskDonoUltraCartaoARec.Text = "", 0, Replace(Replace(mskDonoUltraCartaoARec.Text, ".", ""), ",", "."))
      Report1.Formulas(2) = "Extenso = '(" & Extenso(mskDonoUltraCartaoARec.Text, "Reais", "Real") & ")'"
      '
      Report1.Action = 1
    End If
    
    'REPORT 2
    '
    datData = CDate(mskDtInicio.Text & " " & mskHoraInicio.Text)
    strDataIniFormula = mskDtInicio.Text & " " & mskHoraInicio.Text
    strDataFimFormula = mskDtTermino.Text & " " & mskHoraTermino.Text
    'If cboPrestador.Text = "" Then
    '  strPrestadorFormula = "Prestador = True = true"
    'Else
      strPrestadorFormula = "Prestador = {PRONTUARIO.NOME} = '" & strPrestador & "'"
    'End If
    
    '
    Report2.Destination = 0 'Video
    Report2.CopiesToPrinter = 1
    Report2.WindowState = crptMaximized
    '
    Report2.Formulas(0) = "GRPAGAMENTOID = " & lngGRPAGAMENTOID
    Report2.Formulas(1) = "Prestador = '" & strPrestador & "'"
    Report2.Formulas(2) = "icTipoGR = " & icTipoGR
    'Report2.Formulas(0) = "Prestador = '" & strPrestador & "'"
    'Report2.Formulas(1) = "DataBaseIni = Date(" & Mid(strDataIniFormula, 7, 4) & ", " & Mid(strDataIniFormula, 4, 2) & ", " & Left(strDataIniFormula, 2) & ")"
    'Report2.Formulas(2) = "DataBaseFim = Date(" & Mid(strDataFimFormula, 7, 4) & ", " & Mid(strDataFimFormula, 4, 2) & ", " & Left(strDataFimFormula, 2) & ")"
    'Report2.Formulas(3) = "GRPGTOSTATUS = '" & strStatusGR & "'"
    'Report2.Formulas(4) = strPrestadorFormula
    '
    Report2.Action = 1
  End If
  '
  AmpN
  Exit Sub
  
TratErro:
  TratarErro Err.Number, _
             Err.Description, _
             Err.Source
  AmpN
End Sub

Private Sub grdGRAssoc_UnboundReadDataEx( _
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
               Offset + intI, GRASSOC_LINHASMATRIZ)
  
    ' If the next row is BOF or EOF, then stop fetching
    ' and return any rows fetched up to this point.
    If IsNull(vrtBookmark) Then Exit For
  
    ' Place the record data into the row buffer
    For intJ = 0 To RowBuf.ColumnCount - 1
      intColIndex = RowBuf.ColumnIndex(intI, intJ)
      RowBuf.Value(intI, intJ) = GetUserDataGeral(vrtBookmark, _
                           intColIndex, GRASSOC_COLUNASMATRIZ, GRASSOC_LINHASMATRIZ, GRASSOC_Matriz)
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
  lngNewPosition = IndexFromBookmarkGeral(StartLocation, Offset, GRASSOC_LINHASMATRIZ)
  If lngNewPosition >= 0 Then _
     ApproximatePosition = lngNewPosition
     
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, "[frmUserGRPgtoInc.grdGR_UnboundReadDataEx]"
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
  TratarErro Err.Number, Err.Description, "[frmUserGRPgtoInc.grdGR_UnboundReadDataEx]"
End Sub

Private Sub Form_Load()
  On Error GoTo trata
  Dim strSql      As String
  Dim objRs       As ADODB.Recordset
  Dim objGeral    As busSisMed.clsGeral
  Dim datARec     As Date
  '
  blnFechar = False 'Não Pode Fechar pelo X
  blnRetorno = False
  AmpS
  Me.Height = 8925
  Me.Width = 11610
  CenterForm Me
  Me.Caption = Me.Caption & " - " & strGR
  If Status = tpStatus_Consultar Then
    cmdRelatorio.Enabled = False
    cmdRelatorioARec.Enabled = False
    cmdCadastraItem(0).Enabled = False
    cmdCadastraItem(1).Enabled = False
    cmdCadastraItem(2).Enabled = False
    cmdCadastraItem(3).Enabled = False
  Else
    cmdRelatorio.Enabled = True
    cmdRelatorioARec.Enabled = True
    cmdCadastraItem(0).Enabled = True
    cmdCadastraItem(1).Enabled = True
    cmdCadastraItem(2).Enabled = True
    cmdCadastraItem(3).Enabled = True
  End If

  'Limpar Campos
  LimparCampoMask mskDtInicio
  LimparCampoMask mskHoraInicio
  LimparCampoMask mskDtTermino
  LimparCampoMask mskHoraTermino
  LimparCampoTexto txtPrestador
  LimparCampoTexto txtAcCartao
  '
  'LimparCampoMask mskCasa
  'LimparCampoMask mskPrestador
  'LimparCampoMask mskTotal
  'LimparCampoMask mskTotPagas
  'LimparCampoMask mskTotAPagar
  
  'Tratar Campos
  INCLUIR_VALOR_NO_MASK mskDtInicio, strDataIni, TpMaskData
  INCLUIR_VALOR_NO_MASK mskHoraInicio, strHoraIni, TpMaskData
  INCLUIR_VALOR_NO_MASK mskDtTermino, strDataFim, TpMaskData
  INCLUIR_VALOR_NO_MASK mskHoraTermino, strHoraFim, TpMaskData
  txtPrestador.Text = strPrestador
  '
  datARec = CDate(strDataIni)
  datARec = DateAdd("m", 1, datARec)
  INCLUIR_VALOR_NO_MASK mskDtAReceb, Format(datARec, "DD/MM/YYYY"), TpMaskData
  '
  'Verifica status
  Select Case icTipoGR
  Case tpIcTipoGR_DonoRX: strStatusGR = "DR"
  Case tpIcTipoGR_DonoUltra: strStatusGR = "DU"
  Case tpIcTipoGR_Prest: strStatusGR = "PG"
  Case tpIcTipoGR_TecRX: strStatusGR = "TR"
  Case tpIcTipoGR_CancPont: strStatusGR = "CP"
  Case tpIcTipoGR_CancAut: strStatusGR = "CA"
  Case Else: strStatusGR = ""
  End Select
  '
  LerFiguras Me, tpBmp_Vazio, , , cmdFechar, pbtnImprimir:=cmdRelatorio
  LerFiguras Me, tpBmp_Vazio, , , pbtnImprimir:=cmdRelatorioARec
  'Obter campos
  Set objGeral = New busSisMed.clsGeral
  
  strSql = "SELECT case PRESTADOR.INDACEITACHEQUE WHEN 'S' THEN 'SIM' ELSE 'NÃO' END AS INDACEITACHEQUE " & _
            " From GRPAGAMENTO " & _
            " INNER JOIN PRESTADOR ON PRESTADOR.PRONTUARIOID = GRPAGAMENTO.PRESTADORID " & _
            " WHERE GRPAGAMENTO.PKID = " & Formata_Dados(lngGRPAGAMENTOID, tpDados_Longo)
  Set objRs = objGeral.ExecutarSQL(strSql)
  If Not objRs.EOF Then
    txtAcCartao.Text = objRs.Fields("INDACEITACHEQUE").Value
  End If
  objRs.Close
  
  Set objGeral = Nothing
  '
  Report1.Connect = ConnectRpt
  Report1.ReportFileName = gsReportPath & "Recibo.rpt"
  '
  Report2.Connect = ConnectRpt
  Report2.ReportFileName = gsReportPath & "DemoReceb.rpt"
  
  GR_COLUNASMATRIZ = grdGR.Columns.Count
  GR_LINHASMATRIZ = 0
  GR_MontaMatriz strDataIni & " " & strHoraIni, strDataFim & " " & strHoraFim
  grdGR.ApproxCount = GR_LINHASMATRIZ
  '
  '
  GRASSOC_COLUNASMATRIZ = grdGRAssoc.Columns.Count
  GRASSOC_LINHASMATRIZ = 0
  GRASSOC_MontaMatriz
  grdGRAssoc.ApproxCount = GRASSOC_LINHASMATRIZ
  '

  AmpN
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub TratarAssociacao(pIndice As Integer)
  On Error GoTo trata
  Dim intI          As Long
  Dim objGRPgto     As busSisMed.clsGRPgto
  Dim lngRet        As Long
  Dim blnRet        As Boolean
  Dim intExc        As Long
  Dim strDataIniFormula    As String
  Dim strDataFimFormula    As String
  Dim datData       As Date
  '
  Set objGRPgto = New busSisMed.clsGRPgto
  '
  blnRet = False
  intExc = 0
  '
  Select Case pIndice
  Case 1 'Cadastrar Selecionados
    For intI = 0 To grdGR.SelBookmarks.Count - 1
      grdGR.Bookmark = CLng(grdGR.SelBookmarks.Item(intI))
      'Verificar se item possui estoue suficiente
      objGRPgto.AssociarGRPGTOGR grdGR.Columns("GRID").Text, _
                                 lngGRPAGAMENTOID
      blnRet = True
    Next
  Case 2 'Cadastrar Todos
    For intI = 0 To GR_LINHASMATRIZ - 1
      grdGR.Bookmark = CLng(intI)
      objGRPgto.AssociarGRPGTOGR grdGR.Columns("GRID").Text, _
                                 lngGRPAGAMENTOID
      blnRet = True
    Next
  Case 3 'Retirar Selecionados
    For intI = 0 To grdGRAssoc.SelBookmarks.Count - 1
      grdGRAssoc.Bookmark = CLng(grdGRAssoc.SelBookmarks.Item(intI))
      objGRPgto.DesassociarGRPGTOGR grdGRAssoc.Columns("GRPGTOID").Text
      blnRet = True
    Next
  Case 4 'retirar Todos
    For intI = 0 To GRASSOC_LINHASMATRIZ - 1
      grdGRAssoc.Bookmark = CLng(intI)
      If IsNull(grdGRAssoc.Bookmark) Then grdGRAssoc.Bookmark = CLng(intI)
      objGRPgto.DesassociarGRPGTOGR grdGRAssoc.Columns("GRPGTOID").Text
      blnRet = True
    Next
  End Select
  '
  Set objGRPgto = Nothing
    '
  If blnRet Then 'Houve Auteração, Atualiza grids
    blnRetorno = True
    '
    If Not ValidaCampos Then
      Exit Sub
    End If
    '
    datData = CDate(mskDtInicio.Text & " " & mskHoraInicio.Text)
    strDataIniFormula = mskDtInicio.Text & " " & mskHoraInicio.Text
    strDataFimFormula = mskDtTermino.Text & " " & mskHoraTermino.Text
    '
    GR_COLUNASMATRIZ = grdGR.Columns.Count
    GR_LINHASMATRIZ = 0
    GR_MontaMatriz strDataIniFormula, strDataFimFormula
    grdGR.Bookmark = Null
    grdGR.ReBind
    grdGR.ApproxCount = GR_LINHASMATRIZ
    '
    '
    GRASSOC_COLUNASMATRIZ = grdGRAssoc.Columns.Count
    GRASSOC_LINHASMATRIZ = 0
    GRASSOC_MontaMatriz
    grdGRAssoc.Bookmark = Null
    grdGRAssoc.ReBind
    grdGRAssoc.ApproxCount = GRASSOC_LINHASMATRIZ
    '
  End If
  '
  Exit Sub
trata:
  TratarErro Err.Number, Err.Description, Err.Source
  AmpN
End Sub

Private Sub Form_Unload(Cancel As Integer)
  If Not blnFechar Then Cancel = True
End Sub
