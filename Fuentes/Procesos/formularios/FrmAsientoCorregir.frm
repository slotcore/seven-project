VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Begin VB.Form FrmAsientoCorregir 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Herramientas - Corregir Asientos"
   ClientHeight    =   7530
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11790
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7530
   ScaleWidth      =   11790
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   6435
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   735
      Left            =   2835
      TabIndex        =   1
      Top             =   3210
      Visible         =   0   'False
      Width           =   5805
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   285
         Left            =   90
         TabIndex        =   2
         Top             =   360
         Width           =   5655
         _ExtentX        =   9975
         _ExtentY        =   503
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Height          =   195
         Left            =   4710
         TabIndex        =   16
         Top             =   150
         Width           =   975
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Corregiendo Asientos"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   135
         Width           =   1500
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   5820
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   1
         X1              =   5790
         X2              =   5790
         Y1              =   15
         Y2              =   1155
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   30
         Y2              =   1170
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5820
         Y1              =   15
         Y2              =   0
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsientoCorregir.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsientoCorregir.frx":0544
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsientoCorregir.frx":08D6
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsientoCorregir.frx":0A5A
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsientoCorregir.frx":0EAE
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsientoCorregir.frx":0FC6
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsientoCorregir.frx":150A
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsientoCorregir.frx":1A4E
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsientoCorregir.frx":1B62
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsientoCorregir.frx":1C76
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsientoCorregir.frx":20CA
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsientoCorregir.frx":2236
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmAsientoCorregir.frx":277E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11790
      _ExtentX        =   20796
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Frame Frame1 
      Height          =   1320
      Left            =   0
      TabIndex        =   4
      Top             =   285
      Width           =   11805
      Begin VB.CommandButton Command4 
         Caption         =   "Para Análisis de Cta Cte"
         Height          =   345
         Left            =   5250
         TabIndex        =   22
         Top             =   600
         Width           =   3495
      End
      Begin VB.CheckBox ChkReproceso 
         Caption         =   "Reporoceso"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8790
         TabIndex        =   21
         Top             =   750
         Visible         =   0   'False
         Width           =   2385
      End
      Begin VB.CheckBox chk 
         Caption         =   "Pendientes de Registro"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   255
         Left            =   8790
         TabIndex        =   20
         Top             =   540
         Width           =   2385
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Ajuste Dif Cambio - Orden Despacho"
         Height          =   345
         Left            =   150
         TabIndex        =   19
         Top             =   960
         Width           =   3495
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Ajuste Dif Cambio Bancos"
         Height          =   345
         Left            =   2340
         TabIndex        =   18
         Top             =   600
         Width           =   2085
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Redondeo a Centimos"
         Height          =   345
         Left            =   150
         TabIndex        =   17
         Top             =   600
         Width           =   2145
      End
      Begin VB.CommandButton CmdBancos 
         Caption         =   "Bancos"
         Height          =   345
         Left            =   7380
         TabIndex        =   15
         Top             =   210
         Width           =   1395
      End
      Begin VB.CommandButton CmdBusProv 
         Height          =   230
         Left            =   4980
         Picture         =   "FrmAsientoCorregir.frx":2B10
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   270
         Width           =   210
      End
      Begin VB.CommandButton CmdBusMes 
         Height          =   230
         Left            =   7140
         Picture         =   "FrmAsientoCorregir.frx":2C42
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   280
         Width           =   180
      End
      Begin VB.TextBox TxtMes 
         Height          =   285
         Left            =   5625
         Locked          =   -1  'True
         TabIndex        =   7
         Text            =   "TxtMes"
         Top             =   240
         Width           =   1740
      End
      Begin VB.TextBox TxtLibro 
         Height          =   300
         Left            =   510
         Locked          =   -1  'True
         TabIndex        =   0
         Text            =   "TxtLibro"
         Top             =   240
         Width           =   4695
      End
      Begin VB.Label lbltotalRegistros 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbltotalRegistros"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00400000&
         Height          =   315
         Left            =   9255
         TabIndex        =   14
         Top             =   210
         Width           =   2475
      End
      Begin VB.Label LblIdLibro 
         AutoSize        =   -1  'True
         Caption         =   "LblIdLibro"
         ForeColor       =   &H000000C0&
         Height          =   195
         Left            =   8670
         TabIndex        =   13
         Top             =   270
         Visible         =   0   'False
         Width           =   690
      End
      Begin VB.Label LblIdMes 
         AutoSize        =   -1  'True
         Caption         =   "LblIdMes"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5280
         TabIndex        =   12
         Top             =   120
         Visible         =   0   'False
         Width           =   645
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Libro"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   345
         Width           =   345
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Mes"
         Height          =   195
         Left            =   5250
         TabIndex        =   8
         Top             =   345
         Width           =   300
      End
   End
   Begin TrueOleDBGrid70.TDBGrid grilla 
      Height          =   5850
      Left            =   30
      TabIndex        =   11
      Top             =   1650
      Width           =   11745
      _ExtentX        =   20717
      _ExtentY        =   10319
      _LayoutType     =   4
      _RowHeight      =   13
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "Id"
      Columns(0).DataField=   "id"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   4
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Sel"
      Columns(1).DataField=   "sel"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Nº Reg."
      Columns(2).DataField=   "registro"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Nº R.U.C."
      Columns(3).DataField=   "ruc"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Cliente /Proveedor /Otros"
      Columns(4).DataField=   "nombre"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(5)._VlistStyle=   0
      Columns(5)._MaxComboItems=   5
      Columns(5).Caption=   "T.D."
      Columns(5).DataField=   "abrev"
      Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(6)._VlistStyle=   0
      Columns(6)._MaxComboItems=   5
      Columns(6).Caption=   "Fch. Doc."
      Columns(6).DataField=   "fchdoc"
      Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(7)._VlistStyle=   0
      Columns(7)._MaxComboItems=   5
      Columns(7).Caption=   "Nº Documento"
      Columns(7).DataField=   "numerodoc"
      Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(8)._VlistStyle=   0
      Columns(8)._MaxComboItems=   5
      Columns(8).Caption=   "M"
      Columns(8).DataField=   "moneda"
      Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(9)._VlistStyle=   0
      Columns(9)._MaxComboItems=   5
      Columns(9).Caption=   "Importe"
      Columns(9).DataField=   "total"
      Columns(9).NumberFormat=   "0.00"
      Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   10
      Splits(0)._UserFlags=   0
      Splits(0).MarqueeStyle=   3
      Splits(0).AllowSizing=   -1  'True
      Splits(0).RecordSelectorWidth=   318
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).FilterBar=   -1  'True
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=10"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
      Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
      Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(8)=   "Column(1).Width=635"
      Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=556"
      Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
      Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
      Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(14)=   "Column(2).Width=1588"
      Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1508"
      Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
      Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
      Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(20)=   "Column(3).Width=2434"
      Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2355"
      Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
      Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
      Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(26)=   "Column(4).Width=6720"
      Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=6641"
      Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
      Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
      Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
      Splits(0)._ColumnProps(32)=   "Column(5).Width=820"
      Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
      Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=741"
      Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
      Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
      Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
      Splits(0)._ColumnProps(38)=   "Column(6).Width=1746"
      Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
      Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1667"
      Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
      Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=513"
      Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
      Splits(0)._ColumnProps(44)=   "Column(7).Width=2752"
      Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
      Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2672"
      Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
      Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=516"
      Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
      Splits(0)._ColumnProps(50)=   "Column(8).Width=741"
      Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
      Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=661"
      Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
      Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=513"
      Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
      Splits(0)._ColumnProps(56)=   "Column(9).Width=1746"
      Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
      Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1667"
      Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
      Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=514"
      Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
      Splits.Count    =   1
      PrintInfos(0)._StateFlags=   3
      PrintInfos(0).Name=   "piInternal 0"
      PrintInfos(0).PageHeaderFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageFooterFont=   "Size=8.25,Charset=0,Weight=400,Underline=0,Italic=0,Strikethrough=0,Name=MS Sans Serif"
      PrintInfos(0).PageHeaderHeight=   0
      PrintInfos(0).PageFooterHeight=   0
      PrintInfos.Count=   1
      AllowUpdate     =   0   'False
      Appearance      =   0
      ColumnFooters   =   -1  'True
      DefColWidth     =   0
      HeadLines       =   1.5
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   0
      MaxRows         =   250000
      ViewColumnCaptionWidth=   0
      ViewColumnWidth =   0
      _PropDict       =   "_ExtentX,2003,3;_ExtentY,2004,3;_LayoutType,512,2;_RowHeight,16,3;_StyleDefs,513,0;_WasPersistedAsPixels,516,2"
      _StyleDefs(0)   =   "_StyleRoot:id=0,.parent=-1,.alignment=3,.valignment=0,.bgcolor=&H80000005&"
      _StyleDefs(1)   =   ":id=0,.fgcolor=&H80000008&,.wraptext=0,.locked=0,.transparentBmp=0"
      _StyleDefs(2)   =   ":id=0,.fgpicPosition=0,.bgpicMode=0,.appearance=0,.borderSize=0,.ellipsis=0"
      _StyleDefs(3)   =   ":id=0,.borderColor=&H80000005&,.borderType=0,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(4)   =   ":id=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(5)   =   ":id=0,.fontname=MS Sans Serif"
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0FEFE&,.fgcolor=&H0&,.bold=0"
      _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
      _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(12)  =   ":id=2,.fontname=MS Sans Serif"
      _StyleDefs(13)  =   "FooterStyle:id=3,.parent=1,.namedParent=35,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(14)  =   ":id=3,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(15)  =   ":id=3,.fontname=MS Sans Serif"
      _StyleDefs(16)  =   "InactiveStyle:id=5,.parent=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(17)  =   "SelectedStyle:id=6,.parent=1,.namedParent=36"
      _StyleDefs(18)  =   "EditorStyle:id=7,.parent=1"
      _StyleDefs(19)  =   "HighlightRowStyle:id=8,.parent=1,.namedParent=38"
      _StyleDefs(20)  =   "EvenRowStyle:id=9,.parent=1,.namedParent=39"
      _StyleDefs(21)  =   "OddRowStyle:id=10,.parent=1,.namedParent=40"
      _StyleDefs(22)  =   "RecordSelectorStyle:id=11,.parent=2,.namedParent=41"
      _StyleDefs(23)  =   "FilterBarStyle:id=12,.parent=1,.namedParent=42"
      _StyleDefs(24)  =   "Splits(0).Style:id=13,.parent=1"
      _StyleDefs(25)  =   "Splits(0).CaptionStyle:id=22,.parent=4"
      _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2"
      _StyleDefs(27)  =   "Splits(0).FooterStyle:id=15,.parent=3"
      _StyleDefs(28)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
      _StyleDefs(29)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
      _StyleDefs(30)  =   "Splits(0).EditorStyle:id=17,.parent=7"
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=70,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=67,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=68,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=69,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=74,.parent=13,.alignment=2"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=71,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=72,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=73,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=58,.parent=13"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=55,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=56,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=57,.parent=17"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=62,.parent=13,.alignment=0"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=59,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=60,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=61,.parent=17"
      _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=66,.parent=13,.alignment=2"
      _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=63,.parent=14"
      _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=64,.parent=15"
      _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=65,.parent=17"
      _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=28,.parent=13,.alignment=2"
      _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=25,.parent=14"
      _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=26,.parent=15"
      _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=27,.parent=17"
      _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13"
      _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
      _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
      _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
      _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=46,.parent=13,.alignment=2"
      _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=43,.parent=14"
      _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=44,.parent=15"
      _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=45,.parent=17"
      _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=54,.parent=13,.alignment=1"
      _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=51,.parent=14"
      _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=52,.parent=15"
      _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=53,.parent=17"
      _StyleDefs(76)  =   "Named:id=33:Normal"
      _StyleDefs(77)  =   ":id=33,.parent=0"
      _StyleDefs(78)  =   "Named:id=34:Heading"
      _StyleDefs(79)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(80)  =   ":id=34,.wraptext=-1"
      _StyleDefs(81)  =   "Named:id=35:Footing"
      _StyleDefs(82)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(83)  =   "Named:id=36:Selected"
      _StyleDefs(84)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(85)  =   "Named:id=37:Caption"
      _StyleDefs(86)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(87)  =   "Named:id=38:HighlightRow"
      _StyleDefs(88)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(89)  =   "Named:id=39:EvenRow"
      _StyleDefs(90)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(91)  =   "Named:id=40:OddRow"
      _StyleDefs(92)  =   ":id=40,.parent=33"
      _StyleDefs(93)  =   "Named:id=41:RecordSelector"
      _StyleDefs(94)  =   ":id=41,.parent=34"
      _StyleDefs(95)  =   "Named:id=42:FilterBar"
      _StyleDefs(96)  =   ":id=42,.parent=33"
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "&Activar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "&Desactivar"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "Activar Todos Registros"
      End
      Begin VB.Menu Menu1_5 
         Caption         =   "Desactivar Todos Registros"
      End
   End
End
Attribute VB_Name = "FrmAsientoCorregir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SeEjecuto As Boolean
Dim QueHace As Integer
Dim RstFrm As New ADODB.Recordset '--
Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta

Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE


Sub ActivaTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub


Private Sub CmdBancos_Click()

    pCorregirBancos
''    pCorregirCorrCompras
'    Dim xfrm As New Sgi2_Procesos.Procesos
'    xfrm.BDEvaluar xCon
'    Set xfrm = Nothing
    
    Err.Clear
End Sub

Private Sub Command1_Click()
    FrmRedondeoCentimos1.Show
    FrmRedondeoCentimos1.SetFocus

End Sub

Private Sub Command2_Click()
    FrmAjusteDifCambio.Show
    FrmAjusteDifCambio.SetFocus
End Sub

Private Sub Command3_Click()
    FrmAjusteDifCambioOrdDpcho.Show
    FrmAjusteDifCambioOrdDpcho.SetFocus
End Sub

Private Sub Command4_Click()
AnalisisCtaCte
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
        TxtLibro.Text = ""
        TxtMes.Text = ""
        LblIdLibro.Caption = ""
        LblIdMes.Caption = ""
        lbltotalRegistros.Caption = "Total Registros: 0"
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        BAND_INTERRUMPIR = True '--interrumpir
    End If
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    
    grilla.BatchUpdates = False
    
    grilla.Columns("fchdoc").NumberFormat = FORMAT_DATE
    grilla.Columns("total").NumberFormat = FORMAT_MONTO
    
    QueHace = 3
    
End Sub

Private Function Grabar() As Boolean
    If RstFrm.State = 0 Then Exit Function
    RstFrm.Filter = "sel=-1"
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registros seleccionados", vbExclamation, xTitulo
        RstFrm.Filter = ""
        grilla.SetFocus
        Exit Function
    End If
    
    If MsgBox("Seguro desea Corregir los Asientos de los registros seleccionados", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then
        RstFrm.Filter = ""
        Exit Function
    End If
    
    On Error GoTo LaCague
    
    
    Me.MousePointer = vbHourglass
    
    Frame2.Left = 3090
    Frame2.Top = 3210
    
    Label4.Caption = "Corrigiendo Asientos"
    
    ProgressBar2.Max = RstFrm.RecordCount
    Frame2.Visible = True
    Dim mRow&
    Dim dHora1 As Date '--hora de inicio del proceso
    Dim dHora2 As Date '--hora final de proceso
    Dim qTotalRegistros As Double
    mRow = 1
    dHora1 = Time()
    DoEvents
    qTotalRegistros = RstFrm.RecordCount
    BAND_INTERRUMPIR = False
    Do While Not RstFrm.EOF
        DoEvents
        ProgressBar2.Value = mRow
        Label2.Caption = qTotalRegistros & " / " & mRow
        DoEvents
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        '***************************************
        xCon.BeginTrans
        If GenerarAsiento(xCon, NulosN(LblIdLibro.Caption), RstFrm("id"), AnoTra, NulosN(RstFrm("idmes")), 1, RstFrm("tipmov")) = "" Then GoTo LaCague
        xCon.CommitTrans
        '***************************************
        RstFrm("sel") = 0
        mRow = mRow + 1
        
        RstFrm.MoveNext
    Loop
    
    Frame2.Visible = False
    
    Me.MousePointer = vbDefault
    
    RstFrm.Filter = ""
    dHora2 = Time()
    DoEvents
    MsgBox "Total Registros Corregidos: " & mRow - 1 & vbCr + "Tiempo Transcurrido: " & Format(CDate(CDate(dHora2) - CDate(dHora1)), FORMAT_HORA_LARGO), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    '--limpiar los filtros
    TDB_FiltroLimpiar grilla
    Exit Function
LaCague:
    RstFrm.Filter = ""
    Frame2.Visible = False
    Me.MousePointer = vbDefault
    xCon.RollbackTrans
    MsgBox "No se pudo pueden corregir los asientos por el siguiente motivo :" + Trim(Err.Description)
    Err.Clear
    Exit Function
SALIR:
    Frame2.Visible = False
    Me.MousePointer = vbDefault

    MsgBox "El proceso fue interrumpido", vbExclamation, xTitulo
    
End Function

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este Listo para Importar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    Else
        
    End If

End Sub

Private Sub grilla_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then grilla_DblClick
End Sub

Private Sub grilla_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu Menu1
End Sub

Private Sub Menu1_1_Click()
    TDB_SelDesActCheck grilla, RstFrm, "sel", "-1"
End Sub

Private Sub Menu1_2_Click()
    TDB_SelDesActCheck grilla, RstFrm, "sel", "0"
End Sub

Private Sub Menu1_4_Click()
    TDB_TodosDesActCheck grilla, RstFrm, "sel", "-1"
End Sub

Private Sub Menu1_5_Click()
    TDB_TodosDesActCheck grilla, RstFrm, "sel", "0"
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    
    If Button.Index = 2 Then Grabar
    
        
    If Button.Index = 5 Then
        If Grabar = True Then
            pConsultar
        End If
    End If
    
    If Button.Index = 6 Then
        Unload Me
    End If
End Sub

Sub pConsultar()
    Dim nSQL As String
    Dim nSQLPeriodo As String
    Dim RstTmp As New ADODB.Recordset
    On Error GoTo error
    Select Case NulosN(LblIdLibro.Caption)
    Case 1 '--compras
        If NulosN(LblIdMes.Caption) <> 0 Then nSQLPeriodo = " and (((Month([com_compras].[fchreg])) = " & NulosN(LblIdMes.Caption) & ") )"
        
        nSQL = "SELECT Month([com_compras].[fchreg]) AS idmes, com_compras.id & '' as id, 0 AS sel, IIF([com_compras].[numreg] IS NULL,'PENDIENTE',Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4)) AS registro, mae_prov.numruc AS ruc, mae_prov.nombre AS nombre, IIf([com_compras]![numser] is null or [com_compras]![numser] ='',[com_compras]![numdoc],[com_compras]![numser]+'-'+[com_compras]![numdoc]) AS numerodoc, com_compras.fchdoc & '' as fchdoc, mae_documento.abrev, mae_moneda.simbolo AS moneda, com_compras.imptot & '' AS total,0 as tipmov  " _
            + vbCr + " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (com_compras LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
            + vbCr + " WHERE com_compras.numreg <>'000001' " & nSQLPeriodo _
            + vbCr + " ORDER BY Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4); "
        
        
        grilla.Columns(3).DataField = "ruc"
        grilla.Columns(4).DataField = "nombre"
        
        grilla.Columns(3).Caption = "Nº. R.U.C."
        grilla.Columns(4).Caption = "Proveedor"

    Case 2 '--ventas
    
        If NulosN(LblIdMes.Caption) <> 0 Then nSQLPeriodo = " and (((Month([vta_ventas].[fchreg])) = " & NulosN(LblIdMes.Caption) & ") )"
        
        If chk.Value = 1 Then
'            nSQLPeriodo = nSQLPeriodo & " and ( vta_ventas.numreg is null or vta_ventas.numreg='' ) "
        End If
        
        nSQL = "SELECT Month([vta_ventas].[fchreg])  as idmes, vta_ventas.id & '' as id, 0 AS sel, iif([vta_ventas].[numreg] is null,'PENDIENTE',Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4)) AS registro, IIf([vta_ventas].[anulado]=-1,'',[mae_cliente].[numruc]) AS ruc, IIf(vta_ventas.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, IIf([vta_ventas]![numser] is null or [vta_ventas]![numser]='',[vta_ventas]![numdoc],[vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]) AS numerodoc, vta_ventas.fchdoc & '' as fchdoc, mae_documento.abrev, IIf([vta_ventas].[anulado]=-1,'',[mae_moneda].[simbolo]) AS moneda, vta_ventas.imptotdoc & '' AS total,0 as tipmov " _
        + vbCr + " FROM (mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN vta_ventas ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
        + vbCr + " Where (vta_ventas.numreg<>'000001' or vta_ventas.numreg is null) " & nSQLPeriodo _
        + vbCr + " ORDER BY vta_ventas.id "
        
        grilla.Columns(3).DataField = "ruc"
        grilla.Columns(4).DataField = "nombre"
        
        grilla.Columns(3).Caption = "Nº. R.U.C."
        grilla.Columns(4).Caption = "Cliente"
        
    Case 3 '--provisiones diversas
        
        If NulosN(LblIdMes.Caption) <> 0 Then nSQLPeriodo = " WHERE (((year([con_proviciones].[fchreg]))=" & AnoTra & ") AND (con_proviciones.idmes=" & NulosN(LblIdMes.Caption) & ")) "
        
        nSQL = "SELECT con_proviciones.idmes,con_proviciones.id & '' AS id, 0 AS sel, IIF([con_proviciones].[numreg] IS NULL,'PENDIENTE',Format([con_proviciones].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([con_proviciones].[numreg],3)) AS registro, '' AS ruc, mae_librossub.descripcion AS sublibro, [con_proviciones]![numser]+'-'+[con_proviciones]![numdoc] AS numerodoc, con_proviciones.fchdoc, mae_documento.abrev, mae_moneda.simbolo AS moneda, " _
        + vbCr + " (SELECT Sum([con_provicionesdet]![imp]) AS totdeb  FROM con_provicionesdet WHERE (((con_provicionesdet.id)=con_proviciones.id)  AND ((con_provicionesdet.tipo)=0))) AS total,  " _
        + vbCr + " (SELECT Sum([con_provicionesdet]![imp]) AS totdeb FROM con_provicionesdet WHERE  (((con_provicionesdet.id)=con_proviciones.id)   AND ((con_provicionesdet.tipo)=-1))) AS tothab, " _
        + vbCr + " mae_moneda.descripcion AS mondesc, con_proviciones.glosa,0 as tipmov " _
        + vbCr + " FROM ((((con_proviciones LEFT JOIN mae_libros ON con_proviciones.idlib = mae_libros.id) LEFT JOIN con_meses ON con_proviciones.idmes = con_meses.id) LEFT JOIN mae_moneda ON con_proviciones.idmon = mae_moneda.id) LEFT JOIN mae_documento ON con_proviciones.tipdoc = mae_documento.id) LEFT JOIN mae_librossub ON con_proviciones.idsublib = mae_librossub.id " _
        + vbCr + " " & nSQLPeriodo _
        + vbCr + " ORDER BY con_proviciones.fchreg,con_proviciones.numreg "
        
        grilla.Columns(3).DataField = "sublibro"
        grilla.Columns(4).DataField = "glosa"
        
        grilla.Columns(3).Caption = "SubLibro"
        grilla.Columns(4).Caption = "Glosa"

                        
    Case 4 '--percepciones
        
        If NulosN(LblIdMes.Caption) <> 0 Then nSQLPeriodo = " WHERE (((year(con_percepcion.fchreg))=" & AnoTra & ") AND ((Month(con_percepcion.fchreg))=" & NulosN(LblIdMes.Caption) & ")) "
    
    
        nSQL = "SELECT Month(con_percepcion.fchreg) as idmes,con_percepcion.id & '' as id, 0 AS sel, IIf(con_percepcion.tipo=1,'Compra','Venta') AS tipo, IIF(con_percepcion!numreg IS NULL,'PENDIENTE',Mid(con_percepcion!numreg,1,2)+mae_libros.codsun+Mid(con_percepcion!numreg,3,4)) AS registro, IIf(con_percepcion.tipo=1,mae_prov.numruc,mae_cliente.numruc) AS ruc, IIf(con_percepcion.tipo=1,mae_prov.nombre,mae_cliente.nombre) AS nombre, con_percepcion!numser+'-'+con_percepcion!numdoc AS numerodoc, con_percepcion.fchdoc & '' AS fchdoc, mae_documento.abrev, mae_moneda.simbolo AS moneda, con_percepcion.imptotper & '' AS total,0 as tipmov " _
        + vbCr + " FROM (mae_moneda RIGHT JOIN (((con_percepcion LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id) LEFT JOIN mae_libros ON con_percepcion.idlib = mae_libros.id) LEFT JOIN mae_documento ON con_percepcion.tipdoc = mae_documento.id) ON mae_moneda.id = con_percepcion.idmon) LEFT JOIN mae_cliente ON con_percepcion.idcli = mae_cliente.id " _
        + vbCr + " " & nSQLPeriodo
        
        grilla.Columns(3).DataField = "ruc"
        grilla.Columns(4).DataField = "nombre"
        
        grilla.Columns(3).Caption = "Nº. R.U.C."
        grilla.Columns(4).Caption = "Cliente /Proveedor /Otros"


    Case 5 '--retenciones
    
        If NulosN(LblIdMes.Caption) <> 0 Then nSQLPeriodo = " WHERE (((year([con_retencion].[fchreg]))=" & AnoTra & ") AND ((Month([con_retencion].[fchreg]))=" & NulosN(LblIdMes.Caption) & ")) "
    
    
        nSQL = "SELECT Month([con_retencion].[fchreg]) as idmes,con_retencion.id & '' as id, 0 AS sel, IIf([con_retencion].[tipo]=1,'Compra','Venta') AS tipo, IIF([con_retencion]![numreg] IS NULL,'PENDIENTE',Mid([con_retencion]![numreg],1,2)+[mae_libros].[codsun]+Mid([con_retencion]![numreg],3,4)) AS registro, IIf([con_retencion].[tipo]=1,[mae_prov].[numruc],[mae_cliente].[numruc]) AS ruc, IIf([con_retencion].[tipo]=1,[mae_prov].[nombre],[mae_cliente].[nombre]) AS nombre, [con_retencion]![numser]+'-'+[con_retencion]![numdoc] AS numerodoc, con_retencion.fchemi & '' AS fchdoc, mae_documento.abrev, mae_moneda.simbolo AS moneda, con_retencion.[imp] & '' AS total,0 as tipmov " _
        + vbCr + " FROM (mae_moneda RIGHT JOIN (((con_retencion LEFT JOIN mae_prov ON con_retencion.idpro = mae_prov.id) LEFT JOIN mae_libros ON con_retencion.idlib = mae_libros.id) LEFT JOIN mae_documento ON con_retencion.iddoc = mae_documento.id) ON mae_moneda.id = con_retencion.idmon) LEFT JOIN mae_cliente ON con_retencion.idpro = mae_cliente.id " _
        + vbCr + " " & nSQLPeriodo
        
        grilla.Columns(3).DataField = "ruc"
        grilla.Columns(4).DataField = "nombre"
        
        grilla.Columns(3).Caption = "Nº. R.U.C."
        grilla.Columns(4).Caption = "Cliente /Proveedor /Otros"
    Case 6 '--caja y bancos
        If NulosN(LblIdMes.Caption) <> 0 Then nSQLPeriodo = " WHERE (((year([tes_caja].[fchreg]))=" & AnoTra & ") AND ((Month([tes_caja].[fchreg]))=" & NulosN(LblIdMes.Caption) & ")) "
        
        If chk.Value = 1 Then
            If nSQLPeriodo <> "" Then nSQLPeriodo = nSQLPeriodo & " and ( tes_caja.numreg is null or tes_caja.numreg='' ) "
            If nSQLPeriodo = "" Then nSQLPeriodo = " WHERE ( tes_caja.numreg is null or tes_caja.numreg='' ) "
        End If

        If ChkReproceso.Value = 0 Then
            nSQL = "SELECT Month(tes_caja.fchreg) as idmes,tes_caja.id & '' as id ,  0 AS sel,mae_tipomov.descripcion AS tipo1,iif(tes_caja.numreg is null,'PENDIENTE', Mid(tes_caja.numreg,1,2) & mae_libros.codsun & Mid(tes_caja.numreg,3,4)) AS registro, IIf(IsNull(tes_cajaorigendet!numser)=-1 or tes_cajaorigendet!numser='',tes_cajaorigendet!numdoc,tes_cajaorigendet!numser & '-' & tes_cajaorigendet!numdoc) AS numerodoc, tes_caja.fchope & '' AS fchdoc, tes_documentos.abrev, tes_cajaori.importe & ''  AS total, tes_caja.glosa, mae_moneda.simbolo AS moneda, tes_documentos.descripcion AS descdoc, tes_origen.descripcion AS origen, tes_caja.tipmov " _
            + vbCr + " FROM (tes_origen RIGHT JOIN ((((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) LEFT JOIN tes_cajaori ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (tes_cajaorigendet LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) ON tes_origen.id = tes_cajaori.idori) LEFT JOIN mae_tipomov ON tes_caja.tipmov = mae_tipomov.id " _
            + vbCr + nSQLPeriodo _
            + vbCr + " GROUP BY Month(tes_caja.fchreg), tes_caja.id & '', 0, mae_tipomov.descripcion, IIf(tes_caja.numreg Is Null,'PENDIENTE',Mid(tes_caja.numreg,1,2) & mae_libros.codsun & Mid(tes_caja.numreg,3,4)), IIf(IsNull(tes_cajaorigendet!numser)=-1 Or tes_cajaorigendet!numser='',tes_cajaorigendet!numdoc,tes_cajaorigendet!numser & '-' & tes_cajaorigendet!numdoc), tes_caja.fchope & '', tes_documentos.abrev, tes_cajaori.importe & '', tes_caja.glosa, mae_moneda.simbolo, tes_documentos.descripcion, tes_origen.descripcion, tes_caja.tipmov, Mid(tes_caja.numreg,1,2) & mae_libros.codsun & Mid(tes_caja.numreg,3,4), tes_caja.fchope " _
            + vbCr + " ORDER BY Mid(tes_caja.numreg,1,2) & mae_libros.codsun & Mid(tes_caja.numreg,3,4),tes_caja.fchope ;"
        Else
            ''''''''''''consulta de reproceso
            Dim mItem As Long
            
            mItem = InputBox("Ingrese numero de Reproceso ", "Reproceso Bancos", 0)
            If IsNumeric(mItem) = False Then
                MsgBox "Ingrese correctamente el codigo de reproceso", vbInformation, xTitulo
                Exit Sub
            End If
            If mItem <= 0 Then
                MsgBox "Ingrese correctamente el codigo de reproceso", vbInformation, xTitulo
                Exit Sub
            End If
        
       
            nSQL = "SELECT Month(tes_caja.fchreg) as idmes,tes_caja.id & '' as id ,  0 AS sel,mae_tipomov.descripcion AS tipo1,iif(tes_caja.numreg is null,'PENDIENTE', Mid(tes_caja.numreg,1,2) & mae_libros.codsun & Mid(tes_caja.numreg,3,4)) AS registro, IIf(IsNull(tes_cajaorigendet!numser)=-1 or tes_cajaorigendet!numser='',tes_cajaorigendet!numdoc,tes_cajaorigendet!numser & '-' & tes_cajaorigendet!numdoc) AS numerodoc, tes_caja.fchope & '' AS fchdoc, tes_documentos.abrev, tes_cajaori.importe & ''  AS total, tes_caja.glosa, mae_moneda.simbolo AS moneda, tes_documentos.descripcion AS descdoc, tes_origen.descripcion AS origen, tes_caja.tipmov " _
            + vbCr + " FROM (((((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id) LEFT JOIN mae_tipomov ON tes_caja.tipmov = mae_tipomov.id) LEFT JOIN (tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (tes_cajaorigendet LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) INNER JOIN zzz_tes_caja ON tes_caja.id = zzz_tes_caja.id " _
            + vbCr + " WHERE ((zzz_tes_caja.idreproceso)=" & mItem & ") " _
            + vbCr + " GROUP BY Month(tes_caja.fchreg), tes_caja.id & '', 0, mae_tipomov.descripcion, IIf(tes_caja.numreg Is Null,'PENDIENTE',Mid(tes_caja.numreg,1,2) & mae_libros.codsun & Mid(tes_caja.numreg,3,4)), IIf(IsNull(tes_cajaorigendet!numser)=-1 Or tes_cajaorigendet!numser='',tes_cajaorigendet!numdoc,tes_cajaorigendet!numser & '-' & tes_cajaorigendet!numdoc), tes_caja.fchope & '', tes_documentos.abrev, tes_cajaori.importe & '', tes_caja.glosa, mae_moneda.simbolo, tes_documentos.descripcion, tes_origen.descripcion, tes_caja.tipmov, Mid(tes_caja.numreg,1,2) & mae_libros.codsun & Mid(tes_caja.numreg,3,4), tes_caja.fchope, zzz_tes_caja.idreproceso " _
            + vbCr + " ORDER BY Mid(tes_caja.numreg,1,2) & mae_libros.codsun & Mid(tes_caja.numreg,3,4),tes_caja.fchope ;"

        End If

'
        
        grilla.Columns(3).DataField = "tipo1"
        grilla.Columns(4).DataField = "origen"
        
        grilla.Columns(3).Caption = "Tipo"
        grilla.Columns(4).Caption = "Origen"
            
    Case 8 '--canjes
    
    nSQL = "SELECT 0 AS sel, con_canjes.id, con_canjes.idmes, mae_prov.numruc AS rucpro, mae_prov.nombre AS nompro, mae_cliente.numruc AS ruccli, mae_cliente.nombre AS nomcli, [con_canjes].[numser] & '-' & [con_canjes].[numdoc] AS numerodoc, mae_moneda.simbolo AS moneda, Right([con_canjes].[numreg],2) & [mae_libros].[codsun] & Left([con_canjes].[numreg],4) AS registro, con_canjes.fchemi & '' AS fchdoc, con_canjes.impcan & '' AS total ,0 as tipmov" _
         + vbCr + " FROM (((con_canjes LEFT JOIN mae_prov ON con_canjes.idpro = mae_prov.id) LEFT JOIN mae_cliente ON con_canjes.idcli = mae_cliente.id) LEFT JOIN mae_moneda ON con_canjes.idmon = mae_moneda.id) LEFT JOIN mae_libros ON con_canjes.idlib = mae_libros.id " _
         + vbCr + " ORDER BY con_canjes.idmes, Right([con_canjes].[numreg],2) & [mae_libros].[codsun] & Left([con_canjes].[numreg],4);"

    
    
        grilla.Columns(3).DataField = "rucpro"
        grilla.Columns(4).DataField = "nompro"
        
        grilla.Columns(3).Caption = "Nº. R.U.C."
        grilla.Columns(4).Caption = "Nombres"
    Case 9 '--planillas
    
        grilla.Columns(3).DataField = "dni"
        grilla.Columns(4).DataField = "nombres"
        
        grilla.Columns(3).Caption = "Nº. DNI"
        grilla.Columns(4).Caption = "Nombres"
    
        If NulosN(LblIdMes.Caption) <> 0 Then nSQLPeriodo = " WHERE (((year([pla_boleta].[fchreg]))=" & AnoTra & ") AND ((Month([pla_boleta].[fchreg]))=" & NulosN(LblIdMes.Caption) & ")) "


    nSQL = "SELECT [pla_boleta].[id] & '' AS id, pla_boleta.idmes, 0 AS sel, pla_empleados.numdoc AS dni, pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom AS nombres, pla_boleta.numser & ' ' & pla_boleta.numdoc AS numerodoc, mae_documento.abrev, mae_moneda.simbolo AS moneda, IIf([pla_boleta].[numreg] Is Null Or [pla_boleta].[numreg]='','PENDIENTE',Format([pla_boleta].[idmes],'00') & IIf([mae_libros].[codsun] Is Null Or [mae_libros].[codsun]='','FF',[mae_libros].[codsun]) & Mid([pla_boleta].[numreg],3)) AS registro, pla_boleta.fchdoc & '' AS fchdoc1, pla_boleta.fchpago & '' AS fchpago1, [pla_boleta].[imptot] & '' AS total,0 as tipmov  " _
        + vbCr + " FROM pla_empleados RIGHT JOIN (((mae_moneda RIGHT JOIN pla_boleta ON mae_moneda.id = pla_boleta.idmon) LEFT JOIN mae_documento ON pla_boleta.iddoc = mae_documento.id) LEFT JOIN mae_libros ON pla_boleta.idlib = mae_libros.id) ON pla_empleados.id = pla_boleta.idemp " _
        + vbCr + nSQLPeriodo _
        + vbCr + " ORDER BY pla_boleta.numreg "
        
        
    Case 37 '--canje de letras
    
        grilla.Columns(3).DataField = "ruc"
        grilla.Columns(4).DataField = "nombre"
        
        grilla.Columns(3).Caption = "Nº. R.U.C."
        grilla.Columns(4).Caption = "Nombres"
    
    
        If NulosN(LblIdMes.Caption) <> 0 Then nSQLPeriodo = " and let_letra.idmes= " & NulosN(LblIdMes.Caption) & " "

        nSQL = "SELECT let_letra.idmes, let_letra.ano, let_letra.id, 0 AS sel, IIf([let_letra].[numreg] Is Null,'PENDIENTE',Mid([let_letra].[numreg],1,2)+[mae_libros].[codsun]+Mid([let_letra].[numreg],3,4)) AS registro, mae_cliente.numruc AS ruc, mae_cliente.nombre, '' AS numerodoc, [let_letra].[fchemi] & '' AS fchdoc, mae_documento.abrev, mae_moneda.simbolo AS moneda, let_letra.impcap AS total, 0 as tipmov " _
        + vbCr + " FROM (((mae_moneda RIGHT JOIN (mae_cliente RIGHT JOIN let_letra ON mae_cliente.id = let_letra.idclipro) ON mae_moneda.id = let_letra.idmon) LEFT JOIN mae_documento ON let_letra.tipdoc = mae_documento.id) LEFT JOIN let_letratipoplazo ON let_letra.tipint = let_letratipoplazo.id) LEFT JOIN mae_libros ON let_letra.idlib = mae_libros.id " _
        + vbCr + " WHERE let_letra.ano=" & AnoTra & " " & nSQLPeriodo

        
    
    Case 40 '--honorarios
        nSQL = "SELECT Month([com_honorarios].[fchreg]) as idmes, com_honorarios.id & '' as id, 0 AS sel, IIF([numreg] IS NULL,'PENDIENTE',Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4)) AS registro, mae_prov.numruc AS ruc, mae_prov.nombre AS nombre, IIf([com_honorarios]![numser] is null or [com_honorarios]![numser] ='',[com_honorarios]![numdoc],[com_honorarios]![numser]+'-'+[com_honorarios]![numdoc]) AS numerodoc, com_honorarios.fchdoc & '' as fchdoc, mae_documento.abrev, mae_moneda.simbolo AS moneda, com_honorarios.imptot & '' AS total,0 as tipmov " _
         + vbCr + " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (com_honorarios LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id) ON mae_documento.id = com_honorarios.tipdoc) ON mae_moneda.id = com_honorarios.idmon) ON mae_prov.id = com_honorarios.idpro " _
         + vbCr + " Where (((Month([com_honorarios].[fchreg])) = " & NulosN(LblIdMes.Caption) & ") And ((com_honorarios.importado) = 0)) " _
         + vbCr + " ORDER BY Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4);"
        
        grilla.Columns(3).DataField = "ruc"
        grilla.Columns(4).DataField = "nombre"
        
        grilla.Columns(3).Caption = "Nº. R.U.C."
        grilla.Columns(4).Caption = "Nombres"
        
    Case 41 '--Liquidacion Gasto Debito LGD
        If NulosN(LblIdMes.Caption) <> 0 Then nSQLPeriodo = " Where (((Month([vta_gastodebito].[fchreg])) = " & NulosN(LblIdMes.Caption) & ") )"
    
        
        If chk.Value = 1 Then
            If nSQLPeriodo <> "" Then nSQLPeriodo = nSQLPeriodo & " and ( vta_gastodebito.numreg is null or vta_gastodebito.numreg='' ) "
            If nSQLPeriodo = "" Then nSQLPeriodo = " WHERE ( vta_gastodebito.numreg is null or vta_gastodebito.numreg='' ) "
        End If
        
        nSQL = "SELECT vta_gastodebito.idmes as idmes, vta_gastodebito.id & '' AS id, 0 AS sel, IIF([numreg] IS NULL,'PENDIENTE', Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4)) AS registro, IIF(vta_gastodebito.anulado=-1,'',mae_cliente.numruc) as ruc, IIF(vta_gastodebito.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, [vta_gastodebito].[numser] & '-' & [vta_gastodebito].[numdoc] AS numerodoc, vta_gastodebito.fchemi & '' AS fchdoc, mae_documento.abrev, mae_moneda.simbolo AS moneda, vta_gastodebito.imptot & '' AS total,0 as tipmov " _
         + vbCr + "  FROM (((vta_gastodebito LEFT JOIN mae_cliente ON vta_gastodebito.idcli = mae_cliente.id) LEFT JOIN mae_libros ON vta_gastodebito.idlib = mae_libros.id) LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_gastodebito.idmon = mae_moneda.id " _
         + vbCr + nSQLPeriodo _
         + vbCr + " ORDER BY [vta_gastodebito].[numser] & '-' & [vta_gastodebito].[numdoc]; "

        grilla.Columns(3).DataField = "ruc"
        grilla.Columns(4).DataField = "nombre"
        
        grilla.Columns(3).Caption = "Nº. R.U.C."
        grilla.Columns(4).Caption = "Nombres"
        
    Case 42 'Planilla de Letras Presentadas al Banco
        
        If NulosN(LblIdMes.Caption) <> 0 Then nSQLPeriodo = " Where (((Month([let_planilla].[fchreg])) = " & NulosN(LblIdMes.Caption) & ") )"
    
        
        If chk.Value = 1 Then
            If nSQLPeriodo <> "" Then nSQLPeriodo = nSQLPeriodo & " and ( let_planilla.numreg is null or let_planilla.numreg='' ) "
            If nSQLPeriodo = "" Then nSQLPeriodo = " WHERE ( let_planilla.numreg is null or let_planilla.numreg='' ) "
        End If
                
        
        nSQL = "SELECT Month([let_planilla].[fchreg]) as idmes,let_planilla.id & '' AS id,0 AS sel, let_planilla.numdoc as numerodoc, let_planilla.numlet, let_planilla.imptot & '' as total, let_planilla.fchemi & '' as fchdoc, mae_bancos.descripcion AS banco, let_modalidad.descripcion AS modalidad, " _
            & " let_planilla.numreg,mae_moneda.simbolo AS moneda, let_planilla.idmon,trim(mae_bancos.descripcion ) & ' Nro Cta. ' & mae_banconumcta.numcue as bancocuenta, " _
            & " IIF([let_planilla]![numreg] IS NULL,'PENDIENTE',Mid([let_planilla]![numreg],1,2) & [mae_libros]![codsun] & Mid([let_planilla]![numreg],3,4)) AS registro,0 as tipmov,mae_documento.abrev  " _
            & " FROM ((mae_bancos RIGHT JOIN (((let_planilla LEFT JOIN let_modalidad ON let_planilla.idmod = let_modalidad.id) LEFT JOIN mae_moneda ON let_planilla.idmon = mae_moneda.id)  " _
            & " LEFT JOIN mae_banconumcta ON let_planilla.idbcocta = mae_banconumcta.id) ON mae_bancos.id = mae_banconumcta.idban) LEFT JOIN mae_libros ON let_planilla.idlib = mae_libros.id) LEFT JOIN mae_documento ON let_planilla.tipdoc = mae_documento.id; "
        
        
        grilla.Columns(3).DataField = "modalidad"
        grilla.Columns(4).DataField = "bancocuenta"
        
        grilla.Columns(3).Caption = "Modalidad"
        grilla.Columns(4).Caption = "Banco - Num Cuenta"
        
        
        
    Case Else
        MsgBox "Pendiente", vbInformation
        Exit Sub
    End Select
    '--limpiar los filtros
    TDB_FiltroLimpiar grilla
    '
    Set RstFrm = Nothing
    Set grilla.DataSource = Nothing
    Me.MousePointer = vbHourglass
    DoEvents
    'RST_Busq RstFrm, nSQL, xCon
    'RST_Busq RstTmp, nSQL, xCon
    
    Set RstTmp = xCon.Execute(nSQL)
    
    
    If RstTmp.State = 1 Then
        DEFINIR_RST_TMP RstFrm, RstTmp
        CARGAR_RST_TMP RstFrm, RstTmp
        
        Set grilla.DataSource = RstFrm
        
        lbltotalRegistros.Caption = "Total Registros: " & RstFrm.RecordCount
    End If
    
    Me.MousePointer = vbDefault
    
    Set RstTmp = Nothing
    
    
    Exit Sub
error:
    Me.MousePointer = vbDefault
    MsgBox Err.Description, vbCritical, xTitulo
    Err.Clear
End Sub


Private Sub pCrearFormato()

    On Error GoTo error
    Dim objExcel As Object
    Dim k As Integer
    
    Set objExcel = CreateObject("Excel.Application")
    objExcel.SheetsInNewWorkbook = 1
    objExcel.WindowState = 1
    objExcel.Workbooks.Add
    
    With objExcel.ActiveSheet
        .Cells(1, 1) = "Importar Clientes"
        
        .Cells(3, 1) = "Tipo de Persona"
        .Cells(3, 2) = "Tipo Documento"
        .Cells(3, 3) = "Nº Documento"
        .Cells(3, 4) = "Cliente"
        .Cells(3, 5) = "Nombre 1"
        .Cells(3, 6) = "Nombre 2"
        .Cells(3, 7) = "Apellido 1"
        .Cells(3, 8) = "Apellido 2"
        .Cells(3, 9) = "Dirección"
        .Cells(3, 10) = "Departamento"
        .Cells(3, 11) = "Distrito"
        .Cells(3, 12) = "Teléfono"
        .Cells(3, 13) = "Fax"
        .Cells(3, 14) = "email"
        '---------
        .Columns(1).ColumnWidth = 8
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 15
        '--establecer el ancho
        
        .Columns(1).ColumnWidth = 15
        .Columns(2).ColumnWidth = 15
        .Columns(3).ColumnWidth = 12
        .Columns(4).ColumnWidth = 25.5
        
        .Columns(9).ColumnWidth = 31
        
        For k = 1 To 14
            .Cells(3, k).Font.Bold = True
        Next

                
    End With
    MsgBox "Proceda a ingresar la información según los Parámetros Solicitados" + vbCr + "Luego proceda a Importar...", vbInformation, xTitulo
    objExcel.Visible = True
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
error:
    Set objExcel = Nothing
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then pCrearFormato
End Sub

'**********************
Private Sub CmdBusProv_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "5500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_libros  where activo = -1 ORDER BY descripcion "
    
    xform.Titulo = "Buscando Libro Contable"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        TxtLibro.Text = ""
        TxtLibro.Text = NulosC(xRs("descripcion"))
        LblIdLibro.Caption = NulosC(xRs("id"))
        TxtMes.SetFocus
    End If
    If NulosN(LblIdLibro.Caption) = 6 Then CmdBancos.Visible = True Else CmdBancos.Visible = False
    If NulosN(LblIdLibro.Caption) = 6 Then ChkReproceso.Visible = True Else ChkReproceso.Visible = False
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub CmdBusMes_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    xCampos2(0, 0) = "Descripción":    xCampos2(0, 1) = "descripcion":    xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Código":       xCampos2(1, 1) = "id":             xCampos2(1, 2) = "1000":         xCampos2(1, 3) = "N"

    xform.SQLCad = "SELECT * FROM con_meses"
    xform.Titulo = "Buscando Mes de Trabajo"
        
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtMes.Text = ""
        TxtMes.Text = xRs("descripcion")
        LblIdMes.Caption = xRs("id")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


'*******************************************************
Private Sub grilla_FilterChange()
    TDB_FiltroGenerar grilla, RstFrm
End Sub

Private Sub grilla_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    TDB_Ordenar grilla, ColIndex, RstFrm, AscendenteGrid, True
    
    Err.Clear

End Sub


Private Sub grilla_DblClick()
On Error Resume Next
    If RstFrm.State = 0 Then Exit Sub
    If RstFrm.RecordCount < 1 Then Exit Sub
    RstFrm.Fields("sel") = Not RstFrm.Fields("sel")
Err.Clear
End Sub

Private Sub TxtLibro_Change()
    If NulosN(TxtLibro.Text) = 0 Then
        Set RstFrm = Nothing
        Set grilla.DataSource = Nothing
        DoEvents
    End If
End Sub

Private Sub TxtLibro_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then CmdBusProv_Click
End Sub

Private Sub TxtMes_Change()
    If NulosN(TxtMes.Text) = 0 Then
        Set RstFrm = Nothing
        Set grilla.DataSource = Nothing
        DoEvents
    End If
End Sub

Private Sub TxtMes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then CmdBusMes_Click
End Sub


Private Sub pCorregirBancos()
    Dim rst As New ADODB.Recordset
    Dim Rstdet As New ADODB.Recordset
    Dim nSQL As String
    Dim mCorr As Long
    Dim mIdTes As Long
    
    On Error GoTo error

    If NulosN(LblIdLibro) <> 6 Then Exit Sub
    
    
    nSQL = "SELECT tes_caja.id, tes_caja.numreg FROM tes_caja INNER JOIN tes_cajadestinodet ON tes_caja.id = tes_cajadestinodet.idtes GROUP BY tes_caja.id, tes_caja.numreg, tes_cajadestinodet.corr HAVING (((tes_cajadestinodet.corr)=0 Or (tes_cajadestinodet.corr) Is Null)) " _
    + vbCr + " union " _
    + vbCr + "  SELECT tes_caja.id, tes_caja.numreg FROM tes_caja INNER JOIN tes_cajaorigendet ON tes_caja.id = tes_cajaorigendet.idtes GROUP BY tes_caja.id, tes_caja.numreg, tes_cajaorigendet.corr HAVING (((tes_cajaorigendet.corr)=0 Or (tes_cajaorigendet.corr) Is Null)) "
    
    RST_Busq rst, nSQL, xCon
    
    If rst.RecordCount <> 0 Then
        Me.MousePointer = vbHourglass
        rst.MoveFirst
        
        Do While Not rst.EOF
            DoEvents
            '--corrigiendo el origen
            mCorr = 1
            mIdTes = rst("id")
            Set Rstdet = Nothing
            '--reiniciando los numeros
            xCon.Execute "UPDATE tes_cajaorigendet SET tes_cajaorigendet.corr = 0 WHERE (((tes_cajaorigendet.idtes)=" & rst("id") & "));"
            
            nSQL = "SELECT tes_cajaorigendet.* FROM tes_cajaorigendet WHERE (((tes_cajaorigendet.idtes)=" & mIdTes & ")); "
            RST_Busq Rstdet, nSQL, xCon
            If Rstdet.RecordCount <> 0 Then
                Rstdet.MoveFirst
                Do While Not Rstdet.EOF
                    Rstdet("corr") = mCorr
                    Rstdet.Update
                    mCorr = mCorr + 1
                    Rstdet.MoveNext
                Loop
            End If
            '--corrigiendo los destinos
            mCorr = 1
            Set Rstdet = Nothing
            '--reiniciando los numeros
            xCon.Execute "UPDATE tes_cajadestinodet SET tes_cajadestinodet.corr = 0 WHERE (((tes_cajadestinodet.idtes)=" & rst("id") & "));"
                        
            nSQL = "SELECT tes_cajadestinodet.* FROM tes_cajadestinodet WHERE (((tes_cajadestinodet.idtes)=" & mIdTes & ")); "
            RST_Busq Rstdet, nSQL, xCon
            If Rstdet.RecordCount <> 0 Then
                Rstdet.MoveFirst
                Do While Not Rstdet.EOF
                    Rstdet("corr") = mCorr
                    Rstdet.Update
                    mCorr = mCorr + 1
                    Rstdet.MoveNext
                Loop
            End If
            
            
            rst.MoveNext
        Loop
        Me.MousePointer = vbDefault
        MsgBox "Proceso Terminado", vbInformation, xTitulo
    Else
        MsgBox "No hay Registros de Bancos para Corregir", vbInformation, xTitulo
    End If
    
    Set Rstdet = Nothing
    Set rst = Nothing
    Exit Sub
error:
    Me.MousePointer = vbDefault
    MsgBox Err.Description & vbCr & Err.Source & "Id =>> " & mIdTes, vbCritical, xTitulo
    Err.Clear
End Sub



Private Sub pActualizarDatosDiario()
    Exit Sub
    Dim nSQL As String
    '--compras
    nSQL = "UPDATE (com_compras INNER JOIN con_diario ON com_compras.id = con_diario.idmov) LEFT JOIN mae_libros ON com_compras.idlib = mae_libros.id SET con_diario.ridlib = 1, con_diario.ridtipper = 1, con_diario.ridper = [com_compras].[idpro], con_diario.rtipdoc = [com_compras].[tipdoc], con_diario.rfchope = [com_compras].[fchdoc], con_diario.rnumerodoc = IIf([com_compras].[numser] Is Null Or [com_compras].[numser]='','',[com_compras].[numser] & '-') & [com_compras].[numdoc],con_diario.rglosa=null, con_diario.rglosaope = [com_compras].[glosa], con_diario.rregistro = Left([com_compras].[numreg],2) & [mae_libros].[codsun] & Right([com_compras].[numreg],4), con_diario.ridmon = [com_compras].[idmon],con_diario.idmon = [com_compras].[idmon] " _
        + vbCr + " WHERE (((con_diario.rfchope) Is Null) AND ((con_diario.idlib)=1));"
    xCon.Execute nSQL
    
    '--ventas
    nSQL = "UPDATE (vta_ventas INNER JOIN con_diario ON vta_ventas.id = con_diario.idmov) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id SET con_diario.ridlib = 2, con_diario.ridtipper = 2, con_diario.ridper = [vta_ventas].[idcli], con_diario.rtipdoc = [vta_ventas].[tipdoc], con_diario.rfchope = [vta_ventas].[fchdoc], con_diario.rnumerodoc = IIf([vta_ventas].[numser] Is Null Or [vta_ventas].[numser]='','',[vta_ventas].[numser] & '-') & [vta_ventas].[numdoc], con_diario.rglosa = [vta_ventas].[glosa], con_diario.rregistro = Left([vta_ventas].[numreg],2) & [mae_libros].[codsun] & Right([vta_ventas].[numreg],4) " _
        + vbCr + " WHERE (((con_diario.rfchope) Is Null) AND ((con_diario.idlib)=2)); "

    '--retenciones
    nSQL = "UPDATE ((con_diario LEFT JOIN vta_ventas ON con_diario.iddocpro = vta_ventas.id) RIGHT JOIN (con_retencion LEFT JOIN mae_libros ON con_retencion.idlib = mae_libros.id) ON con_diario.idmov = con_retencion.id) LEFT JOIN mae_libros AS mae_libros_2 ON vta_ventas.idlib = mae_libros_2.id SET con_diario.ridlib = IIf([con_diario].[iddocpro]=0,[mae_libros].[id],[mae_libros_2].[id]), con_diario.ridtipper = 2, con_diario.ridper = IIf([con_diario].[iddocpro]=0,[con_retencion].[idpro],[vta_ventas].[idcli]), con_diario.rtipdoc = IIf([con_diario].[iddocpro]=0,[con_retencion].[iddoc],[vta_ventas].[tipdoc]), " _
        + vbCr + " con_diario.rfchope = IIf([con_diario].[iddocpro]=0,[con_retencion].[fchemi],[vta_ventas].[fchdoc]), con_diario.rnumerodoc = IIf([con_diario].[iddocpro]=0,IIf([con_retencion]![numser] Is Null Or [con_retencion]![numser]='','',[con_retencion]![numser] & '-') & [con_retencion]![numdoc],IIf([vta_ventas]![numser] Is Null Or [vta_ventas]![numser]='','',[vta_ventas]![numser] & '-') & [vta_ventas]![numdoc]), con_diario.rregistro = IIf([con_diario].[iddocpro]=0,Left([con_retencion].[numreg],2) & Format([mae_libros].[codsun],'00') & Right([con_retencion].[numreg],4),Left([vta_ventas].[numreg],2) & Format([mae_libros_2].[codsun],'00') & Right([vta_ventas].[numreg],4)), con_diario.rglosa = IIf([con_diario].[iddocpro]=0,[con_retencion].[glosa],[vta_ventas].[glosa]) " _
        + vbCr + " WHERE (((con_retencion.tipo)=2) AND ((con_diario.idlib)=5));"
    xCon.Execute nSQL
    
    '--honorarios
    nSQL = "UPDATE (com_honorarios INNER JOIN con_diario ON com_honorarios.id = con_diario.idmov) LEFT JOIN mae_libros ON com_honorarios.idlib = mae_libros.id SET con_diario.ridlib = 40, con_diario.ridtipper = 1, con_diario.ridper = [com_honorarios].[idpro], con_diario.rtipdoc = [com_honorarios].[tipdoc], con_diario.rfchope = [com_honorarios].[fchdoc], con_diario.rnumerodoc = IIf([com_honorarios].[numser] Is Null Or [com_honorarios].[numser]='','',[com_honorarios].[numser] & '-') & [com_honorarios].[numdoc], con_diario.rglosa = [com_honorarios].[glosa], con_diario.rregistro = Left([com_honorarios].[numreg],2) & [mae_libros].[codsun] & Right([com_honorarios].[numreg],4) " _
        + vbCr + " WHERE (((con_diario.rfchope) Is Null) AND ((con_diario.idlib)=40));"
    xCon.Execute nSQL

    
    
End Sub


Private Sub pEliminarRegTmp()
    '210109
    Exit Sub
    Dim rst As New ADODB.Recordset
    Dim nSQL As String
    
    nSQL = "select * from zzz_zzz_EliminarCorr"
    
    RST_Busq rst, nSQL, xCon
    
    Do While Not rst.EOF
        xCon.Execute "delete from zzz_regularizados where correlativo=" & rst("correlativo")
        rst.MoveNext
    Loop

End Sub


Private Sub pCorregirCorrCompras()
    Exit Sub
    Dim rst As New ADODB.Recordset
    Dim Rstdet As New ADODB.Recordset
    Dim mCorrelativo&
    Dim nSQL As String
    nSQL = "select * from com_compras where com_compras.numreg not in ('000001','0001') "
    RST_Busq rst, nSQL, xCon
    If rst.RecordCount <> 0 Then
        Do While Not rst.EOF
            Set Rstdet = Nothing
            xCon.Execute "update com_comprasdet set corr=0 where idcom =" & rst("id")
            RST_Busq Rstdet, "select * from com_comprasdet where idcom =" & rst("id"), xCon
            
            mCorrelativo = 1
            If Rstdet.RecordCount <> 0 Then
                Rstdet.MoveFirst
                Do While Not Rstdet.EOF
                    Rstdet("corr") = mCorrelativo
'                    rstDet.Update
                    mCorrelativo = mCorrelativo + 1
                    Rstdet.MoveNext
                Loop
            End If
            rst.MoveNext
        Loop
    End If
    
    Set Rstdet = Nothing
    Set rst = Nothing


End Sub



Private Sub AnalisisCtaCte()
    '===================================================================================================
    'Creado : 16/12/10 Por: Johan Castro
    'Propósito: Grabar registro para analisis de cta cte
    '
    'Entradas:  Ninguno
    '
    'Resultados: Registro en tabla var_analisisctacte
    '
    'Nota:       1.- Consultar el registro
    '            2.- Graba registro(Solo Ventas, LGD, Letras, Abonos
    
    '===================================================================================================
    

    If RstFrm.State = 0 Then Exit Sub
    RstFrm.Filter = "sel=-1"
    If RstFrm.RecordCount = 0 Then
        MsgBox "No hay registros seleccionados", vbExclamation, xTitulo
        RstFrm.Filter = ""
        grilla.SetFocus
        Exit Sub
    End If
    
    If MsgBox("Seguro desea Agregar / Modificar el registro del Análisis de Cuenta Corriente", vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then
        RstFrm.Filter = ""
        Exit Sub
    End If
    
    '--verificando si el libro genera analisis de cta cte
    Select Case NulosN(LblIdLibro.Caption)
        Case 2, 37, 41
        
        Case 6
            MsgBox "Pendiente", vbInformation, xTitulo
        Case Else
            MsgBox "El libro " & TxtLibro.Text & " no genera analisis de cuenta corriente", vbInformation, xTitulo
            Exit Sub
    End Select
    
    On Error GoTo LaCague
    
    
    Me.MousePointer = vbHourglass
    
    Frame2.Left = 3090
    Frame2.Top = 3210
    
    Label4.Caption = "Registrando en ánalisis de cta cte"
    
    ProgressBar2.Max = RstFrm.RecordCount
    Frame2.Visible = True
    Dim mRow&
    Dim dHora1 As Date '--hora de inicio del proceso
    Dim dHora2 As Date '--hora final de proceso
    Dim qTotalRegistros As Double
    mRow = 1
    dHora1 = Time()
    DoEvents
    qTotalRegistros = RstFrm.RecordCount
    BAND_INTERRUMPIR = False
    Do While Not RstFrm.EOF
        DoEvents
        ProgressBar2.Value = mRow
        Label2.Caption = qTotalRegistros & " / " & mRow
        DoEvents
        If BAND_INTERRUMPIR = True Then GoTo SALIR:
        '***************************************
        xCon.BeginTrans
         
        GrabarOperacionCtaCte NulosN(LblIdLibro.Caption), RstFrm("id"), xCon
        
        xCon.CommitTrans
        '***************************************
        RstFrm("sel") = 0
        mRow = mRow + 1
        
        RstFrm.MoveNext
    Loop
    
    Frame2.Visible = False
    
    Me.MousePointer = vbDefault
    
    RstFrm.Filter = ""
    dHora2 = Time()
    DoEvents
    MsgBox "Total Registros en ánalisis de cta cte: " & mRow - 1 & vbCr + "Tiempo Transcurrido: " & Format(CDate(CDate(dHora2) - CDate(dHora1)), FORMAT_HORA_LARGO), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    '--limpiar los filtros
    TDB_FiltroLimpiar grilla
    Exit Sub
LaCague:
    RstFrm.Filter = ""
    Frame2.Visible = False
    Me.MousePointer = vbDefault
    xCon.RollbackTrans
    MsgBox "No se pudo pueden corregir los asientos por el siguiente motivo :" + Trim(Err.Description)
    Err.Clear
    Exit Sub
SALIR:
    Frame2.Visible = False
    Me.MousePointer = vbDefault

    MsgBox "El proceso fue interrumpido", vbExclamation, xTitulo
    
End Sub

