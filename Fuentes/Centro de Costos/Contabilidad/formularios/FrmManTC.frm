VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Begin VB.Form FrmManTC 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Contabilidad - Tipo de Cambio"
   ClientHeight    =   5895
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5895
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog cmm 
      Left            =   6990
      Top             =   315
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FraProgreso 
      BorderStyle     =   0  'None
      Height          =   780
      Left            =   555
      TabIndex        =   35
      Top             =   1455
      Visible         =   0   'False
      Width           =   5085
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   90
         TabIndex        =   36
         Top             =   390
         Width           =   4890
         _ExtentX        =   8625
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
         Scrolling       =   1
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   0
         X2              =   6735
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   15
         X2              =   15
         Y1              =   -15
         Y2              =   1305
      End
      Begin VB.Label LblTituloProg 
         AutoSize        =   -1  'True
         Caption         =   "Exportando a MSExcel"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Left            =   120
         TabIndex        =   38
         Top             =   90
         Width           =   1860
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Interrumpir = ESC"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   2
         Left            =   6000
         TabIndex        =   37
         Top             =   90
         Width           =   1530
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   5070
         X2              =   5070
         Y1              =   0
         Y2              =   1095
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   0
         X2              =   5100
         Y1              =   765
         Y2              =   765
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   -15
      Top             =   2130
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTC.frx":0000
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTC.frx":0544
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTC.frx":06C8
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTC.frx":0B1C
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTC.frx":0C34
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTC.frx":1178
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTC.frx":16BC
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTC.frx":17D0
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTC.frx":18E4
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTC.frx":1D38
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTC.frx":1EA4
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTC.frx":23EC
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTC.frx":2780
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManTC.frx":28D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraImportar 
      BorderStyle     =   0  'None
      Caption         =   "Frame4"
      Height          =   5265
      Left            =   5895
      TabIndex        =   20
      Top             =   495
      Visible         =   0   'False
      Width           =   4140
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   3855
         Picture         =   "FrmManTC.frx":2A6A
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   39
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VB.TextBox txtAnio1 
         Height          =   300
         Left            =   810
         TabIndex        =   8
         Text            =   "txtAnio1"
         Top             =   405
         Width           =   675
      End
      Begin VB.CommandButton CmdImportar 
         Caption         =   "&Cancelar"
         Height          =   465
         Index           =   2
         Left            =   2790
         TabIndex        =   23
         Top             =   4710
         Width           =   1260
      End
      Begin VB.CommandButton CmdImportar 
         Caption         =   "&Grabar"
         Height          =   465
         Index           =   1
         Left            =   1500
         TabIndex        =   22
         Top             =   4710
         Width           =   1260
      End
      Begin VB.CommandButton CmdImportar 
         Caption         =   "&Archivo..."
         Height          =   465
         Index           =   0
         Left            =   195
         TabIndex        =   21
         Top             =   4710
         Width           =   1260
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg1 
         Height          =   3420
         Left            =   150
         TabIndex        =   11
         Top             =   1140
         Width           =   3900
         _cx             =   6879
         _cy             =   6032
         _ConvInfo       =   1
         Appearance      =   2
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   128
         ForeColorSel    =   16777215
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmManTC.frx":2D56
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   -1  'True
         SubtotalPosition=   1
         OutlineBar      =   0
         OutlineCol      =   0
         Ellipsis        =   0
         ExplorerBar     =   0
         PicturesOver    =   0   'False
         FillStyle       =   0
         RightToLeft     =   0   'False
         PictureType     =   0
         TabBehavior     =   0
         OwnerDraw       =   0
         Editable        =   0
         ShowComboButton =   1
         WordWrap        =   0   'False
         TextStyle       =   0
         TextStyleFixed  =   0
         OleDragMode     =   0
         OleDropMode     =   0
         DataMode        =   0
         VirtualData     =   -1  'True
         DataMember      =   ""
         ComboSearch     =   3
         AutoSizeMouse   =   -1  'True
         FrozenRows      =   0
         FrozenCols      =   0
         AllowUserFreezing=   0
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   300
         Left            =   1515
         TabIndex        =   30
         Top             =   390
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1997
         BuddyControl    =   "ImageList1"
         BuddyDispid     =   196646
         OrigLeft        =   4125
         OrigTop         =   600
         OrigRight       =   4365
         OrigBottom      =   975
         Max             =   2999
         Min             =   1997
         Enabled         =   -1  'True
      End
      Begin MSDataListLib.DataCombo DtcTMoneda1 
         Height          =   315
         Left            =   810
         TabIndex        =   10
         Top             =   750
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "DtcTMoneda1"
      End
      Begin MSDataListLib.DataCombo DtcMesFiltro1 
         Height          =   315
         Left            =   2235
         TabIndex        =   9
         Top             =   405
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "DtcMesFiltro1"
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año"
         Height          =   195
         Left            =   165
         TabIndex        =   33
         Top             =   495
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes"
         Height          =   195
         Index           =   1
         Left            =   1890
         TabIndex        =   32
         Top             =   495
         Width           =   300
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda"
         Height          =   195
         Left            =   165
         TabIndex        =   31
         Top             =   855
         Width           =   585
      End
      Begin VB.Label lbl_titulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Importando Datos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Index           =   0
         Left            =   75
         TabIndex        =   24
         Top             =   60
         Width           =   1515
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         Index           =   1
         X1              =   210
         X2              =   3900
         Y1              =   4605
         Y2              =   4605
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   0
         X2              =   15
         Y1              =   0
         Y2              =   3195
      End
      Begin VB.Line ln 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   2
         X1              =   4110
         X2              =   4095
         Y1              =   -285
         Y2              =   5500
      End
      Begin VB.Line ln 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   -210
         X2              =   5475
         Y1              =   5250
         Y2              =   5250
      End
      Begin VB.Line ln 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   -15
         X2              =   6360
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Index           =   0
         Left            =   45
         Top             =   15
         Width           =   4035
      End
   End
   Begin VB.Frame Frame1 
      Height          =   570
      Left            =   -30
      TabIndex        =   25
      Top             =   345
      Width           =   5880
      Begin VB.TextBox TxtAnio 
         Height          =   300
         Left            =   2670
         TabIndex        =   1
         Text            =   "TxtAnio"
         Top             =   180
         Width           =   675
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   300
         Left            =   3360
         TabIndex        =   26
         Top             =   180
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1997
         BuddyControl    =   "ImageList1"
         BuddyDispid     =   196646
         OrigLeft        =   4125
         OrigTop         =   600
         OrigRight       =   4365
         OrigBottom      =   975
         Max             =   2999
         Min             =   1997
         Enabled         =   -1  'True
      End
      Begin MSDataListLib.DataCombo DtcTMoneda_Filtro 
         Height          =   315
         Left            =   750
         TabIndex        =   0
         Top             =   165
         Width           =   1545
         _ExtentX        =   2725
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin MSDataListLib.DataCombo DtcMesFiltro 
         Height          =   315
         Left            =   4035
         TabIndex        =   2
         Top             =   165
         Width           =   1785
         _ExtentX        =   3149
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Año"
         Height          =   195
         Left            =   2310
         TabIndex        =   29
         Top             =   300
         Width           =   285
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Mes"
         Height          =   195
         Index           =   0
         Left            =   3660
         TabIndex        =   28
         Top             =   300
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Moneda"
         Height          =   195
         Left            =   60
         TabIndex        =   27
         Top             =   300
         Width           =   585
      End
   End
   Begin VB.Frame FraRegistrar 
      BorderStyle     =   0  'None
      Height          =   2460
      Left            =   1125
      TabIndex        =   14
      Top             =   1665
      Visible         =   0   'False
      Width           =   3735
      Begin VB.CommandButton CmdCancelar 
         Caption         =   "&Cancelar"
         Height          =   420
         Left            =   1905
         TabIndex        =   13
         Top             =   1920
         Width           =   1020
      End
      Begin VB.CommandButton CmdAceptar 
         Caption         =   "&Aceptar"
         Height          =   420
         Left            =   810
         TabIndex        =   12
         Top             =   1920
         Width           =   1020
      End
      Begin MSDataListLib.DataCombo DtcTMoneda 
         Height          =   315
         Left            =   2160
         TabIndex        =   4
         Top             =   480
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Style           =   2
         Text            =   "DataCombo1"
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFechaRegistro 
         Height          =   300
         Left            =   2160
         TabIndex        =   5
         Top             =   810
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   529
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Valor           =   "10/09/2007"
      End
      Begin VB.TextBox TxtTC_Venta 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2160
         TabIndex        =   7
         Top             =   1440
         Width           =   1335
      End
      Begin VB.TextBox TxtTC_Compra 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   2160
         TabIndex        =   6
         Top             =   1125
         Width           =   1335
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         Index           =   2
         X1              =   105
         X2              =   3605
         Y1              =   1830
         Y2              =   1830
      End
      Begin VB.Label LblTituloFrame 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   " Nuevo Tipo de Cambio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   195
         Left            =   90
         TabIndex        =   19
         Top             =   90
         Width           =   2010
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   300
         Index           =   1
         Left            =   45
         Top             =   45
         Width           =   3660
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   3720
         X2              =   3720
         Y1              =   15
         Y2              =   2460
      End
      Begin VB.Line Line4 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -330
         X2              =   3715
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         X1              =   0
         X2              =   4045
         Y1              =   2445
         Y2              =   2445
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   0
         X1              =   0
         X2              =   0
         Y1              =   15
         Y2              =   2460
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Moneda"
         Height          =   195
         Left            =   195
         TabIndex        =   18
         Top             =   510
         Width           =   1170
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Fecha de Registro"
         Height          =   195
         Left            =   195
         TabIndex        =   17
         Top             =   825
         Width           =   1305
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cambio Venta"
         Height          =   195
         Left            =   195
         TabIndex        =   16
         Top             =   1485
         Width           =   1575
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Tipo de Cambio Compra"
         Height          =   195
         Left            =   195
         TabIndex        =   15
         Top             =   1170
         Width           =   1695
      End
   End
   Begin TrueOleDBGrid70.TDBGrid Dg1 
      Height          =   4935
      Left            =   0
      TabIndex        =   3
      Top             =   945
      Width           =   5835
      _ExtentX        =   10292
      _ExtentY        =   8705
      _LayoutType     =   4
      _RowHeight      =   -2147483647
      _WasPersistedAsPixels=   0
      Columns(0)._VlistStyle=   0
      Columns(0)._MaxComboItems=   5
      Columns(0).Caption=   "ID"
      Columns(0).DataField=   "id"
      Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(1)._VlistStyle=   0
      Columns(1)._MaxComboItems=   5
      Columns(1).Caption=   "Fecha"
      Columns(1).DataField=   "fecha"
      Columns(1).NumberFormat=   "dd/mm/yyyy"
      Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(2)._VlistStyle=   0
      Columns(2)._MaxComboItems=   5
      Columns(2).Caption=   "Moneda"
      Columns(2).DataField=   "descripcion"
      Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(3)._VlistStyle=   0
      Columns(3)._MaxComboItems=   5
      Columns(3).Caption=   "Importe Compra"
      Columns(3).DataField=   "impcom"
      Columns(3).NumberFormat=   "0.0000"
      Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns(4)._VlistStyle=   0
      Columns(4)._MaxComboItems=   5
      Columns(4).Caption=   "Importe Venta"
      Columns(4).DataField=   "impven"
      Columns(4).NumberFormat=   "0.0000"
      Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
      Columns.Count   =   5
      Splits(0)._UserFlags=   0
      Splits(0).RecordSelectorWidth=   265
      Splits(0)._SavedRecordSelectors=   0   'False
      Splits(0).DividerColor=   12632256
      Splits(0).SpringMode=   0   'False
      Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
      Splits(0)._ColumnProps(0)=   "Columns.Count=5"
      Splits(0)._ColumnProps(1)=   "Column(0).Width=1931"
      Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
      Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=1852"
      Splits(0)._ColumnProps(4)=   "Column(0)._ColStyle=516"
      Splits(0)._ColumnProps(5)=   "Column(0).Visible=0"
      Splits(0)._ColumnProps(6)=   "Column(0).Order=1"
      Splits(0)._ColumnProps(7)=   "Column(1).Width=1746"
      Splits(0)._ColumnProps(8)=   "Column(1).DividerColor=0"
      Splits(0)._ColumnProps(9)=   "Column(1)._WidthInPix=1667"
      Splits(0)._ColumnProps(10)=   "Column(1)._ColStyle=516"
      Splits(0)._ColumnProps(11)=   "Column(1).Order=2"
      Splits(0)._ColumnProps(12)=   "Column(2).Width=1905"
      Splits(0)._ColumnProps(13)=   "Column(2).DividerColor=0"
      Splits(0)._ColumnProps(14)=   "Column(2)._WidthInPix=1826"
      Splits(0)._ColumnProps(15)=   "Column(2)._ColStyle=512"
      Splits(0)._ColumnProps(16)=   "Column(2).Order=3"
      Splits(0)._ColumnProps(17)=   "Column(3).Width=2540"
      Splits(0)._ColumnProps(18)=   "Column(3).DividerColor=0"
      Splits(0)._ColumnProps(19)=   "Column(3)._WidthInPix=2461"
      Splits(0)._ColumnProps(20)=   "Column(3)._ColStyle=514"
      Splits(0)._ColumnProps(21)=   "Column(3).Order=4"
      Splits(0)._ColumnProps(22)=   "Column(4).Width=2408"
      Splits(0)._ColumnProps(23)=   "Column(4).DividerColor=0"
      Splits(0)._ColumnProps(24)=   "Column(4)._WidthInPix=2328"
      Splits(0)._ColumnProps(25)=   "Column(4)._ColStyle=514"
      Splits(0)._ColumnProps(26)=   "Column(4).Order=5"
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
      DefColWidth     =   0
      HeadLines       =   1.5
      FootLines       =   1
      MultipleLines   =   0
      CellTipsWidth   =   0
      DeadAreaBackColor=   12632256
      RowDividerColor =   12632256
      RowSubDividerColor=   12632256
      DirectionAfterEnter=   1
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
      _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bold=0,.fontsize=825,.italic=0"
      _StyleDefs(7)   =   ":id=1,.underline=0,.strikethrough=0,.charset=0"
      _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
      _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
      _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.bold=0,.fontsize=825"
      _StyleDefs(11)  =   ":id=2,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
      _StyleDefs(31)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8"
      _StyleDefs(32)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
      _StyleDefs(33)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
      _StyleDefs(34)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
      _StyleDefs(35)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
      _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=28,.parent=13"
      _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=25,.parent=14"
      _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=26,.parent=15"
      _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=27,.parent=17"
      _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13"
      _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
      _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
      _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
      _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=50,.parent=13,.alignment=0"
      _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=47,.parent=14"
      _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=48,.parent=15"
      _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=49,.parent=17"
      _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=46,.parent=13,.alignment=1"
      _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=43,.parent=14"
      _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=44,.parent=15"
      _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=45,.parent=17,.alignment=1"
      _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=54,.parent=13,.alignment=1"
      _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=51,.parent=14"
      _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=52,.parent=15"
      _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=53,.parent=17"
      _StyleDefs(56)  =   "Named:id=33:Normal"
      _StyleDefs(57)  =   ":id=33,.parent=0"
      _StyleDefs(58)  =   "Named:id=34:Heading"
      _StyleDefs(59)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(60)  =   ":id=34,.wraptext=-1"
      _StyleDefs(61)  =   "Named:id=35:Footing"
      _StyleDefs(62)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
      _StyleDefs(63)  =   "Named:id=36:Selected"
      _StyleDefs(64)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(65)  =   "Named:id=37:Caption"
      _StyleDefs(66)  =   ":id=37,.parent=34,.alignment=2"
      _StyleDefs(67)  =   "Named:id=38:HighlightRow"
      _StyleDefs(68)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
      _StyleDefs(69)  =   "Named:id=39:EvenRow"
      _StyleDefs(70)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
      _StyleDefs(71)  =   "Named:id=40:OddRow"
      _StyleDefs(72)  =   ":id=40,.parent=33"
      _StyleDefs(73)  =   "Named:id=41:RecordSelector"
      _StyleDefs(74)  =   ":id=41,.parent=34"
      _StyleDefs(75)  =   "Named:id=42:FilterBar"
      _StyleDefs(76)  =   ":id=42,.parent=33"
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   34
      Top             =   0
      Width           =   5850
      _ExtentX        =   10319
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   15
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Grabar"
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cancelar"
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Quitar Filtro"
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Excel"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Exportar"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "-"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Crear Formato a Importar"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Importar Datos"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   10
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
End
Attribute VB_Name = "FrmManTC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RstTC As New ADODB.Recordset, RsCons As New ADODB.Recordset
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim xTitulo As String, CaracteresNumericos As String
Dim vStrSql As String

'---------
Dim Oleapp As Object

Dim BAND_INTERRUMPIR As Boolean

Function fverif_tc_deundia() As Boolean
    Dim rsverif As New ADODB.Recordset
    vStrSql = "SELECT fecha From con_tc WHERE fecha = DATEVALUE('" & Trim(TxtFechaRegistro.Valor) & "')"
    RST_Busq rsverif, vStrSql, xCon
    If rsverif.RecordCount > 0 Then
        MsgBox "Ya ingreso un tipo de cambio en la fecha seleccionada", vbOKOnly + vbInformation + vbDefaultButton1, xTitulo
        fverif_tc_deundia = True
        TxtFechaRegistro.SetFocus
    Else
        fverif_tc_deundia = False
    End If
    Set rsverif = Nothing
End Function
Sub confpredeter()
    TxtAnio.Text = Year(Date)
    UpDown1.Value = Year(Date)
    '--------

End Sub

'Const CaracteresNumericos = "0123456789." & Chr(8)
Sub centrarFrame()
    FraRegistrar.Top = (Me.Height - FraRegistrar.Height) \ 2
    FraRegistrar.Left = (Me.Width - FraRegistrar.Width) \ 2
End Sub
Sub LlenarGrid(pConFiltroSiNO As Boolean)
    Set RstTC = New ADODB.Recordset
    Select Case pConFiltroSiNO
        Case True 'SI
            RST_Busq RstTC, "SELECT con_tc.id, con_tc.fecha, mae_moneda.descripcion, con_tc.impcom, con_tc.impven " _
                & " FROM mae_moneda INNER JOIN con_tc ON mae_moneda.id = con_tc.idmon WHERE con_tc.idmon = " & Val(DtcTMoneda_Filtro.BoundText) & "" _
                & " AND YEAR(fecha) = " & Val(TxtAnio.Text) & " AND MONTH(fecha) = " & Val(DtcMesFiltro.BoundText) & " ORDER BY con_tc.fecha, con_tc.id", xCon
        Case Else 'NO
            RST_Busq RstTC, "SELECT con_tc.id, con_tc.fecha, mae_moneda.descripcion, con_tc.impcom, con_tc.impven " _
                & " FROM mae_moneda INNER JOIN con_tc ON mae_moneda.id = con_tc.idmon ORDER BY con_tc.fecha, con_tc.id", xCon
    End Select
    Set Dg1.DataSource = RstTC
End Sub

Private Sub LlenarTextosPaModifi()
    RST_Busq RsCons, "SELECT con_tc.id as IDTC, con_tc.fecha, mae_moneda.id AS TMONEDA, con_tc.impcom, con_tc.impven " _
        & " From con_tc INNER JOIN mae_moneda ON con_tc.idmon = mae_moneda.id WHERE con_tc.id = " & Val(Dg1.Columns(0).Text) & "", xCon
    With RsCons
        If .RecordCount > 0 Then
            DtcTMoneda.BoundText = .Fields("TMONEDA")
            TxtFechaRegistro.Valor = Trim(.Fields("fecha"))
            TxtTC_Compra.Text = Format(Val(.Fields("impcom")), "#####0.0000")
            TxtTC_Venta.Text = Format(Val(.Fields("impven")), "#####0.0000")
        End If
    End With
    Set RsCons = Nothing
End Sub

Private Sub CmdAceptar_Click()
    If QueHace = 1 Then 'grabar nuevo reg
        If Grabar = True Then
            '--LIMPIAR PARA CONTINUAR AGREGANDO TC
            TxtFechaRegistro.Valor = DateAdd("d", 1, CDate(TxtFechaRegistro.Valor))
            TxtFechaRegistro.SetFocus
        End If
    ElseIf QueHace = 2 Then
        Grabar
        CmdCancelar.Value = True
    End If
    RstTC.Requery
    Dg1.Refresh
End Sub

Private Sub CmdCancelar_Click()
    Cancelar
'    Bloquea (True)
'    FraRegistrar.Visible = False
    Dg1.SetFocus
End Sub


Private Sub DtcMesFiltro_Change()
    If Trim(TxtAnio.Text) = "" Then
        MsgBox "Falta especificar el año...!", vbExclamation, "Mensaje...!"
        TxtAnio.SetFocus
        Exit Sub
    End If
    If Trim(DtcTMoneda_Filtro.Text) <> "" And Trim(DtcMesFiltro.Text) <> "" Then
        LlenarGrid True
    End If
End Sub


Private Sub DtcTMoneda_Filtro_Change()
    If Trim(DtcTMoneda_Filtro.Text) <> "" And Trim(DtcMesFiltro.Text) <> "" Then
        LlenarGrid True
    End If
End Sub

Private Sub DtcTMoneda_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 27
            Cancelar
            FraRegistrar.Visible = False
        Case 13
            If TxtFechaRegistro.Enabled = True Then
                TxtFechaRegistro.SetFocus
            Else
                TxtTC_Compra.SetFocus
            End If
    End Select
End Sub



Private Sub Form_Activate()
    If SeEjecuto = False Then
        CaracteresNumericos = "0123456789." & Chr(8)
'        RST_Busq RstTC, "SELECT con_tc.id, con_tc.fecha, mae_moneda.descripcion, con_tc.impcom, con_tc.impven " _
'            & " FROM mae_moneda INNER JOIN con_tc ON mae_moneda.id = con_tc.idmon ORDER BY con_tc.id", xCon
'        Set Dg1.DataSource = RstTC
        LlenarGrid False
        
        RST_Busq RsCons, "SELECT id, descripcion From mae_moneda ORDER BY id", xCon
        Set DtcTMoneda_Filtro.RowSource = RsCons
        DtcTMoneda_Filtro.ListField = "descripcion"
        DtcTMoneda_Filtro.BoundColumn = "id"
        
        Set DtcTMoneda.RowSource = RsCons
        DtcTMoneda.ListField = "descripcion"
        DtcTMoneda.BoundColumn = "id"
        Set RsCons = Nothing
        
        RST_Busq RsCons, "SELECT id, descripcion From con_meses ORDER BY id", xCon
        Set DtcMesFiltro.RowSource = RsCons
        DtcMesFiltro.ListField = "descripcion"
        DtcMesFiltro.BoundColumn = "id"
        Set RsCons = Nothing
        
        DtcTMoneda_Filtro.Text = "Dolares"
        DtcMesFiltro.BoundText = Month(Date)
        '**********************************
        Set DtcTMoneda1.RowSource = DtcTMoneda_Filtro.RowSource
        DtcTMoneda1.ListField = "descripcion"
        DtcTMoneda1.BoundColumn = "id"
        
        Set DtcMesFiltro1.RowSource = DtcMesFiltro.RowSource
        DtcMesFiltro1.ListField = "descripcion"
        DtcMesFiltro1.BoundColumn = "id"
        
        DtcTMoneda1.Text = "Dolares"
        DtcMesFiltro1.BoundText = Month(Date)
        
        '**********************************
        confpredeter
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 27
            If FraRegistrar.Visible = True Then CmdCancelar_Click
                
            If fraImportar.Visible = True Then CmdImportar_Click 2

    End Select
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    QueHace = 2
'    Frame1.BackColor = &HC0C0C0
'    Frame2.BackColor = &HC0C0C0
'    TabOne1.CurrTab = 0
End Sub

Sub MuestraSegundoTab()
'    TxtNumRuc.Text = RstTC("numruc")
'    TxtEmpresa.Text = RstTC("nomemp")
'    TxtRutaData.Text = RstTC("ruta")
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 1 Then Exit Sub
        MuestraSegundoTab
    End If
End Sub

Sub Nuevo()
    QueHace = 1
    Toolbar
    Blanquea
    centrarFrame
    FraRegistrar.Visible = True
    Call Bloquea(False)
    TxtFechaRegistro.Enabled = True
'    TabOne1.CurrTab = 1
'    TabOne1.TabEnabled(0) = False
'    Label1.Caption = "Agregando Empresa"
    LblTituloFrame.Caption = " Nuevo Tipo de Cambio"
    DtcTMoneda.SetFocus
End Sub

Sub Modificar()
    LlenarTextosPaModifi
    QueHace = 2
    Toolbar
'    Blanquea
    centrarFrame
    FraRegistrar.Visible = True
    Call Bloquea(False)
    TxtFechaRegistro.Enabled = False
'    TabOne1.CurrTab = 1
'    TabOne1.TabEnabled(0) = False
'    Label1.Caption = "Modificando Empresa"
    LblTituloFrame.Caption = " Modificar Tipo de Cambio"
    DtcTMoneda.SetFocus
End Sub

Sub Blanquea()
    DtcTMoneda.Text = ""
    TxtFechaRegistro.Valor = Format(Date, "dd/mm/yyyy")
    TxtTC_Compra.Text = ""
    TxtTC_Venta.Text = ""
End Sub

Sub Bloquea(pBool As Boolean)
    DtcTMoneda_Filtro.Enabled = pBool
    TxtAnio.Enabled = pBool
    UpDown1.Enabled = pBool
    DtcMesFiltro.Enabled = pBool
    Dg1.Enabled = pBool
End Sub

Sub Eliminar()
    Dim Rpta As Integer
    Rpta = MsgBox("Esta seguro de eliminar el tipo de cambio seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        xCon.Execute "DELETE FROM con_tc WHERE id = " & RstTC("id") & ""
        MsgBox "El tipo de cambio se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstTC.Requery
        Dg1.Refresh
    End If
End Sub

Private Sub pic_Click()
    CmdImportar_Click 2
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Index = 1 Then Exportar
    If ButtonMenu.Index = 3 Then pCrearFormato
    If ButtonMenu.Index = 4 Then ImportarExcel
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        Nuevo
    End If
    If Button.Index = 2 Then
        If Dg1.ApproxCount <= 0 Then
            MsgBox "No hay ningun registro para modificar...!", vbExclamation, "Mensaje...!"
            Exit Sub
        End If
        Modificar
    End If
    If Button.Index = 3 Then
        If Dg1.ApproxCount <= 0 Then
            MsgBox "No hay ningun registro para eliminar...!", vbExclamation, "Mensaje...!"
            Exit Sub
        End If
        Eliminar
    End If
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstTC.Requery
            Dg1.Refresh
        End If
    End If
    
    If Button.Index = 6 Then
        Cancelar
    End If
    
    
    If Button.Index = 15 Then
        Set RstTC = Nothing
        Unload Me
    End If
End Sub

Function Grabar() As Boolean
    Grabar = False
    If NulosC(DtcTMoneda.Text) = "" Then
        MsgBox "No ha seleccionado el tipo de moneda.", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        DtcTMoneda.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtFechaRegistro.Valor) = "" Then
        MsgBox "No ha especificado la fecha de registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFechaRegistro.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtTC_Compra.Text) = "" Then
        MsgBox "No ha especificado el importe del tipo de cambio de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTC_Compra.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtTC_Venta.Text) = "" Then
        MsgBox "No ha especificado el importe del tipo de cambio de venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTC_Venta.SetFocus
        Exit Function
    End If
    If QueHace = 1 Then
        If fverif_tc_deundia = True Then
            Exit Function
        End If
    End If
        
    Dim RstGra As New ADODB.Recordset
    Dim xId As Integer
    
    On Error GoTo LaCague
    
    xCon.BeginTrans
    
    If QueHace = 1 Then
        xId = HallaCodigoTabla("con_tc", xCon, "id")
        
        RST_Busq RstGra, "SELECT top 1 * FROM con_tc", xCon
        RstGra.AddNew
        RstGra("id") = xId
    Else
        RST_Busq RstGra, "SELECT * FROM con_tc WHERE id = " & RstTC("id") & "", xCon
    End If
    
    RstGra("fecha") = Trim(TxtFechaRegistro.Valor)
    RstGra("idmon") = DtcTMoneda.BoundText
    RstGra("impcom") = Val(TxtTC_Compra.Text)
    RstGra("impven") = Val(TxtTC_Venta.Text)
    RstGra.Update
    
    MsgBox "El tipo de cambio se grabo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    xCon.CommitTrans
    Grabar = True
    Exit Function
    
LaCague:
    xCon.RollbackTrans
    Set RstGra = Nothing
    
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

Sub Cancelar()
    Toolbar
    QueHace = 3
    Blanquea
    Bloquea (True)
    FraRegistrar.Visible = False
'    TabOne1.TabEnabled(0) = True
'    TabOne1.CurrTab = 0
    
End Sub

Sub Toolbar()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
'    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
'    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    
    Toolbar1.Buttons(12).Enabled = Not Toolbar1.Buttons(12).Enabled
    
    Toolbar1.Buttons(14).Enabled = Not Toolbar1.Buttons(14).Enabled
End Sub



Private Sub TxtAnio_Change()
    If Trim(DtcTMoneda_Filtro.Text) <> "" And Trim(DtcMesFiltro.Text) <> "" Then
        LlenarGrid True
    End If
End Sub

Private Sub TxtAnio_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub


Private Sub TxtTC_Compra_GotFocus()
    TxtTC_Compra.SelLength = Len(TxtTC_Compra.Text)
End Sub

Private Sub TxtTC_Compra_KeyPress(KeyAscii As Integer)
'    CaracteresNumericos = "0123456789." & Chr(8)
    Select Case KeyAscii
        Case 13
            TxtTC_Venta.SetFocus
        Case 27
            Cancelar
            FraRegistrar.Visible = False
        Case Else
            If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End Select
End Sub

Private Sub TxtTC_Compra_LostFocus()
    TxtTC_Compra.Text = Format(Trim(TxtTC_Compra.Text), "###0.0000")
End Sub

Private Sub TxtTC_Venta_GotFocus()
    TxtTC_Venta.SelLength = Len(TxtTC_Venta.Text)
End Sub

Private Sub TxtTC_Venta_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
        Case 13
            CmdAceptar.SetFocus
'            Call Toolbar1_ButtonClick(Toolbar1.Buttons.Item(5))
        Case 27
            Cancelar
            FraRegistrar.Visible = False
        Case Else
            If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End Select
End Sub

Private Sub TxtTC_Venta_LostFocus()
    TxtTC_Venta.Text = Format(Trim(TxtTC_Venta.Text), "###0.0000")
End Sub

Private Sub UpDown1_Change()
    TxtAnio.Text = UpDown1.Value
End Sub

'--------------

'*******************************************************************

'************************************************************************
'-MODIFICADO AL 25/01/08
'*IMPORTAR DATOS, EXPORTAR DATOS

Private Sub ImportarExcel()
    Bloquea False
    '--------
    txtAnio1.Text = AnoTra
    UpDown2.Value = AnoTra
    
    fraImportar.Visible = True
    fraImportar.Left = 890
    fraImportar.Top = 510
    Fg1.Rows = 1
    CmdImportar(1).Enabled = False
    CmdImportar(0).SetFocus
    
    Dim A As Integer
    For A = 1 To Me.Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
    
End Sub


Private Sub CmdImportar_Click(Index As Integer)
    Select Case Index
        Case 0 '--ARCHIVO
              pImportar
        Case 1 '--GRABAR IPMORTAR
            If GrabarEnLote() = True Then
                TxtAnio_Change
                CmdImportar_Click 2
            End If
        Case 2 '--CANCELAR
            fraImportar.Visible = False
            Bloquea True
            Dim A As Integer
            For A = 1 To Me.Toolbar1.Buttons.Count
                Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
            Next A
            TxtAnio.SetFocus
        
    End Select
End Sub


Private Sub pCrearFormato()

    On Error GoTo error
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    objExcel.SheetsInNewWorkbook = 1
    objExcel.WindowState = 1
    objExcel.Workbooks.Add
    
    With objExcel.ActiveSheet
        .Cells(1, 1) = "Tipo de Cambio"
        .Cells(3, 1) = "Dia"
        .Cells(3, 2) = "Imp. Compra"
        .Cells(3, 3) = "Imp. Venta"
        '---------
        .Columns(1).ColumnWidth = 8
        .Columns(2).ColumnWidth = 15
        .Columns(2).ColumnWidth = 15
        '---
        .Cells(1, 1).Font.Bold = True
        .Cells(1, 1).Font.Size = 15
        .Cells(3, 1).Font.Bold = True
        .Cells(3, 2).Font.Bold = True
        .Cells(3, 3).Font.Bold = True
                
    End With
    MsgBox "Proceda a ingresar la información según los Parámetros Solicitados" + vbCr + "Luego proceda a Importar...", vbInformation, xTitulo
    objExcel.Visible = True
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
error:
    Set objExcel = Nothing
End Sub


Private Sub pImportar()
    Dim vPath As String
    Dim VL_ROW&, VL_COL&
    
    cmm.CancelError = False
    cmm.FileName = ""
    cmm.Filter = "Archivos xls (*.xls)|*.xls"
    cmm.ShowOpen
    vPath = cmm.FileName
    If vPath = "" Then Exit Sub
    If ArchivoExiste(vPath) = False Then
        MsgBox "El archivo no Existe", vbExclamation, xTitulo
        Exit Sub
    End If
    
'''    '--CARGAR DATOS DEL EXCEL AL GRID
'''    Fg1.LoadGrid vPath, flexFileCommaText
'''    '--ELIMINAR TITULOS DEL EXCEL
'''    Fg1.RemoveItem 1
'''    Fg1.RemoveItem 1
'''    Fg1.RemoveItem 1
'''    '--ELIMINAR ULTIMO FILA EN BLANCO(POR DEFECTO)
'''    Fg1.Rows = Fg1.Rows - 1
    
    '** crear formato de grid
    With Fg1
        .Cols = 4
        .Rows = 1
        .TextMatrix(0, 1) = "Dia":          .ColWidth(1) = 870:    .ColAlignment(1) = flexAlignRightBottom
        .TextMatrix(0, 2) = "Imp. Compra":  .ColWidth(2) = 1200:   .ColAlignment(2) = flexAlignRightBottom
        .TextMatrix(0, 3) = "Imp. Venta":   .ColWidth(3) = 1200:   .ColAlignment(3) = flexAlignRightBottom
        .ColFormat(2) = "##.0000"
        .ColFormat(3) = "##.0000"
    End With
    '**********************************************************
    On Error GoTo error
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    Me.MousePointer = vbHourglass
    
    objExcel.Visible = False
    objExcel.SheetsInNewWorkbook = 1
    'Crea el Libro
    objExcel.Workbooks.Open vPath

    Dim xFila&
    xFila = 4
    With objExcel.ActiveSheet
        Do While NulosN(.Cells(xFila, 1)) <> 0
            DoEvents
            Fg1.Rows = Fg1.Rows + 1
            
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosN(.Cells(xFila, 1))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosN(.Cells(xFila, 2))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosN(.Cells(xFila, 3))
   
            xFila = xFila + 1
        Loop
    End With
    Set objExcel = Nothing
    Me.MousePointer = vbDefault
    
    MsgBox "El Archivo se Importó Correctamente", vbInformation, xTitulo
    '--ORDENAR ASCENDENTE DEGUN DIA
    If Fg1.Rows > 1 Then
        Fg1.Select 1, 1
        Fg1.Sort = flexSortGenericAscending
    End If
    '**********************************************************

    CmdImportar(1).Enabled = True
    txtAnio1.SetFocus
    DtcMesFiltro1.Text = ""
    Exit Sub
    
error:
    Me.MousePointer = vbDefault
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "pImportar"
    CmdImportar_Click 2
End Sub


Private Sub Exportar()
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    'On Error GoTo error
    
    If RstTC.RecordCount = 0 Then
        MsgBox "No hay datos para exportar", vbExclamation, xTitulo
        Exit Sub
    End If
    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    'abre el Libro
    objExcel.WindowState = 1
    objExcel.Workbooks.Add  'Trim(App.Path) + "\RegCompras.xls"
    
    LblTituloProg.Caption = "Exportando a excel..."
    FraProgreso.Visible = True
    PgBar.Max = RstTC.RecordCount
    PgBar.Min = 1

    
    With objExcel.ActiveSheet
        
        .Cells(1, 2) = NomEmp
        .Cells(1, 5) = Date
        .Cells(2, 2) = "Nº R.U.C. : " + NumRUC
        
        .Columns(2).ColumnWidth = 6
        .Columns(3).ColumnWidth = 15
        .Columns(4).ColumnWidth = 15
        
        .Cells(4, 2) = "TIPO DE CAMBIO"
        .Cells(5, 2) = " Año: " & Val(TxtAnio.Text) & " Mes: " & DtcMesFiltro.Text
        .Cells(6, 2) = " Moneda: " & DtcTMoneda_Filtro.Text
        xFilas = 8
        .Cells(xFilas, 2) = "Dia"
        .Cells(xFilas, 3) = "Imp. Compra"
        .Cells(xFilas, 4) = "Imp. Venta"
        
        xFilas = xFilas + 1
        RstTC.MoveFirst
        Do While Not RstTC.EOF
        
            PgBar.Value = RstTC.Bookmark
            DoEvents
            If BAND_INTERRUMPIR = True Then
                FraProgreso.Visible = False
                Exit Sub
            End If
            .Cells(xFilas, 2) = Day(RstTC("fecha"))
            .Cells(xFilas, 3) = RstTC("impcom")
            .Cells(xFilas, 4) = RstTC("impven")
            RstTC.MoveNext
            xFilas = xFilas + 1
        Loop
    
    End With
    
    MsgBox "El proceso terminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, "Tipo de Cambio"
    objExcel.WindowState = 1
    Set objExcel = Nothing
    FraProgreso.Visible = False
    Exit Sub
error:
    FraProgreso.Visible = False
    Set objExcel = Nothing
    SHOW_ERROR Me.Name, "ExportarExcelDetalle", , IIf(Err.Number <> 50290, "", "No manipule el archivo hasta que termine de exportar!!!!")
End Sub

Private Function GrabarEnLote() As Boolean
    If NulosC(DtcTMoneda1.Text) = "" Then
        MsgBox "No ha seleccionado el tipo de moneda.", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        DtcTMoneda1.SetFocus
        Exit Function
    End If
    
    If Trim(txtAnio1.Text) = "" Then
        MsgBox "Ingrese un año", vbExclamation, xTitulo
        txtAnio1.SetFocus
        Exit Function
    End If
    
    If (DtcMesFiltro1.BoundText = "0" Or DtcMesFiltro1.BoundText = "13") Or DtcMesFiltro1.MatchedWithList = False Then
        MsgBox "Seleccione un mes Correcto", vbExclamation, xTitulo
        DtcMesFiltro1.SetFocus
        SendKeys "{F4}"
        Exit Function
    End If
    
    If MsgBox("Seguro desea grabar los datos" + vbCr + "Año:          " + txtAnio1.Text + vbCr + "Mes:          " + DtcMesFiltro1.Text + vbCr + "Moneda:   " + DtcTMoneda1.Text, vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Function
    
    Dim RstCab As New ADODB.Recordset
    Dim xId As Integer
    
    On Error GoTo LaCague
    
    xCon.BeginTrans
        
    Dim mRow&
    Dim nSQL As String
    Dim mIdMoneda As String
    Dim mIdMes As String
    Dim rstTmp As New ADODB.Recordset
    
    With Fg1
        mIdMoneda = DtcTMoneda1.BoundText
        mIdMes = DtcMesFiltro1.BoundText
        Me.MousePointer = vbHourglass
        For mRow = 1 To .Rows - 1
            '--GENERANDO LA CONSULTA PARA VER SI EXISTE EL REGISTRO
            nSQL = "SELECT * FROM con_tc  WHERE fecha = cdate('" + .TextMatrix(mRow, 1) + "/" + Format(mIdMes, "00") + "/" + txtAnio1.Text + "') and idmon = " + mIdMoneda + " ;"
            RST_Busq rstTmp, nSQL, xCon
            If rstTmp.RecordCount <> 0 Then
                RST_Busq RstCab, "SELECT * FROM con_tc WHERE id = " & rstTmp("id") & "", xCon
            Else
                RST_Busq RstCab, "SELECT top 1 * FROM con_tc", xCon
                RstCab.AddNew
                xId = HallaCodigoTabla("con_tc", xCon, "id")
                RstCab("id") = xId
            End If
            Set rstTmp = Nothing
            '---
            
            RstCab("fecha") = CDate(.TextMatrix(mRow, 1) + " / " + Format(mIdMes, "00") + " / " + txtAnio1.Text)
            RstCab("idmon") = mIdMoneda
            RstCab("impcom") = NulosN(.TextMatrix(mRow, 2))
            RstCab("impven") = NulosN(.TextMatrix(mRow, 3))
            RstCab.Update
        Next mRow
    End With
    Me.MousePointer = vbDefault
    MsgBox "El tipo de cambio se grabó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    xCon.CommitTrans
    GrabarEnLote = True
    Exit Function
    
LaCague:
    Me.MousePointer = vbDefault
    Set rstTmp = Nothing
    xCon.RollbackTrans
    Set RstCab = Nothing
    
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + vbCr + Trim(Err.Description), vbCritical, "Error"
End Function

Private Sub UpDown2_Change()
    txtAnio1.Text = UpDown2.Value
End Sub

Private Sub txtAnio1_KeyPress(KeyAscii As Integer)
    If validar_numero(KeyAscii) = False Then KeyAscii = 0
    If KeyAscii = 13 Then DtcMesFiltro1.SetFocus
End Sub

Private Sub DtcMesFiltro1_Click(Area As Integer)
    If Area <> 2 Then Exit Sub
    If Trim(txtAnio1.Text) = "" Then
        MsgBox "Falta especificar el año...!", vbExclamation, "Mensaje...!"
        txtAnio1.SetFocus
        Exit Sub
    End If
End Sub

Private Sub DtcMesFiltro1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then DtcTMoneda1.SetFocus
End Sub

Private Sub DtcTMoneda1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then CmdImportar(1).SetFocus
End Sub
