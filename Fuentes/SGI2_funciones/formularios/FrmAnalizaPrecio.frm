VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FrmAnalizaPrecio 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Compras  - Analisis de Precios y Cantidades"
   ClientHeight    =   5955
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   11610
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   5955
   ScaleWidth      =   11610
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   315
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   11610
      _ExtentX        =   20479
      _ExtentY        =   556
      ButtonWidth     =   609
      ButtonHeight    =   556
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Consultar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   6855
         Top             =   45
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   15
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":0000
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":08D6
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":0A5A
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":0EAE
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":0FC6
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":150A
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":1A4E
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":1B62
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":1C76
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":20CA
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":2236
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":277E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":2A98
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmAnalizaPrecio.frx":2E2A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComDlg.CommonDialog Cmmg 
      Left            =   10725
      Top             =   2145
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3090
      TabIndex        =   4
      Top             =   2760
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Index           =   0
         Left            =   210
         TabIndex        =   5
         Top             =   465
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Index           =   1
         Left            =   210
         TabIndex        =   10
         Top             =   795
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label LBL 
         AutoSize        =   -1  'True
         Caption         =   "Datos"
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
         Index           =   5
         Left            =   5025
         TabIndex        =   12
         Top             =   840
         Width           =   495
      End
      Begin VB.Label LBL 
         AutoSize        =   -1  'True
         Caption         =   "Año:"
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
         Index           =   4
         Left            =   5025
         TabIndex        =   11
         Top             =   495
         Width           =   825
      End
      Begin VB.Label LBL 
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
         Height          =   255
         Index           =   2
         Left            =   4275
         TabIndex        =   9
         Top             =   150
         Width           =   1530
      End
      Begin VB.Label LBL 
         AutoSize        =   -1  'True
         Caption         =   "Procesando:"
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
         Index           =   0
         Left            =   225
         TabIndex        =   7
         Top             =   150
         Width           =   1035
      End
      Begin VB.Label LBL 
         AutoSize        =   -1  'True
         Caption         =   "Compras"
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
         Index           =   1
         Left            =   1320
         TabIndex        =   6
         Top             =   150
         Width           =   750
      End
      Begin VB.Shape Shape1 
         Height          =   1065
         Left            =   60
         Top             =   60
         Width           =   5805
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Align           =   2  'Align Bottom
      Height          =   3930
      Left            =   0
      TabIndex        =   8
      Top             =   2025
      Width           =   11610
      _cx             =   20479
      _cy             =   6932
      _ConvInfo       =   1
      Appearance      =   1
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
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
      AllowUserResizing=   1
      SelectionMode   =   0
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmAnalizaPrecio.frx":327C
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
   Begin VB.Frame fr 
      Height          =   1470
      Index           =   5
      Left            =   0
      TabIndex        =   2
      Top             =   375
      Width           =   11595
      Begin VB.CommandButton cb 
         Height          =   210
         Index           =   0
         Left            =   2070
         Picture         =   "FrmAnalizaPrecio.frx":32B8
         Style           =   1  'Graphical
         TabIndex        =   43
         Top             =   1155
         Width           =   195
      End
      Begin VB.Frame Frame1 
         Caption         =   "vs Precio"
         Height          =   960
         Left            =   10395
         TabIndex        =   39
         Top             =   420
         Width           =   1140
         Begin VB.OptionButton opt_precio 
            Caption         =   "Máximo"
            Height          =   195
            Index           =   2
            Left            =   45
            TabIndex        =   42
            Tag             =   "'P. Max'"
            Top             =   705
            Width           =   990
         End
         Begin VB.OptionButton opt_precio 
            Caption         =   "Promedio"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   41
            Tag             =   "'P. Prom'"
            Top             =   465
            Value           =   -1  'True
            Width           =   990
         End
         Begin VB.OptionButton opt_precio 
            Caption         =   "Mínimo"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   40
            Tag             =   "'P. Min'"
            Top             =   225
            Width           =   945
         End
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Volumen"
         Enabled         =   0   'False
         Height          =   195
         Left            =   10410
         TabIndex        =   38
         Top             =   225
         Value           =   1  'Checked
         Width           =   975
      End
      Begin VB.CommandButton cmd 
         Caption         =   "&Editar Color"
         Height          =   285
         Index           =   2
         Left            =   3960
         TabIndex        =   0
         Top             =   210
         Width           =   960
      End
      Begin VB.TextBox txt 
         BackColor       =   &H000000FF&
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
         Index           =   5
         Left            =   3510
         MaxLength       =   2
         TabIndex        =   34
         Text            =   "5"
         Top             =   210
         Width           =   435
      End
      Begin VB.Frame fr 
         Caption         =   "Seleccionar"
         Height          =   1260
         Index           =   0
         Left            =   9000
         TabIndex        =   29
         Top             =   135
         Width           =   1305
         Begin VB.OptionButton opt_mon 
            Caption         =   "Solo S/."
            Height          =   210
            Index           =   0
            Left            =   45
            TabIndex        =   33
            Top             =   270
            Width           =   885
         End
         Begin VB.OptionButton opt_mon 
            Caption         =   "Solo $."
            Height          =   210
            Index           =   1
            Left            =   45
            TabIndex        =   32
            Top             =   495
            Width           =   840
         End
         Begin VB.OptionButton opt_mon 
            Caption         =   "Todo en S/."
            Height          =   210
            Index           =   2
            Left            =   45
            TabIndex        =   31
            Top             =   720
            Value           =   -1  'True
            Width           =   1185
         End
         Begin VB.OptionButton opt_mon 
            Caption         =   "Todo en $."
            Height          =   210
            Index           =   3
            Left            =   45
            TabIndex        =   30
            Top             =   945
            Width           =   1170
         End
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   4
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   28
         Top             =   810
         Width           =   2040
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   3
         Left            =   3855
         Locked          =   -1  'True
         TabIndex        =   27
         Top             =   825
         Width           =   1065
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   2
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   26
         Top             =   510
         Width           =   3960
      End
      Begin VB.TextBox txt 
         Height          =   285
         Index           =   1
         Left            =   960
         Locked          =   -1  'True
         TabIndex        =   25
         Top             =   210
         Width           =   2040
      End
      Begin VB.TextBox txt 
         BackColor       =   &H0000FF00&
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   0
         Left            =   2430
         Locked          =   -1  'True
         TabIndex        =   24
         Top             =   90
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.ListBox ls 
         Height          =   960
         Index           =   1
         Left            =   7410
         Style           =   1  'Checkbox
         TabIndex        =   16
         Top             =   435
         Width           =   1530
      End
      Begin VB.Frame fr 
         Caption         =   "Seleccionar"
         Height          =   1260
         Index           =   2
         Left            =   6300
         TabIndex        =   13
         Top             =   135
         Width           =   1095
         Begin VB.OptionButton opt_estilo 
            Caption         =   "Trimestre"
            Height          =   210
            Index           =   1
            Left            =   60
            TabIndex        =   18
            Top             =   502
            Width           =   960
         End
         Begin VB.OptionButton opt_estilo 
            Caption         =   "Mes"
            Height          =   210
            Index           =   0
            Left            =   60
            TabIndex        =   15
            Top             =   270
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton opt_estilo 
            Caption         =   "Semestre"
            Height          =   210
            Index           =   2
            Left            =   60
            TabIndex        =   14
            Top             =   735
            Visible         =   0   'False
            Width           =   960
         End
      End
      Begin VB.ListBox ls 
         Height          =   960
         Index           =   0
         Left            =   5055
         Style           =   1  'Checkbox
         TabIndex        =   1
         Top             =   435
         Width           =   1200
      End
      Begin VB.TextBox txt_cb 
         Height          =   285
         Index           =   0
         Left            =   960
         MaxLength       =   12
         TabIndex        =   44
         Text            =   "txt_cb(0)"
         Top             =   1125
         Width           =   1335
      End
      Begin VB.Label lbl_cb_capt 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   47
         Top             =   1230
         Width           =   735
      End
      Begin VB.Label lbl_cod 
         BackColor       =   &H000000FF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cod(0)"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   300
         Index           =   0
         Left            =   3240
         TabIndex        =   46
         Top             =   1125
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.Label lbl_cb 
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "lbl_cb(0)"
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
         Height          =   285
         Index           =   0
         Left            =   2280
         TabIndex        =   45
         Top             =   1125
         Width           =   2610
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Top."
         Height          =   195
         Index           =   5
         Left            =   3150
         TabIndex        =   35
         Top             =   270
         Width           =   330
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800000&
         BorderWidth     =   2
         X1              =   4995
         X2              =   4995
         Y1              =   225
         Y2              =   1335
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Item"
         Height          =   195
         Index           =   4
         Left            =   60
         TabIndex        =   23
         Top             =   900
         Width           =   660
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "U.M."
         Height          =   195
         Index           =   3
         Left            =   3390
         TabIndex        =   22
         Top             =   900
         Width           =   345
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Descripción"
         Height          =   195
         Index           =   2
         Left            =   60
         TabIndex        =   21
         Top             =   600
         Width           =   840
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "Código"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   20
         Top             =   300
         Width           =   495
      End
      Begin VB.Label x_lbl 
         AutoSize        =   -1  'True
         Caption         =   "ID"
         Height          =   195
         Index           =   0
         Left            =   2220
         TabIndex        =   19
         Top             =   90
         Visible         =   0   'False
         Width           =   165
      End
      Begin VB.Label LBL 
         AutoSize        =   -1  'True
         Caption         =   "Selecc. Mes"
         Height          =   195
         Index           =   6
         Left            =   7455
         TabIndex        =   17
         Top             =   195
         Width           =   885
      End
      Begin VB.Label LBL 
         AutoSize        =   -1  'True
         Caption         =   "Seleccionar Año"
         Height          =   195
         Index           =   3
         Left            =   5055
         TabIndex        =   3
         Top             =   195
         Width           =   1170
      End
   End
   Begin VB.Label Label1 
      Caption         =   "Para ver el histórico de precios. Seleccione una celda; hacer doble click o presionar Enter."
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   0
      TabIndex        =   36
      Top             =   1845
      Width           =   8925
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Ver Detalle"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Consultar"
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_5 
         Caption         =   "Exportar MSExcel"
      End
      Begin VB.Menu Menu1_6 
         Caption         =   "Imprimir"
      End
   End
End
Attribute VB_Name = "FrmAnalizaPrecio"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'------------------------------------------------------------------------------------------
'PARA QUE ESTE FORMULARIO SE CARGUE ES NECESARIO TENER EL CODIGO DEL ITEM = STRING
'--FUNCION RECIBE_ID_ITEM(ID_ITEM,[true::ventana de compras,false::ventana de ventas])
'------------------------------------------------------------------------------------------
'-- ALMACENAR LOS TOTALES DE TODA LA CONSULTA
Dim Arr_Totales_cols() As Double '--ALMACENAR TOTALES POR TODAS LAS FILAS
'meses(1-12) or trim(1-4), tot, tot gral
'(?,0) precio,
'(?.1) volumen
Dim Arr_Totales_col() As Double     '--ALMACENAR TOTALES POR COLUMNA, SE LIMPIA DESPUES DE CAMBIO DE GRUPO
Dim Arr_Totales_cuenta() As Double     '--ALMACENAR CANTIDAD DE REGISTROS DIFERENTES DE 0 PARA CALCULAR LA MEDIA POR COLUMNA (ACUMULADO/REGISTROS<> DE CERO)

Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
'--DE LA IMPRESION
Dim T_RPT_PERIODO As String '--PERIODO DEL REPORTE
Dim T_RPT_TITULO As String  '--TITULO DE REPORTE
'------------
Dim ARR_ANYO() As String    '--ARRAY DE AÑOS SELECCIONADOS
Dim ARR_XX() As String      '--SE CARGARA CUANDO SE CARGA EL FORMULARIO Y CUANDO SE CAMBIE EL ESTILO(MES, TRIMESTRE,SEMESTRE)
Dim ARR_TMP() As String     '--DEPENDERA DEL ESTILO SOLO CARGARA LO QUE SELECCIONA


                            '--SE USA PARA DAR FORMATO DE LA GRILLA, SEGUN SELECCIONE EL USUARIO
Dim Q_TOTAL_ANYO As Integer '--INDICA LA CANTIDAD DE AÑOS DE BUSQUEDA,
                            '--EJ. 2004,2005 => Q_TOTAL_ANYO = 2
                            '--EJ. 2004,2005,2006 => Q_TOTAL_ANYO = 3
                            
Dim Q_COL_FILA As Integer   '--INDICA LA CANTIDAD DE COLUMNAS ANTES DE LOS MESES
                            '--EJ. IDCLI,CLIENTE => Q_COL_FILA=2
                            '--    IDCLI,ID_PTO_VTA,CLIENTE,PTO_VENTA => Q_COL_FILA=4
                            
Dim Q_POS_MES_INICIO As Integer '--INDICA LA POSICION INICIAL DE LA COLUMNA DEL PRIMER MES, NO CAMBIA
                            '--EJ. Q_POS_MES_INICIO = Q_COL_FILA +1

Dim Q_POS_MES As Integer    '--INDICA LA POSICION DEL MES, ESTO CAMBIA
                            '--UTIL PARA COLOCAR LOS DATOS EN EL GRID

Dim Q_COL_FILA_OCULTA As Integer '--INDICA LAS COLUMNAS QUE CONTENDRAN LOS ID'S, ESTOS SE OCULTARAN
                                '-- -1 NO SE OCULTA, <> -1 SE PROCEDE A ACULTAR
                                'EJ. CLIENTE  vta_ventas.idcli,
                                    'PUNTO DE VENTA vta_guia.idpunven
                                    'PRODUCTO   alm_inventario.tippro
                                    'ITEM       alm_inventario.id
                                    'EMPLEADO   vta_ventas.idven

Dim Q_POSICION_TOTAL  As Integer '--INDICA LA POCISION DE LA COLUMNA DONDE SE COLOCARA EL NOMBRE DEL TOTAL Y TOTAL_GRL
                                 '--OBTENDRA VALOR EN pGenerarConsulta()

Dim Q_COL_COMPARAR_GRUPO As Integer '--INDICA LA COLUMNA PARA COMPARAR EL GRUPO
                                    '--OBTENDRA VALOR EN pGenerarConsulta()

Dim Q_COL_ARR_TOTAL As Integer  '--NOS INDICA EL TOTAL DE COLUMNAS QUE VA A CONTENER LOS ACUMULADOS
                                '--OBTENDRA VALOR EN VALIDAR_CONSULTA()
                                '--SI SEL MES: ENE, FEB, MAR => Q_COL_ARR_TOTAL= 2
                                '--SI SEL TRI: ENE-MAR, ABR-JUN => Q_COL_ARR_TOTAL= 1 OBS: SE INICIA DESDE POS=0

Dim SGIFlex As New SGI2_funciones.JC_VSFlexGrid
Dim SGIVarios As New SGI2_funciones.JC_Varios

Dim F_ES_COMPRA As Boolean '--INDICA SI ES COMPRA O VENTA
                            '--TRUE::ES COMPRA, FALSE::ES VENTA



Public Sub RECIBE_ID_ITEM(ID_ITEM As String, _
                            Optional F_VENTANA_COMPRA As Boolean = True, Optional M_RUC As String = "")
                            
                            
                            
                            
    On Error GoTo ERROR
    Dim rst_select  As New ADODB.Recordset
    Dim sql_select  As String
    Dim Q_ROW       As Integer
    If ID_ITEM = "" Then GoTo salir:
    '--CONSULTA
    sql_select = " SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, mae_tipoproducto.descripcion " + _
                vbCr + " FROM mae_unidades INNER JOIN (mae_tipoproducto INNER JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed " + _
                vbCr + " WHERE (((alm_inventario.id)=" + ID_ITEM + "));"
    
    RST_Busq rst_select, sql_select, xCon
    If rst_select.State = 0 Then GoTo salir:
    If rst_select.RecordCount = 0 Then GoTo salir:
    txt(0).Text = ID_ITEM '--ID_ITEM (IDENTIFICADOR DE REGISTRO)
    For Q_ROW = 0 To rst_select.Fields.Count - 1
        txt(Q_ROW + 1) = rst_select.Fields(Q_ROW) & ""
    Next Q_ROW
    Set rst_select = Nothing
    '--------
    F_ES_COMPRA = F_VENTANA_COMPRA
    If F_ES_COMPRA = False Then
        Me.Caption = "Ventas - Analisis de Precios y Cantidades"
        LBL(1).Caption = "Ventas"
        lbl_cb_capt(0).Caption = "Cliente"
    End If
    If M_RUC <> "" Then
        txt_cb(0).Text = M_RUC
        txt_cb_Validate 0, False
    End If
        
    '----CONSULTAR
    Fg1.Rows = 1
    
    pConsultar
    
    Exit Sub
salir:
    
    
    
    Set rst_select = Nothing
    
    SGIVarios.habilitar cmd, False
    
    Exit Sub
ERROR:
    Set rst_select = Nothing
    SGIVarios.SHOW_ERROR
End Sub


Private Sub cmd_Click(index As Integer)
    Select Case index
        Case 2 '--CAMBIAR COLOR
            Cmmg.CancelError = False
            Cmmg.ShowColor
            txt(5).BackColor = Cmmg.Color
    End Select
End Sub

Private Sub pConsultar()
'    On Error GoTo error
    Dim rst_select As New ADODB.Recordset
    '--
    Dim CN_TMP As New ADODB.Connection '--CONEX TEMPORAL
    Dim Rst_RUTA As New ADODB.Recordset '--CARGA RUTAS DE BD'S
    
    Dim vStrSelect As String '--RECIBIR LA CONSULTA
    '--CARGANDO RUTAS DE LOS AÑOS SELECCIONADOS
    Dim mAnyo As String
    Dim SQL_ANYO As String
    Dim k&, j&
    
    If Validar_Consulta(mAnyo) = False Then Exit Sub
    
    BAND_INTERRUMPIR = False
    '--CONFIGURAR LA PRESENTACION DE LA CONSULTA
    SGIFlex.LimpiarGrid Me.Fg1
    '--INVOCAR A ESTA FUNCION PARA OBTENER LOS VALORES DE
        '--Q_POS_MES , Q_POS_MES_INICIO
    pGenerarConsulta "-1"
    pConfigurarGrilla
    '--CARGANDO RUTAS DE LOS AÑOS SELECCIONADOS
    SQL_ANYO = " AND anotra IN (" + Left(mAnyo, Len(mAnyo) - 1) + ") "
    '--SI LA BASE DE BATOS PRINCIPAL EXISTE
    If SGIVarios.ArchivoExiste(AP_RUTABD + "data.mdb") = False Then
        MsgBox "No existe la ruta a la Base de Datos Principal", vbCritical, xTitulo
        Exit Sub
    End If
    '--ABRIENDO LA CONEXION PARA OBTENER EL LISTADO DE RUTAS A LAS BASES DE DATOS
    SGIVarios.OPEN_CONEX_TMP CN_TMP, AP_RUTABD + "data.mdb"
    If CN_TMP.State = 0 Then Exit Sub
    RST_Busq rst_select, "SELECT ruta,anotra FROM mae_empresa WHERE numruc = '" + NumRUC + "' " + SQL_ANYO + " ORDER BY anotra ASC ", CN_TMP
    '--CARGAR RST TEMPORAL
    SGIVarios.DEFINIR_RST_TMP Rst_RUTA, rst_select
    SGIVarios.CARGAR_RST_TMP Rst_RUTA, rst_select
    If Rst_RUTA.RecordCount = 0 Then
        MsgBox "No hay Base de Datos", vbInformation
        Exit Sub
    End If
    Rst_RUTA.MoveFirst
    Set rst_select = Nothing
    CN_TMP.Close
    '--LIMPIAR ARRAY
    Limpiar_ARRAY_TOTAL True
    '----
    Me.MousePointer = vbHourglass
    DoEvents
    PgBar(1).Min = 0
    PgBar(1).Value = 0
    PosicionarProgBar
    DoEvents
    PgBar(0).Min = 0
    PgBar(0).Max = Rst_RUTA.RecordCount
    
    For k = 0 To Rst_RUTA.RecordCount - 1 '--de los años
        For j = 0 To 1
            LBL(4).Caption = "Año: " + CStr(Rst_RUTA.Fields(1))
            PgBar(0).Value = k + 1
            '------------------------------------------------
            '--ENTRAR SOLO UNA VEZ
            vStrSelect = pGenerarConsulta(CStr(Rst_RUTA.Fields(1)), IIf(j = 0, True, False))
            If k <> 0 Then
                '--EN LOS DEMAS AÑO REEMPLAZAR EL AÑO ANTERIOR POR EL AÑO ACTUAL
                vStrSelect = Replace(vStrSelect, ARR_ANYO(k - 1), CStr(Rst_RUTA.Fields(1)))
            End If
            '------------------------------------------------
            If vStrSelect = "" Then GoTo salir
            '--SI EL ARCHIVO EXISTE
            If SGIVarios.ArchivoExiste(AP_RUTABD & Rst_RUTA.Fields(0) & "") = False Then
                MsgBox "No existe la ruta a la Base de Datos Año: " + CStr(Rst_RUTA.Fields(1)), vbCritical, xTitulo
                GoTo salir
            End If
            '--ABRIENDO LA CONEXION A LA BASE DE DATOS
            SGIVarios.OPEN_CONEX_TMP CN_TMP, AP_RUTABD + Trim(Rst_RUTA.Fields(0)) & ""
            If CN_TMP.State = 0 Then Exit Sub
            '--CARGADO EL RST
            RST_Busq rst_select, vStrSelect, CN_TMP
    
            '--------------------------------------
    '        If rst_select.RecordCount > 0 Then CARGAR_DATOS_TMP CN_TMP, rst_select
            '--CARGA LOS DATOS DEL PRIMER AÑO
            CARGAR_DATOS_GRILLA rst_select, CStr(Rst_RUTA.Fields(1)), IIf(j = 0, True, False)
            CN_TMP.Close
            '--------------------------------------
            
        Next j
        Rst_RUTA.MoveNext
    Next k
    '-----CUANDO LA CONSULTA ES X AÑOS COLOCAR LOS TOTALES
    CARGAR_DATOS_GRILLA_ADD_TOTALES True, "Vol Tot.", True, True, ARR_ANYO(k - 1), , True
    CARGAR_DATOS_GRILLA_ADD_TOTALES True, opt_precio(0).Tag & " Gral", True, True, ARR_ANYO(k - 1), , False
    '
    PgBar(0).Value = PgBar(0).Max
salir:
    FraProgreso.Visible = False
    Set Rst_RUTA = Nothing
    Set rst_select = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
ERROR:
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    Set rst_select = Nothing
    CN_TMP.Close
    SGIVarios.SHOW_ERROR
    
End Sub
Private Sub CARGAR_DATOS_TMP(CN_TMP As ADODB.Connection, _
                             RST_ORIGEN As ADODB.Recordset)

    Dim RST_TMP As New ADODB.Recordset
'''    Dim RST_GRUPO As New ADODB.Recordset
'''    Dim SQL_CONSULTA As String
'''    Dim N_FILTRO As String
'''    Dim Q_ROW_GRUPO As Integer
    Dim Q_ROW1 As Integer
    Dim Q_ROW_TMP As Integer
    Dim Pos As Integer
    '--
    Dim vStrCampo As String
    
    '--ACUMULANDO LOS DATOS
    Dim Arr_Totales_col_TMP(13, 0) As Double
    Dim Arr_Totales_cuenta_TMP(13, 0) As Double
    Dim Q_CUENTA_REG_TMP As Integer
    
    SGIVarios.DEFINIR_RST_TMP RST_TMP, RST_ORIGEN
    SGIVarios.CARGAR_RST_TMP RST_TMP, RST_ORIGEN, , , True
    
    For Q_ROW_TMP = 0 To RST_ORIGEN.RecordCount - 1
        CARGAR_DATOS_ARRAY RST_ORIGEN, Arr_Totales_col_TMP, Arr_Totales_cuenta_TMP
        RST_ORIGEN.MoveNext
    Next Q_ROW_TMP
    
    '--CARGAR DATOS AL RECORD.. TEMP
    For Q_ROW_TMP = 0 To Q_COL_ARR_TOTAL
        If Arr_Totales_cuenta_TMP(Q_ROW_TMP, 0) <> 0 Then RST_TMP.Fields(Q_ROW_TMP + Q_COL_FILA) = Arr_Totales_col_TMP(Q_ROW_TMP, 0) / Arr_Totales_cuenta_TMP(Q_ROW_TMP, 0)
    Next
    RST_TMP.Fields("total") = 0
    For Q_ROW_TMP = 0 To Q_COL_ARR_TOTAL
        If IsNumeric(RST_TMP.Fields(Q_ROW_TMP + Q_COL_FILA)) = True Then
        RST_TMP.Fields("total") = RST_TMP.Fields("total") + RST_TMP.Fields(Q_ROW_TMP + Q_COL_FILA)
        Q_CUENTA_REG_TMP = Q_CUENTA_REG_TMP + 1
        End If
    Next
    If Q_CUENTA_REG_TMP <> 0 Then RST_TMP.Fields("total") = RST_TMP.Fields("total") / Q_CUENTA_REG_TMP


'''
'''
'''        If RST_ORIGEN.RecordCount > 0 Then
'''            '--CARGAR EL PRIMER REGISTRO
'''            CARGAR_RST_TMP rst_tmp, RST_ORIGEN, "", 0, True
'''            '--CARGAR LOS DEMAS REGISTROS
'''            RST_ORIGEN.MoveFirst
'''            If RST_ORIGEN.EOF = False Then RST_ORIGEN.MoveNext
'''            Do While Not RST_ORIGEN.EOF
'''                DoEvents
'''                '-------Q_ROW1 TOMA VALOR DE LAS COLUMNAS
'''                For Q_ROW1 = 0 To RST_ORIGEN.Fields.Count - 1
'''                    '--SI SE NTERRUMPE EL PROCESO => SALIR
'''                    If BAND_INTERRUMPIR = True Then Exit Sub
'''                    vStrCampo = RST_ORIGEN.Fields(Q_ROW1).Name
'''                    Select Case LCase(vStrCampo)
'''                        Case "total", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"
'''                            rst_tmp.Fields(vStrCampo) = NulosN(rst_tmp.Fields(vStrCampo)) + NulosN(RST_ORIGEN.Fields(vStrCampo))
'''                    End Select
'''                Next Q_ROW1
'''                '--------
'''                RST_ORIGEN.MoveNext
'''            Loop
'''        End If
'''        RST_GRUPO.MoveNext
'''    Next Q_ROW_GRUPO

    '--RENOMBRANDO LOS DATOS AL RECORSET PARA QUE SE MUESTRE EN LA GRILLA
    Erase Arr_Totales_col_TMP
    Erase Arr_Totales_cuenta_TMP
    Set RST_ORIGEN = RST_TMP
'''    Set RST_GRUPO = Nothing
    Set RST_TMP = Nothing

End Sub


Private Function CARGAR_DATOS_GRILLA(RST_ORIGEN As ADODB.Recordset, _
                                         mAnyo As String, _
                                         Optional fEsVolumen As Boolean = False)
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    Dim BAND_ADD_REG As Boolean
    
    
    BAND_ADD_REG = True
    
    If RST_ORIGEN.RecordCount > 0 Then
        RST_ORIGEN.MoveFirst
    Else
        Exit Function
    End If
    
    PgBar(1).Min = 0
    PgBar(1).Max = RST_ORIGEN.RecordCount
    
    While Not RST_ORIGEN.EOF
    
    DoEvents
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then Exit Function
        '---------------------------------------------------------
        SGIFlex.ADD_REG Fg1
        '--ACUMULAR EN EL ARRAY_MES
        CARGAR_DATOS_ARRAY RST_ORIGEN, Arr_Totales_col(), Arr_Totales_cuenta(), fEsVolumen
        '--CARGAR A LA GRILLA
        CARGAR_DATOS_GRILLA_ARRAY_TMP RST_ORIGEN, mAnyo, Fg1.Rows - 1, , fEsVolumen
        '---------------------------------------------------------
        RST_ORIGEN.MoveNext
'        --PONER TOTALES AL FINAL DE LA GRILLA
        If Not RST_ORIGEN.EOF Then PgBar(1).Value = CLng(RST_ORIGEN.Bookmark)
           
    Wend
    PgBar(1).Value = 0
    
    '------

End Function

Private Sub CARGAR_DATOS_ARRAY(RST_ORIGEN As ADODB.Recordset, ARR_COL, ARR_CUENTA, _
                               Optional fEsVolumen As Boolean = False)
                               
    '--FUNCION QUE ACUMULARA EN EL ARRAY_TEMP
    Dim vStrCampo As String
    Dim Q_CAMPO As Integer
    Dim Q_POS As Integer
    Q_POS = 0
    '--ASIGNAR LOS DATOS AL RECORDSET TEMPORAL
    For Q_CAMPO = 0 To RST_ORIGEN.Fields.Count - 1
        '--SI SE NTERRUMPE EL PROCESO => SALIR
        If BAND_INTERRUMPIR = True Then Exit Sub
        vStrCampo = RST_ORIGEN.Fields(Q_CAMPO).Name
        '--OBS: SE VA LLENAR EL ARRAY "MONTOS DEL TOTAL" O "MONTOS DEL RESUMEN"
        Select Case LCase(vStrCampo)
            '--ACUMULANDO
            Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"
            '--ARR_TMP(0, 2) INDICA LA PRIMERA COLUMNA A MOSTRAR
                If LCase(vStrCampo) = ARR_TMP(0, 2) Then Q_POS = 0
                If fEsVolumen = False Then
                    If NulosN(RST_ORIGEN.Fields(vStrCampo)) <> 0 Then
                        If opt_precio(0).Value = True Then '--minimo
                            If NulosN(ARR_COL(Q_POS, 0)) = 0 Then ARR_COL(Q_POS, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))
                            If NulosN(ARR_COL(Q_POS, 0)) > NulosN(RST_ORIGEN.Fields(vStrCampo)) Then ARR_COL(Q_POS, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))

                        ElseIf opt_precio(1).Value = True Then '--prom
                            If NulosN(ARR_COL(Q_POS, 0)) = 0 Then
                                ARR_COL(Q_POS, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))
                            Else
                                ARR_COL(Q_POS, 0) = (NulosN(ARR_COL(Q_POS, 0)) + NulosN(RST_ORIGEN.Fields(vStrCampo))) / 2
                            End If
                        Else '--maximo
                            If NulosN(ARR_COL(Q_POS, 0)) = 0 Then ARR_COL(Q_POS, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))
                            If NulosN(ARR_COL(Q_POS, 0)) < NulosN(RST_ORIGEN.Fields(vStrCampo)) Then ARR_COL(Q_POS, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))
                        End If
'                        ARR_COL(Q_POS, 0) = ARR_COL(Q_POS, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
'                        If NulosN(RST_ORIGEN.Fields(vStrCampo)) <> 0 Then
'                            ARR_CUENTA(Q_POS, 0) = ARR_CUENTA(Q_POS, 0) + 1
'                        End If
                    
                    End If
                Else
                    ARR_COL(Q_POS, 1) = ARR_COL(Q_POS, 1) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                End If
                Q_POS = Q_POS + 1

            Case "total":
            
                If fEsVolumen = False Then
                    If opt_precio(0).Value = True Then '--minimo
                        If NulosN(ARR_COL(Q_COL_ARR_TOTAL + 1, 0)) = 0 Then ARR_COL(Q_COL_ARR_TOTAL + 1, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))
                        If NulosN(ARR_COL(Q_COL_ARR_TOTAL + 2, 0)) = 0 Then ARR_COL(Q_COL_ARR_TOTAL + 2, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))
                        
                        If NulosN(ARR_COL(Q_COL_ARR_TOTAL + 1, 0)) > NulosN(RST_ORIGEN.Fields(vStrCampo)) Then ARR_COL(Q_COL_ARR_TOTAL + 1, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))
                        If NulosN(ARR_COL(Q_COL_ARR_TOTAL + 2, 0)) > NulosN(RST_ORIGEN.Fields(vStrCampo)) Then ARR_COL(Q_COL_ARR_TOTAL + 2, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))

                    ElseIf opt_precio(1).Value = True Then '--prom
                        If NulosN(ARR_COL(Q_COL_ARR_TOTAL + 1, 0)) = 0 Then
                            ARR_COL(Q_COL_ARR_TOTAL + 1, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))
                            ARR_COL(Q_COL_ARR_TOTAL + 2, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))
                        Else
                            ARR_COL(Q_COL_ARR_TOTAL + 1, 0) = (NulosN(ARR_COL(Q_COL_ARR_TOTAL + 1, 0)) + NulosN(RST_ORIGEN.Fields(vStrCampo))) / 2
                            ARR_COL(Q_COL_ARR_TOTAL + 2, 0) = (NulosN(ARR_COL(Q_COL_ARR_TOTAL + 2, 0)) + NulosN(RST_ORIGEN.Fields(vStrCampo))) / 2
                        End If
                        
                    Else '--maximo
                        If NulosN(ARR_COL(Q_COL_ARR_TOTAL + 1, 0)) = 0 Then ARR_COL(Q_COL_ARR_TOTAL + 1, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))
                        If NulosN(ARR_COL(Q_COL_ARR_TOTAL + 2, 0)) = 0 Then ARR_COL(Q_COL_ARR_TOTAL + 2, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))
                        
                        If NulosN(ARR_COL(Q_COL_ARR_TOTAL + 1, 0)) < NulosN(RST_ORIGEN.Fields(vStrCampo)) Then ARR_COL(Q_COL_ARR_TOTAL + 1, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))
                        If NulosN(ARR_COL(Q_COL_ARR_TOTAL + 2, 0)) < NulosN(RST_ORIGEN.Fields(vStrCampo)) Then ARR_COL(Q_COL_ARR_TOTAL + 2, 0) = NulosN(RST_ORIGEN.Fields(vStrCampo))
                    End If
                    
'                    ARR_COL(Q_COL_ARR_TOTAL + 1, 0) = ARR_COL(Q_COL_ARR_TOTAL + 1, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
'                    ARR_COL(Q_COL_ARR_TOTAL + 2, 0) = ARR_COL(Q_COL_ARR_TOTAL + 2, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
'                     '--IDENTIFICAR LOS VALORES DIFERENTES A 0 PARA DESPUES HACER LA MEDIA
'                    If NulosN(RST_ORIGEN.Fields(vStrCampo)) <> 0 Then
'                        ARR_CUENTA(Q_COL_ARR_TOTAL + 1, 0) = ARR_CUENTA(Q_COL_ARR_TOTAL + 1, 0) + 1
'                        ARR_CUENTA(Q_COL_ARR_TOTAL + 2, 0) = ARR_CUENTA(Q_COL_ARR_TOTAL + 2, 0) + 1
'                    End If
                    

               Else
                    ARR_COL(Q_COL_ARR_TOTAL + 1, 1) = ARR_COL(Q_COL_ARR_TOTAL + 1, 1) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                    ARR_COL(Q_COL_ARR_TOTAL + 2, 1) = ARR_COL(Q_COL_ARR_TOTAL + 2, 1) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                End If
                

        End Select
    Next Q_CAMPO
    
End Sub


''''Private Sub CARGAR_DATOS_ARRAY(RST_ORIGEN As ADODB.Recordset)
''''    '--FUNCION QUE ACUMULARA EN EL ARRAY_TEMP
''''    Dim vStrCampo As String
''''    Dim Q_CAMPO As Integer
''''    Dim Q_POS As Integer
''''    Q_POS = 0
''''    '--ASIGNAR LOS DATOS AL RECORDSET TEMPORAL
''''    For Q_CAMPO = 0 To RST_ORIGEN.Fields.Count - 1
''''        '--SI SE NTERRUMPE EL PROCESO => SALIR
''''        If BAND_INTERRUMPIR = True Then Exit Sub
''''        vStrCampo = RST_ORIGEN.Fields(Q_CAMPO).Name
''''        '--OBS: SE VA LLENAR EL ARRAY "MONTOS DEL TOTAL" O "MONTOS DEL RESUMEN"
''''        Select Case LCase(vStrCampo)
''''            '--ACUMULANDO X MES
''''
''''            Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"
''''            '--ARR_TMP(0, 2) INDICA LA PRIMERA COLUMNA A MOSTRAR
''''                If LCase(vStrCampo) = ARR_TMP(0, 2) Then Q_POS = 0
''''                Arr_Totales_col(Q_POS, 0) = Arr_Totales_col(Q_POS, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
''''                If NulosN(RST_ORIGEN.Fields(vStrCampo)) <> 0 Then
''''                    Arr_Totales_cuenta(Q_POS, 0) = Arr_Totales_cuenta(Q_POS, 0) + 1
''''                End If
''''                Q_POS = Q_POS + 1
''''
''''            Case "total":
''''                Arr_Totales_col(Q_COL_ARR_TOTAL + 1, 0) = Arr_Totales_col(Q_COL_ARR_TOTAL + 1, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
''''                Arr_Totales_col(Q_COL_ARR_TOTAL + 2, 0) = Arr_Totales_col(Q_COL_ARR_TOTAL + 2, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
''''                '--IDENTIFICAR LOS VALORES DIFERENTES A 0 PARA DESPUES HACER LA MEDIA
''''                If NulosN(RST_ORIGEN.Fields(vStrCampo)) <> 0 Then
''''                    Arr_Totales_cuenta(Q_COL_ARR_TOTAL + 1, 0) = Arr_Totales_cuenta(Q_COL_ARR_TOTAL + 1, 0) + 1
''''                    Arr_Totales_cuenta(Q_COL_ARR_TOTAL + 2, 0) = Arr_Totales_cuenta(Q_COL_ARR_TOTAL + 2, 0) + 1
''''                End If
''''
''''        End Select
''''    Next Q_CAMPO
''''
''''End Sub

Private Function CARGAR_DATOS_GRILLA_ARRAY_TMP(RST_ORIGEN As ADODB.Recordset, _
                                        mAnyo As String, _
                                        Q_ROW As Integer, _
                                        Optional fOtrosAnyos As Boolean = False, _
                                        Optional fEsVolumen As Boolean = False)
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    
    Dim Q_INCREMENTO_X_COL As Integer   '--SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
    Dim Q_POS_MES_TOTAL  As Integer     '--POSICIONA A LA COLUMNA DEL TOTAL X AÑO
    Dim Q_POS As Integer
    Dim Q_CAMPO As Integer
    Dim vStrCampo As String
    
    
    For Q_POS = 0 To UBound(ARR_ANYO) - 1
        If ARR_ANYO(Q_POS) = mAnyo Then
            Q_INCREMENTO_X_COL = Q_POS
            Exit For
        End If
    Next
    '--IDENTIFICAR LA POSICION DE INICIO DE MES(ENERO)
    Q_INCREMENTO_X_COL = 0
    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    '-----------
    
    DoEvents

    
    For Q_CAMPO = 0 To RST_ORIGEN.Fields.Count - 1
        If BAND_INTERRUMPIR = True Then Exit Function
        vStrCampo = RST_ORIGEN.Fields(Q_CAMPO).Name
        If LCase(vStrCampo) = "ene" Then
            Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
        End If
                   
        '--COLOCANDO LOS VALORES EN LA GRILLA
        Select Case LCase(vStrCampo)
            '--DE LOS MESES
            Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"
                '"ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"
                '"ene-mar","abr-jun","jul-sep","oct-dic"
                '"1re sem","2do sem"
                
                '--ARR_TMP(0, 2) INDICA LA PRIMERA COLUMNA A MOSTRAR
                If LCase(vStrCampo) = ARR_TMP(0, 2) Then
                    'If fEsVolumen = True Then Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
                    'If fEsVolumen = False Then Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL + 1
                    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
                End If
    
                Fg1.TextMatrix(Q_ROW, Q_POS_MES) = PONER_FORMATO(NulosN(RST_ORIGEN.Fields(vStrCampo)), , Q_ROW, fEsVolumen)
                'If fEsVolumen = True Then Q_POS_MES = Q_POS_MES + 1
                'If fEsVolumen = False Then Q_POS_MES = Q_POS_MES + 2
                Q_POS_MES = Q_POS_MES + 1
             '--DEL TOTAL DEL AÑO
            Case "total"
                '---vol como nueva fila
                'If fEsVolumen = True Then Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL + 1) * 1
                'If fEsVolumen = False Then Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL + 2) * 1
                Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL + 1) * 1
                '--TOTAL AÑO
                Fg1.TextMatrix(Q_ROW, Q_POS_MES_TOTAL) = PONER_FORMATO(NulosN(RST_ORIGEN.Fields(vStrCampo)), , Q_COL_ARR_TOTAL + 1, fEsVolumen)
            '--DE LOS DEMAS CAMPOS
            Case Else
                '--SOLO SE AGREGARAN EN EL PRIMER AÑO
                If fOtrosAnyos = False Then Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = RST_ORIGEN.Fields(vStrCampo) & ""
        End Select
        '------------
    Next
End Function


Private Sub pExportar()
    On Error GoTo ERROR
    Dim X_EXPORT As New SGI2_funciones.formularios
    If F_ES_COMPRA = False Then T_RPT_TITULO = Replace(T_RPT_TITULO, "COMPRA", "VENTA")
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, T_RPT_TITULO, T_RPT_PERIODO, "", "Precios"
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
ERROR:
    Me.MousePointer = vbDefault
    SGIVarios.SHOW_ERROR Me.Name, "pExportar"
End Sub

Private Sub pImprimir()

    On Error GoTo ERROR
    Dim X_PRINT As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    If F_ES_COMPRA = False Then T_RPT_TITULO = Replace(T_RPT_TITULO, "COMPRA", "VENTA")
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, T_RPT_TITULO + " ", "ITEM: " + txt(2).Text, T_RPT_PERIODO + " ", False, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
ERROR:
    Me.MousePointer = vbDefault
    SGIVarios.SHOW_ERROR Me.Name, "pImprimir"
End Sub


Private Sub Fg1_DblClick()
    Fg1_KeyDown 13, 0
End Sub

Private Sub Fg1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> 13 Then Exit Sub
    If Fg1.Rows = 1 Then Exit Sub
    If Fg1.Row = 0 Or Fg1.Row = Fg1.Rows - 2 Or Fg1.Col <= 3 Or Fg1.Col = Fg1.Cols - 1 Then
        MsgBox "Selecione una Celda Correcta..", vbInformation, xTitulo
        Exit Sub
    End If
    If txt(5).Text = "" Or IsNumeric(txt(5).Text) = False Then
        MsgBox "Ingrese un número a mostrar", vbInformation, xTitulo
        txt(5).SetFocus
        Exit Sub
    End If
    If IsNumeric(Fg1.TextMatrix(Fg1.Row, Fg1.Col)) = False Then
        MsgBox "La celda no es numérico", vbInformation, xTitulo
        Exit Sub
    End If
    
    With FrmAnalizaPrecio_Item
        .RECIBE_ID_ITEM Fg1.TextMatrix(Fg1.Row, 2), Fg1.TextMatrix(0, Fg1.Col), ARR_TMP(), F_ES_COMPRA
        .Show 1
    End With
    
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And Shift = 0 Then PopupMenu Menu1
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
        '--interrumpir
    If KeyCode = vbKeyEscape And Shift = 0 Then BAND_INTERRUMPIR = True
    
End Sub

Private Sub Form_Load()
    On Error GoTo ERROR
    '--CARGAR DATOS
    SGIVarios.CentrarFrm Me
    SGIFlex.LimpiarGrid Me.Fg1
    
    SGIVarios.LimpiaText txt_cb
    SGIVarios.LimpiaText lbl_cb
    SGIVarios.LimpiaText lbl_cod
    
    '--CARGAR LOS AÑOS
    SGIVarios.CARGAR_LISTA_ANYOS_ACTIVOS ls(0), xCon
    SGIVarios.Llenar_Mes ls(1)
    '--CARGANDO LOS MESES
    SGIVarios.CARGAR_ARR_XX ARR_XX(), X_MES
    '--SELECCIONAR EL AÑO ACTUAL
    SGIVarios.ls_activar_chek ls(0), AnoTra
    SGIVarios.ls_activar_chek ls(1)
    '--CONFIGURAR LA GRILLA
    Validar_Consulta "-1"
    pGenerarConsulta "-1"
    pConfigurarGrilla
    Exit Sub
ERROR:
    SGIVarios.SHOW_ERROR
End Sub


Private Sub Form_Unload(Cancel As Integer)
    Erase ARR_TMP
    Erase Arr_Totales_cols
    Erase Arr_Totales_col
    
    Set SGIFlex = Nothing
    Set SGIVarios = Nothing
    
End Sub


Private Sub Menu1_1_Click()
    Fg1_DblClick
End Sub


Private Sub Menu1_3_Click()
    pConsultar
End Sub

Private Sub Menu1_5_Click()
    pExportar
End Sub

Private Sub Menu1_6_Click()
    pImprimir
End Sub

Private Sub opt_estilo_Click(index As Integer)
    Select Case index
        Case 0 '--MES
            SGIVarios.Llenar_Mes ls(1)
            SGIVarios.ls_activar_chek ls(1)
            SGIVarios.CARGAR_ARR_XX ARR_XX(), X_MES
        Case 1 '--TRIMESTRE
            SGIVarios.Llenar_Trimestre ls(1)
            SGIVarios.ls_activar_chek ls(1)
            SGIVarios.CARGAR_ARR_XX ARR_XX(), X_TRIMESTRE
        Case 2 '--SEMESTRE
            SGIVarios.Llenar_Semestre ls(1)
            SGIVarios.ls_activar_chek ls(1)
            SGIVarios.CARGAR_ARR_XX ARR_XX(), X_SEMESTRE
    End Select
    LBL(6).Caption = "Selecc. " + opt_estilo(index).Caption
    
End Sub

'------
Private Function Validar_Consulta(mAnyo As String) As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    '--DE LA FECHA ES NULL
    Dim k&
    mAnyo = ""
    Q_TOTAL_ANYO = 0
    '--RECORRER AÑO A AÑO PARA CARGAR LA DATA
    For k = ls(0).ListCount - 1 To 0 Step -1
        ls(0).ListIndex = k
        If ls(0).Selected(k) = True Then
            mAnyo = mAnyo + ls(0).Text + ","
            Q_TOTAL_ANYO = Q_TOTAL_ANYO + 1
        End If
    Next
    
    If mAnyo = "" Then
       MsgBox "Seleccione un Año como mínimo", vbCritical, xTitulo
'       ls(0).SetFocus
       Exit Function
    End If
    Erase ARR_ANYO '--LIMPIAR ARRAY
    ARR_ANYO = Split(mAnyo, ",") '--ASIGNANDO EL LISTADO DE AÑOS
    
    
    '----------------
    Q_COL_ARR_TOTAL = 0
    For k = ls(1).ListCount - 1 To 0 Step -1
        ls(1).ListIndex = k
        If ls(1).Selected(k) = True Then
            Q_COL_ARR_TOTAL = Q_COL_ARR_TOTAL + 1
        End If
    Next
    If Q_COL_ARR_TOTAL = 0 Then
       MsgBox Replace(LBL(6).Caption, "Selecc.", "Selecc. un ") + " como mínimo...", vbCritical, xTitulo
       ls(1).SetFocus
       Exit Function
    End If
    Q_COL_ARR_TOTAL = Q_COL_ARR_TOTAL - 1
    '-----------
    Erase ARR_TMP
    ReDim ARR_TMP(Q_COL_ARR_TOTAL, 2)
    Dim POS_ARR As Integer
    POS_ARR = 0
    For k = 0 To ls(1).ListCount - 1
        ls(1).ListIndex = k
        If ls(1).Selected(k) = True Then
            ARR_TMP(POS_ARR, 0) = ARR_XX(k, 0)
            ARR_TMP(POS_ARR, 1) = ARR_XX(k, 1)
            ARR_TMP(POS_ARR, 2) = ARR_XX(k, 2)
            POS_ARR = POS_ARR + 1
        End If
    Next
    '-----------
    Validar_Consulta = True

End Function

Private Function pGenerarConsulta(mAnyo As String, Optional fEsVolumen As Boolean = False) As String
    '--FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
    '--
    Dim vStrSelect As String            '--CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    Dim vStrFiltro_ITEM As String       '--SOLO ITEM
    
    Dim vStrFiltro As String

    Dim k As Integer
    '--DEL AÑO
    vStrFiltro = " Year(com_compras.fchdoc)= " + mAnyo + " "
    '--DEL ITEM
    vStrFiltro_ITEM = " AND com_comprasdet.iditem= " + Trim(txt(0).Text) + " "
    '--
    '--SOLO s/.
    If opt_mon(0).Value = True Then vStrFiltro = vStrFiltro + " AND com_compras.idmon= 1 " '--SOLES
    '--SOLO $
    If opt_mon(1).Value = True Then vStrFiltro = vStrFiltro + " AND com_compras.idmon= 2 " '--DOLARES
    
    If NulosN(lbl_cod(0).Caption) <> 0 Then vStrFiltro = vStrFiltro + " AND com_compras.idpro=  " & NulosN(lbl_cod(0).Caption) & " "  '--PROVEEDOR/CLIENTE
    
    
    '--GENERAR LA CONSULTA SEGUN CONDICIONES
    Dim nSQLValor As String
    Dim nSQLCampos As String
    Dim nSQLWhere As String
    Dim nSQLFrom As String
    Dim nSQLGroupBy As String
    Dim nSQLOrderBy As String
    Dim nSQLPivot As String
    Dim nSQLPivot1 As String
    Dim nSQLPivotSalida As String '--ORDENA LOS VALORE MES A MES(ENE,FEB,MAR,ETC.)
    Dim nCampoTmp As String
    If fEsVolumen = True Then
        nCampoTmp = "'Vol'"
    Else
        If opt_precio(0).Value = True Then
            nCampoTmp = "'Prec.Min'"
        ElseIf opt_precio(1).Value = True Then
            nCampoTmp = "'Prec.Prom'"
        Else
            nCampoTmp = "'Prec.Max'"
        End If
        opt_precio(0).Tag = Replace(nCampoTmp, "'", "")
    End If
    
    nSQLWhere = vStrFiltro + vStrFiltro_ITEM

    Q_COL_FILA_OCULTA = 1:        Q_COL_FILA = 3:       Q_POSICION_TOTAL = 3:           Q_COL_COMPARAR_GRUPO = -1
    
    'nSQLCampos = "YEAR(com_compras.fchdoc) AS idanyo,YEAR(com_compras.fchdoc) AS anyo "
    'nSQLGroupBy = "YEAR(com_compras.fchdoc) ,com_comprasdet.preuni "
    'nSQLOrderBy = "YEAR(com_compras.fchdoc) "
    
    
    nSQLCampos = "  VW.idanyo, VW.anyo, " & nCampoTmp & " as tipo "
    nSQLGroupBy = " VW.idanyo, VW.anyo "
    
    '--DE LA CANTIDAD DE COL
    Q_COL_FILA = Q_COL_FILA + 1
    Q_POS_MES_INICIO = Q_COL_FILA '--Q_COL_FILA + CAMPO_TOTAL
    '------------------------------------------
    If opt_estilo(0).Value = True Then '--MES
        T_RPT_TITULO = "ANÁLISIS DE PRECIO DE COMPRAS MENSUAL"
        nSQLPivot = "FORMAT(com_compras.fchdoc,'m') "
        nSQLPivot1 = "FORMAT(vw.fchdoc,'m') "
    ElseIf opt_estilo(1).Value = True Then '--TRIMESTRE
        T_RPT_TITULO = "ANÁLISIS DE PRECIO DE COMPRAS TRIMESTRAL"
        nSQLPivot = "FORMAT(com_compras.fchdoc,'q') "
        nSQLPivot1 = "FORMAT(vw.fchdoc,'q') "
    ElseIf opt_estilo(2).Value = True Then '--SEMESTRE
        T_RPT_TITULO = "ANÁLISIS DE PRECIO DE COMPRAS SEMESTRAL"
        nSQLPivot = "FORMAT(com_compras.fchdoc,'s') "
        nSQLPivot1 = "FORMAT(vw.fchdoc,'s') "
    End If
    '--DEL PIVOT
    For k = 0 To UBound(ARR_TMP)
        nSQLPivotSalida = nSQLPivotSalida + ARR_TMP(k, 2) + ","
    Next k
    nSQLPivotSalida = " IN (" + Left(nSQLPivotSalida, Len(nSQLPivotSalida) - 1) + ") "
    nSQLWhere = nSQLWhere + " AND " + nSQLPivot + nSQLPivotSalida
    'nSQLPivotSalida = " In ('Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic');"
    
    'nSQLFrom = " (con_tc RIGHT JOIN com_compras ON con_tc.fecha = com_compras.fchdoc) INNER JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom "
    nSQLFrom = " (SELECT Year(com_compras.fchdoc) AS idanyo, Year(com_compras.fchdoc) AS anyo, com_compras.fchdoc, mae_moneda.simbolo, mae_unidades.abrev, con_tc.impcom, " _
                    + vbCr + " com_comprasdet.canpro, com_comprasdet.preuni, " _
                    + vbCr + " IIf([com_compras].[idmon]=1,[com_comprasdet].[preuni],IIf([con_tc].[impcom] Is Null,0,[com_comprasdet].[preuni]*[con_tc].[impcom])) AS preunisol, " _
                    + vbCr + " IIf([com_compras].[idmon]=2,[com_comprasdet].[preuni],IIf([con_tc].[impcom] Is Null,0,[com_comprasdet].[preuni])/[con_tc].[impcom]) AS preunidol " _
                    + vbCr + " FROM mae_moneda RIGHT JOIN (mae_unidades RIGHT JOIN ((con_tc RIGHT JOIN com_compras ON con_tc.fecha = com_compras.fchdoc) INNER JOIN com_comprasdet ON com_compras.id = com_comprasdet.idcom) ON mae_unidades.id = com_comprasdet.idunimed) ON mae_moneda.id = com_compras.idmon " _
                    + vbCr + " WHERE " & nSQLWhere _
                    + vbCr + " ORDER BY Year(com_compras.fchdoc), com_compras.fchdoc " _
                + vbCr + " ) AS VW"
    
'    nSQLValor = "AVG(com_comprasdet.preuni) "
'    If Me.opt_mon(2).Value = True Then '--TODO EN SOLES
'        nSQLValor = " AVG(IIF(com_compras.idmon=1,com_comprasdet.preuni,IIF(con_tc.impcom IS NULL,0,con_tc.impcom*com_comprasdet.preuni))) "
'    ElseIf Me.opt_mon(3).Value = True Then '--TODO EN DOLARES
'        nSQLValor = " AVG(IIF(com_compras.idmon=2,com_comprasdet.preuni,IIF(con_tc.impcom IS NULL,0,con_tc.impcom/com_comprasdet.preuni))) "
'    End If
   
    If fEsVolumen = True Then
        nSQLValor = " SUM(VW.canpro) " '--volumen
    Else
        If opt_mon(0).Value = True Or opt_mon(1).Value = True Then
            nSQLValor = " Avg(VW.preuni) " '--soles o dolares
        ElseIf opt_mon(2).Value = True Then
            nSQLValor = " Avg(VW.preunisol) " '--todo en soles
        Else
            nSQLValor = " Avg(VW.preunidol) " '--todo en dolares
        End If
        
        If opt_precio(0).Value = True Then
            nSQLValor = Replace(nSQLValor, "Avg", "Min")
        ElseIf opt_precio(2).Value = True Then
            nSQLValor = Replace(nSQLValor, "Avg", "Max")
        End If
        
    End If
    
    '--GENERANDO LA CONSULTA
'    nSQLCampos = nSQLCampos + "," + nSQLValor + " AS total "
    vStrSelect = " TRANSFORM " + nSQLValor + _
        vbCr + " SELECT " + nSQLCampos + "," + nSQLValor + " AS total " + _
        vbCr + " FROM " + nSQLFrom + _
        vbCr + " GROUP BY " + nSQLGroupBy + _
        vbCr + " PIVOT " + nSQLPivot1 + nSQLPivotSalida
    '--SI ES POR VENTA
    If F_ES_COMPRA = False Then
        vStrSelect = Replace(vStrSelect, "com_comprasdet", "vta_ventasdet")
        vStrSelect = Replace(vStrSelect, "com_compras", "vta_ventas")
        vStrSelect = Replace(vStrSelect, ".idcom", ".idvta")
        vStrSelect = Replace(vStrSelect, ".idpro", ".idcli")
        vStrSelect = Replace(vStrSelect, "WHERE ", "WHERE vta_ventas.anulado=0 AND ")
        
    End If
    '------------------------------------------------------------------------------------
    pGenerarConsulta = vStrSelect
    
    
End Function


Private Sub Limpiar_ARRAY_TOTAL(Optional F_LIMPIA_TOT_GRL As Boolean = False)
    Erase Arr_Totales_col
    ReDim Arr_Totales_col(13, 1) As Double
    Erase Arr_Totales_cuenta
    ReDim Arr_Totales_cuenta(13, 1) As Double
    If F_LIMPIA_TOT_GRL = True Then
        Erase Arr_Totales_cols
        ReDim Arr_Totales_cols(13, 1)
    End If
End Sub
'''
Private Sub CARGAR_DATOS_GRILLA_ADD_TOTALES(BAND_ADD_TOTAL As Boolean, _
                                            Nombre_total As String, _
                                            Optional fTotalGral As Boolean = False, _
                                            Optional fForzarSuma As Boolean = False, _
                                            Optional mAnyo As String, _
                                            Optional fOtrosAnyos As Boolean = False, _
                                            Optional fEsVolumen As Boolean = False)
                
    Dim Q_MES As Integer
    '--AGREGA LOS TOTALES POR CADA GRUPO Y EL TOTAL GENERAL
    '--ACUMULA LOS TOTALES EN EL TOTAL GENERAL
    Dim X_ROW As Long
    'On Error Resume Next
    If fOtrosAnyos = False Then
        X_ROW = Fg1.Rows
        If BAND_ADD_TOTAL = True Then
            '--AGREAGNDO NUEVA FILA
            SGIFlex.ADD_REG Fg1, IIf(fTotalGral = False, Fila_Total, Fila_Total_grl)
    
            'PONIENDO LOS NOMBRES DE LOS TOTALES  Q_POSICION_TOTAL SE OBTIENE DE pGenerarConsulta()
            Fg1.TextMatrix(X_ROW, Q_POSICION_TOTAL) = Nombre_total
            SGIFlex.FORMATO_CELDA Fg1, X_ROW, Q_POSICION_TOTAL
        End If
    Else
        X_ROW = Fg1.Row
    End If

    '--ACUMULANDO LOS TOTALES GRLES
    If fTotalGral = True Then
        For Q_MES = 0 To UBound(Arr_Totales_col())
            If fEsVolumen = False Then
                 Arr_Totales_cols(Q_MES, 0) = Arr_Totales_cols(Q_MES, 0) + Arr_Totales_col(Q_MES, 0)
            Else
                Arr_Totales_cols(Q_MES, 1) = Arr_Totales_cols(Q_MES, 1) + Arr_Totales_col(Q_MES, 1)
            End If
        Next Q_MES
    End If

    '
'--------------------------
    Dim Q_INCREMENTO_X_COL As Integer   '--SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
    Dim Q_POS_MES_TOTAL  As Integer     '--POSICIONA A LA COLUMNA DEL TOTAL X AÑO
    Dim Q_POS_TOTAL As Integer
    
    Q_POS_TOTAL = IIf(fEsVolumen = False, 0, 1)
    
    For Q_MES = 0 To UBound(ARR_ANYO) - 1
        If ARR_ANYO(Q_MES) = mAnyo Then
            Q_INCREMENTO_X_COL = Q_MES
            Exit For
        End If
    Next
    '--IDENTIFICAR LA POSICION DE INICIO DE MES(ENERO)
    Q_INCREMENTO_X_COL = 0
    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    '-----------
'--DE LOS MESES
    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    
    '--DE LOS MESES
    For Q_MES = 0 To Q_COL_ARR_TOTAL
        '--INTERRUMPIR EL PROCESO
        If BAND_INTERRUMPIR = True Then Exit Sub
        Fg1.TextMatrix(X_ROW, Q_POS_MES) = PONER_FORMATO(IIf(fTotalGral = False, Arr_Totales_col(Q_MES, Q_POS_TOTAL), Arr_Totales_cols(Q_MES, Q_POS_TOTAL)), fTotalGral, Q_MES, fEsVolumen)
        SGIFlex.FORMATO_CELDA Fg1, X_ROW, Q_POS_MES
        Q_POS_MES = Q_POS_MES + 1
    Next Q_MES
       
    For Q_MES = Q_COL_ARR_TOTAL + 1 To Q_COL_ARR_TOTAL + 2
        If Q_MES = Q_COL_ARR_TOTAL + 1 Then '--TOTAL
                Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL + 1) * 1
            Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL) = PONER_FORMATO(IIf(fTotalGral = False, Arr_Totales_col(Q_MES, Q_POS_TOTAL), Arr_Totales_cols(Q_MES, Q_POS_TOTAL)), fTotalGral, Q_MES, fEsVolumen)
        ElseIf Q_MES = Q_COL_ARR_TOTAL + 2 Then '--TOTAL GRAL
            Q_POS_MES_TOTAL = Fg1.Cols - 1
            If IsNumeric(Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL)) = False Then
                Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL) = IIf(fTotalGral = False, Arr_Totales_col(Q_MES, Q_POS_TOTAL), Arr_Totales_cols(Q_MES, Q_POS_TOTAL))
            Else
                    Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL) = IIf(fTotalGral = False, Arr_Totales_col(Q_MES, Q_POS_TOTAL), Arr_Totales_cols(Q_MES, Q_POS_TOTAL))
            End If
            Fg1.TextMatrix(X_ROW, Fg1.Cols - 1) = PONER_FORMATO(CDbl(Fg1.TextMatrix(X_ROW, Fg1.Cols - 1)), fTotalGral, Q_MES, fEsVolumen)
        End If
        
        SGIFlex.FORMATO_CELDA Fg1, X_ROW, Q_POS_MES_TOTAL
        
    Next Q_MES

    Err.Clear
End Sub



Private Sub pAcumularDatos()

End Sub


Private Sub pConfigurarGrilla(Optional F_CONSERVAR_FORMATO As Boolean = False)
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    '--DE ACUERDO A LO QUE SE SELECCIONA
    Dim M_ANCHO_COL_MES As Integer '--DEPENDERA DEL TIPO DE PRESENTACION
                                    '--EN DECIMALES, EN MILES
    Dim k&, j&
    
    If F_CONSERVAR_FORMATO = True Then Fg1.Clear
    
    Fg1.FrozenCols = 0
    
    If opt_estilo(0).Value = True Then '--MES
        M_ANCHO_COL_MES = 750
    ElseIf opt_estilo(1).Value = True Then '--TRIMESTRE
        M_ANCHO_COL_MES = 1100
    ElseIf opt_estilo(2).Value = True Then '--SEMESTRE
        M_ANCHO_COL_MES = 1600
    End If
    
    If opt_estilo(0).Value = True Then
        M_ANCHO_COL_MES = M_ANCHO_COL_MES + 250
    Else
        M_ANCHO_COL_MES = M_ANCHO_COL_MES
    End If
    
    With Fg1
        .Rows = 1
        .FixedRows = 1
        '--DATOS DE FILA
        Fg1.Cols = Q_COL_FILA + ((Q_COL_ARR_TOTAL + 2) * 1)
'        UNIR_CELDAS Fg1, 0, Q_COL_FILA, 0, Fg1.Cols - 1, " ", flexAlignCenterTop
                 
        Q_POS_MES = Q_POS_MES_INICIO
        '--DATOS DE COLUMNAS
        For k = 0 To Q_COL_ARR_TOTAL '--MESES DEL AÑO
            '--COLOCANDO LOS MESES
            SGIFlex.UNIR_CELDAS Fg1, 0, Q_POS_MES, 0, Q_POS_MES, ARR_TMP(k, 1), flexAlignCenterTop: .ColWidth(Q_POS_MES) = M_ANCHO_COL_MES
            
'''            '--volumen
'''            UNIR_CELDAS Fg1, 1, Q_POS_MES, 1, Q_POS_MES, "Vol.", flexAlignCenterTop: .ColWidth(Q_POS_MES) = M_ANCHO_COL_MES + 200
'''            '--promedio/min/max
'''            UNIR_CELDAS Fg1, 1, Q_POS_MES + 1, 1, Q_POS_MES + 1, "Prom.", flexAlignCenterTop: .ColWidth(Q_POS_MES + 1) = M_ANCHO_COL_MES
            
            Q_POS_MES = Q_POS_MES + 1
        Next k
        '--COLOCANDO EL TOTAL
        
        .TextMatrix(0, .Cols - 1) = "Totales":         .ColWidth(.Cols - 1) = M_ANCHO_COL_MES + 100
        
        .FrozenCols = Q_POS_MES_INICIO - 1
        .ColWidth(0) = 200
        '--DATOS DE FILA
        .ColAlignment(2) = flexAlignLeftCenter
        .ColAlignment(3) = flexAlignLeftCenter
        .ColAlignment(.Cols - 1) = flexAlignRightCenter
        SGIFlex.UNIR_CELDAS Fg1, 0, 2, 0, 2, "Año", flexAlignCenterTop, False
        .ColWidth(2) = M_ANCHO_COL_MES - 500
        .TextMatrix(0, 3) = "Tipo":         .ColWidth(3) = M_ANCHO_COL_MES

        '--DE LOS ID'S
        For k = 1 To Q_COL_FILA_OCULTA
            .TextMatrix(0, k) = "ID" + CStr(k):         .ColWidth(k) = 500
        Next
        If Q_COL_FILA_OCULTA <> -1 Then SGIFlex.OCULTAR_COL Fg1, 1, Q_COL_FILA_OCULTA
        
    End With
    DoEvents
    
End Sub

Private Sub PosicionarProgBar()
'--POSICIONAR EL PROGRESO DENTRO DEL FORMULARIO
'    FraProgreso.Width = 5820
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    FraProgreso.Visible = True
End Sub


Private Function PONER_FORMATO(S_MONTO As Double, _
                        Optional fTotalGral As Boolean = False, _
                        Optional Q_POS As Integer = -1, _
                        Optional fEsVolumen As Boolean = False) As String
                        
    '--ESTA FUNCION CONVERTIRA AL FORMATO
    If S_MONTO = 0 Then
            If fEsVolumen = False Then PONER_FORMATO = "0.0000"
            If fEsVolumen = True Then PONER_FORMATO = "0.00"
        Exit Function
    End If
    
    If fTotalGral = False Then
        If fEsVolumen = False Then PONER_FORMATO = Format(S_MONTO, SGIFlex.FORMAT_MEDIA)
        If fEsVolumen = True Then PONER_FORMATO = Format(S_MONTO, SGIFlex.FORMAT_MONTO)
    Else
'        If Q_POS <> -1 Then
'            If Arr_Totales_cuenta(Q_POS, 0) = 0 Then
'                If fEsVolumen = False Then PONER_FORMATO = "0.0000"
'                If fEsVolumen = True Then PONER_FORMATO = "0.00"
'                Exit Function
'            End If
'        End If
        'If fEsVolumen = False Then PONER_FORMATO = Format(S_MONTO / Arr_Totales_cuenta(Q_POS, 0), FORMAT_MEDIA)
        If fEsVolumen = False Then PONER_FORMATO = Format(S_MONTO, SGIFlex.FORMAT_MEDIA)
        If fEsVolumen = True Then PONER_FORMATO = Format(S_MONTO, SGIFlex.FORMAT_MONTO)
    End If
    
    
    
End Function

Private Sub txt_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If index <> 5 Then Exit Sub
    If SGIVarios.validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.index = 1 Then pConsultar
    If Button.index = 3 Then pExportar
    If Button.index = 4 Then pImprimir
    If Button.index = 6 Then
        BAND_INTERRUMPIR = True
        Unload Me
    End If
End Sub



'*******************************************************************************************

Private Sub cb_Click(index As Integer)
    Dim xCampos() As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nTitulo As String
    On Error GoTo ERROR
    
    If F_ES_COMPRA = True Then
            
            nSQL = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id FROM mae_prov WHERE (((mae_prov.activo)=-1)) ORDER BY mae_prov.nombre; "
            nTitulo = "Buscando Proveedores"
            
    Else
            nSQL = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id FROM mae_cliente WHERE (((mae_cliente.activo)=-1)) ORDER BY mae_cliente.nombre; "
            nTitulo = "Buscando Clientes"
    End If
    
    ReDim xCampos(3, 3) As String
    xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "R.U.C.":   xCampos(1, 1) = "numruc":    xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "Codigo":   xCampos(2, 1) = "id":        xCampos(2, 2) = "800":    xCampos(2, 3) = "N"
    
    Dim RstTmp As New ADODB.Recordset
    SGIVarios.CARGAR_DLL_EPSBUSCAR xCon, RstTmp, nSQL, xCampos(), nTitulo, "nombre", "nombre", Principio

    If RstTmp.State = 0 Then GoTo salir
    If RstTmp.EOF = True Or RstTmp.BOF = True Or RstTmp.RecordCount = 0 Then GoTo salir

    lbl_cod(index).Tag = lbl_cod(index).Caption

    txt_cb(index).Text = NulosC(RstTmp.Fields(0))  '--TEXTO A MOSTRAR
    lbl_cb(index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
    lbl_cod(index).Caption = NulosN(RstTmp.Fields(2)) '--CODIGO
    lbl_cb(index).ToolTipText = NulosC(RstTmp.Fields(1))  '--NOMBRE
      
salir:
    Set RstTmp = Nothing
Exit Sub
ERROR:
    Set RstTmp = Nothing
    SGIVarios.SHOW_ERROR Me.Name, "cb_Click(" + CStr(index) + ")"
End Sub


Private Sub txt_cb_Change(index As Integer)
    If txt_cb(index).Text = "" Then
        Me.lbl_cb(index).Caption = ""
        Me.lbl_cod(index).Caption = ""
    End If
End Sub

Private Sub txt_cb_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If txt_cb(index).Locked = True Then Exit Sub
    If KeyCode = vbKeyF5 Then
        cb_Click index
        Exit Sub
    End If
End Sub

Private Sub txt_cb_KeyPress(index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        If index <> 1 Then
            SendKeys vbTab
        Else
            If Fg1.Rows >= 2 Then
                Fg1.Row = 1: Fg1.Col = 1
            Else
                Fg1.Row = Fg1.Rows - 1: Fg1.Col = 1
            End If
            Fg1.SetFocus
        End If
        Exit Sub
    End If
    If SGIVarios.validar_numero(KeyAscii) = False Then KeyAscii = 0
End Sub

Private Sub txt_cb_Validate(index As Integer, Cancel As Boolean)

    If txt_cb(index).Text = "" Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    Dim nSQL As String
    On Error GoTo ERROR
    
    If F_ES_COMPRA = True Then
            nSQL = "SELECT mae_prov.numruc, mae_prov.nombre, mae_prov.id FROM mae_prov WHERE (((mae_prov.activo)=-1)) and mae_prov.numruc ='" & NulosC(txt_cb(0).Text) & "' ORDER BY mae_prov.nombre; "
           
    Else
            nSQL = "SELECT mae_cliente.numruc, mae_cliente.nombre, mae_cliente.id FROM mae_cliente WHERE (((mae_cliente.activo)=-1))  and mae_cliente.numruc ='" & NulosC(txt_cb(0).Text) & "'  ORDER BY mae_cliente.nombre; "
    End If
    

    If xCon.State = 0 Then GoTo salir
    RST_Busq RstTmp, nSQL, xCon

    If RstTmp.State = 0 Then GoTo salir

    lbl_cod(index).Tag = lbl_cod(index).Caption

    If RstTmp.RecordCount > 0 Then
        txt_cb(index).Text = NulosC(RstTmp.Fields(0))   '--TEXTO A MOSTRAR
        lbl_cb(index).Caption = NulosC(RstTmp.Fields(1)) '--NOMBRE
        lbl_cod(index).Caption = NulosN(RstTmp.Fields(2)) '--CODIGO
        lbl_cb(index).ToolTipText = NulosC(RstTmp.Fields(1)) '--NOMBRE
    Else
        txt_cb(index).Text = "":    lbl_cb(index).Caption = "":    lbl_cod(index).Caption = ""
    End If
    
    Set RstTmp = Nothing
    Exit Sub
ERROR:
    Set RstTmp = Nothing
    SGIVarios.SHOW_ERROR Me.Name, "txt_cb_Validate (" + CStr(index) + ")"
    Exit Sub
salir:
    Set RstTmp = Nothing
    txt_cb(index).Text = ""
End Sub

'****************************************************************************************

