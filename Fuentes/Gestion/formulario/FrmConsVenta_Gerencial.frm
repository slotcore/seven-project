VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsVenta_Gerencial 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Gestión - Análisis de Ventas"
   ClientHeight    =   8010
   ClientLeft      =   105
   ClientTop       =   600
   ClientWidth     =   11805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8010
   ScaleWidth      =   11805
   Begin VB.Frame FraGraf1 
      Height          =   2385
      Left            =   4320
      TabIndex        =   48
      Top             =   3105
      Visible         =   0   'False
      Width           =   3525
      Begin VB.CommandButton CmdGrafCancel1 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   1800
         TabIndex        =   60
         Top             =   1950
         Width           =   1560
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mostrar"
         Height          =   765
         Left            =   180
         TabIndex        =   57
         Top             =   1500
         Width           =   1515
         Begin VB.CheckBox ChkLeyenda 
            Caption         =   "Leyenda"
            Height          =   195
            Left            =   210
            TabIndex        =   58
            Top             =   300
            Width           =   1005
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Con Datos"
         Height          =   1110
         Left            =   180
         TabIndex        =   54
         Top             =   360
         Width           =   1515
         Begin VB.OptionButton OptconDatosDetalle1 
            Caption         =   "Detallado"
            Height          =   210
            Left            =   165
            TabIndex        =   56
            Top             =   645
            Width           =   1035
         End
         Begin VB.OptionButton OptConDatoResum1 
            Caption         =   "Resumido"
            Height          =   195
            Left            =   165
            TabIndex        =   55
            Top             =   315
            Value           =   -1  'True
            Width           =   1005
         End
      End
      Begin VB.CommandButton CmdGrafAcep1 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   1800
         TabIndex        =   53
         Top             =   1530
         Width           =   1560
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Gráfico"
         Height          =   1110
         Left            =   1800
         TabIndex        =   49
         Top             =   360
         Width           =   1560
         Begin VB.OptionButton OptTipGrafCircular 
            Caption         =   "Circular"
            Height          =   195
            Left            =   165
            TabIndex        =   52
            Top             =   795
            Width           =   1290
         End
         Begin VB.OptionButton OptTipGrafLinea 
            Caption         =   "Lineas"
            Height          =   195
            Left            =   165
            TabIndex        =   51
            Top             =   547
            Width           =   1290
         End
         Begin VB.OptionButton OptTipGrafBarra1 
            Caption         =   "Barras"
            Height          =   195
            Left            =   165
            TabIndex        =   50
            Top             =   300
            Value           =   -1  'True
            Width           =   1290
         End
      End
      Begin VB.Label Label4 
         BackColor       =   &H00800000&
         Caption         =   "  Propiedades de gráfico"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   300
         Left            =   0
         TabIndex        =   59
         Top             =   0
         Width           =   3885
      End
   End
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   3375
      TabIndex        =   8
      Top             =   3615
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   9
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
         TabIndex        =   35
         Top             =   795
         Width           =   4755
         _ExtentX        =   8387
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Label lbl 
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
         TabIndex        =   37
         Top             =   840
         Width           =   495
      End
      Begin VB.Label lbl 
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
         TabIndex        =   36
         Top             =   495
         Width           =   825
      End
      Begin VB.Label lbl 
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
         TabIndex        =   13
         Top             =   150
         Width           =   1530
      End
      Begin VB.Label lbl 
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
         TabIndex        =   11
         Top             =   150
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Ventas"
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
         TabIndex        =   10
         Top             =   150
         Width           =   585
      End
      Begin VB.Shape Shape1 
         Height          =   1065
         Left            =   90
         Top             =   60
         Width           =   5805
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Align           =   2  'Align Bottom
      Height          =   5070
      Left            =   0
      TabIndex        =   12
      Top             =   2940
      Width           =   11805
      _cx             =   20823
      _cy             =   8943
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
      AllowUserResizing=   1
      SelectionMode   =   1
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
      FormatString    =   $"FrmConsVenta_Gerencial.frx":0000
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
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   61
      Top             =   0
      Width           =   11805
      _ExtentX        =   20823
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar a MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Ver Gráfico"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   4860
         Top             =   90
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
               Picture         =   "FrmConsVenta_Gerencial.frx":003C
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Gerencial.frx":0580
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Gerencial.frx":0912
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Gerencial.frx":0A96
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Gerencial.frx":0EEA
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Gerencial.frx":1002
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Gerencial.frx":1546
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Gerencial.frx":1A8A
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Gerencial.frx":1B9E
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Gerencial.frx":1CB2
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Gerencial.frx":2106
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Gerencial.frx":2272
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Gerencial.frx":27BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Gerencial.frx":2AD4
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta_Gerencial.frx":2E66
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fr 
      Height          =   2595
      Index           =   5
      Left            =   0
      TabIndex        =   2
      Top             =   315
      Width           =   11805
      Begin VB.Frame fr 
         Caption         =   "Seleccionar Importe"
         Height          =   645
         Index           =   6
         Left            =   8685
         TabIndex        =   44
         Top             =   120
         Width           =   3075
         Begin VB.OptionButton opt_importe 
            Caption         =   "Sólo Igv"
            Height          =   195
            Index           =   2
            Left            =   2040
            TabIndex        =   47
            Top             =   225
            Width           =   885
         End
         Begin VB.OptionButton opt_importe 
            Caption         =   "Sin Igv"
            Height          =   195
            Index           =   1
            Left            =   1125
            TabIndex        =   46
            Top             =   225
            Width           =   825
         End
         Begin VB.OptionButton opt_importe 
            Caption         =   "Con Igv"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   45
            Top             =   225
            Value           =   -1  'True
            Width           =   900
         End
      End
      Begin VB.ListBox ls 
         Height          =   960
         Index           =   1
         Left            =   4050
         Style           =   1  'Checkbox
         TabIndex        =   41
         Top             =   405
         Width           =   1530
      End
      Begin VB.Frame fr 
         Caption         =   "Seleccionar"
         Height          =   1260
         Index           =   2
         Left            =   2880
         TabIndex        =   38
         Top             =   150
         Width           =   1095
         Begin VB.OptionButton opt_estilo 
            Caption         =   "Trimestre"
            Height          =   210
            Index           =   1
            Left            =   60
            TabIndex        =   43
            Top             =   502
            Width           =   960
         End
         Begin VB.OptionButton opt_estilo 
            Caption         =   "Mes"
            Height          =   210
            Index           =   0
            Left            =   60
            TabIndex        =   40
            Top             =   270
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton opt_estilo 
            Caption         =   "Semestre"
            Height          =   210
            Index           =   2
            Left            =   60
            TabIndex        =   39
            Top             =   735
            Visible         =   0   'False
            Width           =   960
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Tipo de Consulta"
         Height          =   1260
         Index           =   1
         Left            =   30
         TabIndex        =   29
         Top             =   120
         Width           =   1500
         Begin VB.OptionButton opt_consulta 
            Caption         =   "x T. Prod/Item"
            Height          =   195
            Index           =   4
            Left            =   45
            TabIndex        =   34
            Top             =   975
            Width           =   1380
         End
         Begin VB.OptionButton opt_consulta 
            Caption         =   "x Vendedor"
            Height          =   195
            Index           =   3
            Left            =   45
            TabIndex        =   33
            Top             =   780
            Width           =   1380
         End
         Begin VB.OptionButton opt_consulta 
            Caption         =   "x Pto de Venta"
            Height          =   195
            Index           =   2
            Left            =   45
            TabIndex        =   32
            Top             =   585
            Width           =   1380
         End
         Begin VB.OptionButton opt_consulta 
            Caption         =   "x Año"
            Height          =   195
            Index           =   0
            Left            =   45
            TabIndex        =   31
            Top             =   195
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton opt_consulta 
            Caption         =   "x Cliente"
            Height          =   195
            Index           =   1
            Left            =   45
            TabIndex        =   30
            Top             =   390
            Width           =   1380
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Presentación"
         Height          =   645
         Index           =   3
         Left            =   7140
         TabIndex        =   26
         Top             =   120
         Width           =   1485
         Begin VB.OptionButton opt_escala 
            Caption         =   "En Decimales"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   28
            Top             =   195
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton opt_escala 
            Caption         =   "En Miles"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   390
            Width           =   1275
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Seleccionar"
         Height          =   630
         Index           =   4
         Left            =   7140
         TabIndex        =   23
         Top             =   750
         Width           =   1485
         Begin VB.OptionButton opt_totalizar 
            Caption         =   "Cantidades"
            Height          =   195
            Index           =   1
            Left            =   120
            TabIndex        =   25
            Top             =   390
            Width           =   1155
         End
         Begin VB.OptionButton opt_totalizar 
            Caption         =   "Importe"
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   24
            Top             =   180
            Value           =   -1  'True
            Width           =   1155
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Seleccionar"
         Height          =   1260
         Index           =   0
         Left            =   5640
         TabIndex        =   14
         Top             =   120
         Width           =   1485
         Begin VB.OptionButton opt_mon 
            Caption         =   "Todo en $."
            Height          =   210
            Index           =   3
            Left            =   45
            TabIndex        =   18
            Top             =   945
            Width           =   1170
         End
         Begin VB.OptionButton opt_mon 
            Caption         =   "Todo en S/."
            Height          =   210
            Index           =   2
            Left            =   45
            TabIndex        =   17
            Top             =   720
            Width           =   1185
         End
         Begin VB.OptionButton opt_mon 
            Caption         =   "Solo $."
            Height          =   210
            Index           =   1
            Left            =   45
            TabIndex        =   16
            Top             =   495
            Width           =   840
         End
         Begin VB.OptionButton opt_mon 
            Caption         =   "Solo S/."
            Height          =   210
            Index           =   0
            Left            =   45
            TabIndex        =   15
            Top             =   270
            Value           =   -1  'True
            Width           =   885
         End
      End
      Begin VB.ListBox ls 
         Height          =   960
         Index           =   0
         Left            =   1590
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   405
         Width           =   1200
      End
      Begin VB.CheckBox ChkMostrarItem 
         Caption         =   "Mostrar item"
         Height          =   195
         Left            =   9990
         TabIndex        =   4
         Top             =   1515
         Width           =   1155
      End
      Begin VB.CommandButton CmdBusProducto 
         Height          =   225
         Left            =   9165
         Picture         =   "FrmConsVenta_Gerencial.frx":32B8
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   1125
         Width           =   210
      End
      Begin VB.TextBox TxtIdTipProd 
         Height          =   300
         Left            =   8790
         MaxLength       =   3
         TabIndex        =   1
         Top             =   1080
         Width           =   615
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   1080
         Index           =   0
         Left            =   60
         TabIndex        =   19
         Top             =   1455
         Width           =   2835
         _cx             =   5001
         _cy             =   1905
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsVenta_Gerencial.frx":33EA
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
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   1080
         Index           =   1
         Left            =   3010
         TabIndex        =   20
         Top             =   1455
         Width           =   2835
         _cx             =   5001
         _cy             =   1905
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsVenta_Gerencial.frx":3445
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
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   1080
         Index           =   2
         Left            =   5960
         TabIndex        =   21
         Top             =   1455
         Width           =   2835
         _cx             =   5001
         _cy             =   1905
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsVenta_Gerencial.frx":34A7
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
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   1080
         Index           =   3
         Left            =   8910
         TabIndex        =   22
         Top             =   1455
         Width           =   2835
         _cx             =   5001
         _cy             =   1905
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
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsVenta_Gerencial.frx":3503
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Selecc. Mes"
         Height          =   195
         Index           =   6
         Left            =   4095
         TabIndex        =   42
         Top             =   180
         Width           =   885
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Seleccionar Año"
         Height          =   195
         Index           =   3
         Left            =   1605
         TabIndex        =   7
         Top             =   180
         Width           =   1170
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "T.Producto"
         Height          =   195
         Left            =   8790
         TabIndex        =   6
         Top             =   900
         Width           =   795
      End
      Begin VB.Label lblTipProducto 
         BorderStyle     =   1  'Fixed Single
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
         Height          =   300
         Left            =   9420
         TabIndex        =   5
         Top             =   1080
         Width           =   2340
      End
   End
End
Attribute VB_Name = "FrmConsVenta_Gerencial"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'--PARA EXPORTAR A EXCEL
Dim Oleapp As Object
Dim vCantMeses As Integer
'--VARIABLES DE PROPIEDADES DE GRAFICO
Dim vLngTipoGrafico As Long, vTipoDato As Integer
Dim vTituloGraf As String, vViewLeyenda As Boolean
'--FIN PARA EXPORTAR A EXCEL


'-- ALMACENAR LOS TOTALES DE TODA LA CONSULTA
Dim Arr_Totales_cols() As Double '--ALMACENAR TOTALES POR TODAS LAS FILAS
Dim Arr_Totales_col() As Double     '--ALMACENAR TOTALES POR COLUMNA, SE LIMPIA DESPUES DE CAMBIO DE GRUPO
Dim Arr_Totales_row() As Double     '--ALMACENAR TOTALES POR FILA

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
                                 '--OBTENDRA VALOR EN GENERAR_CONSULTA()

Dim Q_COL_COMPARAR_GRUPO As Integer '--INDICA LA COLUMNA PARA COMPARAR EL GRUPO
                                    '--OBTENDRA VALOR EN GENERAR_CONSULTA()

Dim Q_COL_ARR_TOTAL As Integer  '--NOS INDICA EL TOTAL DE COLUMNAS QUE VA A CONTENER LOS ACUMULADOS
                                '--OBTENDRA VALOR EN VALIDAR_CONSULTA()
                                '--SI SEL MES: ENE, FEB, MAR => Q_COL_ARR_TOTAL= 2
                                '--SI SEL TRI: ENE-MAR, ABR-JUN => Q_COL_ARR_TOTAL= 1 OBS: SE INICIA DESDE POS=0

Private Sub CmdBusProducto_Click()
On Error GoTo error
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha

    Dim xCampos(2, 4) As String

    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"

    xform.SQLCad = "SELECT id, descripcion FROM mae_tipoproducto "

    xform.Titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If Me.TxtIdTipProd.Text <> "" And Me.TxtIdTipProd.Text <> CStr(xRs.Fields("id")) Then LimpiarGrid Me.fg(3), True
        TxtIdTipProd.Text = xRs("id")
        lblTipProducto.Caption = xRs("descripcion")
    End If
    
    ChkMostrarItem_Click
     
    Set xform = Nothing
    Set xRs = Nothing
    Exit Sub
error:
    Set xform = Nothing
    Set xRs = Nothing
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"

End Sub




Private Sub CONSULTAR()
    'On Error GoTo error
    Dim rst_select As New ADODB.Recordset
    '--
    Dim CN_TMP As New ADODB.Connection '--CONEX TEMPORAL
    Dim Rst_RUTA As New ADODB.Recordset '--CARGA RUTAS DE BD'S
    
    Dim vStrSelect As String '--RECIBIR LA CONSULTA
    '--CARGANDO RUTAS DE LOS AÑOS SELECCIONADOS
    Dim N_ANYO As String
    Dim SQL_ANYO As String
    Dim k As Integer
    Dim F_CARGAR_1RA_VEZ As Boolean '--TRUE::SE CARGA POR 1RA VEZ LA GRILLA
    
    If Validar_Consulta(N_ANYO) = False Then Exit Sub
    
    BAND_INTERRUMPIR = False
    '--CONFIGURAR LA PRESENTACION DE LA CONSULTA
    LimpiarGrid Me.Fg1
    '--INVOCAR A ESTA FUNCION PARA OBTENER LOS VALORES DE
        '--Q_POS_MES , Q_POS_MES_INICIO
    GENERAR_CONSULTA "-1"
    Configurar_Grilla
    '--CARGANDO RUTAS DE LOS AÑOS SELECCIONADOS
    SQL_ANYO = " AND anotra IN (" + Left(N_ANYO, Len(N_ANYO) - 1) + ") "
    '--SI LA BASE DE BATOS PRINCIPAL EXISTE
    If ArchivoExiste(AP_RUTABD + "data.mdb") = False Then
        MsgBox "No existe la ruta a la Base de Datos Principal", vbCritical, "Mensaje..."
        Exit Sub
    End If
    '--ABRIENDO LA CONEXION PARA OBTENER EL LISTADO DE RUTAS A LAS BASES DE DATOS
    OPEN_CONEX_TMP CN_TMP, AP_RUTABD + "data.mdb"
    If CN_TMP.State = 0 Then Exit Sub
    '----
    RST_Busq rst_select, "SELECT ruta,anotra FROM mae_empresa WHERE numruc = '" + NumRUC + "' " + SQL_ANYO + " ORDER BY anotra ASC ", CN_TMP
    '--CARGAR RST TEMPORAL
    DEFINIR_RST_TMP Rst_RUTA, rst_select
    CARGAR_RST_TMP Rst_RUTA, rst_select
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
    
''''    '-------------------
''''    Dim TMPRstRuta As New ADODB.Recordset
''''
''''    DEFINIR_RST_TMP TMPRstRuta, Rst_RUTA
''''    Rst_RUTA.Filter = "anotra=" + CStr(AnoTra)
''''    If Rst_RUTA.RecordCount <> 0 Then CARGAR_RST_TMP TMPRstRuta, Rst_RUTA, , , True
''''    Rst_RUTA.Filter = ""
''''    Rst_RUTA.Filter = "anotra <> " + CStr(AnoTra)
''''    If Rst_RUTA.RecordCount <> 0 Then CARGAR_RST_TMP TMPRstRuta, Rst_RUTA
''''    Set Rst_RUTA = Nothing
''''    Set Rst_RUTA = TMPRstRuta
''''    Set TMPRstRuta = Nothing
''''    Rst_RUTA.MoveFirst
''''    '--------------------
    
    For k = 0 To Rst_RUTA.RecordCount - 1
    
    
        lbl(4).Caption = "Año: " + CStr(Rst_RUTA.Fields(1))
        PgBar(0).Value = k + 1
        '------------------------------------------------
        If k = 0 Then
            '--ENTRAR SOLO UNA VEZ
            vStrSelect = GENERAR_CONSULTA(CStr(Rst_RUTA.Fields(1)))
        Else
            '--EN LOS DEMAS AÑO REEMPLAZAR EL AÑO ANTERIOR POR EL AÑO ACTUAL
            vStrSelect = Replace(vStrSelect, ARR_ANYO(k - 1), CStr(Rst_RUTA.Fields(1)))
        End If
        '------------------------------------------------
        If vStrSelect = "" Then GoTo salir
        '--SI EL ARCHIVO EXISTE
        If ArchivoExiste(AP_RUTABD + Rst_RUTA.Fields(0) & "") = False Then
            MsgBox "No existe la ruta a la Base de Datos Año: " + CStr(Rst_RUTA.Fields(1)), vbCritical, "Mensaje..."
            GoTo salir
        End If
        '--ABRIENDO LA CONEXION A LA BASE DE DATOS
        OPEN_CONEX_TMP CN_TMP, AP_RUTABD + Rst_RUTA.Fields(0) & ""
        If CN_TMP.State = 0 Then Exit Sub
        '--CARGADO EL RST
        Set rst_select = Nothing
        RST_Busq rst_select, vStrSelect, CN_TMP
        '--SI SELECCIONA TODO EN SOLES O TODO EN DOLARES
        If opt_mon(2).Value = True Or opt_mon(3).Value = True Then
            CARGAR_DATOS_TMP CN_TMP, rst_select, CStr(Rst_RUTA.Fields(1))
        End If
        '--------------------------------------
        If opt_consulta(0).Value = True And (Me.TxtIdTipProd.Text <> "" Or Me.ChkMostrarItem.Value = 1) Then Comparar_Grupo Fg1, Rst_RUTA, False, CStr(Rst_RUTA.Fields(1)), 1
        '--------------------------------------
        If rst_select.RecordCount > 0 Then
            If F_CARGAR_1RA_VEZ = False Or Me.opt_consulta(0).Value = True Then
                '--CARGA LOS DATOS DEL PRIMER AÑO
                CARGAR_DATOS_GRILLA rst_select, CStr(Rst_RUTA.Fields(1))
                F_CARGAR_1RA_VEZ = True
            Else
                '--CUANDO LOS DATOS ESTAN CARGADOS => AGREGAR DATOS A LOS DEMAS AÑOS
                CARGAR_DATOS_GRILLA_OTROS_ANYOS rst_select, CStr(Rst_RUTA.Fields(1))
            End If
        End If
    '        Set Me.Fg1.DataSource = Rst_Select
        CN_TMP.Close
        '--------------------------------------
        Rst_RUTA.MoveNext
    Next k
    '-----CUANDO LA CONSULTA ES X AÑOS COLOCAR LOS TOTALES
    If opt_consulta(0).Value = True Then
        CARGAR_DATOS_GRILLA_ADD_TOTALES True, "Tot Gen:", True, True, ARR_ANYO(k - 1)
    End If
    '
    PgBar(0).Value = PgBar(0).Max
salir:
    FraProgreso.Visible = False
    Set Rst_RUTA = Nothing
    Set rst_select = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    Set rst_select = Nothing
    CN_TMP.Close
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"

End Sub
Private Sub CARGAR_DATOS_TMP(CN_TMP As ADODB.Connection, _
                             RST_ORIGEN As ADODB.Recordset, _
                             M_ANYO As String)

    Dim RST_TMP As New ADODB.Recordset
    Dim RST_GRUPO As New ADODB.Recordset
    Dim SQL_CONSULTA As String
    Dim N_FILTRO As String
    Dim Q_ROW_GRUPO As Integer
    Dim Q_ROW1 As Integer
    Dim Q_ROW_TMP As Integer
    Dim Pos As Integer
    '--
    Dim vStrCampo As String
    
    '--GENERAR LA CONSULTA DE
    SQL_CONSULTA = GENERAR_CONSULTA(M_ANYO, True)
    '--DEFINIR LOS CAMPOS DEL RECORDSET
    DEFINIR_RST_TMP RST_TMP, RST_ORIGEN
    '--CARGAR LOS GRUPOS
    RST_Busq RST_GRUPO, SQL_CONSULTA, CN_TMP
    If RST_GRUPO.RecordCount = 0 Then Exit Sub

    PgBar(1).Min = 0
    If RST_GRUPO.RecordCount = 1 Then
        PgBar(1).Max = 1
    Else
        PgBar(1).Max = RST_GRUPO.RecordCount - 1
    End If
    For Q_ROW_GRUPO = 0 To RST_GRUPO.RecordCount - 1
        '--LOS FILTROS VAN HACER SOBRE LOS ID'S
        PgBar(1).Value = Q_ROW_GRUPO
        DoEvents
        N_FILTRO = ""
        For Q_ROW1 = 0 To Q_COL_FILA_OCULTA - 1
            N_FILTRO = N_FILTRO + RST_GRUPO.Fields(Q_ROW1).Name + "= " + CStr(RST_GRUPO.Fields(Q_ROW1)) + " AND "
        Next Q_ROW1
        N_FILTRO = Left(N_FILTRO, Len(N_FILTRO) - 5) '--QUITO EL ÚLTIMO AND
        RST_ORIGEN.Filter = N_FILTRO '--HACER EL FILTRO
        If RST_ORIGEN.RecordCount > 0 Then
            '--CARGAR EL PRIMER REGISTRO
            CARGAR_RST_TMP RST_TMP, RST_ORIGEN, "", 0, True
            '--CARGAR LOS DEMAS REGISTROS
            RST_ORIGEN.MoveFirst
            If RST_ORIGEN.EOF = False Then RST_ORIGEN.MoveNext
            Do While Not RST_ORIGEN.EOF
                DoEvents
                '-------Q_ROW1 TOMA VALOR DE LAS COLUMNAS
                For Q_ROW1 = 0 To RST_ORIGEN.Fields.Count - 1
                    '--SI SE NTERRUMPE EL PROCESO => SALIR
                    If BAND_INTERRUMPIR = True Then Exit Sub
                    vStrCampo = RST_ORIGEN.Fields(Q_ROW1).Name
                    Select Case LCase(vStrCampo)
                        Case "total", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"
                            RST_TMP.Fields(vStrCampo) = NulosN(RST_TMP.Fields(vStrCampo)) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                    End Select
                Next Q_ROW1
                '--------
                RST_ORIGEN.MoveNext
            Loop
        End If
        RST_GRUPO.MoveNext
    Next Q_ROW_GRUPO

    '--RENOMBRANDO LOS DATOS AL RECORSET PARA QUE SE MUESTRE EN LA GRILLA
    Set RST_ORIGEN = RST_TMP
    Set RST_GRUPO = Nothing
    Set RST_TMP = Nothing

End Sub


Private Function CARGAR_DATOS_GRILLA(RST_ORIGEN As ADODB.Recordset, _
                                         M_ANYO As String)
                                         
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
        Comparar_Grupo Fg1, RST_ORIGEN, BAND_ADD_REG, M_ANYO
        '---------------------------------------------------------
        ADD_REG Fg1
        '--ACUMULAR EN EL ARRAY_MES
        CARGAR_DATOS_ARRAY RST_ORIGEN
        '--CARGAR A LA GRILLA
        CARGAR_DATOS_GRILLA_ARRAY_TMP RST_ORIGEN, M_ANYO, Fg1.Rows - 1
        '---------------------------------------------------------
        RST_ORIGEN.MoveNext
'        --PONER TOTALES AL FINAL DE LA GRILLA
        If RST_ORIGEN.EOF Then
            If opt_consulta(0).Value = False Then
                CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Total:", , , M_ANYO
                CARGAR_DATOS_GRILLA_ADD_TOTALES True, "Tot Gen:", True, True, M_ANYO
            End If
        Else
        
            PgBar(1).Value = CLng(RST_ORIGEN.Bookmark)
            
        End If
    Wend
    PgBar(1).Value = 0
    
    If Me.opt_consulta(0).Value = False Then Limpiar_ARRAY_TOTAL True
    
End Function


Private Sub Comparar_Grupo(GRID As Object, _
                            RST_ORIGEN As ADODB.Recordset, _
                            BAND_ADD_REG As Boolean, _
                            M_ANYO As String, _
                            Optional Q_COL_COMPARAR As Integer = -1)
                            
    '--FUNCION QUE NOS PERMITE ARMAR LOS GRUPOS
    '--COMPARA CUANDO CAMBIAR DE GRUPO
    Dim RST_TEPM_1 As New ADODB.Recordset
    
    '---------------------------------------------------------
    If Q_COL_COMPARAR_GRUPO = -1 Then GoTo salir
    '---------------------------------------------------------
    If Q_COL_COMPARAR = -1 Then Q_COL_COMPARAR = Q_COL_COMPARAR_GRUPO
    
    
    If RST_ORIGEN.Bookmark = 1 Then
        '--SE CARGA EN GENERAR_CONSULTA() Q_COL_COMPARAR_GRUPO
        ADD_REG GRID, Fila_grupo
        UNIR_CELDAS GRID, GRID.Rows - 1, Q_COL_COMPARAR + 1, GRID.Rows - 1, Q_POS_MES_INICIO - 1, RST_ORIGEN.Fields(Q_COL_COMPARAR) & "", flexAlignLeftCenter:        FORMATO_CELDA GRID, GRID.Rows - 1, Q_COL_COMPARAR_GRUPO + 1
    Else
    
        Set RST_TEPM_1 = RST_ORIGEN.Clone
        RST_TEPM_1.Bookmark = RST_ORIGEN.Bookmark
        RST_TEPM_1.MovePrevious
        
        If RST_TEPM_1.Fields(Q_COL_COMPARAR) <> RST_ORIGEN.Fields(Q_COL_COMPARAR) Then
            CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Total:", , , M_ANYO
            
            ADD_REG GRID, Fila_en_Blanco
            UNIR_CELDAS GRID, GRID.Rows - 1, IIf(Q_COL_FILA_OCULTA = -1, 1, Q_COL_FILA_OCULTA + 1), GRID.Rows - 1, Fg1.Cols - 1, " ", flexAlignLeftCenter
            
            Limpiar_ARRAY_TOTAL

            ADD_REG GRID, Fila_grupo
            UNIR_CELDAS GRID, GRID.Rows - 1, Q_COL_COMPARAR + 1, GRID.Rows - 1, Q_POS_MES_INICIO - 1, RST_ORIGEN.Fields(Q_COL_COMPARAR) & "", flexAlignLeftCenter:    FORMATO_CELDA GRID, GRID.Rows - 1, Q_COL_COMPARAR_GRUPO + 1

        End If
    End If
salir:
    Set RST_TEPM_1 = Nothing
End Sub



Private Function CARGAR_DATOS_GRILLA_OTROS_ANYOS(RST_ORIGEN As ADODB.Recordset, _
                                         M_ANYO As String)
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    Dim Q_ROW1 As Integer
    
    If RST_ORIGEN.RecordCount > 0 Then
        RST_ORIGEN.MoveFirst
    Else
        Exit Function
    End If
    
    PgBar(1).Min = 0
    PgBar(1).Max = Fg1.Rows
    
'    Fg1.Row = 2
    Dim Q_ROW As Integer '--INDICA LA POSICION DEL REGISTRO A AGREGAR DATOS
    Dim N_FILTRO As String '--INDICA EL FILTRO QUE SE TENDRA QUE HACER AL RECORDSET
                            '-- DEPENDE DE Q_COL_FILA_OCULTA
                            
    For Q_ROW = 2 To Fg1.Rows - 1
        Fg1.Row = Q_ROW
        PgBar(1).Value = Q_ROW
        N_FILTRO = ""
        '--CONCATENO MI FILTRO
        If Fg1.TextMatrix(Fg1.Row, 1) = e_ESTADO_ROW_GRID.Fila_grupo Then
        
        ElseIf Fg1.TextMatrix(Fg1.Row, 1) = e_ESTADO_ROW_GRID.Fila_Total Then
            CARGAR_DATOS_GRILLA_ADD_TOTALES False, "Total:", , , M_ANYO, True
            Limpiar_ARRAY_TOTAL
        ElseIf Fg1.TextMatrix(Fg1.Row, 1) = e_ESTADO_ROW_GRID.Fila_Total_grl Then
            CARGAR_DATOS_GRILLA_ADD_TOTALES True, "Tot Gen:", True, True, M_ANYO, True
        ElseIf Fg1.TextMatrix(Fg1.Row, 1) = e_ESTADO_ROW_GRID.Fila_en_Blanco Then
        
        Else
            For Q_ROW1 = 0 To Q_COL_FILA_OCULTA - 1
                N_FILTRO = N_FILTRO + RST_ORIGEN.Fields(Q_ROW1).Name + "= " + Fg1.TextMatrix(Fg1.Row, Q_ROW1 + 1) + " AND "
            Next Q_ROW1
            N_FILTRO = Left(N_FILTRO, Len(N_FILTRO) - 5) '--QUITO EL ULTIMO AND
            RST_ORIGEN.Filter = N_FILTRO '--HACER EL FILTRO
            If RST_ORIGEN.RecordCount > 0 Then
                DoEvents
                '--SI SE NTERRUMPE EL PROCESO => SALIR
                If BAND_INTERRUMPIR = True Then Exit Function
                '--ACUMULAR EN EL ARRAY_MES
                CARGAR_DATOS_ARRAY RST_ORIGEN
                '--CARGAR_DATOS A LA GRILLA
                CARGAR_DATOS_GRILLA_ARRAY_TMP RST_ORIGEN, M_ANYO, Q_ROW, True
            End If
        End If
    Next Q_ROW
    Limpiar_ARRAY_TOTAL True
End Function

Private Sub CARGAR_DATOS_ARRAY(RST_ORIGEN As ADODB.Recordset)
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
            '--ACUMULANDO X MES
            
            Case "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12"
                '"ene", "feb", "mar", "abr", "may", "jun", "jul", "ago", "sep", "oct", "nov", "dic"
                '"ene-mar","abr-jun","jul-sep","oct-dic"
                '"1re sem","2do sem"
            '--ARR_TMP(0, 2) INDICA LA PRIMERA COLUMNA A MOSTRAR
                If LCase(vStrCampo) = ARR_TMP(0, 2) Then Q_POS = 0
                Arr_Totales_col(Q_POS, 0) = Arr_Totales_col(Q_POS, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Q_POS = Q_POS + 1

            Case "total":
                Arr_Totales_col(Q_COL_ARR_TOTAL + 1, 0) = Arr_Totales_col(Q_COL_ARR_TOTAL + 1, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Arr_Totales_col(Q_COL_ARR_TOTAL + 2, 0) = Arr_Totales_col(Q_COL_ARR_TOTAL + 2, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
        End Select
    Next Q_CAMPO
    
End Sub


Private Function CARGAR_DATOS_GRILLA_ARRAY_TMP(RST_ORIGEN As ADODB.Recordset, _
                                        M_ANYO As String, _
                                         Q_ROW As Integer, _
                                         Optional F_OTROS_ANYOS As Boolean = False)
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    
    Dim Q_INCREMENTO_X_COL As Integer   '--SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
    Dim Q_POS_MES_TOTAL  As Integer     '--POSICIONA A LA COLUMNA DEL TOTAL X AÑO
    Dim Q_POS As Integer
    Dim Q_CAMPO As Integer
    Dim vStrCampo As String
    
    
    For Q_POS = 0 To UBound(ARR_ANYO) - 1
        If ARR_ANYO(Q_POS) = M_ANYO Then
            Q_INCREMENTO_X_COL = Q_POS
            Exit For
        End If
    Next
    '--IDENTIFICAR LA POSICION DE INICIO DE MES(ENERO)
    If Me.opt_consulta(0).Value = True Then Q_INCREMENTO_X_COL = 0
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
                If LCase(vStrCampo) = ARR_TMP(0, 2) Then Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    
                Fg1.TextMatrix(Q_ROW, Q_POS_MES) = CONVERTIR_A_ESCALA(NulosN(RST_ORIGEN.Fields(vStrCampo)))
                If Me.opt_consulta(0).Value = True Then
                    Q_POS_MES = Q_POS_MES + 1
                Else
                    Q_POS_MES = Q_POS_MES + Q_TOTAL_ANYO
                End If
             '--DEL TOTAL DEL AÑO
            Case "total"
                If Me.opt_consulta(0).Value = True Then
                    Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL + 1) * 1
                Else
                    Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL + 1) * Q_TOTAL_ANYO + Q_INCREMENTO_X_COL
                End If
                '--TOTAL AÑO
                Fg1.TextMatrix(Q_ROW, Q_POS_MES_TOTAL) = CONVERTIR_A_ESCALA(NulosN(RST_ORIGEN.Fields(vStrCampo)))
                '--TOTALIZAR POR FILA
                '--TOTAL GRL
                If Me.opt_consulta(0).Value = False Then
                    If IsNumeric(Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1)) = False Then
                        Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1) = CONVERTIR_A_ESCALA(NulosN(RST_ORIGEN.Fields(vStrCampo)))
                    Else
                        Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1) = CDbl(Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1)) + CONVERTIR_A_ESCALA(NulosN(RST_ORIGEN.Fields(vStrCampo)))
                    End If
                    Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1) = CONVERTIR_A_ESCALA(CDbl(Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1)), True)
                End If
            '--DE LOS DEMAS CAMPOS
            Case Else
                '--SOLO SE AGREGARAN EN EL PRIMER AÑO
                If F_OTROS_ANYOS = False Then Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = RST_ORIGEN.Fields(vStrCampo) & ""
        End Select
        '------------
    Next
End Function



Private Sub pImprimir()

    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, T_RPT_TITULO + " ", "", T_RPT_PERIODO + " ", False, True

    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"

End Sub

Private Sub ChkMostrarItem_Click()
If Me.ChkMostrarItem.Value = 0 Then
    fg(3).Enabled = False
Else
    '--LIMPIAR GRILLA
    fg(3).Enabled = True
    LimpiarGrid fg(3), True
    GRID_COMBOLIST fg(3)
End If
'---BLOQUEAR OPCIONES

If opt_totalizar(0).Value = True Then
    opt_totalizar_Click 0
Else
    opt_totalizar_Click 1
End If
End Sub



Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nOrden As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim nSQLNotIn As String
    
    If Col <> 2 Then Exit Sub
    Select Case Index
    Case 0 '--CLIENTE
            ReDim xCampos(3, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "R.U.C.":   xCampos(1, 1) = "numruc":    xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "Id":   xCampos(2, 1) = "id":        xCampos(2, 2) = "800":   xCampos(2, 3) = "N"
            '--si hay filtros
            nSQLNotIn = GRID_GENERAR_SQL_ID(fg(0), 1, " WHERE mae_cliente.id", "NOT IN", True)
            If NulosC(fg(Index).TextMatrix(Row, Col)) <> "" Then
                nSQLNotIn = IIf(nSQLNotIn = "", " WHERE ", nSQLNotIn & " AND ") & "  (UCASE(mae_cliente.nombre) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' OR UCASE(mae_cliente.nombre) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' ) "
            End If
            '--------------
            nSQL = "SELECT id, nombre,numruc FROM mae_cliente " & nSQLNotIn & "  order by nombre asc"
            
            nTitulo = "Buscando Clientes"
            nOrden = "nombre"
            nCampoBusca = "nombre"
    Case 1 '--PTO VENTA
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4000":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Cliente":  xCampos(1, 1) = "cliente":   xCampos(1, 2) = "3200":   xCampos(1, 3) = "C"
            xCampos(2, 0) = "Id":   xCampos(2, 1) = "id":        xCampos(2, 2) = "800":   xCampos(2, 3) = "N"
           '--si hay filtros
            nSQLNotIn = GRID_GENERAR_SQL_ID(fg(1), 1, " WHERE vta_puntoVenta.id", "NOT IN", True)
            If NulosC(fg(Index).TextMatrix(Row, Col)) <> "" Then
                nSQLNotIn = IIf(nSQLNotIn = "", " WHERE ", nSQLNotIn & " AND ") & "  (UCASE(vta_puntoVenta.descripcion) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' OR UCASE(vta_puntoVenta.descripcion) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' ) "
            End If
            '------
            nSQL = "SELECT vta_puntoVenta.id, vta_puntoVenta.descripcion AS nombre, mae_cliente.nombre as cliente " _
                + vbCr + " FROM vta_puntoVenta INNER JOIN mae_cliente ON vta_puntoVenta.idcli = mae_cliente.id " & nSQLNotIn _
                + vbCr + " ORDER BY mae_cliente.nombre, vta_puntoVenta.descripcion;"
            
            nTitulo = "Buscando Punto de Venta"
            nOrden = "nombre"
            nCampoBusca = "nombre"
    Case 2 '--VENDEDOR
            ReDim xCampos(2, 3) As String
            xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "4500":   xCampos(0, 3) = "C"
            xCampos(1, 0) = "Codigo":   xCampos(1, 1) = "id":        xCampos(1, 2) = "800":   xCampos(1, 3) = "N"
            '--si hay filtros
            nSQLNotIn = GRID_GENERAR_SQL_ID(fg(2), 1, " WHERE vta_vendedores.id", "NOT IN", True)
            If NulosC(fg(Index).TextMatrix(Row, Col)) <> "" Then
                nSQLNotIn = IIf(nSQLNotIn = "", " WHERE ", nSQLNotIn & " AND ") & "  (UCASE(pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' OR UCASE(pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' ) "
            End If
            '-------------
            nSQL = "SELECT vta_vendedores.id, pla_empleados.apepat & ' ' &  pla_empleados.apemat & ' ' & pla_empleados.nom AS nombre " _
                + vbCr + " FROM pla_empleados INNER JOIN vta_vendedores ON pla_empleados.id = vta_vendedores.idper " & nSQLNotIn _
                + vbCr + " ORDER BY pla_empleados.apepat & ' ' & pla_empleados.apemat & ' ' & pla_empleados.nom ;"
    
            nTitulo = "Buscando Vendedores"
            nOrden = "nombre"
            nCampoBusca = "nombre"
    
    Case 3 '--ITEM
        If TxtIdTipProd.Text = "" Then
            MsgBox "Falta especificar el tipo de item...!", vbExclamation, xTitulo
            TxtIdTipProd.SetFocus
            Exit Sub
        End If
        '---
        ReDim xCampos(3, 3) As String
        xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Cod. Prod.":    xCampos(1, 1) = "codpro":         xCampos(1, 2) = "2000":    xCampos(1, 3) = "C"
        xCampos(2, 0) = "Id":            xCampos(2, 1) = "id":             xCampos(2, 2) = "800":         xCampos(2, 3) = "N"
        
        nSQLNotIn = GRID_GENERAR_SQL_ID(fg(3), 1, " and alm_inventario.id", "NOT IN", True)
        If NulosC(fg(Index).TextMatrix(Row, Col)) <> "" Then
            nSQLNotIn = " AND (UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' OR UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' ) "
        End If
        '-------------
        nSQL = "SELECT id, codpro, descripcion as nombre FROM alm_inventario WHERE tippro = " & NulosN(TxtIdTipProd.Text) & nSQLNotIn & ""
        nTitulo = "Buscando Tipo de Item"
        nOrden = "nombre"
        nCampoBusca = "nombre"
    
    End Select
    fg(Index).TextMatrix(Row, Col) = ""
    Dim xRs As New ADODB.Recordset
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, nOrden, nCampoBusca, Principio

    If xRs.State = 0 Then GoTo salir
    If xRs.RecordCount = 0 Then GoTo salir
    
    fg(Index).TextMatrix(Row, 1) = NulosN(xRs("id"))
    fg(Index).TextMatrix(Row, 2) = NulosC(xRs("nombre"))
                
    If fg(Index).Row = fg(Index).Rows - 1 Then fg(Index).AddItem ""
    fg(Index).Row = fg(Index).Rows - 1: fg(Index).Col = 2
        
salir:
    Set xRs = Nothing

Exit Sub
error:
    
    Set xRs = Nothing
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"

End Sub

Private Sub Fg_DblClick(Index As Integer)
    Fg_CellButtonClick Index, fg(Index).Rows - 1, 2
End Sub

Private Sub Fg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If fg(Index).Row = -2 Then Exit Sub
    Select Case KeyCode
        Case 45  'INSERTAR REGI
            fg(Index).AddItem ""
            fg(Index).Row = fg(Index).Rows - 1: fg(Index).Col = 1
        Case 46 'SUPRIMIR/DELETE
            If fg(Index).Rows - 1 >= 2 Then
                fg(Index).RemoveItem fg(Index).Row
            Else
                LimpiarGrid fg(Index), True
                GRID_COMBOLIST fg(Index)
            End If
    End Select
End Sub

Private Sub fg_KeyPressEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If validar_letras(KeyAscii) = False Then
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        '--interrumpir
        BAND_INTERRUMPIR = True
    End If
End Sub

Private Sub Form_Load()
On Error GoTo error
    Dim k As Integer
    '--CARGAR DATOS
    
    CentrarFrm Me
    '--FORMATO DE LAS GRILLAS
    For k = 0 To fg.Count - 1
        GRID_COMBOLIST fg(k)
        fg(k).Tag = fg(k).FormatString
    Next k
    Fg1.Tag = Fg1.FormatString
    
    LimpiarGrid Me.Fg1
    '--CARGAR LOS AÑOS
    If CARGAR_LISTA_ANYOS_ACTIVOS(ls(0), xCon) = False Then Exit Sub
    Llenar_Mes ls(1)
    '--CARGANDO LOS MESES
    CARGAR_ARR_XX ARR_XX(), X_MES
    '--SELECCIONAR EL AÑO ACTUAL
    ls_activar_chek ls(0), AnoTra
    ls_activar_chek ls(1)
    '--CONFIGURAR LA GRILLA
    Validar_Consulta "-1"
    GENERAR_CONSULTA "-1"
    Configurar_Grilla
    Exit Sub
error:
    SHOW_ERROR
End Sub

Private Sub opt_consulta_Click(Index As Integer)
    If opt_totalizar(0).Value = True Then
        opt_totalizar_Click 0
    Else
        opt_totalizar_Click 1
    End If
End Sub

Private Sub opt_estilo_Click(Index As Integer)
    Select Case Index
        Case 0 '--MES
            Llenar_Mes ls(1)
            ls_activar_chek ls(1)
            CARGAR_ARR_XX ARR_XX(), X_MES
        Case 1 '--TRIMESTRE
            Llenar_Trimestre ls(1)
            ls_activar_chek ls(1)
            CARGAR_ARR_XX ARR_XX(), X_TRIMESTRE
        Case 2 '--SEMESTRE
            Llenar_Semestre ls(1)
            ls_activar_chek ls(1)
            CARGAR_ARR_XX ARR_XX(), X_SEMESTRE
    End Select
    lbl(6).Caption = "Selecc. " + opt_estilo(Index).Caption
    
End Sub

Private Sub opt_totalizar_Click(Index As Integer)
    If Index = 0 Then '--importe
        If Me.TxtIdTipProd.Text = "" And Me.ChkMostrarItem.Value = 0 And (opt_consulta(2).Value = False And opt_consulta(4).Value = False) Then
            habilitar opt_importe, True
        Else
            habilitar opt_importe, False
            opt_mon(2).Enabled = False: opt_mon(3).Enabled = False
        End If
        habilitar opt_escala, True
        habilitar opt_mon, True
        opt_mon(0).Value = True
        opt_importe(0).Value = True
        
    Else '--cantidades
        habilitar opt_mon, False
        habilitar opt_importe, False
        habilitar opt_escala, False: opt_escala(0).Value = True
        opt_mon(0).Value = False: opt_mon(1).Value = False: opt_mon(2).Value = False: opt_mon(3).Value = False
        opt_importe(0).Value = False: opt_importe(1).Value = False: opt_importe(2).Value = False
    End If
End Sub



Private Sub TxtIdTipProd_Change()
    If TxtIdTipProd.Text = "" Then
        lblTipProducto.Caption = ""
        If Me.ChkMostrarItem.Value = 1 Then ChkMostrarItem.Value = 0
        LimpiarGrid Me.fg(3), True
        ChkMostrarItem_Click
    End If
End Sub

Private Sub TxtIdTipProd_KeyPress(KeyAscii As Integer)
    On Error GoTo error
    If KeyAscii = 13 Then
        Dim RsTipProd As New ADODB.Recordset
        RsTipProd.CursorLocation = adUseClient
        If TxtIdTipProd.Text <> "" Then
            Set RsTipProd = BuscaConCriterio("SELECT id, descripcion FROM mae_tipoproducto WHERE id =" & Val(TxtIdTipProd.Text) & "", xCon)
            If RsTipProd.RecordCount <> 0 Then
                lblTipProducto.Caption = RsTipProd("descripcion")
            Else
                lblTipProducto.Caption = ""
                TxtIdTipProd.Text = ""
            End If
        End If
        ChkMostrarItem_Click
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
    Set RsTipProd = Nothing
    Exit Sub
error:
    Set RsTipProd = Nothing
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"


End Sub

Private Sub TxtIdTipProd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then  'TECHAL F5
        CmdBusProducto.Value = True
    End If
End Sub

'------
Private Function Validar_Consulta(N_ANYO As String) As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    '--DE LA FECHA ES NULL
    Dim k As Integer
    N_ANYO = ""
    Q_TOTAL_ANYO = 0
    '--RECORRER AÑO A AÑO PARA CARGAR LA DATA
    For k = ls(0).ListCount - 1 To 0 Step -1
        ls(0).ListIndex = k
        If ls(0).Selected(k) = True Then
            N_ANYO = N_ANYO + ls(0).Text + ","
            Q_TOTAL_ANYO = Q_TOTAL_ANYO + 1
        End If
    Next
    
    If N_ANYO = "" Then
       MsgBox "Seleccione un Año como mínimo", vbCritical, "Mensaje..."
'       ls(0).SetFocus
       Exit Function
    End If
    Erase ARR_ANYO '--LIMPIAR ARRAY
    ARR_ANYO = Split(N_ANYO, ",") '--ASIGNANDO EL LISTADO DE AÑOS
    
    
    '----------------
    Q_COL_ARR_TOTAL = 0
    For k = ls(1).ListCount - 1 To 0 Step -1
        ls(1).ListIndex = k
        If ls(1).Selected(k) = True Then
            Q_COL_ARR_TOTAL = Q_COL_ARR_TOTAL + 1
        End If
    Next
    If Q_COL_ARR_TOTAL = 0 Then
       MsgBox Replace(lbl(6).Caption, "Selecc.", "Selecc. un ") + " como mínimo...", vbCritical, "Mensaje..."
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

Private Function GENERAR_CONSULTA(M_ANYO As String, Optional F_TODO_SOL_DOL As Boolean = False) As String
    '--FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
    '--
    Dim vStrSelect As String            '--CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    Dim vStrFiltro_ITEM As String       '--SOLO ITEM
    Dim vStrFiltro_CLI As String        '--SOLO CLIENTES
    Dim vStrFiltro_PTO_VTA As String   '--SOLO PUNTOS DE VENTAS
    Dim vStrFiltro_VEND As String       '--SOLO VENDEDORES
    
    Dim vStrFiltro As String
    Dim vStrFiltro_1 As String      '--ESTE FILTRO SERVIRA PARA CONSULTAR EN EL SUB_SELECT
    Dim k As Integer
    '--DEL AÑO
    vStrFiltro = " Year(vta_ventas.fchdoc)= " + M_ANYO + " "
    '--
    '--DEL CLIENTE
    vStrFiltro_CLI = GRID_GENERAR_SQL_ID(fg(0), 1, " AND vta_ventas.idcli", "IN")
    
    '--DEL LOS PUNTOS DE VENTAS
    vStrFiltro_PTO_VTA = GRID_GENERAR_SQL_ID(fg(1), 1, " AND vta_guia.idpunven", "IN")
    
    '--DEL LOS VENDEDORES
    vStrFiltro_VEND = GRID_GENERAR_SQL_ID(fg(2), 1, " AND vta_ventas.idven", "IN")
    

    '--DEL TIPO DE PRODUCTO
    If TxtIdTipProd.Text <> "" Then vStrFiltro = vStrFiltro + " AND alm_inventario.tippro = " + CStr(TxtIdTipProd.Text) + " "
    '--DEL ITEM
    vStrFiltro_ITEM = GRID_GENERAR_SQL_ID(fg(3), 1, " AND alm_inventario.id", "IN")
    
    '--CONCATENAR FECHA + CLIENTE + PUNTO DE VENTA + VENDEDOR + ITEM
    vStrFiltro = vStrFiltro + vStrFiltro_CLI + vStrFiltro_PTO_VTA + vStrFiltro_VEND + vStrFiltro_ITEM
    '---------------
    '--DE LA MONEDA
    '--SOLO s/.
    If opt_mon(0).Value = True Then vStrFiltro = vStrFiltro + " AND vta_ventas.idmon= 1 " '--SOLES
    '--SOLO $
    If opt_mon(1).Value = True Then vStrFiltro = vStrFiltro + " AND vta_ventas.idmon= 2 " '--DOLARES
    '---------------
    
    vStrFiltro = " AND " + vStrFiltro
    
    '------------------------------------------------------------------------------------
    '''vStrFiltro_1 = Replace(vStrFiltro, "vta_ventas.", "vta_ventas1.")
    '''vStrFiltro_1 = Replace(vStrFiltro_1, "alm_inventario.", "alm_inventario1.")
    
    '--GENERAR LA CONSULTA SEGUN CONDICIONES
    Dim nSQLValor As String
    Dim nSQLCampos As String
    Dim nSQLWhere As String
    Dim nSQLFrom As String
    Dim nSQLGroupBy As String
    Dim nSQLOrderBy As String
    Dim nSQLPivot As String
    Dim nSQLPivotSalida As String '--ORDENA LOS VALORE MES A MES(ENE,FEB,MAR,ETC.)
    nSQLWhere = vStrFiltro
    'mae_cliente.nombre AS nomcliente '--CLIENTE
    'vta_ventas.idcli
    'vta_puntoVenta.descripcion '--PUNTO DE VENTA
    'vta_guia.idpunven
    'mae_tipoproducto.descripcion,    '--PRODUCTO
    'alm_inventario.tippro
    'alm_inventario.descripcion AS desctipcom '--ITEM
    'alm_inventario.id
    'pla_empleados.ape & ' ' & pla_empleados.nom AS nombre --EMPLEADO
    'vta_ventas.idven
    
    '---IMPORTE
    'SUM(vta_ventas.imptotdoc) IMPORTE
    
    'SUM(vta_ventasdet.imptot) IMPORTE
    'SUM(vta_ventasdet.canpro) CANTIDAD
    '--DEL LA FECHA
    
    'FORMAT(vta_guia.fecgiro,'mmm')  --DE GUIA
    'FORMAT(vta_ventas.fchdoc,'mmm') --DE VENTAS
    '--AÑO
    'YEAR(vta_guia.fecgiro)  --DE GUIA
    'YEAR(vta_ventas.fchdoc) --DE VENTAS
    
    
        

    
''            Q_COL_FILA_OCULTA       '--OCULTAR COLUMNAS
''            Q_COL_FILA              '--CANTIDAD DE COLUMNAS QUE SE MOSTRARAN DESCONTANDO LOS MESES Y LOS TOTALES
''            Q_POSICION_TOTAL        '--POSICION DE LA COLUMNA QUE SE PONDRA EL TOTAL Y TOTAL_GRL EJ. TOTAL.(COL=2)   S/. 15000
''            Q_COL_COMPARAR_GRUPO    '--NO HAY GRUPO
      
    If opt_consulta(0).Value = True Then '--X AÑO
'''        If (Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0) Then '--AÑO/PRODUCTO
'''            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 4:        Q_POSICION_TOTAL = 2:        Q_COL_COMPARAR_GRUPO = 0
'''            T_RPT_TITULO = "RESUMEN DE VENTAS POR AÑO CON TIPO PRODUCTO"
'''            nSQLCampos = "YEAR(vta_ventas.fchdoc) AS idanyo,alm_inventario.tippro,  YEAR(vta_ventas.fchdoc) AS anyo, mae_tipoproducto.descripcion "
'''            nSQLGroupBy = "alm_inventario.tippro,YEAR(vta_ventas.fchdoc),mae_tipoproducto.descripcion "
'''            nSQLOrderBy = "mae_tipoproducto.descripcion "
'''        ElseIf Me.ChkMostrarItem.Value = 1 Then '--AÑO/PRODUCTO/ITEM
'''            Q_COL_FILA_OCULTA = 0:       Q_COL_FILA = 6:        Q_POSICION_TOTAL = 3:          Q_COL_COMPARAR_GRUPO = 0
'''            T_RPT_TITULO = "RESUMEN DE VENTAS POR AÑO CON ITEM"
'''            nSQLCampos = "YEAR(vta_ventas.fchdoc) AS idanyo,alm_inventario.tippro,alm_inventario.id,  YEAR(vta_ventas.fchdoc) AS anyo,mae_tipoproducto.descripcion,alm_inventario.descripcion AS desctipcom "
'''            nSQLGroupBy = "alm_inventario.tippro,alm_inventario.id,  YEAR(vta_ventas.fchdoc),mae_tipoproducto.descripcion,alm_inventario.descripcion"
'''            nSQLOrderBy = "mae_tipoproducto.descripcion,alm_inventario.descripcion  "
'''        Else    '--SOLO AÑOS
            Q_COL_FILA_OCULTA = 1:        Q_COL_FILA = 2:       Q_POSICION_TOTAL = 2:           Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR AÑO"
            nSQLCampos = "YEAR(vta_ventas.fchdoc) AS idanyo,YEAR(vta_ventas.fchdoc) AS anyo "
            nSQLGroupBy = "YEAR(vta_ventas.fchdoc) "
            nSQLOrderBy = "YEAR(vta_ventas.fchdoc) "
'''        End If
        
    ElseIf opt_consulta(1).Value = True Then '--X CLIENTE
        If (Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0) Then '--CLIETNE/PRODUCTO
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 4:        Q_POSICION_TOTAL = 4:        Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR PROVEEDOR CON TIPO PRODUCTO"
            nSQLCampos = "vta_ventas.idcli,alm_inventario.tippro,  mae_cliente.nombre AS nomcliente,mae_tipoproducto.descripcion "
            nSQLGroupBy = "vta_ventas.idcli,alm_inventario.tippro,  mae_cliente.nombre,mae_tipoproducto.descripcion "
            nSQLOrderBy = "mae_cliente.nombre,mae_tipoproducto.descripcion "
            nSQLWhere = nSQLWhere + " AND alm_inventario.tippro IS NOT NULL " '--SOLO LOS QUE TIENEN TIPO DE PRODUCTO
            
        ElseIf Me.ChkMostrarItem.Value = 1 Then '--CLIENTE/PRODUCTO/ITEM
            Q_COL_FILA_OCULTA = 3:        Q_COL_FILA = 6:        Q_POSICION_TOTAL = 6:        Q_COL_COMPARAR_GRUPO = 3
            T_RPT_TITULO = "RESUMEN DE VENTAS POR CLIENTE CON ITEM"
            nSQLCampos = "vta_ventas.idcli,alm_inventario.tippro,alm_inventario.id,  mae_cliente.nombre AS nomcliente,mae_tipoproducto.descripcion,alm_inventario.descripcion AS desctipcom "
            nSQLGroupBy = "vta_ventas.idcli,alm_inventario.tippro,alm_inventario.id,  mae_cliente.nombre,mae_tipoproducto.descripcion,alm_inventario.descripcion "
            nSQLOrderBy = "mae_cliente.nombre,mae_tipoproducto.descripcion,alm_inventario.descripcion "
            nSQLWhere = nSQLWhere + " AND alm_inventario.tippro IS NOT NULL " '--SOLO LOS QUE TIENEN TIPO DE PRODUCTO
            
        Else    '--SOLO CLIENTE
            Q_COL_FILA_OCULTA = 1:        Q_COL_FILA = 2:        Q_POSICION_TOTAL = 2:        Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR CLIENTE"
            nSQLCampos = " vta_ventas.idcli,mae_cliente.nombre AS nomcliente "
            nSQLGroupBy = "vta_ventas.idcli,mae_cliente.nombre "
            nSQLOrderBy = "mae_cliente.nombre "
        End If
    
    ElseIf opt_consulta(2).Value = True Then '--X PTO DE VENTA
        If (Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0) Then '--X PTO DE VENTA/PRODUCTO
            Q_COL_FILA_OCULTA = 3:        Q_COL_FILA = 6:        Q_POSICION_TOTAL = 6:        Q_COL_COMPARAR_GRUPO = 3
            T_RPT_TITULO = "RESUMEN DE VENTAS POR PUNTO DE VENTA CON TIPO PRODUCTO"
            nSQLCampos = "vta_ventas.idcli,vta_guia.idpunven,alm_inventario.tippro,  mae_cliente.nombre AS nomcliente,vta_puntoVenta.descripcion,mae_tipoproducto.descripcion "
            nSQLGroupBy = "vta_ventas.idcli,vta_guia.idpunven,alm_inventario.tippro,  mae_cliente.nombre,vta_puntoVenta.descripcion,mae_tipoproducto.descripcion "
            nSQLOrderBy = "mae_cliente.nombre,vta_puntoVenta.descripcion "
            nSQLWhere = nSQLWhere + " AND vta_guia.idpunven <>0 " '--SOLO LOS QUE TIENEN PUNTO DE VENTA
            
        ElseIf Me.ChkMostrarItem.Value = 1 Then '--X PTO DE VENTA/PRODUCTO/ITEM
            Q_COL_FILA_OCULTA = 4:        Q_COL_FILA = 8:        Q_POSICION_TOTAL = 8:        Q_COL_COMPARAR_GRUPO = 4
            T_RPT_TITULO = "RESUMEN DE VENTAS POR PUNTO DE VENTA CON ITEM"
            nSQLCampos = " vta_ventas.idcli,vta_guia.idpunven,alm_inventario.tippro,alm_inventario.id,  mae_cliente.nombre AS nomcliente,vta_puntoVenta.descripcion,mae_tipoproducto.descripcion,alm_inventario.descripcion AS desctipcom "
            nSQLGroupBy = " vta_ventas.idcli,vta_guia.idpunven,alm_inventario.tippro,alm_inventario.id,  mae_cliente.nombre,vta_puntoVenta.descripcion,mae_tipoproducto.descripcion,alm_inventario.descripcion "
            nSQLOrderBy = " mae_cliente.nombre,vta_puntoVenta.descripcion "
            nSQLWhere = nSQLWhere + " AND vta_guia.idpunven <>0 " '--SOLO LOS QUE TIENEN PUNTO DE VENTA
            
        Else    '--X PTO DE VENTA
            Q_COL_FILA_OCULTA = 2:        Q_COL_FILA = 4:        Q_POSICION_TOTAL = 4:        Q_COL_COMPARAR_GRUPO = 2
            T_RPT_TITULO = "RESUMEN DE VENTAS POR PUNTO DE VENTA"
            nSQLCampos = " vta_ventas.idcli,vta_guia.idpunven,  mae_cliente.nombre AS nomcliente,vta_puntoVenta.descripcion "
            nSQLGroupBy = "vta_ventas.idcli,vta_guia.idpunven,  mae_cliente.nombre,vta_puntoVenta.descripcion "
            nSQLOrderBy = "mae_cliente.nombre,vta_puntoVenta.descripcion "
            nSQLWhere = nSQLWhere + " AND vta_guia.idpunven <>0 " '--SOLO LOS QUE TIENEN PUNTO DE VENTA
            
        End If
       nSQLWhere = nSQLWhere + " AND (vta_guia.iddocven <>0 OR vta_guia.iddocven  IS NOT NULL) "
    ElseIf opt_consulta(3).Value = True Then '--X VENDEDOR
        If (Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0) Then '--VENDEDOR/PRODUCTO
            Q_COL_FILA_OCULTA = 2:        Q_COL_FILA = 4:        Q_POSICION_TOTAL = 4:        Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR VENDEDOR CON TIPO PRODUCTO"
            nSQLCampos = " vta_ventas.idven,alm_inventario.tippro,  pla_empleados.ape & ' ' & pla_empleados.nom AS nombre,mae_tipoproducto.descripcion "
            nSQLGroupBy = "vta_ventas.idven,alm_inventario.tippro,  pla_empleados.ape & ' ' & pla_empleados.nom,mae_tipoproducto.descripcion "
            nSQLOrderBy = "pla_empleados.ape & ' ' & pla_empleados.nom "
            nSQLWhere = nSQLWhere + " AND vta_ventas.idven <> 0 " '--SOLO LOS QUE TIENEN VENDEDORES
            
        ElseIf Me.ChkMostrarItem.Value = 1 Then '--VENDEDOR/PRODUCTO/ITEM
            Q_COL_FILA_OCULTA = 3:        Q_COL_FILA = 6:        Q_POSICION_TOTAL = 6:        Q_COL_COMPARAR_GRUPO = 3
            T_RPT_TITULO = "RESUMEN DE VENTAS POR VENDEDOR CON ITEM"
            nSQLCampos = " vta_ventas.idven,alm_inventario.tippro,alm_inventario.id,  pla_empleados.ape & ' ' & pla_empleados.nom AS nombre,mae_tipoproducto.descripcion,alm_inventario.descripcion AS desctipcom "
            nSQLGroupBy = "vta_ventas.idven,alm_inventario.tippro,alm_inventario.id, pla_empleados.ape & ' ' & pla_empleados.nom,mae_tipoproducto.descripcion,alm_inventario.descripcion  "
            nSQLOrderBy = "pla_empleados.ape & ' ' & pla_empleados.nom "
            nSQLWhere = nSQLWhere + " AND vta_ventas.idven <> 0 " '--SOLO LOS QUE TIENEN VENDEDORES
            
        Else    '--SOLO VENDEDOR
            Q_COL_FILA_OCULTA = 1:        Q_COL_FILA = 2:        Q_POSICION_TOTAL = 2:        Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR VENDEDOR"
            nSQLCampos = " vta_ventas.idven,  pla_empleados.ape & ' ' & pla_empleados.nom AS nombre"
            nSQLGroupBy = "vta_ventas.idven,pla_empleados.ape & ' ' & pla_empleados.nom "
            nSQLOrderBy = "pla_empleados.ape & ' ' & pla_empleados.nom "
            nSQLWhere = nSQLWhere + " AND vta_ventas.idven <> 0 " '--SOLO LOS QUE TIENEN VENDEDORES
            
        End If
    
    ElseIf opt_consulta(4).Value = True Then '--X PRODUCTO
        If Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0 Then
            Q_COL_FILA_OCULTA = 1:        Q_COL_FILA = 2:        Q_POSICION_TOTAL = 2:        Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR TIPO PRODUCTO"
            nSQLCampos = "alm_inventario.tippro,  mae_tipoproducto.descripcion "
            nSQLGroupBy = "alm_inventario.tippro,  mae_tipoproducto.descripcion "
            nSQLOrderBy = "mae_tipoproducto.descripcion "
            nSQLWhere = nSQLWhere + " AND alm_inventario.tippro IS NOT NULL " '--SOLO LOS QUE TIENEN TIPO DE PRODUCTO
            
        ElseIf Me.ChkMostrarItem.Value = 1 Then
            Q_COL_FILA_OCULTA = 2:        Q_COL_FILA = 4:        Q_POSICION_TOTAL = 4:        Q_COL_COMPARAR_GRUPO = 2
            T_RPT_TITULO = "RESUMEN DE VENTAS POR TIPO PRODUCTO CON ITEM"
            nSQLCampos = "alm_inventario.tippro,alm_inventario.id,  mae_tipoproducto.descripcion,alm_inventario.descripcion AS desctipcom "
            nSQLGroupBy = "alm_inventario.tippro,alm_inventario.id,  mae_tipoproducto.descripcion,alm_inventario.descripcion "
            nSQLOrderBy = "mae_tipoproducto.descripcion,alm_inventario.descripcion "
            nSQLWhere = nSQLWhere + " AND alm_inventario.tippro IS NOT NULL " '--SOLO LOS QUE TIENEN TIPO DE PRODUCTO
            
        Else '--X FAMILIA
            Q_COL_FILA_OCULTA = 1:        Q_COL_FILA = 2:        Q_POSICION_TOTAL = 2:        Q_COL_COMPARAR_GRUPO = -1
            T_RPT_TITULO = "RESUMEN DE VENTAS POR FAMILIA"
            nSQLCampos = "mae_familia.id,  mae_familia.descripcion "
            nSQLGroupBy = "mae_familia.id,  mae_familia.descripcion "
            nSQLOrderBy = "mae_familia.descripcion  "
            nSQLWhere = nSQLWhere + " AND alm_inventario.idfam IS NOT NULL " '--SOLO LOS QUE TIENEN TIPO DE PRODUCTO
            
        End If
    End If
    
    '--DE LA CANTIDAD DE COL
    Q_COL_FILA = Q_COL_FILA + 1
    Q_POS_MES_INICIO = Q_COL_FILA '--Q_COL_FILA + CAMPO_TOTAL
    '------------------------------------------
    If opt_estilo(0).Value = True Then '--MES
        nSQLPivot = "FORMAT(vta_ventas.fchdoc,'m') "
    ElseIf opt_estilo(1).Value = True Then '--TRIMESTRE
        nSQLPivot = "FORMAT(vta_ventas.fchdoc,'q') "
    ElseIf opt_estilo(2).Value = True Then '--SEMESTRE
        nSQLPivot = "FORMAT(vta_ventas.fchdoc,'s') "
    End If
    '--DEL PIVOT
    For k = 0 To UBound(ARR_TMP)
        nSQLPivotSalida = nSQLPivotSalida + ARR_TMP(k, 2) + ","
    Next k
    nSQLPivotSalida = " IN (" + Left(nSQLPivotSalida, Len(nSQLPivotSalida) - 1) + ") "
    nSQLWhere = nSQLWhere + " AND " + nSQLPivot + nSQLPivotSalida
    'nSQLPivotSalida = " In ('Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic');"
    
    '-------
    
    '-----------------------------------------
    If Me.opt_mon(2).Value = True Then '--TODO EN SOLES
        nSQLValor = " Sum(IIf(vta_ventas.idmon=1,vta_ventas.imptotdoc, IIf((SELECT con_tc.impven From con_tc WHERE con_tc.idmon=2 AND con_tc.fecha=vta_ventas.fchdoc) Is Null,0,(SELECT con_tc.impven From con_tc Where con_tc.idmon = 2 And con_tc.fecha = vta_ventas.fchdoc )*vta_ventas.imptotdoc))) "
        nSQLGroupBy = nSQLGroupBy + ", vta_ventas.fchdoc "
        
    ElseIf Me.opt_mon(3).Value = True Then '--TODO EN DOLARES
        nSQLValor = " Sum(IIf(vta_ventas.idmon=2,vta_ventas.imptotdoc,IIf((SELECT con_tc.impven From con_tc WHERE con_tc.idmon=2 AND con_tc.fecha=vta_ventas.fchdoc) Is Null,0,vta_ventas.imptotdoc/(SELECT con_tc.impven From con_tc Where con_tc.idmon = 2 And con_tc.fecha = vta_ventas.fchdoc )))) "
        nSQLGroupBy = nSQLGroupBy + ", vta_ventas.fchdoc "
        
    End If
    
    '--DEL TIPO DE IMPORTE
    If opt_mon(2).Value = True Or opt_mon(3).Value = True Then
        If opt_importe(0).Value = True Then
            nSQLValor = Replace(nSQLValor, "vta_ventas.imptotdoc", "vta_ventas.imptotdoc")
        ElseIf opt_importe(1).Value = True Then
            nSQLValor = Replace(nSQLValor, "vta_ventas.imptotdoc", "vta_ventas.impbru")
        Else
            nSQLValor = Replace(nSQLValor, "vta_ventas.imptotdoc", "vta_ventas.impigv")
        End If
    Else
        If opt_importe(0).Value = True Then
            nSQLValor = " SUM(vta_ventas.imptotdoc) "
        ElseIf opt_importe(1).Value = True Then
            nSQLValor = " SUM(vta_ventas.impbru) "
        Else
            nSQLValor = " SUM(vta_ventas.impigv) "
        End If
    End If
        
        
    '--DEL FROM ---
    '--SELECC X PTO VENTA O SELECCI. ALGUN REGISTRO DE PTO DEVENTA
    If opt_consulta(2).Value = True Or vStrFiltro_PTO_VTA <> "" Then
        nSQLFrom = " ((((vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id) RIGHT JOIN (((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) INNER JOIN vta_guia ON vta_ventas.id = vta_guia.iddocven) LEFT JOIN vta_puntoVenta ON vta_guia.idpunven = vta_puntoVenta.id) ON vta_vendedores.id = vta_ventas.idven) INNER JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta) INNER JOIN vta_guiadet ON vta_guia.id = vta_guiadet.idgui) INNER JOIN (alm_inventario INNER JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) ON (vta_ventasdet.iditem = alm_inventario.id) AND (vta_guiadet.iditem = alm_inventario.id) "
        '--------
        If Me.opt_mon(2).Value = True Or Me.opt_mon(3).Value = True Then
            nSQLValor = Replace(nSQLValor, "vta_ventas.imptotdoc", "(vta_guiadet.canpro * vta_ventasdet.preuni)")
        Else
            If opt_totalizar(0).Value = True Then
                nSQLValor = " SUM(vta_guiadet.canpro * vta_ventasdet.preuni) " 'IMPORTE
            Else
                nSQLValor = " SUM(vta_guiadet.canpro) " 'CANTIDAD
            End If
        End If
        '--SE BUSCARA POR FECHA DE LA GUIA
        nSQLWhere = Replace(nSQLWhere, "vta_ventas.fchdoc", "vta_guia.fecgiro")
        nSQLPivot = Replace(nSQLPivot, "vta_ventas.fchdoc", "vta_guia.fecgiro")
        
    '--SELECC X T. PROD/ITEM O SELECC. T.PROD O MOSTRAR ITEM
    ElseIf opt_consulta(4).Value = True Or Me.TxtIdTipProd.Text <> "" Or Me.ChkMostrarItem.Value = 1 Then
        nSQLFrom = " (((vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id) RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON vta_vendedores.id = vta_ventas.idven) LEFT JOIN ((alm_inventario RIGHT JOIN vta_ventasdet ON alm_inventario.id = vta_ventasdet.iditem) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id "
    
        If Me.opt_mon(2).Value = True Or Me.opt_mon(3).Value = True Then
            nSQLValor = Replace(nSQLValor, "vta_ventas.imptotdoc", "vta_ventasdet.imptot")
                
        Else
            If opt_totalizar(0).Value = True Then
                nSQLValor = "SUM(vta_ventasdet.imptot) " 'IMPORTE
            Else
                nSQLValor = "SUM(vta_ventasdet.canpro) " 'CANTIDAD
            End If
            
        End If
    '--SELECC. X AÑO, X CLIENTE, X VENDEDOR, CON AGREGAR REG. CLIENTE,VENDEDOR
    Else
        If opt_totalizar(0).Value = True Then '--IMPORTE
            nSQLFrom = " (vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id) RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON vta_vendedores.id = vta_ventas.idven "
        Else 'CANTIDAD
            nSQLFrom = " (((vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id) RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON vta_vendedores.id = vta_ventas.idven) LEFT JOIN ((alm_inventario RIGHT JOIN vta_ventasdet ON alm_inventario.id = vta_ventasdet.iditem) LEFT JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id "
            nSQLValor = " SUM(vta_ventasdet.canpro) "
        End If

    End If
           
    '--GENERANDO LA CONSULTA
    vStrSelect = " TRANSFORM " + nSQLValor + _
        vbCr + " SELECT " + nSQLCampos + "," + nSQLValor + " AS total " + _
        vbCr + " FROM " + nSQLFrom + _
        vbCr + " WHERE (((vta_ventas.anulado) = 0)) " + nSQLWhere + _
        vbCr + " GROUP BY " + nSQLGroupBy + _
        vbCr + " ORDER BY " + nSQLOrderBy + _
        vbCr + " PIVOT " + nSQLPivot + nSQLPivotSalida
    If F_TODO_SOL_DOL = True Then
    '--ESTA CONSULTA NOS SERVIRA CUNADO LA SELECCION DE MONEDA SEA TODOS EN S/. Ó $
    vStrSelect = "SELECT DISTINCT " + nSQLCampos + _
        vbCr + " FROM " + nSQLFrom + _
        vbCr + " WHERE (((vta_ventas.anulado) = 0)) " + nSQLWhere + _
        vbCr + " GROUP BY " + nSQLGroupBy + _
        vbCr + " ORDER BY " + nSQLOrderBy
    End If
    '------------------------------------------------------------------------------------
    GENERAR_CONSULTA = vStrSelect
        
End Function


'--011007
Private Sub Limpiar_ARRAY_TOTAL(Optional F_LIMPIA_TOT_GRL As Boolean = False)
    Erase Arr_Totales_col
    ReDim Arr_Totales_col(13, 0) As Double
    If F_LIMPIA_TOT_GRL = True Then
        Erase Arr_Totales_cols
        ReDim Arr_Totales_cols(13, 0)
    End If
End Sub
'''
Private Sub CARGAR_DATOS_GRILLA_ADD_TOTALES(BAND_ADD_TOTAL As Boolean, Nombre_total As String, _
                Optional Band_Total_gral As Boolean = False, _
                Optional band_forzar_suma As Boolean = False, Optional M_ANYO As String, _
                Optional F_OTROS_ANYOS As Boolean = False)
    Dim Q_MES As Integer
    '--AGREGA LOS TOTALES POR CADA GRUPO Y EL TOTAL GENERAL
    '--ACUMULA LOS TOTALES EN EL TOTAL GENERAL
    Dim X_ROW As Long
    'On Error Resume Next
    If F_OTROS_ANYOS = False Then
        X_ROW = Fg1.Rows
        If BAND_ADD_TOTAL = True Then
            '--AGREAGNDO NUEVA FILA
            ADD_REG Fg1, IIf(Band_Total_gral = False, Fila_Total, Fila_Total_grl)
            'PONIENDO LOS NOMBRES DE LOS TOTALES  Q_POSICION_TOTAL SE OBTIENE DE GENERAR_CONSULTA()
            Fg1.TextMatrix(X_ROW, Q_POSICION_TOTAL) = Nombre_total
            FORMATO_CELDA Fg1, X_ROW, Q_POSICION_TOTAL
        End If
    Else
        X_ROW = Fg1.Row
    End If

    '--ACUMULANDO LOS TOTALES GRLES
    If Me.opt_consulta(0).Value = True Then     '--X AÑO
        If Band_Total_gral = True Or (Me.TxtIdTipProd.Text <> "" Or Me.ChkMostrarItem.Value = 1) Then
            For Q_MES = 0 To UBound(Arr_Totales_col())
                Arr_Totales_cols(Q_MES, 0) = Arr_Totales_cols(Q_MES, 0) + Arr_Totales_col(Q_MES, 0)
            Next Q_MES
        End If
    Else
        If Band_Total_gral = False Then     '--DEMAS
            For Q_MES = 0 To UBound(Arr_Totales_col())
                Arr_Totales_cols(Q_MES, 0) = Arr_Totales_cols(Q_MES, 0) + Arr_Totales_col(Q_MES, 0)
            Next Q_MES
        End If
    End If
    '
'--------------------------
    Dim Q_INCREMENTO_X_COL As Integer   '--SERVIRA PARA POSICIONAR LA PRIMERA COLUMNA DE ENERO DE UN AÑO
    Dim Q_POS_MES_TOTAL  As Integer     '--POSICIONA A LA COLUMNA DEL TOTAL X AÑO
    
    For Q_MES = 0 To UBound(ARR_ANYO) - 1
        If ARR_ANYO(Q_MES) = M_ANYO Then
            Q_INCREMENTO_X_COL = Q_MES
            Exit For
        End If
    Next
    '--IDENTIFICAR LA POSICION DE INICIO DE MES(ENERO)
    If Me.opt_consulta(0).Value = True Then Q_INCREMENTO_X_COL = 0
    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    '-----------
'--DE LOS MESES
    Q_POS_MES = Q_POS_MES_INICIO + Q_INCREMENTO_X_COL
    
    '--DE LOS MESES
    For Q_MES = 0 To Q_COL_ARR_TOTAL
        '--INTERRUMPIR EL PROCESO
        If BAND_INTERRUMPIR = True Then Exit Sub
        Fg1.TextMatrix(X_ROW, Q_POS_MES) = CONVERTIR_A_ESCALA(IIf(Band_Total_gral = False, Arr_Totales_col(Q_MES, 0), Arr_Totales_cols(Q_MES, 0)))
        FORMATO_CELDA Fg1, X_ROW, Q_POS_MES
        If Me.opt_consulta(0).Value = True Then
            Q_POS_MES = Q_POS_MES + 1
        Else
            Q_POS_MES = Q_POS_MES + Q_TOTAL_ANYO
        End If
    Next Q_MES
       
    For Q_MES = Q_COL_ARR_TOTAL + 1 To Q_COL_ARR_TOTAL + 2
        If Q_MES = Q_COL_ARR_TOTAL + 1 Then '--TOTAL
            If Me.opt_consulta(0).Value = True Then
                Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL + 1) * 1
            Else
                Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL + 1) * Q_TOTAL_ANYO + Q_INCREMENTO_X_COL
            End If
            Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL) = CONVERTIR_A_ESCALA(IIf(Band_Total_gral = False, Arr_Totales_col(Q_MES, 0), Arr_Totales_cols(Q_MES, 0)))
        ElseIf Q_MES = Q_COL_ARR_TOTAL + 2 Then '--TOTAL GRAL
            Q_POS_MES_TOTAL = Fg1.Cols - 1
            If IsNumeric(Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL)) = False Then
                Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL) = CONVERTIR_A_ESCALA(IIf(Band_Total_gral = False, Arr_Totales_col(Q_MES, 0), Arr_Totales_cols(Q_MES, 0)))
            Else
                If Me.opt_consulta(0).Value = True Then
                    Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL) = CONVERTIR_A_ESCALA(IIf(Band_Total_gral = False, Arr_Totales_col(Q_MES, 0), Arr_Totales_cols(Q_MES, 0)))
                Else
                    Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL) = CDbl(Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL)) + CONVERTIR_A_ESCALA(IIf(Band_Total_gral = False, Arr_Totales_col(Q_MES, 0), Arr_Totales_cols(Q_MES, 0)))
                End If
            End If
            
            Fg1.TextMatrix(X_ROW, Fg1.Cols - 1) = CONVERTIR_A_ESCALA(CDbl(Fg1.TextMatrix(X_ROW, Fg1.Cols - 1)), True)
        End If
        
        FORMATO_CELDA Fg1, X_ROW, Q_POS_MES_TOTAL
        
    Next Q_MES

    Err.Clear
End Sub


Private Sub Configurar_Grilla(Optional F_CONSERVAR_FORMATO As Boolean = False)
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    '--DE ACUERDO A LO QUE SE SELECCIONA
    Dim M_ANCHO_COL_MES As Integer '--DEPENDERA DEL TIPO DE PRESENTACION
                                    '--EN DECIMALES, EN MILES
    Dim k&, j&
    
    
    If F_CONSERVAR_FORMATO = True Then Fg1.Clear
    
    Fg1.FrozenCols = 0
    
    If opt_estilo(0).Value = True Then '--MES
        M_ANCHO_COL_MES = 800
    ElseIf opt_estilo(1).Value = True Then '--TRIMESTRE
        M_ANCHO_COL_MES = 900
    ElseIf opt_estilo(2).Value = True Then '--SEMESTRE
        M_ANCHO_COL_MES = 1000
    End If
    
    
    If Me.opt_escala(0).Value = True Then
        M_ANCHO_COL_MES = M_ANCHO_COL_MES + 250
    Else
        M_ANCHO_COL_MES = M_ANCHO_COL_MES
    End If
    
    With Fg1
        '-----
        '--DATOS DE FILA
        If opt_consulta(0).Value = True Then
        
            Fg1.Cols = Q_COL_FILA + (Q_COL_ARR_TOTAL + 1) + 1
            UNIR_CELDAS Fg1, 0, Q_COL_FILA, 0, Fg1.Cols - 1, " ", flexAlignCenterTop
            '--DATOS DE FILA
            .ColAlignment(2) = flexAlignLeftCenter
            .TextMatrix(1, 2) = "Año":         .ColWidth(2) = M_ANCHO_COL_MES
                        
            Q_POS_MES = Q_POS_MES_INICIO
            '--DATOS DE COLUMNAS
            For k = 0 To Q_COL_ARR_TOTAL '--MESES DEL AÑO
                '--COLOCANDO LOS MESES
                UNIR_CELDAS Fg1, 1, Q_POS_MES, 1, Q_POS_MES, ARR_TMP(k, 1), flexAlignCenterTop: .ColWidth(k) = M_ANCHO_COL_MES
                .ColAlignment(Q_POS_MES) = flexAlignRightBottom
                .Row = 0:   .Col = Q_POS_MES:   .CellAlignment = flexAlignCenterBottom
                Q_POS_MES = Q_POS_MES + 1
            Next k
            '--COLOCANDO EL TOTAL
            .TextMatrix(1, .Cols - 1) = "Total Gral":         .ColWidth(.Cols - 1) = M_ANCHO_COL_MES + 200
        Else
            '--CANTIDAD DE COLUMNAS
            Fg1.Cols = Q_COL_FILA + ((Q_COL_ARR_TOTAL + 2) * Q_TOTAL_ANYO) + 1
                                    '--total_mes+total_años
            '---
            If opt_consulta(1).Value = True Then '--X CLIENTE
                If (Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0) Then '--CLIETNE/PRODUCTO
                    .TextMatrix(1, 3) = "Cliente":       .ColWidth(3) = 1500:   .ColAlignment(3) = flexAlignLeftBottom
                    .TextMatrix(1, 4) = "Producto":      .ColWidth(4) = 1000:   .ColAlignment(4) = flexAlignLeftBottom
                    .Row = 1:   .Col = 4:  .CellAlignment = flexAlignLeftBottom
                ElseIf Me.ChkMostrarItem.Value = 1 Then '--CLIETNE/PRODUCTO/ITEM
                    .TextMatrix(1, 4) = "Cliente":       .ColWidth(4) = 1500:   .ColAlignment(4) = flexAlignLeftBottom
                    .TextMatrix(1, 5) = "Producto":      .ColWidth(5) = 1000:   .ColAlignment(5) = flexAlignLeftBottom
                    .TextMatrix(1, 6) = "Item":          .ColWidth(6) = 2000:   .ColAlignment(6) = flexAlignLeftBottom
                    .Row = 1:   .Col = 5:  .CellAlignment = flexAlignLeftBottom
                    .Row = 1:   .Col = 6:  .CellAlignment = flexAlignLeftBottom
                Else    '--SOLO CLIENTE
                    .TextMatrix(1, 2) = "Cliente":       .ColWidth(2) = 2500:   .ColAlignment(2) = flexAlignLeftBottom
                    .Row = 1:   .Col = 2:  .CellAlignment = flexAlignLeftBottom
                End If
                
            ElseIf opt_consulta(2).Value = True Then '--X PTO DE VENTA
                If (Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0) Then '--CLIENTE/PTO DE VENTA/PRODUCTO
                    .TextMatrix(1, 4) = "Cliente":        .ColWidth(4) = 1500:   .ColAlignment(4) = flexAlignLeftBottom
                    .TextMatrix(1, 5) = "Punto de Venta": .ColWidth(5) = 1800:   .ColAlignment(5) = flexAlignLeftBottom
                    .TextMatrix(1, 6) = "Producto":       .ColWidth(6) = 1000:   .ColAlignment(6) = flexAlignLeftBottom
                    .Row = 1:   .Col = 5:  .CellAlignment = flexAlignLeftBottom
                    .Row = 1:   .Col = 6:  .CellAlignment = flexAlignLeftBottom
                ElseIf Me.ChkMostrarItem.Value = 1 Then '--CLIENTE/PTO DE VENTA/PRODUCTO/ITEM
                    .TextMatrix(1, 5) = "Cliente":        .ColWidth(5) = 1500:   .ColAlignment(5) = flexAlignLeftBottom
                    .TextMatrix(1, 6) = "Punto de Venta": .ColWidth(6) = 1800:   .ColAlignment(6) = flexAlignLeftBottom
                    .TextMatrix(1, 7) = "Producto":       .ColWidth(7) = 1000:   .ColAlignment(7) = flexAlignLeftBottom
                    .TextMatrix(1, 8) = "Item":           .ColWidth(8) = 2000:   .ColAlignment(8) = flexAlignLeftBottom
                    .Row = 1:   .Col = 6:  .CellAlignment = flexAlignLeftBottom
                    .Row = 1:   .Col = 7:  .CellAlignment = flexAlignLeftBottom
                    .Row = 1:   .Col = 8:  .CellAlignment = flexAlignLeftBottom
                Else    '--SOLO PTO DE VENTA
                .TextMatrix(1, 3) = "Cliente":            .ColWidth(3) = 2000:   .ColAlignment(3) = flexAlignLeftBottom
                .TextMatrix(1, 4) = "Punto de Venta":     .ColWidth(4) = 2500:   .ColAlignment(4) = flexAlignLeftBottom
                .Row = 1:   .Col = 3:  .CellAlignment = flexAlignLeftBottom
                .Row = 1:   .Col = 4:  .CellAlignment = flexAlignLeftBottom
                End If
                
            ElseIf opt_consulta(3).Value = True Then '--X VENDEDOR
                If (Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0) Or (Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0) Then '--VENDEDOR/PRODUCTO
                    .TextMatrix(1, 3) = "Vendedor":       .ColWidth(3) = 1500:   .ColAlignment(3) = flexAlignLeftBottom
                    .TextMatrix(1, 4) = "Producto":       .ColWidth(4) = 1000:   .ColAlignment(4) = flexAlignLeftBottom
                    .Row = 1:   .Col = 4:  .CellAlignment = flexAlignLeftBottom
                ElseIf Me.ChkMostrarItem.Value = 1 Then '--VENDEDOR/PRODUCTO/ITEM
                    .TextMatrix(1, 4) = "Vendedor":       .ColWidth(4) = 1500:   .ColAlignment(4) = flexAlignLeftBottom
                    .TextMatrix(1, 5) = "Producto":       .ColWidth(5) = 1000:   .ColAlignment(5) = flexAlignLeftBottom
                    .TextMatrix(1, 6) = "Item":           .ColWidth(6) = 2000:   .ColAlignment(6) = flexAlignLeftBottom
                    .Row = 1:   .Col = 5:  .CellAlignment = flexAlignLeftBottom
                    .Row = 1:   .Col = 6:  .CellAlignment = flexAlignLeftBottom
                Else    '--SOLO VENDEDOR
                    .TextMatrix(1, 2) = "Vendedor":       .ColWidth(2) = 2000:   .ColAlignment(2) = flexAlignLeftBottom
                    .Row = 1:   .Col = 2:  .CellAlignment = flexAlignLeftBottom
                End If
                
            ElseIf opt_consulta(4).Value = True Then '--X PRODUCTO / ITEM
                If Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0 Then
                    .TextMatrix(1, 2) = "Producto":      .ColWidth(2) = 1000:   .ColAlignment(2) = flexAlignLeftBottom
                    .Row = 1:   .Col = 2:  .CellAlignment = flexAlignLeftBottom
                ElseIf Me.ChkMostrarItem.Value = 1 Then
                    .TextMatrix(1, 3) = "Producto":      .ColWidth(3) = 1000:   .ColAlignment(3) = flexAlignLeftBottom
                    .TextMatrix(1, 4) = "Item":          .ColWidth(4) = 2000:   .ColAlignment(4) = flexAlignLeftBottom
                    .Row = 1:   .Col = 3:  .CellAlignment = flexAlignLeftBottom
                    .Row = 1:   .Col = 4:  .CellAlignment = flexAlignLeftBottom
                Else
                    .TextMatrix(1, 2) = "Familia":       .ColWidth(2) = 2000:   .ColAlignment(2) = flexAlignLeftBottom
                    .Row = 1:   .Col = 2:  .CellAlignment = flexAlignLeftBottom
                End If
            End If
            Q_POS_MES = Q_POS_MES_INICIO
            '--DATOS DE COLUMNAS
            For k = 0 To Q_COL_ARR_TOTAL + 1 '--MESES DEL AÑO + TOTAL
                '--COLOCANDO LOS MESES Y AGRUPANDOLOS
                If k = Q_COL_ARR_TOTAL + 1 Then
                    UNIR_CELDAS Fg1, 0, Q_POS_MES, 0, Q_POS_MES + Q_TOTAL_ANYO - 1, "Totales", flexAlignRightBottom
                Else
                    UNIR_CELDAS Fg1, 0, Q_POS_MES, 0, Q_POS_MES + Q_TOTAL_ANYO - 1, IIf(Q_TOTAL_ANYO > 1, ARR_TMP(k, 0), ARR_TMP(k, 1)), flexAlignRightBottom
                End If
                .ColAlignment(Q_POS_MES) = flexAlignRightBottom
                .Row = 0:   .Col = Q_POS_MES:   .CellAlignment = flexAlignCenterBottom
    '            M_ANCHO_COL_MES = 5
                '--COLOCANDO LOS AÑOS
                For j = 0 To Q_TOTAL_ANYO - 1 '--CANTIDAD DE AÑOS SELECCIONADOS
                    If k = Q_COL_ARR_TOTAL + 1 Then
                        .ColWidth(Q_POS_MES + j) = M_ANCHO_COL_MES + 200 '--DE LOS AÑOS
                    Else
                        .ColWidth(Q_POS_MES + j) = M_ANCHO_COL_MES  '--DE LOS MESES
                    End If
                    UNIR_CELDAS Fg1, 1, Q_POS_MES + j, 1, Q_POS_MES + j, ARR_ANYO(j), flexAlignCenterCenter
                    .Row = 1:   .Col = Q_POS_MES + j:   .CellAlignment = flexAlignCenterCenter
                Next j
                
                Q_POS_MES = Q_POS_MES + Q_TOTAL_ANYO
                
            Next k
            '--COLOCANDO LOS TOTALES
            .ColWidth(.Cols - 1) = M_ANCHO_COL_MES + 400
            UNIR_CELDAS Fg1, 0, .Cols - 1, 0, .Cols - 1, "Total Gral.", flexAlignCenterTop
            'DEL TOTAL GRAL
            UNIR_CELDAS Fg1, 1, .Cols - 1, 1, .Cols - 1, "Total", flexAlignCenterTop
           
            '--OCULTAR EL GRUPO
            .ColWidth(Q_COL_COMPARAR_GRUPO + 1) = 0
           
        End If
        .FrozenCols = Q_POS_MES_INICIO - 1
        .ColWidth(0) = 200
        '--DE LOS ID'S
        For k = 1 To Q_COL_FILA_OCULTA
            .TextMatrix(1, k) = "ID" + CStr(k):         .ColWidth(k) = 500
        Next
        If Q_COL_FILA_OCULTA <> -1 Then OCULTAR_COL Fg1, 1, Q_COL_FILA_OCULTA

    End With
    DoEvents
End Sub



Sub PosicionarProgBar()
'--POSICIONAR EL PROGRESO DENTRO DEL FORMULARIO
'    FraProgreso.Width = 5820
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    FraProgreso.Visible = True
End Sub




Private Function CONVERTIR_A_ESCALA(S_MONTO As Double, _
                                    Optional SOLO_FORMAT As Boolean = False) As String
                                    
    '--ESTA FUNCION CONVERTIRA A LA ESCALA QUE SELECCIONA EL USUARIO
    '--EN DECIMALES O MILES
    If S_MONTO = 0 Then
        If opt_totalizar(0).Value = True Then
            CONVERTIR_A_ESCALA = "0.00"
        Else
            CONVERTIR_A_ESCALA = "0"
        End If
        Exit Function
    End If
    
    If opt_totalizar(0).Value = True Then '--MONTO
        If SOLO_FORMAT = True Then
            CONVERTIR_A_ESCALA = Format(S_MONTO, FORMAT_MONTO)
            Exit Function
        End If
        If Me.opt_escala(0).Value = True Then
            CONVERTIR_A_ESCALA = Format(S_MONTO, FORMAT_MONTO)
        Else
            CONVERTIR_A_ESCALA = Format(S_MONTO / 1000, FORMAT_MONTO)
        End If
    Else '--CANTIDAD
        If SOLO_FORMAT = True Then
            CONVERTIR_A_ESCALA = Format(S_MONTO, FORMAT_CANTIDAD)
            Exit Function
        End If
        If Me.opt_escala(0).Value = True Then
            CONVERTIR_A_ESCALA = Format(S_MONTO, FORMAT_CANTIDAD)
        Else
            CONVERTIR_A_ESCALA = Format(S_MONTO / 1000, FORMAT_CANTIDAD)
        End If
    End If
    
    
    
End Function


'---DEL GRAFICO
'--251007


Private Sub CmdGrafAcep1_Click()
    If OptTipGrafBarra1.Value = True Then
        vLngTipoGrafico = 51
    ElseIf OptTipGrafLinea.Value = True Then
        vLngTipoGrafico = 65
    ElseIf OptTipGrafCircular.Value = True Then
        vLngTipoGrafico = 5
    End If
    
    If OptConDatoResum1.Value = True Then
        vTipoDato = 0
    ElseIf OptconDatosDetalle1.Value = True Then
        vTipoDato = 1
    End If
    
    If ChkLeyenda.Value = 1 Then
        vViewLeyenda = True
    Else
        vViewLeyenda = False
    End If
    
    GrafEstilo_TotGral_0_1
    FraGraf1.Visible = False
End Sub

Private Sub CmdGrafCancel1_Click()
    FraGraf1.Visible = False
End Sub


Private Function fTituloGrafico() As String
    If OptConDatoResum1.Value = True Then
        fTituloGrafico = "RESUMIDO POR AÑO"
    ElseIf OptconDatosDetalle1.Value = True Then
        fTituloGrafico = "DETALLADO POR AÑO"
    End If
End Function

Private Sub GenerarGraf_TotGral_0_1(pRango As String, pTipoGraf As Long, pTitulo As String, pTipoDato As Integer)
    With Oleapp
        '--MACRO 1
    '    .Sheets("Hoja1").Select
    '    .Sheets("Hoja1").Name = "dato"
    '    .Range(pRango).Select
        .Charts.Add
        '.ActiveChart.ChartType = xlColumnClustered
        .ActiveChart.ChartType = pTipoGraf
        '.ActiveChart.SetSourceData Source:=Sheets("dato").Range("A3:B5"), PlotBy:=xlColumns
        If OptTipGrafLinea.Value = True Then
            .ActiveChart.SetSourceData Source:=.Sheets("datos").Range(pRango), PlotBy:=1
        Else
            If OptconDatosDetalle1.Value = True Then
                .ActiveChart.SetSourceData Source:=.Sheets("datos").Range(pRango), PlotBy:=1
            Else
                .ActiveChart.SetSourceData Source:=.Sheets("datos").Range(pRango), PlotBy:=2
            End If
        End If
        '.ActiveChart.Location Where:=xlLocationAsNewSheet
        .ActiveChart.Location Where:=1
'        If pTipoDato = 1 Then
'            ActiveChart.HasLegend = True
'        End If
        '----
        Select Case pTipoGraf
            Case 51 'BARRAS
                If pTipoDato = 0 Then
                    .ActiveChart.ChartArea.Select
                    .ActiveChart.ApplyDataLabels Type:=2, LegendKey:=False
                End If
            Case 5 'CIRCULAR
                .ActiveChart.HasLegend = True
                .ActiveChart.Legend.Select
                .Selection.Position = -4152
                .ActiveChart.ApplyDataLabels Type:=3, LegendKey:=True _
                    , HasLeaderLines:=True
        End Select
        '-----
        '--PONER TITULO
        .ActiveChart.ChartArea.Select
        With .ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = pTitulo
        End With
        On Error Resume Next
        .ActiveChart.ChartArea.Select
        .ActiveChart.HasLegend = vViewLeyenda
        
    End With
End Sub

Private Sub GrafEstilo_TotGral_0_1()
    'GRAFICO POR ANIO POR TOTAL GENERAL
    Dim i_row As Long, i_col As Long, fs As Variant, NFILA As Long
    Dim nArchivo As String, NCOLUMN As Long, vRangSelect As String
    Dim vColTotMesAnio As Long, vIniCol_Grilla As Integer, vColIndexVarible As Long
    'VARIABLES PARA TRABAJAR CON LA SELECCION DE CELDAS DE EXCEL
    Dim vRango1 As String, vRango2 As String, vRangoCelSelecTotal As String
    '----------------------------------------------------------
    Set fs = CreateObject("Scripting.FileSystemObject")
    If vTipoDato = 0 Then
        nArchivo = "C:\grafico_x_anio.XLS"
    Else
        nArchivo = "C:\grafico_x_anio_Detallado.XLS"
    End If
    Set Oleapp = CreateObject("excel.application")
    Oleapp.Visible = True
    With Oleapp
        .WindowState = 1
        .Workbooks.Add
        .Sheets(1).Select
        .Sheets(1).Name = "datos"
                
        NCOLUMN = 2 'COLUMNA INICIO PARA EXCEL
        vIniCol_Grilla = 3
        
        If vTipoDato = 0 Then
            vColTotMesAnio = vCantMeses + 1 'MESES + 1(TOTAL GENERAL)
        Else
            vColTotMesAnio = vCantMeses
        End If
        '--LE SUMA EL VALOR DE INICIO DE LA COLUMNA DE INICIO
        vColTotMesAnio = vColTotMesAnio + vIniCol_Grilla - 1
        '--PONEL EL ENCABEZADO DEL TOTAL GRAL. O DE LOS MESES
        If vTipoDato = 0 Then 'SOLO CON TOTAL GENERAL
            .Cells(3, NCOLUMN) = Fg1.TextMatrix(1, vColTotMesAnio)
        Else 'PARA DETALLADO
            For i_col = vIniCol_Grilla To vColTotMesAnio
                .Cells(3, NCOLUMN) = Fg1.TextMatrix(1, i_col)
                NCOLUMN = NCOLUMN + 1
            Next
        End If
        
        '--PONE LOS ANIO COMO REGISTROS EN LA COLUMNA 1 EN EXCEL
        NFILA = 4
        NCOLUMN = 2 'COLUMNA INICIO PARA EXCEL
        For i_row = 2 To Fg1.Rows - 2
            .Cells(NFILA, 1) = Trim(Fg1.TextMatrix(i_row, 2))
            NFILA = NFILA + 1
        Next
        
        'LLENAR LOS DATOS DEL DETALLE DE LA GRILLA
        If vTipoDato = 0 Then 'SOLO POR TOT GENERAL
            vColTotMesAnio = vCantMeses + 1 'MESES + 1(TOTAL GENERAL)
        Else 'SOLO PARA DETALLADO
            vColTotMesAnio = vCantMeses
        End If
        '--LE SUMA EL VALOR DE INICIO DE LA COLUMNA DE INICIO
        vColTotMesAnio = vColTotMesAnio + vIniCol_Grilla - 1
        NCOLUMN = 2 'COLUMNA INICIO PARA EXCEL
        NFILA = 4
        For i_row = 2 To Fg1.Rows - 2
            If vTipoDato = 0 Then
                .Cells(NFILA, NCOLUMN) = Fg1.TextMatrix(i_row, vColTotMesAnio)
                NFILA = NFILA + 1
            ElseIf vTipoDato = 1 Then
                NCOLUMN = 2
                For i_col = vIniCol_Grilla To vColTotMesAnio
                    .Cells(NFILA, NCOLUMN) = Fg1.TextMatrix(i_row, i_col)
                    NCOLUMN = NCOLUMN + 1
                Next
                NFILA = NFILA + 1
            End If
        Next
        '--GENERA EL GRAFICO
        vRango1 = .Cells(3, 1).Address
        If vTipoDato = 0 Then
            vRango2 = .Cells(NFILA - 1, 2).Address
        Else
            vColTotMesAnio = vCantMeses + 1
            vRango2 = .Cells(NFILA - 1, vColTotMesAnio).Address
        End If
        vRangSelect = vRango1 & ":" & vRango2
        
        vTituloGraf = fTituloGrafico
        'vLngTipoGrafico = 51 barras
'        vLngTipoGrafico = 5 'pie
        GenerarGraf_TotGral_0_1 vRangSelect, vLngTipoGrafico, vTituloGraf, vTipoDato
'        Oleapp.ActiveWorkbook.SaveAs (nArchivo)
        Oleapp.WindowState = 1
        '.ActiveWindow.Zoom = 75
    End With
'    vRangSelect = "A" & CStr(3) & ":M" & CStr(NFILA)
'    GeneraGrafico vRangSelect, "Grafico por Año"
'    Oleapp.Quit
    Set Oleapp = Nothing   ' la aplicación; después libera la referenci
    Set fs = Nothing
    MsgBox "Los datos han sido exportados correctamente", vbInformation, "Aviso"
End Sub

Private Sub UnirCeldaEnExcel(pRango As String)
    With Oleapp
        .Range(pRango).Select
        With .Selection
            .HorizontalAlignment = -4108
            .VerticalAlignment = -4107
            .WrapText = False
            .Orientation = 0
            .ShrinkToFit = False
            .MergeCells = False
        End With
        .Selection.Merge
    End With
End Sub

Sub GraficoEstilo_0() 'SOLO POR ANIO
    Dim i_row As Long, i_col As Long, fs As Variant, NFILA As Long
    Dim nArchivo As String, NCOLUMN As Long, vRangSelect As String
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    nArchivo = "C:\XANIO_MESES_GRAFIC.XLS"
    
    Set Oleapp = CreateObject("excel.application")
    With Oleapp
        .Workbooks.Add
        .Sheets(1).Select
        .Sheets(1).Name = "datos"
               
'        .CELLS(1, 3).Value = "PADRÓN DE ARTICULOS"
        NCOLUMN = 2
        For i_col = 2 To Fg1.Cols - 2
            .Cells(3, NCOLUMN) = Fg1.TextMatrix(1, i_col)
            NCOLUMN = NCOLUMN + 1
        Next
               
        NFILA = 4: NCOLUMN = 1
        For i_row = 2 To Fg1.Rows - 2
            NCOLUMN = 1
            For i_col = 1 To Fg1.Cols - 2
                .Cells(NFILA, NCOLUMN) = Fg1.TextMatrix(i_row, i_col)
                NCOLUMN = NCOLUMN + 1
            Next
            NFILA = NFILA + 1
        Next
        NFILA = 3
        For i_row = 2 To Fg1.Rows - 2
            NFILA = NFILA + 1
        Next
        Oleapp.ActiveWorkbook.SaveAs (nArchivo)
    End With
    vRangSelect = "A" & CStr(3) & ":M" & CStr(NFILA)
    GeneraGrafico vRangSelect, "Grafico por Año"
    Oleapp.Quit
    Set Oleapp = Nothing   ' la aplicación; después libera la referenci
    Set fs = Nothing
    MsgBox "Los datos han sido exportados correctamente", vbInformation, "Aviso"
End Sub

Sub GraficoEstilo_1() 'SOLO POR PROVEEDOR
    Dim i_row As Long, i_col As Long, fs As Variant, NFILA As Long
    Dim nArchivo As String, NCOLUMN As Long, vRangSelect As String
    Dim vColTotMesAnio As Long, vIniCol_Grilla As Integer, vColIndexVarible As Long
    'VARIABLES PARA TRABAJAR CON LA SELECCION DE CELDAS DE EXCEL
    Dim vRango1 As String, vRango2 As String, vRangoCelSelecTotal As String
    '----------------------------------------------------------
    Set fs = CreateObject("Scripting.FileSystemObject")
    nArchivo = "C:\grafico_x_proveedor.XLS"
    Set Oleapp = CreateObject("excel.application")
    With Oleapp
        .Workbooks.Add
        .Sheets(1).Select
        .Sheets(1).Name = "datos"

        NCOLUMN = 2 'COLUMNA INICIO PARA EXCEL
        vIniCol_Grilla = 3
        vColTotMesAnio = (vCantMeses * Q_TOTAL_ANYO) - 1 'MESES X ANIOS
        '--PONE EL ENCABEZADO DE ANIOS
        For i_col = vIniCol_Grilla To vColTotMesAnio + vIniCol_Grilla
            .Cells(3, NCOLUMN) = Fg1.TextMatrix(0, i_col)
            .Cells(4, NCOLUMN) = Fg1.TextMatrix(1, i_col)
            NCOLUMN = NCOLUMN + 1
        Next
        '--UNE CELDAS DE LOS MESES
        'ESTA VARIABLE vCantMeses ME INDICA LA CANTIDAD DE MESES SELECCIONADOS
        vColTotMesAnio = (vCantMeses * Q_TOTAL_ANYO)
        vColIndexVarible = 2
        For i_col = 1 To vCantMeses
            vRango1 = .Cells(3, vColIndexVarible).Address
            vRango2 = .Cells(3, vColIndexVarible + (Q_TOTAL_ANYO - 1)).Address
            vRangoCelSelecTotal = vRango1 & ":" & vRango2 'ejemplo B3:C3
            On Error Resume Next
            UnirCeldaEnExcel vRangoCelSelecTotal
            vColIndexVarible = vColIndexVarible + Q_TOTAL_ANYO
        Next
        'LLENAR NOMBRES DE PROVEEDORES
        NFILA = 5
        For i_row = 2 To Fg1.Rows - 1
            .Cells(NFILA, 1) = Trim(Fg1.TextMatrix(i_row, 2))
            NFILA = NFILA + 1
        Next
        'LLENAR LOS DATOS DEL DETALLE DE LA GRILLA
        vColTotMesAnio = (vCantMeses * Q_TOTAL_ANYO) - 1
        NFILA = 5: NCOLUMN = 2
        For i_row = 2 To Fg1.Rows - 1
            NCOLUMN = 2
            For i_col = 3 To (3 + vColTotMesAnio)
                
                .Cells(NFILA, NCOLUMN) = Fg1.TextMatrix(i_row, i_col)
                NCOLUMN = NCOLUMN + 1
            Next
            NFILA = NFILA + 1
        Next
        Oleapp.ActiveWorkbook.SaveAs (nArchivo)
    End With
'    vRangSelect = "A" & CStr(3) & ":M" & CStr(NFILA)
'    GeneraGrafico vRangSelect, "Grafico por Año"
    Oleapp.Quit
    Set Oleapp = Nothing   ' la aplicación; después libera la referenci
    Set fs = Nothing
    MsgBox "Los datos han sido exportados correctamente", vbInformation, "Aviso"
End Sub

Sub GeneraGrafico(pRango As String, pTitGrafico As String)
    With Oleapp
        .Charts.Add
        .ActiveChart.ChartType = 65
        .ActiveChart.SetSourceData Source:=.Sheets("datos").Range(pRango), PlotBy:=1
        .ActiveChart.Location Where:=1
        .ActiveChart.ChartArea.Select
        .Selection.AutoScaleFont = True
        With .Selection.Font
            .Name = "Arial"
            .Size = 8
            .Strikethrough = False
            .Superscript = False
            .Subscript = False
            .OutlineFont = False
            .Shadow = False
            .Underline = -4142
            .ColorIndex = -4105
            .Background = -4105
        End With
        '-----------
'        .ActiveChart.ChartArea.Select
'        .ActiveChart.ApplyDataLabels AutoText:=True, LegendKey:=False, _
'        HasLeaderLines:=False, ShowSeriesName:=False, ShowCategoryName:=False, _
'        ShowValue:=True, ShowPercentage:=False, ShowBubbleSize:=False, Separator _
'        :=" "
        
        'PARA OFF 97
        .ActiveChart.ChartArea.Select
        .ActiveChart.ApplyDataLabels Type:=2, LegendKey:=False
        '------------
        With .ActiveChart
            .HasTitle = True
            .ChartTitle.Characters.Text = pTitGrafico
        End With
    End With
End Sub
'--FIN CODIGO DE GRAFICO------------------------------------------


Private Sub CmdExportar_Click()
    FraGraf1.Visible = False
    EXPORTAR
End Sub


Private Sub EXPORTAR()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios

    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, T_RPT_TITULO, T_RPT_PERIODO, "", "Ventas"
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub



'************************************************

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then CONSULTAR
    If Button.Index = 5 Then EXPORTAR
    If Button.Index = 6 Then pVerGrafico
    If Button.Index = 7 Then pImprimir
    If Button.Index = 9 Then
        BAND_INTERRUMPIR = True
        Unload Me
    End If
End Sub

'************************************************

Private Sub pVerGrafico()
    If Fg1.Rows = 2 Then
        MsgBox "No hay datos para el gráfico.", vbInformation, xTitulo
        Exit Sub
    End If
    
    Dim vEstilo As Integer
'''    vEstilo = ESTILO_CONSULTA
    vCantMeses = Q_COL_ARR_TOTAL + 1
    FraGraf1.Left = (Me.Width - FraGraf1.Width) \ 2
    FraGraf1.Top = (Me.Height - FraGraf1.Height) \ 2
    FraGraf1.Visible = True

End Sub

