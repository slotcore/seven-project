VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomctl.ocx"
Begin VB.Form FrmConsProd_Gerencial 
   Caption         =   "Gestión - Análisis de Producción"
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   615
   ClientWidth     =   12435
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   12435
   Begin VB.Frame FraGraf1 
      Height          =   2325
      Left            =   4440
      TabIndex        =   30
      Top             =   2790
      Visible         =   0   'False
      Width           =   3525
      Begin VB.CommandButton CmdGrafCancel1 
         Caption         =   "&Cancelar"
         Height          =   345
         Left            =   1920
         TabIndex        =   41
         Top             =   1920
         Width           =   1425
      End
      Begin VB.Frame Frame1 
         Caption         =   "Mostrar"
         Height          =   765
         Left            =   210
         TabIndex        =   39
         Top             =   1500
         Width           =   1635
         Begin VB.CheckBox ChkLeyenda 
            Caption         =   "Leyenda"
            Height          =   195
            Left            =   240
            TabIndex        =   40
            Top             =   360
            Width           =   1005
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "Con Datos"
         Height          =   1110
         Left            =   180
         TabIndex        =   36
         Top             =   360
         Width           =   1635
         Begin VB.OptionButton OptconDatosDetalle1 
            Caption         =   "Detallado"
            Height          =   210
            Left            =   165
            TabIndex        =   38
            Top             =   645
            Width           =   1155
         End
         Begin VB.OptionButton OptConDatoResum1 
            Caption         =   "Resumido"
            Height          =   195
            Left            =   165
            TabIndex        =   37
            Top             =   315
            Value           =   -1  'True
            Width           =   1155
         End
      End
      Begin VB.CommandButton CmdGrafAcep1 
         Caption         =   "&Aceptar"
         Height          =   345
         Left            =   1920
         TabIndex        =   35
         Top             =   1530
         Width           =   1425
      End
      Begin VB.Frame Frame3 
         Caption         =   "Tipo de Gráfico"
         Height          =   1110
         Left            =   1920
         TabIndex        =   31
         Top             =   360
         Width           =   1425
         Begin VB.OptionButton OptTipGrafCircular 
            Caption         =   "Circular"
            Height          =   195
            Left            =   165
            TabIndex        =   34
            Top             =   795
            Width           =   1080
         End
         Begin VB.OptionButton OptTipGrafLinea 
            Caption         =   "Lineas"
            Height          =   195
            Left            =   165
            TabIndex        =   33
            Top             =   547
            Width           =   1080
         End
         Begin VB.OptionButton OptTipGrafBarra1 
            Caption         =   "Barras"
            Height          =   195
            Left            =   165
            TabIndex        =   32
            Top             =   300
            Value           =   -1  'True
            Width           =   1080
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
         TabIndex        =   42
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
      TabIndex        =   3
      Top             =   3615
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Index           =   0
         Left            =   225
         TabIndex        =   4
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
         TabIndex        =   19
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
         TabIndex        =   21
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
         TabIndex        =   20
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
         Height          =   195
         Index           =   2
         Left            =   4275
         TabIndex        =   7
         Top             =   180
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
         TabIndex        =   6
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Producción"
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
         TabIndex        =   5
         Top             =   180
         Width           =   930
      End
      Begin VB.Shape Shape1 
         Height          =   1065
         Left            =   90
         Top             =   60
         Width           =   5805
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   43
      Top             =   0
      Width           =   12435
      _ExtentX        =   21934
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
               Picture         =   "FrmConsProd_Gerencial.frx":0000
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProd_Gerencial.frx":0544
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProd_Gerencial.frx":08D6
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProd_Gerencial.frx":0A5A
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProd_Gerencial.frx":0EAE
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProd_Gerencial.frx":0FC6
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProd_Gerencial.frx":150A
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProd_Gerencial.frx":1A4E
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProd_Gerencial.frx":1B62
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProd_Gerencial.frx":1C76
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProd_Gerencial.frx":20CA
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProd_Gerencial.frx":2236
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProd_Gerencial.frx":277E
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProd_Gerencial.frx":2A98
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsProd_Gerencial.frx":2E2A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame fr 
      Height          =   2595
      Index           =   5
      Left            =   0
      TabIndex        =   1
      Top             =   315
      Width           =   11805
      Begin VB.Frame Frame2 
         Caption         =   "Seleccionar"
         Height          =   1185
         Left            =   1590
         TabIndex        =   44
         Top             =   105
         Width           =   1125
         Begin VB.OptionButton opt_tipo 
            Caption         =   "&Resumen"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   46
            Top             =   240
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton opt_tipo 
            Caption         =   "&Detalle"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   45
            Top             =   510
            Width           =   825
         End
      End
      Begin VB.ListBox ls 
         Height          =   960
         Index           =   1
         Left            =   5250
         Style           =   1  'Checkbox
         TabIndex        =   25
         Top             =   330
         Width           =   1530
      End
      Begin VB.Frame fr 
         Caption         =   "Seleccionar"
         Height          =   1185
         Index           =   2
         Left            =   4050
         TabIndex        =   22
         Top             =   105
         Width           =   1095
         Begin VB.OptionButton opt_estilo 
            Caption         =   "Trimestre"
            Height          =   210
            Index           =   1
            Left            =   60
            TabIndex        =   27
            Top             =   390
            Width           =   960
         End
         Begin VB.OptionButton opt_estilo 
            Caption         =   "Mes"
            Height          =   210
            Index           =   0
            Left            =   60
            TabIndex        =   24
            Top             =   195
            Value           =   -1  'True
            Width           =   960
         End
         Begin VB.OptionButton opt_estilo 
            Caption         =   "Semestre"
            Height          =   210
            Index           =   2
            Left            =   60
            TabIndex        =   23
            Top             =   585
            Visible         =   0   'False
            Width           =   960
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Presentación"
         Height          =   1185
         Index           =   3
         Left            =   8340
         TabIndex        =   12
         Top             =   105
         Width           =   1485
         Begin VB.OptionButton opt_escala 
            Caption         =   "En Decimales"
            Height          =   225
            Index           =   0
            Left            =   120
            TabIndex        =   14
            Top             =   195
            Value           =   -1  'True
            Width           =   1275
         End
         Begin VB.OptionButton opt_escala 
            Caption         =   "En Miles"
            Height          =   225
            Index           =   1
            Left            =   120
            TabIndex        =   13
            Top             =   487
            Width           =   1275
         End
      End
      Begin VB.Frame fr 
         Caption         =   "Seleccionar"
         Height          =   1185
         Index           =   0
         Left            =   6825
         TabIndex        =   8
         Top             =   105
         Width           =   1485
         Begin VB.OptionButton opt_sel 
            Caption         =   "Desviación"
            Enabled         =   0   'False
            Height          =   210
            Index           =   2
            Left            =   45
            TabIndex        =   11
            Top             =   780
            Width           =   1260
         End
         Begin VB.OptionButton opt_sel 
            Caption         =   "Cant. Producc"
            Enabled         =   0   'False
            Height          =   210
            Index           =   1
            Left            =   45
            TabIndex        =   10
            Top             =   487
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton opt_sel 
            Caption         =   "Cant. Teórica"
            Enabled         =   0   'False
            Height          =   210
            Index           =   0
            Left            =   45
            TabIndex        =   9
            Top             =   195
            Width           =   1320
         End
      End
      Begin VB.ListBox ls 
         Height          =   960
         Index           =   0
         Left            =   2790
         Style           =   1  'Checkbox
         TabIndex        =   0
         Top             =   330
         Width           =   1200
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   1185
         Index           =   0
         Left            =   75
         TabIndex        =   28
         ToolTipText     =   "Buscar Producto"
         Top             =   1365
         Width           =   5790
         _cx             =   10213
         _cy             =   2090
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
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsProd_Gerencial.frx":327C
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
         Height          =   1185
         Index           =   1
         Left            =   5940
         TabIndex        =   29
         ToolTipText     =   "Buscar Insumo"
         Top             =   1365
         Width           =   5790
         _cx             =   10213
         _cy             =   2090
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsProd_Gerencial.frx":32F7
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
         PicturesOver    =   -1  'True
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
         Caption         =   "Tipo de Consulta"
         Height          =   1185
         Index           =   1
         Left            =   30
         TabIndex        =   15
         Top             =   105
         Width           =   1500
         Begin VB.OptionButton opt_consulta 
            Caption         =   "x &Insumo /  M. P. / P. I."
            Height          =   345
            Index           =   2
            Left            =   60
            TabIndex        =   18
            Top             =   750
            Width           =   1290
         End
         Begin VB.OptionButton opt_consulta 
            Caption         =   "x &Producto"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   17
            Top             =   255
            Value           =   -1  'True
            Width           =   1380
         End
         Begin VB.OptionButton opt_consulta 
            Caption         =   "x &Familia"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   16
            Top             =   510
            Width           =   1380
         End
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Selecc. Mes"
         Height          =   195
         Index           =   6
         Left            =   5295
         TabIndex        =   26
         Top             =   135
         Width           =   885
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "Seleccionar Año"
         Height          =   195
         Index           =   3
         Left            =   2805
         TabIndex        =   2
         Top             =   135
         Width           =   1170
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Height          =   4650
      Left            =   0
      TabIndex        =   47
      Top             =   2940
      Width           =   11820
      _cx             =   20849
      _cy             =   8202
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
      BackColor       =   14745342
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   128
      ForeColorSel    =   16777215
      BackColorBkg    =   -2147483636
      BackColorAlternate=   14745342
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
      FormatString    =   $"FrmConsProd_Gerencial.frx":33D1
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
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "Seleccionar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "Eliminar"
      End
   End
   Begin VB.Menu Menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu Menu2_1 
         Caption         =   "Agregar"
      End
      Begin VB.Menu Menu2_4 
         Caption         =   "Seleccionar"
      End
      Begin VB.Menu Menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu Menu2_3 
         Caption         =   "Eliminar"
      End
   End
End
Attribute VB_Name = "FrmConsProd_Gerencial"
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
                                    
Dim Q_COL_GRUPO_TERMINA As Integer '--INDICA EL TERMINO DEL GRUPO, UNE LAS CELDAS DE 1 HASTA Q_COL_GRUPO_TERMINA

Dim Q_COL_ARR_TOTAL As Integer  '--NOS INDICA EL TOTAL DE COLUMNAS QUE VA A CONTENER LOS ACUMULADOS
                                '--OBTENDRA VALOR EN VALIDAR_CONSULTA()
                                '--SI SEL MES: ENE, FEB, MAR => Q_COL_ARR_TOTAL= 2
                                '--SI SEL TRI: ENE-MAR, ABR-JUN => Q_COL_ARR_TOTAL= 1 OBS: SE INICIA DESDE POS=0
Dim F_SELECCION As Boolean '--INDICA SI SE VA SELECCIONAR LOS REGISTROS DE: PROVEEDOR, PRODUCTO
                           '--FALSE = SELECCIONA UN REGISTRO; TRUE = SELECCIONAR VARIOS REGISTROS



Private Sub CONSULTAR()
    On Error GoTo error
    Dim rst_select As New ADODB.Recordset
    Dim rst_select_1 As New ADODB.Recordset '--RECORDSET TEMPORAL
    '--
    Dim CN_TMP As New ADODB.Connection '--CONEX TEMPORAL
    Dim Rst_RUTA As New ADODB.Recordset '--CARGA RUTAS DE BD'S
    
    Dim vStrSelect As String    '--RECIBIR LA CONSULTA
    Dim vStrSelect_1 As String  '--CONSULTA OPCIONAL
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
    '--CARGAR DATOS EMPRESA
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
    For k = 0 To Rst_RUTA.RecordCount - 1
        lbl(4).Caption = "Año: " + CStr(Rst_RUTA.Fields(1))
        PgBar(0).Value = k + 1
        '------------------------------------------------
        If BAND_INTERRUMPIR = True Then GoTo salir
        '------------------------------------------------
        If k = 0 Then
            '--ENTRAR SOLO UNA VEZ
            If opt_consulta(0).Value = True And opt_tipo(1).Value = True Then
                vStrSelect_1 = GENERAR_CONSULTA(CStr(Rst_RUTA.Fields(1)), False, True)
            End If
            vStrSelect = GENERAR_CONSULTA(CStr(Rst_RUTA.Fields(1)))
        Else
            '--EN LOS DEMAS AÑO REEMPLAZAR EL AÑO ANTERIOR POR EL AÑO ACTUAL
            vStrSelect_1 = Replace(vStrSelect_1, ARR_ANYO(k - 1), CStr(Rst_RUTA.Fields(1)))
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
        If vStrSelect_1 <> "" Then RST_Busq rst_select_1, vStrSelect_1, CN_TMP
        
        '--------------------------------------
        If rst_select.RecordCount > 0 Then
            If F_CARGAR_1RA_VEZ = False Then
                '--CARGA LOS DATOS DEL PRIMER AÑO
                CARGAR_DATOS_GRILLA rst_select, CStr(Rst_RUTA.Fields(1)), rst_select_1
                F_CARGAR_1RA_VEZ = True
            Else
                '--CUANDO LOS DATOS ESTAN CARGADOS => AGREGAR DATOS A LOS DEMAS AÑOS
                CARGAR_DATOS_GRILLA_OTROS_ANYOS rst_select, CStr(Rst_RUTA.Fields(1)), rst_select_1
            End If
        End If
        CN_TMP.Close
        '--------------------------------------
        Rst_RUTA.MoveNext
    Next k

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
                                         M_ANYO As String, _
                                         RST_TMP As ADODB.Recordset)
                                         
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
        Comparar_Grupo Fg1, RST_ORIGEN, BAND_ADD_REG, M_ANYO, , RST_TMP
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
            CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Total:", , , M_ANYO
            CARGAR_DATOS_GRILLA_ADD_TOTALES True, "Tot Gen:", True, True, M_ANYO
        Else
        
            PgBar(1).Value = CLng(RST_ORIGEN.Bookmark)
            
        End If
    Wend
    PgBar(1).Value = 0
    
    Limpiar_ARRAY_TOTAL True
    
End Function


Private Sub Comparar_Grupo(GRID As Object, _
                            RST_ORIGEN As ADODB.Recordset, _
                            BAND_ADD_REG As Boolean, _
                            M_ANYO As String, _
                            Optional Q_COL_COMPARAR As Integer = -1, _
                            Optional RST_TMP As ADODB.Recordset)
                            
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
        '--DATOS ADICIONALES AL GRUPO
        CARGAR_DATOS_GRILLA_GRUPO_ADDICIONALES RST_ORIGEN, M_ANYO, Q_COL_COMPARAR, RST_TMP, GRID.Rows - 1, RST_ORIGEN.Fields(Q_COL_COMPARAR) & ""
        GRID.TextMatrix(GRID.Rows - 1, 1) = e_ESTADO_ROW_GRID.Fila_grupo
        '-----
        If opt_consulta(0).Value = True And opt_tipo(1).Value = True Then
            UNIR_CELDAS GRID, GRID.Rows - 1, Q_COL_COMPARAR + 1, GRID.Rows - 1, Q_COL_GRUPO_TERMINA, RST_ORIGEN.Fields(Q_COL_COMPARAR) & "", flexAlignLeftCenter, , flexMergeFixedOnly:
        Else
            UNIR_CELDAS GRID, GRID.Rows - 1, Q_COL_COMPARAR + 1, GRID.Rows - 1, Q_COL_GRUPO_TERMINA, RST_ORIGEN.Fields(Q_COL_COMPARAR) & "", flexAlignLeftCenter
        End If
        FORMATO_CELDA GRID, GRID.Rows - 1, Q_COL_COMPARAR_GRUPO + 1
        
        
    Else
    
        Set RST_TEPM_1 = RST_ORIGEN.Clone
        RST_TEPM_1.Bookmark = RST_ORIGEN.Bookmark
        RST_TEPM_1.MovePrevious
        
        If RST_TEPM_1.Fields(Q_COL_COMPARAR) <> RST_ORIGEN.Fields(Q_COL_COMPARAR) Then
            CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Total:", , , M_ANYO
            ADD_REG GRID, Fila_en_Blanco
            UNIR_CELDAS GRID, GRID.Rows - 1, IIf(Q_COL_FILA_OCULTA = -1, 1, Q_COL_FILA_OCULTA + 1), GRID.Rows - 1, GRID.Cols - 1, " ", flexAlignLeftCenter
            
            Limpiar_ARRAY_TOTAL

            ADD_REG GRID, Fila_grupo
            
            '--DATOS ADICIONALES AL GRUPO
            CARGAR_DATOS_GRILLA_GRUPO_ADDICIONALES RST_ORIGEN, M_ANYO, Q_COL_COMPARAR, RST_TMP, GRID.Rows - 1, RST_ORIGEN.Fields(Q_COL_COMPARAR) & ""
            GRID.TextMatrix(GRID.Rows - 1, 1) = e_ESTADO_ROW_GRID.Fila_grupo
            
            If opt_consulta(0).Value = True And opt_tipo(1).Value = True Then
                UNIR_CELDAS GRID, GRID.Rows - 1, Q_COL_COMPARAR + 1, GRID.Rows - 1, Q_COL_GRUPO_TERMINA, RST_ORIGEN.Fields(Q_COL_COMPARAR) & "", flexAlignLeftCenter, , flexMergeFixedOnly:
            Else
                UNIR_CELDAS GRID, GRID.Rows - 1, Q_COL_COMPARAR + 1, GRID.Rows - 1, Q_COL_GRUPO_TERMINA, RST_ORIGEN.Fields(Q_COL_COMPARAR) & "", flexAlignLeftCenter
            End If
            FORMATO_CELDA GRID, GRID.Rows - 1, Q_COL_COMPARAR_GRUPO + 1
            
        End If
    End If
salir:
    Set RST_TEPM_1 = Nothing
End Sub



Private Function CARGAR_DATOS_GRILLA_OTROS_ANYOS(RST_ORIGEN As ADODB.Recordset, _
                                         M_ANYO As String, _
                                         RST_TMP As ADODB.Recordset)
                                         
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    Dim Q_ROW1 As Integer
    
    If RST_ORIGEN.RecordCount > 0 Then
        RST_ORIGEN.MoveFirst
    Else
        Exit Function
    End If
    
    PgBar(1).Min = 0
    PgBar(1).Max = Fg1.Rows
    
    Dim Q_ROW As Long '--INDICA LA POSICION DEL REGISTRO A AGREGAR DATOS
    Dim N_FILTRO As String '--INDICA EL FILTRO QUE SE TENDRA QUE HACER AL RECORDSET
                            '-- DEPENDE DE Q_COL_FILA_OCULTA

    For Q_ROW = 2 To Fg1.Rows - 1
        Fg1.Row = Q_ROW
        PgBar(1).Value = Q_ROW
        '------------------------------------------------
        If BAND_INTERRUMPIR = True Then GoTo salir
        '------------------------------------------------
        '--CONCATENO MI FILTRO
        If Fg1.TextMatrix(Fg1.Row, 1) = e_ESTADO_ROW_GRID.Fila_grupo Then
            '--DATOS ADICIONALES AL GRUPO
            CARGAR_DATOS_GRILLA_GRUPO_ADDICIONALES RST_ORIGEN, M_ANYO, Q_COL_COMPARAR_GRUPO, RST_TMP, Q_ROW, Fg1.TextMatrix(Q_ROW, Q_COL_COMPARAR_GRUPO + 1)
            Fg1.TextMatrix(Fg1.Row, 1) = e_ESTADO_ROW_GRID.Fila_grupo
            
        ElseIf Fg1.TextMatrix(Fg1.Row, 1) = e_ESTADO_ROW_GRID.Fila_Total Then
            CARGAR_DATOS_GRILLA_ADD_TOTALES False, "Total:", , , M_ANYO, True
            Limpiar_ARRAY_TOTAL
            
        ElseIf Fg1.TextMatrix(Fg1.Row, 1) = e_ESTADO_ROW_GRID.Fila_Total_grl Then
            CARGAR_DATOS_GRILLA_ADD_TOTALES True, "Tot Gen:", True, True, M_ANYO, True
            
        ElseIf Fg1.TextMatrix(Fg1.Row, 1) = e_ESTADO_ROW_GRID.Fila_en_Blanco Then
        
        Else
            N_FILTRO = ""
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
salir:
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
                                         Q_ROW As Long, _
                                         Optional F_OTROS_ANYOS As Boolean = False, _
                                         Optional F_DATOS_ADICIONALES As Boolean = False)
                                         
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
                If F_DATOS_ADICIONALES = False Then
                    If opt_sel(2).Value = True Then
                        If NulosN(RST_ORIGEN.Fields(vStrCampo)) > 0 Then '--azul (consumo ahorrado)
                            FORMATO_CELDA Fg1, Q_ROW, Q_POS_MES, &HFF0000, False, &HFFFFFF, CONVERTIR_A_ESCALA(NulosN(RST_ORIGEN.Fields(vStrCampo)))
                        ElseIf NulosN(RST_ORIGEN.Fields(vStrCampo)) < 0 Then '--rojo (consumo adicional)
                            FORMATO_CELDA Fg1, Q_ROW, Q_POS_MES, &HFF, False, &HFFFFFF, CONVERTIR_A_ESCALA(NulosN(RST_ORIGEN.Fields(vStrCampo)))
                        Else
                            Fg1.TextMatrix(Q_ROW, Q_POS_MES) = CONVERTIR_A_ESCALA(NulosN(RST_ORIGEN.Fields(vStrCampo)))
                        End If
                    Else
                        Fg1.TextMatrix(Q_ROW, Q_POS_MES) = CONVERTIR_A_ESCALA(NulosN(RST_ORIGEN.Fields(vStrCampo)))
                    End If
                Else
                    FORMATO_CELDA Fg1, Q_ROW, Q_POS_MES, &H800000, False, , CONVERTIR_A_ESCALA(NulosN(RST_ORIGEN.Fields(vStrCampo)))
                End If
                Q_POS_MES = Q_POS_MES + Q_TOTAL_ANYO
                
             '--DEL TOTAL DEL AÑO
            Case "total"
                Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL + 1) * Q_TOTAL_ANYO + Q_INCREMENTO_X_COL
                '--TOTAL AÑO
                
                
                If F_DATOS_ADICIONALES = False Then
                    Fg1.TextMatrix(Q_ROW, Q_POS_MES_TOTAL) = CONVERTIR_A_ESCALA(NulosN(RST_ORIGEN.Fields(vStrCampo)))
                Else
                    FORMATO_CELDA Fg1, Q_ROW, Q_POS_MES_TOTAL, &H800000, False, , CONVERTIR_A_ESCALA(NulosN(RST_ORIGEN.Fields(vStrCampo)))
                End If
                
                '--TOTALIZAR POR FILA
                '--TOTAL GRL
                If IsNumeric(Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1)) = False Then
                    Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1) = CONVERTIR_A_ESCALA(NulosN(RST_ORIGEN.Fields(vStrCampo)))
                Else
                    Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1) = CDbl(Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1)) + CONVERTIR_A_ESCALA(NulosN(RST_ORIGEN.Fields(vStrCampo)))
                End If
                
                If F_DATOS_ADICIONALES = False Then
                    Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1) = CONVERTIR_A_ESCALA(CDbl(Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1)), True)
                Else
                    FORMATO_CELDA Fg1, Q_ROW, Fg1.Cols - 1, &H800000, False, , CONVERTIR_A_ESCALA(CDbl(Fg1.TextMatrix(Q_ROW, Fg1.Cols - 1)), True)
                End If
            '--DE LOS DEMAS CAMPOS
            Case Else
                '--SOLO SE AGREGARAN EN EL PRIMER AÑO
                
                If F_OTROS_ANYOS = False Then
                    If F_DATOS_ADICIONALES = False Then
                        Fg1.TextMatrix(Q_ROW, Q_CAMPO + 1) = RST_ORIGEN.Fields(vStrCampo) & ""
                    Else
                        FORMATO_CELDA Fg1, Q_ROW, Q_CAMPO + 1, &H800000, False, , RST_ORIGEN.Fields(vStrCampo) & ""
                    End If
                End If
        End Select
        '------------
    Next
End Function

Private Sub pImprimir()

    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.Formularios
    Me.MousePointer = vbHourglass
    
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, T_RPT_TITULO + " ", "", T_RPT_PERIODO + " ", False, True
    
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "CmdImprimir_Click"
End Sub

Private Sub Fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    On Error GoTo error
    Dim xCampos() As String
    Dim nTitulo As String
    Dim nOrden As String
    Dim nCampoBusca As String
    Dim nSQL As String
    Dim SQL_NOT_IN As String
    Dim Q_ROW As Long
    If Col <> 2 Then Exit Sub
    '--DE LOS REGISTROS YA SELECCIONADOS
    
    '----

    Select Case Index
    Case 0 '--PRODUCTO
        ReDim xCampos(2, 3) As String
        xCampos(0, 0) = "Descripción":   xCampos(0, 1) = "proddesc":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
        xCampos(1, 0) = "Familia":      xCampos(1, 1) = "famdesc":    xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
        '--si se ingresa algun filtro adicional
        SQL_NOT_IN = GENERAR_SQL_ID(fg(Index), 1, " AND alm_inventario.id", "NOT IN")
        If NulosC(fg(Index).TextMatrix(Row, Col)) <> "" Then
            SQL_NOT_IN = SQL_NOT_IN & " AND (UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' OR UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' ) "
        End If
        '-----------------------
        
        nSQL = "SELECT alm_inventario.id, alm_inventario.descripcion AS proddesc, mae_familia.descripcion AS famdesc,0 as xsel " _
            + vbCr + " FROM alm_inventario INNER JOIN mae_familia ON alm_inventario.idfam = mae_familia.id " _
            + vbCr + " WHERE (((alm_inventario.tippro) = 3)) " + SQL_NOT_IN _
            + vbCr + " ORDER BY mae_familia.descripcion, alm_inventario.descripcion; "
        
        nTitulo = "Buscando Producto"
        nOrden = "proddesc"
        nCampoBusca = "proddesc"
    Case 1 '--INSUMO
        ReDim xCampos(3, 3) As String
        xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "insdesc":     xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
        xCampos(1, 0) = "Tipo Producto":    xCampos(1, 1) = "tipprodesc":  xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
        xCampos(2, 0) = "Familia":          xCampos(2, 1) = "famdesc":     xCampos(2, 2) = "1500":   xCampos(2, 3) = "C"
        '---------------
        SQL_NOT_IN = GENERAR_SQL_ID(fg(Index), 1, " WHERE alm_inventario.id", "NOT IN")
        If NulosC(fg(Index).TextMatrix(Row, Col)) <> "" Then
            SQL_NOT_IN = SQL_NOT_IN & " AND (UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' OR UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(fg(Index).TextMatrix(Row, Col))) & "%' ) "
        End If
        '---------------
        nSQL = "SELECT alm_inventario.id, alm_inventario.descripcion AS insdesc, mae_tipoproducto.descripcion AS tipprodesc, mae_familia.descripcion AS famdesc,0 as xsel " _
             + vbCr + " FROM (pro_receta INNER JOIN ((mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) INNER JOIN pro_recetains ON alm_inventario.id = pro_recetains.iditem) ON pro_receta.id = pro_recetains.idrec) LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id " _
                        + SQL_NOT_IN _
             + vbCr + " GROUP BY alm_inventario.id, alm_inventario.descripcion, mae_tipoproducto.descripcion, mae_familia.descripcion " _
             + vbCr + " ORDER BY alm_inventario.descripcion, mae_tipoproducto.descripcion, mae_familia.descripcion;"

         
         nTitulo = "Buscando Insumos"
         nOrden = "insdesc"
         nCampoBusca = "insdesc"
    End Select

    Dim xRs As New ADODB.Recordset
    
    If F_SELECCION = False Then
        '--PERMITIRA MOSTRAR LA VENTANA PARA AGREGAR UN REGISTRO
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), nTitulo, nOrden, nCampoBusca, Principio

    Else
        '--PERMITIRA MOSTRAR LA VENTANA PARA SELECCIONAR UNO O VARIOS REGISTROS
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), nTitulo
        
    End If
    
    If xRs.State = 0 Then GoTo salir
    If xRs.EOF = True Or xRs.BOF = True Or xRs.RecordCount = 0 Then GoTo salir
    
    Do While Not xRs.EOF
        '--IMPRIMIR LOS DATOS EN PANTALLA
        For Q_ROW = 0 To xRs.Fields.Count - 2
            fg(Index).TextMatrix(fg(Index).Row, Q_ROW + 1) = xRs.Fields(Q_ROW) & ""
            
        Next Q_ROW
            
        fg(Index).AddItem ""
        fg(Index).Row = fg(Index).Rows - 1:
        fg(Index).Col = 1
        
        '--VERIFICAR SI SOLAMENTE SE AGREGA UN REGISTRO
        If F_SELECCION = False Then
            Exit Do
        Else
            Row = fg(Index).Rows - 1
        End If
        xRs.MoveNext
    Loop
    
salir:

    Set xRs = Nothing
    F_SELECCION = False
Exit Sub
error:
    Set xRs = Nothing
    SHOW_ERROR Me.Name, "Fg_CellButtonClick(" + CStr(Index) + ")"

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


Private Sub Fg_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii <> 13 Then KeyAscii = 0
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
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    
    CentrarFrm Me
    '--FORMATO DE LAS GRILLAS
    '--FORMATO DE LAS GRILLAS
    GRID_COMBOLIST fg(0), 2:        fg(0).Tag = fg(0).FormatString
    GRID_COMBOLIST fg(1), 2:        fg(1).Tag = fg(1).FormatString

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

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub

    If Me.Height > 3500 Then
        Fg1.Top = 2940
        Fg1.Width = Me.Width - 150
        Fg1.Height = Me.Height - 3350
    End If
End Sub

Private Sub Menu1_4_Click()
    F_SELECCION = True
    Fg_CellButtonClick 0, fg(0).Row, fg(0).Col
End Sub

Private Sub Menu2_4_Click()
    F_SELECCION = True
    Fg_CellButtonClick 1, fg(1).Row, fg(1).Col
End Sub

Private Sub opt_consulta_Click(Index As Integer)
    If opt_consulta(2).Value = True Then
        habilitar opt_tipo, False
        opt_tipo(1).Value = True
    Else
        habilitar opt_tipo, True
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

Private Function GENERAR_CONSULTA(M_ANYO As String, _
                        Optional F_TODO_SOL_DOL As Boolean = False, _
                        Optional MAS_DATOS As Boolean = False) As String
                        
    '--FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
    '--MAS_DATOS::TRUE=>> GENERAR LA CONSULTA DE TOTALES DE PRODUCCION MES A MES POR PRODUCTO
    
    '--
    Dim vStrSelect As String            '--CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    Dim SQL_PROD As String
    Dim SQL_INSUMO As String

    
    Dim vStrFiltro As String
    Dim vStrFiltro_1 As String      '--ESTE FILTRO SERVIRA PARA CONSULTAR EN EL SUB_SELECT
    Dim T_CONSULTA As Integer '--DEL TIPO DE CONSULTA, SE FORMARA EL ENCABEZADO DEL GRID
    
    Dim k As Integer
    '--DEL AÑO
    vStrFiltro = " Year(pro_produccion.dia)= " + M_ANYO + " "
    '--
    '--DE LOS PRODUCTOS
    SQL_PROD = GENERAR_SQL_ID(fg(0), 1, " AND pro_receta.iditem", "IN")
    
    '--DE LOS INSUMOS
    SQL_INSUMO = GENERAR_SQL_ID(fg(1), 1, " AND pro_producciondetins.iditem", "IN")
    
    vStrFiltro = vStrFiltro + SQL_PROD + SQL_INSUMO
    '------------------------------------------------------------------------------------
    
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

''            Q_COL_FILA_OCULTA       '--OCULTAR COLUMNAS
''            Q_COL_FILA              '--CANTIDAD DE COLUMNAS QUE SE MOSTRARAN DESCONTANDO LOS MESES Y LOS TOTALES
''            Q_POSICION_TOTAL        '--POSICION DE LA COLUMNA QUE SE PONDRA EL TOTAL Y TOTAL_GRL EJ. TOTAL.(COL=2)   S/. 15000
''            Q_COL_COMPARAR_GRUPO    '--NO HAY GRUPO
    If MAS_DATOS = True Then GoTo IR_MAS_DATOS
    
    '--BUSCAR TIPO DE CONSULTA
    T_CONSULTA = ESTILO_CONSULTA()
    
    Select Case T_CONSULTA
        Case 0 '--X PRODUCTO / RESUMEN
IR_MAS_DATOS:
            Q_COL_FILA_OCULTA = 1:         Q_COL_FILA = 3:        Q_POSICION_TOTAL = 3:        Q_COL_COMPARAR_GRUPO = -1
            Q_COL_GRUPO_TERMINA = 3
            T_RPT_TITULO = "REPORTE DE PRODUCCIÓN DE PRODUCTOS"
            nSQLCampos = " alm_inventario.id AS proid, " + IIf(MAS_DATOS = True, "'' as tmp0,", "") + " alm_inventario.descripcion AS prodesc, " + IIf(MAS_DATOS = True, "'' as tmp1,'' as tmp2,", "") + " mae_unidades.abrev AS prounidabrev "
            nSQLGroupBy = " alm_inventario.id, alm_inventario.descripcion, mae_unidades.abrev "
            nSQLOrderBy = " alm_inventario.descripcion "
            
        Case 1 '--X PRODUCTO / DETALLE
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 6:        Q_POSICION_TOTAL = 6:        Q_COL_COMPARAR_GRUPO = 2
            Q_COL_GRUPO_TERMINA = 5
            T_RPT_TITULO = "REPORTE DE PRODUCCIÓN DE PRODUCTOS CON INSUMOS"
            nSQLCampos = " alm_inventario.id AS proid, pro_producciondetins.iditem AS insid, alm_inventario.descripcion AS prodesc,  mae_tipoproducto.descripcion AS instipprodesc, alm_inventario_1.descripcion AS insdesc, mae_unidades_1.abrev AS insunidabrev "
            nSQLGroupBy = " alm_inventario.id, pro_producciondetins.iditem, alm_inventario.descripcion,  mae_tipoproducto.descripcion, alm_inventario_1.descripcion, mae_unidades_1.abrev "
            nSQLOrderBy = " alm_inventario.descripcion, mae_tipoproducto.descripcion, alm_inventario_1.descripcion "
            
        
        Case 2 '--X FAMILIA / RESUMEN
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 5:        Q_POSICION_TOTAL = 5:        Q_COL_COMPARAR_GRUPO = 2
            Q_COL_GRUPO_TERMINA = 5
            T_RPT_TITULO = "REPORTE DE PRODUCCIÓN DE FAMILIA CON PRODUCTO"
            nSQLCampos = " mae_familia.id AS famid,  alm_inventario.id AS proid, mae_familia.descripcion AS famdesc, alm_inventario.descripcion AS prodesc, mae_unidades.abrev AS prounidabrev "
            nSQLGroupBy = " mae_familia.id, alm_inventario.id, mae_familia.descripcion, alm_inventario.descripcion, mae_unidades.abrev "
            nSQLOrderBy = " mae_familia.descripcion, alm_inventario.descripcion "
        
        Case 3 '--X FAMILIA / DETALLE
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 6:        Q_POSICION_TOTAL = 6:        Q_COL_COMPARAR_GRUPO = 2
            Q_COL_GRUPO_TERMINA = 6
            T_RPT_TITULO = "REPORTE DE PRODUCCIÓN DE FAMILIA CON INSUMOS"
            nSQLCampos = " mae_familia.id AS famid, pro_producciondetins.iditem AS insid, mae_familia.descripcion AS famdesc, mae_tipoproducto.descripcion AS instipprodesc, alm_inventario_1.descripcion AS insdesc, mae_unidades_1.abrev AS insunidabrev "
            nSQLGroupBy = " mae_familia.id, pro_producciondetins.iditem, mae_familia.descripcion, mae_tipoproducto.descripcion, alm_inventario_1.descripcion, mae_unidades_1.abrev "
            nSQLOrderBy = " mae_familia.descripcion, mae_tipoproducto.descripcion, alm_inventario_1.descripcion "
        
        
        Case 4 '--X INSUMO / RESUMEN
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 7:        Q_POSICION_TOTAL = 7:        Q_COL_COMPARAR_GRUPO = 3
            Q_COL_GRUPO_TERMINA = 7
            T_RPT_TITULO = "REPORTE DE INSUMOS POR PRODUCTO"
            nSQLCampos = " pro_producciondetins.iditem AS insid, alm_inventario.id AS proid, mae_tipoproducto.descripcion AS instipprodesc, alm_inventario_1.descripcion AS insdesc, mae_unidades_1.abrev AS insuniabrev, alm_inventario.descripcion AS prodesc, mae_unidades.abrev AS prounidabrev "
            nSQLGroupBy = " pro_producciondetins.iditem, alm_inventario.id, mae_tipoproducto.descripcion, alm_inventario_1.descripcion, mae_unidades_1.abrev, alm_inventario.descripcion, mae_unidades.abrev "
            nSQLOrderBy = " mae_tipoproducto.descripcion, alm_inventario_1.descripcion, alm_inventario.descripcion "
        
        
        Case 5 '--X INSUMO / DETALLE
            Q_COL_FILA_OCULTA = 2:         Q_COL_FILA = 6:        Q_POSICION_TOTAL = 6:        Q_COL_COMPARAR_GRUPO = 3
            Q_COL_GRUPO_TERMINA = 6
            T_RPT_TITULO = "REPORTE DE INSUMOS POR PRODUCTO"
            nSQLCampos = " pro_producciondetins.iditem AS insid, alm_inventario.id AS proid, mae_tipoproducto.descripcion AS instipprodesc, alm_inventario_1.descripcion AS insdesc, alm_inventario.descripcion AS prodesc, mae_unidades_1.abrev AS insuniabrev "
            nSQLGroupBy = " pro_producciondetins.iditem, alm_inventario.id, mae_tipoproducto.descripcion, alm_inventario_1.descripcion, alm_inventario.descripcion, mae_unidades_1.abrev "
            nSQLOrderBy = " mae_tipoproducto.descripcion, alm_inventario_1.descripcion, alm_inventario.descripcion "
    End Select
    
    '--DE LA CANTIDAD DE COL
    Q_COL_FILA = Q_COL_FILA + 1
    Q_POS_MES_INICIO = Q_COL_FILA '--Q_COL_FILA + CAMPO_TOTAL
    '------------------------------------------
    If opt_estilo(0).Value = True Then '--MES
        nSQLPivot = "FORMAT(pro_produccion.dia,'m') "
    ElseIf opt_estilo(1).Value = True Then '--TRIMESTRE
        nSQLPivot = "FORMAT(pro_produccion.dia,'q') "
    ElseIf opt_estilo(2).Value = True Then '--SEMESTRE
        nSQLPivot = "FORMAT(pro_produccion.dia,'s') "
    End If
    '--DEL PIVOT
    For k = 0 To UBound(ARR_TMP)
        nSQLPivotSalida = nSQLPivotSalida + ARR_TMP(k, 2) + ","
    Next k
    nSQLPivotSalida = " IN (" + Left(nSQLPivotSalida, Len(nSQLPivotSalida) - 1) + ") "
    nSQLWhere = nSQLWhere + " AND " + nSQLPivot + nSQLPivotSalida
    
    '--EJM. FORMATO DE SALIDA
    'nSQLPivotSalida = " In ('Ene','Feb','Mar','Abr','May','Jun','Jul','Ago','Sep','Oct','Nov','Dic');"
    
    '-------
    
    If (opt_tipo(0).Value = True Or MAS_DATOS = True) And SQL_INSUMO = "" Then '--CANTIDAD POR RECETA
        nSQLValor = " SUM(pro_producciondet.cantidad) "
    Else
        If opt_sel(0).Value = True Then '--CANT TEORICA
            nSQLValor = " Sum(IIf(pro_producciondetins.canpro Is Null,0,(pro_producciondetins.canpro*pro_producciondet.cantidad))) "
        ElseIf opt_sel(1).Value = True Then '--CANT PRODUCIDA
            nSQLValor = " SUM(pro_producciondetins.canutil) "
        ElseIf opt_sel(2).Value = True Then '-- DESVIACION
            nSQLValor = " Sum(IIf(pro_producciondetins.canpro Is Null,0,(pro_producciondetins.canpro*pro_producciondet.cantidad))-pro_producciondetins.canutil) "
        ElseIf opt_sel(3).Value = True Then '--% DESVICACION
            nSQLValor = " SUM(IIF(pro_producciondetins.canpro IS NULL OR pro_producciondet.cantidad IS NULL,0,Sum((IIf(pro_producciondetins.canpro Is Null,0,(pro_producciondetins.canpro*pro_producciondet.cantidad))-pro_producciondetins.canutil)/(pro_producciondetins.canpro*pro_producciondet.cantidad)*100))) "
        End If
    End If
    
    '--DEL FROM ---
    If (opt_tipo(0).Value = True Or MAS_DATOS = True) And SQL_INSUMO = "" Then
        nSQLFrom = " (mae_familia RIGHT JOIN alm_inventario ON mae_familia.id = alm_inventario.idfam) RIGHT JOIN (pro_produccion INNER JOIN ((pro_producciondet INNER JOIN pro_receta ON pro_producciondet.idrec = pro_receta.id) LEFT JOIN mae_unidades ON pro_receta.idunimed = mae_unidades.id) ON pro_produccion.id = pro_producciondet.idpro) ON alm_inventario.id = pro_receta.iditem "
    Else
        nSQLFrom = "  pro_produccion INNER JOIN (((alm_inventario LEFT JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) RIGHT JOIN (mae_unidades RIGHT JOIN (pro_receta INNER JOIN pro_producciondet ON pro_receta.id = pro_producciondet.idrec) ON mae_unidades.id = pro_producciondet.idunimed) ON alm_inventario.id = pro_receta.iditem) INNER JOIN ((mae_unidades AS mae_unidades_1 RIGHT JOIN (pro_producciondetins LEFT JOIN alm_inventario AS alm_inventario_1 ON pro_producciondetins.iditem = alm_inventario_1.id) ON mae_unidades_1.id = pro_producciondetins.idunimed) LEFT JOIN mae_tipoproducto ON alm_inventario_1.tippro = mae_tipoproducto.id) ON (pro_producciondet.idrec = pro_producciondetins.idrec) AND (pro_producciondet.numparte = pro_producciondetins.numparte) AND (pro_producciondet.idpro = pro_producciondetins.idpro)) ON pro_produccion.id = pro_producciondet.idpro "
    End If
           
    '--GENERANDO LA CONSULTA
    vStrSelect = " TRANSFORM " + nSQLValor + _
        vbCr + " SELECT " + nSQLCampos + "," + nSQLValor + " AS total " + _
        vbCr + " FROM " + nSQLFrom + _
        vbCr + " WHERE " + nSQLWhere + _
        vbCr + " GROUP BY " + nSQLGroupBy + _
        vbCr + " ORDER BY " + nSQLOrderBy + _
        vbCr + " PIVOT " + nSQLPivot + nSQLPivotSalida

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
    
    If Band_Total_gral = False Then     '--DEMAS
        For Q_MES = 0 To UBound(Arr_Totales_col())
            Arr_Totales_cols(Q_MES, 0) = Arr_Totales_cols(Q_MES, 0) + Arr_Totales_col(Q_MES, 0)
        Next Q_MES
    End If

        
    '---
    If opt_sel(2).Value = True Or (opt_tipo(1).Value = True And opt_consulta(2).Value = False) Then
        
        Exit Sub
    End If
    '---

    
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
        Q_POS_MES = Q_POS_MES + Q_TOTAL_ANYO
    Next Q_MES
       
    For Q_MES = Q_COL_ARR_TOTAL + 1 To Q_COL_ARR_TOTAL + 2
        If Q_MES = Q_COL_ARR_TOTAL + 1 Then '--TOTAL
            Q_POS_MES_TOTAL = Q_POS_MES_INICIO + (Q_COL_ARR_TOTAL + 1) * Q_TOTAL_ANYO + Q_INCREMENTO_X_COL
            Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL) = CONVERTIR_A_ESCALA(IIf(Band_Total_gral = False, Arr_Totales_col(Q_MES, 0), Arr_Totales_cols(Q_MES, 0)))
        ElseIf Q_MES = Q_COL_ARR_TOTAL + 2 Then '--TOTAL GRAL
            Q_POS_MES_TOTAL = Fg1.Cols - 1
            If IsNumeric(Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL)) = False Then
                Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL) = CONVERTIR_A_ESCALA(IIf(Band_Total_gral = False, Arr_Totales_col(Q_MES, 0), Arr_Totales_cols(Q_MES, 0)))
            Else
                Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL) = CDbl(Fg1.TextMatrix(X_ROW, Q_POS_MES_TOTAL)) + CONVERTIR_A_ESCALA(IIf(Band_Total_gral = False, Arr_Totales_col(Q_MES, 0), Arr_Totales_cols(Q_MES, 0)))
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
    Dim k, j As Integer
    Dim T_CONSULTA As Integer
    
    If F_CONSERVAR_FORMATO = True Then Fg1.Clear
    
    Fg1.FrozenCols = 0
    
    If opt_estilo(0).Value = True Then '--MES
        M_ANCHO_COL_MES = 700
    ElseIf opt_estilo(1).Value = True Then '--TRIMESTRE
        M_ANCHO_COL_MES = 700
    ElseIf opt_estilo(2).Value = True Then '--SEMESTRE
        M_ANCHO_COL_MES = 800
    End If
    
    
    If Me.opt_escala(0).Value = True Then
        If opt_sel(2).Value = False Then M_ANCHO_COL_MES = M_ANCHO_COL_MES + 200
    Else
        M_ANCHO_COL_MES = M_ANCHO_COL_MES
    End If
    
    
    With Fg1
        '-----
        '--CANTIDAD DE COLUMNAS
        Fg1.Cols = Q_COL_FILA + ((Q_COL_ARR_TOTAL + 2) * Q_TOTAL_ANYO) + 1
                                '--total_mes+total_años
        
    T_CONSULTA = ESTILO_CONSULTA()
    Select Case T_CONSULTA
        Case 0 '--X PRODUCTO / RESUMEN
            .TextMatrix(1, 2) = "Producto":         .ColWidth(2) = 3500:    .ColAlignment(2) = flexAlignLeftBottom:
            .Row = 1:   .Col = 2:   .CellAlignment = flexAlignLeftBottom
            .TextMatrix(1, 3) = "U.M.":             .ColWidth(3) = 450:     .ColAlignment(3) = flexAlignLeftBottom
            .Row = 1:   .Col = 3:   .CellAlignment = flexAlignLeftBottom
        Case 1 '--X PRODUCTO / DETALLE
            .TextMatrix(1, 3) = "Producto":         .ColWidth(3) = 0:      .ColAlignment(3) = flexAlignLeftBottom
            .TextMatrix(1, 4) = "Tipo Producto":    .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftBottom
            .TextMatrix(1, 5) = "Insumo":           .ColWidth(5) = 4000:    .ColAlignment(5) = flexAlignLeftBottom
            .TextMatrix(1, 6) = "U.M.":             .ColWidth(6) = 450:     .ColAlignment(6) = flexAlignLeftBottom
        
            If opt_consulta(0).Value = True And opt_tipo(1).Value = True Then
                .ColWidth(4) = 0
                .ColWidth(5) = 4500
            End If
        Case 2 '--X FAMILIA / RESUMEN
            .TextMatrix(1, 3) = "Familia":          .ColWidth(3) = 0:       .ColAlignment(3) = flexAlignLeftBottom
            .TextMatrix(1, 4) = "Producto":         .ColWidth(4) = 3500:    .ColAlignment(4) = flexAlignLeftBottom
            .TextMatrix(1, 5) = "U.M.":             .ColWidth(5) = 450:     .ColAlignment(5) = flexAlignLeftBottom
        
        
        Case 3 '--X FAMILIA / DETALLE
            .TextMatrix(1, 3) = "Familia":          .ColWidth(3) = 0:       .ColAlignment(3) = flexAlignLeftBottom
            .TextMatrix(1, 4) = "Tipo Producto":    .ColWidth(4) = 1200:    .ColAlignment(4) = flexAlignLeftBottom
            .TextMatrix(1, 5) = "Insumo":           .ColWidth(5) = 4000:    .ColAlignment(5) = flexAlignLeftBottom
            .TextMatrix(1, 6) = "U.M.":             .ColWidth(6) = 450:     .ColAlignment(6) = flexAlignLeftBottom
        
        
        Case 4 '--X INSUMO / RESUMEN
            .TextMatrix(1, 3) = "Tipo Producto":    .ColWidth(3) = 0:    .ColAlignment(3) = flexAlignLeftBottom
            .TextMatrix(1, 4) = "Insumo":           .ColWidth(4) = 0:    .ColAlignment(4) = flexAlignLeftBottom
            .TextMatrix(1, 5) = "U.M.":             .ColWidth(5) = 0:     .ColAlignment(5) = flexAlignLeftBottom
        
            .TextMatrix(1, 6) = "Producto":         .ColWidth(6) = 4500:    .ColAlignment(6) = flexAlignLeftBottom
            .TextMatrix(1, 7) = "U.M.":             .ColWidth(7) = 450:     .ColAlignment(7) = flexAlignLeftBottom

        Case 5 '--X INSUMO / DETALLE
            .TextMatrix(1, 3) = "Tipo Producto":    .ColWidth(3) = 0:    .ColAlignment(3) = flexAlignLeftBottom
            .TextMatrix(1, 4) = "Insumo":           .ColWidth(4) = 0:    .ColAlignment(4) = flexAlignLeftBottom
        
            .TextMatrix(1, 5) = "Producto":         .ColWidth(5) = 4500:    .ColAlignment(5) = flexAlignLeftBottom
            .TextMatrix(1, 6) = "U.M.":             .ColWidth(6) = 450:     .ColAlignment(6) = flexAlignLeftBottom
    End Select
        '---
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
                UNIR_CELDAS Fg1, 1, Q_POS_MES + j, 1, Q_POS_MES + j, ARR_ANYO(j), flexAlignCenterTop
                .ColAlignment(Q_POS_MES + j) = flexAlignRightBottom
                .Row = 1:       .Col = Q_POS_MES + j:           .CellAlignment = flexAlignCenterBottom
            Next j
            
            Q_POS_MES = Q_POS_MES + Q_TOTAL_ANYO
            
        Next k
        '--COLOCANDO LOS TOTALES
        .ColWidth(.Cols - 1) = M_ANCHO_COL_MES + 400
        .ColAlignment(.Cols - 1) = flexAlignRightBottom
        UNIR_CELDAS Fg1, 0, .Cols - 1, 0, .Cols - 1, "Total Gral.", flexAlignCenterTop
        
        'DEL TOTAL GRAL
        UNIR_CELDAS Fg1, 1, .Cols - 1, 1, .Cols - 1, "Total", flexAlignCenterTop
       


        .FrozenCols = Q_POS_MES_INICIO - 1
        .ColWidth(0) = 200
        '--DE LOS ID'S
        For k = 1 To Q_COL_FILA_OCULTA
            .TextMatrix(1, k) = "ID" + CStr(k):         .ColWidth(k) = 500
        Next
        If Q_COL_FILA_OCULTA <> -1 Then OCULTAR_COL Fg1, 1, Q_COL_FILA_OCULTA
        UNIR_CELDAS Fg1, 0, 1, 0, Q_POS_MES_INICIO - 1, " ", flexAlignCenterTop
        
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
        CONVERTIR_A_ESCALA = "0.00"
        Exit Function
    End If
    
    If SOLO_FORMAT = True Then
        CONVERTIR_A_ESCALA = Format(S_MONTO, FORMAT_MONTO)
        Exit Function
    End If
    If Me.opt_escala(0).Value = True Then
        CONVERTIR_A_ESCALA = Format(S_MONTO, FORMAT_MONTO)
    Else
        CONVERTIR_A_ESCALA = Format(S_MONTO / 1000, FORMAT_MONTO)
    End If
     
End Function

Private Function ESTILO_CONSULTA() As Integer
    Dim M_ESTILO As Integer
    If opt_consulta(0).Value = True Then '--X PRODUCTO
        If opt_tipo(0).Value = True Then M_ESTILO = 0 '--X PRODUCTO/RESUMEN
        If opt_tipo(1).Value = True Then M_ESTILO = 1 '--X PRODUCTO/DETALLE
        
    ElseIf opt_consulta(1).Value = True Then '--X FAMILIA
        If opt_tipo(0).Value = True Then M_ESTILO = 2 '--X FAMILIA/RESUMEN
        If opt_tipo(1).Value = True Then M_ESTILO = 3 '--X FAMILIA/DETALLE
        
    ElseIf opt_consulta(2).Value = True Then '--X INSUMO
        If opt_tipo(0).Value = True Then M_ESTILO = 4 '--X INSUMO/RESUMEN
        If opt_tipo(1).Value = True Then M_ESTILO = 5 '--X INSUMO/DETALLE
    End If
    ESTILO_CONSULTA = M_ESTILO
End Function

Private Sub opt_sel_Click(Index As Integer)
    If Index = 2 Then
        habilitar opt_escala, False
        opt_escala(0).Value = True
    Else
        habilitar opt_escala, True
    End If
End Sub

Private Sub opt_tipo_Click(Index As Integer)
    If Index = 0 Then
        habilitar opt_sel, False
        opt_sel(1).Value = True
    Else
        habilitar opt_sel, True
    End If
End Sub


Private Sub CARGAR_DATOS_GRILLA_GRUPO_ADDICIONALES(RST_ORIGEN As ADODB.Recordset, _
                                                    M_ANYO As String, _
                                                    Q_COL_COMPARAR As Integer, _
                                                    RST_TMP As ADODB.Recordset, _
                                                    Q_ROW1 As Long, N_TEXTO_COMPARA As String)

    '--ESTA FUNCION AGREGARA DATOS ADICIONALES A LA FILA Q_ROW1 DESDE RST_TMP
    Dim N_FILTRO  As String
    If RST_TMP.State = 0 Then Exit Sub
    RST_TMP.Filter = ""
    If RST_TMP.EOF = False Or RST_TMP.BOF = False Or RST_TMP.RecordCount <> 0 Then
        N_FILTRO = RST_ORIGEN.Fields(Q_COL_COMPARAR).Name + "= '" + N_TEXTO_COMPARA + "'"
        RST_TMP.Filter = N_FILTRO '--HACER EL FILTRO
        If RST_TMP.RecordCount > 0 Then
            CARGAR_DATOS_GRILLA_ARRAY_TMP RST_TMP, M_ANYO, Q_ROW1, False, True
        End If
    End If

End Sub

'----
Private Sub fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Index = 0 Then PopupMenu Menu1
        If Index = 1 Then PopupMenu Menu2
    End If
End Sub

'--DEL PRODUCTO
Private Sub menu1_1_Click()
    F_SELECCION = False
    Fg_CellButtonClick 0, fg(0).Rows - 1, 2
End Sub

Private Sub menu1_3_Click()
    Fg_KeyDown 0, 46, 0
End Sub
'--DE LOS INSUMOS
Private Sub Menu2_1_Click()
    F_SELECCION = False
    Fg_CellButtonClick 1, fg(1).Rows - 1, 2
End Sub

Private Sub Menu2_3_Click()
    Fg_KeyDown 1, 46, 0
End Sub
'--------


Private Sub EXPORTAR()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.Formularios
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, T_RPT_TITULO, T_RPT_PERIODO, "", "Producción"
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

