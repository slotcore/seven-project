VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FormReportePedidos2 
   Caption         =   "Ventas  -  Reporte de Pedidos"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11850
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7230
   ScaleWidth      =   11850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame6 
      Caption         =   "Reporte"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4995
      Left            =   30
      TabIndex        =   2
      Top             =   2190
      Width           =   11820
      Begin VB.Frame FraProgreso 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   945
         Left            =   2760
         TabIndex        =   32
         Top             =   1200
         Visible         =   0   'False
         Width           =   5940
         Begin MSComctlLib.ProgressBar PgBar 
            Height          =   255
            Left            =   225
            TabIndex        =   33
            Top             =   465
            Width           =   5505
            _ExtentX        =   9710
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   1
            Appearance      =   0
            Scrolling       =   1
         End
         Begin VB.Shape Shape1 
            Height          =   765
            Left            =   60
            Top             =   90
            Width           =   5805
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "Pedidos"
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
            TabIndex        =   36
            Top             =   180
            Width           =   660
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
            TabIndex        =   35
            Top             =   180
            Width           =   1035
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
            TabIndex        =   34
            Top             =   180
            Width           =   1530
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid FgDet 
         Height          =   1980
         Left            =   90
         TabIndex        =   28
         Top             =   2670
         Visible         =   0   'False
         Width           =   11580
         _cx             =   20426
         _cy             =   3492
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
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   100
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
         ScrollTrack     =   0   'False
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   0
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   0   'False
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
      Begin VSFlex7Ctl.VSFlexGrid FgReporteProd 
         Height          =   2370
         Left            =   90
         TabIndex        =   3
         Top             =   270
         Width           =   11655
         _cx             =   20558
         _cy             =   4180
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
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   0
         Cols            =   10
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   2805
      Left            =   30
      TabIndex        =   0
      Top             =   -150
      Width           =   11895
      Begin VB.Frame Frame2 
         Caption         =   "[Tipo Consulta]"
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
         Height          =   1760
         Left            =   9570
         TabIndex        =   25
         Top             =   600
         Width           =   2175
         Begin VB.OptionButton OptionFechas 
            Caption         =   "Por Fechas"
            Enabled         =   0   'False
            Height          =   255
            Left            =   600
            TabIndex        =   31
            Top             =   870
            Width           =   1335
         End
         Begin VB.OptionButton OptionClientes 
            Caption         =   "Por Clientes"
            Enabled         =   0   'False
            Height          =   255
            Left            =   600
            TabIndex        =   30
            Top             =   1110
            Width           =   1365
         End
         Begin VB.OptionButton OptionProductos 
            Caption         =   "Por Productos"
            Enabled         =   0   'False
            Height          =   255
            Left            =   600
            TabIndex        =   29
            Top             =   1380
            Width           =   1365
         End
         Begin VB.CheckBox CheckDetallado 
            Caption         =   "Detallado"
            Height          =   195
            Left            =   300
            TabIndex        =   27
            Top             =   615
            Width           =   1005
         End
         Begin VB.CheckBox CheckResumido 
            Caption         =   "Resumido"
            Height          =   195
            Left            =   300
            TabIndex        =   26
            Top             =   330
            Width           =   1065
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "[Fech. a Entregar]"
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
         Height          =   885
         Left            =   6690
         TabIndex        =   20
         Top             =   600
         Width           =   2835
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEntDesde 
            Height          =   300
            Left            =   1080
            TabIndex        =   21
            Top             =   225
            Width           =   1275
            _ExtentX        =   2249
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
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEntHasta 
            Height          =   300
            Left            =   1080
            TabIndex        =   22
            Top             =   525
            Width           =   1275
            _ExtentX        =   2249
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
         End
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   375
            TabIndex        =   24
            Top             =   255
            Width           =   510
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Left            =   375
            TabIndex        =   23
            Top             =   555
            Width           =   465
         End
      End
      Begin VB.Frame Frame15 
         Caption         =   "[Condicion]"
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
         Height          =   885
         Left            =   6690
         TabIndex        =   16
         Top             =   1470
         Width           =   2835
         Begin VB.CheckBox CheckVigent 
            Caption         =   "Vigente"
            Height          =   195
            Left            =   1080
            TabIndex        =   19
            Top             =   285
            Width           =   1065
         End
         Begin VB.CheckBox CheckAnul 
            Caption         =   "Anulado"
            Height          =   195
            Left            =   1080
            TabIndex        =   18
            Top             =   585
            Width           =   1005
         End
         Begin VB.CheckBox CheckCond 
            Height          =   195
            Left            =   600
            TabIndex        =   17
            Top             =   285
            Width           =   170
         End
      End
      Begin VB.Frame Frame4 
         Caption         =   "[Productos]"
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
         Height          =   880
         Left            =   90
         TabIndex        =   14
         Top             =   1470
         Width           =   3250
         Begin VSFlex7Ctl.VSFlexGrid FgProductos 
            Height          =   630
            Left            =   90
            TabIndex        =   15
            ToolTipText     =   "Buscar Productos"
            Top             =   195
            Width           =   3105
            _cx             =   5477
            _cy             =   1111
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
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FormReportePedidos2.frx":0000
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   2
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
      End
      Begin VB.Frame Frame13 
         Caption         =   "[Clientes]"
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
         Height          =   880
         Left            =   90
         TabIndex        =   12
         Top             =   600
         Width           =   3250
         Begin VSFlex7Ctl.VSFlexGrid FgClientes 
            Height          =   630
            Left            =   90
            TabIndex        =   13
            ToolTipText     =   "Buscar Clientes"
            Top             =   195
            Width           =   3105
            _cx             =   5477
            _cy             =   1111
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
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FormReportePedidos2.frx":003D
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   2
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
      End
      Begin VB.Frame Frame14 
         Caption         =   "[Ord. de Pedido]"
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
         Height          =   880
         Left            =   3390
         TabIndex        =   10
         Top             =   600
         Width           =   3250
         Begin VSFlex7Ctl.VSFlexGrid FgOC 
            Height          =   625
            Left            =   120
            TabIndex        =   11
            ToolTipText     =   "Buscar Ordenes de Pedido"
            Top             =   195
            Width           =   3045
            _cx             =   5371
            _cy             =   1102
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
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   1
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FormReportePedidos2.frx":007A
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   2
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
      End
      Begin VB.Frame Frame16 
         Caption         =   "[Fech. Emision]"
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
         Height          =   885
         Left            =   3390
         TabIndex        =   5
         Top             =   1470
         Width           =   3255
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmiDesde 
            Height          =   300
            Left            =   1260
            TabIndex        =   6
            Top             =   210
            Width           =   1275
            _ExtentX        =   2249
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
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmiHasta 
            Height          =   300
            Left            =   1260
            TabIndex        =   7
            Top             =   540
            Width           =   1275
            _ExtentX        =   2249
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
         End
         Begin VB.Label Label28 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Left            =   555
            TabIndex        =   9
            Top             =   585
            Width           =   465
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   555
            TabIndex        =   8
            Top             =   255
            Width           =   510
         End
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   345
         Left            =   0
         TabIndex        =   1
         Top             =   180
         Width           =   11900
         _ExtentX        =   20981
         _ExtentY        =   609
         ButtonWidth     =   609
         ButtonHeight    =   556
         Appearance      =   1
         Style           =   1
         ImageList       =   "ImageList1"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Buscar"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
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
               Object.Visible         =   0   'False
               Object.ToolTipText     =   "Imprimir"
               ImageIndex      =   12
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.ToolTipText     =   "Salir"
               ImageIndex      =   11
            EndProperty
         EndProperty
         BorderStyle     =   1
         Begin MSComctlLib.ImageList ImageList1 
            Left            =   8250
            Top             =   30
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
                  Picture         =   "FormReportePedidos2.frx":00B7
                  Key             =   "IMG1"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReportePedidos2.frx":05FB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReportePedidos2.frx":098D
                  Key             =   "IMG2"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReportePedidos2.frx":0B11
                  Key             =   "IMG3"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReportePedidos2.frx":0F65
                  Key             =   "IMG4"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReportePedidos2.frx":107D
                  Key             =   "IMG5"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReportePedidos2.frx":15C1
                  Key             =   "IMG6"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReportePedidos2.frx":1B05
                  Key             =   "IMG7"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReportePedidos2.frx":1C19
                  Key             =   "IMG8"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReportePedidos2.frx":1D2D
                  Key             =   "IMG9"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReportePedidos2.frx":2181
                  Key             =   "IMG10"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReportePedidos2.frx":22ED
                  Key             =   "IMG11"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReportePedidos2.frx":2835
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormReportePedidos2.frx":2B4F
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PRODUCTO_01"
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
      Left            =   1560
      TabIndex        =   4
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "FormReportePedidos2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cargo As Boolean
Dim interrumpir As Boolean
Dim estadoCheckCondicion As Boolean

Private Function verificarDatos() As Boolean
    verificarDatos = True
    If TxtFchEmiDesde.Valor = "" Then MsgBox "Ingrese Fecha de Emision Desde": verificarDatos = False
    If TxtFchEmiHasta.Valor = "" Then MsgBox "Ingrese Fecha de Emision Hasta": verificarDatos = False
    If TxtFchEntDesde.Valor = "" Then MsgBox "Ingrese Fecha de Entrega Desde": verificarDatos = False
    If TxtFchEntHasta.Valor = "" Then MsgBox "Ingrese Fecha de Entrega Hasta": verificarDatos = False
End Function

Private Sub generarConsultaResumido()
    Dim RstLis As New ADODB.Recordset
    Dim A As Integer
    Dim fila As Integer
    
    Dim c_PRODUCTOS As String
    Dim c_CLIENTES As String
    Dim c_OC As String
    
    Dim c_FECHA_EMI As String
    Dim c_FECHA_PLAZ1 As String
    Dim c_FECHA_PLAZ2 As String
    Dim c_FECHA_PLAZ3 As String
    
    Dim c_ESTADO As String
    Dim c_SITUACION As String
    Dim c_SITUACION2 As String
    Dim c_CONDICION As String
    Dim c_CATEGORIA As String
    Dim c_CATEGORIA2 As String
        
    Dim c_FECHA_ENT As String
    
    Dim c_SQL As String
    
    limpiarReporte
    
    'Consulta para Productos
    FgProductos.Row = 0
    FgProductos.Col = 1
    c_PRODUCTOS = "((alm_inventario.descripcion)= '" + FgProductos.Text + "'"
    If (FgProductos.TextMatrix(0, 1) = "Todos") Then
        c_PRODUCTOS = ""
    Else
        For A = 0 To FgProductos.Rows - 1
            FgProductos.Row = A
            FgProductos.Col = 1
            c_PRODUCTOS = c_PRODUCTOS + " OR " + "(alm_inventario.descripcion)= '" + FgProductos.Text + "'"
        Next A
        c_PRODUCTOS = c_PRODUCTOS + ") AND "
    End If
    'Consulta para Clientes
    FgClientes.Row = 0
    FgClientes.Col = 1
    c_CLIENTES = "((mae_cliente.nombre)= '" + FgClientes.Text + "')"
    If (FgClientes.TextMatrix(0, 1) = "Todos") Then
        c_CLIENTES = ""
    Else
        For A = 1 To FgClientes.Rows - 2
            FgClientes.Row = A
            FgClientes.Col = 1
            c_CLIENTES = c_CLIENTES + " OR ((mae_cliente.nombre)= '" + FgClientes.Text + "')"
        Next A
        c_CLIENTES = "(" + c_CLIENTES + ") AND "
    End If
    'Consulta para Ordenes de Pedido
    FgOC.Row = 0
    FgOC.Col = 1
    c_OC = "((ped_pedido.oc)= '" + FgOC.Text + "'"
    If (FgOC.TextMatrix(0, 1) = "Todos") Then
        c_OC = ""
    Else
        For A = 0 To FgOC.Rows - 1
            FgOC.Row = A
            FgOC.Col = 1
            c_OC = c_OC + " OR " + "(ped_pedido.oc)= '" + FgOC.Text + "'"
        Next A
        c_OC = c_OC + ") AND "
    End If
    
    'Consulta para Fecha de Emision
    c_FECHA_EMI = "((ped_pedido.fchemi)>=CDate('" & TxtFchEmiDesde.Valor & "')"
    c_FECHA_EMI = c_FECHA_EMI & " AND (ped_pedido.fchemi)<=CDate('" & TxtFchEmiHasta.Valor & "'))"
    'Consulta para Fecha a Entregar
    c_FECHA_PLAZ1 = "((ped_pedido.fchent)>=CDate('" & TxtFchEntDesde.Valor & "')"
    c_FECHA_PLAZ1 = c_FECHA_PLAZ1 & " AND (ped_pedido.fchent)<=CDate('" & TxtFchEntHasta.Valor & "'))"
    
    c_FECHA_PLAZ2 = "((ped_pedidodetent.fchent)>=CDate('" & TxtFchEntDesde.Valor & "')"
    c_FECHA_PLAZ2 = c_FECHA_PLAZ2 & " AND (ped_pedidodetent.fchent)<=CDate('" & TxtFchEntHasta.Valor & "'))"
    
    c_FECHA_PLAZ3 = "((IIf([ped_pedidodet].[fchent] Is Not Null,[ped_pedidodet].[fchent],[ped_pedido].[fchent]))>=CDate('" & TxtFchEntDesde.Valor & "')"
    c_FECHA_PLAZ3 = c_FECHA_PLAZ3 & " AND (IIf([ped_pedidodet].[fchent] Is Not Null,[ped_pedidodet].[fchent],[ped_pedido].[fchent]))<=CDate('" & TxtFchEntHasta.Valor & "'))"
    
    'Condicion para entregados y no entregados
    If CheckCond.Value = 1 Then
        c_CONDICION = ""
    Else
        Dim eVigente As String
        Dim eAnulado As String
        
        If CheckVigent.Value = 1 Then
            eVigente = "((IIf([ped_pedido].[anulado]=0,'VIGENTE','ANULADO'))='VIGENTE')"
        Else
            eVigente = ""
        End If
        If CheckAnul.Value = 1 Then
            If CheckVigent.Value = 0 Then
                eAnulado = "((IIf([ped_pedido].[anulado]=0,'VIGENTE','ANULADO'))='ANULADO')"
            Else
                eAnulado = " OR ((IIf([ped_pedido].[anulado]=0,'VIGENTE','ANULADO'))='ANULADO') "
            End If
        Else
            eAnulado = ""
        End If
        c_CONDICION = "(" & eVigente & eAnulado & ") AND "
    End If
    
    c_SQL = "SELECT ped_pedido.oc, mae_cliente.nombre AS nombre, ped_pedido.id, ped_pedidodet.iditem, alm_inventario.descripcion, ped_pedido.idcli, ped_pedido.idpunvecli, ped_pedido.tipdoc, ped_pedido.numser, ped_pedido.numdoc, ped_pedido.idconpag, ped_pedido.fchemi, ped_pedido.glosa, ped_pedido.idtipped, IIf([ped_pedidodet].[fchent] Is Not Null,[ped_pedidodet].[fchent],[ped_pedido].[fchent]) AS fEnt, ped_pedidodet.canpro, IIf([ped_pedido].[anulado]=0,'VIGENTE','ANULADO') AS Anulado, ped_pedido.numreg, ped_pedido.idlib, ped_pedido.proceso, IIf(IsNull(ped_pedido!numser)=-1,ped_pedido!numdoc,ped_pedido!numser+'-'+ped_pedido!numdoc) AS numerodoc, mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, mae_condpago.abrev AS conpagabre, ped_pedido.fchemi & '' AS fchemi1, vta_puntoVenta.descripcion AS ptovta, ped_tipo.descripcion AS tipped, mae_unidades.abrev " _
            + vbCr + "FROM (((ped_tipo RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN (vta_puntoVenta RIGHT JOIN (ped_pedido LEFT JOIN mae_condpago ON ped_pedido.idconpag = mae_condpago.id) ON vta_puntoVenta.id = ped_pedido.idpunvecli) ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) ON ped_tipo.id = ped_pedido.idtipped) LEFT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id " _
            + vbCr + "GROUP BY ped_pedido.oc, mae_cliente.nombre, ped_pedido.id, ped_pedidodet.iditem, alm_inventario.descripcion, ped_pedido.idcli, ped_pedido.idpunvecli, ped_pedido.tipdoc, ped_pedido.numser, ped_pedido.numdoc, ped_pedido.idconpag, ped_pedido.fchemi, ped_pedido.glosa, ped_pedido.idtipped, IIf([ped_pedidodet].[fchent] Is Not Null,[ped_pedidodet].[fchent],[ped_pedido].[fchent]), ped_pedidodet.canpro, IIf([ped_pedido].[anulado]=0,'VIGENTE','ANULADO'), ped_pedido.numreg, ped_pedido.idlib, ped_pedido.proceso, IIf(IsNull(ped_pedido!numser)=-1,ped_pedido!numdoc,ped_pedido!numser+'-'+ped_pedido!numdoc), mae_documento.descripcion, mae_condpago.descripcion, mae_documento.abrev, mae_cliente.numruc, mae_condpago.abrev, ped_pedido.fchemi & '', vta_puntoVenta.descripcion, ped_tipo.descripcion, mae_unidades.abrev " _
            + vbCr + "HAVING (" & c_CLIENTES & c_OC & c_PRODUCTOS & c_CONDICION & c_FECHA_EMI & " And " & c_FECHA_PLAZ3 & ") " _
            + vbCr + "ORDER BY IIf([ped_pedidodet].[fchent] Is Not Null,[ped_pedidodet].[fchent],[ped_pedido].[fchent]) DESC; " _
            + vbCr + "Union " _
            + vbCr + "SELECT ped_pedido.oc, mae_cliente.nombre AS nombre, ped_pedido.id, ped_pedidodetent.iditem, alm_inventario.descripcion, ped_pedido.idcli, ped_pedido.idpunvecli, ped_pedido.tipdoc, ped_pedido.numser, ped_pedido.numdoc, ped_pedido.idconpag, ped_pedido.fchemi, ped_pedido.glosa, ped_pedido.idtipped, ped_pedidodetent.fchent As fEnt, ped_pedidodetent.canpro, IIf([ped_pedido].[anulado]=0,'VIGENTE','ANULADO') AS Anulado, ped_pedido.numreg, ped_pedido.idlib, ped_pedido.proceso, IIf(IsNull(ped_pedido!numser)=-1,ped_pedido!numdoc,ped_pedido!numser+'-'+ped_pedido!numdoc) AS numerodoc, mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, mae_condpago.abrev AS conpagabre, ped_pedido.fchemi & '' AS fchemi1, vta_puntoVenta.descripcion AS ptovta, ped_tipo.descripcion AS tipped, mae_unidades.abrev " _
            + vbCr + "FROM (((ped_tipo RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN (vta_puntoVenta RIGHT JOIN (ped_pedido LEFT JOIN mae_condpago ON ped_pedido.idconpag = mae_condpago.id) ON vta_puntoVenta.id = ped_pedido.idpunvecli) ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) ON ped_tipo.id = ped_pedido.idtipped) LEFT JOIN ped_pedidodetent ON ped_pedido.id = ped_pedidodetent.idped) LEFT JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodetent.idunimed = mae_unidades.id " _
            + vbCr + "Where (((ped_pedido.idtipped) = 2)) " _
            + vbCr + "GROUP BY ped_pedido.id, ped_pedido.idcli, ped_pedido.idpunvecli, ped_pedido.tipdoc, ped_pedido.numser, ped_pedido.numdoc, ped_pedido.idconpag, ped_pedido.fchemi, ped_pedido.glosa, ped_pedido.idtipped, ped_pedidodetent.fchent, IIf([ped_pedido].[anulado]=0,'VIGENTE','ANULADO'), ped_pedido.numreg, ped_pedido.idlib, ped_pedido.oc, ped_pedido.proceso, mae_cliente.nombre, IIf(IsNull(ped_pedido!numser)=-1,ped_pedido!numdoc,ped_pedido!numser+'-'+ped_pedido!numdoc), mae_documento.descripcion, mae_condpago.descripcion, mae_documento.abrev, mae_cliente.numruc, mae_condpago.abrev, ped_pedido.fchemi & '', vta_puntoVenta.descripcion, ped_tipo.descripcion, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodetent.canpro, ped_pedidodetent.iditem " _
            + vbCr + "HAVING (" & c_CLIENTES & c_OC & c_PRODUCTOS & c_CONDICION & c_FECHA_EMI & " And " & c_FECHA_PLAZ2 & ");"

    
'    c_SQL = "SELECT ped_pedido.oc, mae_cliente.nombre AS nombre, ped_pedido.id, ped_pedidodet.iditem, alm_inventario.descripcion, ped_pedido.idcli, ped_pedido.idpunvecli, ped_pedido.tipdoc, ped_pedido.numser, ped_pedido.numdoc, ped_pedido.idconpag, ped_pedido.fchemi, ped_pedido.glosa, ped_pedido.idtipped, ped_pedido.fchent As fEnt, ped_pedidodet.canpro, IIf([ped_pedido].[anulado]=0,'VIGENTE','ANULADO') AS Anulado, ped_pedido.numreg, ped_pedido.idlib, ped_pedido.proceso, IIf(IsNull(ped_pedido!numser)=-1,ped_pedido!numdoc,ped_pedido!numser+'-'+ped_pedido!numdoc) AS numerodoc, mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, mae_condpago.abrev AS conpagabre, ped_pedido.fchemi & '' AS fchemi1, vta_puntoVenta.descripcion AS ptovta, ped_tipo.descripcion AS tipped, mae_unidades.abrev " _
'            + vbCr + "FROM (((ped_tipo RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN (vta_puntoVenta RIGHT JOIN (ped_pedido LEFT JOIN mae_condpago ON ped_pedido.idconpag = mae_condpago.id) ON vta_puntoVenta.id = ped_pedido.idpunvecli) ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) ON ped_tipo.id = ped_pedido.idtipped) LEFT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id " _
'            + vbCr + "Where (((ped_pedido.idtipped) = 1)) " _
'            + vbCr + "GROUP BY ped_pedido.id, ped_pedido.idcli, ped_pedido.idpunvecli, ped_pedido.tipdoc, ped_pedido.numser, ped_pedido.numdoc, ped_pedido.idconpag, ped_pedido.fchemi, ped_pedido.glosa, ped_pedido.idtipped, ped_pedido.fchent, IIf([ped_pedido].[anulado]=0,'VIGENTE','ANULADO'), ped_pedido.numreg, ped_pedido.idlib, ped_pedido.oc, ped_pedido.proceso, mae_cliente.nombre, IIf(IsNull(ped_pedido!numser)=-1,ped_pedido!numdoc,ped_pedido!numser+'-'+ped_pedido!numdoc), mae_documento.descripcion, mae_condpago.descripcion, mae_documento.abrev, mae_cliente.numruc, mae_condpago.abrev, ped_pedido.fchemi & '', vta_puntoVenta.descripcion, ped_tipo.descripcion, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodet.canpro, ped_pedidodet.iditem " _
'            + vbCr + "HAVING (" & c_CLIENTES & c_OC & c_PRODUCTOS & c_CONDICION & c_FECHA_EMI & " And " & c_FECHA_PLAZ1 & ") " _
'            + vbCr + "ORDER BY ped_pedido.fchent DESC; " _
'            + vbCr + "Union " _
'            + vbCr + "SELECT ped_pedido.oc, mae_cliente.nombre AS nombre, ped_pedido.id, ped_pedidodet.iditem, alm_inventario.descripcion, ped_pedido.idcli, ped_pedido.idpunvecli, ped_pedido.tipdoc, ped_pedido.numser, ped_pedido.numdoc, ped_pedido.idconpag, ped_pedido.fchemi, ped_pedido.glosa, ped_pedido.idtipped, ped_pedidodet.fchent AS fEnt, ped_pedidodet.canpro, IIf([ped_pedido].[anulado]=0,'VIGENTE','ANULADO') AS Anulado, ped_pedido.numreg, ped_pedido.idlib, ped_pedido.proceso, IIf(IsNull(ped_pedido!numser)=-1,ped_pedido!numdoc,ped_pedido!numser+'-'+ped_pedido!numdoc) AS numerodoc, mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, mae_condpago.abrev AS conpagabre, ped_pedido.fchemi & '' AS fchemi1, vta_puntoVenta.descripcion AS ptovta, ped_tipo.descripcion AS tipped, mae_unidades.abrev " _
'            + vbCr + "FROM (((ped_tipo RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN (vta_puntoVenta RIGHT JOIN (ped_pedido LEFT JOIN mae_condpago ON ped_pedido.idconpag = mae_condpago.id) ON vta_puntoVenta.id = ped_pedido.idpunvecli) ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) ON ped_tipo.id = ped_pedido.idtipped) LEFT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id " _
'            + vbCr + "GROUP BY ped_pedido.oc, mae_cliente.nombre, ped_pedido.id, ped_pedidodet.iditem, alm_inventario.descripcion, ped_pedido.idcli, ped_pedido.idpunvecli, ped_pedido.tipdoc, ped_pedido.numser, ped_pedido.numdoc, ped_pedido.idconpag, ped_pedido.fchemi, ped_pedido.glosa, ped_pedido.idtipped, ped_pedidodet.fchent, ped_pedidodet.canpro, IIf([ped_pedido].[anulado]=0,'VIGENTE','ANULADO'), ped_pedido.numreg, ped_pedido.idlib, ped_pedido.proceso, IIf(IsNull(ped_pedido!numser)=-1,ped_pedido!numdoc,ped_pedido!numser+'-'+ped_pedido!numdoc), mae_documento.descripcion, mae_condpago.descripcion, mae_documento.abrev, mae_cliente.numruc, mae_condpago.abrev, ped_pedido.fchemi & '', vta_puntoVenta.descripcion, ped_tipo.descripcion, mae_unidades.abrev " _
'            + vbCr + "HAVING (" & c_CLIENTES & c_OC & c_PRODUCTOS & c_CONDICION & c_FECHA_EMI & " And " & c_FECHA_PLAZ3 & ");" _
'            + vbCr + "Union " _
'            + vbCr + "SELECT ped_pedido.oc, mae_cliente.nombre AS nombre, ped_pedido.id, ped_pedidodetent.iditem, alm_inventario.descripcion, ped_pedido.idcli, ped_pedido.idpunvecli, ped_pedido.tipdoc, ped_pedido.numser, ped_pedido.numdoc, ped_pedido.idconpag, ped_pedido.fchemi, ped_pedido.glosa, ped_pedido.idtipped, ped_pedidodetent.fchent As fEnt, ped_pedidodetent.canpro, IIf([ped_pedido].[anulado]=0,'VIGENTE','ANULADO') AS Anulado, ped_pedido.numreg, ped_pedido.idlib, ped_pedido.proceso, IIf(IsNull(ped_pedido!numser)=-1,ped_pedido!numdoc,ped_pedido!numser+'-'+ped_pedido!numdoc) AS numerodoc, mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, mae_condpago.abrev AS conpagabre, ped_pedido.fchemi & '' AS fchemi1, vta_puntoVenta.descripcion AS ptovta, ped_tipo.descripcion AS tipped, mae_unidades.abrev " _
'            + vbCr + "FROM (((ped_tipo RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN (vta_puntoVenta RIGHT JOIN (ped_pedido LEFT JOIN mae_condpago ON ped_pedido.idconpag = mae_condpago.id) ON vta_puntoVenta.id = ped_pedido.idpunvecli) ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) ON ped_tipo.id = ped_pedido.idtipped) LEFT JOIN ped_pedidodetent ON ped_pedido.id = ped_pedidodetent.idped) LEFT JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodetent.idunimed = mae_unidades.id " _
'            + vbCr + "Where (((ped_pedido.idtipped) = 2)) " _
'            + vbCr + "GROUP BY ped_pedido.id, ped_pedido.idcli, ped_pedido.idpunvecli, ped_pedido.tipdoc, ped_pedido.numser, ped_pedido.numdoc, ped_pedido.idconpag, ped_pedido.fchemi, ped_pedido.glosa, ped_pedido.idtipped, ped_pedidodetent.fchent, IIf([ped_pedido].[anulado]=0,'VIGENTE','ANULADO'), ped_pedido.numreg, ped_pedido.idlib, ped_pedido.oc, ped_pedido.proceso, mae_cliente.nombre, IIf(IsNull(ped_pedido!numser)=-1,ped_pedido!numdoc,ped_pedido!numser+'-'+ped_pedido!numdoc), mae_documento.descripcion, mae_condpago.descripcion, mae_documento.abrev, mae_cliente.numruc, mae_condpago.abrev, ped_pedido.fchemi & '', vta_puntoVenta.descripcion, ped_tipo.descripcion, alm_inventario.descripcion, mae_unidades.abrev, ped_pedidodetent.canpro, ped_pedidodetent.iditem " _
'            + vbCr + "HAVING (" & c_CLIENTES & c_OC & c_PRODUCTOS & c_CONDICION & c_FECHA_EMI & " And " & c_FECHA_PLAZ2 & ");"
            
            RST_Busq RstLis, c_SQL, xCon
    Set FgReporteProd.DataSource = RstLis.DataSource
    configurarFgReporte 0
    
    FgReporteProd.ColSel = 15
    FgReporteProd.Sort = flexSortGenericDescending
    Set RstLis = Nothing
End Sub

Private Sub generarConsultaDetallado()
    generarConsultaResumido
    FgDet.Rows = FgReporteProd.Rows * 20
    FgDet.Cols = FgReporteProd.Cols * 2
    If OptionFechas.Value = True Then genConDetFechas
    If OptionClientes.Value = True Then genConDetClientes
    If OptionProductos.Value = True Then genConDetProductos
End Sub

Private Sub genConDetFechas()
    Dim filRes As Integer
    Dim filDet As Integer
    Dim fechAux As String
    
    filRes = 1
    filDet = 1
    fechAux = FgReporteProd.TextMatrix(1, 15)
    filDet = filDet + 1
    FgDet.TextMatrix(1, 1) = fechAux
    
    FgDet.Select 1, 1, 1, 1
    FgDet.FillStyle = flexFillRepeat
    FgDet.CellForeColor = &H8000000D
    
    While filRes < FgReporteProd.Rows - 1
        If fechAux <> FgReporteProd.TextMatrix(filRes, 15) Then
            filDet = filDet + 1
            fechAux = FgReporteProd.TextMatrix(filRes, 15)
            FgDet.TextMatrix(filDet, 1) = fechAux
            
            FgDet.Select filDet, 1, filDet, 1
            FgDet.FillStyle = flexFillRepeat
            FgDet.CellForeColor = &H8000000D
            
            filDet = filDet + 1
        End If
        FgDet.TextMatrix(filDet, 1) = FgReporteProd.TextMatrix(filRes, 21)
        FgDet.TextMatrix(filDet, 2) = FgReporteProd.TextMatrix(filRes, 1)
        FgDet.TextMatrix(filDet, 3) = FgReporteProd.TextMatrix(filRes, 5)
        FgDet.TextMatrix(filDet, 4) = FgReporteProd.TextMatrix(filRes, 12)
        FgDet.TextMatrix(filDet, 5) = FgReporteProd.TextMatrix(filRes, 16)
        FgDet.TextMatrix(filDet, 6) = FgReporteProd.TextMatrix(filRes, 30)
        FgDet.TextMatrix(filDet, 7) = FgReporteProd.TextMatrix(filRes, 2)
        FgDet.TextMatrix(filDet, 8) = FgReporteProd.TextMatrix(filRes, 25)
        FgDet.TextMatrix(filDet, 9) = FgReporteProd.TextMatrix(filRes, 28)
        FgDet.TextMatrix(filDet, 10) = FgReporteProd.TextMatrix(filRes, 17)
        filRes = filRes + 1
        filDet = filDet + 1
    Wend
    
    FgDet.Cols = 11
    FgDet.Rows = filRes - 1
    configurarFgDetallado
End Sub

Private Sub genConDetClientes()
    Dim filRes As Integer
    Dim filDet As Integer
    Dim CliAux As String
    
    FgReporteProd.ColSel = 2
    FgReporteProd.Sort = flexSortGenericAscending
    
    filRes = 1
    filDet = 1
    CliAux = FgReporteProd.TextMatrix(1, 2)
    filDet = filDet + 1
    FgDet.TextMatrix(1, 1) = CliAux
            
    FgDet.Select 1, 1, 1, 1
    FgDet.CellAlignment = flexAlignLeftCenter
    FgDet.FillStyle = flexFillRepeat
    FgDet.CellForeColor = &H8000000D
    
    While filRes < FgReporteProd.Rows - 1
        If CliAux <> FgReporteProd.TextMatrix(filRes, 2) Then
            filDet = filDet + 1
            CliAux = FgReporteProd.TextMatrix(filRes, 2)
            FgDet.TextMatrix(filDet, 1) = CliAux
            
            FgDet.Select filDet, 1, filDet, 1
            FgDet.CellAlignment = flexAlignLeftCenter
            
            FgDet.FillStyle = flexFillRepeat
            FgDet.CellForeColor = &H8000000D
            
            filDet = filDet + 1
        End If
        FgDet.TextMatrix(filDet, 1) = FgReporteProd.TextMatrix(filRes, 15)
        FgDet.TextMatrix(filDet, 2) = FgReporteProd.TextMatrix(filRes, 1)
        FgDet.TextMatrix(filDet, 3) = FgReporteProd.TextMatrix(filRes, 5)
        FgDet.TextMatrix(filDet, 4) = FgReporteProd.TextMatrix(filRes, 12)
        FgDet.TextMatrix(filDet, 5) = FgReporteProd.TextMatrix(filRes, 16)
        FgDet.TextMatrix(filDet, 6) = FgReporteProd.TextMatrix(filRes, 30)
        FgDet.TextMatrix(filDet, 7) = FgReporteProd.TextMatrix(filRes, 21)
        FgDet.TextMatrix(filDet, 8) = FgReporteProd.TextMatrix(filRes, 25)
        FgDet.TextMatrix(filDet, 9) = FgReporteProd.TextMatrix(filRes, 28)
        FgDet.TextMatrix(filDet, 10) = FgReporteProd.TextMatrix(filRes, 17)
        filRes = filRes + 1
        filDet = filDet + 1
    Wend
    FgDet.Cols = 11
    FgDet.Rows = filRes
    configurarFgDetallado
End Sub

Private Sub llenarGuias(ByRef filDetAux As Integer, ByRef filRes As Integer, ByRef sumGuia As Double)
    Dim RstGuia As New ADODB.Recordset
    Dim cSQL As String
    Dim A As Integer
    
    cSQL = "SELECT vta_guia.id, vta_guia.numordcom, vta_guia.idcli, vta_guiadet.iditem, vta_guia.fchentord, vta_guiadet.canpro, IIf(IsNull([vta_guia]![numser])=-1,[vta_guia]![numdoc],[vta_guia]![numser]+'-'+[vta_guia]![numdoc]) AS numerodoc " _
        + vbCr + "FROM vta_guia INNER JOIN vta_guiadet ON vta_guia.id = vta_guiadet.idgui " _
        + vbCr + "GROUP BY vta_guia.id, vta_guia.numordcom, vta_guia.idcli, vta_guiadet.iditem, vta_guia.fchentord, vta_guiadet.canpro, IIf(IsNull([vta_guia]![numser])=-1,[vta_guia]![numdoc],[vta_guia]![numser]+'-'+[vta_guia]![numdoc]), vta_guia.Anulado " _
        + vbCr + "Having (((vta_guia.numordcom) = '" & FgReporteProd.TextMatrix(filRes - 1, 1) & "') And ((vta_guiadet.iditem) = " & NulosN(FgReporteProd.TextMatrix(filRes - 1, 4)) & ") And ((vta_guia.idcli) =" & FgReporteProd.TextMatrix(filRes - 1, 6) & ") And ((vta_guia.fchentord) > CDate('" & TxtFchEntDesde.Valor & "') And (vta_guia.fchentord) < CDate('" & TxtFchEntHasta.Valor & "')) And ((vta_guia.Anulado) = 0)) " _
        + vbCr + "ORDER BY vta_guia.fchentord ASC;"
    
    RST_Busq RstGuia, cSQL, xCon
    
    With FgDet
        If Not RstGuia.EOF Then
            RstGuia.MoveFirst
            While Not RstGuia.EOF
                sumGuia = sumGuia + NulosN(RstGuia("canpro"))
                .TextMatrix(filDetAux, 5) = RstGuia("canpro")
                .TextMatrix(filDetAux, 7) = RstGuia("fchentord")
                '.TextMatrix(filDetAux, 8) = RstGuia("numordcom")
                .TextMatrix(filDetAux, 9) = RstGuia("numerodoc")
                '.TextMatrix(filDetAux, 10) = RstGuia("id")
                filDetAux = filDetAux + 1
                RstGuia.MoveNext
            Wend
        Else
            .TextMatrix(filDetAux, 5) = "-"
            .TextMatrix(filDetAux, 7) = "-"
            '.TextMatrix(filDetAux, 8) = "-"
            '.TextMatrix(filDetAux, 9) = "-"
            '.TextMatrix(filDetAux, 10) = "-"
            filDetAux = filDetAux + 1
        End If
    End With
End Sub

Private Sub llenarPed(ByRef filDet As Integer, ByRef filRes As Integer, ByRef sumPed As Double)
    'Se llena el primer detalle de un pedido
    With FgDet
        .TextMatrix(filDet, 1) = FgReporteProd.TextMatrix(filRes, 2)
        .ColAlignment(1) = flexAlignRightCenter
        .TextMatrix(filDet, 2) = FgReporteProd.TextMatrix(filRes, 1)
        .TextMatrix(filDet, 3) = FgReporteProd.TextMatrix(filRes, 15)
        .TextMatrix(filDet, 4) = FgReporteProd.TextMatrix(filRes, 16)
        .TextMatrix(filDet, 8) = FgReporteProd.TextMatrix(filRes, 25)
        .TextMatrix(filDet, 10) = FgReporteProd.TextMatrix(filRes, 17)
        sumPed = sumPed + NulosN(FgReporteProd.TextMatrix(filRes, 16))
        filDet = filDet + 1
        filRes = filRes + 1
    End With
End Sub

Private Sub llenarTotales(ByRef filDet As Integer, ByRef sumPed As Double, sumGuia As Double)
    With FgDet
        .Select filDet, 3, filDet, 7
        .FillStyle = flexFillRepeat
        .CellBackColor = &HC0FFFF
        .CellForeColor = &H8000000D
        
        .TextMatrix(filDet, 3) = "Total"
        .TextMatrix(filDet, 4) = NulosN(sumPed)
        .TextMatrix(filDet, 5) = NulosN(sumGuia)
        .TextMatrix(filDet, 6) = sumGuia - sumPed
        
        If (sumPed - sumGuia) < 0 Then .TextMatrix(filDet, 7) = "EXCEDIDO": .Select filDet, 7, filDet, 7: .FillStyle = flexFillRepeat: .CellForeColor = &HC0C000
        If (sumPed - sumGuia) > 0 Then .TextMatrix(filDet, 7) = "INCOMPLETO": .Select filDet, 7, filDet, 7: .FillStyle = flexFillRepeat: .CellForeColor = &HFF&
        If (sumPed - sumGuia) = 0 Then .TextMatrix(filDet, 7) = "COMPLETO": .Select filDet, 7, filDet, 7: .FillStyle = flexFillRepeat: .CellForeColor = &HC000&
        
        sumPed = 0
        sumGuia = 0
    End With
End Sub

Private Sub genConDetProductos()
    Dim RstLis As New ADODB.Recordset
    Dim RstGuia As New ADODB.Recordset
    Dim filRes As Integer
    Dim filDet As Integer
    Dim filDetAux As Integer
    Dim ProdAux As String
    Dim OrdAux As String
    Dim cSQL As String
    Dim sumPed As Double
    Dim sumGuia As Double
    
    FgReporteProd.ColSel = 5
    FgReporteProd.Sort = flexSortGenericAscending
    
    configurarFgDetallado
    
    filRes = 1
    filDet = 1
    ProdAux = FgReporteProd.TextMatrix(1, 5)
    OrdAux = FgReporteProd.TextMatrix(1, 1)
    
    FgDet.TextMatrix(1, 1) = ProdAux
    filDet = filDet + 1
    
    llenarPed filDet, filRes, sumPed
    filDetAux = filDet - 1
    llenarGuias filDetAux, filRes, sumGuia
    
    FgDet.Select 1, 1, 1, 1
    FgDet.FillStyle = flexFillRepeat
    FgDet.CellForeColor = &H8000000D '&H8000000F&
    
    'Se recorre la consulta resumida
    FraProgreso.Visible = True
    FraProgreso.Refresh
    PgBar.Max = FgReporteProd.Rows
    While filRes < FgReporteProd.Rows - 1
        DoEvents
        If interrumpir = True Then Exit Sub
        PgBar.Value = filRes
        'Si es otro Producto
        If ProdAux <> FgReporteProd.TextMatrix(filRes, 5) Then
            'Escribimos Cantidades Totales
            If filDetAux > filDet Then filDet = filDetAux
            
            llenarTotales filDet, sumPed, sumGuia
            
            filDet = filDet + 2
            ProdAux = FgReporteProd.TextMatrix(filRes, 5)
            OrdAux = FgReporteProd.TextMatrix(filRes, 1)
            
            FgDet.TextMatrix(filDet, 1) = ProdAux
            FgDet.Select filDet, 1, filDet, 1
            FgDet.FillStyle = flexFillRepeat
            FgDet.CellForeColor = &H8000000D
            
            filDet = filDet + 1
            llenarPed filDet, filRes, sumPed
            filDetAux = filDet - 1
            llenarGuias filDetAux, filRes, sumGuia
            
        Else
            'Si es otra Orden
            If FgReporteProd.TextMatrix(filRes, 1) <> OrdAux Then
                If filDetAux > filDet Then filDet = filDetAux
                OrdAux = FgReporteProd.TextMatrix(filRes, 1)
                'Escribimos Cantidades
                llenarTotales filDet, sumPed, sumGuia
                
                filDet = filDet + 1
                llenarPed filDet, filRes, sumPed
                filDetAux = filDet - 1
                llenarGuias filDetAux, filRes, sumGuia
            Else
                llenarPed filDet, filRes, sumPed
            End If
        End If
    Wend
    If filDetAux > filDet Then filDet = filDetAux
    
    llenarTotales filDet, sumPed, sumGuia
    FraProgreso.Visible = False
    
    FgDet.Cols = 11
    FgDet.Rows = filDet + 1
End Sub

Private Sub iniciarCampos()
    cargo = False
    
    Set FgProductos.DataSource = Nothing
    Set FgClientes.DataSource = Nothing
    Set FgOC.DataSource = Nothing
    'Se inicializa:
    'datos para productos
    FgProductos.Rows = 1
    FgProductos.Cols = 2
    FgProductos.Row = 0
    FgProductos.Col = 1
    FgProductos.Text = "Todos"
    'datos para clientes
    FgClientes.Rows = 1
    FgClientes.Cols = 2
    FgClientes.Row = 0
    FgClientes.Col = 1
    FgClientes.Text = "Todos"
    'datos para Ordenes de Compra
    FgOC.Rows = 1
    FgOC.Cols = 2
    FgOC.Row = 0
    FgOC.Col = 1
    FgOC.Text = "Todos"
    
    'datos para fechas
    TxtFchEmiDesde.Valor = CDate("01/01/2010")
    TxtFchEmiHasta.Valor = Date
    TxtFchEntDesde.Valor = CDate("01/" + CStr(Month(Date) - 1) + "/" + CStr(Year(Date)))
    TxtFchEntHasta.Valor = Date
    
    'configuracion de Grids
    FgProductos.Editable = flexEDKbdMouse
    FgProductos.ColComboList(1) = "..."
    FgProductos.ShowComboButton = flexSBAlways
    
    FgClientes.Editable = flexEDKbdMouse
    FgClientes.ColComboList(1) = "..."
    FgClientes.ShowComboButton = flexSBAlways
    
    FgOC.Editable = flexEDKbdMouse
    FgOC.ColComboList(1) = "..."
    FgOC.ShowComboButton = flexSBAlways
    
    CheckCond.Value = 1
    CheckResumido.Value = 1
    
    FgReporteProd.AllowUserResizing = flexResizeColumns
    FgReporteProd.AutoSearch = flexSearchFromTop
    FgReporteProd.ExplorerBar = flexExSortShowAndMove
    FgReporteProd.SelectionMode = flexSelectionByRow
    FgReporteProd.ForeColorSel = &H0&
    FgReporteProd.BackColorSel = &HC0E0FF
    
    FgDet.AllowUserResizing = flexResizeColumns
    FgDet.AutoSearch = flexSearchFromTop
    FgDet.SelectionMode = flexSelectionByRow
    FgDet.ForeColorSel = &H0&
    FgDet.BackColorSel = &HC0E0FF
    
    FgReporteProd.Top = 270
    FgReporteProd.Left = 90
    FgReporteProd.Height = 4650
    FgReporteProd.Width = 11655
End Sub

Private Sub configurarFgReporte(tipo As Integer)
    Select Case tipo
        Case 0
            configurarFgResumido
        Case 1
            configurarFgDetallado
    End Select
End Sub

Private Sub configurarFgResumido()
    FgReporteProd.Cols = 31
    FgReporteProd.FixedRows = 1
    FgReporteProd.FrozenCols = 5
    FgReporteProd.ColWidth(0) = 0
    FgReporteProd.RowHeight(0) = 300
    FgReporteProd.Row = 0
    
    FgReporteProd.Col = 1
    FgReporteProd.ColAlignment(1) = flexAlignLeftCenter
    FgReporteProd.Text = "Ord. Pedido"
    FgReporteProd.ColWidth(1) = 1000
    
    FgReporteProd.ColWidth(2) = 0
    
    FgReporteProd.Col = 2
    FgReporteProd.Text = "Cliente"
    FgReporteProd.ColWidth(2) = 1800
    
    FgReporteProd.ColWidth(3) = 0
    FgReporteProd.ColWidth(4) = 0

    FgReporteProd.Col = 5
    FgReporteProd.Text = "Producto"
    FgReporteProd.ColWidth(5) = 3700
    
    FgReporteProd.ColWidth(6) = 0
    FgReporteProd.ColWidth(7) = 0
    FgReporteProd.ColWidth(8) = 0
    FgReporteProd.ColWidth(9) = 0
    FgReporteProd.ColWidth(10) = 0
    FgReporteProd.ColWidth(11) = 0
    
    FgReporteProd.Col = 12
    FgReporteProd.Text = "Fech. Emision"
    FgReporteProd.ColWidth(12) = 1200
    
    FgReporteProd.ColWidth(13) = 0
    FgReporteProd.ColWidth(14) = 0
    
    FgReporteProd.Col = 15
    FgReporteProd.Text = "Fech. A Entregar"
    FgReporteProd.ColWidth(15) = 1300
    
    FgReporteProd.Col = 16
    FgReporteProd.Text = "Cant. A Entregar"
    FgReporteProd.ColWidth(16) = 1300
    
    FgReporteProd.Col = 17
    FgReporteProd.Text = "Condicion"
    FgReporteProd.ColWidth(17) = 1000
    
    FgReporteProd.ColWidth(18) = 0
    FgReporteProd.ColWidth(19) = 0
    FgReporteProd.ColWidth(20) = 0
    
    FgReporteProd.Col = 21
    FgReporteProd.Text = "Num. Documento"
    FgReporteProd.ColWidth(21) = 0
    
    FgReporteProd.ColWidth(22) = 0
    FgReporteProd.ColWidth(23) = 0
    FgReporteProd.ColWidth(24) = 0
    
    FgReporteProd.Col = 25
    FgReporteProd.Text = "Num. RUC"
    FgReporteProd.ColWidth(25) = 0
    
    FgReporteProd.ColWidth(26) = 0
    FgReporteProd.ColWidth(27) = 0

    FgReporteProd.Col = 28
    FgReporteProd.Text = "Direccion"
    FgReporteProd.ColWidth(28) = 0
    
    FgReporteProd.ColWidth(29) = 0
    FgReporteProd.ColWidth(30) = 0
End Sub

Private Sub configurarFgDetallado()
    If OptionFechas.Value = True Then configFgDetFechas
    If OptionClientes.Value = True Then configFgDetClientes
    If OptionProductos.Value = True Then configFgDetProductos
End Sub

Private Sub configFgDetFechas()
    FgDet.Cols = 11
    FgDet.FixedRows = 1
    FgDet.FrozenCols = 3
    FgDet.ColWidth(0) = 0
    FgDet.RowHeight(0) = 300
    FgDet.Row = 0

    FgDet.Col = 1
    FgDet.Text = "Num. Doc."
    FgDet.ColWidth(1) = 1400

    FgDet.Col = 2
    FgDet.Text = "Ord. Pedido"
    FgDet.ColWidth(2) = 1000

    FgDet.Col = 3
    FgDet.Text = "Producto"
    FgDet.ColWidth(3) = 3000

    FgDet.Col = 4
    FgDet.Text = "Fech. Emi."
    FgDet.ColWidth(4) = 1000

    FgDet.Col = 5
    FgDet.Text = "Cant."
    FgDet.ColWidth(5) = 500

    FgDet.Col = 6
    FgDet.Text = "Unid"
    FgDet.ColWidth(6) = 430
    
    FgDet.Col = 7
    FgDet.Text = "Cliente"
    FgDet.ColWidth(7) = 1800

    FgDet.Col = 8
    FgDet.Text = "Num. RUC"
    FgDet.ColWidth(8) = 1100

    FgDet.Col = 9
    FgDet.Text = "Direccion"
    FgDet.ColWidth(9) = 1990

    FgDet.Col = 10
    FgDet.Text = "Condicion"
    FgDet.ColWidth(10) = 1100
End Sub

Private Sub configFgDetClientes()
    FgDet.Cols = 11
    FgDet.FixedRows = 1
    FgDet.FrozenCols = 1
    FgDet.ColWidth(0) = 0
    FgDet.RowHeight(0) = 300
    FgDet.Row = 0

    FgDet.Col = 1
    FgDet.Text = "Fch. A Entreg."
    FgDet.ColWidth(1) = 2000
    FgDet.ColAlignment(1) = flexAlignRightCenter

    FgDet.Col = 2
    FgDet.Text = "Ord. Pedido"
    FgDet.ColWidth(2) = 1100

    FgDet.Col = 3
    FgDet.Text = "Producto"
    FgDet.ColWidth(3) = 3700

    FgDet.Col = 4
    FgDet.Text = "Fech. Emi."
    FgDet.ColWidth(4) = 1000

    FgDet.Col = 5
    FgDet.Text = "Cantidad"
    FgDet.ColWidth(5) = 800

    FgDet.Col = 6
    FgDet.Text = "Unid"
    FgDet.ColWidth(6) = 500
    
    FgDet.Col = 7
    FgDet.Text = "Num. Doc."
    FgDet.ColWidth(7) = 1100

    FgDet.Col = 8
    FgDet.Text = "Num. RUC"
    FgDet.ColWidth(8) = 1100

    FgDet.Col = 9
    FgDet.Text = "Direccion"
    FgDet.ColWidth(9) = 1990

    FgDet.Col = 10
    FgDet.Text = "Condicion"
    FgDet.ColWidth(10) = 1100
End Sub

Private Sub configFgDetProductos()
    FgDet.Cols = 11
    FgDet.FixedRows = 1
    FgDet.FrozenCols = 1
    FgDet.ColWidth(0) = 0
    FgDet.RowHeight(0) = 300
    FgDet.Row = 0

    FgDet.Col = 1
    FgDet.Text = "Cliente"
    FgDet.ColWidth(1) = 3500

    FgDet.Col = 2
    FgDet.Text = "Ord. Pedido"
    FgDet.ColWidth(2) = 1000

    FgDet.Col = 3
    FgDet.Text = "Fch.A Entreg."
    FgDet.ColWidth(3) = 1000

    FgDet.Col = 4
    FgDet.Text = "Cant.AEnt."
    FgDet.ColWidth(4) = 850

    FgDet.Col = 5
    FgDet.Text = "Cant.Entr."
    FgDet.ColWidth(5) = 800

    FgDet.Col = 6
    FgDet.Text = "Dif."
    FgDet.ColWidth(6) = 600
    
    FgDet.Col = 7
    FgDet.Text = "Fch.Entreg."
    FgDet.ColWidth(7) = 1000

    FgDet.Col = 8
    FgDet.Text = "Num. RUC"
    FgDet.ColWidth(8) = 1100

    FgDet.Col = 9
    FgDet.Text = "Num. Doc."
    FgDet.ColWidth(9) = 1990

    FgDet.Col = 10
    FgDet.Text = "Condicion"
    FgDet.ColWidth(10) = 1100
End Sub

Private Sub limpiarReporte()
    FgReporteProd.Cols = 2
    FgReporteProd.Rows = 2
    FgReporteProd.FixedRows = 1
    FgReporteProd.FixedCols = 1
    
    FgDet.Cols = 2
    FgDet.Rows = 2
    FgDet.FixedRows = 1
    FgDet.FixedCols = 1
End Sub

Private Sub verificarCheckCondicion()
    If CheckVigent.Value = 1 And CheckAnul.Value = 1 Then CheckCond.Value = 1
    If CheckVigent.Value = 0 Or CheckAnul.Value = 0 Then CheckCond.Value = 0
    If CheckVigent.Value = 0 And CheckAnul.Value = 0 Then CheckCond.Value = 1
End Sub

Private Sub CheckDetallado_Click()
    OptionFechas.Enabled = True
    OptionClientes.Enabled = True
    OptionProductos.Enabled = True
    OptionFechas.Value = True
End Sub

Private Sub CheckResumido_Click()
    OptionFechas.Enabled = False
    OptionClientes.Enabled = False
    OptionProductos.Enabled = False
    
    OptionFechas.Value = False
    OptionClientes.Value = False
    OptionProductos.Value = False
End Sub

Private Sub CheckResumido_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    verificarCheckTipo
End Sub

Private Sub CheckDetallado_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    verificarCheckTipo
End Sub

Private Sub CheckVigent_Click()
    verificarCheckCondicion
End Sub

Private Sub CheckAnul_Click()
    verificarCheckCondicion
End Sub

Private Sub verificarCheckTipo()
    If CheckResumido.Value = 0 Then CheckDetallado.Value = 0
    If CheckDetallado.Value = 0 Then CheckResumido.Value = 0
End Sub

Private Sub CheckCond_Click()
    If estadoCheckCondicion Then
        estadoCheckCondicion = False
    Else
        estadoCheckCondicion = True
        CheckVigent.Value = 1
        CheckAnul.Value = 1
    End If
End Sub

Private Sub Form_Load()
    iniciarCampos
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        interrumpir = True
        FraProgreso.Visible = False
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        If verificarDatos Then
            cargo = True
            If CheckResumido.Value = 1 Then
                generarConsultaResumido
                FgReporteProd.Top = 270
                FgReporteProd.Left = 90
                FgReporteProd.Height = 4650
                FgReporteProd.Width = 11655
                FgReporteProd.Visible = True
                FgDet.Visible = False
            Else
                generarConsultaDetallado
                FgDet.Top = 270
                FgDet.Left = 90
                FgDet.Height = 4650
                FgDet.Width = 11655
                FgDet.Visible = True
                FgReporteProd.Visible = False
            End If
        End If
    End If
    If Button.Index = 5 Then
        If Not cargo Then MsgBox "No se ha procesado ninguna Consulta, procesela antes de Exportar", vbCritical + vbOKOnly, "Reporte de Pedidos": Exit Sub
        If CheckDetallado.Value = 1 Then ExportarExcel FgDet
        If CheckResumido.Value = 1 Then ExportarExcel FgReporteProd
    End If
    If Button.Index = 8 Then
        Unload Me
    End If
End Sub

Private Sub FgProductos_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    FgProductos.ShowComboButton = flexSBFocus
    
    Dim nSQL As String
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 3) As String
    xCampos(0, 0) = "Nombre":          xCampos(0, 1) = "descripcion":     xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "codpro":            xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"

    nSQL = "SELECT alm_inventario.descripcion, alm_inventario.codpro " _
        + vbCr + "From alm_inventario " _
        + vbCr + "WHERE (((alm_inventario.activo)=-1) AND ((alm_inventario.tippro) In (1,3)) AND ((alm_inventario.idcuentaven)<>0)) " _
        + vbCr + "GROUP BY alm_inventario.descripcion, alm_inventario.codpro " _
        + vbCr + "ORDER BY alm_inventario.descripcion;"
        
    xform.SQLCad = nSQL
    
    xform.Titulo = "Buscando Productos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        FgProductos.TextMatrix(FgProductos.Row, 1) = xRs.Fields(0) & ""
        If FgProductos.Row = FgProductos.Rows - 1 Then FgProductos.AddItem ""
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub FgClientes_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    FgClientes.ShowComboButton = flexSBFocus
    
    Dim nSQL As String
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    Dim xCampos(2, 3) As String
    xCampos(0, 0) = "Nombre":          xCampos(0, 1) = "nombre":     xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.UC.":        xCampos(1, 1) = "numruc":     xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
    
    nSQL = "SELECT mae_cliente.nombre, mae_cliente.numruc " _
           + vbCr + "From mae_cliente " _
           + vbCr + "WHERE mae_cliente.nombre <> ''" _
           + vbCr + "GROUP BY mae_cliente.id, mae_cliente.nombre, mae_cliente.numruc " _
           + vbCr + "ORDER BY mae_cliente.nombre;"
    xform.SQLCad = nSQL
    
    xform.Titulo = "Buscando Clientes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        FgClientes.TextMatrix(FgClientes.Row, 1) = xRs.Fields(0) & ""
        If FgClientes.Row = FgClientes.Rows - 1 Then FgClientes.AddItem ""
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub FgProductos_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        If FgProductos.Rows = 0 Then FgProductos.AddItem ("Todos")
    End If
    If KeyCode = 46 Then
        On Error GoTo MAY
        FgProductos.RemoveItem FgProductos.Row
        If FgProductos.Rows = 0 Then FgProductos.AddItem (""): FgProductos.TextMatrix(0, 1) = "Todos"
    End If
    Exit Sub
MAY:
    Exit Sub
End Sub

Private Sub FgClientes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        If FgClientes.Rows = 0 Then FgClientes.AddItem ("Todos")
    End If
    If KeyCode = 46 Then
        On Error GoTo MAY
        FgClientes.RemoveItem FgClientes.Row
        If FgClientes.Rows = 0 Then FgClientes.AddItem (""): FgClientes.TextMatrix(0, 1) = "Todos"
    End If
    Exit Sub
MAY:
    Exit Sub
End Sub

Private Sub FgOC_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    FgOC.ShowComboButton = flexSBFocus
    
    Dim nSQL As String
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 3) As String
    
    xCampos(0, 0) = "Orden de Compra":    xCampos(0, 1) = "oc":       xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id Pedido":          xCampos(1, 1) = "idped":    xCampos(1, 2) = "2500":   xCampos(1, 3) = "C"

'SELECT DISTINCT ped_pedido.oc, ped_pedido.id AS idped
'From ped_pedido
'GROUP BY ped_pedido.oc, ped_pedido.id
'HAVING (((ped_pedido.oc) Is Not Null And (ped_pedido.oc)<>'S/N' And (ped_pedido.oc)<>'') AND ((ped_pedido.id) Is Not Null))
'ORDER BY ped_pedido.oc;

    nSQL = "SELECT DISTINCT ped_pedido.oc, ped_pedido.id AS idped " _
         + vbCr + "From ped_pedido " _
         + vbCr + "GROUP BY ped_pedido.oc, ped_pedido.id " _
         + vbCr + "HAVING (((ped_pedido.oc) Is Not Null And (ped_pedido.oc)<>'S/N' And (ped_pedido.oc)<>'') AND ((ped_pedido.id) Is Not Null)) " _
         + vbCr + "ORDER BY ped_pedido.oc;"

'    nSQL = "SELECT DISTINCT ped_pedido.oc, ped_pedidodetent.idped " _
'         + vbCr + "FROM ped_pedidodetent RIGHT JOIN ped_pedido ON ped_pedidodetent.idped = ped_pedido.id " _
'         + vbCr + "GROUP BY ped_pedido.oc, ped_pedidodetent.idped " _
'         + vbCr + "Having (((ped_pedido.oc) Is Not Null And (ped_pedido.oc) <> 'S/N' And (ped_pedido.oc) <> '') And ((ped_pedidodetent.idped) Is Not Null)) " _
'         + vbCr + "ORDER BY ped_pedido.oc;"
         
    xform.SQLCad = nSQL
    
    xform.Titulo = "Buscando Ordenes de Compra"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "oc"
    xform.CampoBusca = "oc"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        FgOC.TextMatrix(FgOC.Row, 1) = xRs.Fields(0) & ""
        If FgOC.Row = FgOC.Rows - 1 Then FgOC.AddItem ""
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Sub ExportarExcel(ByRef FgGrid As VSFlexGrid)
    Dim A As Integer
    Dim B As Integer
    Dim xFilas As Integer
    Dim xCad As String
    Dim objExcel As Object
    
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    objExcel.WindowState = 2
    objExcel.Workbooks.Add
   
    With objExcel.ActiveSheet
        .cells(1, 2) = "Fecha de Hoy: "
        .cells(1, 3) = "'" & Date
        .cells(2, 2) = "Fecha de Emision del Pedido   Desde: "
        .cells(2, 3) = "'" + TxtFchEmiDesde.Valor
        .cells(2, 4) = "Hasta: "
        .cells(2, 5) = "'" + TxtFchEmiHasta.Valor
        .cells(3, 2) = "Fecha a Entregar del Pedido   Desde: "
        .cells(3, 3) = "'" + TxtFchEntDesde.Valor
        .cells(3, 4) = "Hasta: "
        .cells(3, 5) = "'" + TxtFchEntHasta.Valor
        
        .cells(4, 2) = "Clientes: "
        xFilas = 5
        For A = 0 To FgClientes.Rows - 1
            .cells(xFilas, 3) = FgClientes.TextMatrix(A, 1)
            xFilas = xFilas + 1
        Next A
        .cells(4, 5) = "Productos: "
        Dim xFilasAux As Integer
        xFilasAux = 5
        For A = 0 To FgProductos.Rows - 1
            .cells(xFilasAux, 6) = FgProductos.TextMatrix(A, 1)
            xFilasAux = xFilasAux + 1
        Next A
        
        If (xFilas < xFilasAux) Then xFilas = xFilasAux
        
        xFilas = xFilas + 1
        For A = 0 To FgGrid.Rows - 1
            For B = 1 To FgGrid.Cols - 1
                If A = 0 Then
                    .cells(xFilas, B + 1) = "'" + FgGrid.TextMatrix(A, B)
                Else
                    If (B = 1 Or B = 3 Or B = 7) Then
                        .cells(xFilas, B + 1) = FgGrid.TextMatrix(A, B)
                    Else
                        .cells(xFilas, B + 1) = "'" + FgGrid.TextMatrix(A, B)
                    End If
                End If
            Next B
            xFilas = xFilas + 1
        Next A
    End With
    
    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Reporte de Pedidos"
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
End Sub

Private Sub FgOC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        If FgOC.Rows = 0 Then FgOC.AddItem ("")
    End If
    If KeyCode = 46 Then
        On Error GoTo MAY
        FgOC.RemoveItem FgOC.Row
        If FgOC.Rows = 0 Then FgOC.AddItem (""): FgOC.TextMatrix(0, 1) = "Todos"
    End If
    Exit Sub
MAY:
    Exit Sub
End Sub

