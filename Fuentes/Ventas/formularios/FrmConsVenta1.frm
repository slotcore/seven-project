VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsVenta1 
   Caption         =   "Ventas - Consulta de Ventas"
   ClientHeight    =   7995
   ClientLeft      =   120
   ClientTop       =   615
   ClientWidth     =   11895
   HasDC           =   0   'False
   Icon            =   "FrmConsVenta1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7995
   ScaleWidth      =   11895
   Begin SizerOneLibCtl.TabOne TabOne2 
      Height          =   1305
      Left            =   30
      TabIndex        =   7
      Top             =   360
      Width           =   11880
      _cx             =   20955
      _cy             =   2302
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   0
      FrontTabForeColor=   -2147483630
      Caption         =   "Inicio|Mas"
      Align           =   0
      CurrTab         =   0
      FirstTab        =   0
      Style           =   3
      Position        =   2
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   -1  'True
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   1215
         Left            =   345
         TabIndex        =   9
         Top             =   45
         Width           =   11490
         Begin VB.Frame Frame1 
            Caption         =   "Moneda"
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
            Height          =   1125
            Left            =   5040
            TabIndex        =   20
            Top             =   30
            Width           =   1080
            Begin VB.OptionButton OptMonTodos 
               Caption         =   "Todos"
               Height          =   195
               Left            =   90
               TabIndex        =   23
               Top             =   270
               Value           =   -1  'True
               Width           =   840
            End
            Begin VB.OptionButton OptSol 
               Caption         =   "Soles"
               Height          =   195
               Left            =   90
               TabIndex        =   22
               Top             =   540
               Width           =   750
            End
            Begin VB.OptionButton OptDol 
               Caption         =   "Dolares"
               Height          =   195
               Left            =   90
               TabIndex        =   21
               Top             =   810
               Width           =   840
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Fecha"
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
            Height          =   1125
            Left            =   3780
            TabIndex        =   14
            Top             =   30
            Width           =   1215
            Begin VB.OptionButton OptReg 
               Caption         =   "Fch. Reg."
               Height          =   195
               Left            =   60
               TabIndex        =   41
               Top             =   810
               Width           =   1080
            End
            Begin VB.OptionButton OptVenc 
               Caption         =   "Fch. Venc."
               Height          =   195
               Left            =   60
               TabIndex        =   16
               Top             =   540
               Width           =   1080
            End
            Begin VB.OptionButton OptEmi 
               Caption         =   "Fch. Emi."
               Height          =   195
               Left            =   60
               TabIndex        =   15
               Top             =   270
               Value           =   -1  'True
               Width           =   975
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "[ Tipo Consulta ]"
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
            Height          =   525
            Left            =   0
            TabIndex        =   17
            Top             =   630
            Width           =   3735
            Begin VB.OptionButton OptResum 
               Caption         =   "Resumen"
               Height          =   195
               Left            =   330
               TabIndex        =   19
               Top             =   240
               Value           =   -1  'True
               Width           =   1005
            End
            Begin VB.OptionButton OptDetalle 
               Caption         =   "Detallado"
               Height          =   195
               Left            =   2130
               TabIndex        =   18
               Top             =   240
               Width           =   975
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "[ Seleccionar Fecha ]"
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
            Height          =   585
            Left            =   0
            TabIndex        =   36
            Top             =   30
            Width           =   3735
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   570
               TabIndex        =   37
               Top             =   210
               Width           =   1305
               _ExtentX        =   2302
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
               Valor           =   "11/09/2008"
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
               Height          =   300
               Left            =   2370
               TabIndex        =   38
               Top             =   210
               Width           =   1305
               _ExtentX        =   2302
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
               Valor           =   "11/09/2008"
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Hasta"
               Height          =   195
               Index           =   2
               Left            =   1920
               TabIndex        =   40
               Top             =   300
               Width           =   420
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Desde"
               Height          =   195
               Index           =   1
               Left            =   60
               TabIndex        =   39
               Top             =   300
               Width           =   465
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Ordenar Por"
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
            Height          =   810
            Left            =   9960
            TabIndex        =   32
            Top             =   1260
            Visible         =   0   'False
            Width           =   1320
            Begin VB.OptionButton opt_orden 
               Caption         =   "Num. Reg."
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   35
               Top             =   180
               Value           =   -1  'True
               Width           =   1155
            End
            Begin VB.OptionButton opt_orden 
               Caption         =   "Nº Doc."
               Height          =   195
               Index           =   1
               Left            =   120
               TabIndex        =   34
               Top             =   375
               Width           =   1095
            End
            Begin VB.OptionButton opt_orden 
               Caption         =   "Fch. Doc"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   33
               Top             =   570
               Width           =   1125
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Seleccionar"
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
            Height          =   1125
            Left            =   6180
            TabIndex        =   25
            Top             =   30
            Width           =   1230
            Begin VB.OptionButton OptPag 
               Caption         =   "Pagados"
               Height          =   195
               Left            =   60
               TabIndex        =   28
               Top             =   810
               Width           =   945
            End
            Begin VB.OptionButton OptPend 
               Caption         =   "Pendientes"
               Height          =   195
               Left            =   60
               TabIndex        =   27
               Top             =   540
               Width           =   1095
            End
            Begin VB.OptionButton OptTodos 
               Caption         =   "Todos"
               Height          =   195
               Left            =   60
               TabIndex        =   26
               Top             =   270
               Value           =   -1  'True
               Width           =   840
            End
         End
         Begin VB.CommandButton CmdBusProducto 
            Height          =   240
            Left            =   7845
            Picture         =   "FrmConsVenta1.frx":000C
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   315
            Width           =   225
         End
         Begin VB.CheckBox ChkMostrarItem 
            Caption         =   "Mostrar item"
            Height          =   195
            Left            =   7455
            TabIndex        =   13
            Top             =   690
            Width           =   1275
         End
         Begin VB.CheckBox chkAnioPasados 
            Caption         =   "Considerar Años Anteriores"
            Height          =   195
            Left            =   7455
            TabIndex        =   12
            Top             =   960
            Width           =   2595
         End
         Begin VB.TextBox TxtIdTipProd 
            Height          =   300
            Left            =   7455
            MaxLength       =   5
            TabIndex        =   29
            Text            =   "TxtIdTipProd"
            Top             =   285
            Width           =   645
         End
         Begin VB.Label lblTipProducto 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipProducto"
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
            Left            =   8100
            TabIndex        =   31
            Top             =   285
            Width           =   3330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Producto"
            Height          =   195
            Index           =   0
            Left            =   7455
            TabIndex        =   30
            Top             =   75
            Width           =   1230
         End
      End
      Begin VB.Frame Frame11 
         BorderStyle     =   0  'None
         Height          =   1215
         Left            =   12825
         TabIndex        =   8
         Top             =   45
         Width           =   11490
         Begin VB.TextBox TxtFiltroItem 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1800
            TabIndex        =   43
            Text            =   "TxtFiltroItem"
            Top             =   840
            Width           =   3915
         End
         Begin VB.CheckBox ChkFiltroItem 
            Caption         =   "Filtro Item Masivo"
            Height          =   195
            Left            =   90
            TabIndex        =   42
            Top             =   870
            Width           =   1605
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg3 
            Height          =   790
            Left            =   60
            TabIndex        =   10
            Top             =   0
            Width           =   5670
            _cx             =   10001
            _cy             =   1402
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
            FormatString    =   $"FrmConsVenta1.frx":013E
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
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   790
            Left            =   5775
            TabIndex        =   11
            Top             =   0
            Width           =   5670
            _cx             =   10001
            _cy             =   1393
            _ConvInfo       =   1
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   0   'False
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
            FormatString    =   $"FrmConsVenta1.frx":01A2
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
   End
   Begin VB.Frame FraProgreso 
      BorderStyle     =   0  'None
      Height          =   705
      Left            =   2730
      TabIndex        =   2
      Top             =   3630
      Visible         =   0   'False
      Width           =   5760
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   90
         TabIndex        =   3
         Top             =   345
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   450
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
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
         Left            =   1185
         TabIndex        =   6
         Top             =   75
         Width           =   585
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
         Left            =   90
         TabIndex        =   5
         Top             =   75
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
         Height          =   255
         Index           =   2
         Left            =   4140
         TabIndex        =   4
         Top             =   75
         Width           =   1530
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         X1              =   5745
         X2              =   5745
         Y1              =   -90
         Y2              =   4800
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         X1              =   -150
         X2              =   5895
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00000000&
         BorderWidth     =   2
         Index           =   2
         X1              =   -60
         X2              =   6360
         Y1              =   690
         Y2              =   690
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   3
         Index           =   3
         X1              =   0
         X2              =   15
         Y1              =   15
         Y2              =   5070
      End
   End
   Begin VSFlex7Ctl.VSFlexGrid Fg1 
      Align           =   2  'Align Bottom
      Height          =   6345
      Left            =   0
      TabIndex        =   0
      Top             =   1650
      Width           =   11895
      _cx             =   20981
      _cy             =   11192
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
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   28
      FixedRows       =   2
      FixedCols       =   1
      RowHeightMin    =   0
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"FrmConsVenta1.frx":021D
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   0
      MergeCompare    =   0
      AutoResize      =   -1  'True
      AutoSizeMode    =   1
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
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   1058
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Exportar con Formato"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Exportar sin Formato"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            NumListImages   =   14
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta1.frx":04C6
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta1.frx":0A0A
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta1.frx":0D9C
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta1.frx":0F20
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta1.frx":1374
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta1.frx":148C
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta1.frx":19D0
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta1.frx":1F14
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta1.frx":2028
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta1.frx":213C
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta1.frx":2590
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta1.frx":26FC
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta1.frx":2C44
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmConsVenta1.frx":2F5E
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Menu mnItem 
      Caption         =   "Item"
      Visible         =   0   'False
      Begin VB.Menu mnItemAdd 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mnItemSel 
         Caption         =   "Seleccionar"
      End
   End
   Begin VB.Menu mnCliente 
      Caption         =   "Cliente"
      Visible         =   0   'False
      Begin VB.Menu mnCliAdd 
         Caption         =   "Agregar"
      End
      Begin VB.Menu mnCliSel 
         Caption         =   "Seleccionar"
      End
   End
End
Attribute VB_Name = "FrmConsVenta1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--modificado 04/08/10 por johan Castro
'                      Agregar columnas en consulta del detalle[glosa,Referencia1(cuando doc sea nota credito),Referencia2(orden de despacho)]

Option Explicit

Dim vStrCons As String, vFormatString As String, vFormatStrGridItem As String, vFormatGridProv As String
Dim CaracteresNumericos As String

'-- ALMACENAR LOS TOTALES DE TODA LA CONSULTA
Dim Arr_Totales_grls() As Double
Dim Arr_Totales() As Double

Dim BAND_INTERRUMPIR As Boolean '--SE USARA PARA INTERRUMPIR LOS PROCESOS DE CONSULTA
                                '--TRUE SE INTERRUMPE
                                
Dim Q_POSICION_TOTAL  As Integer '--INDICA LA POCISION DE LA COLUMNA DONDE SE COLOCARA EL NOMBRE DEL TOTAL Y TOTAL_GRL
                                 '--OBTENDRA VALOR EN pGenerarConsulta()
                                
                                
                                
Dim T_RPT_PERIODO As String
Dim T_RPT_TITULO As String

Dim SeEjecuto As Boolean
    

Private Sub ChkFiltroItem_Click()
    If ChkFiltroItem.Value = 1 Then
        TxtFiltroItem.Enabled = True
    Else
        TxtFiltroItem.Enabled = False
    End If
End Sub

Private Sub CmdBusProducto_Click()
    On Error GoTo error
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "800":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT id, descripcion FROM mae_tipoproducto"
    
    xform.titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdTipProd.Text = xRs("id")
        lblTipProducto.Caption = xRs("descripcion")
        '--activar por defecto la seleccion de item's
        ChkMostrarItem.Value = 1
    End If

    Set xform = Nothing
    Set xRs = Nothing
    ChkMostrarItem.SetFocus
    Exit Sub
error:
    Set xform = Nothing
    Set xRs = Nothing
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"
    
End Sub

Private Sub ExportarExcelSinFormato()
    Dim mDataBase As New SistemaData.EDataBase
    Dim mRecord As New ADODB.Recordset
    Dim nSQL As String
    Dim F As New SistemaLogica.Funciones
    
    If Validar_Consulta() = False Then Exit Sub
    nSQL = pGenerarConsulta(True)
    If nSQL = "" Then Exit Sub
    
    
    Me.MousePointer = vbHourglass
    Set mDataBase.Connection = xCon
    mDataBase.CommandText = nSQL
    Set mRecord = mDataBase.GetRecordset
            
    If mRecord.RecordCount = 0 Then
        F.MostrarMensajeError "No se encontraron registros para la busqueda actual", "Costo de Movimientos"
        Me.MousePointer = vbDefault
        Exit Sub
    End If
    
    Me.MousePointer = vbDefault
    ' Se exporta a excel el recordset
    F.ExportarExcelRecordSet mRecord
End Sub


Private Sub pConsultar()
'    On Error GoTo error
    Dim rst_select As New ADODB.Recordset
       
    Dim nSQL As String '--RECIBIR LA CONSULTA
    If Validar_Consulta() = False Then Exit Sub
    nSQL = pGenerarConsulta()  '--DEVUELVE LA CONSULTA
    If nSQL = "" Then Exit Sub
        
'    nSQL = "SELECT registro,tdocabrev,nomcliente,numerodoc,fchdoc,fchven,conpagabre,diasvenc,glosa,simbolo,tipcam,desctipcom,codigo,descripcion,prodabrev,canpro,pumn,impdmn,pume,impdme,impdexpmn,impdexpme " _
'        + vbCr + "FROM ( " _
'        + vbCr + "SELECT Left([vta_ventas].[numreg],2) & Format(mae_libros.codsun,'00') & Right([vta_ventas].[numreg],4) AS registro, mae_cliente.numruc, mae_cliente.nombre AS nomcliente, mae_documento.abrev AS tdocabrev, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_condpago.abrev AS conpagabre, " _
'        + "IIf([vta_ventas].[impsal]<>0 And Date()>=[vta_ventas.fchven],Date()-[vta_ventas.fchven],'') AS diasvenc, mae_moneda.simbolo, IIf([vta_ventas].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, " _
'        + "vta_ventas.glosa, vta_ventas.idcli, vta_ventas.tipdoc, vta_ventas.idmon, " _
'        + "mae_tipoproducto.descripcion AS desctipcom, alm_inventario.codpro as codigo ,alm_inventario.descripcion, mae_unidades.abrev AS prodabrev, IIf([vta_ventas].[tipdoc]=7,(-1) * vta_ventasdet.canpro,vta_ventasdet.canpro) as canpro, " _
'        + "IIf([vta_ventas].[tipdoc]=7,(-1)*[vta_ventasdet].[imptot],[vta_ventasdet].[imptot]) AS impdreal, " _
'        + "IIf([vta_ventas].[idmon]=2,[vta_ventasdet].[preuni],0) AS pume, " _
'        + "IIf([vta_ventas].[idmon]=2,[impdreal],0) AS impdme, " _
'        + "IIf([vta_ventas].[idmon]=1,[vta_ventasdet].[preuni],0) AS pumn, " _
'        + "IIf([vta_ventas].[idmon]=1,[impdreal],0) AS impdmn, " _
'        + "IIf([vta_ventas].[idmon]=1,[impdmn],[impdreal]*[tipcam]) AS impdexpmn, " _
'        + "IIf([vta_ventas].[idmon] = 2, [impdme], IIf([tipcam] = 0, 0, [impdreal] / [tipcam])) As impdexpme " _
'        + vbCr + "FROM ((mae_cliente RIGHT JOIN (((((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_condpago ON vta_ventas.idconpag = mae_condpago.id) ON mae_cliente.id = vta_ventas.idcli) INNER JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN (mae_unidades RIGHT JOIN (mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed) ON vta_ventasdet.iditem = alm_inventario.id " _
'        + vbCr + "WHERE  vta_ventas.anulado = 0 AND  vta_ventas.fchdoc between cdate('01/01/2016') AND cdate('31/12/2016')  AND alm_inventario.tippro = 3   AND alm_inventario.descripcion LIKE '%DR.KU%'   and vta_ventas.numreg<>'000001' " _
'         + vbCr + ") AS vista " _
'        + vbCr + "ORDER BY vista.nomcliente,vista.fchdoc, vista.numerodoc "
 
    BAND_INTERRUMPIR = False
    LimpiarGrid Me.Fg1
    pConfigurarGrilla
    Me.MousePointer = vbHourglass
    DoEvents
    RST_Busq rst_select, nSQL, xCon
    PosicionarProgBar
    CARGAR_DATOS_GRILLA rst_select
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    Exit Sub
error:
    Me.MousePointer = vbDefault
    FraProgreso.Visible = False
    Set rst_select = Nothing
    MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"
End Sub

Private Function CARGAR_DATOS_GRILLA(RST_ORIGEN As ADODB.Recordset) As ADODB.Recordset
    '--FUNCION QUE AGREGARA LOS REGISTROS A LA GRILLA
    Dim vStrCampo As String
    Dim vCampos As Long
    Dim BAND_ADD_REG As Boolean
    Dim i As Integer
    
    BAND_ADD_REG = True
    
    vCampos = RST_ORIGEN.Fields.Count
    '--Libera la memoria usada por la matriz.
    Erase Arr_Totales
    Erase Arr_Totales_grls
    
    '--ARRAY QUE ACUMULARA LOS TOTALES
    ReDim Arr_Totales(10, 0)
    ReDim Arr_Totales_grls(10, 0)
    
    If RST_ORIGEN.RecordCount > 0 Then
        RST_ORIGEN.MoveFirst
    Else
        Exit Function
    End If
    
    PgBar.Min = 0
    PgBar.Max = RST_ORIGEN.RecordCount
    While Not RST_ORIGEN.EOF
    
    DoEvents
        '--SI SE NTERRUMPE EL PROCESO
        If BAND_INTERRUMPIR = True Then Exit Function
        '------CREANDO LOS GRUPOS
        If ((Me.OptDetalle.Value = True) Or (Me.OptResum.Value = True And (Trim(Me.TxtIdTipProd.Text) <> "" Or Me.ChkMostrarItem.Value = 1))) And RST_ORIGEN.Bookmark = 1 Then
            ADD_REG Fg1
            UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 5, RST_ORIGEN.Fields("nomcliente") & "", flexAlignLeftCenter: FORMATO_CELDA Fg1, Fg1.Rows - 1, 1
        End If
    
        Comparar_Grupo RST_ORIGEN, BAND_ADD_REG
        ADD_REG Fg1
        '--ASIGNAR LOS DATOS AL RECORDSET TEMPORAL
        For i = 0 To vCampos - 1
            vStrCampo = RST_ORIGEN.Fields(i).Name
            '--OBS: SE VA LLENAR EL ARRAY "MONTOS DEL TOTAL" O "MONTOS DEL RESUMEN"
            Select Case LCase(vStrCampo)
                
                '***************************************************************************
                '--MONTOS DEL TOTAL
                Case "candoc":                      Arr_Totales(0, 0) = Arr_Totales(0, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Case "canpro", "totcan":            Arr_Totales(0, 0) = Arr_Totales(0, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                
                Case "totmn", "impmn":              Arr_Totales(1, 0) = Arr_Totales(1, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Case "totsalmn", "impsalmn":        Arr_Totales(2, 0) = Arr_Totales(2, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                
                Case "totme", "impme":              Arr_Totales(3, 0) = Arr_Totales(3, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Case "totsalme", "impsalme":        Arr_Totales(4, 0) = Arr_Totales(4, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                
                
                Case "totexpmn", "impexpmn":        Arr_Totales(5, 0) = Arr_Totales(5, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Case "totaboexpmn", "impaboexpmn":  Arr_Totales(6, 0) = Arr_Totales(6, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Case "totsalexpmn", "impsalexpmn":  Arr_Totales(7, 0) = Arr_Totales(7, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                
                Case "totexpme":        Arr_Totales(8, 0) = Arr_Totales(8, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Case "totaboexpme":     Arr_Totales(9, 0) = Arr_Totales(9, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Case "totsalexpme":     Arr_Totales(10, 0) = Arr_Totales(10, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                
                Case "impdmn":          Arr_Totales(1, 0) = Arr_Totales(1, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Case "impdme":          Arr_Totales(2, 0) = Arr_Totales(2, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Case "impdexpmn":       Arr_Totales(3, 0) = Arr_Totales(3, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Case "impdexpme":       Arr_Totales(4, 0) = Arr_Totales(4, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                                
                Case "pumn":            Arr_Totales(4, 0) = Arr_Totales(4, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Case "pumn":            Arr_Totales(5, 0) = Arr_Totales(5, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                
                
                Case "totdmn":          Arr_Totales(1, 0) = Arr_Totales(1, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Case "totdme":          Arr_Totales(2, 0) = Arr_Totales(2, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Case "totdexpmn":       Arr_Totales(3, 0) = Arr_Totales(3, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                Case "totdexpme":       Arr_Totales(4, 0) = Arr_Totales(4, 0) + NulosN(RST_ORIGEN.Fields(vStrCampo))
                
                '***************************************************************************
                
                
            End Select
            
            If Me.OptDetalle.Value = True And Me.ChkMostrarItem.Value = 1 And LCase(vStrCampo) = "total_pu_mn" Then
                '--PARA ACUMULAR LOS REGISTROS ENCONTRDOS POR CLIENTE Y A LA VEZ ACUMULAR LOS REGISTROS ENCONTRADO DE TODA LA CONSULTA
                '--NOS SERVIRA PARA CALCULAR EL PRE. PROM. POR CLIENTE Y PRE. PROM. GRAL
                Arr_Totales(6, 0) = Arr_Totales(6, 0) + 1
            End If
            '--
            Select Case LCase(vStrCampo)
                
                '********************************
                Case "totme", "totsalme", "totmn", "totsalmn", "totexpmn", "totaboexpmn", "totsalexpmn", "totexpme", "totaboexpme", "totsalexpme"
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
                    
                Case "impme", "impsalme", "impmn", "impsalmn", "impexpmn", "impaboexpmn", "impsalexpmn"
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
                    
                Case "impdme", "impdmn", "impdexpmn", "impdexpme"
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
                    
                Case "pumn", "pume"
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
                    
                Case "totdmn", "totdme", "totdexpmn", "totdexpme"
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_MONTO)
                    
                '********************************
                
                Case "canpro"
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_CANTIDAD)
                Case "impven"
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = Format(NulosN(RST_ORIGEN.Fields(vStrCampo)), FORMAT_IMPUESTO)
                    
                Case "fchdoc", "fchven", "ref1fchdoc", "ref2fchdoc"
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = Format(RST_ORIGEN.Fields(vStrCampo), FORMAT_DATE)
                Case Else
                    Fg1.TextMatrix(Fg1.Rows - 1, i + 1) = RST_ORIGEN.Fields(vStrCampo) & ""
            End Select
    
        Next
        RST_ORIGEN.MoveNext
        '--PONER TOTALES AL FINAL DE LA GRILLA
        If RST_ORIGEN.EOF Then
            If Me.OptDetalle.Value = True Or (Me.TxtIdTipProd.Text <> "" Or Me.ChkMostrarItem.Value = 1) Then
                '--mostrar los totales del ultimo registro
                CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Total:"
            End If
            If Verificar_Poner_Datos_Grls() = True Then
                If OptResum.Value = True Then
                    If Me.TxtIdTipProd.Text <> "" Or Me.ChkMostrarItem.Value = 1 Then
                        CARGAR_DATOS_GRILLA_ADD_TOTALES True, "Tot Gen:", True, True
                    Else
                        CARGAR_DATOS_GRILLA_ADD_TOTALES True, "Tot Gen:", False, False
                    End If
                Else
                    CARGAR_DATOS_GRILLA_ADD_TOTALES True, "Tot Gen:", True, True
                End If
            End If
            '--DEL PRECIO PROMEDIO
            If VERIFICAR_PONER_PRECIO_PROMEDIO() = True Then
                CARGAR_DATOS_GRILLA_ADD_TOTALES True, "P. Prom"
                If Verificar_Poner_Datos_Grls() = True Then CARGAR_DATOS_GRILLA_ADD_TOTALES True, "P. Prom. Gen", True, True
            End If
        Else
            PgBar.Value = CLng(RST_ORIGEN.Bookmark)
        End If
        
    Wend
End Function

Private Sub pImprimir()
    On Error GoTo error
    Dim X_PRINT As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    X_PRINT.Imprimir_x_VSFlexGrid Fg1, T_RPT_TITULO, " ", T_RPT_PERIODO, True, True
    Set X_PRINT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR
End Sub
Private Sub ChkMostrarItem_Click()
    If Me.ChkMostrarItem.Value = 0 Then
        Fg2.Enabled = False
    Else
        '--LIMPIAR GRILLA
        Fg2.Enabled = True
        OptTodos.Value = True
        LimpiarGrid Fg2, True, 2
        GRID_COMBOLIST Fg2
    End If

End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then pCargaItem False
End Sub

Private Sub Fg2_DblClick()
    Fg2_CellButtonClick Fg2.Rows - 1, 2
End Sub

Private Sub Fg2_KeyDown(KeyCode As Integer, Shift As Integer)
    If Fg2.Row = -2 Then Exit Sub
    Select Case KeyCode
        Case 45  'INSERTAR REGI
            Fg2.AddItem ""
            Fg2.Row = Fg2.Rows - 1: Fg2.Col = 2
        Case 46 'SUPRIMIR/DELETE
            If Fg2.Rows - 1 >= 2 Then
                Fg2.RemoveItem Fg2.Row
                Fg2.Row = Fg2.Rows - 1: Fg2.Col = 2
            Else
                LimpiarGrid Fg2, True, 2
                GRID_COMBOLIST Fg2
            End If
    End Select
End Sub

Private Sub Fg2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If validar_letras(KeyAscii) = False Then
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub Fg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnItem
End Sub

Private Sub Fg3_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 2 Then pCargarCliente False
End Sub

Private Sub Fg3_DblClick()
    Fg3_CellButtonClick Fg3.Rows - 1, 2
End Sub

Private Sub Fg3_KeyDown(KeyCode As Integer, Shift As Integer)
    If Fg3.Row = -2 Then Exit Sub
    Select Case KeyCode
        Case 45  'INSERTAR REGI
            Fg3.AddItem ""
            Fg3.Row = Fg3.Rows - 1: Fg3.Col = 2
        Case 46
            If Fg3.Rows - 1 >= 2 Then
                Fg3.RemoveItem Fg3.Row
                Fg3.Row = Fg3.Rows - 1: Fg3.Col = 2
            Else
                LimpiarGrid Fg3, True, 2
                GRID_COMBOLIST Fg3
            End If
    End Select
End Sub

Private Sub Fg3_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then Exit Sub
    If validar_letras(KeyAscii) = False Then
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub Fg3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnCliente
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        TabOne2.CurrTab = 0
        SeEjecuto = True
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape And Shift = 0 Then
        '--interrumpir
        BAND_INTERRUMPIR = True
    End If
End Sub

Private Sub Form_Load()
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    CentrarFrm Me
    SeEjecuto = False
    GRID_COMBOLIST Fg2
    GRID_COMBOLIST Fg3
    
    vFormatString = Fg1.FormatString
    Fg2.Tag = Fg2.FormatString
    Fg3.Tag = Fg3.FormatString
 
    TxtIdTipProd.Text = ""
    lblTipProducto.Caption = ""
    CaracteresNumericos = "0123456789." & Chr(8)
    TxtFiltroItem.Text = ""
    
    TxtFchIni.Valor = CDate("01/01/" + CStr(AnoTra))
    TxtFchFin.Valor = CDate("31/12/" + CStr(AnoTra))
    
    LimpiarGrid Me.Fg1
    pConfigurarGrilla
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub

    If Me.Height > 2100 Then
        Fg1.Top = 1650
        'Fg1.Width = Me.Width - 150
        Fg1.Height = Me.Height - 2220
    End If
End Sub

Private Sub mnCliAdd_Click()
    pCargarCliente False
End Sub

Private Sub mnCliSel_Click()
    pCargarCliente True
End Sub

Private Sub mnItemAdd_Click()
    pCargaItem False
End Sub

Private Sub mnItemSel_Click()
    pCargaItem True
End Sub

Private Sub OptDetalle_Click()
    habilitar opt_orden, True
End Sub

Private Sub OptResum_Click()
    habilitar opt_orden, False
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 5 Then
        'Exportar con Formato
        If ButtonMenu.Index = 1 Then
            pExportarExcel
        End If
        'Exportar sin Formato
        If ButtonMenu.Index = 2 Then
            ExportarExcelSinFormato
        End If
    End If
End Sub

Private Sub TxtIdTipProd_Change()
    If TxtIdTipProd.Text = "" Then
        lblTipProducto.Caption = ""
        If Me.ChkMostrarItem.Value = 1 Then ChkMostrarItem.Value = 0
        LimpiarGrid Fg2, True
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
                '--activar por defecto la seleccion de item's
                ChkMostrarItem.Value = 1
            Else
                lblTipProducto.Caption = ""
                TxtIdTipProd.Text = ""
                ChkMostrarItem.Value = 0
            End If
        End If
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    Set RsTipProd = Nothing
    Exit Sub
error:
    Set RsTipProd = Nothing
    SHOW_ERROR

End Sub

Private Sub TxtIdTipProd_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then  'TECHAL F5
        CmdBusProducto.Value = True
    End If
End Sub

'------
Private Function Validar_Consulta() As Boolean
    '--FUNCION QUE VALIDARA LA CONSULTA
    '--DE LA FECHA ES NULL
    
    '--posicionar en la primera pestaña del menu de opciones de consulta
    TabOne2.CurrTab = 0
    '---
            
    If TxtFchIni.Valor = "" Or TxtFchFin.Valor = "" Then
        MsgBox "Ingrese una fecha", vbExclamation, xTitulo
        If TxtFchIni.Valor = "" Then TxtFchIni.SetFocus Else TxtFchFin.SetFocus
        Exit Function
    End If
    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        MsgBox "La fecha inicial es superior al Final", vbExclamation, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    End If
        If (Year(TxtFchIni.Valor) <> Year(TxtFchFin.Valor)) Then
        MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    ElseIf Year(TxtFchIni.Valor) <> CStr(AnoTra) Then
        MsgBox "El rango de fechas debe estar en el Año de Trabajo : " + CStr(AnoTra), vbExclamation, xTitulo
        TxtFchIni.SetFocus
        Exit Function
    End If
    Validar_Consulta = True

End Function

Private Function pGenerarConsulta(Optional MostrarCabeceraDetallada As Boolean = False) As String
    '--FUNCION QUE NOS PERMITIRA GENERAR LA CONSULTA DE ACUERDO A LO QUE SELECCIONE EL USUARIO
    '--
    Dim nSQL As String        '--CONSULTA GENERAL, ESTO PERMITIRA HACER LA CONSULTA
    Dim nSQLItem As String   '--SOLO ITEM
    Dim nSQLFiltro_CLI As String    '--SOLO CLIENTES
    Dim nSQLTipoItem As String
    Dim nSQLFecha As String
    Dim nSQLFiltro As String
    Dim vFiltro As String
    Dim k  As Integer
    
    '--DE LA FECHA
    If CDate(TxtFchIni.Valor) < CDate(TxtFchFin.Valor) Then
        'nSQLFecha = " vta_ventas.fchdoc >= cdate('" + Format(TxtFchIni.Valor, "dd/mm/yyyy") + "') AND vta_ventas.fchdoc<= cdate('" + Format(TxtFchFin.Valor, "dd/mm/yyyy") + "') "
        nSQLFecha = " vta_ventas.fchdoc between cdate('" + Format(TxtFchIni.Valor, "dd/mm/yyyy") + "') AND cdate('" + Format(TxtFchFin.Valor, "dd/mm/yyyy") + "') "
        T_RPT_PERIODO = " Del: " + TxtFchIni.Valor + " Al: " + TxtFchFin.Valor
    Else
        nSQLFecha = " vta_ventas.fchdoc = cdate('" + Format(TxtFchIni.Valor, "dd/mm/yyyy") + "') "
        T_RPT_PERIODO = "Al: " + TxtFchFin.Valor
    End If
    
    '--SI OPCION DE SELECCIONAR POR FECHA DE VENCIMIENTO
    If Me.OptVenc.Value = True Then nSQLFecha = Replace(nSQLFecha, "vta_ventas.fchdoc", "vta_ventas.fchven")
    '--SI OPCION SELECCIONA POR FECHA DE REGISTRO
    If Me.OptReg.Value = True Then nSQLFecha = Replace(nSQLFecha, "vta_ventas.fchdoc", "vta_ventas.fchreg")
    
    '--DEL TIPO DE PRODUCTO
    If TxtIdTipProd.Text <> "" Then vFiltro = vFiltro + " AND alm_inventario.tippro = " + CStr(TxtIdTipProd.Text) + " "
    
    '--DEL ITEM
    vFiltro = vFiltro & GRID_GENERAR_SQL_ID(Fg2, 3, " AND alm_inventario.id", "IN")
    
    '--DEL CLIENTE
    vFiltro = vFiltro & GRID_GENERAR_SQL_ID(Fg3, 1, " AND vta_ventas.idcli", "IN")
 
    '--DE LA MONEDA
    If OptSol.Value = True Then vFiltro = vFiltro + " AND vta_ventas.idmon= 1 "       '--SOLES
    If Me.OptDol.Value = True Then vFiltro = vFiltro + " AND vta_ventas.idmon= 2 "    '--DOLARES
    '---------------
    
    If OptPag.Value = True Then         '---SI ES CANCELADO
        vFiltro = vFiltro + " AND vta_ventas.impsal = 0 "
        
    ElseIf OptPend.Value = True Then    '---SI ES PENDIENTE DE PAGO
        vFiltro = vFiltro + " AND vta_ventas.impsal > 0 "
        
    End If
        
    nSQLFiltro = " vta_ventas.anulado = 0 AND " + nSQLFecha + vFiltro
    
    If ChkMostrarItem.Value = 1 Then
        nSQLFiltro = nSQLFiltro & " AND alm_inventario.descripcion LIKE '%" & TxtFiltroItem.Text & "%' "
    End If
    
    '--incluir documentos de apertura
    If chkAnioPasados.Value = 0 Then nSQLFiltro = nSQLFiltro & " and vta_ventas.numreg<>'000001'"
        
    '------------------------------------------------------------------------------------
    
    If OptResum.Value = True Then '--RESUMEN
        If ChkMostrarItem.Value = 1 Or TxtIdTipProd.Text <> "" Then
            If TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0 Then
            '--MOSTRAR SOLO PRODUCTO
   
    
        T_RPT_TITULO = "REPORTE DE VENTAS RESUMIDO POR TIPO PRODUCTO"

        
        nSQL = "SELECT vista.numruc, vista.nomcliente, vista.desctipcom, Sum(vista.impdmn) AS totdmn, Sum(vista.impdme) AS totdme, Sum(vista.impdexpmn) AS totdexpmn, Sum(vista.impdexpme) AS totdexpme " _
            + vbCr + " FROM ( " _
            + vbCr + " SELECT Left([vta_ventas].[numreg],2) & Format(mae_libros.codsun,'00') & Right([vta_ventas].[numreg],4) AS registro, mae_cliente.numruc, mae_cliente.nombre AS nomcliente, mae_documento.abrev AS tdocabrev, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_condpago.abrev AS conpagabre, " _
            + vbCr + " IIf([vta_ventas].[impsal]<>0 And Date()>=[vta_ventas.fchven],Date()-[vta_ventas.fchven],'') AS diasvenc, mae_moneda.simbolo, IIf([vta_ventas].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, " _
            + vbCr + " vta_ventas.idcli, vta_ventas.tipdoc, vta_ventas.idmon, " _
            + vbCr + " mae_tipoproducto.descripcion AS desctipcom, alm_inventario.codpro as codigo ,alm_inventario.descripcion, mae_unidades.abrev AS prodabrev, vta_ventasdet.canpro, " _
            + vbCr + " IIf([vta_ventas].[tipdoc]=7,(-1)*[vta_ventasdet].[imptot],[vta_ventasdet].[imptot]) AS impdreal, " _
            + vbCr + " IIf([vta_ventas].[idmon]=2,[vta_ventasdet].[preuni],0) AS pume, " _
            + vbCr + " IIf([vta_ventas].[idmon]=2,[impdreal],0) AS impdme, " _
            + vbCr + " IIf([vta_ventas].[idmon]=1,[vta_ventasdet].[preuni],0) AS pumn, " _
            + vbCr + " IIf([vta_ventas].[idmon]=1,[impdreal],0) AS impdmn, " _
            + vbCr + " IIf([vta_ventas].[idmon]=1,[impdmn],[impdreal]*[tipcam]) AS impdexpmn, " _
            + vbCr + " IIf([vta_ventas].[idmon] = 2, [impdme], IIf([tipcam] = 0, 0, [impdreal] / [tipcam])) As impdexpme " _
            + vbCr + " FROM ((mae_cliente RIGHT JOIN (((((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_condpago ON vta_ventas.idconpag = mae_condpago.id) ON mae_cliente.id = vta_ventas.idcli) INNER JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN (mae_unidades RIGHT JOIN (mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed) ON vta_ventasdet.iditem = alm_inventario.id " _
            + vbCr + " WHERE " & nSQLFiltro _
            + vbCr + "  ) AS vista " _
            + vbCr + " GROUP BY vista.numruc, vista.nomcliente, vista.desctipcom" _
            + vbCr + " ORDER BY vista.nomcliente,vista.desctipcom "
                                
        If MostrarCabeceraDetallada Then
            nSQL = "SELECT c.numruc AS [Numero Ruc], c.nomcliente AS [Cliente], c.desctipcom AS [Tipo Producto], c.totmn AS [Total MN], c.totdme AS [Total ME], c.totdexpmn AS [Expresado MN], c.impdexpme AS [Saldo ME],  c.totexpmn AS [Expresado ME] " _
            + vbCr + " FROM ( " _
            + vbCr + nSQL _
            + vbCr + " ) AS c"
        End If
        
        Q_POSICION_TOTAL = 3
        
                       
            Else
            '--MOSTRAR PRODUCTO Y ITEM
            
                T_RPT_TITULO = "REPORTE DE VENTAS RESUMIDO POR TIPO PRODUCTO CON ITEM"
            
                '--en esta consulta solo se considera la base imponible
                nSQL = "SELECT vista.numruc, vista.nomcliente, vista.desctipcom, vista.codigo, vista.descripcion,vista.prodabrev, Sum(vista.canpro) AS totcan, Sum(vista.impdmn) AS totdmn, Sum(vista.impdme) AS totdme, Sum(vista.impdexpmn) AS totdexpmn, Sum(vista.impdexpme) AS totdexpme " _
                    + vbCr + " FROM ( " _
                    + vbCr + " SELECT Left([vta_ventas].[numreg],2) & Format(mae_libros.codsun,'00') & Right([vta_ventas].[numreg],4) AS registro, mae_cliente.numruc, mae_cliente.nombre AS nomcliente, mae_documento.abrev AS tdocabrev, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_condpago.abrev AS conpagabre, " _
                    + vbCr + " IIf([vta_ventas].[impsal]<>0 And Date()>=[vta_ventas.fchven],Date()-[vta_ventas.fchven],'') AS diasvenc, mae_moneda.simbolo, IIf([vta_ventas].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, " _
                    + vbCr + " vta_ventas.idcli, vta_ventas.tipdoc, vta_ventas.idmon, " _
                    + vbCr + " mae_tipoproducto.descripcion AS desctipcom, alm_inventario.codpro as codigo ,alm_inventario.descripcion, mae_unidades.abrev AS prodabrev, IIf([vta_ventas].[tipdoc]=7,(-1) * vta_ventasdet.canpro,vta_ventasdet.canpro) as canpro,  " _
                    + vbCr + " IIf([vta_ventas].[tipdoc]=7,(-1)*[vta_ventasdet].[imptot],[vta_ventasdet].[imptot]) AS impdreal, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=2,[vta_ventasdet].[preuni],0) AS pume, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=2,[impdreal],0) AS impdme, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=1,[vta_ventasdet].[preuni],0) AS pumn, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=1,[impdreal],0) AS impdmn, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=1,[impdmn],[impdreal]*[tipcam]) AS impdexpmn, " _
                    + vbCr + " IIf([vta_ventas].[idmon] = 2, [impdme], IIf([tipcam] = 0, 0, [impdreal] / [tipcam])) As impdexpme " _
                    + vbCr + " FROM ((mae_cliente RIGHT JOIN (((((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_condpago ON vta_ventas.idconpag = mae_condpago.id) ON mae_cliente.id = vta_ventas.idcli) INNER JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN (mae_unidades RIGHT JOIN (mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed) ON vta_ventasdet.iditem = alm_inventario.id " _
                    + vbCr + " WHERE " & nSQLFiltro _
                    + vbCr + "  ) AS vista " _
                    + vbCr + " GROUP BY vista.numruc, vista.nomcliente, vista.codigo, vista.desctipcom, vista.descripcion, vista.prodabrev " _
                    + vbCr + " ORDER BY vista.nomcliente,vista.desctipcom,vista.descripcion "
                                
                If MostrarCabeceraDetallada Then
                    nSQL = "SELECT c.numruc AS [Numero Ruc], c.nomcliente AS [Cliente], c.desctipcom AS [Tipo Producto], c.codigo AS [Codigo Prod], c.descripcion AS [Descripcion Prod], c.prodabrev AS [Unidad], c.totcan AS [Cantidad], c.totdmn AS [Total MN], c.totdme AS [Total ME], c.totdexpmn AS [Expresado MN], c.totdexpme AS [Expresado ME] " _
                        + vbCr + " FROM ( " _
                        + vbCr + nSQL _
                        + vbCr + " ) AS c"
                End If
            
                Q_POSICION_TOTAL = 5
                
            End If
        Else '--GENERAL
                
                T_RPT_TITULO = "REPORTE DE VENTAS RESUMIDO POR CLIENTE"
                                
                nSQL = "SELECT vista.numruc, vista.nomcliente , Count(vista.numruc) AS candoc, Sum(vista.impmn) AS totmn, Sum(vista.impsalmn) AS totsalmn, Sum(vista.impme) AS totme, Sum(vista.impsalme) AS totsalme,  Sum(vista.impexpmn) AS totexpmn, totexpmn-totsalexpmn AS totaboexpmn , Sum(vista.impsalexpmn) AS totsalexpmn,Sum(vista.impexpme) AS totexpme, totexpme-totsalexpme AS totaboexpme , Sum(vista.impsalexpme) AS totsalexpme " _
                    + vbCr + "  FROM ( " _
                        + vbCr + " SELECT Left([vta_ventas].[numreg],2) & Format(mae_libros.codsun,'00') & Right([vta_ventas].[numreg],4) AS registro, mae_cliente.numruc, mae_cliente.nombre as nomcliente, mae_documento.abrev AS tdocabrev, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_condpago.descripcion AS conpagabre, " _
                        + vbCr + " IIf([vta_ventas].[impsal]<>0 And Date()>=[vta_ventas.fchven],Date()-[vta_ventas.fchven],'') AS diasvenc, mae_moneda.simbolo, " _
                        + vbCr + " IIf([vta_ventas].[tc]=0,IIf([con_tc].[impven] is null,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, " _
                        + vbCr + " IIf(vta_ventas.tipdoc=7,(-1)*vta_ventas.imptotdoc,vta_ventas.imptotdoc) AS impreal, " _
                        + vbCr + " IIf(vta_ventas.idmon=1,impreal,0) AS impmn, " _
                        + vbCr + " IIf(vta_ventas.idmon=2,impreal,0) AS impme, " _
                        + vbCr + " IIf(vta_ventas.tipdoc=7,(-1)*vta_ventas.impsal,vta_ventas.impsal) AS impsalreal, " _
                        + vbCr + " IIf(vta_ventas.idmon=1,impsalreal,0) AS impsalmn, " _
                        + vbCr + " IIf(vta_ventas.idmon=2,impsalreal,0) AS impsalme, " _
                        + vbCr + " (impmn + impme * tipcam) as impexpmn, " _
                        + vbCr + " (impme +  iif(tipcam=0,0, impmn / tipcam) )  as impexpme, " _
                        + vbCr + " (impsalmn + impsalme * tipcam) as impsalexpmn, " _
                        + vbCr + " (impsalme +  iif(tipcam=0,0, impsalmn / tipcam) )  as impsalexpme, " _
                        + vbCr + " vta_ventas.idcli , vta_ventas.TipDoc, vta_ventas.idmon " _
                        + vbCr + " FROM mae_cliente RIGHT JOIN (((((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_condpago ON vta_ventas.idconpag = mae_condpago.id) ON mae_cliente.id = vta_ventas.idcli " _
                        + vbCr + " WHERE  " & nSQLFiltro _
                    + vbCr + "  ) AS vista " _
                    + vbCr + " GROUP BY vista.numruc, vista.nomcliente " _
                    + vbCr + " ORDER BY vista.nomcliente "

                If MostrarCabeceraDetallada Then
                    nSQL = "SELECT c.numruc AS [Numero Ruc], c.nomcliente AS [Cliente], c.candoc AS [Total Doc], c.totmn AS [Total MN], c.totsalmn AS [Saldo MN], c.totme AS [Total ME], c.totsalme AS [Saldo ME],  c.totexpmn AS [Expresado MN], c.totaboexpmn AS [Abono Expresado MN], c.totsalexpmn AS [Saldo Expresado MN], c.totexpme AS [Expresado ME], c.totaboexpme AS [Abono Expresado ME], c.totsalexpme AS [Saldo Expresado ME] " _
                    + vbCr + " FROM ( " _
                    + vbCr + nSQL _
                    + vbCr + " ) AS c"
                End If
                
            Q_POSICION_TOTAL = 2
        End If
    
    
    Else '--DETALLADO
        If ChkMostrarItem.Value = 1 Or TxtIdTipProd.Text <> "" Then
            If TxtIdTipProd.Text <> "" And ChkMostrarItem.Value = 0 Then '--MOSTRAR SOLO PRODUCTO
                
                T_RPT_TITULO = "REPORTE DETALLADO POR TIPO PRODUCTO"
                
                '--en esta consulta solo se considera la base imponible
                nSQL = "SELECT registro, tdocabrev, nomcliente, numerodoc, fchdoc, fchven, conpagabre, diasvenc,glosa, simbolo, tipcam, desctipcom, Sum(impdmn) AS totdmn, Sum(impdme) AS totdme, Sum(impdexpmn) AS totdexpmn, Sum(impdexpme) AS totdexpme " _
                    + vbCr + " FROM ( " _
                    + vbCr + " SELECT Left([vta_ventas].[numreg],2) & Format(mae_libros.codsun,'00') & Right([vta_ventas].[numreg],4) AS registro, mae_cliente.numruc, mae_cliente.nombre AS nomcliente, mae_documento.abrev AS tdocabrev, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_condpago.abrev AS conpagabre, " _
                    + vbCr + " IIf([vta_ventas].[impsal]<>0 And Date()>=[vta_ventas.fchven],Date()-[vta_ventas.fchven],'') AS diasvenc, mae_moneda.simbolo, IIf([vta_ventas].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, " _
                    + vbCr + " vta_ventas.glosa, vta_ventas.idcli, vta_ventas.tipdoc, vta_ventas.idmon, " _
                    + vbCr + " mae_tipoproducto.descripcion AS desctipcom, alm_inventario.codpro as codigo ,alm_inventario.descripcion, mae_unidades.abrev AS prodabrev, IIf([vta_ventas].[tipdoc]=7,(-1) * vta_ventasdet.canpro,vta_ventasdet.canpro) as canpro, " _
                    + vbCr + " IIf([vta_ventas].[tipdoc]=7,(-1)*[vta_ventasdet].[imptot],[vta_ventasdet].[imptot]) AS impdreal, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=2,[vta_ventasdet].[preuni],0) AS pume, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=2,[impdreal],0) AS impdme, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=1,[vta_ventasdet].[preuni],0) AS pumn, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=1,[impdreal],0) AS impdmn, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=1,[impdmn],[impdreal]*[tipcam]) AS impdexpmn, " _
                    + vbCr + " IIf([vta_ventas].[idmon] = 2, [impdme], IIf([tipcam] = 0, 0, [impdreal] / [tipcam])) As impdexpme " _
                    + vbCr + " FROM ((mae_cliente RIGHT JOIN (((((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_condpago ON vta_ventas.idconpag = mae_condpago.id) ON mae_cliente.id = vta_ventas.idcli) INNER JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN (mae_unidades RIGHT JOIN (mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed) ON vta_ventasdet.iditem = alm_inventario.id " _
                    + vbCr + " WHERE " & nSQLFiltro _
                    + vbCr + "  ) AS vista " _
                    + vbCr + " GROUP BY vista.registro, vista.tdocabrev, vista.nomcliente, vista.numerodoc, vista.fchdoc, vista.fchven, vista.conpagabre, vista.diasvenc, vista.simbolo, vista.tipcam, vista.desctipcom " _
                    + vbCr + " ORDER BY vista.nomcliente,vista.fchdoc, vista.numerodoc "
                                  
                If MostrarCabeceraDetallada Then
                    nSQL = "SELECT c.registro AS [N° Registro], c.tdocabrev AS [Tipo Documento], c.nomcliente AS [Cliente], c.numerodoc AS [Numero Documento], c.fchdoc AS [Fecha Documento], c.fchven AS [Fecha Vencimiento], conpagabre, diasvenc, c.glosa As [Glosa], c.simbolo AS [Moneda], c.tipcam AS [Tipo de Cambio], c.desctipcom AS [Tipo Producto], c.totdmn AS [Total MN], c.totdme AS [Total ME], c.totdexpmn AS [Expresado MN], c.totdexpme AS [Expresado ME] " _
                        + vbCr + " FROM ( " _
                        + vbCr + nSQL _
                        + vbCr + " ) AS c"
                End If
                Q_POSICION_TOTAL = 6
            Else
            '--MOSTRAR PRODUCTO Y ITEM
                T_RPT_TITULO = "REPORTE DE VENTAS DETALLADO POR TIPO PRODUCTO CON ITEM"
                
                    '--en esta consulta solo se considera la base imponible
                    nSQL = "SELECT registro,tdocabrev,nomcliente,numerodoc,fchdoc,fchven,conpagabre,diasvenc,glosa,simbolo,tipcam,desctipcom,codigo,descripcion,prodabrev,canpro,pumn,impdmn,pume,impdme,impdexpmn,impdexpme " _
                    + vbCr + " FROM ( " _
                    + vbCr + " SELECT Left([vta_ventas].[numreg],2) & Format(mae_libros.codsun,'00') & Right([vta_ventas].[numreg],4) AS registro, mae_cliente.numruc, mae_cliente.nombre AS nomcliente, mae_documento.abrev AS tdocabrev, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_condpago.abrev AS conpagabre, " _
                    + vbCr + " IIf([vta_ventas].[impsal]<>0 And Date()>=[vta_ventas.fchven],Date()-[vta_ventas.fchven],'') AS diasvenc, mae_moneda.simbolo, IIf([vta_ventas].[tc]=0,IIf([con_tc].[impven] Is Null,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, " _
                    + vbCr + " vta_ventas.glosa, vta_ventas.idcli, vta_ventas.tipdoc, vta_ventas.idmon, " _
                    + vbCr + " mae_tipoproducto.descripcion AS desctipcom, alm_inventario.codpro as codigo ,alm_inventario.descripcion, mae_unidades.abrev AS prodabrev, IIf([vta_ventas].[tipdoc]=7,(-1) * vta_ventasdet.canpro,vta_ventasdet.canpro) as canpro, " _
                    + vbCr + " IIf([vta_ventas].[tipdoc]=7,(-1)*[vta_ventasdet].[imptot],[vta_ventasdet].[imptot]) AS impdreal, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=2,[vta_ventasdet].[preuni],0) AS pume, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=2,[impdreal],0) AS impdme, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=1,[vta_ventasdet].[preuni],0) AS pumn, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=1,[impdreal],0) AS impdmn, " _
                    + vbCr + " IIf([vta_ventas].[idmon]=1,[impdmn],[impdreal]*[tipcam]) AS impdexpmn, " _
                    + vbCr + " IIf([vta_ventas].[idmon] = 2, [impdme], IIf([tipcam] = 0, 0, [impdreal] / [tipcam])) As impdexpme " _
                    + vbCr + " FROM ((mae_cliente RIGHT JOIN (((((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_condpago ON vta_ventas.idconpag = mae_condpago.id) ON mae_cliente.id = vta_ventas.idcli) INNER JOIN vta_ventasdet ON vta_ventas.id = vta_ventasdet.idvta) LEFT JOIN (mae_unidades RIGHT JOIN (mae_tipoproducto RIGHT JOIN alm_inventario ON mae_tipoproducto.id = alm_inventario.tippro) ON mae_unidades.id = alm_inventario.idunimed) ON vta_ventasdet.iditem = alm_inventario.id " _
                    + vbCr + " WHERE " & nSQLFiltro _
                    + vbCr + "  ) AS vista " _
                    + vbCr + " ORDER BY vista.nomcliente,vista.fchdoc, vista.numerodoc "

                If MostrarCabeceraDetallada Then
                    nSQL = "SELECT c.registro AS [N° Registro], c.tdocabrev AS [Tipo Documento], c.nomcliente AS [Cliente], c.numerodoc AS [Numero Documento], c.fchdoc AS [Fecha Documento], c.fchven AS [Fecha Vencimiento], c.conpagabre AS [Condicion Pago], c.diasvenc As [Dias Retraso], c.glosa As [Glosa], c.simbolo AS [Moneda], c.tipcam AS [Tipo de Cambio], c.desctipcom AS [Tipo Producto], c.codigo AS [Codigo Producto], c.descripcion AS [Descripcion Producto], c.prodabrev AS [Unidad], c.canpro AS [Cantidad], c.pumn AS [Precio Unitario MN], c.impdmn AS [Total MN], c.pume AS [Precio Unitario ME], c.impdme AS [Total ME], c.impdexpmn AS [Expresado MN], c.impdexpme AS [Expresado ME] " _
                        + vbCr + " FROM ( " _
                        + vbCr + nSQL _
                        + vbCr + " ) AS c"
                End If
                           
                Q_POSICION_TOTAL = 6
            End If
        Else '--MOSTRAR SIN DETALLE
        
            T_RPT_TITULO = "REPORTE DE VENTAS DETALLADO POR CLIENTE"

            nSQL = "SELECT registro,tdocabrev,nomcliente,numerodoc,fchdoc,fchven,conpagabre,diasvenc,glosa,simbolo,tipcam,impmn,impsalmn,impme,impsalme,impexpmn, (impexpmn-impsalexpmn) as impaboexpmn ,impsalexpmn, ref1registro,ref1abrev,ref1numdoc,ref1fchdoc, ref2abrev,ref2numdoc,ref2fchdoc,ref2cliruc,ref2clinombre " _
                + vbCr + "  FROM ( " _
                    + vbCr + " SELECT Left([vta_ventas].[numreg],2) & Format(mae_libros.codsun,'00') & Right([vta_ventas].[numreg],4) AS registro, mae_cliente.numruc, mae_cliente.nombre as nomcliente, mae_documento.abrev AS tdocabrev, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, vta_ventas.fchdoc, vta_ventas.fchven, mae_condpago.abrev AS conpagabre, " _
                    + vbCr + " IIf([vta_ventas].[impsal]<>0 And Date()>=[vta_ventas.fchven],Date()-[vta_ventas.fchven],'') AS diasvenc, mae_moneda.simbolo, " _
                    + vbCr + " IIf([vta_ventas].[tc]=0,IIf([con_tc].[impven] is null,0,[con_tc].[impven]),[vta_ventas].[tc]) AS tipcam, " _
                    + vbCr + " IIf(vta_ventas.tipdoc=7,(-1)*vta_ventas.imptotdoc,vta_ventas.imptotdoc) AS impreal, " _
                    + vbCr + " IIf(vta_ventas.idmon=1,impreal,0) AS impmn, " _
                    + vbCr + " IIf(vta_ventas.idmon=2,impreal,0) AS impme, " _
                    + vbCr + " IIf(vta_ventas.tipdoc=7,(-1)*vta_ventas.impsal,vta_ventas.impsal) AS impsalreal, " _
                    + vbCr + " IIf(vta_ventas.idmon=1,impsalreal,0) AS impsalmn, " _
                    + vbCr + " IIf(vta_ventas.idmon=2,impsalreal,0) AS impsalme, " _
                    + vbCr + " (impmn + impme * tipcam) as impexpmn, " _
                    + vbCr + " (impme +  iif(tipcam=0,0, impmn / tipcam) )  as impexpme, " _
                    + vbCr + " (impsalmn + impsalme * tipcam) as impsalexpmn, " _
                    + vbCr + " (impsalme +  iif(tipcam=0,0, impsalmn / tipcam) )  as impsalexpme, " _
                    + vbCr + " vta_ventas.glosa, vta_ventas.idcli, vta_ventas.tipdoc, vta_ventas.idmon, " _
                    + vbCr + " Left([vta_ventas_1].[numreg],2) & Format([mae_libros_1].[codsun],'00') & Right([vta_ventas_1].[numreg],4) AS ref1registro, mae_documento_1.abrev AS ref1abrev, IIf([vta_ventas].[iddocref]=0,'',[vta_ventas_1].[numser] & '-' & [vta_ventas_1].[numdoc]) AS ref1numdoc, vta_ventas_1.fchdoc AS ref1fchdoc,  " _
                    + vbCr + " IIf(vta_ventas.idtipdocref=4,mae_cliente_1.numruc,'') AS ref2cliruc,  IIf(vta_ventas.idtipdocref=4,mae_cliente_1.nombre,'') AS ref2clinombre,IIF(vta_ventas.idtipdocref=4,mae_documento_2.abrev,'') AS ref2abrev, IIF(vta_ventas.idtipdocref=4,var_ordendespacho.numerodoc, vta_ventas.numerodocref) AS ref2numdoc, IIF(vta_ventas.idtipdocref=4,var_ordendespacho.fchemi,'') AS ref2fchdoc " _
                    + vbCr + " FROM ((((((mae_cliente RIGHT JOIN (((((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id) LEFT JOIN mae_condpago ON vta_ventas.idconpag = mae_condpago.id) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN (vta_ventas AS vta_ventas_1 LEFT JOIN mae_documento AS mae_documento_1 ON vta_ventas_1.tipdoc = mae_documento_1.id) ON vta_ventas.iddocref = vta_ventas_1.id) " _
                    + vbCr + " LEFT JOIN mae_libros AS mae_libros_1 ON vta_ventas_1.idlib = mae_libros_1.id) LEFT JOIN var_ordendespacho ON vta_ventas.iddocref2 = var_ordendespacho.id) LEFT JOIN mae_docreferencia ON vta_ventas.idtipdocref = mae_docreferencia.id) LEFT JOIN mae_documento AS mae_documento_2 ON mae_docreferencia.iddoc = mae_documento_2.id) LEFT JOIN mae_cliente AS mae_cliente_1 ON var_ordendespacho.idcli = mae_cliente_1.id " _
                    + vbCr + " WHERE  " & nSQLFiltro _
                + vbCr + "  ) AS vista " _
                + vbCr + " ORDER BY vista.nomcliente,vista.fchdoc, vista.numerodoc "
                    
            If MostrarCabeceraDetallada Then
                nSQL = "SELECT c.registro AS [N° Registro], c.tdocabrev AS [Tipo Documento], c.nomcliente AS [Cliente], c.numerodoc AS [Numero Documento], c.fchdoc AS [Fecha Documento], c.fchven AS [Fecha Vencimiento], c.conpagabre AS [Condicion Pago], c.diasvenc AS [Dias Atraso], c.glosa As [Glosa], c.simbolo AS [Moneda], c.tipcam AS [Tipo de Cambio], c.impmn AS [Total MN], c.impsalmn AS [Saldo MN], c.impme As [Total ME], c.impsalme AS [Saldo ME], c.impexpmn AS [Expresado MN], c.impaboexpmn AS [Abono Expresado MN], c.impsalexpmn AS [Saldo Expresado MN], c.ref1registro AS [Ref1 N° Registro], c.ref1abrev AS [Ref1 Tipo Documento], c.ref1numdoc AS [Ref1 Numero Documento], c.ref1fchdoc AS [Ref1 Fecha Documento], c.ref2abrev AS [Ref2 Tipo Documento], c.ref2numdoc AS [Ref2 Numero Documento], c.ref2fchdoc AS [Ref2 Fecha Documento], c.ref2cliruc AS [Ref2 Ruc], c.ref2clinombre AS [Ref2 Cliente] " _
                    + vbCr + " FROM ( " _
                    + vbCr + nSQL _
                    + vbCr + " ) AS c"
            End If
            Q_POSICION_TOTAL = 9
        End If
    End If
    '------------------------------------------------------------------------------------
    pGenerarConsulta = nSQL
End Function



'--011007
Private Sub Comparar_Grupo(RST_ORIGEN As ADODB.Recordset, BAND_ADD_REG As Boolean)
    '--FUNCION QUE NOS PERMITE ARMAR LOS GRUPOS POR EL CLIENTE
    '--CUANDO SE GENERA EL GRUPO SE ARGEGA EL NOMBRE DEL CLIENTE COMO CABECERA
    '--COMPARA CUANDO CAMBIAR DE GRUPO
    Dim RST_TEPM_1 As New ADODB.Recordset
    
    Set RST_TEPM_1 = RST_ORIGEN.Clone
    RST_TEPM_1.Bookmark = RST_ORIGEN.Bookmark
    RST_TEPM_1.MovePrevious

    If RST_ORIGEN.Bookmark = 1 Then
        If OptDetalle.Value = False Then
            'ADD_REG Fg1
        End If
        Exit Sub
    End If
    
    '---------------------------------------------------------
    If RST_ORIGEN.Bookmark <> 1 Then
        If NulosC(RST_TEPM_1.Fields("nomcliente")) <> NulosC(RST_ORIGEN.Fields("nomcliente")) Then  '--CLIENTE
            If Me.OptResum.Value = True And (Trim(Me.TxtIdTipProd.Text) = "" And Me.ChkMostrarItem.Value = 0) Then Exit Sub
            CARGAR_DATOS_GRILLA_ADD_TOTALES BAND_ADD_REG, "Total:"
            '--DEL PRECIO PROMEDIO
            If VERIFICAR_PONER_PRECIO_PROMEDIO() = True Then
                CARGAR_DATOS_GRILLA_ADD_TOTALES True, "P. Prom"
            End If
            ADD_REG Fg1
            LimpiarArray
            If OptDetalle.Value = True Or (Me.OptResum.Value = True And (Trim(Me.TxtIdTipProd.Text) <> "" Or Me.ChkMostrarItem.Value = 1)) Then
                ADD_REG Fg1
                UNIR_CELDAS Fg1, Fg1.Rows - 1, 1, Fg1.Rows - 1, 5, NulosC(RST_ORIGEN.Fields("nomcliente")), flexAlignLeftCenter:      FORMATO_CELDA Fg1, Fg1.Rows - 1, 1
            End If
            Exit Sub
        End If
    End If
    Set RST_TEPM_1 = Nothing
End Sub

Private Sub LimpiarArray()
    Erase Arr_Totales
    ReDim Arr_Totales(10, 0) As Double
End Sub

Private Sub CARGAR_DATOS_GRILLA_ADD_TOTALES(BAND_ADD_TOTAL As Boolean, Nombre_total As String, Optional Band_Total_gral As Boolean = False, Optional band_forzar_suma As Boolean = False)
    '--AGREGA LOS TOTALES POR CADA GRUPO Y EL TOTAL GENERAL
    '--ACUMULA LOS TOTALES EN EL TOTAL GENERAL
    Dim X_ROW As Long
    Dim k As Integer
    
    'On Error Resume Next
    X_ROW = Fg1.Rows - 1
    If BAND_ADD_TOTAL = True Then
        ADD_REG Fg1
        X_ROW = Fg1.Rows - 1
        'PONIENDO LOS NOMBRES DE LOS TOTALES
        Fg1.TextMatrix(X_ROW, Q_POSICION_TOTAL) = Nombre_total
        FORMATO_CELDA Fg1, X_ROW, Q_POSICION_TOTAL
    End If
    
    '-----------------------------------------------------------------------------
    '--ACUMULANDO LOS TOTALES GRLES
    If Me.OptResum.Value = True Then    '--RESUMEN
        If Band_Total_gral = False And (Me.TxtIdTipProd.Text <> "" Or Me.ChkMostrarItem.Value = 1) Then
            For k = 0 To UBound(Arr_Totales())
                Arr_Totales_grls(k, 0) = Arr_Totales_grls(k, 0) + Arr_Totales(k, 0)
            Next k
        End If
    Else
        If Band_Total_gral = False Then     '--DETALLE
            For k = 0 To UBound(Arr_Totales())
                Arr_Totales_grls(k, 0) = Arr_Totales_grls(k, 0) + Arr_Totales(k, 0)
            Next k
        End If
    End If
    '-----------------------------------------------------------------------------
    
    '
    If Me.OptResum.Value = True Then
        '--RESUMEN
            With Fg1
            If Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0 Then '--PRODUCTO
                .TextMatrix(X_ROW, 4) = Format(IIf(Band_Total_gral = False, Arr_Totales(1, 0), Arr_Totales_grls(1, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 4    '"Imp. MN"
                .TextMatrix(X_ROW, 5) = Format(IIf(Band_Total_gral = False, Arr_Totales(2, 0), Arr_Totales_grls(2, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 5   '"Imp. ME"
                .TextMatrix(X_ROW, 6) = Format(IIf(Band_Total_gral = False, Arr_Totales(3, 0), Arr_Totales_grls(3, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 6    '"Total MN"
                .TextMatrix(X_ROW, 7) = Format(IIf(Band_Total_gral = False, Arr_Totales(4, 0), Arr_Totales_grls(4, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 7    '"Total ME"
                
            ElseIf Me.ChkMostrarItem.Value = 1 Then '--PRODUCTO Y ITEM
                .TextMatrix(X_ROW, 7) = Format(IIf(Band_Total_gral = False, Arr_Totales(0, 0), Arr_Totales_grls(0, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 7    '"CANTIDAD"
                .TextMatrix(X_ROW, 8) = Format(IIf(Band_Total_gral = False, Arr_Totales(1, 0), Arr_Totales_grls(1, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 8    '"Imp. ME"
                .TextMatrix(X_ROW, 9) = Format(IIf(Band_Total_gral = False, Arr_Totales(2, 0), Arr_Totales_grls(2, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 9    '"Imp. MN"
                .TextMatrix(X_ROW, 10) = Format(IIf(Band_Total_gral = False, Arr_Totales(3, 0), Arr_Totales_grls(3, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 10    '"Total mn"
                .TextMatrix(X_ROW, 11) = Format(IIf(Band_Total_gral = False, Arr_Totales(4, 0), Arr_Totales_grls(4, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 11    '"Total me"
                
            Else
                .TextMatrix(X_ROW, 3) = IIf(Band_Total_gral = False, Arr_Totales(0, 0), Arr_Totales_grls(0, 0)):: FORMATO_CELDA Fg1, X_ROW, 3    '"# Doc"
                .TextMatrix(X_ROW, 4) = Format(IIf(Band_Total_gral = False, Arr_Totales(1, 0), Arr_Totales_grls(1, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 4   '"Imp. ME"
                .TextMatrix(X_ROW, 5) = Format(IIf(Band_Total_gral = False, Arr_Totales(2, 0), Arr_Totales_grls(2, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 5    '"Saldo ME"
                .TextMatrix(X_ROW, 6) = Format(IIf(Band_Total_gral = False, Arr_Totales(3, 0), Arr_Totales_grls(3, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 6    ' "Imp. MN"
                .TextMatrix(X_ROW, 7) = Format(IIf(Band_Total_gral = False, Arr_Totales(4, 0), Arr_Totales_grls(4, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 7   '"Saldo MN"
                .TextMatrix(X_ROW, 8) = Format(IIf(Band_Total_gral = False, Arr_Totales(5, 0), Arr_Totales_grls(5, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 8    '"Total EXP MN"
                .TextMatrix(X_ROW, 9) = Format(IIf(Band_Total_gral = False, Arr_Totales(6, 0), Arr_Totales_grls(6, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 9    '"Abono EXP MN"
                .TextMatrix(X_ROW, 10) = Format(IIf(Band_Total_gral = False, Arr_Totales(7, 0), Arr_Totales_grls(7, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 10    ' "Saldo EXP MN"
                
                .TextMatrix(X_ROW, 11) = Format(IIf(Band_Total_gral = False, Arr_Totales(8, 0), Arr_Totales_grls(8, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 11    '"Total EXP ME"
                .TextMatrix(X_ROW, 12) = Format(IIf(Band_Total_gral = False, Arr_Totales(9, 0), Arr_Totales_grls(9, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 12    '"Abono EXP ME"
                .TextMatrix(X_ROW, 13) = Format(IIf(Band_Total_gral = False, Arr_Totales(10, 0), Arr_Totales_grls(10, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 13    ' "Saldo EXP ME"
                
                
            End If
        End With
    Else '-DETALLE
        With Fg1
            If Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0 Then '--PRODUCTO
                .TextMatrix(X_ROW, 13) = Format(IIf(Band_Total_gral = False, Arr_Totales(1, 0), Arr_Totales_grls(1, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 13    '"TOTAL MN"
                .TextMatrix(X_ROW, 14) = Format(IIf(Band_Total_gral = False, Arr_Totales(2, 0), Arr_Totales_grls(2, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 14    '"TOTAL ME"
                .TextMatrix(X_ROW, 15) = Format(IIf(Band_Total_gral = False, Arr_Totales(3, 0), Arr_Totales_grls(3, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 15    '"EXP MN"
                .TextMatrix(X_ROW, 16) = Format(IIf(Band_Total_gral = False, Arr_Totales(4, 0), Arr_Totales_grls(4, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 16    '"EXP ME"
                
            ElseIf Me.ChkMostrarItem.Value = 1 Then '--PRODUCTO E ITEM
                If VERIFICAR_PONER_PRECIO_PROMEDIO() = True And (Nombre_total = "P. Prom" Or Nombre_total = "P. Prom. Gen") Then
                    'Calcular_Precio_Promedio (Band_Total_gral)
                    .TextMatrix(X_ROW, 17) = CALCULAR_PRECIO_PROMEDIO(Band_Total_gral, 4): FORMATO_CELDA Fg1, X_ROW, 17 '"PRECIO PROM MN
                    .TextMatrix(X_ROW, 19) = CALCULAR_PRECIO_PROMEDIO(Band_Total_gral, 5): FORMATO_CELDA Fg1, X_ROW, 19  '"PRECIO PROM ME
                    Exit Sub
                End If
    
            
                .TextMatrix(X_ROW, 16) = Format(IIf(Band_Total_gral = False, Arr_Totales(0, 0), Arr_Totales_grls(0, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 16    '"cantidad"
                .TextMatrix(X_ROW, 18) = Format(IIf(Band_Total_gral = False, Arr_Totales(1, 0), Arr_Totales_grls(1, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 18    '"Imp. Total MN"
                .TextMatrix(X_ROW, 20) = Format(IIf(Band_Total_gral = False, Arr_Totales(2, 0), Arr_Totales_grls(2, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 20    '"Imp. Total ME"
                .TextMatrix(X_ROW, 21) = Format(IIf(Band_Total_gral = False, Arr_Totales(3, 0), Arr_Totales_grls(3, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 21    '"EXP. MN"
                .TextMatrix(X_ROW, 22) = Format(IIf(Band_Total_gral = False, Arr_Totales(4, 0), Arr_Totales_grls(4, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 22    '"EXP. ME"
                 
            Else '--SIN PRODUCTO E ITEM
                .TextMatrix(X_ROW, 12) = Format(IIf(Band_Total_gral = False, Arr_Totales(1, 0), Arr_Totales_grls(1, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 12   '"Imp. ME"
                .TextMatrix(X_ROW, 13) = Format(IIf(Band_Total_gral = False, Arr_Totales(2, 0), Arr_Totales_grls(2, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 13    '"Saldo ME"
                .TextMatrix(X_ROW, 14) = Format(IIf(Band_Total_gral = False, Arr_Totales(3, 0), Arr_Totales_grls(3, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 14     '"Imp. MN"
                .TextMatrix(X_ROW, 15) = Format(IIf(Band_Total_gral = False, Arr_Totales(4, 0), Arr_Totales_grls(4, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 15    '"Saldo MN"
                .TextMatrix(X_ROW, 16) = Format(IIf(Band_Total_gral = False, Arr_Totales(5, 0), Arr_Totales_grls(5, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 16     '"Total"
                .TextMatrix(X_ROW, 17) = Format(IIf(Band_Total_gral = False, Arr_Totales(6, 0), Arr_Totales_grls(6, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 17    '"Abono"
                .TextMatrix(X_ROW, 18) = Format(IIf(Band_Total_gral = False, Arr_Totales(7, 0), Arr_Totales_grls(7, 0)), FORMAT_MONTO): FORMATO_CELDA Fg1, X_ROW, 18     '"Saldo"
                
            End If
    
        End With
    End If
    'Err.Clear
End Sub
    
Private Sub pConfigurarGrilla()
    '--PERMITIRA CONFIGURAR EL FORMATO DE LA CONSULTA
    '--DE ACUERDO A LO QUE SE SELECCIONA
    Fg1.FrozenCols = 0
    If Me.OptResum.Value = True Then '--RESUMEN
        With Fg1
            If Trim(Me.TxtIdTipProd.Text) <> "" Or Me.ChkMostrarItem.Value = 1 Then
                .ColWidth(1) = 0 'RUC
                .ColWidth(2) = 0 'CLIENTE
            Else
                .TextMatrix(1, 1) = "RUC":          .ColWidth(1) = 1200:    .ColAlignment(1) = flexAlignCenterBottom
                .TextMatrix(1, 2) = "Cliente":      .ColWidth(2) = 2500:    .ColAlignment(2) = flexAlignLeftBottom
            End If
            If Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0 Then
                '--SOLO PRODUCTO
                '.FrozenCols = 4
                .TextMatrix(1, 3) = "Producto": .ColWidth(3) = 3500:    .ColAlignment(3) = flexAlignLeftBottom
                .TextMatrix(1, 4) = "MN":  .ColWidth(4) = 1100:   .ColAlignment(4) = flexAlignRightBottom
                .TextMatrix(1, 5) = "ME":  .ColWidth(5) = 1100:   .ColAlignment(5) = flexAlignRightBottom
                .TextMatrix(1, 6) = "MN": .ColWidth(6) = 1200:    .ColAlignment(6) = flexAlignRightBottom
                .TextMatrix(1, 7) = "ME": .ColWidth(7) = 1200:    .ColAlignment(7) = flexAlignRightBottom
                '--SOLO DOLARES OCULTAR SOLES
                'If Me.OptDol.Value = True Then .ColWidth(5) = 0
                '--SOLO SOLES OCULTAR DOLARES
                'If Me.OptSol.Value = True Then .ColWidth(4) = 0
                UNIR_CELDAS Fg1, 0, 1, 0, 3, " "
                UNIR_CELDAS Fg1, 0, 4, 0, 5, "TOTALES EN"
                UNIR_CELDAS Fg1, 0, 6, 0, 7, "EXPRESADO EN"
                OCULTAR_COL Fg1, 8, Fg1.Cols - 1
                
            ElseIf Me.ChkMostrarItem.Value = 1 Then
                '--CON PRODUCTO E ITEM
                .FrozenCols = 0
                .TextMatrix(1, 3) = "Producto":     .ColWidth(3) = 800:     .ColAlignment(3) = flexAlignLeftBottom
                .TextMatrix(1, 4) = "Código":       .ColWidth(4) = 800:      .ColAlignment(4) = flexAlignLeftBottom
                .TextMatrix(1, 5) = "Item":         .ColWidth(5) = 3000:    .ColAlignment(5) = flexAlignLeftBottom
                .TextMatrix(1, 6) = "U.M.":         .ColWidth(6) = 500:     .ColAlignment(6) = flexAlignLeftBottom
                .TextMatrix(1, 7) = "Cant.":        .ColWidth(7) = 800:     .ColAlignment(7) = flexAlignRightBottom
                
                .TextMatrix(1, 8) = "MN":       .ColWidth(8) = 1100:    .ColAlignment(8) = flexAlignRightBottom
                .TextMatrix(1, 9) = "ME":     .ColWidth(9) = 1100:    .ColAlignment(9) = flexAlignRightBottom
                .TextMatrix(1, 10) = "MN":   .ColWidth(10) = 1300:   .ColAlignment(10) = flexAlignRightBottom
                .TextMatrix(1, 11) = "ME":   .ColWidth(11) = 1300:   .ColAlignment(11) = flexAlignRightBottom
                 '--SOLO DOLARES OCULTAR SOLES
'                If Me.OptDol.Value = True Then .ColWidth(9) = 0
                '--SOLO SOLES OCULTAR DOLARES
'                If Me.OptSol.Value = True Then .ColWidth(8) = 0
                UNIR_CELDAS Fg1, 0, 1, 0, 5, " "
                UNIR_CELDAS Fg1, 0, 8, 0, 9, "TOTALES EN"
                UNIR_CELDAS Fg1, 0, 10, 0, 11, "EXPRESADO EN"
                
                OCULTAR_COL Fg1, 12, Fg1.Cols - 1
            Else
                .FrozenCols = 3
                .TextMatrix(1, 3) = "# Doc":        .ColWidth(3) = 650:     .ColAlignment(3) = flexAlignRightBottom
                '.TextMatrix(1, 4) = "M":            .ColWidth(4) = 500:     .ColAlignment(4) = flexAlignLeftBottom
                .TextMatrix(1, 4) = "Imp.":         .ColWidth(4) = 1000:    .ColAlignment(4) = flexAlignRightBottom
                .TextMatrix(1, 5) = "Saldo":        .ColWidth(5) = 1000:    .ColAlignment(5) = flexAlignRightBottom
                .TextMatrix(1, 6) = "Imp.":         .ColWidth(6) = 1000:    .ColAlignment(6) = flexAlignRightBottom
                .TextMatrix(1, 7) = "Saldo":        .ColWidth(7) = 1000:    .ColAlignment(7) = flexAlignRightBottom
                .TextMatrix(1, 8) = "Total":        .ColWidth(8) = 1200:    .ColAlignment(8) = flexAlignRightBottom
                .TextMatrix(1, 9) = "Abono":        .ColWidth(9) = 1200:     .ColAlignment(9) = flexAlignRightBottom
                .TextMatrix(1, 10) = "Saldo":       .ColWidth(10) = 1200:   .ColAlignment(10) = flexAlignRightBottom
                
                .TextMatrix(1, 11) = "Total":       .ColWidth(11) = 1200:    .ColAlignment(11) = flexAlignRightBottom
                .TextMatrix(1, 12) = "Abono":       .ColWidth(12) = 1200:    .ColAlignment(12) = flexAlignRightBottom
                .TextMatrix(1, 13) = "Saldo":       .ColWidth(13) = 1200:    .ColAlignment(13) = flexAlignRightBottom
                                
                '--SOLO PAGADO OCULTAR SALDOS
                If Me.OptPag.Value = True Then .ColWidth(5) = 0: .ColWidth(7) = 0: .ColWidth(10) = 0
'                '--SOLO DOLARES OCULTAR SOLES
'                If Me.OptDol.Value = True Then .ColWidth(6) = 0: .ColWidth(7) = 0
'                '--SOLO SOLES OCULTAR DOLARES
'                If Me.OptSol.Value = True Then .ColWidth(4) = 0: .ColWidth(5) = 0
                UNIR_CELDAS Fg1, 0, 1, 0, 3, " "
                UNIR_CELDAS Fg1, 0, 4, 0, 5, "MN"
                UNIR_CELDAS Fg1, 0, 6, 0, 8, "ME"
                UNIR_CELDAS Fg1, 0, 8, 0, 10, "EXPRESADO EN MN"
                UNIR_CELDAS Fg1, 0, 11, 0, 13, "EXPRESADO EN ME"
                OCULTAR_COL Fg1, 14, Fg1.Cols - 1
            End If
        End With
    Else '--DETALLE
        With Fg1
            .TextMatrix(1, 1) = "N°.Reg.":    .ColWidth(1) = 820:   .ColAlignment(1) = flexAlignLeftBottom
            .TextMatrix(1, 2) = "T.D.":       .ColWidth(2) = 420:   .ColAlignment(2) = flexAlignCenterBottom
            .TextMatrix(1, 3) = "Cliente":    .ColWidth(3) = 0:     .ColAlignment(3) = flexAlignLeftBottom
            
            .TextMatrix(1, 4) = "Num. Documento":   .ColWidth(4) = 1400:    .ColAlignment(4) = flexAlignCenterBottom
            .TextMatrix(1, 5) = "Fch.Doc.":         .ColWidth(5) = 840:     .ColAlignment(5) = flexAlignCenterBottom
            .TextMatrix(1, 6) = "Fch.Venc.":        .ColWidth(6) = 840:     .ColAlignment(6) = flexAlignCenterBottom
            .TextMatrix(1, 7) = "Cond. Pago":       .ColWidth(7) = 950:     .ColAlignment(7) = flexAlignRightBottom
            .TextMatrix(1, 8) = "Dias Atra..":      .ColAlignment(8) = flexAlignRightBottom
            If Me.OptPag.Value = True Then
                .ColWidth(8) = 0
            Else
                .ColWidth(8) = 800
            End If
            .TextMatrix(1, 9) = "Glosa":    .ColWidth(9) = 1500:     .ColAlignment(9) = flexAlignLeftBottom
            .TextMatrix(1, 10) = "M":       .ColWidth(10) = 450:     .ColAlignment(10) = flexAlignLeftBottom
            .TextMatrix(1, 11) = "T.C.":    .ColWidth(11) = 550:     .ColAlignment(11) = flexAlignRightBottom
            
            If Me.TxtIdTipProd.Text <> "" And Me.ChkMostrarItem.Value = 0 Then '--SOLO PRODUCTO
            
                .FrozenCols = 0
                
                .ColWidth(7) = 0
                .ColWidth(8) = 0
                 
                .TextMatrix(1, 12) = "Producto":    .ColWidth(12) = 1200: .ColAlignment(12) = flexAlignLeftBottom
                .TextMatrix(1, 13) = "MN":          .ColWidth(13) = 1000: .ColAlignment(13) = flexAlignRightBottom
                .TextMatrix(1, 14) = "ME":          .ColWidth(14) = 1000: .ColAlignment(14) = flexAlignRightBottom
                .TextMatrix(1, 15) = "MN":          .ColWidth(15) = 1000: .ColAlignment(15) = flexAlignRightBottom
                .TextMatrix(1, 16) = "ME":          .ColWidth(16) = 1000: .ColAlignment(16) = flexAlignRightBottom
                UNIR_CELDAS Fg1, 0, 1, 0, 12, " "
                UNIR_CELDAS Fg1, 0, 13, 0, 14, "TOTALES"
                UNIR_CELDAS Fg1, 0, 15, 0, 16, "EXPRESADO EN"
                OCULTAR_COL Fg1, 17, Fg1.Cols - 1
                
            ElseIf Me.ChkMostrarItem.Value = 1 Then '--ITEM
                .FrozenCols = 6
                
                 .ColWidth(7) = 0
                 .ColWidth(8) = 0
                 
                .TextMatrix(1, 12) = "Producto":    .ColWidth(12) = 900:    .ColAlignment(12) = flexAlignLeftBottom
                .TextMatrix(1, 13) = "Código":      .ColWidth(13) = 900:    .ColAlignment(13) = flexAlignLeftBottom
                .TextMatrix(1, 14) = "Item":        .ColWidth(14) = 2800:   .ColAlignment(14) = flexAlignLeftBottom
                
                .TextMatrix(1, 15) = "U.M.":        .ColWidth(15) = 500:    .ColAlignment(15) = flexAlignLeftBottom
                .TextMatrix(1, 16) = "Cant.":       .ColWidth(16) = 700:    .ColAlignment(16) = flexAlignRightBottom
                .TextMatrix(1, 17) = "P/U":         .ColWidth(17) = 500:    .ColAlignment(17) = flexAlignRightBottom
                .TextMatrix(1, 18) = "Imp.Total":   .ColWidth(18) = 900:    .ColAlignment(18) = flexAlignRightBottom
                .TextMatrix(1, 19) = "P/U":         .ColWidth(19) = 500:    .ColAlignment(19) = flexAlignRightBottom
                .TextMatrix(1, 20) = "Imp.Total":   .ColWidth(20) = 900:    .ColAlignment(20) = flexAlignRightBottom
                .TextMatrix(1, 21) = "MN":          .ColWidth(21) = 1000:   .ColAlignment(21) = flexAlignRightBottom
                .TextMatrix(1, 22) = "ME":          .ColWidth(22) = 1000:   .ColAlignment(22) = flexAlignRightBottom
                
                UNIR_CELDAS Fg1, 0, 1, 0, 16, " "
                UNIR_CELDAS Fg1, 0, 17, 0, 18, "MN"
                UNIR_CELDAS Fg1, 0, 19, 0, 20, "ME"
                UNIR_CELDAS Fg1, 0, 21, 0, 22, "EXPRESADO EN"
                
                OCULTAR_COL Fg1, 23, Fg1.Cols - 1
            Else
                .FrozenCols = 5
                UNIR_CELDAS Fg1, 0, 12, 0, 13, "MN"
                .TextMatrix(1, 12) = "Imp.":    .ColWidth(12) = 900: .ColAlignment(12) = flexAlignRightBottom
                .TextMatrix(1, 13) = "Saldo":   .ColWidth(13) = 900: .ColAlignment(13) = flexAlignRightBottom
                
                UNIR_CELDAS Fg1, 0, 14, 0, 15, "ME"
                .TextMatrix(1, 14) = "Imp.":    .ColWidth(14) = 900: .ColAlignment(14) = flexAlignRightBottom
                .TextMatrix(1, 15) = "Saldo":   .ColWidth(15) = 900: .ColAlignment(15) = flexAlignRightBottom
                
                UNIR_CELDAS Fg1, 0, 16, 0, 18, "EXPRESADO EN MN"
                .TextMatrix(1, 16) = "Total":   .ColWidth(16) = 1100: .ColAlignment(16) = flexAlignRightBottom
                .TextMatrix(1, 17) = "Abono":   .ColWidth(17) = 1100: .ColAlignment(17) = flexAlignRightBottom
                .TextMatrix(1, 18) = "Saldo":   .ColWidth(18) = 1200: .ColAlignment(18) = flexAlignRightBottom
                
                
                UNIR_CELDAS Fg1, 0, 19, 0, 22, "REFERENCIA 1"
                .TextMatrix(1, 19) = "N°.Reg.":         .ColWidth(19) = 820: .ColAlignment(19) = flexAlignLeftCenter
                .TextMatrix(1, 20) = "T.D.":            .ColWidth(20) = 420: .ColAlignment(20) = flexAlignCenterBottom
                .TextMatrix(1, 21) = "Num. Documento":  .ColWidth(21) = 1400: .ColAlignment(21) = flexAlignCenterBottom
                .TextMatrix(1, 22) = "Fch.Doc.":        .ColWidth(22) = 840: .ColAlignment(22) = flexAlignCenterBottom
                
                UNIR_CELDAS Fg1, 0, 23, 0, 27, "REFERENCIA 2"
                .TextMatrix(1, 23) = "T.D.":            .ColWidth(23) = 420:    .ColAlignment(23) = flexAlignCenterBottom
                .TextMatrix(1, 24) = "Num. Documento":  .ColWidth(24) = 1500:   .ColAlignment(24) = flexAlignCenterBottom
                .TextMatrix(1, 25) = "Fch.Doc.":        .ColWidth(25) = 840:    .ColAlignment(25) = flexAlignCenterBottom
                .TextMatrix(1, 26) = "Ruc":             .ColWidth(26) = 1100:   .ColAlignment(25) = flexAlignLeftCenter
                .TextMatrix(1, 27) = "Cliente":         .ColWidth(27) = 2000:   .ColAlignment(27) = flexAlignLeftCenter
                
                
                '--SOLO DOLARES OCULTAR SOLES
                If Me.OptDol.Value = True Then
                    .ColWidth(14) = 0: .ColWidth(15) = 0
                    .ColWidth(12) = 1000: .ColWidth(13) = 1000
                    .ColWidth(16) = 1000: .ColWidth(17) = 1000: .ColWidth(18) = 1250
                End If
                '--SOLO SOLES OCULTAR DOLARES
                If Me.OptSol.Value = True Then
                    .ColWidth(12) = 0: .ColWidth(13) = 0
                    .ColWidth(14) = 1000: .ColWidth(15) = 1000
                    .ColWidth(16) = 1000: .ColWidth(17) = 1000: .ColWidth(18) = 1000
                End If
                '--SOLO PAGADO OCULTAR SALDOS
                If Me.OptPag.Value = True Then .ColWidth(13) = 0: .ColWidth(15) = 0: .ColWidth(18) = 0
    
                UNIR_CELDAS Fg1, 0, 1, 0, 11, " "
                
                OCULTAR_COL Fg1, 28, Fg1.Cols - 1
            End If
        End With
    End If
End Sub




Private Sub PosicionarProgBar()
    '--POSICIONAR EL PROGRESO DENTRO DEL FORMULARIO
    '    FraProgreso.Width = 5820
    FraProgreso.Left = (Me.Width - FraProgreso.Width) \ 2
    FraProgreso.Top = (Me.Height - FraProgreso.Height) \ 2
    FraProgreso.Visible = True
End Sub


Private Function VERIFICAR_PONER_PRECIO_PROMEDIO() As Boolean
    '--VERIFICAR SI INSERTARA EL PRECIO PROMEDIO
    '--SOLO ESTA ACTIVO CUANDO SELECCIONE UN ITEM, LA SELECCION DE CLIENTE
    '--SI INSERTA SERA EN LA SIGUIENTE FILA DE LOS TOTALES
    Dim k, M_CANTIDAD_REGI As Integer
    
    '--DEL ITEM: M_CANTIDAD_REGI_CLI = 0
    If Me.OptResum.Value = True Then Exit Function
    If Me.ChkMostrarItem.Value = 0 Then GoTo Salir_FUNC
    With Fg2
        For k = 0 To .Rows - 1
            If Me.ChkMostrarItem.Value = 0 Then GoTo Salir_FUNC '--SALIR SI NO SELECCIONA MOSTRAR ITEM
            If k + 1 = .Rows Then Exit For
            'M_CANTIDAD_REGI = M_CANTIDAD_REGI + 1
            If CStr(.TextMatrix(k + 1, 3)) <> "" Then M_CANTIDAD_REGI = M_CANTIDAD_REGI + 1
        Next k
    End With
    
    If M_CANTIDAD_REGI = 1 Then VERIFICAR_PONER_PRECIO_PROMEDIO = True
    Exit Function
Salir_FUNC:
End Function

Private Function Verificar_Poner_Datos_Grls() As Boolean
    '--VERIFICAR SI INSERTARA LOS DATOS GENERALES TANTO PARA LOS MONTOS Y PRECIO PROMEDIO
    '--NO INSERTARA CUANDO SELECCIONA UN CLIENTE
    Dim k, M_CANTIDAD_REGI_CLI As Integer
    '--DEL ITEM
    M_CANTIDAD_REGI_CLI = 0
    With Fg3
        For k = 0 To .Rows - 1
            If k + 1 = .Rows Then Exit For
            If CStr(.TextMatrix(k + 1, 1)) <> "" Then M_CANTIDAD_REGI_CLI = M_CANTIDAD_REGI_CLI + 1
        Next k
    End With
    '---
    If M_CANTIDAD_REGI_CLI = 1 Then Exit Function
    Verificar_Poner_Datos_Grls = True
End Function


Private Function CALCULAR_PRECIO_PROMEDIO(Band_Total_gral As Boolean, M_POS As Integer) As String
    '--M_POS = 4: PU MN
    '--M_POS = 5: PU ME
    '--M_POS = 6 CANTIDAD DE REGISTROS
    If (Arr_Totales(M_POS, 0) = 0) And (Arr_Totales_grls(M_POS, 0) = 0) Then
        CALCULAR_PRECIO_PROMEDIO = ""
        Exit Function
    End If
    
    
    If Band_Total_gral = False Then
        If NulosN(Arr_Totales(6, 0)) = 0 Then
            CALCULAR_PRECIO_PROMEDIO = 0
        Else
            CALCULAR_PRECIO_PROMEDIO = Format(NulosN(Arr_Totales(M_POS, 0)) / NulosN(Arr_Totales(6, 0)), FORMAT_MONTO)
        End If
    Else
        If NulosN(Arr_Totales(6, 0)) = 0 Then
            CALCULAR_PRECIO_PROMEDIO = 0
        Else
            CALCULAR_PRECIO_PROMEDIO = Format(NulosN(Arr_Totales_grls(M_POS, 0)) / NulosN(Arr_Totales_grls(6, 0)), FORMAT_MONTO)
        End If
    End If
    
End Function

'--------
Private Sub pExportarExcel()
    On Error GoTo error
    Dim X_EXPORT As New SGI2_funciones.formularios

    Me.MousePointer = vbHourglass
    X_EXPORT.VSFlexGrid_Exportar_MSExcel xCon, Fg1, T_RPT_TITULO + " ", "", T_RPT_PERIODO, "Ventas"
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
error:
    Set X_EXPORT = Nothing
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "pExportarExcel"
End Sub



'************************************************

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then pConsultar
    If Button.Index = 5 Then pExportarExcel
    If Button.Index = 6 Then pImprimir
    If Button.Index = 8 Then
        Unload Me
    End If
End Sub

'************************************************


Private Sub pCargaItem(Seleccionar As Boolean)
    Dim nSQL As String
    Dim nSQLNotIn  As String
    Dim xRs As New ADODB.Recordset
    
    On Error GoTo error
    
        If TxtIdTipProd.Text = "" Then
            '--posicionar en la primera pestaña del menu de opciones de consulta
            TabOne2.CurrTab = 0
            '---
            MsgBox "Falta especificar el tipo de item!", vbExclamation, xTitulo
            TxtIdTipProd.SetFocus
            Exit Sub
        End If
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        
        Dim xCampos(3, 4) As String
        
        xCampos(0, 0) = "Descripción":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Código":        xCampos(1, 1) = "codpro":         xCampos(1, 2) = "2000":         xCampos(1, 3) = "c"
        xCampos(2, 0) = "Id":            xCampos(2, 1) = "id":             xCampos(2, 2) = "600":          xCampos(2, 3) = "N"
                
        nSQLNotIn = GRID_GENERAR_SQL_ID(Fg2, 3, " AND alm_inventario.id", "NOT IN", True)
        
        '--si se ingresa algun filtro adicional
        If NulosC(Fg2.TextMatrix(Fg2.Row, Fg2.Col)) <> "" Then
            nSQLNotIn = nSQLNotIn & " AND (UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(Fg2.TextMatrix(Fg2.Row, Fg2.Col))) & "%' OR UCASE(alm_inventario.descripcion) LIKE '%" & UCase(NulosC(Fg2.TextMatrix(Fg2.Row, Fg2.Col))) & "%' ) "
        End If
        
        Fg2.TextMatrix(Fg2.Row, Fg2.Col) = ""
        
        nSQL = "SELECT 0 as xsel,id, codpro, descripcion FROM alm_inventario WHERE tippro = " & NulosN(TxtIdTipProd.Text) & nSQLNotIn & ""
        
        '--muestra pantalla para buscar o seleccionar datos
        If Seleccionar = False Then
            CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Tipo de Item", "descripcion", "descripcion", Principio
        Else
            CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscando Tipo de Item"
        End If
        
        If xRs.State = 0 Then GoTo SALIR
        If xRs.RecordCount = 0 Then GoTo SALIR
        
        Do While Not xRs.EOF
            Fg2.TextMatrix(Fg2.Row, 1) = NulosC(xRs("codpro"))
            Fg2.TextMatrix(Fg2.Row, 2) = NulosC(xRs("descripcion"))
            Fg2.TextMatrix(Fg2.Row, 3) = NulosN(xRs("id"))
            '--agrega nuevo registro
            If Fg2.Row = Fg2.Rows - 1 Then Fg2.AddItem ""
            '--posicionando el cursor en el siguiente registro
            Fg2.Row = Fg2.Rows - 1: Fg2.Col = 2
            If Seleccionar = False Then Exit Do
            xRs.MoveNext
        Loop
SALIR:
        Set xRs = Nothing
    
    Exit Sub
error:
        Set xRs = Nothing
        MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"
End Sub

Private Sub TxtIdTipProd_Validate(Cancel As Boolean)
    
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
    SendKeys vbTab
    
    Set RsTipProd = Nothing
    
End Sub

Private Sub pCargarCliente(Seleccionar As Boolean)
    Dim nSQLNotIn As String
    Dim nSQL As String
    On Error GoTo error
    
        
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(3, 4) As String
    
    xCampos(0, 0) = "Nombre":   xCampos(0, 1) = "nombre":    xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Ruc":      xCampos(1, 1) = "numruc":    xCampos(1, 2) = "1500":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "Id":       xCampos(2, 1) = "id":        xCampos(2, 2) = "800":   xCampos(2, 3) = "N"
    
    nSQLNotIn = GRID_GENERAR_SQL_ID(Fg3, 1, " WHERE mae_cliente.id", "NOT IN", True)
    
    '--si se ingresa algun filtro adicional
    If NulosC(Fg3.TextMatrix(Fg3.Row, Fg3.Col)) <> "" Then
        nSQLNotIn = IIf(nSQLNotIn = "", " WHERE ", nSQLNotIn & " AND ") & "  (UCASE(mae_cliente.nombre) LIKE '%" & UCase(NulosC(Fg3.TextMatrix(Fg3.Row, Fg3.Col))) & "%' OR UCASE(mae_cliente.nombre) LIKE '%" & UCase(NulosC(Fg3.TextMatrix(Fg3.Row, Fg3.Col))) & "%' ) "
    End If
    
    Fg3.TextMatrix(Fg3.Row, Fg3.Col) = ""
    
    nSQL = "SELECT 0 as xsel,mae_cliente.* FROM mae_cliente " & nSQLNotIn & " order by nombre asc"
    
    
    If Seleccionar = False Then
        CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Clientes", "nombre", "nombre", Principio
    Else
        CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), "Buscando Clientes"
    End If
    
    If xRs.State = 0 Then GoTo SALIR
    If xRs.RecordCount = 0 Then GoTo SALIR
    
    Do While Not xRs.EOF
        Fg3.TextMatrix(Fg3.Row, 1) = Trim(xRs("id"))
        Fg3.TextMatrix(Fg3.Row, 2) = NulosC(xRs("nombre"))
        
        '--agrega nuevo registro
        If Fg3.Row = Fg3.Rows - 1 Then Fg3.AddItem ""
        '--posicionando el cursor en el siguiente registro
        Fg3.Row = Fg3.Rows - 1: Fg3.Col = 2
    
        If Seleccionar = False Then Exit Do
        xRs.MoveNext
    Loop
    
SALIR:
        
        Set xRs = Nothing
    
    Exit Sub
error:
        
        Set xRs = Nothing
        MsgBox Err.Description + vbCr + Err.Source + vbCr + CStr(Err.Number), vbCritical, "Error"

End Sub
