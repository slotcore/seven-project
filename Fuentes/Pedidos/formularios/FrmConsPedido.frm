VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsPedido 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas  -  Reporte de Pedidos"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   12000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   12000
   StartUpPosition =   3  'Windows Default
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6885
      Left            =   0
      TabIndex        =   1
      Top             =   390
      Width           =   12000
      _cx             =   21167
      _cy             =   12144
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   -1  'True
      Appearance      =   1
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   -2147483630
      Caption         =   "&Entregados|&No Entregados"
      Align           =   0
      CurrTab         =   1
      FirstTab        =   0
      Style           =   0
      Position        =   0
      AutoSwitch      =   -1  'True
      AutoScroll      =   -1  'True
      TabPreview      =   -1  'True
      ShowFocusRect   =   -1  'True
      TabsPerPage     =   0
      BorderWidth     =   0
      BoldCurrent     =   0   'False
      DogEars         =   -1  'True
      MultiRow        =   0   'False
      MultiRowOffset  =   200
      CaptionStyle    =   0
      TabHeight       =   0
      TabCaptionPos   =   4
      TabPicturePos   =   0
      CaptionEmpty    =   ""
      Separators      =   0   'False
      Begin VB.Frame Frame12 
         BorderStyle     =   0  'None
         Height          =   6510
         Left            =   45
         TabIndex        =   59
         Top             =   330
         Width           =   11910
         Begin VB.Frame Frame13 
            BorderStyle     =   0  'None
            Height          =   6510
            Left            =   0
            TabIndex        =   60
            Top             =   0
            Width           =   11910
            Begin VB.Frame Frame17 
               BorderStyle     =   0  'None
               Height          =   2505
               Left            =   0
               TabIndex        =   74
               Top             =   0
               Width           =   11900
               Begin VB.Frame Frame22 
                  Caption         =   "Situacion"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1995
                  Left            =   9360
                  TabIndex        =   92
                  Top             =   450
                  Width           =   2400
                  Begin VB.CheckBox CheckATiempo2 
                     Caption         =   "CheckATiempo"
                     Height          =   195
                     Left            =   650
                     TabIndex        =   96
                     Top             =   570
                     Width           =   170
                  End
                  Begin VB.CheckBox CheckAntesDeTiempo2 
                     Caption         =   "CheckAntesDeTiempo"
                     Height          =   195
                     Left            =   650
                     TabIndex        =   95
                     Top             =   990
                     Width           =   170
                  End
                  Begin VB.CheckBox CheckADestiempo2 
                     Caption         =   "CheckADestiempo"
                     Height          =   195
                     Left            =   650
                     TabIndex        =   94
                     Top             =   1410
                     Width           =   170
                  End
                  Begin VB.CheckBox CheckSituacion2 
                     Caption         =   "CheckSituacion"
                     Height          =   195
                     Left            =   300
                     TabIndex        =   93
                     Top             =   570
                     Width           =   170
                  End
                  Begin VB.Label Label42 
                     AutoSize        =   -1  'True
                     Caption         =   "A Tiempo"
                     Height          =   195
                     Left            =   930
                     TabIndex        =   99
                     Top             =   570
                     Width           =   675
                  End
                  Begin VB.Label Label41 
                     AutoSize        =   -1  'True
                     Caption         =   "Antes de Tiermpo"
                     Height          =   195
                     Left            =   930
                     TabIndex        =   98
                     Top             =   990
                     Width           =   1245
                  End
                  Begin VB.Label Label40 
                     AutoSize        =   -1  'True
                     Caption         =   "A Destiempo"
                     Height          =   195
                     Left            =   930
                     TabIndex        =   97
                     Top             =   1410
                     Width           =   900
                  End
               End
               Begin VB.Frame Frame18 
                  Caption         =   "Detalles"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1990
                  Left            =   6840
                  TabIndex        =   81
                  Top             =   450
                  Width           =   2415
                  Begin AspaTextBoxFecha.TextBoxFecha TextBoxFechaEmision1_2 
                     Height          =   300
                     Left            =   1000
                     TabIndex        =   82
                     Top             =   420
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
                  Begin AspaTextBoxFecha.TextBoxFecha TextBoxFechaEmision2_2 
                     Height          =   300
                     Left            =   1000
                     TabIndex        =   83
                     Top             =   750
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
                  Begin AspaTextBoxFecha.TextBoxFecha TextBoxFechaPlazo1_2 
                     Height          =   300
                     Left            =   1000
                     TabIndex        =   84
                     Top             =   1300
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
                  Begin AspaTextBoxFecha.TextBoxFecha TextBoxFechaPlazo2_2 
                     Height          =   300
                     Left            =   1000
                     TabIndex        =   85
                     Top             =   1600
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
                  Begin VB.Label Label38 
                     AutoSize        =   -1  'True
                     Caption         =   "Hasta:"
                     Height          =   195
                     Left            =   450
                     TabIndex        =   91
                     Top             =   1650
                     Width           =   465
                  End
                  Begin VB.Label Label37 
                     AutoSize        =   -1  'True
                     Caption         =   "Desde:"
                     Height          =   195
                     Left            =   450
                     TabIndex        =   90
                     Top             =   1350
                     Width           =   510
                  End
                  Begin VB.Label Label36 
                     AutoSize        =   -1  'True
                     Caption         =   "Hasta:"
                     Height          =   195
                     Left            =   450
                     TabIndex        =   89
                     Top             =   795
                     Width           =   465
                  End
                  Begin VB.Label Label35 
                     AutoSize        =   -1  'True
                     Caption         =   "Desde:"
                     Height          =   195
                     Left            =   450
                     TabIndex        =   88
                     Top             =   465
                     Width           =   510
                  End
                  Begin VB.Label Label34 
                     AutoSize        =   -1  'True
                     Caption         =   "Fech. a Entregar"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   90
                     TabIndex        =   87
                     Top             =   1100
                     Width           =   1185
                  End
                  Begin VB.Label Label33 
                     AutoSize        =   -1  'True
                     Caption         =   "Fech. Emision"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   400
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     Height          =   195
                     Left            =   120
                     TabIndex        =   86
                     Top             =   200
                     Width           =   990
                  End
               End
               Begin VB.Frame Frame21 
                  Caption         =   "Ordenes de Pedido"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1990
                  Left            =   3450
                  TabIndex        =   79
                  Top             =   450
                  Width           =   3280
                  Begin VSFlex7Ctl.VSFlexGrid VSFlexGridOC2 
                     Height          =   1610
                     Left            =   90
                     TabIndex        =   80
                     ToolTipText     =   "Buscar Ordenes de Pedido"
                     Top             =   330
                     Width           =   3100
                     _cx             =   5468
                     _cy             =   2840
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
                     FormatString    =   $"FrmConsPedido.frx":0000
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
               Begin VB.Frame Frame20 
                  Height          =   495
                  Left            =   10
                  TabIndex        =   77
                  Top             =   -50
                  Width           =   11865
                  Begin VB.Label Label39 
                     AutoSize        =   -1  'True
                     Caption         =   "Reporte de Pedidos No Entregados"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   9.75
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   240
                     Left            =   3900
                     TabIndex        =   78
                     Top             =   150
                     Width           =   3735
                  End
               End
               Begin VB.Frame Frame19 
                  Caption         =   "Clientes"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1990
                  Left            =   120
                  TabIndex        =   75
                  Top             =   450
                  Width           =   3280
                  Begin VSFlex7Ctl.VSFlexGrid VSFlexGridClientes2 
                     Height          =   1605
                     Left            =   90
                     TabIndex        =   76
                     ToolTipText     =   "Buscar Clientes"
                     Top             =   330
                     Width           =   3100
                     _cx             =   5468
                     _cy             =   2831
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
                     FormatString    =   $"FrmConsPedido.frx":003D
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
            End
            Begin VB.Frame Frame14 
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
               Height          =   3915
               Left            =   0
               TabIndex        =   61
               Top             =   2500
               Width           =   11865
               Begin VB.Frame FrameEspecificaciones2 
                  Caption         =   "Especificaciones"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   3850
                  Left            =   11010
                  TabIndex        =   62
                  Top             =   510
                  Visible         =   0   'False
                  Width           =   11805
                  Begin VB.Frame Frame16 
                     Height          =   885
                     Left            =   150
                     TabIndex        =   63
                     Top             =   540
                     Width           =   11445
                     Begin VB.Label Label29 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "SITUACION:"
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
                        Height          =   195
                        Left            =   6120
                        TabIndex        =   67
                        Top             =   300
                        Width           =   1080
                     End
                     Begin VB.Label LabelDetalleSituacion2 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "RETRASO_01"
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
                        Left            =   7500
                        TabIndex        =   66
                        Top             =   300
                        Width           =   1215
                     End
                     Begin VB.Label LabelDetalleEntgr2 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "ENTREGAR_01"
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
                        Left            =   3100
                        TabIndex        =   65
                        Top             =   300
                        Width           =   1350
                     End
                     Begin VB.Label Label20 
                        AutoSize        =   -1  'True
                        BackStyle       =   0  'Transparent
                        Caption         =   "CANT. A ENTREGAR:"
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
                        Height          =   195
                        Left            =   285
                        TabIndex        =   64
                        Top             =   300
                        Width           =   1890
                     End
                  End
                  Begin VSFlex7Ctl.VSFlexGrid VSFlexGridDetalle2_2 
                     Height          =   1500
                     Left            =   105
                     TabIndex        =   68
                     Top             =   1815
                     Width           =   11550
                     _cx             =   20373
                     _cy             =   2646
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
                     Rows            =   50
                     Cols            =   10
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
                  Begin VB.Label Label23 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "PEDIDOS"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Left            =   5310
                     TabIndex        =   102
                     Top             =   1530
                     Width           =   840
                  End
                  Begin VB.Label LabelCerrar2 
                     AutoSize        =   -1  'True
                     Caption         =   "X"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   12
                        Charset         =   0
                        Weight          =   700
                        Underline       =   -1  'True
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H000000FF&
                     Height          =   300
                     Left            =   11580
                     TabIndex        =   71
                     Top             =   30
                     Width           =   195
                  End
                  Begin VB.Label LabelDetalleOC2 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "OC_01"
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
                     Left            =   3300
                     TabIndex        =   70
                     Top             =   360
                     Width           =   585
                  End
                  Begin VB.Label Label30 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "Nº ORDEN DE PEDIDO:"
                     BeginProperty Font 
                        Name            =   "MS Sans Serif"
                        Size            =   8.25
                        Charset         =   0
                        Weight          =   700
                        Underline       =   0   'False
                        Italic          =   0   'False
                        Strikethrough   =   0   'False
                     EndProperty
                     ForeColor       =   &H00FF0000&
                     Height          =   195
                     Left            =   450
                     TabIndex        =   69
                     Top             =   375
                     Width           =   2085
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid VSFlexGridReporte2 
                  Height          =   3000
                  Left            =   90
                  TabIndex        =   72
                  Top             =   270
                  Width           =   11685
                  _cx             =   20611
                  _cy             =   5292
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
                  Rows            =   50
                  Cols            =   10
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
               Begin VB.CommandButton Command1 
                  Caption         =   "Limpiar Todo"
                  Height          =   465
                  Left            =   10050
                  TabIndex        =   73
                  Top             =   3360
                  Width           =   1725
               End
            End
         End
      End
      Begin VB.Frame FrameContEnt 
         BorderStyle     =   0  'None
         Height          =   6510
         Left            =   -12555
         TabIndex        =   3
         Top             =   330
         Width           =   11910
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
            Height          =   3915
            Left            =   0
            TabIndex        =   39
            Top             =   2500
            Width           =   11865
            Begin VB.Frame FrameEspecificaciones 
               Caption         =   "Especificaciones"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   3850
               Left            =   10560
               TabIndex        =   40
               Top             =   300
               Visible         =   0   'False
               Width           =   11805
               Begin VB.Frame Frame5 
                  Height          =   1275
                  Left            =   180
                  TabIndex        =   41
                  Top             =   540
                  Width           =   11445
                  Begin VB.Line Line1 
                     X1              =   210
                     X2              =   4500
                     Y1              =   810
                     Y2              =   810
                  End
                  Begin VB.Label LabelResto 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "RESTO_01"
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
                     Left            =   3100
                     TabIndex        =   51
                     Top             =   900
                     Width           =   960
                  End
                  Begin VB.Label Label14 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "RESTO:"
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
                     Height          =   195
                     Left            =   285
                     TabIndex        =   50
                     Top             =   900
                     Width           =   705
                  End
                  Begin VB.Label Label17 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "CANT. A ENTREGAR:"
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
                     Height          =   195
                     Left            =   285
                     TabIndex        =   49
                     Top             =   250
                     Width           =   1890
                  End
                  Begin VB.Label LabelDetalleEntgr 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "ENTREGAR_01"
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
                     Left            =   3100
                     TabIndex        =   48
                     Top             =   285
                     Width           =   1350
                  End
                  Begin VB.Label LabelDetalleEntgda 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "ENTREGADA_01"
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
                     Left            =   3100
                     TabIndex        =   47
                     Top             =   585
                     Width           =   1470
                  End
                  Begin VB.Label Label8 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "CANT. ENTREGADA:"
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
                     Height          =   195
                     Left            =   285
                     TabIndex        =   46
                     Top             =   550
                     Width           =   1830
                  End
                  Begin VB.Label Label7 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "ESTADO:"
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
                     Height          =   195
                     Left            =   6135
                     TabIndex        =   45
                     Top             =   255
                     Width           =   825
                  End
                  Begin VB.Label LabelDetalleRetraso 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "RETRASO_01"
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
                     Left            =   7500
                     TabIndex        =   44
                     Top             =   255
                     Width           =   1215
                  End
                  Begin VB.Label LabelDetalleSituacion 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "RETRASO_01"
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
                     Left            =   7500
                     TabIndex        =   43
                     Top             =   555
                     Width           =   1215
                  End
                  Begin VB.Label Label16 
                     AutoSize        =   -1  'True
                     BackStyle       =   0  'Transparent
                     Caption         =   "SITUACION:"
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
                     Height          =   195
                     Left            =   6120
                     TabIndex        =   42
                     Top             =   555
                     Width           =   1080
                  End
               End
               Begin VSFlex7Ctl.VSFlexGrid VSFlexGridDetalle2 
                  Height          =   1500
                  Left            =   100
                  TabIndex        =   52
                  Top             =   2150
                  Width           =   6000
                  _cx             =   10583
                  _cy             =   2646
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
                  Rows            =   50
                  Cols            =   10
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
               Begin VSFlex7Ctl.VSFlexGrid VSFlexGridDetalle 
                  Height          =   1500
                  Left            =   6100
                  TabIndex        =   53
                  Top             =   2150
                  Width           =   5600
                  _cx             =   9878
                  _cy             =   2646
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
                  Rows            =   50
                  Cols            =   10
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
               Begin VB.Label Label15 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "ENTREGAS"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Left            =   8370
                  TabIndex        =   101
                  Top             =   1890
                  Width           =   1020
               End
               Begin VB.Label Label6 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "PEDIDOS"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Left            =   2520
                  TabIndex        =   100
                  Top             =   1890
                  Width           =   840
               End
               Begin VB.Label Label9 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Nº ORDEN DE PEDIDO:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H00FF0000&
                  Height          =   195
                  Left            =   450
                  TabIndex        =   56
                  Top             =   375
                  Width           =   2085
               End
               Begin VB.Label LabelDetalleOC 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "OC_01"
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
                  Left            =   3300
                  TabIndex        =   55
                  Top             =   360
                  Width           =   585
               End
               Begin VB.Label LabelCerrar 
                  AutoSize        =   -1  'True
                  Caption         =   "X"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   12
                     Charset         =   0
                     Weight          =   700
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   300
                  Left            =   11580
                  TabIndex        =   54
                  Top             =   30
                  Width           =   195
               End
            End
            Begin VSFlex7Ctl.VSFlexGrid VSFlexGridReporte 
               Height          =   3000
               Left            =   90
               TabIndex        =   58
               Top             =   270
               Width           =   11685
               _cx             =   20611
               _cy             =   5292
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
               Rows            =   50
               Cols            =   10
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
            Begin VB.CommandButton CommandLimpiar 
               Caption         =   "Limpiar Todo"
               Height          =   465
               Left            =   10050
               TabIndex        =   57
               Top             =   3360
               Width           =   1725
            End
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   2565
            Left            =   0
            TabIndex        =   4
            Top             =   0
            Width           =   11900
            Begin VB.Frame Frame7 
               Caption         =   "Detalles"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1990
               Left            =   6780
               TabIndex        =   28
               Top             =   450
               Width           =   2415
               Begin AspaTextBoxFecha.TextBoxFecha TextBoxFechaEmision1 
                  Height          =   300
                  Left            =   1000
                  TabIndex        =   29
                  Top             =   420
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
               Begin AspaTextBoxFecha.TextBoxFecha TextBoxFechaEmision2 
                  Height          =   300
                  Left            =   1000
                  TabIndex        =   30
                  Top             =   750
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
               Begin AspaTextBoxFecha.TextBoxFecha TextBoxFechaPlazo1 
                  Height          =   300
                  Left            =   1000
                  TabIndex        =   31
                  Top             =   1300
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
               Begin AspaTextBoxFecha.TextBoxFecha TextBoxFechaPlazo2 
                  Height          =   300
                  Left            =   1000
                  TabIndex        =   32
                  Top             =   1600
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
               Begin VB.Label Label2 
                  AutoSize        =   -1  'True
                  Caption         =   "Fech. Emision"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   120
                  TabIndex        =   38
                  Top             =   200
                  Width           =   990
               End
               Begin VB.Label Label3 
                  AutoSize        =   -1  'True
                  Caption         =   "Fech. a Entregar"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   400
                     Underline       =   -1  'True
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   90
                  TabIndex        =   37
                  Top             =   1100
                  Width           =   1185
               End
               Begin VB.Label Label5 
                  AutoSize        =   -1  'True
                  Caption         =   "Desde:"
                  Height          =   195
                  Left            =   450
                  TabIndex        =   36
                  Top             =   465
                  Width           =   510
               End
               Begin VB.Label Label10 
                  AutoSize        =   -1  'True
                  Caption         =   "Hasta:"
                  Height          =   195
                  Left            =   450
                  TabIndex        =   35
                  Top             =   795
                  Width           =   465
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  Caption         =   "Desde:"
                  Height          =   195
                  Left            =   450
                  TabIndex        =   34
                  Top             =   1350
                  Width           =   510
               End
               Begin VB.Label Label12 
                  AutoSize        =   -1  'True
                  Caption         =   "Hasta:"
                  Height          =   195
                  Left            =   450
                  TabIndex        =   33
                  Top             =   1650
                  Width           =   465
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "Clientes"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1990
               Left            =   120
               TabIndex        =   26
               Top             =   450
               Width           =   3280
               Begin VSFlex7Ctl.VSFlexGrid VSFlexGridClientes 
                  Height          =   1605
                  Left            =   90
                  TabIndex        =   27
                  ToolTipText     =   "Buscar Clientes"
                  Top             =   330
                  Width           =   3100
                  _cx             =   5468
                  _cy             =   2831
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
                  FormatString    =   $"FrmConsPedido.frx":007A
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
            Begin VB.Frame Frame3 
               Height          =   495
               Left            =   10
               TabIndex        =   24
               Top             =   -50
               Width           =   11865
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  Caption         =   "Reporte de Pedidos Entregados"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   9.75
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   240
                  Left            =   4000
                  TabIndex        =   25
                  Top             =   150
                  Width           =   3375
               End
            End
            Begin VB.Frame Frame8 
               Caption         =   "Ordenes de Pedido"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1990
               Left            =   3450
               TabIndex        =   22
               Top             =   450
               Width           =   3280
               Begin VSFlex7Ctl.VSFlexGrid VSFlexGridOC 
                  Height          =   1610
                  Left            =   90
                  TabIndex        =   23
                  ToolTipText     =   "Buscar Ordenes de Pedido"
                  Top             =   330
                  Width           =   3100
                  _cx             =   5468
                  _cy             =   2840
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
                  FormatString    =   $"FrmConsPedido.frx":00B7
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
            Begin VB.Frame Frame9 
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2055
               Left            =   9270
               TabIndex        =   5
               Top             =   450
               Width           =   2535
               Begin VB.Frame Frame10 
                  Caption         =   "Estado"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   915
                  Left            =   60
                  TabIndex        =   14
                  Top             =   5
                  Width           =   2400
                  Begin VB.CheckBox CheckCompleto 
                     Caption         =   "CheckCompleto"
                     Height          =   195
                     Left            =   650
                     TabIndex        =   18
                     Top             =   215
                     Width           =   170
                  End
                  Begin VB.CheckBox CheckExcedido 
                     Caption         =   "CheckExcedido"
                     Height          =   195
                     Left            =   650
                     TabIndex        =   17
                     Top             =   425
                     Width           =   170
                  End
                  Begin VB.CheckBox CheckIncompleto 
                     Caption         =   "CheckIncompleto"
                     Height          =   195
                     Left            =   650
                     TabIndex        =   16
                     Top             =   635
                     Width           =   170
                  End
                  Begin VB.CheckBox CheckEstado 
                     Caption         =   "CheckEstado"
                     Height          =   195
                     Left            =   300
                     TabIndex        =   15
                     Top             =   215
                     Width           =   170
                  End
                  Begin VB.Label Label4 
                     AutoSize        =   -1  'True
                     Caption         =   "Completo"
                     Height          =   195
                     Left            =   900
                     TabIndex        =   21
                     Top             =   210
                     Width           =   660
                  End
                  Begin VB.Label Label13 
                     AutoSize        =   -1  'True
                     Caption         =   "Excedido"
                     Height          =   195
                     Left            =   900
                     TabIndex        =   20
                     Top             =   420
                     Width           =   660
                  End
                  Begin VB.Label Label18 
                     AutoSize        =   -1  'True
                     Caption         =   "Incompleto"
                     Height          =   195
                     Left            =   900
                     TabIndex        =   19
                     Top             =   630
                     Width           =   780
                  End
               End
               Begin VB.Frame Frame11 
                  Caption         =   "Situacion"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   945
                  Left            =   60
                  TabIndex        =   6
                  Top             =   1050
                  Width           =   2400
                  Begin VB.CheckBox CheckSituacion 
                     Caption         =   "CheckSituacion"
                     Height          =   195
                     Left            =   300
                     TabIndex        =   10
                     Top             =   215
                     Width           =   170
                  End
                  Begin VB.CheckBox CheckADestiempo 
                     Caption         =   "CheckADestiempo"
                     Height          =   195
                     Left            =   650
                     TabIndex        =   9
                     Top             =   635
                     Width           =   170
                  End
                  Begin VB.CheckBox CheckAntesDeTiempo 
                     Caption         =   "CheckAntesDeTiempo"
                     Height          =   195
                     Left            =   650
                     TabIndex        =   8
                     Top             =   425
                     Width           =   170
                  End
                  Begin VB.CheckBox CheckATiempo 
                     Caption         =   "CheckATiempo"
                     Height          =   195
                     Left            =   650
                     TabIndex        =   7
                     Top             =   215
                     Width           =   170
                  End
                  Begin VB.Label Label19 
                     AutoSize        =   -1  'True
                     Caption         =   "A Destiempo"
                     Height          =   195
                     Left            =   900
                     TabIndex        =   13
                     Top             =   630
                     Width           =   900
                  End
                  Begin VB.Label Label21 
                     AutoSize        =   -1  'True
                     Caption         =   "Antes de Tiermpo"
                     Height          =   195
                     Left            =   900
                     TabIndex        =   12
                     Top             =   420
                     Width           =   1245
                  End
                  Begin VB.Label Label22 
                     AutoSize        =   -1  'True
                     Caption         =   "A Tiempo"
                     Height          =   195
                     Left            =   900
                     TabIndex        =   11
                     Top             =   210
                     Width           =   675
                  End
               End
            End
         End
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   20000
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   345
         Left            =   0
         TabIndex        =   2
         Top             =   0
         Width           =   20000
         _ExtentX        =   35269
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
                  Picture         =   "FrmConsPedido.frx":00F4
                  Key             =   "IMG1"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConsPedido.frx":0638
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConsPedido.frx":09CA
                  Key             =   "IMG2"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConsPedido.frx":0B4E
                  Key             =   "IMG3"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConsPedido.frx":0FA2
                  Key             =   "IMG4"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConsPedido.frx":10BA
                  Key             =   "IMG5"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConsPedido.frx":15FE
                  Key             =   "IMG6"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConsPedido.frx":1B42
                  Key             =   "IMG7"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConsPedido.frx":1C56
                  Key             =   "IMG8"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConsPedido.frx":1D6A
                  Key             =   "IMG9"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConsPedido.frx":21BE
                  Key             =   "IMG10"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConsPedido.frx":232A
                  Key             =   "IMG11"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConsPedido.frx":2872
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FrmConsPedido.frx":2B8C
                  Key             =   ""
               EndProperty
            EndProperty
         End
      End
   End
End
Attribute VB_Name = "FrmConsPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstLis As New ADODB.Recordset

Dim presionado As Boolean
Dim cargo As Boolean
Dim limpio As Boolean

Dim cambioTextEmsion1 As Boolean
Dim cambioTextPlazo1 As Boolean

Dim fechaEmision1 As String
Dim fechaEmision2 As String
Dim fechaPlazo1 As String
Dim fechaPlazo2 As String

Dim estadoCheckEstado As Boolean
Dim estadoCheckSituacion As Boolean
Dim estadoCheckSituacion2 As Boolean

Private Sub generarConsulta(tipo As Integer)
    Select Case tipo
        Case 0
            generarConsultaEntregados
        Case 1
            generarConsultaNoEntregados
    End Select
End Sub

Private Sub generarConsultaEntregados()
    Dim A As Integer
    Dim fila As Integer
    
    Dim c_PRODUCTOS As String
    Dim c_CLIENTES As String
    Dim c_OC As String
    
    Dim c_FECHA_EMI As String
    Dim c_FECHA_PLAZ As String
    
    Dim c_ESTADO As String
    Dim c_SITUACION As String
        
    Dim c_FECHA_ENT As String
    
    Dim c_SQL As String
    
    limpiarVSFlexGridReporte 0
    
    'Consulta para Clientes
    VSFlexGridClientes.Row = 0
    VSFlexGridClientes.Col = 1
    c_CLIENTES = "((Tabla001.nombre)= '" + VSFlexGridClientes.Text + "'"
    If (VSFlexGridClientes.TextMatrix(0, 1) = "Todos") Then
        c_CLIENTES = ""
    Else
        For A = 0 To VSFlexGridClientes.Rows - 1
            VSFlexGridClientes.Row = A
            VSFlexGridClientes.Col = 1
            c_CLIENTES = c_CLIENTES + " OR " + "(Tabla001.nombre)= '" + VSFlexGridClientes.Text + "'"
        Next A
        c_CLIENTES = c_CLIENTES + ") AND "
    End If
    'Consulta para Ordenes de Pedido
    VSFlexGridOC.Row = 0
    VSFlexGridOC.Col = 1
    c_OC = "((ped_pedido.oc)= '" + VSFlexGridOC.Text + "'"
    If (VSFlexGridOC.TextMatrix(0, 1) = "Todos") Then
        c_OC = ""
    Else
        For A = 0 To VSFlexGridOC.Rows - 1
            VSFlexGridOC.Row = A
            VSFlexGridOC.Col = 1
            c_OC = c_OC + " OR " + "(ped_pedido.oc)= '" + VSFlexGridOC.Text + "'"
        Next A
        c_OC = c_OC + ") AND "
    End If
    'Consulta para Fecha de Emision
    c_FECHA_EMI = "((Tabla001.fEmi)>=CDate('" & TextBoxFechaEmision1.Valor & "')"
    c_FECHA_EMI = c_FECHA_EMI & " AND (Tabla001.fEmi)<=CDate('" & TextBoxFechaEmision2.Valor & "'))"
    'Consulta para Fecha a Entregar
    c_FECHA_PLAZ = "((Tabla001.fAEnt)>=CDate('" & TextBoxFechaPlazo1.Valor & "')"
    c_FECHA_PLAZ = c_FECHA_PLAZ & " AND (Tabla001.fAEnt)<=CDate('" & TextBoxFechaPlazo2.Valor & "'))"
        
    If CheckEstado.Value = 1 Then
        c_ESTADO = ""
    Else
        Dim eCompleto As String
        Dim eExcedido As String
        Dim einCompleto As String
        
        If CheckCompleto.Value = 1 Then
            eCompleto = "(IIf([Tabla002].[cantEnt]<[Tabla001].[cantAEnt],'INCOMPLETO',IIf([Tabla002].[cantEnt]>[Tabla001].[cantAEnt],'EXCEDIDO','COMPLETO'))='COMPLETO')"
        Else
            eCompleto = ""
        End If
        If CheckExcedido.Value = 1 Then
            If CheckCompleto.Value = 0 Then
                eExcedido = "(IIf([Tabla002].[cantEnt]<[Tabla001].[cantAEnt],'INCOMPLETO',IIf([Tabla002].[cantEnt]>[Tabla001].[cantAEnt],'EXCEDIDO','COMPLETO'))='EXCEDIDO')"
            Else
                eExcedido = " OR (IIf([Tabla002].[cantEnt]<[Tabla001].[cantAEnt],'INCOMPLETO',IIf([Tabla002].[cantEnt]>[Tabla001].[cantAEnt],'EXCEDIDO','COMPLETO'))='EXCEDIDO') "
            End If
        Else
            eExcedido = ""
        End If
        If CheckIncompleto.Value = 1 Then
            If CheckExcedido.Value = 0 And CheckCompleto.Value = 0 Then
                einCompleto = "(IIf([Tabla002].[cantEnt]<[Tabla001].[cantAEnt],'INCOMPLETO',IIf([Tabla002].[cantEnt]>[Tabla001].[cantAEnt],'EXCEDIDO','COMPLETO'))='INCOMPLETO')"
            Else
                einCompleto = " OR (IIf([Tabla002].[cantEnt]<[Tabla001].[cantAEnt],'INCOMPLETO',IIf([Tabla002].[cantEnt]>[Tabla001].[cantAEnt],'EXCEDIDO','COMPLETO'))='INCOMPLETO')"
            End If
        Else
            einCompleto = ""
        End If
        c_ESTADO = "(" & eCompleto & eExcedido & einCompleto & ") AND "
    End If
    
    If CheckSituacion.Value = 1 Then
        c_SITUACION = ""
    Else
        Dim eATiempo As String
        Dim eADestiempo As String
        Dim eAntesDeTiempo As String
        
        If CheckATiempo.Value = 1 Then
            eATiempo = "(IIf([Tabla002].[fEnt]>[Tabla001].[fAEnt],'A DESTIEMPO',IIf([Tabla002].[fEnt]<[Tabla001].[fAEnt],'ANTES DE TIEMPO','A TIEMPO'))='A TIEMPO')"
        Else
            eATiempo = ""
        End If
        If CheckADestiempo.Value = 1 Then
            If CheckATiempo.Value = 0 Then
                eADestiempo = "(IIf([Tabla002].[fEnt]>[Tabla001].[fAEnt],'A DESTIEMPO',IIf([Tabla002].[fEnt]<[Tabla001].[fAEnt],'ANTES DE TIEMPO','A TIEMPO'))='A DESTIEMPO')"
            Else
                eADestiempo = " OR (IIf([Tabla002].[fEnt]>[Tabla001].[fAEnt],'A DESTIEMPO',IIf([Tabla002].[fEnt]<[Tabla001].[fAEnt],'ANTES DE TIEMPO','A TIEMPO'))='A DESTIEMPO')"
            End If
        Else
            eADestiempo = ""
        End If
        If CheckAntesDeTiempo.Value = 1 Then
            If CheckADestiempo.Value = 0 And CheckATiempo.Value = 0 Then
                eAntesDeTiempo = "(IIf([Tabla002].[fEnt]>[Tabla001].[fAEnt],'A DESTIEMPO',IIf([Tabla002].[fEnt]<[Tabla001].[fAEnt],'ANTES DE TIEMPO','A TIEMPO'))='ANTES DE TIEMPO')"
            Else
                eAntesDeTiempo = " OR (IIf([Tabla002].[fEnt]>[Tabla001].[fAEnt],'A DESTIEMPO',IIf([Tabla002].[fEnt]<[Tabla001].[fAEnt],'ANTES DE TIEMPO','A TIEMPO'))='ANTES DE TIEMPO')"
            End If
        Else
            eAntesDeTiempo = ""
        End If
        c_SITUACION = "(" & eATiempo & eADestiempo & eAntesDeTiempo & ") AND "
    End If
    'Se hace la consulta general
    c_SQL = "SELECT Tabla001.oc, Tabla001.nombre, Tabla001.fEmi, Tabla001.fAEnt, Tabla001.cantAEnt, Tabla002.FEnt, Tabla002.cantEnt, IIf([Tabla002].[cantEnt]<[Tabla001].[cantAEnt],'INCOMPLETO',IIf([Tabla002].[cantEnt]>[Tabla001].[cantAEnt],'EXCEDIDO','COMPLETO')) AS estado, IIf([Tabla002].[fEnt]>[Tabla001].[fAEnt],'A DESTIEMPO',IIf([Tabla002].[fEnt]<[Tabla001].[fAEnt],'ANTES DE TIEMPO','A TIEMPO')) AS situacion " _
            + vbCr + "FROM [SELECT ped_pedido.oc, mae_cliente.nombre, Min(ped_pedido.fchemi) AS fEmi, Max(ped_pedidodetent.fchent) AS fAEnt, Sum(ped_pedidodetent.canpro) AS cantAEnt FROM (ped_pedidodetent LEFT JOIN ped_pedido ON ped_pedidodetent.idped = ped_pedido.id) LEFT JOIN mae_cliente ON ped_pedido.idcli = mae_cliente.id GROUP BY ped_pedido.oc, mae_cliente.nombre ]. AS Tabla001, [SELECT  vta_guia.numordcom,Max(vta_guia.fchentord) AS FEnt, Sum(vta_guiadet.canpro) AS cantEnt FROM vta_guiadet INNER JOIN vta_guia ON vta_guiadet.idgui = vta_guia.id GROUP BY vta_guia.numordcom ]. AS Tabla002 " _
            + vbCr + "WHERE (((Tabla001.oc) = [Tabla002].[numordcom])) " _
            + vbCr + "GROUP BY Tabla001.oc, Tabla001.nombre, Tabla001.fEmi, Tabla001.fAEnt, Tabla001.cantAEnt, Tabla002.FEnt, Tabla002.cantEnt, IIf([Tabla002].[cantEnt]<[Tabla001].[cantAEnt],'INCOMPLETO',IIf([Tabla002].[cantEnt]>[Tabla001].[cantAEnt],'EXCEDIDO','COMPLETO')), IIf([Tabla002].[fEnt]>[Tabla001].[fAEnt],'A DESTIEMPO',IIf([Tabla002].[fEnt]<[Tabla001].[fAEnt],'ANTES DE TIEMPO','A TIEMPO')) " _
            + vbCr + "HAVING (" & c_CLIENTES & c_OC & c_ESTADO & c_SITUACION & c_FECHA_EMI & " And " & c_FECHA_PLAZ & "AND ((Tabla001.nombre)<>'') AND ((Tabla001.oc)<>''))" _
            + vbCr + "ORDER BY Tabla001.fAEnt;"
            
    RST_Busq RstLis, c_SQL, xCon
    Set VSFlexGridReporte.DataSource = RstLis.DataSource
    configurarVSFlexGridReporte 0
    Set RstLis = Nothing
End Sub

Private Sub generarConsultaNoEntregados()
    Dim A As Integer
    Dim fila As Integer
    
    Dim c_PRODUCTOS As String
    Dim c_CLIENTES As String
    Dim c_OC As String
    
    Dim c_FECHA_EMI As String
    Dim c_FECHA_PLAZ As String
    
    'Dim c_ESTADO As String
    Dim c_SITUACION As String
        
    Dim c_FECHA_ENT As String
    
    Dim c_SQL As String
    
    limpiarVSFlexGridReporte 1
    
    'Consulta para Clientes
    VSFlexGridClientes2.Row = 0
    VSFlexGridClientes2.Col = 1
    c_CLIENTES = "((Tabla001.nombre)= '" + VSFlexGridClientes2.Text + "'"
    If (VSFlexGridClientes2.TextMatrix(0, 1) = "Todos") Then
        c_CLIENTES = ""
    Else
        For A = 0 To VSFlexGridClientes2.Rows - 1
            VSFlexGridClientes2.Row = A
            VSFlexGridClientes2.Col = 1
            c_CLIENTES = c_CLIENTES + " OR " + "(Tabla001.nombre)= '" + VSFlexGridClientes2.Text + "'"
        Next A
        c_CLIENTES = c_CLIENTES + ") AND "
    End If
    'Consulta para Ordenes de Pedido
    VSFlexGridOC2.Row = 0
    VSFlexGridOC2.Col = 1
    c_OC = "((ped_pedido.oc)= '" + VSFlexGridOC2.Text + "'"
    If (VSFlexGridOC2.TextMatrix(0, 1) = "Todos") Then
        c_OC = ""
    Else
        For A = 0 To VSFlexGridOC2.Rows - 1
            VSFlexGridOC2.Row = A
            VSFlexGridOC2.Col = 1
            c_OC = c_OC + " OR " + "(ped_pedido.oc)= '" + VSFlexGridOC2.Text + "'"
        Next A
        c_OC = c_OC + ") AND "
    End If
    'Consulta para Fecha de Emision
    c_FECHA_EMI = "((Tabla001.fEmi)>=CDate('" & TextBoxFechaEmision1_2.Valor & "')"
    c_FECHA_EMI = c_FECHA_EMI & " AND (Tabla001.fEmi)<=CDate('" & TextBoxFechaEmision2_2.Valor & "'))"
    'Consulta para Fecha a Entregar
    c_FECHA_PLAZ = "((Tabla001.fAEnt)>=CDate('" & TextBoxFechaPlazo1_2.Valor & "')"
    c_FECHA_PLAZ = c_FECHA_PLAZ & " AND (Tabla001.fAEnt)<=CDate('" & TextBoxFechaPlazo2_2.Valor & "'))"
    
    If CheckSituacion2.Value = 1 Then
        c_SITUACION = ""
    Else
        Dim eATiempo As String
        Dim eADestiempo As String
        Dim eAntesDeTiempo As String
        
        If CheckATiempo2.Value = 1 Then
            eATiempo = "(IIf([Tabla001].[fAEnt]> " & Date & ",'A DESTIEMPO',IIf([Tabla001].[fAEnt]< " & Date & ",'ANTES DE TIEMPO','A TIEMPO'))='A TIEMPO')"
        Else
            eATiempo = ""
        End If
        If CheckADestiempo2.Value = 1 Then
            If CheckATiempo2.Value = 0 Then
                eADestiempo = "(IIf([Tabla001].[fAEnt]> " & Date & ",'A DESTIEMPO',IIf([Tabla001].[fAEnt]< " & Date & ",'ANTES DE TIEMPO','A TIEMPO'))='A DESTIEMPO')"
            Else
                eADestiempo = " OR (IIf([Tabla001].[fAEnt]> " & Date & ",'A DESTIEMPO',IIf([Tabla001].[fAEnt]< " & Date & ",'ANTES DE TIEMPO','A TIEMPO'))='A DESTIEMPO')"
            End If
        Else
            eADestiempo = ""
        End If
        If CheckAntesDeTiempo2.Value = 1 Then
            If CheckADestiempo2.Value = 0 And CheckATiempo2.Value = 0 Then
                eAntesDeTiempo = "(IIf([Tabla001].[fAEnt]> " & Date & ",'A DESTIEMPO',IIf([Tabla001].[fAEnt]< " & Date & ",'ANTES DE TIEMPO','A TIEMPO'))='ANTES DE TIEMPO')"
            Else
                eAntesDeTiempo = " OR (IIf([Tabla001].[fAEnt]> " & Date & ",'A DESTIEMPO',IIf([Tabla001].[fAEnt]< " & Date & ",'ANTES DE TIEMPO','A TIEMPO'))='ANTES DE TIEMPO')"
            End If
        Else
            eAntesDeTiempo = ""
        End If
        c_SITUACION = "(" & eATiempo & eADestiempo & eAntesDeTiempo & ") AND "
    End If
    'Se hace la consulta general
    
'SELECT Tabla001.oc, Tabla001.nombre, Tabla001.fEmi, Tabla001.fAEnt, Tabla001.cantAEnt, IIf([Tabla001].[fAEnt]>27/12/10,'A DESTIEMPO',IIf([Tabla001].[fAEnt]<27/12/10,'ANTES DE TIEMPO','A TIEMPO')) AS situacion
'FROM [SELECT ped_pedido.oc, mae_cliente.nombre, Min(ped_pedido.fchemi) AS fEmi, Max(ped_pedidodetent.fchent) AS fAEnt, Sum(ped_pedidodetent.canpro) AS cantAEnt FROM (ped_pedidodetent LEFT JOIN ped_pedido ON ped_pedidodetent.idped = ped_pedido.id) LEFT JOIN mae_cliente ON ped_pedido.idcli = mae_cliente.id GROUP BY ped_pedido.oc, mae_cliente.nombre ]. AS Tabla001, [SELECT  vta_guia.numordcom,Max(vta_guia.fchentord) AS FEnt, Sum(vta_guiadet.canpro) AS cantEnt FROM vta_guiadet INNER JOIN vta_guia ON vta_guiadet.idgui = vta_guia.id GROUP BY vta_guia.numordcom ]. AS Tabla002
'Where (((Tabla001.oc) <> [Tabla002].[numordcom]))
'GROUP BY Tabla001.oc, Tabla001.nombre, Tabla001.fEmi, Tabla001.fAEnt, Tabla001.cantAEnt, IIf([Tabla001].[fAEnt]>27/12/10,'A DESTIEMPO',IIf([Tabla001].[fAEnt]<27/12/10,'ANTES DE TIEMPO','A TIEMPO'))
'HAVING (((Tabla001.oc)<>'') AND ((Tabla001.nombre)<>'') AND ((Tabla001.fEmi)>=CDate('01/01/2010') And (Tabla001.fEmi)<=CDate('27/12/2010')) AND ((Tabla001.fAEnt)>=CDate('01/12/2010') And (Tabla001.fAEnt)<=CDate('26/05/2011')))
'ORDER BY Tabla001.fAEnt;

    c_SQL = "SELECT Tabla001.oc, Tabla001.nombre, Tabla001.fEmi, Tabla001.fAEnt, Tabla001.cantAEnt, IIf([Tabla001].[fAEnt]> " & Date & ",'A DESTIEMPO',IIf([Tabla001].[fAEnt]< " & Date & ",'ANTES DE TIEMPO','A TIEMPO')) AS situacion " _
            + vbCr + "FROM [SELECT ped_pedido.oc, mae_cliente.nombre, Min(ped_pedido.fchemi) AS fEmi, Max(ped_pedidodetent.fchent) AS fAEnt, Sum(ped_pedidodetent.canpro) AS cantAEnt FROM (ped_pedidodetent LEFT JOIN ped_pedido ON ped_pedidodetent.idped = ped_pedido.id) LEFT JOIN mae_cliente ON ped_pedido.idcli = mae_cliente.id GROUP BY ped_pedido.oc, mae_cliente.nombre ]. AS Tabla001, [SELECT  vta_guia.numordcom,Max(vta_guia.fchentord) AS FEnt, Sum(vta_guiadet.canpro) AS cantEnt FROM vta_guiadet INNER JOIN vta_guia ON vta_guiadet.idgui = vta_guia.id GROUP BY vta_guia.numordcom ]. AS Tabla002 " _
            + vbCr + "Where (((Tabla001.oc) <> [Tabla002].[numordcom])) " _
            + vbCr + "GROUP BY Tabla001.oc, Tabla001.nombre, Tabla001.fEmi, Tabla001.fAEnt, Tabla001.cantAEnt, IIf([Tabla001].[fAEnt]> " & Date & ",'A DESTIEMPO',IIf([Tabla001].[fAEnt]< " & Date & ",'ANTES DE TIEMPO','A TIEMPO')) " _
            + vbCr + "HAVING (" & c_CLIENTES & c_OC & c_SITUACION & c_FECHA_EMI & " And " & c_FECHA_PLAZ & "AND ((Tabla001.nombre)<>'') AND ((Tabla001.oc)<>''))" _
            + vbCr + "ORDER BY Tabla001.fAEnt;"
            
    RST_Busq RstLis, c_SQL, xCon
    Set VSFlexGridReporte2.DataSource = RstLis.DataSource
    configurarVSFlexGridReporte 1
    Set RstLis = Nothing
End Sub

Private Sub bloquear(tipo As Integer)
    Select Case tipo
        Case 0
            bloquearEntregados
        Case 1
            bloquearNoEntregados
    End Select
    
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
End Sub

Private Sub bloquearEntregados()
    VSFlexGridClientes.Enabled = Not VSFlexGridClientes.Enabled
    VSFlexGridOC.Enabled = Not VSFlexGridOC.Enabled
    
    TextBoxFechaEmision1.Enabled = Not TextBoxFechaEmision1.Enabled
    TextBoxFechaEmision2.Enabled = Not TextBoxFechaEmision2.Enabled
    TextBoxFechaPlazo1.Enabled = Not TextBoxFechaPlazo1.Enabled
    TextBoxFechaPlazo2.Enabled = Not TextBoxFechaPlazo2.Enabled
    
    CheckEstado.Enabled = Not CheckEstado.Enabled
    CheckCompleto.Enabled = Not CheckCompleto.Enabled
    CheckExcedido.Enabled = Not CheckExcedido.Enabled
    CheckIncompleto.Enabled = Not CheckIncompleto.Enabled
    
    CheckSituacion.Enabled = Not CheckSituacion.Enabled
    CheckATiempo.Enabled = Not CheckATiempo.Enabled
    CheckAntesDeTiempo.Enabled = Not CheckAntesDeTiempo.Enabled
    CheckADestiempo.Enabled = Not CheckADestiempo.Enabled
End Sub

Private Sub bloquearNoEntregados()
    VSFlexGridClientes2.Enabled = Not VSFlexGridClientes.Enabled
    VSFlexGridOC2.Enabled = Not VSFlexGridOC.Enabled
    
    TextBoxFechaEmision1_2.Enabled = Not TextBoxFechaEmision1.Enabled
    TextBoxFechaEmision2_2.Enabled = Not TextBoxFechaEmision2.Enabled
    TextBoxFechaPlazo1_2.Enabled = Not TextBoxFechaPlazo1.Enabled
    TextBoxFechaPlazo2_2.Enabled = Not TextBoxFechaPlazo2.Enabled
    
    CheckSituacion2.Enabled = Not CheckSituacion.Enabled
    CheckATiempo2.Enabled = Not CheckATiempo.Enabled
    CheckAntesDeTiempo2.Enabled = Not CheckAntesDeTiempo.Enabled
    CheckADestiempo2.Enabled = Not CheckADestiempo.Enabled
End Sub

Private Sub mostrarDetalle(tipo As Integer, fil As Integer)
    Select Case tipo
        Case 1
            mostrarDetalleEntregados fil
        Case 2
            mostrarDetalleNoEntregados fil
    End Select
End Sub

Private Sub mostrarDetalleEntregados(fil As Integer)
    Dim c_SQL As String
    Dim c_SQL2 As String
    Dim RstLisAux As New ADODB.Recordset
    Dim RstLisAux2 As New ADODB.Recordset
    Dim prodAct As String
    Dim numOrdAct As String

    Dim cantPed As Double
    Dim cantEnt As Double
    
    bloquear 0
    limpiarVSFlexGridDetalle 0

    numOrdAct = VSFlexGridReporte.TextMatrix(fil, 1)
    prodAct = VSFlexGridReporte.TextMatrix(fil, 4)

    'Se hace la consulta para pedidos
    c_SQL = "SELECT alm_inventario.descripcion, ped_pedido.fchemi, ped_pedidodetent.fchent, ped_pedidodetent.canpro " _
            + vbCr + "FROM (ped_pedidodetent LEFT JOIN ped_pedido ON ped_pedidodetent.idped = ped_pedido.id) LEFT JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id " _
            + vbCr + "WHERE (((ped_pedido.oc)= '" & numOrdAct & "')) " _
            + vbCr + "ORDER BY ped_pedidodetent.iditem;"

    RST_Busq RstLisAux, c_SQL, xCon
    Set VSFlexGridDetalle2.DataSource = RstLisAux.DataSource

    'Se hace la consulta para entregas
    c_SQL2 = "SELECT alm_inventario.descripcion, vta_guia.fchentord, vta_guiadet.canpro " _
            + vbCr + "FROM (vta_guiadet INNER JOIN vta_guia ON vta_guiadet.idgui = vta_guia.id) INNER JOIN alm_inventario ON vta_guiadet.iditem = alm_inventario.id " _
            + vbCr + "WHERE (((vta_guia.numordcom)= '" & numOrdAct & "')) " _
            + vbCr + "ORDER BY vta_guiadet.iditem, vta_guia.fchentord;"

    RST_Busq RstLisAux2, c_SQL2, xCon
    Set VSFlexGridDetalle.DataSource = RstLisAux2.DataSource
    
    If (RstLisAux2.EOF) Then
        VSFlexGridDetalle.AddItem ("")
        VSFlexGridDetalle.TextMatrix(1, 3) = "No hay Informacion"
        LabelDetalleRetraso = VSFlexGridReporte.TextMatrix(fil, 8)
        LabelDetalleSituacion = VSFlexGridReporte.TextMatrix(fil, 9)

        cantPed = 0
        cantEnt = 0
        RstLisAux.MoveFirst
        While (Not RstLisAux.EOF)
            cantPed = cantPed + RstLisAux("canpro")
            RstLisAux.MoveNext
        Wend
        LabelDetalleEntgr = cantPed
        LabelDetalleEntgda = cantEnt
    Else
        cantPed = 0
        cantEnt = 0
        RstLisAux.MoveFirst
        RstLisAux2.MoveFirst
        While (Not RstLisAux.EOF Or Not RstLisAux2.EOF)
            If (Not RstLisAux.EOF) Then
                cantPed = cantPed + RstLisAux("canpro")
                RstLisAux.MoveNext
            End If
            If (Not RstLisAux2.EOF) Then
                cantEnt = cantEnt + RstLisAux2("canpro")
                RstLisAux2.MoveNext
            End If
        Wend
        LabelDetalleEntgr = cantPed
        LabelDetalleEntgda = cantEnt
        RstLisAux.MoveLast
        RstLisAux2.MoveLast
        LabelDetalleRetraso = VSFlexGridReporte.TextMatrix(fil, 8)
        LabelDetalleSituacion = VSFlexGridReporte.TextMatrix(fil, 9)
    End If

    RstLisAux.MoveFirst

    configurarVSFlexGridDetalle 0
    LabelDetalleOC = VSFlexGridReporte.TextMatrix(fil, 1)
    LabelResto = LabelDetalleEntgda - LabelDetalleEntgr

    Set RstLisAux = Nothing
    Set RstLisAux2 = Nothing
    c_SQL = ""
    c_SQL2 = ""
End Sub

Private Sub mostrarDetalleNoEntregados(fil As Integer)
    Dim c_SQL As String
    Dim c_SQL2 As String
    Dim RstLisAux As New ADODB.Recordset
    Dim RstLisAux2 As New ADODB.Recordset
    Dim prodAct As String
    Dim numOrdAct As String

    Dim cantPed As Double
    Dim cantEnt As Double
    
    bloquear 1
    limpiarVSFlexGridDetalle 1

    numOrdAct = VSFlexGridReporte2.TextMatrix(fil, 1)

    'Se hace la consulta para pedidos
    c_SQL = "SELECT alm_inventario.descripcion, ped_pedido.fchemi, ped_pedidodetent.fchent, ped_pedidodetent.canpro " _
            + vbCr + "FROM (ped_pedidodetent LEFT JOIN ped_pedido ON ped_pedidodetent.idped = ped_pedido.id) LEFT JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id " _
            + vbCr + "WHERE (((ped_pedido.oc)= '" & numOrdAct & "')) " _
            + vbCr + "ORDER BY ped_pedidodetent.iditem;"

    RST_Busq RstLisAux, c_SQL, xCon
    Set VSFlexGridDetalle2_2.DataSource = RstLisAux.DataSource
    cantPed = 0
    cantEnt = 0
    RstLisAux.MoveFirst
    While (Not RstLisAux.EOF)
        If (Not RstLisAux.EOF) Then
            cantPed = cantPed + RstLisAux("canpro")
            RstLisAux.MoveNext
        End If
    Wend
    LabelDetalleEntgr2 = cantPed
    RstLisAux.MoveLast
    RstLisAux.MoveFirst

    configurarVSFlexGridDetalle 1
    LabelDetalleSituacion2 = VSFlexGridReporte2.TextMatrix(fil, 6)
    LabelDetalleOC2 = VSFlexGridReporte2.TextMatrix(fil, 1)
    'LabelResto = LabelDetalleEntgda - LabelDetalleEntgr

    Set RstLisAux = Nothing
    'Set RstLisAux2 = Nothing
    c_SQL = ""
    'c_SQL2 = ""
End Sub

Private Sub iniciarCampos()
    Dim fechAct As Date
    fechAct = Date
    presionado = False
    Set VSFlexGridClientes.DataSource = Nothing
    Set VSFlexGridClientes2.DataSource = Nothing
    Set VSFlexGridOC.DataSource = Nothing
    Set VSFlexGridOC2.DataSource = Nothing
    
    Set VSFlexGridDetalle.DataSource = Nothing
    Set VSFlexGridDetalle2.DataSource = Nothing
    Set VSFlexGridDetalle2_2.DataSource = Nothing
    'Se inicializa:
    'datos para clientes
    VSFlexGridClientes.Rows = 1
    VSFlexGridClientes.Cols = 2
    VSFlexGridClientes.Row = 0
    VSFlexGridClientes.Col = 1
    VSFlexGridClientes.Text = "Todos"
    
    VSFlexGridClientes2.Rows = 1
    VSFlexGridClientes2.Cols = 2
    VSFlexGridClientes2.Row = 0
    VSFlexGridClientes2.Col = 1
    VSFlexGridClientes2.Text = "Todos"
    'datos para Ordenes de Compra
    VSFlexGridOC.Rows = 1
    VSFlexGridOC.Cols = 2
    VSFlexGridOC.Row = 0
    VSFlexGridOC.Col = 1
    VSFlexGridOC.Text = "Todos"
    
    VSFlexGridOC2.Rows = 1
    VSFlexGridOC2.Cols = 2
    VSFlexGridOC2.Row = 0
    VSFlexGridOC2.Col = 1
    VSFlexGridOC2.Text = "Todos"
    'datos para detalles
    TextBoxFechaEmision1.Valor = CDate("01/01/" + CStr(Year(Date)))
    TextBoxFechaEmision2.Valor = Date
    TextBoxFechaPlazo1.Valor = CDate("01/" + CStr(Month(Date)) + "/" + CStr(Year(Date)))
    TextBoxFechaPlazo2.Valor = Date + 150
    
    TextBoxFechaEmision1_2.Valor = CDate("01/01/" + CStr(Year(Date)))
    TextBoxFechaEmision2_2.Valor = Date
    TextBoxFechaPlazo1_2.Valor = CDate("01/" + CStr(Month(Date)) + "/" + CStr(Year(Date)))
    TextBoxFechaPlazo2_2.Valor = Date + 150
    
    VSFlexGridClientes.Editable = flexEDKbdMouse
    VSFlexGridClientes.ColComboList(1) = "..."
    VSFlexGridClientes.ShowComboButton = flexSBAlways
    
    VSFlexGridClientes2.Editable = flexEDKbdMouse
    VSFlexGridClientes2.ColComboList(1) = "..."
    VSFlexGridClientes2.ShowComboButton = flexSBAlways
    
    VSFlexGridOC.Editable = flexEDKbdMouse
    VSFlexGridOC.ColComboList(1) = "..."
    VSFlexGridOC.ShowComboButton = flexSBAlways
    
    VSFlexGridOC2.Editable = flexEDKbdMouse
    VSFlexGridOC2.ColComboList(1) = "..."
    VSFlexGridOC2.ShowComboButton = flexSBAlways
    
    CheckEstado.Value = 1
    CheckSituacion.Value = 1
    CheckSituacion2.Value = 1
    
    FrameEspecificaciones.Top = -30
    FrameEspecificaciones.Left = 10
    FrameEspecificaciones.Width = 11850
    FrameEspecificaciones.Height = 3950
    
    FrameEspecificaciones2.Top = -30
    FrameEspecificaciones2.Left = 10
    FrameEspecificaciones2.Width = 11850
    FrameEspecificaciones2.Height = 3950
    
    VSFlexGridReporte.AllowUserResizing = flexResizeColumns
    VSFlexGridReporte.AutoSearch = flexSearchFromTop
    VSFlexGridReporte.ExplorerBar = flexExSortShowAndMove
    
    VSFlexGridReporte2.AllowUserResizing = flexResizeColumns
    VSFlexGridReporte2.AutoSearch = flexSearchFromTop
    VSFlexGridReporte2.ExplorerBar = flexExSortShowAndMove
    
    VSFlexGridDetalle2.AllowUserResizing = flexResizeColumns
    VSFlexGridDetalle2.AutoSearch = flexSearchFromTop
    VSFlexGridDetalle2.ExplorerBar = flexExSortShowAndMove
    
    VSFlexGridDetalle2_2.AllowUserResizing = flexResizeColumns
    VSFlexGridDetalle2_2.AutoSearch = flexSearchFromTop
    VSFlexGridDetalle2_2.ExplorerBar = flexExSortShowAndMove
    
    VSFlexGridDetalle.AllowUserResizing = flexResizeColumns
    VSFlexGridDetalle.AutoSearch = flexSearchFromTop
    VSFlexGridDetalle.ExplorerBar = flexExSortShowAndMove
    
    TabOne1.CurrTab = 0
End Sub

Private Sub Command2_Click()
    iniciarCampos
End Sub

Private Sub configurarVSFlexGridReporte(tipo As Integer)
    Select Case tipo
        Case 0
            configurarVSFlexGridReporteEntregados
        Case 1
            configurarVSFlexGridReporteNoEntregados
    End Select
End Sub

Private Sub configurarVSFlexGridReporteEntregados()
    VSFlexGridReporte.Cols = 10
    VSFlexGridReporte.FixedRows = 1
    VSFlexGridReporte.FixedCols = 3
    VSFlexGridReporte.ColWidth(0) = 250
    VSFlexGridReporte.RowHeight(0) = 400
    VSFlexGridReporte.Row = 0
    
    VSFlexGridReporte.Col = 1
    VSFlexGridReporte.Text = "ORD. PED."
    VSFlexGridReporte.ColWidth(1) = 1200
    VSFlexGridReporte.CellFontBold = True
    
    VSFlexGridReporte.Col = 2
    VSFlexGridReporte.Text = "CLIENTE"
    VSFlexGridReporte.ColWidth(3) = 2300
    VSFlexGridReporte.CellFontBold = True
    
    VSFlexGridReporte.Col = 3
    VSFlexGridReporte.Text = "FECH. EMISION"
    VSFlexGridReporte.ColWidth(3) = 1950
    VSFlexGridReporte.CellFontBold = True
    
    VSFlexGridReporte.Col = 4
    VSFlexGridReporte.Text = "FECH. A ENTREGAR"
    VSFlexGridReporte.ColWidth(4) = 1950
    VSFlexGridReporte.CellFontBold = True
    
    VSFlexGridReporte.Col = 5
    VSFlexGridReporte.Text = "CANT. A ENTREGAR"
    VSFlexGridReporte.ColWidth(5) = 1950
    VSFlexGridReporte.CellFontBold = True
    
    VSFlexGridReporte.Col = 6
    VSFlexGridReporte.Text = "FECH. ENTREGADA"
    VSFlexGridReporte.ColWidth(6) = 1950
    VSFlexGridReporte.CellFontBold = True
    
    VSFlexGridReporte.Col = 7
    VSFlexGridReporte.Text = "CANT. ENTREGADA"
    VSFlexGridReporte.ColWidth(7) = 1950
    VSFlexGridReporte.CellFontBold = True
    
    VSFlexGridReporte.Col = 8
    VSFlexGridReporte.Text = "ESTADO"
    VSFlexGridReporte.ColWidth(8) = 1950
    VSFlexGridReporte.CellFontBold = True
    
    VSFlexGridReporte.Col = 9
    VSFlexGridReporte.Text = "SITUACION"
    VSFlexGridReporte.ColWidth(9) = 1950
    VSFlexGridReporte.CellFontBold = True
'
End Sub

Private Sub configurarVSFlexGridReporteNoEntregados()
    VSFlexGridReporte2.Cols = 7
    VSFlexGridReporte2.FixedRows = 1
    VSFlexGridReporte2.FixedCols = 3
    VSFlexGridReporte2.ColWidth(0) = 250
    VSFlexGridReporte2.RowHeight(0) = 400
    VSFlexGridReporte2.Row = 0
    
    VSFlexGridReporte2.Col = 1
    VSFlexGridReporte2.Text = "ORD. PED."
    VSFlexGridReporte2.ColWidth(1) = 1200
    VSFlexGridReporte2.CellFontBold = True
    
    VSFlexGridReporte2.Col = 2
    VSFlexGridReporte2.Text = "CLIENTE"
    VSFlexGridReporte2.ColWidth(3) = 2300
    VSFlexGridReporte2.CellFontBold = True
    
    VSFlexGridReporte2.Col = 3
    VSFlexGridReporte2.Text = "FECH. EMISION"
    VSFlexGridReporte2.ColWidth(3) = 1950
    VSFlexGridReporte2.CellFontBold = True
    
    VSFlexGridReporte2.Col = 4
    VSFlexGridReporte2.Text = "FECH. A ENTREGAR"
    VSFlexGridReporte2.ColWidth(4) = 1950
    VSFlexGridReporte2.CellFontBold = True
    
    VSFlexGridReporte2.Col = 5
    VSFlexGridReporte2.Text = "CANT. A ENTREGAR"
    VSFlexGridReporte2.ColWidth(5) = 1950
    VSFlexGridReporte2.CellFontBold = True
    
    VSFlexGridReporte2.Col = 6
    VSFlexGridReporte2.Text = "SITUACION"
    VSFlexGridReporte2.ColWidth(6) = 1950
    VSFlexGridReporte2.CellFontBold = True
End Sub

Private Sub configurarVSFlexGridDetalle(tipo As Integer)
    Select Case tipo
        Case 0
            configurarVSFlexGridDetalleEntregados
        Case 1
            configurarVSFlexGridDetalleNoEntregados
    End Select
End Sub

Private Sub configurarVSFlexGridDetalleEntregados()
    'detalle 1
    VSFlexGridDetalle2.Cols = 5
    VSFlexGridDetalle2.FixedRows = 1
    VSFlexGridDetalle2.FixedCols = 1
    VSFlexGridDetalle2.ColWidth(0) = 0
    VSFlexGridDetalle2.RowHeight(0) = 300
    VSFlexGridDetalle2.Row = 0
    
    VSFlexGridDetalle2.Col = 1
    VSFlexGridDetalle2.Text = "PRODUCTO"
    VSFlexGridDetalle2.ColWidth(1) = 2610
    
    VSFlexGridDetalle2.Col = 2
    VSFlexGridDetalle2.Text = "FECH. EMISION"
    VSFlexGridDetalle2.ColWidth(2) = 0
    
    VSFlexGridDetalle2.Col = 3
    VSFlexGridDetalle2.Text = "FECH. A ENTREGAR"
    VSFlexGridDetalle2.ColWidth(3) = 1650
    
    VSFlexGridDetalle2.Col = 4
    VSFlexGridDetalle2.Text = "CANT. A ENTREGAR"
    VSFlexGridDetalle2.ColWidth(4) = 1650
    
    'detalle 2
    VSFlexGridDetalle.Cols = 4
    VSFlexGridDetalle.FixedRows = 1
    VSFlexGridDetalle.FixedCols = 0
    VSFlexGridDetalle.ColWidth(0) = 0
    VSFlexGridDetalle.RowHeight(0) = 300
    VSFlexGridDetalle.Row = 0
    
    VSFlexGridDetalle.Col = 1
    VSFlexGridDetalle.Text = "PRODUCTO"
    VSFlexGridDetalle.ColWidth(1) = 2300
    
    VSFlexGridDetalle.Col = 2
    VSFlexGridDetalle.Text = "FECH. ENTREGADA"
    VSFlexGridDetalle.ColWidth(2) = 1600
    
    VSFlexGridDetalle.Col = 3
    VSFlexGridDetalle.Text = "CANT. ENTREGADA"
    VSFlexGridDetalle.ColWidth(3) = 1600
End Sub

Private Sub configurarVSFlexGridDetalleNoEntregados()
    'detalle 1
    VSFlexGridDetalle2_2.Cols = 5
    VSFlexGridDetalle2_2.FixedRows = 1
    VSFlexGridDetalle2_2.FixedCols = 1
    VSFlexGridDetalle2_2.ColWidth(0) = 0
    VSFlexGridDetalle2_2.RowHeight(0) = 300
    VSFlexGridDetalle2_2.Row = 0
    
    VSFlexGridDetalle2_2.Col = 1
    VSFlexGridDetalle2_2.Text = "PRODUCTO"
    VSFlexGridDetalle2_2.ColWidth(1) = 5300
    
    VSFlexGridDetalle2_2.Col = 2
    VSFlexGridDetalle2_2.Text = "FECH. EMISION"
    VSFlexGridDetalle2_2.ColWidth(2) = 1900
    
    VSFlexGridDetalle2_2.Col = 3
    VSFlexGridDetalle2_2.Text = "FECH. A ENTREGAR"
    VSFlexGridDetalle2_2.ColWidth(3) = 2000
    
    VSFlexGridDetalle2_2.Col = 4
    VSFlexGridDetalle2_2.Text = "CANT. A ENTREGAR"
    VSFlexGridDetalle2_2.ColWidth(4) = 2000
End Sub

Private Sub limpiarVSFlexGridReporte(tipo As Integer)
    Select Case tipo
        Case 0
            limpiarVSFlexGridReporteEntregados
        Case 1
            limpiarVSFlexGridReporteNoEntregados
    End Select
End Sub

Private Sub limpiarVSFlexGridReporteEntregados()
    VSFlexGridReporte.Cols = 2
    VSFlexGridReporte.Rows = 2
    VSFlexGridReporte.FixedRows = 1
    VSFlexGridReporte.FixedCols = 1
End Sub

Private Sub limpiarVSFlexGridReporteNoEntregados()
    VSFlexGridReporte2.Cols = 2
    VSFlexGridReporte2.Rows = 2
    VSFlexGridReporte2.FixedRows = 1
    VSFlexGridReporte2.FixedCols = 1
End Sub

Private Sub limpiarVSFlexGridDetalle(tipo As Integer)
    Select Case tipo
        Case 0
            limpiarVSFlexGridDetalleEntregados
        Case 1
            limpiarVSFlexGridDetalleNoEntregados
    End Select
End Sub

Private Sub limpiarVSFlexGridDetalleEntregados()
    VSFlexGridDetalle.Cols = 2
    VSFlexGridDetalle.Rows = 2
    VSFlexGridDetalle.FixedRows = 1
    VSFlexGridDetalle.FixedCols = 1
    
    VSFlexGridDetalle2.Cols = 2
    VSFlexGridDetalle2.Rows = 2
    VSFlexGridDetalle2.FixedRows = 1
    VSFlexGridDetalle2.FixedCols = 1
End Sub

Private Sub limpiarVSFlexGridDetalleNoEntregados()
    VSFlexGridDetalle.Cols = 2
    VSFlexGridDetalle.Rows = 2
    VSFlexGridDetalle.FixedRows = 1
    VSFlexGridDetalle.FixedCols = 1
    
    VSFlexGridDetalle2.Cols = 2
    VSFlexGridDetalle2.Rows = 2
    VSFlexGridDetalle2.FixedRows = 1
    VSFlexGridDetalle2.FixedCols = 1
End Sub

Private Sub verficarCheckSituacion(tipo As Integer)
    Select Case tipo
        Case 0
            verficarCheckSituacionEntregados
        Case 1
            verficarCheckSituacionNoEntregados
    End Select
End Sub

Private Sub verficarCheckSituacionEntregados()
    If CheckADestiempo.Value = 1 And CheckAntesDeTiempo.Value = 1 And CheckATiempo.Value = 1 Then CheckSituacion.Value = 1
    If CheckADestiempo.Value = 0 Or CheckAntesDeTiempo.Value = 0 Or CheckATiempo.Value = 0 Then CheckSituacion.Value = 0
    If CheckADestiempo.Value = 0 And CheckAntesDeTiempo.Value = 0 And CheckATiempo.Value = 0 Then CheckSituacion.Value = 1
End Sub

Private Sub verficarCheckSituacionNoEntregados()
    If CheckADestiempo2.Value = 1 And CheckAntesDeTiempo2.Value = 1 And CheckATiempo2.Value = 1 Then CheckSituacion2.Value = 1
    If CheckADestiempo2.Value = 0 Or CheckAntesDeTiempo2.Value = 0 Or CheckATiempo2.Value = 0 Then CheckSituacion2.Value = 0
    If CheckADestiempo2.Value = 0 And CheckAntesDeTiempo2.Value = 0 And CheckATiempo2.Value = 0 Then CheckSituacion2.Value = 1
End Sub

Private Sub CheckADestiempo_Click()
    verficarCheckSituacion 0
End Sub

Private Sub CheckADestiempo2_Click()
    verficarCheckSituacion 1
End Sub

Private Sub CheckAntesDeTiempo_Click()
    verficarCheckSituacion 0
End Sub

Private Sub CheckAntesDeTiempo2_Click()
    verficarCheckSituacion 1
End Sub

Private Sub CheckATiempo_Click()
    verficarCheckSituacion 0
End Sub

Private Sub verficarCheckEstado()
    If CheckIncompleto.Value = 1 And CheckExcedido.Value = 1 And CheckCompleto.Value = 1 Then CheckEstado.Value = 1
    If CheckIncompleto.Value = 0 Or CheckExcedido.Value = 0 Or CheckCompleto.Value = 0 Then CheckEstado.Value = 0
    If CheckIncompleto.Value = 0 And CheckExcedido.Value = 0 And CheckCompleto.Value = 0 Then CheckEstado.Value = 1
End Sub

Private Sub CheckATiempo2_Click()
    verficarCheckSituacion 1
End Sub

Private Sub CheckCompleto_Click()
    verficarCheckEstado
End Sub

Private Sub CheckExcedido_Click()
    verficarCheckEstado
End Sub

Private Sub CheckIncompleto_Click()
    verficarCheckEstado
End Sub

Private Sub CheckEstado_Click()
    If estadoCheckEstado Then
        estadoCheckEstado = False
    Else
        estadoCheckEstado = True
        CheckCompleto.Value = 1
        CheckIncompleto.Value = 1
        CheckExcedido.Value = 1
    End If
End Sub

Private Sub CheckSituacion_Click()
    If estadoCheckSituacion Then
        estadoCheckSituacion = False
    Else
        estadoCheckSituacion = True
        CheckATiempo.Value = 1
        CheckAntesDeTiempo.Value = 1
        CheckADestiempo.Value = 1
    End If
End Sub

Private Sub CheckSituacion2_Click()
    If estadoCheckSituacion2 Then
        estadoCheckSituacion2 = False
    Else
        estadoCheckSituacion2 = True
        CheckATiempo2.Value = 1
        CheckAntesDeTiempo2.Value = 1
        CheckADestiempo2.Value = 1
    End If
End Sub

Private Sub CommandLimpiar_Click()
    limpio = True
    iniciarCampos
    limpio = False
End Sub

Private Sub Form_Activate()
    generarConsulta 0
End Sub

Private Sub Form_Load()
    
    limpio = False
    cargo = False
    cambioTextPlazo1 = False
    iniciarCampos
End Sub

Private Sub LabelCerrar_Click()
    FrameEspecificaciones.Visible = False
    bloquear 0
End Sub

Private Sub LabelCerrar2_Click()
    FrameEspecificaciones2.Visible = False
    bloquear 1
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If Not cargo Then
        If OldTab = 0 Then generarConsulta 1: cargo = True
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        If TabOne1.CurrTab = 0 Then
            generarConsulta 0
        Else
            generarConsulta 1
        End If
    End If
    If Button.Index = 5 Then
        If TabOne1.CurrTab = 0 Then
            ExportarExcel 0
        Else
            ExportarExcel 1
        End If
    End If
    If Button.Index = 8 Then
        Unload Me
    End If
End Sub

Private Sub VSFlexGridClientes_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    VSFlexGridClientes.ShowComboButton = flexSBFocus
    
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
        VSFlexGridClientes.TextMatrix(VSFlexGridClientes.Row, 1) = xRs.Fields(0) & ""
        If VSFlexGridClientes.Row = VSFlexGridClientes.Rows - 1 Then VSFlexGridClientes.AddItem ""
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub VSFlexGridClientes_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        If VSFlexGridClientes.Rows = 0 Then VSFlexGridClientes.AddItem ("Todos")
    End If
    If KeyCode = 46 Then
        On Error GoTo MAY
        VSFlexGridClientes.RemoveItem VSFlexGridClientes.Row
        If VSFlexGridClientes.Rows = 0 Then VSFlexGridClientes.AddItem (""): VSFlexGridClientes.TextMatrix(0, 1) = "Todos"
    End If
    Exit Sub
MAY:
    Exit Sub
End Sub

Private Sub VSFlexGridClientes2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    VSFlexGridClientes2.ShowComboButton = flexSBFocus
    
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
        VSFlexGridClientes2.TextMatrix(VSFlexGridClientes2.Row, 1) = xRs.Fields(0) & ""
        If VSFlexGridClientes2.Row = VSFlexGridClientes2.Rows - 1 Then VSFlexGridClientes2.AddItem ""
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub VSFlexGridClientes2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        If VSFlexGridClientes2.Rows = 0 Then VSFlexGridClientes2.AddItem ("Todos")
    End If
    If KeyCode = 46 Then
        On Error GoTo MAY
        VSFlexGridClientes2.RemoveItem VSFlexGridClientes2.Row
        If VSFlexGridClientes2.Rows = 0 Then VSFlexGridClientes2.AddItem (""): VSFlexGridClientes2.TextMatrix(0, 1) = "Todos"
    End If
    Exit Sub
MAY:
    Exit Sub
End Sub

Private Sub VSFlexGridOC_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    VSFlexGridOC.ShowComboButton = flexSBFocus
    
    Dim nSQL As String
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 3) As String
    
    xCampos(0, 0) = "Orden de Compra":    xCampos(0, 1) = "oc":       xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id Pedido":          xCampos(1, 1) = "idped":    xCampos(1, 2) = "2500":   xCampos(1, 3) = "C"

    nSQL = "SELECT DISTINCT ped_pedido.oc, ped_pedidodetent.idped " _
         + vbCr + "FROM ped_pedidodetent RIGHT JOIN ped_pedido ON ped_pedidodetent.idped = ped_pedido.id " _
         + vbCr + "GROUP BY ped_pedido.oc, ped_pedidodetent.idped " _
         + vbCr + "Having (((ped_pedido.oc) Is Not Null And (ped_pedido.oc) <> 'S/N' And (ped_pedido.oc) <> '') And ((ped_pedidodetent.idped) Is Not Null)) " _
         + vbCr + "ORDER BY ped_pedido.oc;"
         
    xform.SQLCad = nSQL
    
    xform.Titulo = "Buscando Ordenes de Compra"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "oc"
    xform.CampoBusca = "oc"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        VSFlexGridOC.TextMatrix(VSFlexGridOC.Row, 1) = xRs.Fields(0) & ""
        If VSFlexGridOC.Row = VSFlexGridOC.Rows - 1 Then VSFlexGridOC.AddItem ""
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub VSFlexGridOC_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        If VSFlexGridOC.Rows = 0 Then VSFlexGridOC.AddItem ("")
    End If
    If KeyCode = 46 Then
        On Error GoTo MAY
        VSFlexGridOC.RemoveItem VSFlexGridOC.Row
        If VSFlexGridOC.Rows = 0 Then VSFlexGridOC.AddItem (""): VSFlexGridOC.TextMatrix(0, 1) = "Todos"
    End If
    Exit Sub
MAY:
    Exit Sub
End Sub

Private Sub VSFlexGridOC2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    VSFlexGridOC2.ShowComboButton = flexSBFocus
    
    Dim nSQL As String
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 3) As String
    
    xCampos(0, 0) = "Orden de Compra":    xCampos(0, 1) = "oc":       xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Id Pedido":          xCampos(1, 1) = "idped":    xCampos(1, 2) = "2500":   xCampos(1, 3) = "C"

    nSQL = "SELECT DISTINCT ped_pedido.oc, ped_pedidodetent.idped " _
         + vbCr + "FROM ped_pedidodetent RIGHT JOIN ped_pedido ON ped_pedidodetent.idped = ped_pedido.id " _
         + vbCr + "GROUP BY ped_pedido.oc, ped_pedidodetent.idped " _
         + vbCr + "Having (((ped_pedido.oc) Is Not Null And (ped_pedido.oc) <> 'S/N' And (ped_pedido.oc) <> '') And ((ped_pedidodetent.idped) Is Not Null)) " _
         + vbCr + "ORDER BY ped_pedido.oc;"
         
    xform.SQLCad = nSQL
    
    xform.Titulo = "Buscando Ordenes de Compra"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "oc"
    xform.CampoBusca = "oc"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        VSFlexGridOC2.TextMatrix(VSFlexGridOC2.Row, 1) = xRs.Fields(0) & ""
        If VSFlexGridOC2.Row = VSFlexGridOC2.Rows - 1 Then VSFlexGridOC2.AddItem ""
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub VSFlexGridOC2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        If VSFlexGridOC2.Rows = 0 Then VSFlexGridOC2.AddItem ("")
    End If
    If KeyCode = 46 Then
        On Error GoTo MAY
        VSFlexGridOC2.RemoveItem VSFlexGridOC2.Row
        If VSFlexGridOC2.Rows = 0 Then VSFlexGridOC2.AddItem (""): VSFlexGridOC2.TextMatrix(0, 1) = "Todos"
    End If
    Exit Sub
MAY:
    Exit Sub
End Sub

Sub ExportarExcel(tipo As Integer)
    Select Case tipo
        Case 0
            ExportarExcelEntregados
        Case 1
            ExportarExcelNoEntregados
    End Select
End Sub

Sub ExportarExcelEntregados()
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
        .cells(2, 3) = "'" + TextBoxFechaEmision1.Valor
        .cells(2, 4) = "Hasta: "
        .cells(2, 5) = "'" + TextBoxFechaEmision2.Valor
        .cells(3, 2) = "Fecha a Entregar del Pedido   Desde: "
        .cells(3, 3) = "'" + TextBoxFechaPlazo1.Valor
        .cells(3, 4) = "Hasta: "
        .cells(3, 5) = "'" + TextBoxFechaPlazo2.Valor
        
        .cells(4, 2) = "Clientes: "
        xFilas = 5
        For A = 0 To VSFlexGridClientes.Rows - 1
            .cells(xFilas, 3) = VSFlexGridClientes.TextMatrix(A, 1)
            xFilas = xFilas + 1
        Next A
        xFilas = xFilas + 1
        For A = 0 To VSFlexGridReporte.Rows - 1
            For B = 1 To VSFlexGridReporte.Cols - 1
                If A = 0 Then
                    .cells(xFilas, B + 1) = "'" + VSFlexGridReporte.TextMatrix(A, B)
                Else
                    If (B = 1 Or B = 3 Or B = 7) Then
                        .cells(xFilas, B + 1) = NulosN(VSFlexGridReporte.TextMatrix(A, B))
                    Else
                        .cells(xFilas, B + 1) = "'" + VSFlexGridReporte.TextMatrix(A, B)
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

Sub ExportarExcelNoEntregados()
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
        .cells(2, 3) = "'" + TextBoxFechaEmision1_2.Valor
        .cells(2, 4) = "Hasta: "
        .cells(2, 5) = "'" + TextBoxFechaEmision2_2.Valor
        .cells(3, 2) = "Fecha a Entregar del Pedido   Desde: "
        .cells(3, 3) = "'" + TextBoxFechaPlazo1_2.Valor
        .cells(3, 4) = "Hasta: "
        .cells(3, 5) = "'" + TextBoxFechaPlazo2_2.Valor
        
        .cells(4, 2) = "Clientes: "
        xFilas = 5
        For A = 0 To VSFlexGridClientes2.Rows - 1
            .cells(xFilas, 3) = VSFlexGridClientes2.TextMatrix(A, 1)
            xFilas = xFilas + 1
        Next A
        xFilas = xFilas + 1
        For A = 0 To VSFlexGridReporte2.Rows - 1
            For B = 1 To VSFlexGridReporte2.Cols - 1
                If A = 0 Then
                    .cells(xFilas, B + 1) = "'" + VSFlexGridReporte2.TextMatrix(A, B)
                Else
                    If (B = 1 Or B = 3 Or B = 7) Then
                        .cells(xFilas, B + 1) = NulosN(VSFlexGridReporte2.TextMatrix(A, B))
                    Else
                        .cells(xFilas, B + 1) = "'" + VSFlexGridReporte2.TextMatrix(A, B)
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

Private Sub VSFlexGridReporte_DblClick()
    If (VSFlexGridReporte.Row <> 0) Then
        mostrarDetalle 1, VSFlexGridReporte.Row
        FrameEspecificaciones.Visible = True
    End If
End Sub

Private Sub VSFlexGridReporte2_DblClick()
    If (VSFlexGridReporte2.Row <> 0) Then
        mostrarDetalle 2, VSFlexGridReporte2.Row
        FrameEspecificaciones2.Visible = True
    End If
End Sub
