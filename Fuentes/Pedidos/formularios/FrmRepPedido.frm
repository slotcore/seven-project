VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmRepPedido 
   Caption         =   "Ventas  -  Reporte de Pedidos"
   ClientHeight    =   8355
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13095
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   8355
   ScaleWidth      =   13095
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.Frame FraProgreso 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   945
      Left            =   690
      TabIndex        =   37
      Top             =   4470
      Visible         =   0   'False
      Width           =   5940
      Begin MSComctlLib.ProgressBar PgBar 
         Height          =   255
         Left            =   225
         TabIndex        =   38
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
         Index           =   0
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
         TabIndex        =   41
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
         TabIndex        =   40
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label lbl 
         Alignment       =   1  'Right Justify
         Caption         =   "No Interrumpir"
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
         Left            =   4170
         TabIndex        =   39
         Top             =   180
         Width           =   1530
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   850
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   13065
      Begin VB.CommandButton Cmd 
         Caption         =   "Consultar"
         Height          =   330
         Index           =   0
         Left            =   11760
         TabIndex        =   30
         ToolTipText     =   "Eliminar Todos"
         Top             =   510
         Visible         =   0   'False
         Width           =   1275
      End
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
         Height          =   885
         Left            =   9990
         TabIndex        =   14
         Top             =   -30
         Width           =   1755
         Begin VB.OptionButton OptTipo 
            Caption         =   "Vista Historica"
            Height          =   195
            Index           =   2
            Left            =   60
            TabIndex        =   36
            Top             =   210
            Width           =   1365
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Vista Detallada"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   35
            Top             =   420
            Width           =   1395
         End
         Begin VB.OptionButton OptTipo 
            Caption         =   "Vista Simple"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   34
            Top             =   630
            Width           =   1185
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
         Left            =   8100
         TabIndex        =   9
         Top             =   -30
         Width           =   1875
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEntDesde 
            Height          =   300
            Left            =   555
            TabIndex        =   10
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEntHasta 
            Height          =   300
            Left            =   555
            TabIndex        =   11
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
         Begin VB.Label Label27 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   45
            TabIndex        =   13
            Top             =   255
            Width           =   510
         End
         Begin VB.Label Label26 
            AutoSize        =   -1  'True
            Caption         =   "Hasta:"
            Height          =   195
            Left            =   45
            TabIndex        =   12
            Top             =   555
            Width           =   465
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
         Left            =   6210
         TabIndex        =   4
         Top             =   -30
         Width           =   1875
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmiDesde 
            Height          =   300
            Left            =   555
            TabIndex        =   5
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
            Left            =   555
            TabIndex        =   6
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
            Left            =   45
            TabIndex        =   8
            Top             =   585
            Width           =   465
         End
         Begin VB.Label Label29 
            AutoSize        =   -1  'True
            Caption         =   "Desde:"
            Height          =   195
            Left            =   45
            TabIndex        =   7
            Top             =   255
            Width           =   510
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   800
         Index           =   3
         Left            =   30
         TabIndex        =   31
         ToolTipText     =   "Buscar Clientes"
         Top             =   75
         Width           =   2175
         _cx             =   3836
         _cy             =   1411
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
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRepPedido.frx":0000
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
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   795
         Index           =   4
         Left            =   2220
         TabIndex        =   32
         ToolTipText     =   "Buscar Productos"
         Top             =   60
         Width           =   2415
         _cx             =   4260
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
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRepPedido.frx":0047
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
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   795
         Index           =   5
         Left            =   4630
         TabIndex        =   33
         ToolTipText     =   "Buscar Ordenes de Pedido"
         Top             =   60
         Width           =   1575
         _cx             =   2778
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
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRepPedido.frx":008F
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
   Begin VB.Frame Frame10 
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
      ForeColor       =   &H80000001&
      Height          =   4320
      Left            =   5430
      TabIndex        =   17
      Top             =   2190
      Visible         =   0   'False
      Width           =   6330
      Begin VB.Frame Frame8 
         Height          =   1275
         Left            =   90
         TabIndex        =   19
         Top             =   300
         Width           =   6135
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   4770
            TabIndex        =   27
            Top             =   870
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
            Locked          =   -1  'True
         End
         Begin VB.Label LblCliente 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblProd2"
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   1080
            TabIndex        =   29
            Top             =   510
            Width           =   4950
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Left            =   90
            TabIndex        =   28
            Top             =   525
            Width           =   480
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emision"
            Height          =   195
            Index           =   0
            Left            =   3390
            TabIndex        =   26
            Top             =   930
            Width           =   1260
         End
         Begin VB.Label lbl_cb_capt 
            AutoSize        =   -1  'True
            Caption         =   "Ord. Pedido"
            Height          =   195
            Index           =   9
            Left            =   90
            TabIndex        =   22
            Top             =   930
            Width           =   840
         End
         Begin VB.Label LblProd 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblProd2"
            ForeColor       =   &H000000FF&
            Height          =   300
            Left            =   1080
            TabIndex        =   21
            Top             =   150
            Width           =   4950
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "Producto"
            Height          =   195
            Left            =   90
            TabIndex        =   20
            Top             =   165
            Width           =   645
         End
         Begin VB.Label LblOrden 
            BackStyle       =   0  'Transparent
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTarea2"
            ForeColor       =   &H00800000&
            Height          =   300
            Left            =   1080
            TabIndex        =   23
            Top             =   870
            Width           =   1725
         End
      End
      Begin VB.PictureBox PbCerrar 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   3
         Left            =   6060
         Picture         =   "FrmRepPedido.frx":00D9
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   18
         ToolTipText     =   "Cerrar"
         Top             =   30
         Width           =   195
      End
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   2550
         Index           =   2
         Left            =   90
         TabIndex        =   24
         Top             =   1620
         Width           =   6120
         _cx             =   10795
         _cy             =   4498
         _ConvInfo       =   1
         Appearance      =   0
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
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRepPedido.frx":03C5
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
         Editable        =   2
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
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de Pedido"
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
         Left            =   105
         TabIndex        =   25
         Top             =   60
         Width           =   1530
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   5
         X1              =   30
         X2              =   6300
         Y1              =   4290
         Y2              =   4290
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   4
         X1              =   0
         X2              =   8295
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00C00000&
         BorderWidth     =   2
         Index           =   5
         X1              =   6300
         X2              =   6300
         Y1              =   0
         Y2              =   4290
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   255
         Index           =   2
         Left            =   30
         Top             =   30
         Width           =   6240
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   570
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   13095
      _ExtentX        =   23098
      _ExtentY        =   1005
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
         Left            =   11070
         Top             =   0
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
               Picture         =   "FrmRepPedido.frx":0492
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepPedido.frx":09D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepPedido.frx":0D68
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepPedido.frx":0EEC
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepPedido.frx":1340
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepPedido.frx":1458
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepPedido.frx":199C
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepPedido.frx":1EE0
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepPedido.frx":1FF4
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepPedido.frx":2108
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepPedido.frx":255C
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepPedido.frx":26C8
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepPedido.frx":2C10
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmRepPedido.frx":2F2A
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame6 
      BorderStyle     =   0  'None
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
      Height          =   7125
      Left            =   30
      TabIndex        =   1
      Top             =   1200
      Width           =   13020
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   4695
         Index           =   1
         Left            =   0
         TabIndex        =   15
         Top             =   2400
         Visible         =   0   'False
         Width           =   12750
         _cx             =   22490
         _cy             =   8281
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
         Rows            =   1
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRepPedido.frx":32BC
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
      Begin VSFlex7Ctl.VSFlexGrid fg 
         Height          =   2370
         Index           =   0
         Left            =   0
         TabIndex        =   2
         Top             =   45
         Width           =   12765
         _cx             =   22516
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
         Rows            =   1
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmRepPedido.frx":3474
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
      TabIndex        =   3
      Top             =   0
      Width           =   1365
   End
   Begin VB.Menu menu 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu00 
         Caption         =   "Insertar Item"
      End
      Begin VB.Menu separador 
         Caption         =   "-"
      End
      Begin VB.Menu menu01 
         Caption         =   "Eliminar Item"
      End
   End
End
Attribute VB_Name = "FrmRepPedido"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim cargo As Boolean
Dim cSQL As String
Dim RstResumido As New ADODB.Recordset
Dim RstDetallado As New ADODB.Recordset
Dim INDICE_ As Integer
Dim OrigFX As Long
Dim OrigFY As Long

Private Sub Buscar()
    Frame10.Visible = False
    ' Se verfican si esta correcta la informacion
    If Not verificarDatos Then Exit Sub
    
    cargo = True
    If OptTipo(2).Value = True Then ' Opcion Historico
        generarConsulta False, False, True
        fg(1).Top = 45
        fg(1).Left = 0
        fg(1).Width = Frame6.Width - 150
        fg(1).Height = Frame6.Height - 100
        fg(1).Visible = True
        fg(0).Visible = False
        Exit Sub
    End If
    If OptTipo(1).Value = True Then ' Opcion Detallado
        generarConsulta False, True
        fg(1).Top = 45
        fg(1).Left = 0
        fg(1).Width = Frame6.Width - 150
        fg(1).Height = Frame6.Height - 100
        fg(1).Visible = True
        fg(0).Visible = False
    Else                            ' Opcion Resumido
        generarConsulta True, False
        fg(0).Top = 45
        fg(0).Left = 0
        fg(1).Width = Frame6.Width - 150
        fg(1).Height = Frame6.Height - 100
        fg(0).Visible = True
        fg(1).Visible = False
    End If
End Sub

Private Function verificarDatos() As Boolean
    Dim VERIFICO_ As Boolean
    Dim MENSAJE_ As String
    
    VERIFICO_ = True
    If (TxtFchEmiDesde.Valor = "" Or TxtFchEmiHasta.Valor = "") _
                                        Or (CDate(TxtFchEmiHasta.Valor) < CDate(TxtFchEmiDesde.Valor)) Then
        MENSAJE_ = "Ingrese un valor adecuado para la Fecha de Emision"
        VERIFICO_ = False
    End If
    
    If (TxtFchEntDesde.Valor = "" Or TxtFchEntHasta.Valor = "") _
                                        Or (CDate(TxtFchEntHasta.Valor) < CDate(TxtFchEntDesde.Valor)) Then
        MENSAJE_ = "Ingrese un valor adecuado para la Fecha de Entrega"
        VERIFICO_ = False
    End If
    If Not VERIFICO_ Then MsgBox MENSAJE_, vbCritical + vbOKOnly, "Reporte de Pedidos"
    verificarDatos = VERIFICO_
End Function

Private Sub generarConsulta(RESUMIDO_ As Boolean, DETALLADO_ As Boolean, _
                                                    Optional HISTORICO_ As Boolean = False)
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    Dim c_PRODUCTOS As String
    Dim c_CLIENTES As String
    Dim c_OC As String
    
    If RESUMIDO_ Then
        'Consulta para Productos
        If (fg(4).Rows = 1) Then
            c_PRODUCTOS = ""
        Else
            For A = 1 To fg(4).Rows - 1
                If A = 1 Then
                    c_PRODUCTOS = "((alm_inventario.descripcion)= '" + fg(4).TextMatrix(A, 1) + "'"
                Else
                    c_PRODUCTOS = c_PRODUCTOS + " OR " + "(alm_inventario.descripcion)= '" & fg(4).TextMatrix(A, 1) & "'"
                End If
            Next A
            c_PRODUCTOS = c_PRODUCTOS + ") AND "
        End If
        'Consulta para Clientes
        If (fg(3).Rows = 1) Then
            c_CLIENTES = ""
        Else
            For A = 1 To fg(3).Rows - 1
                If A = 1 Then
                    c_CLIENTES = "((mae_cliente.nombre)= '" + fg(3).TextMatrix(A, 1) + "')"
                Else
                    c_CLIENTES = c_CLIENTES + " OR ((mae_cliente.nombre)= '" + fg(3).TextMatrix(A, 1) + "')"
                End If
            Next A
            c_CLIENTES = "(" + c_CLIENTES + ") AND "
        End If
        'Consulta para Ordenes de Pedido
        If (fg(5).Rows = 1) Then
            c_OC = ""
        Else
            For A = 1 To fg(5).Rows - 1
                If A = 1 Then
                    c_OC = "((ped_pedido.oc)= '" + fg(5).TextMatrix(A, 1) + "'"
                Else
                    c_OC = c_OC + " OR " + "(ped_pedido.oc)= '" + fg(5).TextMatrix(A, 1) + "'"
                End If
            Next A
            c_OC = c_OC + ") AND "
        End If
        
        cSQL = "SELECT ped_pedido.id AS idped, ped_pedido.oc AS numped, mae_cliente.nombre AS nomcli, ped_pedido.idcli, ped_pedidodet.iditem AS idpro, alm_inventario.descripcion AS despro, mae_unidades.abrev AS unimed, ped_pedido.fchemi, IIf([ped_pedidodet].[fchent] Is Not Null,[ped_pedidodet].[fchent],[ped_pedido].[fchent]) AS fchent, ped_pedidodet.canpro, IIf(IsNull([ped_pedido]![numser])=-1,[ped_pedido]![numdoc],[ped_pedido]![numser]+'-'+[ped_pedido]![numdoc]) AS numdocped " _
                + vbCr + "FROM (((ped_tipo RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN (vta_puntoVenta RIGHT JOIN (ped_pedido LEFT JOIN mae_condpago ON ped_pedido.idconpag = mae_condpago.id) ON vta_puntoVenta.id = ped_pedido.idpunvecli) ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) ON ped_tipo.id = ped_pedido.idtipped) LEFT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id " _
                + vbCr + "WHERE (((ped_pedido.anulado) = 0) And " & c_CLIENTES & c_OC & c_PRODUCTOS & "((ped_pedido.fchemi)>=CDate('" & NulosC(TxtFchEmiDesde.Valor) & "') And (ped_pedido.fchemi)<=CDate('" & NulosC(TxtFchEmiHasta.Valor) & "')) AND ((IIf([ped_pedidodet].[fchent] Is Not Null,[ped_pedidodet].[fchent],[ped_pedido].[fchent]))>=CDate('" & NulosC(TxtFchEntDesde.Valor) & "') And (IIf([ped_pedidodet].[fchent] Is Not Null,[ped_pedidodet].[fchent],[ped_pedido].[fchent]))<=CDate('" & NulosC(TxtFchEntHasta.Valor) & "'))) " _
                + vbCr + "GROUP BY ped_pedido.id, ped_pedido.oc, mae_cliente.nombre, ped_pedido.idcli, ped_pedidodet.iditem, alm_inventario.descripcion, mae_unidades.abrev, ped_pedido.fchemi, IIf([ped_pedidodet].[fchent] Is Not Null,[ped_pedidodet].[fchent],[ped_pedido].[fchent]), ped_pedidodet.canpro, IIf(IsNull([ped_pedido]![numser])=-1,[ped_pedido]![numdoc],[ped_pedido]![numser]+'-'+[ped_pedido]![numdoc]); " _
                + vbCr + "Union " _
                + vbCr + "SELECT ped_pedido.id AS idped, ped_pedido.oc AS numped, mae_cliente.nombre AS nomcli, ped_pedido.idcli, ped_pedidodetent.iditem AS idpro, alm_inventario.descripcion AS despro, mae_unidades.abrev AS unimed, ped_pedido.fchemi, ped_pedidodetent.fchent, ped_pedidodetent.canpro, IIf(IsNull([ped_pedido]![numser])=-1,[ped_pedido]![numdoc],[ped_pedido]![numser]+'-'+[ped_pedido]![numdoc]) AS numdocped " _
                + vbCr + "FROM (((ped_tipo RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN (vta_puntoVenta RIGHT JOIN (ped_pedido LEFT JOIN mae_condpago ON ped_pedido.idconpag = mae_condpago.id) ON vta_puntoVenta.id = ped_pedido.idpunvecli) ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) ON ped_tipo.id = ped_pedido.idtipped) LEFT JOIN ped_pedidodetent ON ped_pedido.id = ped_pedidodetent.idped) LEFT JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodetent.idunimed = mae_unidades.id " _
                + vbCr + "WHERE (((ped_pedido.anulado) = 0) And " & c_CLIENTES & c_OC & c_PRODUCTOS & "((ped_pedido.idtipped)=2) AND ((ped_pedido.fchemi)>=CDate('" & NulosC(TxtFchEmiDesde.Valor) & "') And (ped_pedido.fchemi)<=CDate('" & NulosC(TxtFchEmiHasta.Valor) & "')) AND ((ped_pedidodetent.fchent)>=CDate('" & NulosC(TxtFchEntDesde.Valor) & "') And (ped_pedidodetent.fchent)<=CDate('" & NulosC(TxtFchEntHasta.Valor) & "'))) " _
                + vbCr + "GROUP BY ped_pedido.id, ped_pedido.oc, mae_cliente.nombre, ped_pedido.idcli, ped_pedidodetent.iditem, alm_inventario.descripcion, mae_unidades.abrev, ped_pedido.fchemi, ped_pedidodetent.fchent, ped_pedidodetent.canpro, IIf(IsNull([ped_pedido]![numser])=-1,[ped_pedido]![numdoc],[ped_pedido]![numser]+'-'+[ped_pedido]![numdoc]);"
            
        RST_Busq xRs, cSQL, xCon
        
        ' Se genera el recordset
        llenarDefinirRST True, False, xRs
        ' se llenan los datos de la consulta
        LlenarDatos True, False
        ' se configura el grid segun consulta
        configurarGrid True, False
        
        Set xRs = Nothing
    End If
    
    If DETALLADO_ Then
        ' Se genera el recordset
        llenarDefinirRST False, True
        ' se llenan los datos de la consulta
        LlenarDatos False, True
        ' se configura el grid segun consulta
        configurarGrid False, True
    End If
    
    If HISTORICO_ Then
        Set xRs = Nothing
        fg(1).Rows = 2
        Set xRs = Nothing
        
        ' Se verifica el estado del recordset
        generarConsulta True, False ' Se genera el recordset resumido
        If RstResumido.RecordCount = 0 Then Exit Sub
        
        If xRs.State = 0 Then DEFINIR_RST_TMP xRs, RstResumido
        If xRs.RecordCount <> 0 Then limpiarRST xRs, True
        CARGAR_RST_TMP xRs, RstResumido
        
        ' Se ordena el recordset
        RstResumido.Sort = "idpro"
        xRs.Sort = "idpro"
        
        ' Se recorre el reporte
        RstResumido.MoveFirst
        XcentrarFrm FraProgreso
        FraProgreso.Visible = True
        PgBar.Min = 0
        PgBar.Max = RstResumido.RecordCount
        For A = 1 To RstResumido.RecordCount
            PgBar.Value = A
            FraProgreso.Refresh
            
            xRs.Filter = "idpro = " & NulosN(RstResumido("idpro")) & " And numped = " & NulosN(RstResumido("numped")) & ""
            If xRs.RecordCount = 0 Then GoTo SIGUIENTE
            
            ' Se escribe el titulo del producto
            fg(1).Rows = fg(1).Rows + 1
            
            fg(1).Select fg(1).Rows - 1, 1
            fg(1).CellForeColor = &HC00000
            fg(1).TextMatrix(fg(1).Rows - 1, 1) = xRs("despro")
            
            xRs.MoveFirst
            ' Se llena el detalle con historico
            llenarVentana xRs("idpro"), xRs("numped"), fg(1), , , , , , , True, 4, True, True, True
            ' Se limpia parte del Recordset Temporal
            limpiarRST xRs, False
            xRs.Filter = adFilterNone
SIGUIENTE:
            RstResumido.MoveNext
        Next A
        FraProgreso.Visible = False
        configurarGrid False, False, False, True
    End If
End Sub

Private Sub llenarDefinirRST(RESUMIDO_ As Boolean, DETALLADO_ As Boolean, Optional RSTORIG As ADODB.Recordset)
    If RESUMIDO_ Then
        ' Se crea el recordset si no esta creado
        If RstResumido.State = 0 Then DEFINIR_RST_TMP RstResumido, RSTORIG
        limpiarRST RstResumido, True ' Se limpia el Rst
        CARGAR_RST_TMP RstResumido, RSTORIG ' Se carga el Rst
    End If
    If DETALLADO_ Then
        Set RstDetallado = Nothing
        limpiarRST RstDetallado, True ' Se limpia el Rst
        
        generarConsulta True, False
        llenarDetallado
        RstDetallado.Filter = adFilterNone
    End If
End Sub

Private Sub limpiarRST(Rst As ADODB.Recordset, Optional TODO As Boolean = True)
    If Rst.State = 0 Then Exit Sub
    With Rst
        If TODO Then .Filter = adFilterNone
        If .RecordCount = 0 Then Exit Sub
        .MoveFirst
        While Not .EOF
            .Delete
            .MoveNext
        Wend
    End With
End Sub

Sub preparaRST(ByRef RST_ As ADODB.Recordset)
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(12, 3) As String

    xCampos(0, 0) = "numped":       xCampos(0, 1) = "D":      xCampos(0, 2) = "2"
    xCampos(1, 0) = "idcli":        xCampos(1, 1) = "D":      xCampos(1, 2) = "2"
    xCampos(2, 0) = "nomcli":       xCampos(2, 1) = "C":      xCampos(2, 2) = "100"
    xCampos(3, 0) = "idpro":        xCampos(3, 1) = "D":      xCampos(3, 2) = "2"
    xCampos(4, 0) = "nompro":       xCampos(4, 1) = "C":      xCampos(4, 2) = "100"
    xCampos(5, 0) = "cantotped":    xCampos(5, 1) = "D":      xCampos(5, 2) = "2"
    xCampos(6, 0) = "cantotent":    xCampos(6, 1) = "D":      xCampos(6, 2) = "2"
    xCampos(7, 0) = "fchemi":       xCampos(7, 1) = "C":      xCampos(7, 2) = "20"
    xCampos(8, 0) = "canultent":    xCampos(8, 1) = "D":      xCampos(8, 2) = "2"
    xCampos(9, 0) = "fchultent":    xCampos(9, 1) = "C":      xCampos(9, 2) = "20"
    xCampos(10, 0) = "canparcped":  xCampos(10, 1) = "D":     xCampos(10, 2) = "2"
    xCampos(11, 0) = "fchparcped":  xCampos(11, 1) = "C":     xCampos(11, 2) = "20"
    xCampos(12, 0) = "cantotaent":  xCampos(12, 1) = "D":     xCampos(12, 2) = "2"
    
    Set RST_ = xFun.CrearRstTMP(xCampos)
    RST_.Open
End Sub

Private Sub llenarDetallado(Optional TODO_ As Boolean = True, Optional PEDIDOS_ As Boolean = False, _
                                                            Optional ENTREGAS_ As Boolean = False, _
                                                            Optional NUMPEDIDO_ As Double = 0, _
                                                            Optional IDPRODUCTO_ As Double = 0)
    Dim xRs As New ADODB.Recordset
    Dim xRs1 As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    Dim A As Integer
    Dim CANTIDAD_ As Double
    Dim NUMPED_ As Double
    Dim IDPRO_ As Double
    
    NUMPED_ = NUMPEDIDO_
    IDPRO_ = IDPRODUCTO_
    
    If TODO_ Then
        ' Se muestran todos los pedidos
        RstResumido.Filter = adFilterNone
        ' Se verifica el estado del recordset
        If RstResumido.State = 0 Then Exit Sub
        If RstResumido.RecordCount = 0 Then Exit Sub
        ' Se carga el recordset auxiliar
        DEFINIR_RST_TMP xRs, RstResumido
        CARGAR_RST_TMP xRs, RstResumido
    
        preparaRST RstDetallado ' Se define el recordset de pedidos detallados
        
        RstResumido.MoveFirst
        
        XcentrarFrm FraProgreso
        Frame10.Visible = False
        FraProgreso.Visible = True
        FraProgreso.Refresh
        PgBar.Min = 0
        PgBar.Max = RstResumido.RecordCount
        
        For A = 1 To RstResumido.RecordCount
            PgBar.Value = A
            NUMPED_ = NulosN(RstResumido("numped"))
            IDPRO_ = NulosN(RstResumido("idpro"))
            xRs.Filter = "numped = " & NUMPED_ & " And idpro = " & IDPRO_ & ""
            
            If xRs.RecordCount <> 0 Then
                xRs.MoveFirst
                ' Se filtra el pedido especificado
                RstDetallado.Filter = "numped = " & NUMPED_ & " And idpro = " & IDPRO_ & ""
                
                ' Se ve si ya se evaluo ese pedido
                If RstDetallado.RecordCount = 0 Then
                    llenarDetallado False, True, False, NUMPED_, IDPRO_ ' Se llenan los pedidos
                    llenarDetallado False, False, True, NUMPED_, IDPRO_ ' Se llenan las entregas
                    ' Se llenan los Totales
                    RstDetallado("cantotaent") = NulosN(RstDetallado("canparcped")) - NulosN(RstDetallado("cantotent"))
                End If
            End If
            
            RstResumido.MoveNext
        Next A
        FraProgreso.Visible = False
    End If
    If PEDIDOS_ Then
        Set xRs1 = Nothing
    
        cSQL = "SELECT ped_pedido.oc, ped_pedido.fchemi, ped_pedidodet.fchent, ped_pedido.idcli, mae_cliente.nombre, ped_pedidodet.iditem, alm_inventario.descripcion, ped_pedidodet.canpro " _
            + vbCr + "FROM ((ped_pedido LEFT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN mae_cliente ON ped_pedido.idcli = mae_cliente.id) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id " _
            + vbCr + "Where (((ped_pedidodet.idItem) = " & IDPRO_ & ") And ((ped_pedido.oc) = '" & NUMPED_ & "') And ((ped_pedidodet.fchent) Is Not Null And (ped_pedidodet.fchent) > #1/1/2011#)) " _
            + vbCr + "GROUP BY ped_pedido.oc, ped_pedido.fchemi, ped_pedidodet.fchent, ped_pedido.idcli, mae_cliente.nombre, ped_pedidodet.iditem, alm_inventario.descripcion, ped_pedidodet.canpro; " _
            + vbCr + "Union " _
            + vbCr + "SELECT ped_pedido.oc, ped_pedido.fchemi, ped_pedidodetent.fchent, ped_pedido.idcli, mae_cliente.nombre, ped_pedidodetent.iditem, alm_inventario.descripcion, ped_pedidodetent.canpro " _
            + vbCr + "FROM ((ped_pedido LEFT JOIN mae_cliente ON ped_pedido.idcli = mae_cliente.id) LEFT JOIN ped_pedidodetent ON ped_pedido.id = ped_pedidodetent.idped) LEFT JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id " _
            + vbCr + "Where (((ped_pedidodetent.idItem) = " & IDPRO_ & ") And ((ped_pedido.oc) = '" & NUMPED_ & "') And ((ped_pedidodetent.fchent) Is Not Null And (ped_pedidodetent.fchent) > #1/1/2011#)) " _
            + vbCr + "GROUP BY ped_pedido.oc, ped_pedido.fchemi, ped_pedidodetent.fchent, ped_pedido.idcli, mae_cliente.nombre, ped_pedidodetent.iditem, alm_inventario.descripcion, ped_pedidodetent.canpro;"
        
        RST_Busq xRs1, cSQL, xCon
        xRs1.Sort = "fchent"
        
        If xRs1.State = 0 Then Exit Sub
        If xRs1.RecordCount = 0 Then Exit Sub
        
        RstDetallado.AddNew
        For A = 1 To xRs1.RecordCount
            CANTIDAD_ = CANTIDAD_ + NulosN(xRs1("canpro"))
            
            If CDate(xRs1("fchent")) <= CDate(TxtFchEntHasta.Valor) Then
                RstDetallado("canparcped") = CANTIDAD_
                RstDetallado("fchparcped") = NulosC(xRs1("fchent"))
            End If
            
            If A = xRs1.RecordCount Then
                RstDetallado("numped") = NulosN(xRs1("oc"))
                RstDetallado("idcli") = NulosN(xRs1("idcli"))
                RstDetallado("nomcli") = NulosC(xRs1("nombre"))
                RstDetallado("idpro") = NulosN(xRs1("iditem"))
                RstDetallado("nompro") = NulosC(xRs1("descripcion"))
                RstDetallado("fchemi") = NulosC(xRs1("fchemi"))
                RstDetallado("cantotped") = CANTIDAD_
                
                RstDetallado.Update
            End If
            xRs1.MoveNext
        Next A
    End If
    If ENTREGAS_ Then
        Set xRs2 = Nothing
        
        cSQL = "SELECT vta_guia.numordcom, vta_guiadet.iditem, vta_guia.idcli, vta_guia.fchentord, vta_guiadet.canpro " _
            + vbCr + "FROM vta_guia INNER JOIN vta_guiadet ON vta_guia.id = vta_guiadet.idgui " _
            + vbCr + "Where (((vta_guia.numordcom) = '" & NUMPED_ & "') And ((vta_guiadet.idItem) = " & IDPRO_ & ")) " _
            + vbCr + "GROUP BY vta_guia.numordcom, vta_guiadet.iditem, vta_guia.idcli, vta_guia.fchentord, vta_guiadet.canpro;"
        
        RST_Busq xRs2, cSQL, xCon
        xRs2.Sort = "fchentord"
        
        If xRs2.State = 0 Then Exit Sub
        If xRs2.RecordCount = 0 Then Exit Sub
        
        ' Se filtra el pedido especificado
        RstDetallado.Filter = "numped = " & NUMPED_ & " And idpro = " & IDPRO_ & ""
        
        If RstDetallado.RecordCount = 0 Then Exit Sub
        
        For A = 1 To xRs2.RecordCount
            CANTIDAD_ = CANTIDAD_ + NulosN(xRs2("canpro"))
                        
            If A = xRs2.RecordCount Then
                RstDetallado("cantotent") = CANTIDAD_
                RstDetallado("canultent") = NulosN(xRs2("canpro"))
                RstDetallado("fchultent") = NulosC(xRs2("fchentord"))
                RstDetallado.Update
            End If
            xRs2.MoveNext
        Next A
    End If
End Sub

Private Sub llenarVentana(IDPRO_ As Double, NUMPED_ As Double, ByRef FgGrid As VSFlexGrid, _
                                    Optional PEDIDO_ As Boolean = True, _
                                    Optional ENTREGA_ As Boolean = False, _
                                    Optional TOTAL_ As Boolean = False, _
                                    Optional ByRef FILAPED_ As Double, _
                                    Optional ByRef SUMAPED_ As Double, _
                                    Optional ByRef SUMAGUIA_ As Double, _
                                    Optional HISTORIAL As Boolean = False, _
                                    Optional COLUMNAINICIO_ As Double = 1, _
                                    Optional MOSTRARCLIENTE_ As Boolean = False, _
                                    Optional MOSTRARFECHAEMISION_ As Boolean = False, _
                                    Optional MOSTRARORDEN_ As Boolean = False)
    Dim xRs As New ADODB.Recordset
    Dim A As Integer
    
    If PEDIDO_ Then
        Set xRs = Nothing
        cSQL = "SELECT ped_pedido.oc, ped_pedido.fchemi, ped_pedidodet.fchent, ped_pedido.idcli, mae_cliente.nombre, ped_pedidodet.iditem, alm_inventario.descripcion, ped_pedidodet.canpro " _
                + vbCr + "FROM ((ped_pedido LEFT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN mae_cliente ON ped_pedido.idcli = mae_cliente.id) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id " _
                + vbCr + "Where (((ped_pedidodet.idItem) = " & IDPRO_ & ") And ((ped_pedido.oc) = '" & NUMPED_ & "') And ((ped_pedidodet.fchent) Is Not Null And (ped_pedidodet.fchent) > #1/1/2011#)) " _
                + vbCr + "GROUP BY ped_pedido.oc, ped_pedido.fchemi, ped_pedidodet.fchent, ped_pedido.idcli, mae_cliente.nombre, ped_pedidodet.iditem, alm_inventario.descripcion, ped_pedidodet.canpro; " _
                + vbCr + "Union " _
                + vbCr + "SELECT ped_pedido.oc, ped_pedido.fchemi, ped_pedidodetent.fchent, ped_pedido.idcli, mae_cliente.nombre, ped_pedidodetent.iditem, alm_inventario.descripcion, ped_pedidodetent.canpro " _
                + vbCr + "FROM ((ped_pedido LEFT JOIN mae_cliente ON ped_pedido.idcli = mae_cliente.id) LEFT JOIN ped_pedidodetent ON ped_pedido.id = ped_pedidodetent.idped) LEFT JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id " _
                + vbCr + "Where (((ped_pedidodetent.idItem) = " & IDPRO_ & ") And ((ped_pedido.oc) = '" & NUMPED_ & "') And ((ped_pedidodetent.fchent) Is Not Null And (ped_pedidodetent.fchent) > #1/1/2011#)) " _
                + vbCr + "GROUP BY ped_pedido.oc, ped_pedido.fchemi, ped_pedidodetent.fchent, ped_pedido.idcli, mae_cliente.nombre, ped_pedidodetent.iditem, alm_inventario.descripcion, ped_pedidodetent.canpro;"
        
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        ' Se llena el primer detalle de un pedido
        With FgGrid
            ' Si no se estan mostrando historicos
            If Not HISTORIAL Then .Rows = 2
            xRs.MoveFirst
            FILAPED_ = 0
            For A = 1 To xRs.RecordCount
                .Rows = .Rows + 1
                If MOSTRARCLIENTE_ Then
                    .TextMatrix(.Rows - 1, COLUMNAINICIO_ - 3) = NulosC(xRs("nombre"))
                    .Select .Rows - 1, COLUMNAINICIO_ - 3
                    .CellAlignment = flexAlignRightCenter
                End If
                If MOSTRARORDEN_ Then .TextMatrix(.Rows - 1, COLUMNAINICIO_ - 2) = Format(NulosC(xRs("oc")), "0000000000")
                If MOSTRARFECHAEMISION_ Then .TextMatrix(.Rows - 1, COLUMNAINICIO_ - 1) = NulosC(xRs("fchemi"))
                .TextMatrix(.Rows - 1, COLUMNAINICIO_) = NulosC(xRs("fchent"))
                .TextMatrix(.Rows - 1, COLUMNAINICIO_ + 1) = Format(NulosN(xRs("canpro")), "0.00")
                SUMAPED_ = SUMAPED_ + NulosN(xRs("canpro"))
                FILAPED_ = FILAPED_ + 1 ' Numero de filas que tiene el pedido
                xRs.MoveNext
            Next A
        End With
        
        SUMAGUIA_ = 0
        ' Se llenan Entregas
        llenarVentana IDPRO_, NUMPED_, FgGrid, False, True, False, FILAPED_, SUMAPED_, SUMAGUIA_, HISTORIAL, COLUMNAINICIO_, MOSTRARORDEN_
        ' Se llenan los totales
        llenarVentana IDPRO_, NUMPED_, FgGrid, False, False, True, FILAPED_, SUMAPED_, SUMAGUIA_, HISTORIAL, COLUMNAINICIO_, MOSTRARORDEN_
    End If
    
    If ENTREGA_ Then
        Dim FILAENTREGA_ As Double
        Set xRs = Nothing
        
        cSQL = "SELECT vta_guia.numordcom, vta_guiadet.iditem, vta_guia.idcli, vta_guia.fchentord, vta_guiadet.canpro " _
                + vbCr + "FROM vta_guia INNER JOIN vta_guiadet ON vta_guia.id = vta_guiadet.idgui " _
                + vbCr + "Where (((vta_guia.numordcom) = '" & NUMPED_ & "') And ((vta_guiadet.idItem) = " & IDPRO_ & ")) " _
                + vbCr + "GROUP BY vta_guia.numordcom, vta_guiadet.iditem, vta_guia.idcli, vta_guia.fchentord, vta_guiadet.canpro;"
        
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        'Se llena el primer detalle de un pedido
        With FgGrid
            xRs.MoveFirst
            FILAENTREGA_ = 0
            For A = 1 To xRs.RecordCount
                If FILAPED_ >= A Then
                    FILAENTREGA_ = (.Rows - 1) - (FILAPED_ - A)
                Else
                    FILAENTREGA_ = .Rows - 1
                End If
                
                .TextMatrix(FILAENTREGA_, COLUMNAINICIO_ + 2) = NulosC(xRs("fchentord"))
                .TextMatrix(FILAENTREGA_, COLUMNAINICIO_ + 3) = Format(NulosN(xRs("canpro")), "0.00")
                SUMAGUIA_ = SUMAGUIA_ + NulosN(xRs("canpro"))
                ' si es el ultimo bucle ya no se aumenta una fila
                If A = xRs.RecordCount Then GoTo SIGUIENTE
                If FILAENTREGA_ >= .Rows - 1 Then .Rows = .Rows + 1
SIGUIENTE:
                xRs.MoveNext
            Next A
        End With
    End If
    
    If TOTAL_ Then
        With FgGrid
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, COLUMNAINICIO_) = "TOTAL"
            .TextMatrix(.Rows - 1, COLUMNAINICIO_ + 1) = Format(NulosN(SUMAPED_), "0.00")
            .TextMatrix(.Rows - 1, COLUMNAINICIO_ + 3) = Format(NulosN(SUMAGUIA_), "0.00")
            .TextMatrix(.Rows - 1, COLUMNAINICIO_ + 4) = Format(NulosN(SUMAPED_ - SUMAGUIA_), "0.00")
            
            .Select .Rows - 1, COLUMNAINICIO_, .Rows - 1, COLUMNAINICIO_ + 4
            .FillStyle = flexFillRepeat
            .CellBackColor = &HC0FFFF
            .CellForeColor = &H8000000D
            
            .Select .Rows - 1, COLUMNAINICIO_ + 4
            If .TextMatrix(.Rows - 1, COLUMNAINICIO_ + 4) <= 0 Then
                ' Si hay stock para entrega
                .CellForeColor = &H8000000D
            Else
                ' Si no hay stock para entrega
                .CellForeColor = &HFF&
            End If
            
            SUMAPED_ = 0
            SUMAGUIA_ = 0
            .Select 2, 1, 2, 5
        End With
    End If
End Sub

Private Sub LlenarDatos(RESUMIDO_ As Boolean, DETALLADO_ As Boolean)
    Dim A As Integer
    If RESUMIDO_ Then
        fg(0).Rows = 1
        If RstResumido.State = 0 Then Exit Sub
        If RstResumido.RecordCount = 0 Then Exit Sub
        RstResumido.Sort = "fchent Desc"
        
        RstResumido.MoveFirst
        XcentrarFrm FraProgreso
        FraProgreso.Visible = True
        PgBar.Min = 0
        PgBar.Max = RstResumido.RecordCount
        For A = 1 To RstResumido.RecordCount
            FraProgreso.Refresh
            PgBar.Value = A
            fg(0).Rows = fg(0).Rows + 1
            fg(0).TextMatrix(A, 1) = Format(NulosN(RstResumido("numped")), "0000000000")
            fg(0).TextMatrix(A, 2) = NulosC(RstResumido("nomcli"))
            fg(0).TextMatrix(A, 3) = NulosC(RstResumido("despro"))
            fg(0).TextMatrix(A, 4) = NulosC(RstResumido("fchemi"))
            fg(0).TextMatrix(A, 5) = NulosC(RstResumido("fchent"))
            fg(0).TextMatrix(A, 6) = Format(NulosN(RstResumido("canpro")), "0.00")
            fg(0).TextMatrix(A, 7) = NulosC(RstResumido("unimed"))
            fg(0).TextMatrix(A, 8) = "" 'NulosC(RstResumido("anulado"))
            fg(0).TextMatrix(A, 9) = NulosC(RstResumido("numdocped"))
            fg(0).TextMatrix(A, 10) = NulosN(RstResumido("idped"))
            fg(0).TextMatrix(A, 11) = NulosN(RstResumido("idcli"))
            fg(0).TextMatrix(A, 12) = NulosN(RstResumido("idpro"))
            RstResumido.MoveNext
        Next A
        FraProgreso.Visible = False
    End If
    
    If DETALLADO_ Then
        fg(1).Rows = 2
        If RstDetallado.State = 0 Then Exit Sub
        If RstDetallado.RecordCount = 0 Then Exit Sub
        RstDetallado.MoveFirst
        
        For A = 2 To RstDetallado.RecordCount + 1
            fg(1).Rows = fg(1).Rows + 1
            fg(1).TextMatrix(A, 1) = Format(NulosN(RstDetallado("numped")), "0000000000")
            fg(1).TextMatrix(A, 2) = NulosC(RstDetallado("nomcli"))
            fg(1).TextMatrix(A, 3) = NulosC(RstDetallado("nompro"))
            fg(1).TextMatrix(A, 4) = NulosC(RstDetallado("fchemi"))
            fg(1).TextMatrix(A, 5) = Format(NulosN(RstDetallado("cantotped")), "0.00")
            fg(1).TextMatrix(A, 6) = Format(NulosN(RstDetallado("cantotent")), "0.00")
            fg(1).TextMatrix(A, 7) = NulosC(RstDetallado("fchultent"))
            fg(1).TextMatrix(A, 8) = Format(NulosN(RstDetallado("canultent")), "0.00")
            fg(1).TextMatrix(A, 9) = NulosC(RstDetallado("fchparcped"))
            fg(1).TextMatrix(A, 10) = Format(NulosN(RstDetallado("canparcped")), "0.00")
            fg(1).TextMatrix(A, 11) = Format(NulosN(RstDetallado("cantotaent")), "0.00")
            fg(1).TextMatrix(A, 12) = NulosN(RstDetallado("idpro"))
            fg(1).TextMatrix(A, 13) = NulosN(RstDetallado("idcli"))
            
            ' Se da color a la entrega
            fg(1).Select fg(1).Rows - 1, 11, fg(1).Rows - 1, 11
            If NulosN(fg(1).TextMatrix(A, 11)) <= 0 Then
                ' Si hay stock para entrega
                fg(1).CellForeColor = &H8000000D
            Else
                ' Si no hay stock para entrega
                fg(1).CellForeColor = &HFF&
            End If
                        
            RstDetallado.MoveNext
        Next A
    End If
End Sub

Private Sub XcentrarFrm(ByRef frm As Frame)
    With frm
        .Left = ((Me.Width - .Width) / 2)
        .Top = ((Me.Height - .Height) / 2)
    End With
End Sub

Private Sub iniciarCampos()
    Dim MES_ As Integer
    Dim ANIO_ As Integer
    
    cargo = False
    
    Set fg(3).DataSource = Nothing
    Set fg(4).DataSource = Nothing
    Set fg(5).DataSource = Nothing
    'Se inicializa:
    'datos para clientes
    fg(3).Rows = 1
    GRID_COMBOLIST fg(3), 1
    fg(3).Editable = flexEDKbdMouse
    'datos para productos
    fg(4).Rows = 1
    GRID_COMBOLIST fg(4), 1
    fg(4).Editable = flexEDKbdMouse
    'datos para Ordenes de Compra
    fg(5).Rows = 1
    GRID_COMBOLIST fg(5), 1
    fg(5).Editable = flexEDKbdMouse
    
    'datos para fechas
    TxtFchEmiDesde.Valor = CDate("01/01/" & AnoTra & "")
    TxtFchEmiHasta.Valor = Date
    TxtFchEntDesde.Valor = CDate("01/" + CStr(Month(Date)) + "/" + CStr(Year(Date)))
    
    MES_ = Month(Date) + 1
    ANIO_ = Year(Date)
    If MES_ > 12 Then MES_ = 1: ANIO_ = ANIO_ + 1
    TxtFchEntHasta.Valor = CDate("01/" + CStr(MES_) + "/" + CStr(ANIO_)) - 1
    ' datos para el check Entregas
    OptTipo(1).Value = True
    
    ' datos para el reporte Simple
    fg(0).AllowUserResizing = flexResizeColumns
    fg(0).AutoSearch = flexSearchFromTop
    fg(0).ExplorerBar = flexExSortShowAndMove
    fg(0).SelectionMode = flexSelectionByRow
    fg(0).ForeColorSel = &H80000005
    fg(0).BackColorSel = &H80&
    
    fg(0).ColWidth(0) = 0
    fg(0).ColWidth(9) = 0
    fg(0).ColWidth(10) = 0
    fg(0).ColWidth(11) = 0
    fg(0).ColWidth(12) = 0
    
    fg(0).Top = 45
    fg(0).Left = 0
    fg(0).Width = Frame6.Width - 150
    fg(0).Height = Frame6.Height - 100
    
    ' datos para el reporte Compuesto
    fg(1).AllowUserResizing = flexResizeColumns
    fg(1).AutoSearch = flexSearchFromTop
    fg(1).ExplorerBar = flexExSortShowAndMove
    fg(1).SelectionMode = flexSelectionByRow
    fg(1).ForeColorSel = &H80000005
    fg(1).BackColorSel = &H80&
    
    fg(1).Rows = 2
    fg(1).FixedRows = 2
    fg(1).FrozenCols = 3
    
    fg(2).Rows = 2
    fg(2).FixedRows = 2
    
    configurarGrid True, False
End Sub

Private Sub configurarGrid(RESUMIDO_ As Boolean, DETALLADO_ As Boolean, _
                                        Optional VENTANA_ As Boolean = False, _
                                        Optional HISTORICO_ As Boolean = False)
    If RESUMIDO_ Then
        fg(0).ColWidth(8) = 0
        fg(0).ColWidth(9) = 0
        fg(0).ColWidth(10) = 0
        fg(0).ColWidth(11) = 0
        fg(0).ColWidth(12) = 0
        
        fg(0).RowHeight(0) = 480
    End If
    
    If DETALLADO_ Then
        fg(1).ColWidth(0) = 0
        fg(1).ColWidth(12) = 0
        fg(1).ColWidth(13) = 0
        GRID_COMBINAR fg(1), 0, 1, 1, 1, "Ord. Pedido", flexAlignCenterCenter, False, flexMergeFixedOnly, , &H8000000F, False
        GRID_COMBINAR fg(1), 0, 2, 1, 2, "Cliente", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR fg(1), 0, 3, 1, 3, "Producto", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR fg(1), 0, 4, 1, 4, "Fch. Emision", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR fg(1), 0, 5, 0, 6, "Totales Cantidad", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(1), 1, 5, 1, 5, "Pedido", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(1), 1, 6, 1, 6, "Entregado", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(1), 0, 7, 0, 8, "Ultima Entrega", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(1), 1, 7, 1, 7, "Fecha", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(1), 1, 8, 1, 8, "Cantidad", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(1), 0, 9, 0, 10, "Hasta el " & NulosC(TxtFchEntHasta.Valor) & "", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(1), 1, 9, 1, 9, "Fch. de Entrega", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(1), 1, 10, 1, 10, "Total Pedido", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(1), 0, 11, 1, 11, "Total a Entregar", flexAlignCenterCenter, False, , , &H8000000F, True
        
        fg(1).MergeCells = flexMergeFixedOnly
        
        fg(1).ColWidth(0) = 0
        fg(1).ColWidth(1) = 1020
        fg(1).ColWidth(2) = 2430
        fg(1).ColWidth(3) = 4500
        fg(1).ColWidth(4) = 1005
        fg(1).ColWidth(5) = 915
        fg(1).ColWidth(6) = 915
        fg(1).ColWidth(7) = 1140
        fg(1).ColWidth(8) = 915
        fg(1).ColWidth(9) = 1290
        fg(1).ColWidth(10) = 1035
        fg(1).ColWidth(11) = 1515
        fg(1).ColWidth(12) = 0
        fg(1).ColWidth(13) = 0
        
        If fg(1).Rows > 2 Then
            fg(1).Select 2, fg(1).Cols - 5, fg(1).Rows - 1, fg(1).Cols - 5
            fg(1).FillStyle = flexFillRepeat
            fg(1).CellBackColor = &HDDFFFF        '&H8000000F&
            
            fg(1).Select 2, fg(1).Cols - 3, fg(1).Rows - 1, fg(1).Cols - 3
            fg(1).FillStyle = flexFillRepeat
            fg(1).CellBackColor = &HDDFFFF
            
            fg(1).Select 2, 1, 2, 1
        End If
    End If
    
    If VENTANA_ Then
        fg(2).ColWidth(0) = 0
        GRID_COMBINAR fg(2), 0, 1, 0, 2, "Pedido", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(2), 1, 1, 1, 1, "Fecha", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(2), 1, 2, 1, 2, "Cantidad", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(2), 0, 3, 0, 4, "Entrega", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(2), 1, 3, 1, 3, "Fecha", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(2), 1, 4, 1, 4, "Cantidad", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(2), 0, 5, 1, 5, "Total Restante", flexAlignCenterCenter, False, , , &H8000000F, False
        fg(2).MergeCells = flexMergeFixedOnly
    End If
    
    If HISTORICO_ Then
        GRID_COMBINAR fg(1), 0, 1, 1, 1, "Producto / Cliente", flexAlignCenterCenter, False, flexMergeFixedOnly, , &H8000000F, False
        GRID_COMBINAR fg(1), 0, 2, 0, 5, "Pedido", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(1), 1, 2, 1, 2, "Ord. Pedido", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR fg(1), 1, 3, 1, 3, "Fch. Emision", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR fg(1), 1, 4, 1, 4, "Fch. Entrega", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR fg(1), 1, 5, 1, 5, "Cantidad", flexAlignCenterCenter, False, , , &H8000000F, False
        GRID_COMBINAR fg(1), 0, 6, 0, 7, "Entrega", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(1), 1, 6, 1, 6, "Fch. Entregada", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(1), 1, 7, 1, 7, "Cantidad", flexAlignCenterCenter, True, , , &H8000000F, False
        GRID_COMBINAR fg(1), 0, 8, 1, 8, "Total Restante", flexAlignCenterCenter, False, , , &H8000000F, False
        
        fg(1).MergeCells = flexMergeFixedOnly
        fg(1).FrozenCols = 0
        
        fg(1).ColWidth(0) = 0
        fg(1).ColWidth(1) = 4500
        fg(1).ColWidth(2) = 1200
        fg(1).ColWidth(3) = 1200
        fg(1).ColWidth(4) = 1200
        fg(1).ColWidth(5) = 1200
        fg(1).ColWidth(6) = 1200
        fg(1).ColWidth(7) = 1200
        fg(1).ColWidth(8) = 1200
        fg(1).ColWidth(9) = 0
        fg(1).ColWidth(10) = 0
        fg(1).ColWidth(11) = 0
        fg(1).ColWidth(12) = 0
        fg(1).ColWidth(13) = 0
        
        If fg(1).Rows > 2 Then
            fg(1).Select 2, fg(1).Cols - 5, fg(1).Rows - 1, fg(1).Cols - 5
            fg(1).FillStyle = flexFillRepeat
            fg(1).CellBackColor = &HDDFFFF        '&H8000000F&
            
            fg(1).Select 2, fg(1).Cols - 3, fg(1).Rows - 1, fg(1).Cols - 3
            fg(1).FillStyle = flexFillRepeat
            fg(1).CellBackColor = &HDDFFFF
            
            fg(1).Select 2, 1, 2, 1
        End If
    End If
End Sub

Private Sub Cmd_Click(Index As Integer)
    If Index = 0 Then ' Boton Consultar
        Buscar
    End If
End Sub

Private Sub fg_CellButtonClick(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim xCampos() As String
    Dim xRs As New ADODB.Recordset
    
    If Index = 3 Then ' Clientes
        ReDim xCampos(2, 3) As String
        Set xRs = Nothing
        
        xCampos(0, 0) = "Nombre":          xCampos(0, 1) = "nombre":     xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
        xCampos(1, 0) = "Nº R.UC.":        xCampos(1, 1) = "numruc":     xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
        
        cSQL = "SELECT mae_cliente.nombre, mae_cliente.numruc " _
               + vbCr + "From mae_cliente " _
               + vbCr + "WHERE mae_cliente.nombre <> ''" _
               + vbCr + "GROUP BY mae_cliente.id, mae_cliente.nombre, mae_cliente.numruc " _
               + vbCr + "ORDER BY mae_cliente.nombre;"
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando Clientes", "nombre", "nombre", Principio
        
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        fg(Index).TextMatrix(fg(Index).Row, 1) = xRs("nombre")
    End If
    
    If Index = 4 Then ' Productos
        ReDim xCampos(2, 3) As String
        Set xRs = Nothing
        
        xCampos(0, 0) = "Nombre":    xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
        xCampos(1, 0) = "Codigo":    xCampos(1, 1) = "codpro":        xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
    
        cSQL = "SELECT alm_inventario.descripcion, alm_inventario.codpro " _
            + vbCr + "From alm_inventario " _
            + vbCr + "WHERE (((alm_inventario.activo)=-1) AND ((alm_inventario.tippro) In (1,3)) AND ((alm_inventario.idcuentaven)<>0)) " _
            + vbCr + "GROUP BY alm_inventario.descripcion, alm_inventario.codpro " _
            + vbCr + "ORDER BY alm_inventario.descripcion;"
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando Productos", "descripcion", "descripcion", Principio
         
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        fg(Index).TextMatrix(fg(Index).Row, 1) = xRs("descripcion")
    End If
    
    If Index = 5 Then ' Ordenes de Pedido
        ReDim xCampos(2, 3) As String
        Set xRs = Nothing
        
        xCampos(0, 0) = "Orden de Compra":    xCampos(0, 1) = "oc":       xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
        xCampos(1, 0) = "Id Pedido":          xCampos(1, 1) = "idped":    xCampos(1, 2) = "2500":   xCampos(1, 3) = "C"
    
        cSQL = "SELECT DISTINCT ped_pedido.oc, ped_pedido.id AS idped " _
             + vbCr + "From ped_pedido " _
             + vbCr + "GROUP BY ped_pedido.oc, ped_pedido.id " _
             + vbCr + "HAVING (((ped_pedido.oc) Is Not Null And (ped_pedido.oc)<>'S/N' And (ped_pedido.oc)<>'') AND ((ped_pedido.id) Is Not Null)) " _
             + vbCr + "ORDER BY ped_pedido.oc;"
        
        CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando Ordenes de Compra", "oc", "oc", CualquierParte
                
        If xRs.State = 0 Then Exit Sub
        If xRs.RecordCount = 0 Then Exit Sub
        
        fg(Index).TextMatrix(fg(Index).Row, 1) = xRs("oc")
    End If
    
    Set xRs = Nothing
End Sub

Private Sub fg_DblClick(Index As Integer)
    Dim FILA_ As Double
    Dim SUMPEDIDO_ As Double
    Dim SUMENTREGA_ As Double
    Dim NUMORDEN_ As Double
    Dim IDPRO_ As Double
    
    FILA_ = 0
    SUMPEDIDO_ = 0
    SUMENTREGA_ = 0
    NUMORDEN_ = NulosN(fg(1).TextMatrix(fg(1).Row, 1))
    IDPRO_ = NulosN(fg(1).TextMatrix(fg(1).Row, 12))
    
    If Index = 1 Then ' Reporte
        ' Si no es consulta detallada se sale
        If OptTipo(1).Value = False Then Exit Sub
        
        LblProd.Caption = NulosC(fg(1).TextMatrix(fg(1).Row, 3))
        LblCliente.Caption = NulosC(fg(1).TextMatrix(fg(1).Row, 2))
        LblOrden.Caption = NulosC(NUMORDEN_)
        TxtFchEmi.Valor = NulosC(fg(1).TextMatrix(fg(1).Row, 4))
        
        XcentrarFrm Frame10
        Frame10.Visible = True
        
        configurarGrid False, False, True
        llenarVentana IDPRO_, NUMORDEN_, fg(2)
    End If
End Sub

Private Sub fg_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Select Case Index
        Case 3, 4, 5
            If KeyCode = vbKeyInsert Then ' Agregar
                menu00_Click
            End If
            If KeyCode = vbKeyDelete Then ' Eliminar
                menu01_Click
            End If
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub fg_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    Select Case Index
        Case 3, 4, 5
            INDICE_ = Index
            PopupMenu menu
        Case Else
            Exit Sub
    End Select
End Sub

Private Sub fg_RowColChange(Index As Integer)
    If Index <> 1 Then Exit Sub
    If Frame10.Visible = False Then Exit Sub
    
    Dim FILA_ As Double
    Dim SUMPEDIDO_ As Double
    Dim SUMENTREGA_ As Double
    Dim NUMORDEN_ As Double
    Dim IDPRO_ As Double
    
    FILA_ = 0
    SUMPEDIDO_ = 0
    SUMENTREGA_ = 0
    NUMORDEN_ = NulosN(fg(1).TextMatrix(fg(1).Row, 1))
    IDPRO_ = NulosN(fg(1).TextMatrix(fg(1).Row, 12))
    
    If Index = 1 Then ' Reporte
        LblProd.Caption = NulosC(fg(1).TextMatrix(fg(1).Row, 3))
        LblCliente.Caption = NulosC(fg(1).TextMatrix(fg(1).Row, 2))
        LblOrden.Caption = NulosC(NUMORDEN_)
        TxtFchEmi.Valor = NulosC(fg(1).TextMatrix(fg(1).Row, 4))
                
        configurarGrid False, False, True
        llenarVentana IDPRO_, NUMORDEN_, fg(2)
    End If
End Sub

Private Sub Form_Load()
    iniciarCampos
End Sub

Private Sub Form_Resize()
    ' Si esta minimizado no se hace nada
    If Me.WindowState = 1 Then Exit Sub
    
    If Me.Width <= 13200 Then Me.Width = 13200
    If Me.Height <= 2850 Then Me.Height = 2850
        
    ' Se dimensiona el contenido
    Frame6.Width = Me.Width - 60
    Frame6.Height = Me.Height - 1700
    
    fg(0).Width = Frame6.Width - 100
    fg(0).Height = Frame6.Height - 100
    
    fg(1).Width = Frame6.Width - 100
    fg(1).Height = Frame6.Height - 100
    
    If Frame10.Visible = True Then XcentrarFrm Frame10
End Sub

Private Sub menu00_Click()
    fg(INDICE_).Rows = fg(INDICE_).Rows + 1
    fg(INDICE_).Select fg(INDICE_).Rows - 1, 1
    If fg(INDICE_).Rows > 2 Then fg(INDICE_).TopRow = fg(INDICE_).Rows - 2
    fg_CellButtonClick INDICE_, fg(INDICE_).Rows - 1, 1
End Sub

Private Sub menu01_Click()
    If fg(INDICE_).Row < fg(INDICE_).FixedRows Then Exit Sub
    fg(INDICE_).RemoveItem fg(INDICE_).Row
End Sub

Private Sub PbCerrar_Click(Index As Integer)
    Frame10.Visible = False
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then
        If verificarDatos Then
            Buscar
        End If
    End If
    
    If Button.Index = 5 Then
        If Not cargo Then MsgBox "No se ha procesado ninguna Consulta, procesela antes de Exportar", vbCritical + vbOKOnly, "Reporte de Pedidos": Exit Sub
        If fg(1).Visible = True Then ExportarExcel fg(1)
        If fg(0).Visible = True Then ExportarExcel fg(0)
    End If
    
    If Button.Index = 8 Then
        Unload Me
    End If
End Sub

Private Sub crearCabeceraExcel(ByRef OBJETOEXCEL_ As Object, ByRef FILA_ As Integer, _
                                                ByRef COLUMNA_ As Integer, ByRef GRID_ As VSFlexGrid, _
                                                Optional SIMPLE_ As Boolean = False, _
                                                Optional DETALLADO_ As Boolean = True, _
                                                Optional HISTORICO_ As Boolean = False)
    Dim TITULO_ As String
    
    With OBJETOEXCEL_.ActiveSheet
        If SIMPLE_ Then
            FILA_ = FILA_ + 2
            ' Se llena la cabecera
            ' Ord. Pedido
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
            ' Cliente
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 35
            ' Producto
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 60
            ' Fecha de Emision
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
            ' Fecha de Entrega
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
            ' Cantidad
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
            ' Unidad de Medida
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 5
            ' Condicion
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
            
            TITULO_ = "REPORTE SIMPLE DE PEDIDOS   " & TxtFchEntDesde.Valor & " - " & TxtFchEntHasta.Valor & ""
        End If
        
        If DETALLADO_ Then
            FILA_ = FILA_ + 2
            ' Se llena la cabecera
            ' Ord. Pedido
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
            ' Cliente
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 35
            ' Producto
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 60
            ' Fecha de Emision
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
            ' Totales Cantidad
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_ + 1) + CStr(FILA_)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_ + 1) + CStr(FILA_)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
                ' Pedido
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(1, COLUMNA_ - 65)
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
                ' Entregado
                COLUMNA_ = COLUMNA_ + 1
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(1, COLUMNA_ - 65)
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
            ' Ultima Entrega
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_ + 1) + CStr(FILA_)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_ + 1) + CStr(FILA_)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
                ' Pedido
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(1, COLUMNA_ - 65)
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
                ' Entregado
                COLUMNA_ = COLUMNA_ + 1
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(1, COLUMNA_ - 65)
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
                
            ' Hasta la Fecha
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_ + 1) + CStr(FILA_)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_ + 1) + CStr(FILA_)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
                ' Fecha de Entrega
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(1, COLUMNA_ - 65)
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 16
                ' Total Pedido
                COLUMNA_ = COLUMNA_ + 1
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(1, COLUMNA_ - 65)
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 16
            ' Total a Entregar
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 16
            
            TITULO_ = "REPORTE DETALLADO DE PEDIDOS   " & TxtFchEntDesde.Valor & " - " & TxtFchEntHasta.Valor & ""
        End If
    
        If HISTORICO_ Then
            FILA_ = FILA_ + 2
            ' Se llena la cabecera
            ' Pedido / Cliente
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 60
            ' Pedido
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_ + 3) + CStr(FILA_)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_ + 3) + CStr(FILA_)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
                ' Ord. Pedido
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(1, COLUMNA_ - 65)
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
                ' Fecha de Emision
                COLUMNA_ = COLUMNA_ + 1
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(1, COLUMNA_ - 65)
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
                ' Fecha de Entrega
                COLUMNA_ = COLUMNA_ + 1
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(1, COLUMNA_ - 65)
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
                ' Cantidad
                COLUMNA_ = COLUMNA_ + 1
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(1, COLUMNA_ - 65)
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
            ' Entrega
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_ + 1) + CStr(FILA_)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_ + 1) + CStr(FILA_)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
                ' Fecha Entregada
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(1, COLUMNA_ - 65)
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 14
                ' Cantidad
                COLUMNA_ = COLUMNA_ + 1
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(1, COLUMNA_ - 65)
                .Range(Chr(COLUMNA_) + CStr(FILA_ + 1) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 12
            ' Total Restante
            COLUMNA_ = COLUMNA_ + 1
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = "'" + GRID_.TextMatrix(0, COLUMNA_ - 65)
            .Range(Chr(COLUMNA_) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).ColumnWidth = 16
            
            TITULO_ = "REPORTE HISTORICO DE PEDIDOS   " & TxtFchEntDesde.Valor & " - " & TxtFchEntHasta.Valor & ""
        End If
        
        ' Se da Formato a las celdas que conforman la cabecera
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).HorizontalAlignment = -4108  'xlCenter ' Alineacion Horizontal
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).VerticalAlignment = -4108  'xlCenter ' Alineacion Vertical
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Font.Bold = True ' Letra en Negrita
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Font.Size = 12 ' Tamaño de letra
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Interior.Color = RGB(234, 234, 234) ' Color de Fondo
        ' Se dan formato a los Bordes de las celdas
            ' xlContinuous se puede cambiar por xlDouble, xlThick
        
''        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Borders(xlTop).LineStyle =  xlContinuous
''        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Borders(xlBottom).LineStyle =  xlContinuous
''        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Borders(xlRight).LineStyle =  xlContinuous
''        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Borders(xlLeft).LineStyle =  xlContinuous
        
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Borders(-4160).LineStyle = 1  'xlContinuous
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Borders(-4107).LineStyle = 1 'xlContinuous
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Borders(-4152).LineStyle = 1 'xlContinuous
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Borders(-4131).LineStyle = 1 'xlContinuous
        FILA_ = FILA_ - 2
        ' Se llena el Titulo
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Merge
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Value = TITULO_
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).HorizontalAlignment = -4108  'xlCenter  ' Alineacion Horizontal
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).VerticalAlignment = -4108  'xlCenter  ' Alineacion Vertical
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Font.Bold = True ' Letra en Negrita
        .Range(Chr(66) + CStr(FILA_) + ":" + Chr(COLUMNA_) + CStr(FILA_ + 1)).Font.Size = 15 ' Tamaño de letra
        FILA_ = FILA_ + 4
    End With
End Sub

Sub ExportarExcel(ByRef GRID_ As VSFlexGrid)
    Dim A As Integer
    Dim B As Integer
    Dim FILA_ As Integer
    Dim xCad As String
    Dim objExcel As Object
    Dim COLUMNA_ As Integer
    
    Set objExcel = CreateObject("Excel.Application")
    
    objExcel.Visible = True
    'determina el numero de hojas que se mostrara en el Excel
    objExcel.SheetsInNewWorkbook = 1
    
    objExcel.WindowState = 2
    objExcel.Workbooks.Add
    ' Se aplica zoom al 75%
    objExcel.ActiveWindow.Zoom = 75
   
    With objExcel.ActiveSheet
        FILA_ = 2
        COLUMNA_ = 66  ' codigo Ascii de la 'B'
             
        If OptTipo(0).Value = True Then
            crearCabeceraExcel objExcel, FILA_, COLUMNA_, GRID_, True, False, False ' Simple
        Else
            If OptTipo(1).Value = True Then
                crearCabeceraExcel objExcel, FILA_, COLUMNA_, GRID_ ' Detallado
            Else
                crearCabeceraExcel objExcel, FILA_, COLUMNA_, GRID_, False, False, True ' Historico
            End If
        End If
             
        For A = 2 To GRID_.Rows - 1
            For B = 1 To COLUMNA_ - 65
                If IsNumeric(GRID_.TextMatrix(A, B)) Then
                    If GRID_.TextMatrix(1, B) = "Ord. Pedido" Then GoTo PROCESARCOMOCADENA
                    .Cells(FILA_, B + 1) = GRID_.TextMatrix(A, B)
                    .Cells(FILA_, B + 1).NumberFormat = "#,##0.00"
                Else
PROCESARCOMOCADENA:
                    .Cells(FILA_, B + 1) = "'" + GRID_.TextMatrix(A, B)
                End If
            Next B
            FILA_ = FILA_ + 1
        Next A
    End With
    
    MsgBox "El proceso de exportacion termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, "Reporte de Pedidos"
    objExcel.WindowState = 1
    Set objExcel = Nothing
    Exit Sub
End Sub

'Metodos para arrastrar el Frame
''''''''''''''''''''''''''''''''
Private Sub Frame10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    Frame10.ZOrder 0
End Sub

Private Sub Frame10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With Frame10
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub


