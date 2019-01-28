VERSION 5.00
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FormReportePedidos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas  -  Reporte de Pedidos"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   11925
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   11925
   StartUpPosition =   3  'Windows Default
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
      Height          =   1875
      Left            =   6780
      TabIndex        =   8
      Top             =   780
      Width           =   5025
      Begin VB.TextBox TextOrdCompra 
         Height          =   285
         Left            =   1500
         TabIndex        =   32
         Top             =   1500
         Width           =   1215
      End
      Begin AspaTextBoxFecha.TextBoxFecha TextBoxFechaEmision1 
         Height          =   300
         Left            =   1500
         TabIndex        =   14
         Top             =   390
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
         Left            =   3510
         TabIndex        =   16
         Top             =   390
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
         Left            =   1500
         TabIndex        =   18
         Top             =   930
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
         Left            =   3510
         TabIndex        =   20
         Top             =   930
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
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Nº :"
         Height          =   195
         Left            =   1080
         TabIndex        =   21
         Top             =   1560
         Width           =   270
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   2910
         TabIndex        =   19
         Top             =   980
         Width           =   465
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   870
         TabIndex        =   17
         Top             =   980
         Width           =   510
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Hasta:"
         Height          =   195
         Left            =   2910
         TabIndex        =   15
         Top             =   430
         Width           =   465
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Desde:"
         Height          =   195
         Left            =   870
         TabIndex        =   13
         Top             =   430
         Width           =   510
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Orden de Compra"
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
         TabIndex        =   11
         Top             =   1305
         Width           =   1245
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
         Left            =   120
         TabIndex        =   10
         Top             =   750
         Width           =   1185
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
         TabIndex        =   9
         Top             =   200
         Width           =   990
      End
   End
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
      Height          =   4545
      Left            =   30
      TabIndex        =   7
      Top             =   2760
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
         Height          =   4455
         Left            =   2130
         TabIndex        =   12
         Top             =   150
         Visible         =   0   'False
         Width           =   11805
         Begin VB.Frame Frame5 
            Height          =   1245
            Left            =   180
            TabIndex        =   38
            Top             =   750
            Width           =   11445
            Begin VB.Line Line1 
               X1              =   210
               X2              =   3500
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
               Left            =   2325
               TabIndex        =   48
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
               TabIndex        =   47
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
               TabIndex        =   46
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
               Left            =   2325
               TabIndex        =   45
               Top             =   250
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
               Left            =   2325
               TabIndex        =   44
               Top             =   550
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
               TabIndex        =   43
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
               Left            =   3825
               TabIndex        =   42
               Top             =   250
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
               Left            =   5025
               TabIndex        =   41
               Top             =   250
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
               Left            =   5025
               TabIndex        =   40
               Top             =   550
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
               Left            =   3810
               TabIndex        =   39
               Top             =   550
               Width           =   1080
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid VSFlexGridDetalle2 
            Height          =   2300
            Left            =   150
            TabIndex        =   28
            Top             =   2050
            Width           =   6645
            _cx             =   11721
            _cy             =   4057
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
            Height          =   2300
            Left            =   6900
            TabIndex        =   27
            Top             =   2050
            Width           =   4725
            _cx             =   8334
            _cy             =   4057
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
         Begin VB.Label LabelDetalleID 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID_01"
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
            Left            =   2520
            TabIndex        =   37
            Top             =   560
            Width           =   525
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ID PRODUCTO:"
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
            Left            =   510
            TabIndex        =   36
            Top             =   560
            Width           =   1365
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Nº O.C. :"
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
            Left            =   3990
            TabIndex        =   35
            Top             =   560
            Width           =   780
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
            Left            =   5190
            TabIndex        =   34
            Top             =   560
            Width           =   585
         End
         Begin VB.Label LabelDetalleProducto 
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
            Left            =   2520
            TabIndex        =   30
            Top             =   250
            Width           =   1365
         End
         Begin VB.Label Label15 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "PRODUCTO:"
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
            Left            =   510
            TabIndex        =   29
            Top             =   250
            Width           =   1110
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
            TabIndex        =   25
            Top             =   30
            Width           =   195
         End
      End
      Begin VB.CommandButton CommandMostOcult 
         Caption         =   "Mostrar / Ocultar Ord. Comp"
         Height          =   465
         Left            =   60
         TabIndex        =   26
         Top             =   3990
         Width           =   2745
      End
      Begin VB.CommandButton CommandLimpiar 
         Caption         =   "Limpiar Todo"
         Height          =   465
         Left            =   2940
         TabIndex        =   23
         Top             =   3990
         Width           =   1725
      End
      Begin VSFlex7Ctl.VSFlexGrid VSFlexGridReporte 
         Height          =   3615
         Left            =   90
         TabIndex        =   24
         Top             =   270
         Width           =   11685
         _cx             =   20611
         _cy             =   6376
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
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   225
         Left            =   5460
         TabIndex        =   33
         Top             =   4230
         Width           =   6255
         _ExtentX        =   11033
         _ExtentY        =   397
         _Version        =   393216
         Appearance      =   1
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2895
      Left            =   30
      TabIndex        =   0
      Top             =   -150
      Width           =   11865
      Begin VB.Frame Frame3 
         Caption         =   "Productos"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1875
         Left            =   60
         TabIndex        =   4
         Top             =   930
         Width           =   3290
         Begin VSFlex7Ctl.VSFlexGrid VSFlexGridProductos 
            Height          =   1485
            Left            =   90
            TabIndex        =   6
            ToolTipText     =   "Buscar Producto"
            Top             =   330
            Width           =   3135
            _cx             =   5530
            _cy             =   2619
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
            FormatString    =   $"FormRepPedidos.frx":0000
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   -1  'True
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
            ExplorerBar     =   7
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
      Begin VB.Frame Frame2 
         Height          =   495
         Left            =   60
         TabIndex        =   1
         Top             =   450
         Width           =   11835
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Reporte de Pedidos"
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
            Left            =   4980
            TabIndex        =   2
            Top             =   150
            Width           =   2115
         End
      End
      Begin MSComctlLib.Toolbar Toolbar1 
         Height          =   345
         Left            =   60
         TabIndex        =   3
         Top             =   180
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
            NumButtons      =   8
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
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
                  Picture         =   "FormRepPedidos.frx":003D
                  Key             =   "IMG1"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRepPedidos.frx":0581
                  Key             =   ""
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRepPedidos.frx":0913
                  Key             =   "IMG2"
               EndProperty
               BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRepPedidos.frx":0A97
                  Key             =   "IMG3"
               EndProperty
               BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRepPedidos.frx":0EEB
                  Key             =   "IMG4"
               EndProperty
               BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRepPedidos.frx":1003
                  Key             =   "IMG5"
               EndProperty
               BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRepPedidos.frx":1547
                  Key             =   "IMG6"
               EndProperty
               BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRepPedidos.frx":1A8B
                  Key             =   "IMG7"
               EndProperty
               BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRepPedidos.frx":1B9F
                  Key             =   "IMG8"
               EndProperty
               BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRepPedidos.frx":1CB3
                  Key             =   "IMG9"
               EndProperty
               BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRepPedidos.frx":2107
                  Key             =   "IMG10"
               EndProperty
               BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRepPedidos.frx":2273
                  Key             =   "IMG11"
               EndProperty
               BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRepPedidos.frx":27BB
                  Key             =   ""
               EndProperty
               BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "FormRepPedidos.frx":2AD5
                  Key             =   ""
               EndProperty
            EndProperty
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
         Height          =   1875
         Left            =   3390
         TabIndex        =   5
         Top             =   930
         Width           =   3290
         Begin VSFlex7Ctl.VSFlexGrid VSFlexGridClientes 
            Height          =   1485
            Left            =   60
            TabIndex        =   22
            ToolTipText     =   "Buscar Producto"
            Top             =   330
            Width           =   3135
            _cx             =   5530
            _cy             =   2619
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
            FormatString    =   $"FormRepPedidos.frx":2E67
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
      TabIndex        =   31
      Top             =   0
      Width           =   1365
   End
End
Attribute VB_Name = "FormReportePedidos"
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


Private Sub generarConsulta()
    Dim A As Integer
    Dim fila As Integer
    
    Dim c_PRODUCTOS As String
    Dim c_CLIENTES As String
    Dim c_FECHA_EMI As String
    Dim c_FECHA_PLAZ As String
    
    Dim c_FECHA_EMI_2 As String
    Dim c_FECHA_PLAZ_2 As String
    
    Dim c_FECHA_ENT As String
    
    Dim c_ORD_COMP_1 As String
    Dim c_ORD_COMP_2 As String
    
    Dim c_SQL As String
    
    limpiarVSFlexGridReporte
    
    VSFlexGridProductos.Row = 0
    VSFlexGridProductos.Col = 1
    c_PRODUCTOS = "((alm_inventario.descripcion)='" + VSFlexGridProductos.Text + "'"
    If (VSFlexGridProductos.TextMatrix(0, 1) = "Todos") Then
        c_PRODUCTOS = ""
    Else
        For A = 0 To VSFlexGridProductos.Rows - 1
            VSFlexGridProductos.Row = A
            VSFlexGridProductos.Col = 1
            c_PRODUCTOS = c_PRODUCTOS + " OR " + "(alm_inventario.descripcion)='" + VSFlexGridProductos.Text + "'"
        Next A
        c_PRODUCTOS = c_PRODUCTOS + ") AND "
    End If
    
    VSFlexGridClientes.Row = 0
    VSFlexGridClientes.Col = 1
    c_CLIENTES = "((mae_cliente.nombre)= '" + VSFlexGridClientes.Text + "'"
    If (VSFlexGridClientes.TextMatrix(0, 1) = "Todos") Then
        c_CLIENTES = ""
    Else
        For A = 0 To VSFlexGridClientes.Rows - 1
            VSFlexGridClientes.Row = A
            VSFlexGridClientes.Col = 1
            c_CLIENTES = c_CLIENTES + " OR " + "(mae_cliente.nombre)= '" + VSFlexGridClientes.Text + "'"
        Next A
        c_CLIENTES = c_CLIENTES + ") AND "
    End If
    c_FECHA_EMI = "((ped_pedido.fchemi)>=CDate('" & TextBoxFechaEmision1.Valor & "')"
    c_FECHA_EMI = c_FECHA_EMI & " AND (ped_pedido.fchemi)<=CDate('" & TextBoxFechaEmision2.Valor & "'))"
    
    c_FECHA_PLAZ = "((ped_pedidodetent.fchent)>=CDate('" & TextBoxFechaPlazo1.Valor & "')"
    c_FECHA_PLAZ = c_FECHA_PLAZ & " AND (ped_pedidodetent.fchent)<=CDate('" & TextBoxFechaPlazo2.Valor & "'))"
    
    c_FECHA_EMI_2 = "((vta_pedido.fchemi)>=CDate('" & TextBoxFechaEmision1.Valor & "')"
    c_FECHA_EMI_2 = c_FECHA_EMI_2 & " AND (vta_pedido.fchemi)<=CDate('" & TextBoxFechaEmision2.Valor & "'))"
    
    c_FECHA_PLAZ_2 = "((vta_pedido.fchent)>=CDate('" & TextBoxFechaPlazo1.Valor & "')"
    c_FECHA_PLAZ_2 = c_FECHA_PLAZ_2 & " AND (vta_pedido.fchent)<=CDate('" & TextBoxFechaPlazo2.Valor & "'))"
    
    If (TextOrdCompra.Text = "Todos") Then
        c_ORD_COMP_1 = ""
        c_ORD_COMP_2 = ""
    Else
        c_ORD_COMP_1 = " And ((ped_pedido.oc) = '" & TextOrdCompra.Text & "')"
        c_ORD_COMP_2 = " And ((Mid([vta_pedido].[numcen],11,10)) = '" & TextOrdCompra.Text & "')"
    End If
    
'SELECT ped_pedido.oc, mae_cliente.nombre, alm_inventario.id, alm_inventario.descripcion, ped_pedido.fchemi, ped_pedidodetent.fchent, ped_pedidodetent.canpro
'FROM ((ped_pedido INNER JOIN mae_cliente ON ped_pedido.idcli = mae_cliente.id) INNER JOIN ped_pedidodetent ON ped_pedido.id = ped_pedidodetent.idped) INNER JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id
'Where (((ped_pedido.oc) = "105934") And ((alm_inventario.id) = 536))
'ORDER BY ped_pedido.fchemi DESC , ped_pedidodetent.fchent
'Union
'SELECT Mid([vta_pedido].[numcen],11,10) AS oc, mae_cliente.nombre, alm_inventario.id, alm_inventario.descripcion, vta_pedido.fchemi, vta_pedido.fchent, vta_pedidodet.canpro
'FROM ((((vta_pedidodet INNER JOIN vta_pedido ON vta_pedidodet.idped = vta_pedido.id) INNER JOIN mae_productoscen ON vta_pedidodet.codpro = mae_productoscen.codcen) INNER JOIN alm_inventario ON mae_productoscen.iditem = alm_inventario.id) INNER JOIN vta_puntoVenta ON vta_pedido.idpunvecli = vta_puntoVenta.id) INNER JOIN mae_cliente ON vta_puntoVenta.idcli = mae_cliente.id
'WHERE (((Mid([vta_pedido].[numcen],11,10))=100733) AND ((alm_inventario.id)=779) AND ((vta_pedido.fchemi)>=CDate('1/1/10') And (vta_pedido.fchemi)<=CDate('31/12/10')) AND ((vta_pedido.fchent)>=CDate('1/1/10') And (vta_pedido.fchent)<=CDate('31/12/10')));
'
    'Se hace la consulta general
    c_SQL = "SELECT ped_pedido.oc, mae_cliente.nombre, alm_inventario.id, alm_inventario.descripcion, ped_pedido.fchemi, ped_pedidodetent.fchent, ped_pedidodetent.canpro " _
            + vbCr + "FROM ((ped_pedido INNER JOIN mae_cliente ON ped_pedido.idcli = mae_cliente.id) INNER JOIN ped_pedidodetent ON ped_pedido.id = ped_pedidodetent.idped) INNER JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id " _
            + vbCr + "WHERE (" & c_CLIENTES & c_PRODUCTOS & c_FECHA_EMI & c_ORD_COMP_1 & " AND " & c_FECHA_PLAZ & ")" _
            + vbCr + "ORDER BY ped_pedido.fchemi, ped_pedidodetent.fchent;"
'            + vbCr + "UNION " _
'            + vbCr + "SELECT Mid([vta_pedido].[numcen],11,10) AS oc, mae_cliente.nombre, alm_inventario.id, alm_inventario.descripcion, vta_pedido.fchemi, vta_pedido.fchent, vta_pedidodet.canpro " _
'            + vbCr + "FROM ((((vta_pedidodet INNER JOIN vta_pedido ON vta_pedidodet.idped = vta_pedido.id) INNER JOIN mae_productoscen ON vta_pedidodet.codpro = mae_productoscen.codcen) INNER JOIN alm_inventario ON mae_productoscen.iditem = alm_inventario.id) INNER JOIN vta_puntoVenta ON vta_pedido.idpunvecli = vta_puntoVenta.id) INNER JOIN mae_cliente ON vta_puntoVenta.idcli = mae_cliente.id" _
'            + vbCr + "WHERE (" & c_CLIENTES & c_PRODUCTOS & c_FECHA_EMI_2 & c_ORD_COMP_2 & " AND " & c_FECHA_PLAZ_2 & ");"
            
    RST_Busq RstLis, c_SQL, xCon
    Set VSFlexGridReporte.DataSource = RstLis.DataSource
    configurarVSFlexGridReporte
    rellenarPendientes RstLis
    Set RstLis = Nothing
    'CommandMostOcult.SetFocus
End Sub

Private Sub mostrarDetalle(fil As Integer)
    Dim c_SQL As String
    Dim c_SQL2 As String
    Dim RstLisAux As New ADODB.Recordset
    Dim RstLisAux2 As New ADODB.Recordset
    Dim prodAct As String
    Dim numOrdAct As String
    
    Dim cantPed As Double
    Dim cantEnt As Double
    
    limpiarVSFlexGridDetalle
    
    numOrdAct = VSFlexGridReporte.TextMatrix(fil, 1)
    prodAct = VSFlexGridReporte.TextMatrix(fil, 3)

    'Se hace la consulta general
    c_SQL = "SELECT ped_pedido.oc, alm_inventario.id, alm_inventario.descripcion, ped_pedido.fchemi, ped_pedidodetent.fchent, ped_pedidodetent.canpro " _
            + vbCr + "FROM ((ped_pedido INNER JOIN mae_cliente ON ped_pedido.idcli = mae_cliente.id) INNER JOIN ped_pedidodetent ON ped_pedido.id = ped_pedidodetent.idped) INNER JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id " _
            + vbCr + "WHERE (((ped_pedido.oc)= '" & numOrdAct & "') AND ((alm_inventario.id)=" & prodAct & "))" _
            + vbCr + "ORDER BY ped_pedido.fchemi, ped_pedidodetent.fchent;"
'            + vbCr + "UNION " _
'            + vbCr + "SELECT Mid([vta_pedido].[numcen],11,10) AS oc, alm_inventario.id, alm_inventario.descripcion, vta_pedido.fchemi, vta_pedido.fchent, vta_pedidodet.canpro " _
'            + vbCr + "FROM ((((vta_pedidodet INNER JOIN vta_pedido ON vta_pedidodet.idped = vta_pedido.id) INNER JOIN mae_productoscen ON vta_pedidodet.codpro = mae_productoscen.codcen) INNER JOIN alm_inventario ON mae_productoscen.iditem = alm_inventario.id) INNER JOIN vta_puntoVenta ON vta_pedido.idpunvecli = vta_puntoVenta.id) INNER JOIN mae_cliente ON vta_puntoVenta.idcli = mae_cliente.id" _
'            + vbCr + "WHERE (((Mid([vta_pedido].[numcen],11,10))= '" & numOrdAct & "') AND ((alm_inventario.id)=" & prodAct & "));"
                
    RST_Busq RstLisAux, c_SQL, xCon
    Set VSFlexGridDetalle2.DataSource = RstLisAux.DataSource

    c_SQL2 = "SELECT vta_guia.numordcom, alm_inventario.descripcion, vta_guia.fecgiro, vta_guiadet.canpro " _
            + vbCr + "FROM (vta_guiadet INNER JOIN vta_guia ON vta_guiadet.idgui = vta_guia.id) INNER JOIN alm_inventario ON vta_guiadet.iditem = alm_inventario.id " _
            + vbCr + "WHERE (((vta_guia.numordcom)= '" & numOrdAct & "') AND ((alm_inventario.id)=" & prodAct & "))" _
            + vbCr + "ORDER BY vta_guia.fecgiro, alm_inventario.descripcion;"
            
    RST_Busq RstLisAux2, c_SQL2, xCon
    Set VSFlexGridDetalle.DataSource = RstLisAux2.DataSource
    If (RstLisAux2.EOF) Then
        VSFlexGridDetalle.AddItem ("")
        VSFlexGridDetalle.TextMatrix(1, 4) = "No hay Informacion"
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
    
    configurarVSFlexGridDetalle
    LabelDetalleProducto = RstLisAux("descripcion")
    LabelDetalleID = VSFlexGridReporte.TextMatrix(fil, 3)
    LabelDetalleOC = VSFlexGridReporte.TextMatrix(fil, 1)
    LabelResto = LabelDetalleEntgda - LabelDetalleEntgr
    
    Set RstLisAux = Nothing
    Set RstLisAux2 = Nothing
    c_SQL = ""
    c_SQL2 = ""
End Sub

Private Function generarPendientes(fil As Integer) As String()
    Dim rpta(2) As String
    Dim c_SQL As String
    Dim c_SQL2 As String
    Dim RstLisAux As New ADODB.Recordset
    Dim RstLisAux2 As New ADODB.Recordset
    Dim prodAct As String
    Dim numOrdAct As String
    
    Dim cantPed As Double
    Dim cantEnt As Double
    
    limpiarVSFlexGridDetalle
    numOrdAct = VSFlexGridReporte.TextMatrix(fil, 1)
    prodAct = VSFlexGridReporte.TextMatrix(fil, 3)

'SELECT ped_pedido.oc, alm_inventario.id, Sum(ped_pedidodetent.canpro) AS SumaDecanpro, Max(ped_pedidodetent.fchent) AS MáxDefchent
'FROM (ped_pedido INNER JOIN ped_pedidodetent ON ped_pedido.id = ped_pedidodetent.idped) INNER JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id
'GROUP BY ped_pedido.oc, alm_inventario.id
'HAVING (((ped_pedido.oc)='14008382') AND ((alm_inventario.id)=1230));
'Union
'SELECT Mid([vta_pedido].[numcen],11,10) AS oc, alm_inventario.id, Sum(vta_pedidodet.canpro) AS SumaDecanpro, Max(vta_pedido.fchent) AS MáxDefchent
'FROM ((vta_pedidodet INNER JOIN vta_pedido ON vta_pedidodet.idped = vta_pedido.id) INNER JOIN mae_productoscen ON vta_pedidodet.codpro = mae_productoscen.codcen) INNER JOIN alm_inventario ON mae_productoscen.iditem = alm_inventario.id
'GROUP BY Mid([vta_pedido].[numcen],11,10), alm_inventario.id
'HAVING (((Mid([vta_pedido].[numcen],11,10))='14008382') AND ((alm_inventario.id)=1230));
    
    c_SQL = "SELECT ped_pedido.oc, alm_inventario.id, Sum(ped_pedidodetent.canpro) AS SumaDecanpro, Max(ped_pedidodetent.fchent) AS MáxDefchent " _
            + vbCr + "FROM (ped_pedido INNER JOIN ped_pedidodetent ON ped_pedido.id = ped_pedidodetent.idped) INNER JOIN alm_inventario ON ped_pedidodetent.iditem = alm_inventario.id " _
            + vbCr + "GROUP BY ped_pedido.oc, alm_inventario.id " _
            + vbCr + "HAVING (((ped_pedido.oc)= '" & numOrdAct & "') AND ((alm_inventario.id)=" & prodAct & "));"
'            + vbCr + "UNION " _
'            + vbCr + "SELECT Mid([vta_pedido].[numcen],11,10) AS oc, alm_inventario.id, Sum(vta_pedidodet.canpro) AS SumaDecanpro, Max(vta_pedido.fchent) AS MáxDefchent " _
'            + vbCr + "FROM ((vta_pedidodet INNER JOIN vta_pedido ON vta_pedidodet.idped = vta_pedido.id) INNER JOIN mae_productoscen ON vta_pedidodet.codpro = mae_productoscen.codcen) INNER JOIN alm_inventario ON mae_productoscen.iditem = alm_inventario.id " _
'            + vbCr + "GROUP BY Mid([vta_pedido].[numcen],11,10), alm_inventario.id " _
'            + vbCr + "HAVING (((Mid([vta_pedido].[numcen],11,10))= '" & numOrdAct & "') AND ((alm_inventario.id)=" & prodAct & "));"

'SELECT vta_guia.numordcom, alm_inventario.id, Sum(vta_guiadet.canpro) AS SumaDecanpro, Max(vta_guia.fecgiro) AS MáxDefecgiro
'FROM (vta_guiadet INNER JOIN vta_guia ON vta_guiadet.idgui = vta_guia.id) INNER JOIN alm_inventario ON vta_guiadet.iditem = alm_inventario.id
'GROUP BY vta_guia.numordcom, alm_inventario.id
'HAVING (((vta_guia.numordcom)='14008382') AND ((alm_inventario.id)=1230));
                
    RST_Busq RstLisAux, c_SQL, xCon
    c_SQL2 = "SELECT vta_guia.numordcom, alm_inventario.id, Sum(vta_guiadet.canpro) AS SumaDecanpro, Max(vta_guia.fecgiro) AS MáxDefecgiro " _
            + vbCr + "FROM (vta_guiadet INNER JOIN vta_guia ON vta_guiadet.idgui = vta_guia.id) INNER JOIN alm_inventario ON vta_guiadet.iditem = alm_inventario.id " _
            + vbCr + "GROUP BY vta_guia.numordcom, alm_inventario.id " _
            + vbCr + "HAVING (((vta_guia.numordcom)= '" & numOrdAct & "') AND ((alm_inventario.id)=" & prodAct & "));"
            
    RST_Busq RstLisAux2, c_SQL2, xCon
    
    If (RstLisAux2.EOF) Then
        rpta(0) = "PENDIENTE"
        If (RstLisAux("MáxDefchent") < Date) Then
            rpta(1) = "VENCIDO  " & (Date - RstLisAux("MáxDefchent")) & " dia(s)"
        Else
            rpta(1) = "POR VENCER  " & (RstLisAux("MáxDefchent") - Date) & " dia(s)"
        End If
    Else
        cantPed = CDbl(RstLisAux("SumaDecanpro"))
        cantEnt = CDbl(RstLisAux2("SumaDecanpro"))
        If ((cantEnt - cantPed) < 0) Then
            rpta(0) = "INCOMPLETO"
            If (RstLisAux("MáxDefchent") < RstLisAux2("MáxDefecgiro")) Then
                rpta(1) = "VENCIDO  " & (RstLisAux2("MáxDefecgiro") - RstLisAux("MáxDefchent")) & " dia(s)"
            Else
                rpta(1) = "POR VENCER  " & (RstLisAux("MáxDefchent") - RstLisAux2("MáxDefecgiro")) & " dia(s)"
            End If
        Else
            If ((cantEnt - cantPed) > 0) Then
                rpta(0) = "EXCEDIDO"
                If (RstLisAux("MáxDefchent") <= RstLisAux2("MáxDefecgiro")) Then
                    rpta(1) = "A TIEMPO  "
                Else
                    rpta(1) = "A DESTIEMPO  " & (RstLisAux2("MáxDefecgiro") - RstLisAux("MáxDefchent")) & " dia(s)"
                End If
            Else
                rpta(0) = "COMPLETO"
                If (RstLisAux("MáxDefchent") < RstLisAux2("MáxDefecgiro")) Then
                    rpta(1) = "A DESTIEMPO  " & (RstLisAux2("MáxDefecgiro") - RstLisAux("MáxDefchent")) & " dia(s)"
                Else
                    rpta(1) = "A TIEMPO  "
                End If
            End If
        End If
    End If
    generarPendientes = rpta
End Function


Private Sub rellenarPendientes(ByRef T1 As ADODB.Recordset)
    Dim fila As Integer
    Dim dif() As String
    fila = 1
    If (Not T1.EOF) Then
        ProgressBar1.Max = T1.RecordCount
        T1.MoveFirst
        While (Not T1.EOF)
            ProgressBar1.Value = fila
            dif = generarPendientes(fila)
            VSFlexGridReporte.TextMatrix(fila, 8) = dif(0)
            VSFlexGridReporte.TextMatrix(fila, 9) = dif(1)
            
            If (VSFlexGridReporte.TextMatrix(fila, 8) = "PENDIENTE") Then
                VSFlexGridReporte.Select fila, 8
                VSFlexGridReporte.CellForeColor = &HFF&
            Else
                If (VSFlexGridReporte.TextMatrix(fila, 8) = "COMPLETO") Then
                    VSFlexGridReporte.Select fila, 8
                    VSFlexGridReporte.CellForeColor = &H8000000D
                Else
                    VSFlexGridReporte.Select fila, 8
                    VSFlexGridReporte.CellForeColor = &HC000&
                End If
            End If
            fila = fila + 1
            T1.MoveNext
        Wend
    End If
End Sub

Private Sub iniciarCampos()
    Dim fechAct As Date
    fechAct = Date
    presionado = False
    Set VSFlexGridClientes.DataSource = Nothing
    Set VSFlexGridProductos.DataSource = Nothing
    Set VSFlexGridDetalle.DataSource = Nothing
    Set VSFlexGridDetalle2.DataSource = Nothing
    'Se inicializa:
    'datos para productos
    VSFlexGridProductos.Rows = 1
    VSFlexGridProductos.Cols = 2
    VSFlexGridProductos.Row = 0
    VSFlexGridProductos.Col = 1
    VSFlexGridProductos.Text = "Todos"
    'datos para clientes
    VSFlexGridClientes.Rows = 1
    VSFlexGridClientes.Cols = 2
    VSFlexGridClientes.Row = 0
    VSFlexGridClientes.Col = 1
    VSFlexGridClientes.Text = "Todos"
    
    TextOrdCompra.Text = "Todos"
    'datos para detalles
    TextBoxFechaEmision1.Valor = CDate("01/01/" + CStr(Year(Date)))
    TextBoxFechaEmision2.Valor = Date
    TextBoxFechaPlazo1.Valor = CDate("01/" + CStr(Month(Date)) + "/" + CStr(Year(Date)))
    TextBoxFechaPlazo2.Valor = Date + 150
    
    VSFlexGridClientes.Editable = flexEDKbdMouse
    VSFlexGridClientes.ColComboList(1) = "..."
    VSFlexGridClientes.ShowComboButton = flexSBAlways
    
    VSFlexGridProductos.Editable = flexEDKbdMouse
    VSFlexGridProductos.ColComboList(1) = "..."
    VSFlexGridProductos.ShowComboButton = flexSBAlways
    
    FrameEspecificaciones.Top = -30
    FrameEspecificaciones.Left = 35
    FrameEspecificaciones.Width = 11805
    FrameEspecificaciones.Height = 4550
    
    VSFlexGridReporte.AllowUserResizing = flexResizeColumns
    VSFlexGridReporte.AutoSearch = flexSearchFromTop
    VSFlexGridReporte.ExplorerBar = flexExSortShowAndMove

End Sub

Private Sub Command2_Click()
    iniciarCampos
End Sub

Private Sub configurarVSFlexGridReporte()
    VSFlexGridReporte.Cols = 10
    VSFlexGridReporte.FixedRows = 1
    VSFlexGridReporte.FixedCols = 5
    VSFlexGridReporte.ColWidth(0) = 250
    VSFlexGridReporte.RowHeight(0) = 500
    VSFlexGridReporte.Row = 0
    
    
    VSFlexGridReporte.Col = 1
    VSFlexGridReporte.Text = "ORD. COMP."
    VSFlexGridReporte.ColWidth(1) = 0
    VSFlexGridReporte.CellFontBold = True
    
    VSFlexGridReporte.Col = 2
    VSFlexGridReporte.Text = "CLIENTE"
    VSFlexGridReporte.ColWidth(2) = 2300
    VSFlexGridReporte.CellFontBold = True
    
    VSFlexGridReporte.Col = 3
    VSFlexGridReporte.Text = "ID"
    VSFlexGridReporte.ColWidth(3) = 0
    VSFlexGridReporte.CellFontBold = True
    
    VSFlexGridReporte.Col = 4
    VSFlexGridReporte.Text = "PRODUCTO"
    VSFlexGridReporte.ColWidth(4) = 4800
    VSFlexGridReporte.CellFontBold = True
    
    VSFlexGridReporte.Col = 5
    VSFlexGridReporte.Text = "FECH. EMISION"
    VSFlexGridReporte.ColWidth(5) = 1950
    VSFlexGridReporte.CellFontBold = True
    
    VSFlexGridReporte.Col = 6
    VSFlexGridReporte.Text = "FECH. A ENTREGAR"
    VSFlexGridReporte.ColWidth(6) = 1950
    VSFlexGridReporte.CellFontBold = True
    
    VSFlexGridReporte.Col = 7
    VSFlexGridReporte.Text = "CANT. A ENTREGAR"
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

Private Sub configurarVSFlexGridDetalle()
    'detalle 1
    VSFlexGridDetalle2.Cols = 7
    VSFlexGridDetalle2.FixedRows = 1
    VSFlexGridDetalle2.FixedCols = 4
    VSFlexGridDetalle2.ColWidth(0) = 250
    VSFlexGridDetalle2.RowHeight(0) = 500
    VSFlexGridDetalle2.Row = 0
    
    VSFlexGridDetalle2.Col = 1
    VSFlexGridDetalle2.Text = "ORD. COMP."
    VSFlexGridDetalle2.ColWidth(1) = 0
    VSFlexGridDetalle2.CellFontBold = True
    
    VSFlexGridDetalle2.Col = 2
    VSFlexGridDetalle2.Text = "CLIENTE"
    VSFlexGridDetalle2.ColWidth(2) = 0
    VSFlexGridDetalle2.CellFontBold = True
    
    VSFlexGridDetalle2.Col = 3
    VSFlexGridDetalle2.Text = "PRODUCTO"
    VSFlexGridDetalle2.ColWidth(3) = 0
    VSFlexGridDetalle2.CellFontBold = True
    
    VSFlexGridDetalle2.Col = 4
    VSFlexGridDetalle2.Text = "FECH. EMISION"
    VSFlexGridDetalle2.ColWidth(4) = 1950
    VSFlexGridDetalle2.CellFontBold = True
    
    VSFlexGridDetalle2.Col = 5
    VSFlexGridDetalle2.Text = "FECH. A ENTREGAR"
    VSFlexGridDetalle2.ColWidth(5) = 1950
    VSFlexGridDetalle2.CellFontBold = True
    
    VSFlexGridDetalle2.Col = 6
    VSFlexGridDetalle2.Text = "CANT. A ENTREGAR"
    VSFlexGridDetalle2.ColWidth(6) = 1950
    VSFlexGridDetalle2.CellFontBold = True
    
    'detalle 2
    VSFlexGridDetalle.Cols = 5
    VSFlexGridDetalle.FixedRows = 1
    VSFlexGridDetalle.FixedCols = 3
    VSFlexGridDetalle.ColWidth(0) = 250
    VSFlexGridDetalle.RowHeight(0) = 500
    VSFlexGridDetalle.Row = 0
    
    VSFlexGridDetalle.Col = 1
    VSFlexGridDetalle.Text = "ORD. COMP."
    VSFlexGridDetalle.ColWidth(1) = 0
    VSFlexGridDetalle.CellFontBold = True
    
    VSFlexGridDetalle.Col = 2
    VSFlexGridDetalle.Text = "PRODUCTO"
    VSFlexGridDetalle.ColWidth(2) = 0
    VSFlexGridDetalle.CellFontBold = True
    
    VSFlexGridDetalle.Col = 3
    VSFlexGridDetalle.Text = "FECH. ENTREGADA"
    VSFlexGridDetalle.ColWidth(3) = 1950
    VSFlexGridDetalle.CellFontBold = True
    
    VSFlexGridDetalle.Col = 4
    VSFlexGridDetalle.Text = "CANT. ENTREGADA"
    VSFlexGridDetalle.ColWidth(4) = 1950
    VSFlexGridDetalle.CellFontBold = True

End Sub

Private Sub limpiarVSFlexGridReporte()
    VSFlexGridReporte.Cols = 2
    VSFlexGridReporte.Rows = 2
    VSFlexGridReporte.FixedRows = 1
    VSFlexGridReporte.FixedCols = 1
End Sub

Private Sub limpiarVSFlexGridDetalle()
    VSFlexGridDetalle.Cols = 2
    VSFlexGridDetalle.Rows = 2
    VSFlexGridDetalle.FixedRows = 1
    VSFlexGridDetalle.FixedCols = 1
    
    VSFlexGridDetalle2.Cols = 2
    VSFlexGridDetalle2.Rows = 2
    VSFlexGridDetalle2.FixedRows = 1
    VSFlexGridDetalle2.FixedCols = 1
End Sub

Private Sub CommandLimpiar_Click()
    limpio = True
    iniciarCampos
    limpio = False
End Sub

Private Sub CommandMostOcult_Click()
    presionado = Not presionado
    If (presionado) Then
        VSFlexGridReporte.ColWidth(1) = 2000
    Else
        VSFlexGridReporte.ColWidth(1) = 0
    End If
End Sub

Private Sub Form_Activate()
    generarConsulta
    ProgressBar1.Visible = False
    cargo = True
End Sub

Private Sub Form_Load()
    limpio = False
    cargo = False
    cambioTextPlazo1 = False
    iniciarCampos
    Main
End Sub

Private Sub LabelCerrar_Click()
    FrameEspecificaciones.Visible = False
End Sub

Private Sub TextBoxFechaEmision1_Change()
    If (TextBoxFechaEmision1.Valor = "") Then TextBoxFechaEmision1.Valor = fechaEmision1
    If cargo Then
        If Not limpio Then
            If (fechaEmision1 <> TextBoxFechaEmision1.Valor) Then
                ProgressBar1.Visible = True
                generarConsulta
                ProgressBar1.Visible = False
                CommandMostOcult.SetFocus
            End If
        End If
    End If
    fechaEmision1 = TextBoxFechaEmision1.Valor
End Sub

Private Sub TextBoxFechaEmision2_Change()
    If (TextBoxFechaEmision2.Valor = "") Then TextBoxFechaEmision2.Valor = fechaEmision2
    If cargo Then
        If Not limpio Then
            If (fechaEmision2 <> TextBoxFechaEmision2.Valor) Then
                ProgressBar1.Visible = True
                generarConsulta
                ProgressBar1.Visible = False
                CommandMostOcult.SetFocus
            End If
        End If
    End If
    fechaEmision2 = TextBoxFechaEmision2.Valor
End Sub

Private Sub TextBoxFechaPlazo1_Change()
    If (TextBoxFechaPlazo1.Valor = "") Then TextBoxFechaPlazo1.Valor = fechaPlazo1
    If cargo Then
        If (Not limpio) Then
            If (fechaPlazo1 <> TextBoxFechaPlazo1.Valor) Then
                cambioTextPlazo1 = True
                ProgressBar1.Visible = True
                generarConsulta
                ProgressBar1.Visible = False
                CommandMostOcult.SetFocus
            End If
        Else
            If (fechaPlazo1 <> TextBoxFechaPlazo1.Valor) Then
                cambioTextPlazo1 = True
                ProgressBar1.Visible = True
                generarConsulta
                ProgressBar1.Visible = False
                CommandMostOcult.SetFocus
            End If
        End If
    End If
    fechaPlazo1 = TextBoxFechaPlazo1.Valor
End Sub

Private Sub TextBoxFechaPlazo2_Change()
    If (TextBoxFechaPlazo2.Valor = "") Then TextBoxFechaPlazo2.Valor = fechaPlazo2
    If cargo Then
        If Not limpio Then
            If (fechaPlazo2 <> TextBoxFechaPlazo2.Valor) Then
                ProgressBar1.Visible = True
                generarConsulta
                ProgressBar1.Visible = False
                CommandMostOcult.SetFocus
            End If
        Else
            If Not cambioTextPlazo1 Then
                If (fechaPlazo2 <> TextBoxFechaPlazo2.Valor) Then
                    ProgressBar1.Visible = True
                    generarConsulta
                    ProgressBar1.Visible = False
                    CommandMostOcult.SetFocus
                End If
            End If
        End If
    End If
    fechaPlazo2 = TextBoxFechaPlazo2.Valor
End Sub

Private Sub TextOrdCompra_Change()
    If cargo Then
        If Not limpio Then
            ProgressBar1.Visible = True
            generarConsulta
            ProgressBar1.Visible = False
        Else
            ProgressBar1.Visible = True
            generarConsulta
            ProgressBar1.Visible = False
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 5 Then
        ExportarExcel
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
    'Dim xCampos(3, 3) As String
    
    Dim xCampos(2, 3) As String
    xCampos(0, 0) = "Nombre":          xCampos(0, 1) = "nombre":     xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.UC.":        xCampos(1, 1) = "numruc":     xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
    
    nSQL = "SELECT mae_cliente.nombre, mae_cliente.numruc " _
           + vbCr + "From mae_cliente " _
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
        generarConsulta
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub VSFlexGridClientes_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If cargo Then
        If Not limpio Then
            ProgressBar1.Visible = True
            generarConsulta
            ProgressBar1.Visible = False
        Else
            ProgressBar1.Visible = True
            generarConsulta
            ProgressBar1.Visible = False
        End If
    End If
End Sub

Private Sub VSFlexGridProductos_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    VSFlexGridProductos.ShowComboButton = flexSBFocus
    
    Dim nSQL As String
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 3) As String
    
    xCampos(0, 0) = "Descripción":      xCampos(0, 1) = "insdesc":     xCampos(0, 2) = "5000":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "Tip. Prod":        xCampos(1, 1) = "tipprodesc":      xCampos(1, 2) = "1300":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "Familia":          xCampos(2, 1) = "famdesc":      xCampos(2, 2) = "1500":   xCampos(2, 3) = "C"
    
    nSQL = "SELECT DISTINCT alm_inventario.id, alm_inventario.descripcion AS insdesc, mae_tipoproducto.descripcion AS tipprodesc, mae_familia.descripcion AS famdesc " _
         + vbCr + "FROM ((ges_plaprod INNER JOIN (mae_unidades INNER JOIN (ges_plaproddet INNER JOIN alm_inventario ON ges_plaproddet.codpro = alm_inventario.id) ON mae_unidades.id = alm_inventario.idunimed) ON ges_plaprod.id = ges_plaproddet.idpv) INNER JOIN mae_familia ON alm_inventario.idfam = mae_familia.id) INNER JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id " _
         + vbCr + "Where (((ges_plaprod.activo) = -1) And ((ges_plaproddet.idmes) <> 13)) " _
         + vbCr + "Union " _
         + vbCr + "SELECT DISTINCT alm_inventario.id, alm_inventario.descripcion AS insdesc, mae_tipoproducto.descripcion AS tipprodesc, mae_familia.descripcion AS famdesc " _
         + vbCr + "FROM ((mae_unidades RIGHT JOIN (ges_plaprod INNER JOIN (ges_plaproddet2 INNER JOIN alm_inventario ON ges_plaproddet2.codpro = alm_inventario.id) ON ges_plaprod.id = ges_plaproddet2.idpv) ON mae_unidades.id = alm_inventario.idunimed) INNER JOIN mae_tipoproducto ON alm_inventario.tippro = mae_tipoproducto.id) INNER JOIN mae_familia ON alm_inventario.idfam = mae_familia.id " _
         + vbCr + "WHERE (((ges_plaprod.activo)=-1) AND ((ges_plaproddet2.idmes)<>13));"
         
    xform.SQLCad = nSQL
    
    xform.Titulo = "Buscando Productos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "insdesc"
    xform.CampoBusca = "insdesc"
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        VSFlexGridProductos.TextMatrix(VSFlexGridProductos.Row, 1) = xRs.Fields(1) & ""
        If VSFlexGridProductos.Row = VSFlexGridProductos.Rows - 1 Then VSFlexGridProductos.AddItem ""
        generarConsulta
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub VSFlexGridProductos_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If cargo Then
        If Not limpio Then
            ProgressBar1.Visible = True
            generarConsulta
            ProgressBar1.Visible = False
        Else
            ProgressBar1.Visible = True
            generarConsulta
            ProgressBar1.Visible = False
        End If
    End If
End Sub


Sub ExportarExcel()
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
        .cells(4, 5) = "Productos: "
        Dim xFilasAux As Integer
        xFilasAux = 5
        For A = 0 To VSFlexGridProductos.Rows - 1
            .cells(xFilasAux, 6) = VSFlexGridProductos.TextMatrix(A, 1)
            xFilasAux = xFilasAux + 1
        Next A
        
        If (xFilas < xFilasAux) Then xFilas = xFilasAux
        
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

Private Sub VSFlexGridReporte_DblClick()
    If (VSFlexGridReporte.Row <> 0) Then
        mostrarDetalle VSFlexGridReporte.Row
        FrameEspecificaciones.Visible = True
    End If
End Sub
