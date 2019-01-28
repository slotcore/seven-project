VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmManGuias 
   Caption         =   "Ventas - Emision de Guias"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11865
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7560
   ScaleWidth      =   11865
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Fraseldoc 
      BorderStyle     =   0  'None
      Caption         =   "Documento"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   1570
      Left            =   8370
      TabIndex        =   116
      Top             =   9060
      Visible         =   0   'False
      Width           =   5595
      Begin VB.CommandButton CmdBusNumSer2 
         Height          =   240
         Left            =   2085
         Picture         =   "FrmManGuias.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   122
         Top             =   780
         Width           =   240
      End
      Begin VB.Frame Frame8 
         Height          =   930
         Left            =   3150
         TabIndex        =   119
         Top             =   450
         Width           =   2280
         Begin VB.CommandButton cmdokseldoc 
            Height          =   510
            Left            =   90
            Picture         =   "FrmManGuias.frx":0132
            Style           =   1  'Graphical
            TabIndex        =   121
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmdsalirseldoc 
            Height          =   510
            Left            =   1140
            Picture         =   "FrmManGuias.frx":24B8
            Style           =   1  'Graphical
            TabIndex        =   120
            Top             =   240
            Width           =   1050
         End
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmiAnul 
         Height          =   300
         Left            =   1440
         TabIndex        =   117
         Top             =   420
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
      End
      Begin VB.TextBox TxtNumSer2 
         Height          =   300
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   123
         Text            =   "TxtNumSer2"
         Top             =   750
         Width           =   915
      End
      Begin VB.TextBox TxtNumDocGen 
         Height          =   300
         Left            =   1425
         MaxLength       =   10
         TabIndex        =   118
         Text            =   "TxtNumDocGen"
         Top             =   1065
         Width           =   1335
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   0
         X2              =   5550
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   3
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   2235
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   2
         X1              =   5550
         X2              =   5550
         Y1              =   0
         Y2              =   1530
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Emisión de Documentos Anulados"
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
         Left            =   135
         TabIndex        =   127
         Top             =   105
         Width           =   2880
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Nº Serie"
         Height          =   195
         Left            =   165
         TabIndex        =   126
         Top             =   780
         Width           =   585
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   2
         X1              =   30
         X2              =   7440
         Y1              =   1530
         Y2              =   1530
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   165
         TabIndex        =   125
         Top             =   1095
         Width           =   1050
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Documento"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   124
         Top             =   480
         Width           =   1185
      End
      Begin VB.Shape Shape6 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Left            =   45
         Top             =   45
         Width           =   5475
      End
   End
   Begin VB.Frame Frame7 
      BorderStyle     =   0  'None
      Caption         =   "Frame7"
      Height          =   5520
      Left            =   11970
      TabIndex        =   80
      Top             =   450
      Visible         =   0   'False
      Width           =   9105
      Begin VB.CommandButton CmdBusCli 
         Height          =   240
         Left            =   8760
         Picture         =   "FrmManGuias.frx":430A
         Style           =   1  'Graphical
         TabIndex        =   92
         Top             =   465
         Width           =   240
      End
      Begin VB.TextBox TxtCliente 
         Height          =   300
         Left            =   1275
         TabIndex        =   90
         Text            =   "TxtCliente"
         Top             =   435
         Width           =   7755
      End
      Begin VB.CommandButton CmdAcepta22 
         Caption         =   "&Aceptar"
         Height          =   420
         Left            =   3420
         TabIndex        =   83
         Top             =   4980
         Width           =   1125
      End
      Begin VB.CommandButton CmdCancela22 
         Caption         =   "&Cancelar"
         Height          =   420
         Left            =   4680
         TabIndex        =   82
         Top             =   4995
         Width           =   1125
      End
      Begin VB.CommandButton CmdBusPro22 
         Height          =   240
         Left            =   8760
         Picture         =   "FrmManGuias.frx":443C
         Style           =   1  'Graphical
         TabIndex        =   81
         Top             =   2820
         Width           =   240
      End
      Begin VSFlex7Ctl.VSFlexGrid fg5 
         Height          =   1710
         Left            =   75
         TabIndex        =   85
         Top             =   1005
         Width           =   8955
         _cx             =   15796
         _cy             =   3016
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
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmManGuias.frx":456E
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
      Begin VSFlex7Ctl.VSFlexGrid Fg6 
         Height          =   1710
         Left            =   75
         TabIndex        =   86
         Top             =   3180
         Width           =   8955
         _cx             =   15796
         _cy             =   3016
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
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmManGuias.frx":4634
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
      Begin VB.TextBox TxtProducto 
         Height          =   300
         Left            =   1275
         TabIndex        =   84
         Text            =   "TxtProducto"
         Top             =   2790
         Width           =   7755
      End
      Begin VB.Label LblIdProd2 
         AutoSize        =   -1  'True
         Caption         =   "LblIdProd2"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   5385
         TabIndex        =   94
         Top             =   780
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.Label LblIdcli2 
         AutoSize        =   -1  'True
         Caption         =   "LblIdcli2"
         ForeColor       =   &H000000FF&
         Height          =   195
         Left            =   4545
         TabIndex        =   93
         Top             =   765
         Visible         =   0   'False
         Width           =   585
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Left            =   75
         TabIndex        =   91
         Top             =   465
         Width           =   480
      End
      Begin VB.Line Line14 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   9150
         Y1              =   5505
         Y2              =   5505
      End
      Begin VB.Line Line14 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   9075
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line15 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   9090
         X2              =   9090
         Y1              =   15
         Y2              =   5520
      End
      Begin VB.Line Line15 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   0
         Y1              =   0
         Y2              =   5475
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Correccion Item de Guias"
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
         Left            =   165
         TabIndex        =   89
         Top             =   105
         Width           =   2160
      End
      Begin VB.Label Label17 
         Caption         =   "Guias a Modificar"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   75
         TabIndex        =   88
         Top             =   780
         Width           =   1635
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "Producto"
         Height          =   195
         Left            =   75
         TabIndex        =   87
         Top             =   2820
         Width           =   645
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00400000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Left            =   45
         Top             =   45
         Width           =   9015
      End
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Height          =   3630
      Left            =   7920
      TabIndex        =   60
      Top             =   7620
      Visible         =   0   'False
      Width           =   2970
      Begin VB.CommandButton CmdSalir 
         Height          =   660
         Left            =   1500
         Picture         =   "FrmManGuias.frx":472C
         Style           =   1  'Graphical
         TabIndex        =   64
         Top             =   2865
         Width           =   720
      End
      Begin VB.CommandButton CmdOk 
         Height          =   660
         Left            =   735
         Picture         =   "FrmManGuias.frx":4A36
         Style           =   1  'Graphical
         TabIndex        =   63
         Top             =   2865
         Width           =   720
      End
      Begin MSComCtl2.MonthView MonthView1 
         Height          =   2370
         Left            =   210
         TabIndex        =   61
         Top             =   420
         Width           =   2595
         _ExtentX        =   4577
         _ExtentY        =   4180
         _Version        =   393216
         ForeColor       =   -2147483630
         BackColor       =   -2147483633
         Appearance      =   1
         StartOfWeek     =   85983234
         CurrentDate     =   38919
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Seleccionar Mes"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   195
         Left            =   150
         TabIndex        =   62
         Top             =   75
         Width           =   1425
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H80000002&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H80000002&
         Height          =   315
         Left            =   30
         Top             =   30
         Width           =   2910
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   0
         X2              =   2955
         Y1              =   3615
         Y2              =   3615
      End
      Begin VB.Line Line5 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   -30
         X2              =   -30
         Y1              =   -360
         Y2              =   3240
      End
      Begin VB.Line Line4 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         X1              =   2955
         X2              =   2955
         Y1              =   15
         Y2              =   3630
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   0
         X2              =   2955
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.Frame Fradocsproc 
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
      ForeColor       =   &H00000080&
      Height          =   3495
      Left            =   11970
      TabIndex        =   69
      Top             =   6030
      Visible         =   0   'False
      Width           =   3825
      Begin VB.CommandButton cmdSalirdocsproc 
         Height          =   630
         Left            =   2625
         Picture         =   "FrmManGuias.frx":4D40
         Style           =   1  'Graphical
         TabIndex        =   72
         Top             =   2760
         Width           =   750
      End
      Begin VB.CommandButton cmdOKdocsproc 
         Height          =   630
         Left            =   480
         Picture         =   "FrmManGuias.frx":504A
         Style           =   1  'Graphical
         TabIndex        =   71
         Top             =   2760
         Width           =   750
      End
      Begin VB.CommandButton cmdEliminarOKdocsproc 
         Height          =   630
         Left            =   1260
         Picture         =   "FrmManGuias.frx":5354
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   2760
         Width           =   765
      End
      Begin VSFlex7Ctl.VSFlexGrid fgdocsproc 
         Height          =   2190
         Left            =   135
         TabIndex        =   73
         Top             =   450
         Width           =   3555
         _cx             =   6271
         _cy             =   3863
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
         Rows            =   50
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmManGuias.frx":5456
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
         Caption         =   "Guias Adjuntas"
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
         Left            =   195
         TabIndex        =   78
         Top             =   90
         Width           =   1290
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Left            =   45
         Top             =   30
         Width           =   3735
      End
      Begin VB.Line Line13 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         X1              =   3810
         X2              =   3810
         Y1              =   15
         Y2              =   3510
      End
      Begin VB.Line Line11 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   3810
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line12 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   3465
      End
      Begin VB.Line Line11 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   0
         X1              =   -15
         X2              =   3780
         Y1              =   3480
         Y2              =   3480
      End
   End
   Begin VB.Frame fraconsdocref 
      BorderStyle     =   0  'None
      Caption         =   "Documentos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   4470
      Left            =   30
      TabIndex        =   65
      Top             =   7590
      Visible         =   0   'False
      Width           =   7800
      Begin VB.CommandButton CmdOkRef 
         Height          =   630
         Left            =   3120
         Picture         =   "FrmManGuias.frx":54D0
         Style           =   1  'Graphical
         TabIndex        =   67
         Top             =   3720
         Width           =   750
      End
      Begin VB.CommandButton CmdSalirRef 
         Height          =   630
         Left            =   3915
         Picture         =   "FrmManGuias.frx":57DA
         Style           =   1  'Graphical
         TabIndex        =   66
         Top             =   3720
         Width           =   750
      End
      Begin VSFlex7Ctl.VSFlexGrid Fgdocref 
         Height          =   3150
         Left            =   105
         TabIndex        =   68
         Top             =   450
         Width           =   7590
         _cx             =   13388
         _cy             =   5556
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmManGuias.frx":5AE4
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
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Guias a Adicionar"
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
         Left            =   240
         TabIndex        =   77
         Top             =   90
         Width           =   1515
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   300
         Left            =   60
         Top             =   45
         Width           =   7695
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   7905
         Y1              =   4455
         Y2              =   4455
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   1
         X1              =   30
         X2              =   30
         Y1              =   0
         Y2              =   4470
      End
      Begin VB.Line Line10 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   7785
         X2              =   7785
         Y1              =   15
         Y2              =   4485
      End
      Begin VB.Line Line9 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   7770
         Y1              =   15
         Y2              =   15
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8250
      Top             =   45
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManGuias.frx":5BCC
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManGuias.frx":6110
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManGuias.frx":64A2
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManGuias.frx":6626
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManGuias.frx":6A7A
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManGuias.frx":6B92
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManGuias.frx":70D6
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManGuias.frx":761A
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManGuias.frx":772E
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManGuias.frx":7842
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManGuias.frx":7C96
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmManGuias.frx":7E02
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7215
      Left            =   0
      TabIndex        =   20
      Top             =   360
      Width           =   11880
      _cx             =   20955
      _cy             =   12726
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9
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
      FrontTabForeColor=   8388608
      Caption         =   "  &Consulta  |   &Detalle  "
      Align           =   0
      CurrTab         =   0
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
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6795
         Left            =   45
         TabIndex        =   56
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   57
            Top             =   345
            Width           =   11745
            _ExtentX        =   20717
            _ExtentY        =   11324
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Id"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   0
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Nº Guia"
            Columns(1).DataField=   "numguia"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Emi"
            Columns(2).DataField=   "fecgiro1"
            Columns(2).NumberFormat=   "Short Date"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Cliente"
            Columns(3).DataField=   "nombre"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Nº Pedido"
            Columns(4).DataField=   "numordcom"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Punto venta"
            Columns(5).DataField=   "despunven"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Condicion"
            Columns(6).DataField=   "Anulado2"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Doc Ref"
            Columns(7).DataField=   "numdocref"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   8
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=8"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=2434"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=2355"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1879"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1799"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=4180"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=4101"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=512"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=2858"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=2778"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=3916"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=3836"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=1905"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1826"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=2540"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2461"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
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
            HeadLines       =   1
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
            _StyleDefs(6)   =   "Style:id=1,.parent=0,.namedParent=33,.bgcolor=&HE0FEFE&,.fgcolor=&H0&,.bold=0"
            _StyleDefs(7)   =   ":id=1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(8)   =   ":id=1,.fontname=MS Sans Serif"
            _StyleDefs(9)   =   "CaptionStyle:id=4,.parent=2,.namedParent=37"
            _StyleDefs(10)  =   "HeadingStyle:id=2,.parent=1,.namedParent=34,.alignment=2,.fgcolor=&H800000&"
            _StyleDefs(11)  =   ":id=2,.bold=-1,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=54,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=51,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=52,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=53,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=28,.parent=13"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=25,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=26,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=27,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=43,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=44,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=45,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=47,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=48,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=49,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=29,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=30,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=31,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=55,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=56,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=57,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=62,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=59,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=60,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=61,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=66,.parent=13"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=63,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=64,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=65,.parent=17"
            _StyleDefs(68)  =   "Named:id=33:Normal"
            _StyleDefs(69)  =   ":id=33,.parent=0"
            _StyleDefs(70)  =   "Named:id=34:Heading"
            _StyleDefs(71)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(72)  =   ":id=34,.wraptext=-1"
            _StyleDefs(73)  =   "Named:id=35:Footing"
            _StyleDefs(74)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(75)  =   "Named:id=36:Selected"
            _StyleDefs(76)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(77)  =   "Named:id=37:Caption"
            _StyleDefs(78)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(79)  =   "Named:id=38:HighlightRow"
            _StyleDefs(80)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(81)  =   "Named:id=39:EvenRow"
            _StyleDefs(82)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(83)  =   "Named:id=40:OddRow"
            _StyleDefs(84)  =   ":id=40,.parent=33"
            _StyleDefs(85)  =   "Named:id=41:RecordSelector"
            _StyleDefs(86)  =   ":id=41,.parent=34"
            _StyleDefs(87)  =   "Named:id=42:FilterBar"
            _StyleDefs(88)  =   ":id=42,.parent=33"
         End
         Begin VB.Label LblMes 
            AutoSize        =   -1  'True
            Caption         =   "LblMes"
            BeginProperty Font 
               Name            =   "Times New Roman"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   285
            Left            =   9120
            TabIndex        =   101
            Top             =   30
            Width           =   735
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label2 
            Alignment       =   2  'Center
            Caption         =   "Consulta Guias de Remisión"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   90
            TabIndex        =   58
            Top             =   45
            Width           =   11610
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6795
         Left            =   12525
         TabIndex        =   55
         Top             =   375
         Width           =   11790
         Begin VB.CommandButton CmdBusAlm 
            Height          =   240
            Left            =   8700
            Picture         =   "FrmManGuias.frx":834A
            Style           =   1  'Graphical
            TabIndex        =   128
            Top             =   990
            Width           =   240
         End
         Begin VB.CommandButton cmd 
            Caption         =   "&Elim. Todos Item"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   10320
            TabIndex        =   115
            Top             =   6420
            Width           =   1440
         End
         Begin VB.CommandButton cmd 
            Caption         =   "&Sel Item"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   10320
            TabIndex        =   114
            Top             =   5700
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.Frame Frame10 
            Height          =   465
            Left            =   9720
            TabIndex        =   102
            Top             =   300
            Width           =   2115
            Begin VB.Label LblPeriodo2 
               Alignment       =   2  'Center
               Caption         =   "LblPeriodo2"
               BeginProperty Font 
                  Name            =   "Times New Roman"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   -1  'True
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   300
               Left            =   90
               TabIndex        =   103
               Top             =   120
               Width           =   1860
            End
         End
         Begin VB.CommandButton CmdAddEntrega 
            Caption         =   "Agregar Entrega"
            Enabled         =   0   'False
            Height          =   315
            Left            =   10320
            TabIndex        =   99
            Top             =   3900
            Width           =   1440
         End
         Begin VB.CommandButton CmdBusIdTipDocRef 
            Height          =   240
            Left            =   1995
            Picture         =   "FrmManGuias.frx":847C
            Style           =   1  'Graphical
            TabIndex        =   95
            Top             =   1320
            Width           =   240
         End
         Begin VB.CommandButton cmdbuscotizacion 
            Enabled         =   0   'False
            Height          =   240
            Left            =   10080
            Picture         =   "FrmManGuias.frx":85AE
            Style           =   1  'Graphical
            TabIndex        =   75
            Top             =   1350
            Width           =   240
         End
         Begin VB.TextBox TxtNumCotizacion 
            Height          =   300
            Left            =   8265
            Locked          =   -1  'True
            TabIndex        =   7
            Text            =   "txtnumcotizacion"
            Top             =   1305
            Width           =   2085
         End
         Begin VB.CommandButton cmdcotizprocesadas 
            Caption         =   "Ver Cotizac Adic."
            Height          =   315
            Left            =   10320
            TabIndex        =   74
            Top             =   3540
            Visible         =   0   'False
            Width           =   1440
         End
         Begin VB.CommandButton CmdBusTipItem 
            Height          =   240
            Left            =   1995
            Picture         =   "FrmManGuias.frx":86E0
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   375
            Width           =   240
         End
         Begin VB.TextBox TxtTipItem 
            Height          =   300
            Left            =   1350
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   0
            Text            =   "TxtTipItem"
            Top             =   345
            Width           =   915
         End
         Begin VB.CommandButton CmdDatosAdicion 
            Caption         =   "&Datos Adicionales"
            Enabled         =   0   'False
            Height          =   315
            Left            =   10320
            TabIndex        =   52
            Top             =   4260
            Width           =   1440
         End
         Begin VB.CommandButton CmdDelItem 
            Caption         =   "&Eliminar Item"
            Enabled         =   0   'False
            Height          =   315
            Left            =   10320
            TabIndex        =   19
            Top             =   6090
            Width           =   1440
         End
         Begin VB.CommandButton CmdAddItem 
            Caption         =   "&Agregar Item"
            Enabled         =   0   'False
            Height          =   315
            Left            =   10320
            TabIndex        =   18
            Top             =   5370
            Width           =   1440
         End
         Begin VB.CommandButton CmdPlaCar 
            Enabled         =   0   'False
            Height          =   240
            Left            =   11070
            Picture         =   "FrmManGuias.frx":8812
            Style           =   1  'Graphical
            TabIndex        =   48
            Top             =   2820
            Width           =   240
         End
         Begin VB.CommandButton CmdCho 
            Enabled         =   0   'False
            Height          =   240
            Left            =   11400
            Picture         =   "FrmManGuias.frx":8944
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   2505
            Width           =   240
         End
         Begin VB.CommandButton CmdEmpTra 
            Enabled         =   0   'False
            Height          =   240
            Left            =   11400
            Picture         =   "FrmManGuias.frx":8A76
            Style           =   1  'Graphical
            TabIndex        =   43
            Top             =   2205
            Width           =   240
         End
         Begin VB.CommandButton CmdMot 
            Enabled         =   0   'False
            Height          =   240
            Left            =   1995
            Picture         =   "FrmManGuias.frx":8BA8
            Style           =   1  'Graphical
            TabIndex        =   30
            Top             =   1635
            Width           =   240
         End
         Begin VB.TextBox TxtMotivo 
            Height          =   300
            Left            =   1350
            Locked          =   -1  'True
            TabIndex        =   8
            Text            =   "TxtMotivo"
            Top             =   1605
            Width           =   915
         End
         Begin VB.Frame Frame4 
            Caption         =   "[ Datos de Transporte ]"
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
            Height          =   1560
            Left            =   5925
            TabIndex        =   41
            Top             =   1905
            Width           =   5835
            Begin VB.CommandButton CmdAddChofer 
               Caption         =   "Adicionar &Chofer"
               Enabled         =   0   'False
               Height          =   300
               Left            =   120
               TabIndex        =   100
               Top             =   1200
               Width           =   1830
            End
            Begin VB.TextBox TxtNumPlaCar 
               Height          =   300
               Left            =   3960
               Locked          =   -1  'True
               TabIndex        =   17
               Text            =   "TxtNumPlaCar"
               Top             =   885
               Width           =   1455
            End
            Begin VB.TextBox TxtNumBre 
               Height          =   300
               Left            =   1215
               Locked          =   -1  'True
               TabIndex        =   16
               Text            =   "TxtNumBre"
               Top             =   885
               Width           =   1455
            End
            Begin VB.TextBox TxtChofer 
               Height          =   300
               Left            =   1215
               Locked          =   -1  'True
               TabIndex        =   15
               Text            =   "TxtChofer"
               Top             =   570
               Width           =   4530
            End
            Begin VB.TextBox TxtDescTrans 
               Height          =   300
               Left            =   1215
               Locked          =   -1  'True
               TabIndex        =   14
               Text            =   "TxtDescTrans"
               Top             =   270
               Width           =   4530
            End
            Begin VB.Label LblIdNumPlaca 
               AutoSize        =   -1  'True
               Caption         =   "LblIdNumPlaca"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   4470
               TabIndex        =   51
               Top             =   90
               Visible         =   0   'False
               Width           =   1080
            End
            Begin VB.Label LblIdCho 
               AutoSize        =   -1  'True
               Caption         =   "LblIdCho"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   3555
               TabIndex        =   50
               Top             =   90
               Visible         =   0   'False
               Width           =   630
            End
            Begin VB.Label LblIdTran 
               AutoSize        =   -1  'True
               Caption         =   "LblIdTran"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   2610
               TabIndex        =   49
               Top             =   90
               Visible         =   0   'False
               Width           =   675
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Placa Auto"
               Height          =   195
               Index           =   3
               Left            =   3075
               TabIndex        =   47
               Top             =   930
               Width           =   780
            End
            Begin VB.Label Label12 
               AutoSize        =   -1  'True
               Caption         =   "Nº Brevete"
               Height          =   195
               Index           =   2
               Left            =   120
               TabIndex        =   46
               Top             =   930
               Width           =   780
            End
            Begin VB.Label Label8 
               AutoSize        =   -1  'True
               Caption         =   "Chofer"
               Height          =   195
               Left            =   120
               TabIndex        =   44
               Top             =   630
               Width           =   465
            End
            Begin VB.Label Label7 
               AutoSize        =   -1  'True
               Caption         =   "Transportista"
               Height          =   195
               Left            =   120
               TabIndex        =   42
               Top             =   315
               Width           =   915
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "[ Datos de la Orden de compra ]"
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
            Height          =   1560
            Left            =   45
            TabIndex        =   32
            Top             =   1905
            Width           =   5835
            Begin VB.CommandButton CmdBusDir 
               Enabled         =   0   'False
               Height          =   240
               Left            =   5505
               Picture         =   "FrmManGuias.frx":8CDA
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   1200
               Width           =   240
            End
            Begin VB.CommandButton CmdBusPunVen 
               Enabled         =   0   'False
               Height          =   240
               Left            =   5505
               Picture         =   "FrmManGuias.frx":8E0C
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   900
               Width           =   240
            End
            Begin VB.TextBox TxtDirPunVen 
               Height          =   300
               Left            =   1245
               Locked          =   -1  'True
               TabIndex        =   13
               Text            =   "TxtDirPunVen"
               Top             =   1170
               Width           =   4530
            End
            Begin VB.TextBox TxtPunVen 
               Height          =   300
               Left            =   1245
               Locked          =   -1  'True
               TabIndex        =   12
               Text            =   "TxtPunVen"
               Top             =   870
               Width           =   4530
            End
            Begin VB.TextBox TxtNumOrd 
               Height          =   300
               Left            =   1245
               Locked          =   -1  'True
               MaxLength       =   50
               TabIndex        =   9
               Text            =   "TxtNumOrd"
               Top             =   270
               Width           =   4530
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmiPed 
               Height          =   300
               Left            =   1245
               TabIndex        =   10
               Top             =   570
               Width           =   1350
               _ExtentX        =   2381
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
               Valor           =   "06/02/2006"
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchEnt 
               Height          =   300
               Left            =   4230
               TabIndex        =   11
               Top             =   570
               Width           =   1350
               _ExtentX        =   2381
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
               Valor           =   "06/02/2006"
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Entrega"
               Height          =   195
               Index           =   4
               Left            =   3105
               TabIndex        =   35
               Top             =   615
               Width           =   915
            End
            Begin VB.Label LblIdPunVen 
               AutoSize        =   -1  'True
               Caption         =   "LblIdPunVen"
               ForeColor       =   &H000000FF&
               Height          =   195
               Left            =   3495
               TabIndex        =   40
               Top             =   90
               Visible         =   0   'False
               Width           =   915
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Fch. Pedido"
               Height          =   195
               Index           =   3
               Left            =   105
               TabIndex        =   34
               Top             =   615
               Width           =   855
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Lugar  Entrega"
               Height          =   195
               Index           =   2
               Left            =   105
               TabIndex        =   38
               Top             =   1215
               Width           =   1050
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Punto Venta"
               Height          =   195
               Index           =   1
               Left            =   105
               TabIndex        =   36
               Top             =   915
               Width           =   885
            End
            Begin VB.Label Label5 
               AutoSize        =   -1  'True
               Caption         =   "Nº Orden"
               Height          =   195
               Index           =   0
               Left            =   105
               TabIndex        =   33
               Top             =   315
               Width           =   660
            End
         End
         Begin VB.CommandButton CmdCli 
            Enabled         =   0   'False
            Height          =   240
            Left            =   6960
            Picture         =   "FrmManGuias.frx":8F3E
            Style           =   1  'Graphical
            TabIndex        =   28
            Top             =   1005
            Width           =   240
         End
         Begin VB.TextBox TxtCli 
            Height          =   300
            Left            =   1350
            Locked          =   -1  'True
            TabIndex        =   4
            Text            =   "TxtCli"
            Top             =   975
            Width           =   5880
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2940
            Locked          =   -1  'True
            TabIndex        =   3
            Text            =   "TxtNumDoc"
            Top             =   660
            Width           =   3195
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmi 
            Height          =   300
            Left            =   8010
            TabIndex        =   1
            Top             =   345
            Width           =   1350
            _ExtentX        =   2381
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
            Valor           =   "06/02/2006"
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   3210
            Left            =   45
            TabIndex        =   54
            Top             =   3510
            Width           =   10215
            _cx             =   18018
            _cy             =   5662
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
            Rows            =   2
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmManGuias.frx":9070
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
            Begin VB.Frame Frame6 
               BorderStyle     =   0  'None
               Caption         =   "Frame6"
               Height          =   2130
               Left            =   180
               TabIndex        =   104
               Top             =   480
               Visible         =   0   'False
               Width           =   4440
               Begin VB.CommandButton CmdCan 
                  Caption         =   "&Cancelar"
                  Height          =   375
                  Left            =   2220
                  TabIndex        =   107
                  Top             =   1605
                  Width           =   1300
               End
               Begin VB.CommandButton CmdAcep 
                  Caption         =   "&Aceptar"
                  Height          =   375
                  Left            =   885
                  TabIndex        =   106
                  Top             =   1605
                  Width           =   1300
               End
               Begin VB.TextBox TxtNumLote 
                  Height          =   300
                  Left            =   1440
                  MaxLength       =   20
                  TabIndex        =   105
                  Text            =   "TxtNumLote"
                  Top             =   510
                  Width           =   2850
               End
               Begin AspaTextBoxFecha.TextBoxFecha TxtFchProd 
                  Height          =   300
                  Left            =   1440
                  TabIndex        =   108
                  Top             =   825
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
                  Valor           =   "08/12/2006"
               End
               Begin AspaTextBoxFecha.TextBoxFecha TxtFchVen 
                  Height          =   300
                  Left            =   1440
                  TabIndex        =   109
                  Top             =   1140
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
                  Valor           =   "08/12/2006"
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Nº de Lote"
                  Height          =   195
                  Index           =   2
                  Left            =   120
                  TabIndex        =   113
                  Top             =   540
                  Width           =   765
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Fch. Vencimiento"
                  Height          =   195
                  Index           =   1
                  Left            =   120
                  TabIndex        =   112
                  Top             =   1170
                  Width           =   1230
               End
               Begin VB.Label Label13 
                  AutoSize        =   -1  'True
                  Caption         =   "Fch. Produccion"
                  Height          =   195
                  Index           =   0
                  Left            =   120
                  TabIndex        =   111
                  Top             =   870
                  Width           =   1170
               End
               Begin VB.Label Label11 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "Datos Adicionales"
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
                  Left            =   100
                  TabIndex        =   110
                  Top             =   80
                  Width           =   1545
               End
               Begin VB.Line Line8 
                  BorderColor     =   &H80000003&
                  BorderWidth     =   2
                  Index           =   0
                  X1              =   4425
                  X2              =   4425
                  Y1              =   15
                  Y2              =   2115
               End
               Begin VB.Line Line7 
                  BorderColor     =   &H80000003&
                  BorderWidth     =   2
                  X1              =   30
                  X2              =   4395
                  Y1              =   2115
                  Y2              =   2115
               End
               Begin VB.Line Line8 
                  BorderColor     =   &H80000005&
                  BorderWidth     =   2
                  Index           =   1
                  X1              =   15
                  X2              =   15
                  Y1              =   0
                  Y2              =   2100
               End
               Begin VB.Line Line6 
                  BorderColor     =   &H80000005&
                  BorderWidth     =   2
                  X1              =   15
                  X2              =   4410
                  Y1              =   15
                  Y2              =   15
               End
               Begin VB.Shape Shape2 
                  BackColor       =   &H80000002&
                  BackStyle       =   1  'Opaque
                  BorderColor     =   &H80000002&
                  FillColor       =   &H00800000&
                  Height          =   250
                  Left            =   45
                  Top             =   45
                  Width           =   4350
               End
            End
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1350
            Locked          =   -1  'True
            TabIndex        =   2
            Text            =   "TxtNumSer"
            Top             =   660
            Width           =   1200
         End
         Begin VB.TextBox TxtIdTipDoc 
            Height          =   300
            Left            =   1350
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   6
            Text            =   "TxtIdTipDoc"
            Top             =   1290
            Width           =   915
         End
         Begin VB.TextBox TxtIdAlm 
            Height          =   300
            Left            =   8265
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   5
            Text            =   "TxtIdAlm"
            Top             =   960
            Width           =   705
         End
         Begin VB.Shape Shape7 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2670
            Top             =   780
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   195
            Index           =   11
            Left            =   7440
            TabIndex        =   130
            Top             =   1005
            Width           =   615
         End
         Begin VB.Label LblAlmacen 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblAlmacen"
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
            Left            =   9000
            TabIndex        =   129
            Top             =   960
            Width           =   2700
         End
         Begin VB.Label LblIdDocRef 
            Caption         =   "LblIdDocRef"
            ForeColor       =   &H000000FF&
            Height          =   180
            Left            =   10710
            TabIndex        =   98
            Top             =   1350
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Label LblDescTipDocRef 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblDescTipDocRef"
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
            Left            =   2325
            TabIndex        =   97
            Top             =   1290
            Width           =   4890
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tip. Doc. Ref."
            Height          =   195
            Index           =   8
            Left            =   60
            TabIndex        =   96
            ToolTipText     =   "Tipo de Documento de Referencia"
            Top             =   1335
            Width           =   1005
         End
         Begin VB.Label LblMotivo 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblMotivo"
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
            Left            =   2325
            TabIndex        =   79
            Top             =   1605
            Width           =   4890
         End
         Begin VB.Label lblcotizacion 
            AutoSize        =   -1  'True
            Caption         =   "Nº Pedido"
            Height          =   195
            Left            =   7440
            TabIndex        =   76
            Top             =   1365
            Width           =   720
         End
         Begin VB.Label LblTipoItem 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblTipoItem"
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
            Left            =   2325
            TabIndex        =   24
            Top             =   345
            Width           =   3810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Item"
            Height          =   195
            Index           =   6
            Left            =   60
            TabIndex        =   22
            Top             =   390
            Width           =   660
         End
         Begin VB.Label Lblidmottra 
            AutoSize        =   -1  'True
            Caption         =   "Lblidmottra"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   7290
            TabIndex        =   31
            Top             =   1665
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.Label Label9 
            AutoSize        =   -1  'True
            Caption         =   "Motivo Traslado"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   53
            Top             =   1635
            Width           =   1140
         End
         Begin VB.Label LblIdCli 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCli"
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   6720
            TabIndex        =   29
            Top             =   720
            Visible         =   0   'False
            Width           =   510
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "Nº Documento"
            Height          =   195
            Left            =   60
            TabIndex        =   26
            Top             =   720
            Width           =   1050
         End
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            Caption         =   "Detalle Guia de Remisión"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   90
            TabIndex        =   21
            Top             =   45
            Width           =   11610
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   27
            Top             =   1020
            Width           =   480
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Emision"
            Height          =   195
            Left            =   6915
            TabIndex        =   25
            Top             =   390
            Width           =   900
         End
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   59
      Top             =   0
      Width           =   11865
      _ExtentX        =   20929
      _ExtentY        =   1058
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
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Guias"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Restaurar Guia"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Items de Guias"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   3
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Anular Guia"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar Guia"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Emitir Guia Anulada"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Grabar"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Object.ToolTipText     =   "Cancelar"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Facturadas"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "No Facturadas"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Anuladas"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Todas"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Guia"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            ImageIndex      =   10
            Style           =   3
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu_1 
      Caption         =   "menu"
      Visible         =   0   'False
      Begin VB.Menu menu_1_1 
         Caption         =   "Agregar item             "
      End
      Begin VB.Menu menu_1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu_1_3 
         Caption         =   "Eliminar item         "
      End
   End
End
Attribute VB_Name = "FrmManGuias"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre            : FrmManGuias
'* Tipo              : FORMULARIO
'* Descripcion       : FORMULARIO PARA EL INGRESO Y MANTENIMIENTO DE LAS GUIAS DE REMISION
'* DISEÑADO POR      : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION   : 24/09/09
'* VERSION           : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstGui As New ADODB.Recordset       ' RECORDSET PRINCIPAL QUE ALMACENARA LOS REGISTRO DE LA TABLA vta_guias, ESTOS DATOS SE MOSTRARAN EN LA PESTAÑA CONSULTA
Dim QueHace As Integer                  ' ESPECIFICA EN QUE MODO SE ENCUENTRA EL FORMULARIO  1 = ADICIONA, 2 = MODIFICA, 3 = SOLO LECTURA
Dim SeEjecuto As Boolean                ' VARIABLE QUE INDICA QUE EL EVENTO ACTIVATE DEBE DE EJECUTARSE UNA SOLA VEZ
Dim xFchIni, xFchFin As String          ' VARIABLE QUE ALMACENA LA HORA DE INICIO Y LA HORA FINAL
Dim CaracteresNumericos As String       ' VARIABLE QUE ALAMCENA LOS CARACTERES NUMERICOS QUE SERAN UTLIZADOS EN LOS CONTROLES TextBox
Dim xnumreg As Integer                  ' VARIBALE QUE ALMACENA EN NUMERO DE REGISTRO
Dim swguiafact                          ' 0 No se Procesaron , 1 Se Procesaron
Dim xHorIni As Date                     ' VARAIABLE QUE ALMACENA LA HORA DE INICIO DE EDICION DEL REGISTRO
Dim fOrdenLista As Boolean              ' especfica el orden de la lista de la consulta
Dim Agregando As Boolean               ' para saber cuando se este agregando FILAS AL CONTROL grid de productos

Dim JALOPEDIDO As Boolean               ' ESPECIFICS SI LAS GUIAS SE GENERAN DE UNA O VAIAS ORDENES DE PEDIDO
Dim VAR_IDPEDIDO As Integer             ' VARIABLE QUE ESPECIFICA EL ID DEL PEDIDO
Dim VAR_FECHAPEDIDO As String           ' VARIABLEA QUE INDICA LA FECHA DEL PEDIDO

Dim mIdRegistro&                        ' identificador del registro
Dim mMesActivo As Integer               ' ESPECFICA EL MES ACTUAL

Dim fCierrePeriodo As Boolean           ' --indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)
Dim IdMenuActivo As Integer             'INDICA EL CODIGO DEL MENU ACTIVO
Dim cSQL As String
Dim F As New SistemaLogica.Funciones

'Estado de la cotizacones
'1 pendiente
'2 aprobada
'3 procesada
'4 rechazada
'QueHace 1 ADICIONAR
'QueHace 2 MODIFICAR

'*****************************************************************************************************
'* Nombre           : ActualizarStock
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTUALIZ EL STOCK ACTUAL DE LOS ITEMS
'* Paranetros       : NOMBRE   |  TIPO    |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    TIPO     |  STRING  |  ESPECIFICA EL TIPO DE MOVIMIENTO S= SALIDA; E = ENTRADA
'*                    numid    |  INTEGER |  ESPECIFICA EL NUMERO DE REGISTRO
'* Devuelve         :
'*****************************************************************************************************
Sub ActualizarStock(TIPO As String, numid As Double)
    'Tipo S= Salida
    'Tipo E= Extorno por Anular Guia , Eliminar Guia

    Dim RstDet As New ADODB.Recordset
    Dim Rstitem As New ADODB.Recordset

    RST_Busq RstDet, "SELECT vta_guiadet.* FROM vta_guiadet WHERE idgui = " & numid & "", xCon

    Do While Not RstDet.EOF
        RST_Busq Rstitem, "SELECT alm_inventario.* FROM alm_inventario WHERE id = " & NulosN(RstDet("iditem")) & "", xCon
        If Rstitem.RecordCount > 0 Then
            If TIPO = "S" Then
                Rstitem("stckact") = Rstitem("stckact") - RstDet("canpro")
            ElseIf TIPO = "E" Then
                Rstitem("stckact") = Rstitem("stckact") + RstDet("canpro")
            End If
            Rstitem.Update
        End If
        RstDet.MoveNext
    Loop
    
    Set RstDet = Nothing
    Set Rstitem = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO, ESTA INFORMACION SE MUESTRA EN LA
'*                    PESTAÑA DETALLE DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Blanquea
    If RstGui.RecordCount = 0 Or RstGui.BOF = True Or RstGui.EOF = True Then Exit Sub
    
    ' MUESTRA LOS DATOS PRINCIPALES DEL REGISTRO
    TxtTipItem.Text = RstGui("tippro")
    TxtFchEmi.Valor = RstGui("fecgiro")
    TxtNumSer.Text = Format(RstGui("numser"), "0000")
    TxtNumDoc.Text = Format(RstGui("numdoc"), "0000000000")
    TxtCli.Text = NulosC(RstGui("nombre"))
    LblIdCli.Caption = RstGui("idcli")
    TxtIdTipDoc.Text = NulosN(RstGui("idtipdocref"))
    TxtIdTipDoc.Tag = NulosN(RstGui("idtipdocref"))
    LblDescTipDocRef.Caption = Busca_Codigo(NulosN(TxtIdTipDoc.Text), "id", "descripcion", "mae_docreferencia", "N", xCon)
    
    TxtIdAlm.Text = NulosN(RstGui("idalm"))
    LblAlmacen.Caption = NulosC(RstGui("desalm"))
    
    If NulosC(LblDescTipDocRef.Caption) = "" Then
        TxtIdTipDoc.Text = ""
    End If
        
    LblMotivo.Caption = NulosC(RstGui("descmotgui"))
    TxtMotivo.Text = NulosN(RstGui("idmottra"))
    LblTipoItem.Caption = Busca_Codigo(RstGui("tippro"), "id", "descripcion", "mae_tipoproducto", "N", xCon)
    
    If NulosN(RstGui("idmottra")) = 1 Then
        Fg1.ColHidden(Fg1.ColIndex("PREUNI")) = False
    Else
        Fg1.ColHidden(Fg1.ColIndex("PREUNI")) = True
    End If
    
    ' Datos de la orden de compra
    TxtNumOrd.Text = NulosC(RstGui("numordcom"))
    LblIdDocRef.Caption = Busca_Codigo(NulosN(TxtNumOrd.Text), "oc", "id", "ped_pedido", "C", xCon)
    
    TxtFchEmiPed.Valor = Format(RstGui("fchemiord"), "dd/mm/yy")
    TxtFchEnt.Valor = Format(RstGui("fchentord"), "dd/mm/yy")
    
    TxtPunVen.Text = NulosC(RstGui("despunven"))
    LblIdPunVen.Caption = NulosN(RstGui("idpunven"))
    TxtDirPunVen.Text = NulosC(RstGui("direccion"))
    
    ' Datos del transporte
    TxtDescTrans.Text = NulosC(RstGui("desemptra"))
    LblIdTran.Caption = NulosC(RstGui("idemptra"))
    LblIdCho.Caption = NulosN(RstGui("idcho"))
    LblIdNumPlaca.Caption = NulosN(RstGui("idveh"))
    TxtNumBre.Text = NulosC(RstGui("numbre"))
    TxtNumPlaCar.Text = NulosC(RstGui("numpla"))
    
    ' MOSTRAMOS LOS DATOS ADICIONALES
    TxtFchProd.Valor = NulosC(RstGui("fchpro"))
    TxtFchVen.Valor = NulosC(RstGui("fchven"))
    TxtNumLote.Text = NulosC(RstGui("numlote"))
    
    ' cargamos las variables para mostrar el pedido al que esta relacionada la guia, sis es que tuviera un pedido asociado
    If NulosC(TxtNumCotizacion.Text) <> "-" And NulosC(TxtNumCotizacion.Text) <> "" Then
        Dim Rst2 As New ADODB.Recordset
'        RST_Busq Rst2, "SELECT DISTINCT ped_pedidodetent.idped, ped_pedidodetent.fchent, ped_pedidodetent.idtipdoc, ped_pedidodetent.iddocven" _
'            & " From ped_pedidodetent WHERE (((ped_pedidodetent.idtipdoc)=1) AND ((ped_pedidodetent.iddocven)=" & NulosN(RstGui("id")) & "))", xCon
        Set Rst2 = Nothing
        RST_Busq Rst2, "SELECT DISTINCT ped_pedidodet.idped, ped_pedidodet.fchent, ped_pedidodet.idtipdoc, ped_pedidodet.iddocven" _
            & " From ped_pedidodet WHERE (((ped_pedidodet.idtipdoc)=1) AND ((ped_pedidodet.iddocven)=" & NulosN(RstGui("id")) & "))", xCon
        
        If Rst2.RecordCount <> 0 Then
            VAR_IDPEDIDO = Rst2("idped")
            VAR_FECHAPEDIDO = Rst2("fchent")
        End If
        Set Rst2 = Nothing
    End If
    
    Habilitar_Control_Pedido
    
    
    Dim Rst As New ADODB.Recordset
    Dim A As Integer
    ' MOSTRAMOS LOS DATOS DEL VENDEDOR
    RST_Busq Rst, "SELECT mae_chofer.*, UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom] AS apenom " _
        & " FROM pla_empleados RIGHT JOIN mae_chofer ON pla_empleados.id = mae_chofer.idper Where (((mae_chofer.id) = " & RstGui("idcho") & ")) ORDER BY " _
        & " UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom]", xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        TxtChofer.Text = NulosC(Rst("apenom"))
    Else
        TxtChofer.Text = ""
    End If
    Set Rst = Nothing
    
    Agregando = True
    
    ' MOSTRAMOS EL DETALLE DE LA GUIA
    Dim cSQL As String
    cSQL = "SELECT vta_guiadet.idgui, vta_guiadet.iditem, vta_guiadet.lote, vta_guiadet.preuni, alm_inventario.descripcion, alm_inventario.codpro, vta_guiadet.idunimed, MAE_Unidades.descripcion AS UnidMed, vta_guiadet.canpro, vta_guiadet.iddocref " _
        + vbCr + "FROM (vta_guiadet LEFT JOIN alm_inventario ON vta_guiadet.iditem = alm_inventario.id) LEFT JOIN MAE_Unidades ON vta_guiadet.idunimed = MAE_Unidades.id " _
        + vbCr + "Where (((vta_guiadet.idgui) = " & RstGui("id") & "))" _
        + vbCr + "ORDER BY alm_inventario.descripcion;"
        
    RST_Busq Rst, cSQL, xCon

    Fg1.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 0) = NulosN(Rst("iddocref"))
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(Rst("codpro"))
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(Rst("descripcion"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(Rst("Unidmed"))
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(NulosN(Rst("canpro")), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosN(Rst("idunimed"))
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosN(Rst("iditem"))
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(NulosN(Rst("canpro")), "0.00")
            
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosC(Rst("lote"))
            Fg1.TextMatrix(Fg1.Rows - 1, Fg1.ColIndex("PREUNI")) = Format(NulosN(Rst("preuni")), "0.00")
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
    
    Agregando = True
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BLANQUEA LOS CONTROLES TextBox, PREPARA EL FORMULARIO PARA EL INGRESO DE UN
'*                    NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtCli.Text = ""
    TxtMotivo.Text = ""
    TxtFchEmi.Valor = ""
    TxtNumCotizacion.Text = ""
    TxtIdTipDoc.Text = ""
    
    'datos del pedido
    TxtNumOrd.Text = ""
    TxtFchEmiPed.Valor = ""
    TxtFchEnt.Valor = ""
    TxtPunVen.Text = ""
    TxtDirPunVen.Text = ""
    
    'datos del transporte
    TxtDescTrans.Text = ""
    TxtChofer.Text = ""
    TxtNumBre.Text = ""
    TxtNumPlaCar.Text = ""
    TxtIdAlm.Text = ""
    
    LblIdCli.Caption = ""
    Lblidmottra.Caption = ""
    LblIdPunVen.Caption = ""
    LblIdTran.Caption = ""
    LblIdCho.Caption = ""
    LblIdNumPlaca.Caption = ""
    LblTipoItem.Caption = ""
    LblMotivo.Caption = ""
    LblDescTipDocRef.Caption = ""
    LblIdDocRef.Caption = ""
    LblAlmacen.Caption = ""
    
    TxtTipItem.Text = ""
    TxtFchProd.Valor = ""
    TxtFchVen.Valor = ""
    TxtNumLote.Text = ""
    TxtNumCotizacion.Text = ""
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TextBox, PREPARA LOS CONTROLES PARA LA ADICION
'*                    O MODIFICACION DE UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea()
    TxtFchEmi.Locked = Not TxtFchEmi.Locked
    
    'If QueHace = 1 Then
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    'End If
    
    TxtCli.Locked = Not TxtCli.Locked
    TxtMotivo.Locked = Not TxtMotivo.Locked
    TxtTipItem.Locked = Not TxtTipItem.Locked
    TxtIdTipDoc.Locked = Not TxtIdTipDoc.Locked
    
    'datos del pedido
    TxtNumOrd.Locked = Not TxtNumOrd.Locked
    TxtFchEmiPed.Locked = Not TxtFchEmiPed.Locked
    TxtFchEnt.Locked = Not TxtFchEnt.Locked
    TxtPunVen.Locked = Not TxtPunVen.Locked
    TxtDirPunVen.Locked = Not TxtDirPunVen.Locked
    
    'datos del transporte
    TxtDescTrans.Locked = Not TxtDescTrans.Locked
    TxtChofer.Locked = Not TxtChofer.Locked
    TxtNumBre.Locked = Not TxtNumBre.Locked
    TxtNumPlaCar.Locked = Not TxtNumPlaCar.Locked
    
    'CmdBusNumSer.Enabled = Not CmdBusNumSer.Enabled
    CmdCli.Enabled = Not CmdCli.Enabled
    CmdMot.Enabled = Not CmdMot.Enabled
    CmdBusTipItem.Enabled = Not CmdBusTipItem.Enabled
    CmdBusPunVen.Enabled = Not CmdBusPunVen.Enabled
    CmdEmpTra.Enabled = Not CmdEmpTra.Enabled
    CmdCho.Enabled = Not CmdCho.Enabled
    CmdPlaCar.Enabled = Not CmdPlaCar.Enabled
    CmdBusDir.Enabled = Not CmdBusDir.Enabled
    
    CmdAddItem.Enabled = Not CmdAddItem.Enabled
    CmdDelItem.Enabled = Not CmdDelItem.Enabled
    CmdDatosAdicion.Enabled = Not CmdDatosAdicion.Enabled
    cmdbuscotizacion.Enabled = Not cmdbuscotizacion.Enabled
    cmdcotizprocesadas.Enabled = Not cmdcotizprocesadas.Enabled
    CmdAddChofer.Enabled = Not CmdAddChofer.Enabled
    
    cmd(1).Enabled = Not cmd(1).Enabled
End Sub

Private Sub cmd_Click(Index As Integer)
    If QueHace = 3 Then Exit Sub
    
    Select Case Index
        Case 0 ' Seleccionar
        
        Case 1 ' Eliminar Todos
            If MsgBox("¿Esta seguro de eliminar todos los registros?", _
                        vbQuestion + vbYesNo + vbDefaultButton2, xTitulo) = vbNo Then Exit Sub
            
            Fg1.Rows = Fg1.FixedRows
            
    End Select
End Sub

Private Sub CmdAcep_Click()
    ' VALIDA QUE LOS DATOS ADICIONALES SE HAYAN INGRESADO CORRECTAMENTE
    
    ' VERIFICAMOS QUE LOS DATOS SE HAYAN INGRESADO CORRECTAMENTE
    If NulosC(TxtFchProd.Valor) = "" Then
        MsgBox "No ha especificado la fecha de produccion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchProd.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtFchVen.Valor) = "" Then
        MsgBox "No ha especificado la fecha de vencimiento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchVen.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtNumLote.Text) = "" Then
        MsgBox "No ha especificado el numero de lote del producto", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumLote.SetFocus
        Exit Sub
    End If
    TabOne1.Enabled = True
    Toolbar1.Enabled = True
    Frame6.Visible = False
End Sub

Private Sub CmdAcepta22_Click()
    ' ACTUALIZA LAS CANTIDADES DE LOS ITEMS EN LAS GUIAS ESPECIFICADAS
    Dim Rpta As Integer
    Dim A As Integer
    
    Rpta = MsgBox("Esta seguro de actualizar las cantidades a los items especificados", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        For A = 1 To Fg6.Rows - 1
            xCon.Execute "UPDATE vta_guiadet SET vta_guiadet.canpro = " & NulosN(Fg6.TextMatrix(A, 5)) & " WHERE " _
                & " ((vta_guiadet.idgui=  " & NulosN(Fg6.TextMatrix(A, 6)) & " ) AND (vta_guiadet.iditem= " & NulosN(Fg6.TextMatrix(A, 7)) & " ))"
            
        Next A
        MsgBox "El proceso de actualizacion de items termino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        CmdCancela22_Click
    End If
End Sub

Private Sub CmdAddChofer_Click()
    'LLAMA AL FORMULARIO FrmIngChofer PARA EL INGRESO DE UN NUEVO CHOFER
    FrmIngChofer.Show vbModal
End Sub

Private Sub CmdAddEntrega_Click()
    Dim A1 As Integer
    Dim Rst2 As New ADODB.Recordset
    Dim A As Integer
    Dim nTitulo As String
    Dim xform As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    Dim xCampos(5, 4) As String
    Dim IDPED_ As Integer
    Dim IDCLI_ As Integer
    
    ' EJECUTA LA BUSQUEDA DE UN PEDIDO
    If NulosN(TxtNumCotizacion.Text) = 0 Then
        MsgBox "No ha especificado la Orden de Pedido", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumCotizacion.SetFocus
        Exit Sub
    End If
    
    'descripcion                                'campo                          'tamaño                        'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":              xCampos(0, 1) = "desitem":      xCampos(0, 2) = "4500":        xCampos(0, 3) = "C":
    xCampos(1, 0) = "Uni. Med":                 xCampos(1, 1) = "desunimed":    xCampos(1, 2) = "1000":        xCampos(1, 3) = "C":
    xCampos(2, 0) = "Dia/":                     xCampos(2, 1) = "dia":          xCampos(2, 2) = "400":         xCampos(2, 3) = "C":
    xCampos(3, 0) = "Mes/":                     xCampos(3, 1) = "mes":          xCampos(3, 2) = "450":         xCampos(3, 3) = "C":
    xCampos(4, 0) = "Año":                      xCampos(4, 1) = "anio":         xCampos(4, 2) = "500":         xCampos(4, 3) = "C":
    
    'CARGAMOS LOS PEDIDOS PENDIENTES
    IDPED_ = NulosN(LblIdDocRef.Caption)
    IDCLI_ = NulosN(LblIdCli.Caption)

    cSQL = "SELECT DISTINCT 0 AS xsel, ped_pedidodet.idped, ped_pedidodet.idpeddet, Day([ped_pedidodet].[fchent]) AS dia, Month([ped_pedidodet].[fchent]) AS mes, Year([ped_pedidodet].[fchent]) AS anio, ped_pedidodet.iditem, alm_inventario.codpro, alm_inventario.descripcion AS desitem, ped_pedidodet.fchent, ped_pedidodet.idunimed, mae_unidades.abrev AS desunimed " _
        + vbCr + "FROM ((ped_pedido LEFT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id " _
        + vbCr + "WHERE (((ped_pedido.id)=" & IDPED_ & ") AND ((ped_pedido.idcli)=" & IDCLI_ & "));"


'    cSQL = "SELECT DISTINCT 0 AS xsel, ped_pedidodet.idped, ped_pedidodet.idpeddet, ped_pedido.idcli, Day([ped_pedidodet].[fchent]) AS dia, Month([ped_pedidodet].[fchent]) AS mes, Year([ped_pedidodet].[fchent]) AS anio, alm_inventario.descripcion, mae_unidades.abrev AS abreunimed, [ped_pedidodet].[canpro]-[ped_pedidodet].[canproent] AS canprorest " _
'        + vbCr + "FROM (((mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN ped_pedido ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) LEFT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id " _
'        + vbCr + "WHERE (((ped_pedido.id)=" & NulosN(LblIdDocRef.Caption) & ") AND ((ped_pedido.idcli)=" & NulosN(LblIdCli.Caption) & ") AND (([ped_pedidodet].[canpro]-[ped_pedidodet].[canproent])>0) AND ((ped_pedidodet.fchent) Is Not Null) AND ((alm_inventario.descripcion) Is Not Null))"
        
    nTitulo = "Buscando Pedidos por Despachar"
    xform.SQLCad = cSQL
        
    xform.titulo = nTitulo
    Set xform.Coneccion = xCon
    Set xRs = Nothing
    Set xRs = xform.Seleccionar(xCampos)
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    Fg1.Rows = 1
    
    While Not xRs.EOF
        Fg1.Rows = Fg1.Rows + 1
        Fg1.TextMatrix(Fg1.Rows - 1, 0) = NulosN(xRs("idpeddet"))
        Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("codpro"))
        Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRs("desitem"))
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = NulosC(xRs("desunimed"))
        Fg1.TextMatrix(Fg1.Rows - 1, 4) = 0
        Fg1.TextMatrix(Fg1.Rows - 1, 5) = NulosN(xRs("idunimed"))
        Fg1.TextMatrix(Fg1.Rows - 1, 6) = NulosN(xRs("iditem"))
        xRs.MoveNext
    
    
'        VAR_IDPEDIDO = NulosN(xRs("idped"))
'        VAR_FECHAPEDIDO = CDate(xRs("dia") & "/" & xRs("mes") & "/" & xRs("anio"))
'
'        ' CARGAMOS LOS ITEMS DEL PEDIDO
'        cSQL = "SELECT ped_pedido.id, ped_pedido.idcli, ped_pedido.fchemi, mae_documento.abrev, ped_pedido!numser & '-' & ped_pedido!numdoc AS numdoc, ped_pedidodet.fchent, mae_cliente.nombre, alm_inventario.descripcion, mae_unidades.abrev AS unimed, alm_inventario.codpro, ped_pedidodet.canpro, ped_pedidodet.idunimed, ped_pedidodet.iditem " _
'                + vbCr + "FROM (((mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN ped_pedido ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) INNER JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id " _
'                + vbCr + "WHERE (((ped_pedido.id) = " & xRs("idped") & ") And ((ped_pedido.idcli) = " & NulosN(LblIdCli.Caption) & ") And ((ped_pedidodet.fchent) = CDate('" & VAR_FECHAPEDIDO & "')))" _
'                + vbCr + "ORDER BY ped_pedidodet.fchent, mae_cliente.nombre"
'
'         RST_Busq Rst2, cSQL, xCon
'
'        If Rst2.RecordCount <> 0 Then
'            Rst2.MoveFirst
'            While Not Rst2.EOF
'                Fg1.Rows = Fg1.Rows + 1
'                Fg1.TextMatrix(A1, 0) = xRs("idpeddet")
'                Fg1.TextMatrix(A1, 1) = Rst2("codpro")
'                Fg1.TextMatrix(A1, 2) = Rst2("descripcion")
'                Fg1.TextMatrix(A1, 3) = Rst2("unimed")
'                Fg1.TextMatrix(A1, 4) = xRs("canprorest")
'                Fg1.TextMatrix(A1, 5) = Rst2("idunimed")
'                Fg1.TextMatrix(A1, 6) = Rst2("IdItem")
'
'                Rst2.MoveNext
'            Wend
'        End If
'        xRs.MoveNext
    Wend
    
    Set xform = Nothing
    Set xRs = Nothing
    Set Rst2 = Nothing
    TxtMotivo.SetFocus
End Sub

Private Sub CmdAddItem_Click()
    ' AGREGA UNA FILA AL CONTROL FlexGrid Fg1
    AddItem
End Sub

Private Sub CmdBusAlm_Click()
    ' EJECUTA LA BUSQUEDA DE ALMACENES
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT alm_almacenes.* FROM alm_almacenes"
    
    xform.titulo = "Buscando Almacenes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        LblAlmacen.Caption = NulosC(xRs("descripcion"))
        TxtIdAlm.Text = NulosN(xRs("id"))
        TxtIdTipDoc.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub TxtIdAlm_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdAlm_Validate(Cancel As Boolean)
    ' VALIDA EL ID DEL ALMACEN
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtIdAlm.Text) <> "" Then
        LblAlmacen.Caption = Busca_Codigo(NulosN(TxtIdAlm.Text), "id", "descripcion", "alm_almacenes", "N", xCon)
        If LblAlmacen.Caption = "" Then
            TxtIdAlm.Text = ""
        End If
    Else
        LblAlmacen.Caption = ""
    End If
End Sub

Private Sub CmdBusCli_Click()
    ' ejecuta la busqueda de un cliente
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Razon Social":  xCampos(0, 1) = "nombre":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "R.U.C.":        xCampos(1, 1) = "numruc":   xCampos(1, 2) = "1400":    xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT MAE_Cliente.id, MAE_Cliente.nombre, MAE_Cliente.numruc, " _
        & " MAE_Cliente.dir From MAE_Cliente ORDER BY MAE_Cliente.Nombre"
    
    xform.titulo = "Buscando Clientes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtCliente.Text = xRs("Nombre")
            LblIdcli2.Caption = xRs("id")
            fg5.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

'Private Function encontrarRutas(ByRef rut() As String) As Boolean
'    Dim xConAux As New ADODB.Connection
'    Dim xFun As New eps_librerias.FuncionesData
'    Dim Rst As New ADODB.Recordset
'    Dim NumRUC As String
'    Dim xCad As String
'    Dim cSQL As String
'    Dim rutas() As String
'    Dim cant As Integer
'    Dim A As Integer
'
'    cSQL = "SELECT numruc FROM mae_empresa"
'    RST_Busq Rst, cSQL, xCon
'
'    NumRUC = Rst("numruc")
'    xCad = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
'
'    xFun.F_BASEDATOS = xCad + "data.mdb"
'    xFun.F_GRUPOTRABAJO = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS") + "seven.mdw"
'    xFun.F_PASSWORD = Eps_Pass
'    xFun.F_USUARIO = Eps_User
'    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
'
'    Set xConAux = xFun.AbrirConeccion
'
'    cSQL = "SELECT mae_empresa.numruc, mae_empresa.ruta, mae_empresa.anotra, mae_empresa.activo " _
'        + vbCr + "From mae_empresa " _
'        + vbCr + "WHERE (((mae_empresa.numruc)= '" & NumRUC & "') AND ((mae_empresa.activo)=-1))"
'
'    RST_Busq Rst, cSQL, xConAux
'    If Rst.RecordCount <> 0 Then
'        cant = Rst.RecordCount
'        ReDim rutas(1 To cant, 1 To 2) As String
'        Rst.MoveFirst
'        For A = 1 To cant
'            rutas(A, 2) = Rst("ruta")
'            rutas(A, 1) = Rst("anotra")
'            Rst.MoveNext
'        Next A
'        rut = rutas
'        encontrarRutas = True
'    Else
'        encontrarRutas = False
'    End If
'End Function

Function generarConsulta(RutaData As String) As ADODB.Recordset
    Dim Rst As New ADODB.Recordset
    Dim xCad As String
    
    Dim xFun As New eps_librerias.FuncionesData
    Dim xRutaData As String
    Dim xRst As New ADODB.Recordset
    Dim xCon2 As New ADODB.Connection
    Dim cSQL As String
    
    xCad = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTABD", "RUTAS")
    
    xFun.F_BASEDATOS = xCad + RutaData
    xFun.F_GRUPOTRABAJO = LeerLineaINI(Trim(App.Path) + "\seven.ini", "RUTASY", "RUTAS") + "seven.mdw"
    xFun.F_PASSWORD = Eps_Pass
    xFun.F_USUARIO = Eps_User
    xFun.F_PROVEEDOR = "Microsoft.Jet.OLEDB.4.0"
    
    Set xCon2 = xFun.AbrirConeccion

'    cSQL = "SELECT DISTINCT ped_pedido.id, ped_pedido.idcli, ped_pedido.fchemi, ped_pedidodetent.fchent, mae_documento.abrev AS abredoc, ped_pedido!numser & '-' & ped_pedido!numdoc AS numdoc, alm_inventario.descripcion, mae_unidades.abrev AS abreunimed, ped_pedidodet.canpro, ped_pedido.oc " _
'        & " FROM ped_pedidodetent RIGHT JOIN ((((mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN ped_pedido ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) RIGHT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id) ON ped_pedidodetent.idped = ped_pedido.id " _
'        & " Where (((ped_pedido.idcli) = " & NulosN(LblIdCli.Caption) & ") And ((ped_pedidodet.estado) <> 1)) " _
'        & " ORDER BY ped_pedido.fchemi DESC"
        
    cSQL = "SELECT DISTINCT ped_pedido.id, ped_pedido.idcli, ped_pedido.fchemi, ped_pedidodet.fchent, mae_documento.abrev AS abredoc, ped_pedido!numser & '-' & ped_pedido!numdoc AS numdoc, alm_inventario.descripcion, mae_unidades.abrev AS abreunimed, ped_pedidodet.canpro, ped_pedido.oc " _
        & " FROM ped_pedidodet RIGHT JOIN ((((mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN ped_pedido ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) RIGHT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id) ON ped_pedidodet.idped = ped_pedido.id " _
        & " Where (((ped_pedido.idcli) = " & NulosN(LblIdCli.Caption) & ") And ((ped_pedidodet.estado) <> 1)) " _
        & " ORDER BY ped_pedido.fchemi DESC"
        
    RST_Busq Rst, cSQL, xCon2
    
'    Dim rst3 As New ADODB.Recordset'
'    DEFINIR_RST_TMP rst3, RstAño'
'    CARGAR_RST_TMP rst3, RstAño'
'    CARGAR_RST_TMP rst3, RstAño1

    Set generarConsulta = Rst
End Function

Private Sub cmdbuscotizacion_Click()
    Dim Rst1 As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    Dim RstTEMP As New ADODB.Recordset
    Dim cSQL As String
    Dim xCampos(3, 4) As String
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    ' EJECUTA LA BUSQUEDA DE UN PEDIDO
    If NulosC(TxtIdTipDoc.Text) = "" Then Exit Sub
    If NulosN(TxtIdTipDoc.Text) <> 5 Then Exit Sub
    
    If NulosC(TxtCli.Text) = "" Then
        MsgBox "No ha especificado el cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCli.SetFocus
        Exit Sub
    End If
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Num. Orden":       xCampos(0, 1) = "oc":           xCampos(0, 2) = "1300":         xCampos(0, 3) = "C":
    xCampos(1, 0) = "Fch. Emi.":        xCampos(1, 1) = "fchemi":       xCampos(1, 2) = "1000":         xCampos(1, 3) = "D":
    xCampos(2, 0) = "Num. Doc.":        xCampos(2, 1) = "numdoc":       xCampos(2, 2) = "2000":         xCampos(2, 3) = "C":
    
    ' CARGAMOS LOS PEDIDOS
    cSQL = "SELECT ped_pedido.id, ped_pedido.oc, ped_pedido.fchemi, ped_pedido!numser & '-' & ped_pedido!numdoc AS numdoc " _
        + vbCr + "FROM ped_pedido " _
        + vbCr + "WHERE (((ped_pedido.idcli)=" & NulosN(LblIdCli.Caption) & "));"
    
'    cSQL = "SELECT ped_pedido.id, ped_pedido.idcli, ped_pedido.oc, mae_cliente.nombre, ped_pedido.fchemi, mae_documento.abrev AS abredoc, ped_pedido!numser & '-' & ped_pedido!numdoc AS numdoc, Sum([ped_pedidodet].[canpro]-[ped_pedidodet].[canproent]) AS canprorest " _
'        + vbCr + "FROM (((mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN ped_pedido ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) RIGHT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id " _
'        + vbCr + "GROUP BY ped_pedido.id, ped_pedido.idcli, ped_pedido.oc, mae_cliente.nombre, ped_pedido.fchemi, mae_documento.abrev, ped_pedido!numser & '-' & ped_pedido!numdoc, ped_pedidodet.estado " _
'        + vbCr + "Having (((ped_pedido.idcli) = " & NulosN(LblIdCli.Caption) & ") And ((Sum([ped_pedidodet].[canpro] - [ped_pedidodet].[canproent])) > 0) And ((ped_pedidodet.estado) = 2)) " _
'        + vbCr + "ORDER BY ped_pedido.fchemi DESC;"
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos, "Buscando Pedidos Pendientes", "fchemi", "oc"
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub

    JALOPEDIDO = True
    TxtNumCotizacion.Text = NulosC(xRs("oc"))
    LblIdDocRef.Caption = NulosN(xRs("id"))
    TxtNumOrd.Text = NulosC(xRs("oc"))
    TxtFchEmiPed.Valor = xRs("fchemi")
    TxtFchEmiPed.Enabled = False
    TxtFchEnt.Valor = Date 'xRs("fchent")
    TxtFchEnt.Enabled = False
    TxtNumOrd.Enabled = False
    CmdAddEntrega.Enabled = True
    TxtNumCotizacion_Validate True
    TxtMotivo.SetFocus
    ' Se vacia el Grid
    Fg1.Rows = Fg1.FixedRows


    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDir_Click()
    ' EJECUTA LA BUSQUEDA DE LA DIRECCION DE UN PUNTO DE VENTA
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Direcc":  xCampos(0, 1) = "dir":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":  xCampos(1, 1) = "id":    xCampos(1, 2) = "800":     xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT vta_puntoVenta.Dir, vta_puntoVenta.id" _
        & " From vta_puntoVenta WHERE vta_puntoVenta.idcli =" & Val(LblIdCli.Caption) & ""
    
    xform.titulo = "Buscando Direccion de Destino"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "dir"
    xform.CampoBusca = "dir"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtDirPunVen.Text = xRs("Dir")
            TxtDescTrans.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusIdTipDocRef_Click()
    ' EJECUTA LA BUSQUEDA DEL DOCUMENTO DE REFERENCIA
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripción":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Código":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_docreferencia ORDER BY descripcion"
    
    xform.titulo = "Buscando Tipo de Documento de Referencia"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    '--verificar si cambian de cliente => limpiar campos
    If xRs("id") <> NulosN(TxtIdTipDoc.Text) And NulosN(TxtIdTipDoc.Text) <> 0 Then
        LblDescTipDocRef.Caption = ""
        Fg1.Rows = 1
    End If

    TxtIdTipDoc.Text = xRs("id")
    LblDescTipDocRef.Caption = NulosC(xRs("descripcion"))
    If xRs("id") = 5 Then
        JALOPEDIDO = True
        TxtIdTipDoc_Validate True
        TxtNumCotizacion.Text = ""
        LblIdDocRef.Caption = ""
        TxtNumCotizacion.SetFocus
    Else
        JALOPEDIDO = False
        TxtIdTipDoc_Validate True
        TxtMotivo.SetFocus
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusNumSer_Click()
    ' EJECUTA LA BUSQUEDA DE UN NUMERO DE SERIE
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Codigo":         xCampos(1, 1) = "iddoc":       xCampos(0, 2) = "1500":    xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":    xCampos(0, 1) = "descripcion": xCampos(1, 2) = "1500":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Serie":          xCampos(2, 1) = "numser":      xCampos(2, 2) = "1500":    xCampos(2, 3) = "C"
    xCampos(3, 0) = "Nro Documento":  xCampos(3, 1) = "numdoc":      xCampos(3, 2) = "1500":    xCampos(3, 3) = "C"
    
    xform.SQLCad = "SELECT mae_documento.descripcion, mae_series.iddoc, mae_series.numser, mae_series.numdoc " & _
                   " FROM mae_documento INNER JOIN mae_series ON mae_documento.id = mae_series.iddoc where iddoc = 9"

    xform.titulo = "Buscando Guias"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numser"
    xform.CampoBusca = "numser"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumSer.Text = Format(xRs("numser"), "0000")
            TxtNumDoc = HallaNumdocVenta(9, NulosC(TxtNumSer.Text), xCon)
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusPro22_Click()
    ' EJECUTA LA BUSQUEDA DE PRODUCTOS REMITIDOS
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim X, A As Integer
    Dim xCadWhere As String
    Dim xCampos(2, 4) As String
    
    ' PREPARAMOS LA CADENA WHERE PARA LA CONSULTA
    xCadWhere = ""
    For A = 1 To fg5.Rows - 1
        'A = fg5.TextMatrix(A, 5)
        xCadWhere = xCadWhere + "(vta_guiadet.idgui = " & NulosN(fg5.TextMatrix(A, 5)) & ")"
        
        If A = fg5.Rows - 1 Then Exit For
        xCadWhere = xCadWhere + " OR "
    Next A
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "6000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "id":             xCampos(1, 2) = "1500":    xCampos(1, 3) = "C"
    
    ' CARGAMOS LOS DATOS
    xform.SQLCad = "SELECT DISTINCT vta_guiadet.iditem, alm_inventario.descripcion, mae_unidades.abrev FROM mae_unidades RIGHT JOIN (vta_guiadet LEFT JOIN " _
        & " alm_inventario ON vta_guiadet.iditem = alm_inventario.id) ON mae_unidades.id = alm_inventario.idunimed " _
        & " WHERE " & Trim(xCadWhere)
    
    xform.titulo = "Buscando Productos Remitidos"
    xform.FormaBusca = CualquierParte
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"

    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtProducto.Text = xRs("descripcion")
            LblIdProd2.Caption = xRs("iditem")
            CargarItemGuias xCadWhere
        End If
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : CargarItemGuias
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS ITEMS DE UNA GUIA
'* Paranetros       : NOMBRE    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    CadWhere  |  STRING    |  ESPECIFICA LA CADENA WHERE PARA LA CONSULTA
'* Devuelve         :
'*****************************************************************************************************
Sub CargarItemGuias(CadWhere As String)
    Dim Rst As New ADODB.Recordset
    Dim xCad As String
    Dim A As Integer
    
    xCad = "SELECT DISTINCT [vta_guia]![numser]+'-'+[vta_guia]![numdoc] AS numdoc, vta_guiadet.iditem, alm_inventario.descripcion, mae_unidades.abrev AS desunimed, " _
        & " vta_guiadet.canpro, vta_guia.id FROM vta_guia LEFT JOIN (mae_unidades RIGHT JOIN (vta_guiadet LEFT JOIN alm_inventario ON vta_guiadet.iditem = alm_inventario.id) " _
        & " ON mae_unidades.id = alm_inventario.idunimed) ON vta_guia.id = vta_guiadet.idgui " _
        & " WHERE (vta_guiadet.iditem=" & NulosN(LblIdProd2.Caption) & ") AND (" & NulosC(CadWhere) & ")"
    
    RST_Busq Rst, xCad, xCon
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Fg6.Rows = 1
        For A = 1 To Rst.RecordCount
            Fg6.Rows = Fg6.Rows + 1
            Fg6.TextMatrix(A, 1) = Rst("numdoc")
            Fg6.TextMatrix(A, 2) = Rst("descripcion")
            Fg6.TextMatrix(A, 3) = Rst("desunimed")
            Fg6.TextMatrix(A, 4) = Format(Rst("canpro"), "0.0000")
            Fg6.TextMatrix(A, 5) = NulosN(Rst("canpro"))
            Fg6.TextMatrix(A, 6) = Rst("id")
            Fg6.TextMatrix(A, 7) = Rst("iditem")
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
End Sub

Private Sub CmdBusPunVen_Click()
    ' EJECUTA LA BUSQUEDA DEL PUNTO DE VENTA
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":  xCampos(1, 1) = "codcen":        xCampos(1, 2) = "1400":    xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT VTA_PuntoVenta.idcli, VTA_PuntoVenta.codcen, VTA_PuntoVenta.descripcion, VTA_PuntoVenta.id, VTA_PuntoVenta.dir " _
        & " From VTA_PuntoVenta Where (((VTA_PuntoVenta.idcli) = " & Val(LblIdCli.Caption) & " )) ORDER BY VTA_PuntoVenta.descripcion"

    xform.titulo = "Buscando Punto de Venta"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtPunVen.Text = xRs("descripcion")
            TxtDirPunVen.Text = xRs("dir")
            LblIdPunVen.Caption = xRs("id")
            TxtDirPunVen.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipItem_Click()
    ' EJECUTA LA BUSQUEDA DEL TIPO DE ITEM
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":   xCampos(0, 1) = "descripcion":    xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":        xCampos(1, 1) = "id":             xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_tipoproducto.* FROM mae_tipoproducto"
    
    xform.titulo = "Buscando Tipo de Producto"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipItem.Text = xRs("id")
            LblTipoItem = xRs("descripcion")
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing

End Sub

Private Sub CmdCan_Click()
    If NulosC(TxtNumLote.Text) = "" Then
        TxtFchProd.Valor = ""
        TxtFchVen.Valor = ""
    End If
    
    TabOne1.Enabled = True
    Toolbar1.Enabled = True
    Frame6.Visible = False
End Sub

Private Sub CmdCancela22_Click()
    Toolbar1.Enabled = True
    TabOne1.Enabled = True
    Frame7.Visible = False
End Sub

Private Sub CmdCho_Click()
    ' EJECUTA LA BUSQUEDA DE UN CHOFER
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(5, 4) As String
    
    xCampos(0, 0) = "Apellidos y Nombres":  xCampos(0, 1) = "apenom":     xCampos(0, 2) = "3500":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Categoria":            xCampos(1, 1) = "categoria":  xCampos(1, 2) = "1000":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Nº Brevete":           xCampos(2, 1) = "numbre":     xCampos(2, 2) = "1200":    xCampos(2, 3) = "C"
    xCampos(3, 0) = "Vehiculo":             xCampos(3, 1) = "marca":      xCampos(3, 2) = "1500":    xCampos(3, 3) = "C"
    xCampos(4, 0) = "Nº Placa":             xCampos(4, 1) = "numpla":     xCampos(4, 2) = "1000":    xCampos(4, 3) = "C"
    
    xform.SQLCad = "SELECT mae_chofer.*, UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom] AS apenom, mae_vehiculo.marca, mae_vehiculo.numpla" _
        & " FROM mae_vehiculo RIGHT JOIN (pla_empleados RIGHT JOIN mae_chofer ON pla_empleados.id = mae_chofer.idper) ON mae_vehiculo.id = mae_chofer.idvehiculo"

    xform.titulo = "Buscando Chofer"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "apenom"
    xform.CampoBusca = "apenom"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtChofer.Text = NulosC(xRs("apenom"))
            TxtNumBre.Text = NulosC(xRs("numbre"))
            LblIdCho.Caption = NulosN(xRs("id"))
            TxtNumPlaCar.SetFocus
            
            LblIdNumPlaca.Caption = NulosN(xRs("idvehiculo"))
            TxtNumPlaCar.Text = NulosC(xRs("numpla"))
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
    Set xRs2 = Nothing
End Sub

Private Sub CmdCli_Click()
    ' EJECUTA LA BUSQUEDA DE UN CLIENTE
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Razon Social":  xCampos(0, 1) = "nombre":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "R.U.C.":        xCampos(1, 1) = "numruc":   xCampos(1, 2) = "1400":    xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT MAE_Cliente.id, MAE_Cliente.nombre, MAE_Cliente.numruc, " _
        & " MAE_Cliente.dir From MAE_Cliente ORDER BY MAE_Cliente.Nombre"
    
    xform.titulo = "Buscando Clientes"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            '--verificar si cambian de cliente => limpiar campos
            If xRs("id") <> NulosN(LblIdCli.Caption) And NulosN(LblIdCli.Caption) <> 0 Then
                TxtIdTipDoc.Text = ""
                LblDescTipDocRef.Caption = ""
                
                TxtNumCotizacion.Text = ""
                LblIdDocRef.Caption = ""
                Fg1.Rows = 1
            End If
        
            TxtCli.Text = xRs("Nombre")
            LblIdCli.Caption = xRs("id")
            TxtDirPunVen.Text = xRs("dir")
            TxtCli_Validate True
            TxtIdTipDoc.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDatosAdicion_Click()
    ' MUESTRA EL CONTROL Frame Frame6 PARA EL INGRESO O MODIFICACION DE DATOS ADICIONALES
    'TabOne1.Enabled = False
    Toolbar1.Enabled = False
    
'    Frame6.Left = 3720
'    Frame6.Top = 2130
    Frame6.Visible = True
    Frame6.Enabled = True
    
    'TxtNumLote.SetFocus
End Sub

Private Sub CmdDelItem_Click()
    DelItem
End Sub

Private Sub cmdEliminarOKdocsproc_Click()
    Dim Rst As New ADODB.Recordset
    Dim X As Integer

    If fgdocsproc.Rows - 1 > 0 Then
        
        If fgdocsproc.Rows - 1 = 1 Then
            fgdocsproc.Rows = 1
            Fg1.Rows = 1
            Exit Sub
        Else
            With Me.Fg1
                For X = 1 To Me.Fg1.Rows - 1
                    RST_Busq Rst, "Select vta_cotizaciondet.* From vta_cotizaciondet where vta_cotizaciondet.Idvta = " & Val(fgdocsproc.TextMatrix(fgdocsproc.Row, 1)) & " and vta_cotizaciondet.IdItem = " & Val(Fg1.TextMatrix(X, 6)) & "", xCon
                    
                    If Rst.RecordCount > 0 Then
                        .TextMatrix(X, 4) = Val(.TextMatrix(X, 4)) - Rst("canpro")
                    End If
                Next
                
                ' Colocamos en el campo idest 2  de la tabla cotizacion  que indica que no esta procesado
                xCon.Execute " UPDATE vta_cotizacion  SET vta_cotizacion.idEst = 2 WHERE vta_cotizacion.id = " & Val(fgdocsproc.TextMatrix(fgdocsproc.Row, 1)) & ""
            End With
            fgdocsproc.RemoveItem fgdocsproc.Row
        End If
    End If
    Set Rst = Nothing
End Sub

Private Sub CmdEmpTra_Click()
    'EJECUTAMOS LA BUSQUEFA DE LA EMPRESA DE TRANSPORTES
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Nombre":  xCampos(0, 1) = "nombre":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":  xCampos(1, 1) = "id":        xCampos(1, 2) = "1400":    xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_emptra.nombre, mae_emptra.id From mae_emptra ORDER BY mae_emptra.nombre"
    
    xform.titulo = "Buscando Empresas de Transporte"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtDescTrans.Text = xRs("nombre")
            LblIdTran.Caption = xRs("id")
            TxtChofer.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdMot_Click()
    ' EJECUTAMOS LA MIUSQUEDA DE MOTIVOS DE GUIA
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":  xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "6200":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":       xCampos(1, 1) = "id":            xCampos(1, 2) = "1200":    xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_mottra.descripcion, mae_mottra.id From mae_mottra" _
        & " ORDER BY mae_mottra.descripcion"
    
    xform.titulo = "Buscando Motivo de Guia"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtMotivo.Text = xRs("id")
            LblMotivo.Caption = xRs("descripcion")
            'TxtNumOrd.SetFocus
            If NulosN(TxtMotivo.Text) = 1 Then
                Fg1.ColHidden(Fg1.ColIndex("PREUNI")) = False
            Else
                Fg1.ColHidden(Fg1.ColIndex("PREUNI")) = True
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdOk_Click()
    ' MUESTRA LAS GUIAS CORRESPONDIENTES AL PERIODO ESPECIFICADO
    Dim xFecha As String
    xFecha = Format(MonthView1.Value, "dd/mm/yy")
       
    xFchIni = "01/" + Mid(Format(CDate(xFecha), "dd/mm/yy"), 4, 5)
    xFchFin = Trim(Format(HallaDiasMes(CDate(xFecha)), "00")) + "/" + Mid(Format(CDate(xFecha), "dd/mm/yy"), 4, 5)
        
    TDB_FiltroLimpiar Dg1
    
    Set RstGui = Nothing
    Set Dg1.DataSource = Nothing
    DoEvents
    ' CARGA LOS REGISTROS, ESTOS DATOS SERAN VISUALIZADOS EN LA PESTAÑA COSULTA
    RST_Busq RstGui, "SELECT Format(vta_guia.numser,'0000')+'-'+Format(vta_guia.numdoc,'0000000000') AS numguia, VTA_Guia.*, MAE_Cliente.nombre, MAE_Cliente.dir, " _
        & " VTA_PuntoVenta.descripcion AS despunven, IIf(vta_guia.Anulado=0,'','Anulado') AS Anulado, VTA_PuntoVenta.dir AS Direccion, mae_emptra.nombre AS desemptra, " _
        & " mae_emptra.numruc, mae_mottra.descripcion AS descmotgui, MAE_Cliente.numruc AS RucCli, UCase([pla_empleados].[apepat])+' '+UCase([pla_empleados].[apemat])+', '+[pla_empleados].[nom] AS apenomcho, " _
        & " mae_chofer.numbre, mae_vehiculo.marca AS marcacar, mae_vehiculo.numpla, Format([vta_ventas].[numser],'0000')+'-'+Format([vta_ventas].[numdoc],'0000000000') AS numdocref , VTA_Guia.fecgiro & '' as fecgiro1 " _
        & " FROM pla_empleados RIGHT JOIN (vta_ventas RIGHT JOIN ((((((VTA_Guia LEFT JOIN MAE_Cliente ON VTA_Guia.idcli = MAE_Cliente.id) LEFT JOIN VTA_PuntoVenta " _
        & " ON VTA_Guia.idpunven = VTA_PuntoVenta.id) LEFT JOIN mae_emptra ON VTA_Guia.idemptra = mae_emptra.id) LEFT JOIN mae_mottra ON VTA_Guia.idmottra = mae_mottra.id) " _
        & " LEFT JOIN mae_chofer ON VTA_Guia.idcho = mae_chofer.id) LEFT JOIN mae_vehiculo ON VTA_Guia.idveh = mae_vehiculo.id) ON vta_ventas.id = VTA_Guia.iddocven) " _
        & " ON pla_empleados.id = mae_chofer.idper WHERE (((VTA_Guia.fecgiro)>=CDate('" & xFchIni & "') And (VTA_Guia.fecgiro)<=CDate('" & xFchFin & "'))) " _
        & " ORDER BY VTA_Guia.numser, VTA_Guia.numdoc DESC", xCon
    
    Set Dg1.DataSource = RstGui
    CmdSalir_Click
End Sub

Private Sub cmdOKdocsproc_Click()
    ' MUESTRA LOS ITEMS DE LOS DOCUMENTOS ADJUNTOS A LA GUIA
    Dim xRs As New ADODB.Recordset
    
    If fgdocsproc.Rows - 1 <= 0 Then
       cmdSalirdocsproc_Click
       Exit Sub
    End If
         
    fraconsdocref.Height = 4800
    fraconsdocref.Width = 7620
    fraconsdocref.Visible = True

    ' CARGAMOS LOS ITEMS
    RST_Busq xRs, " SELECT vta_cotizacion.numdoc, alm_inventario.descripcion, vta_cotizaciondet.canpro, mae_unidades.abrev " _
        & " FROM vta_cotizacion INNER JOIN (mae_unidades INNER JOIN (alm_inventario INNER JOIN vta_cotizaciondet ON alm_inventario.id = vta_cotizaciondet.iditem) " _
        & " ON mae_unidades.id = alm_inventario.idunimed) ON vta_cotizacion.id = vta_cotizaciondet.idvta " _
        & " WHERE vta_cotizacion.id =  " & Val(fgdocsproc.TextMatrix(Me.fgdocsproc.Row, 1)), xCon

     With Me.Fgdocref
         .Rows = 1
         .Cols = 5
                  
        .ColWidth(1) = 1500  'Nro Doc
        .ColWidth(2) = 3500 'Item
        .ColWidth(3) = 800 'Cantidad
        .ColWidth(4) = 800 'Uni Med
     
        .TextMatrix(0, 1) = "Nro Doc"
        .TextMatrix(0, 2) = "Item"
        .TextMatrix(0, 3) = "Cantidad"
        .TextMatrix(0, 4) = "Unid Med"
       
       Do While Not xRs.EOF
            .AddItem ""
            .TextMatrix(.Rows - 1, 1) = Format(xRs("numdoc"), "000000000")
            .TextMatrix(.Rows - 1, 2) = xRs("descripcion")
            .TextMatrix(.Rows - 1, 3) = xRs("canpro")
            .TextMatrix(.Rows - 1, 4) = xRs("abrev")
             xRs.MoveNext
       Loop
         
    End With
    Fradocsproc.Visible = False
    Set xRs = Nothing
End Sub

Private Sub CmdOkRef_Click()
    Dim Rst As New ADODB.Recordset
    Dim swaviso As Byte
    Dim Pos As Integer
    Dim X As Integer

    If Fgdocref.Rows - 1 = 0 Then
        CmdSalirRef_Click
        Exit Sub
    End If

    RST_Busq Rst, " SELECT vta_cotizaciondet.iditem,  alm_inventario.idtipven, alm_inventario.descripcion, vta_cotizaciondet.canpro, " _
        & " vta_cotizaciondet.idunimed, mae_unidades.abrev, vta_cotizacion.idcli,vta_cotizacion.id, mae_cliente.nombre, " _
        & " mae_cliente.numruc FROM mae_cliente INNER JOIN (vta_cotizacion INNER JOIN (mae_unidades INNER JOIN (alm_inventario " _
        & " INNER JOIN vta_cotizaciondet ON alm_inventario.id = vta_cotizaciondet.iditem) ON mae_unidades.id = alm_inventario.idunimed) " _
        & " ON vta_cotizacion.id = vta_cotizaciondet.idvta) ON mae_cliente.id = vta_cotizacion.idcli " _
        & " WHERE vta_cotizacion.id =" & Val(Fgdocref.TextMatrix(Fgdocref.Row, 1)) & " ", xCon
    
    If Rst.RecordCount > 0 Then
        TxtCli = Rst("nombre")
        LblIdCli = Rst("idcli")
                
        'Añadimos a la lista de documentos cotizados
        With fgdocsproc
            .AddItem ""
            .TextMatrix(.Rows - 1, 1) = Rst("id")
            .TextMatrix(.Rows - 1, 2) = Fgdocref.TextMatrix(Fgdocref.Row, 2)
            .TextMatrix(.Rows - 1, 3) = Trim(Fgdocref.TextMatrix(Fgdocref.Row, 5))
        End With
        
        With Fg1
            Do While Not Rst.EOF
                swaviso = 0
                Pos = 0
                    
                For X = 1 To .Rows - 1
                    If Rst("iditem") = Val(.TextMatrix(X, 1)) Then
                        swaviso = 1
                        Pos = X
                        Exit For
                    End If
                Next
                    
                If swaviso = 0 Then
                    .AddItem ""
                    .TextMatrix(.Rows - 1, 1) = Rst("iditem")
                    .TextMatrix(.Rows - 1, 2) = Rst("descripcion")
                    .TextMatrix(.Rows - 1, 3) = Rst("abrev")
                    .TextMatrix(.Rows - 1, 4) = Rst("canpro")
                    .TextMatrix(.Rows - 1, 5) = Rst("idunimed")
                Else
                    .TextMatrix(Pos, 4) = Val(.TextMatrix(Pos, 4)) + Rst("canpro")
                End If
                
                Rst.MoveNext
            Loop
            
            'Estado de la cotizacones
            '1 pendiente
            '2 aprobada
            '3 procesada
            '4 rechazada
                
            'Colocamos en el campo idest 3  de la tabla guia que indica que esta procesado
            'xCon.Execute " UPDATE vta_cotizacion SET vta_cotizacion.idEst = 3 WHERE vta_cotizacion.id = " & Val(Fgdocref.TextMatrix(Fgdocref.Row, 1)) & ""
        End With
    End If

    Toolbar1.Enabled = True
    TabOne1.Enabled = True
    fraconsdocref.Visible = False
    
    Set Rst = Nothing
End Sub

Private Sub cmdokseldoc_Click()
    Dim Rpta As Integer
    Dim RstCab As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim xId As Double
    Dim xnumdoc As String
    Dim xnumser As String
    Dim xFecha As String
            
    ' EMITE UN DOCUMENTO ANULADO
    ' VERIFICA QUE LOS DASTOS NECESARIOS SE HAYAN INGRESADO CORRECTAMENTE
    If NulosC(TxtNumSer2.Text) = "" Then
        MsgBox "No ha especificado el número de serie del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer2.SetFocus
        Exit Sub
    End If
    
    If TxtNumDocGen.Text = "" Then
        MsgBox "No ha especificado el número del documento a generar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDocGen.SetFocus
        Exit Sub
    End If

    xFecha = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    
    ' VERIFICAMOS QUE LA FECHA DE EMISION DEL DOCUMENTO CORRESPONDA AL PERIODO ESPECIFICADO
    If CDate(TxtFchEmiAnul.Valor) < CDate(xFecha) Then
        MsgBox "La fecha del documento no corresponde la periodo contable especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("Esta seguro de emitir una guia como anulada", vbYesNo + vbQuestion + vbDefaultButton1, xTitulo)
    If Rpta = vbNo Then
        Exit Sub
    End If
    
On Error GoTo LaCague
    xCon.BeginTrans
    
    RST_Busq RstCab, "SELECT * FROM vta_guia", xCon
    
    xId = HallaCodigoTabla("vta_guia", xCon, "id")
    RstCab.AddNew
    RstCab("id") = xId
    RstCab("numser") = NulosC(TxtNumSer2.Text)
    xnumser = NulosC(TxtNumSer2.Text)
    RstCab("numdoc") = NulosC(TxtNumDocGen.Text)
    xnumdoc = NulosC(TxtNumDocGen.Text)
    RstCab("idcli") = 0
    RstCab("tipdoc") = 9
    RstCab("fecgiro") = TxtFchEmiAnul.Valor 'CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    RstCab("numordcom") = "ANULADA"
    RstCab("FchEmiord") = TxtFchEmiAnul.Valor
    RstCab("FchEntord") = TxtFchEmiAnul.Valor
    RstCab("idpunven") = 0
    RstCab("idmottra") = 0
    RstCab("idemptra") = 0
    RstCab("idcho") = 0
    RstCab("idveh") = 0
    RstCab("anulado") = -1
    RstCab("fchpro") = TxtFchEmiAnul.Valor
    RstCab("fchven") = TxtFchEmiAnul.Valor
    RstCab("numlote") = "ANULADO"
    RstCab.Update
        
    ' Validar si el nro de documento existe solo en modo adicionar documento
'    RstCab.AddNew
'    xId = HallaCodigoTabla("vta_guia", xCon, "id")
'    RstCab("id") = xId
'    RstCab("numser") = xnumser
'    xnumdoc = HallaNumdocVenta(9, Format(xnumser, "0000"), xCon)
'    RstCab("numdoc") = CLng(xnumdoc)
'    RstCab("idcli") = 1
'    RstCab("tipdoc") = 9
'    RstCab("fecgiro") = CDate(TxtFchEmi.Valor)
'    RstCab("numordcom") = "ANULADA"
'    If NulosC(TxtFchEmiPed.Valor) <> "" Then RstCab("FchEmiord") = TxtFchEmiPed.Valor
'    If NulosC(TxtFchEnt.Valor) <> "" Then RstCab("FchEntord") = TxtFchEnt.Valor
'    RstCab("idpunven") = 0
'    RstCab("idmottra") = 0
'    RstCab("idemptra") = 0
'    RstCab("idcho") = 0
'    RstCab("idveh") = 0
'    RstCab("anulado") = -1
'    RstCab("Estado") = 0 'sin facturar
'    'Grabamos datos adicionales
'    If NulosC(TxtFchProd.Valor) <> "" Then RstCab("fchpro") = TxtFchProd.Valor
'    If NulosC(TxtFchVen.Valor) <> "" Then RstCab("fchven") = TxtFchVen.Valor
'    RstCab("numlote") = "ANULADO"
'    RstCab.Update
    
    'Actualizamos el numero de documento en la tabla Mae_series
    Call ActualizaNroDocumento(CLng(xnumdoc), 9, CLng(xnumser))
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, 1, QueHace, xHorIni, Time, Date, xCon, xId
    xCon.CommitTrans
    
    MsgBox "La guia anulada se genero con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    cmdsalirseldoc_Click
    RstGui.Requery
    Dg1.Refresh
    Exit Sub

LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set xRs = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Exit Sub
End Sub

Private Sub CmdPlaCar_Click()
    ' EJECUTA LA BUSQUEDA DE UN VEHICULO DE TRANSPORTE
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Placa":   xCampos(0, 1) = "numpla":   xCampos(0, 2) = "5000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Marca":   xCampos(1, 1) = "marca":   xCampos(1, 2) = "1400":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Codigo":  xCampos(2, 1) = "id":      xCampos(2, 2) = "1400":    xCampos(2, 3) = "C"
    
    xform.SQLCad = "SELECT mae_vehiculo.numpla, mae_vehiculo.marca, mae_vehiculo.id From mae_vehiculo " _
        & " ORDER BY mae_vehiculo.numpla"
    
    xform.titulo = "Buscando Vehiculos"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numpla"
    xform.CampoBusca = "numpla"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumPlaCar.Text = xRs("numpla")
            LblIdNumPlaca.Caption = xRs("id")
            Fg1.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdSalir_Click()
    Toolbar1.Enabled = True
    TabOne1.Enabled = True
    Frame5.Visible = False
End Sub

Private Sub cmdSalirdocsproc_Click()
    Fradocsproc.Visible = False
    Toolbar1.Enabled = True
    TabOne1.Enabled = True
End Sub

Private Sub CmdSalirRef_Click()
    fraconsdocref.Visible = False
    Toolbar1.Enabled = True
    TabOne1.Enabled = True
End Sub

Private Sub cmdsalirseldoc_Click()
    ' SALE DEL FRAME Fraseldoc
    QueHace = 3
    ActivarEntorno
    Fraseldoc.Visible = False
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstGui
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LAS COLUMNAS DEL CONTROL DataGrid Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstGui.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstGui("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    ' EJECUTA LA BUSQUEDA DE ITEMS
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim X As Integer
    Dim nSQLId As String
    Dim xCampos(2, 4) As String
    Dim nTitulo As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "codpro":           xCampos(0, 2) = "1500":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":      xCampos(1, 2) = "6000":    xCampos(1, 3) = "C"
    
    nSQLId = GENERAR_SQL_ID(Fg1, Fg1.ColIndex("IDITEM"), " AND alm_inventario.id", "NOT IN", True)
    
    cSQL = "SELECT alm_inventario.id, alm_inventario.codpro, alm_inventario.Descripcion, MAE_Unidades.Descripcion as [UnidMed], MAE_Unidades.id as [iduni] " _
        + vbCr + "FROM alm_inventario LEFT JOIN MAE_Unidades ON alm_inventario.idunimed = MAE_Unidades.id " _
        + vbCr + "WHERE  alm_inventario.tippro = " & NulosN(TxtTipItem.Text) & "  AND Alm_inventario.activo = -1 " & nSQLId _
        + vbCr + "ORDER BY alm_inventario.Descripcion"
    
    nTitulo = "Buscando Productos"
     
    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos(), "Buscando " & nTitulo, "codpro", "codpro", Principio
    
    If xRs.State = 1 Then
        If xRs.RecordCount = 0 Then Exit Sub
        Fg1.TextMatrix(Fg1.Row, 1) = NulosC(xRs("codpro"))
        Fg1.TextMatrix(Fg1.Row, 2) = NulosC(xRs("descripcion"))
        Fg1.TextMatrix(Fg1.Row, 3) = NulosC(xRs("unidmed"))
        Fg1.TextMatrix(Fg1.Row, 5) = NulosN(xRs("iduni"))
        Fg1.TextMatrix(Fg1.Row, 6) = NulosN(xRs("id"))
        With Fg1
            .Select Fg1.Rows - 1, 4, Fg1.Rows - 1, 4
        End With
        Fg1.SetFocus
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    If Fg1.TextMatrix(Fg1.Row, 1) = "" Then
        Fg1.TextMatrix(Fg1.Row, 1) = ""
        Fg1.TextMatrix(Fg1.Row, 2) = ""
        Fg1.TextMatrix(Fg1.Row, 3) = ""
        Fg1.TextMatrix(Fg1.Row, 4) = ""
        Fg1.TextMatrix(Fg1.Row, 5) = ""
    End If
End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        Fg1.Editable = flexEDNone
        Exit Sub
    End If
    
    If Fg1.Col = 1 Or Fg1.Col = 4 Or Fg1.Col = 8 Or Fg1.Col = Fg1.ColIndex("PREUNI") Then
        Fg1.Editable = flexEDKbdMouse
    Else
        Fg1.Editable = flexEDNone
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If Fg1.Rows = 1 Then Exit Sub
    
    If KeyCode = 45 Then
        Fg1.Rows = Fg1.Rows + 1
        With Fg1
            .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, 1
        End With
        Fg1_CellButtonClick Fg1.Rows - 1, 1
    End If
    
    If KeyCode = 46 Then
        DelItem
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If QueHace = 3 Then Exit Sub
        PopupMenu menu_1
    End If
End Sub

Private Sub fg5_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Col = 1 Then
        If NulosN(LblIdcli2.Caption) = 0 Then
            MsgBox "No ha especificado el cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtCliente.SetFocus
            Exit Sub
        End If
        CargarGuia
    End If
End Sub

Private Sub fg5_EnterCell()
    If fg5.Col = 1 Then
        fg5.Editable = flexEDKbdMouse
    Else
        fg5.Editable = flexEDNone
    End If
End Sub

Private Sub Fg6_EnterCell()
    If Fg6.Col = 5 Then
        Fg6.Editable = flexEDKbdMouse
    Else
        Fg6.Editable = flexEDNone
    End If
End Sub

Private Sub Fgdocref_DblClick()
    Call CmdOkRef_Click
End Sub

Private Sub fgdocsproc_DblClick()
    Call cmdOKdocsproc_Click
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    Dim Rpta As Integer
    If SeEjecuto = False Then
        SeEjecuto = True
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        mMesActivo = xMes
        Dim xFecha As String
        TabOne1.CurrTab = 0
        xFecha = CDate("01/" & Format(mMesActivo, "00") & "/" & AnoTra)
            
        xFchIni = "01/" + Mid(Format(CDate(xFecha), "dd/mm/yy"), 4, 5)
        xFchFin = Trim(Format(HallaDiasMes(CDate(xFecha)), "00")) + "/" + Mid(Format(CDate(xFecha), "dd/mm/yy"), 4, 5)
        ' CARGA LOS REGISTRO DE LA TABLA Vta_guias
        pCargarGrid
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    SeEjecuto = False
    QueHace = 3
    iniciarCampos
End Sub

Private Sub iniciarCampos()
    Frame1.BackColor = &H8000000F
    Frame2.BackColor = &H8000000F
    
    Fg1.ColWidth(0) = 0
    Fg1.ColWidth(5) = 0
    Fg1.ColWidth(6) = 0
    Fg1.ColWidth(7) = 0
    Fg1.GridLines = flexGridInset
    
    fgdocsproc.Rows = 1
    CaracteresNumericos = "0123456789." & Chr(8)
    Fg1.SelectionMode = flexSelectionByRow
    TxtFchEmi.Valor = Date
    TxtFchEmiPed.Valor = Date
    TxtFchEnt.Valor = Date
    TxtFchProd.Valor = Date
    TxtFchVen.Valor = ""
    
    swguiafact = 0
    TxtFchEmi.Valor = Date
    TxtFchEmiPed.Valor = Date
    TxtFchEnt.Valor = Date
    MonthView1.Value = Date
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede cerrar este formulario mientras este agregando o modificando datos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = 1
        Exit Sub
    End If
End Sub

Private Sub menu_1_1_Click()
    AddItem
End Sub

Private Sub menu_1_3_Click()
    DelItem
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace = 1 Then Exit Sub
        'Validamos si la cuadricula tiene datos
        If RstGui.RecordCount = 0 Then
            MsgBox "No existe información para visualizar", vbInformation, Me.Caption
            Blanquea
            Exit Sub
        Else
            Blanquea
            MuestraSegundoTab
        End If
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Anular
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ANULA UNA GUIA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Anular()
    Dim xRs As New ADODB.Recordset
    ' Validamos si la guia esta facturada
    RST_Busq xRs, " SELECT mae_documento.descripcion, vta_ventas.numser, vta_ventas.numdoc " & _
        " FROM mae_documento INNER JOIN (vta_ventas INNER JOIN vta_guia ON vta_ventas.id = vta_guia.iddocven) " _
        & " ON mae_documento.id = vta_ventas.tipdoc WHERE vta_ventas.id = " & RstGui("iddocven") & "", xCon
           
    If xRs.RecordCount > 0 Then
        MsgBox "La guia esta facturada " & " Según " & xRs!Descripcion & " Nº " & RellenaNumdoc(xRs("numser"), xRs("numdoc"))
        Set xRs = Nothing
        Exit Sub
    End If
    
    Dim Rpta As Integer
    Dim A As Integer
    Rpta = MsgBox("¿Esta seguro de anular la guia Nº " + Format(RstGui("numser"), "0000") & "-" & Format(RstGui("numdoc"), "0000000000") + "?", vbYesNo + vbDefaultButton1 + vbQuestion, xTitulo)
    
    xnumreg = RstGui.RecordCount
    
    If Rpta = vbYes Then
        mIdRegistro = RstGui("id")
        ' PROCEDEMOS A AUNLAR LA GUIA
        xCon.Execute "UPDATE vta_guia SET vta_guia.Anulado = -1, " _
            & " vta_guia.idcli = 1,  vta_guia.numordcom= 'ANULADA', idpunven = 0" _
            & " WHERE vta_guia.id = " & RstGui("id") & " "
        
        ' ACTUALIZAMOS EL STOCK
        Call ActualizarStock("E", RstGui("id"))
        xCon.Execute "DELETE * FROM vta_guiadet WHERE vta_guiadet.idgui = " & RstGui("id") & ""
        
        If NulosN(TxtIdTipDoc.Text) = 5 Then
            ' eliminamos la referencia del documento a la orden de pedido
'            xCon.Execute "UPDATE ped_pedidodetent SET ped_pedidodetent.idtipdoc = 0, ped_pedidodetent.iddocven = 0, ped_pedidodetent.estado = 2" _
'                & " WHERE (((ped_pedidodetent.idtipdoc)=1) AND ((ped_pedidodetent.iddocven)=" & RstGui("id") & "))"
            
            xCon.Execute "UPDATE ped_pedidodet SET ped_pedidodet.idtipdoc = 0, ped_pedidodet.iddocven = 0, ped_pedidodet.estado = 2" _
                & " WHERE (((ped_pedidodet.idtipdoc)=1) AND ((ped_pedidodet.iddocven)=" & RstGui("id") & "))"
        End If
        
        RstGui.Requery
        Dg1.Refresh
        
        
        If RstGui.RecordCount <> 0 Then
            RstGui.MoveFirst
            RstGui.Find "id=" & mIdRegistro
            If RstGui.EOF = True Then RstGui.MoveFirst
        End If
        
        MsgBox "La guia se anulo con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA vta_guia
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim xRs As New ADODB.Recordset
    Dim Rpta As Integer
    Dim A As Integer
    Dim cSQL As String
    Dim F As New SistemaLogica.Funciones
    
    If RstGui.RecordCount = 0 Then
        MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de Eliminar la guia Nº " + Format(RstGui("numser"), "0000") & "-" & Format(RstGui("numdoc"), "0000000000") + "?", vbYesNo + vbDefaultButton1 + vbQuestion, xTitulo)
    
    If Rpta = vbYes Then
        '***********
        '-'-'-'-'-'-
        ' se actualiza las cantidades entregadas de la tabla ped_pedidodet
        cSQL = "SELECT ped_pedidodet.canpro, vta_guiadet.canpro As canprogui, ped_pedidodet.idpeddet" _
             + vbCr + "FROM vta_guiadet LEFT JOIN ped_pedidodet ON vta_guiadet.iddocref = ped_pedidodet.idpeddet " _
             + vbCr + "WHERE (((ped_pedidodet.canpro) Is Not Null) AND ((vta_guiadet.idgui)=" & RstGui("id") & "));"
        
        RST_Busq xRs, cSQL, xCon
        
        If xRs.State = 1 Then
            If Not xRs.EOF Then
                xRs.MoveFirst
                While Not xRs.EOF
                    xCon.Execute "UPDATE ped_pedidodet SET ped_pedidodet.canproent = [ped_pedidodet]![canproent] - " & xRs("canprogui") & " " _
                        & "WHERE (((ped_pedidodet.idpeddet)=" & xRs("idpeddet") & "));"
                    
                    xRs.MoveNext
                Wend
            End If
        End If
        '-'-'-'-'-'-
        '***********
        
        xnumreg = RstGui.RecordCount
        
        Call ActualizarStock("E", RstGui("id"))
        
        xCon.Execute "DELETE * FROM vta_guia    WHERE vta_guia.id = " & RstGui("id") & ""
        xCon.Execute "DELETE * FROM vta_guiadet WHERE vta_guiadet.idgui =" & RstGui("id") & ""
        
        '***********************************
        ' Eliminamos los movimientos generados
        If F.NuloNumeric(F.KeyValue("CreacionMovimientoAutoGuiaRemision", xCon)) = -1 Then
            ' Verificamos si ya tiene registro en movimientos
            Dim database As New SistemaData.EDataBase
            Dim record As New ADODB.Recordset
            Dim Movimiento As New AlmacenEntidad.EMovimiento
            
            Set database.Connection = xCon
            database.CommandText = "SELECT alm_ingreso.id AS idmov " _
                        + vbCr + "FROM alm_ingreso " _
                        + vbCr + "WHERE (((alm_ingreso.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoGuiaRemision", xCon)) & ") AND ((alm_ingreso.iddocref)=" & F.NuloNumeric(RstGui("id")) & "))"
            Set record = database.GetRecordset
            If record.RecordCount > 0 Then
                Movimiento.IdMovimiento = F.NuloNumeric(record("idmov"))
                Set Movimiento.Conexion = xCon
                Movimiento.Delete CLng(xIdUsuario), F.MachineName
            End If
        End If
        '***********************************
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstGui("id") & " AND idform = " & IdMenuActivo
        
        
        RstGui.Requery
        Dg1.Refresh
        MsgBox "La guia se elimino con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
    Set xRs = Nothing
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then
        If RstGui.RecordCount = 0 Then
            MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
            Exit Sub
        End If
        
        If RstGui("anulado") = -1 Then
            MsgBox "No puede modificar una guia anulada", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
            Exit Sub
        Else
            Modificar
        End If
    End If
    
    If Button.Index = 3 Then
        If RstGui.RecordCount = 0 Then
            MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
            Exit Sub
        End If
    
        'Validamos si la guia esta anulada
        If RstGui("Anulado") = -1 Then
            MsgBox "La guia ya fue anulada, seleccione otra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If

        Anular
    End If
    
    If Button.Index = 5 Then
'        xnumreg = RstGui.RecordCount
        If Grabar = True Then
'            Cancelar
'            If QueHace = 1 Then
'                While (RstGui.RecordCount < (xnumreg + 1))
'                    RstGui.Requery
'                    Dg1.Refresh
'                Wend
'            Else
'                Dim A As Integer
'                For A = 1 To 5
'                    RstGui.Requery
'                    Dg1.Refresh
'                Next A
'            End If
            
            
            Cancelar
            RstGui.Requery
            Dg1.Refresh
            
            If RstGui.RecordCount <> 0 Then
                RstGui.MoveFirst
                RstGui.Find "id=" & mIdRegistro
                If RstGui.EOF = True Then RstGui.MoveFirst
            End If
            
        End If
    End If
    
    If Button.Index = 6 Then
        Cancelar
    End If
    
    If Button.Index = 8 Then
        FiltrarGuias (1)
    End If
    
    If Button.Index = 9 Then '--quitar filtro
        TDB_FiltroLimpiar Dg1
        RstGui.Filter = ""
        RstGui.Requery
        
    End If
    
    If Button.Index = 11 Then
        CambiarMes
    End If
    
    If Button.Index = 13 Then
        Imprimir
    End If
    
    If Button.Index = 15 Then
        Set RstGui = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : CambiarMes
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CAMBIA EL MES DE TRABAJO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CambiarMes()
    TabOne1.CurrTab = 0
    mMesActivo = SeleccionaMes(xCon)
    pCargarGrid
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN LA TABLA Vta_Guia, ESTA FUNCION DEVUELVE VERDADERO CUANDO
'*                    TIENE EXITO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    Dim F As New SistemaLogica.Funciones
    'VERIFICAOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If TxtFchEmi.Valor = "" Then
        MsgBox "No ha especificado la fecha de emision de la guia", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchEmi.SetFocus
        Exit Function
    End If
    
    If Year(TxtFchEmi.Valor) <> AnoTra Then
        MsgBox "El año de trabajo no es correcto" & vbCr & "Año de trabajo: " & AnoTra, vbInformation + vbOKOnly
        
       Exit Function
    End If
    If NulosC(TxtCli.Text) = "" Then
        MsgBox "No ha especificado el cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtCli.SetFocus
        Exit Function
    End If
    
    If NulosC(TxtMotivo.Text) = "" Then
        MsgBox "No ha especificado de traslado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtMotivo.SetFocus
        Exit Function
    End If
    
'    If NulosC(TxtDirPunVen.Text) = "" Then
'        MsgBox "No ha especificado el lugar de entrega", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtDirPunVen.SetFocus
'        Exit Function
'    End If
    
'    If NulosC(TxtDescTrans.Text) = "" Then
'        MsgBox "No ha especificado la empresa de transportes", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtDescTrans.SetFocus
'        Exit Function
'    End If
    
'    If NulosC(TxtChofer.Text) = "" Then
'        MsgBox "No ha especificado el nombre del chofer", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtChofer.SetFocus
'        Exit Function
'    End If
    
'    If NulosC(TxtNumPlaCar.Text) = "" Then
'        MsgBox "No ha especificado el numero de placa del vehiculo de transporte", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtNumPlaCar.SetFocus
'        Exit Function
'    End If
    
    If NulosN(TxtIdAlm.Text) = 0 Then
        MsgBox "No ha especificado el Almacen", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdAlm.SetFocus
        Exit Function
    End If
    
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado items para el documento de salida", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    Dim A As Integer
    Dim SALIR As Boolean
    
    SALIR = False
    ' VERIFICAMOS QUE LOS ITEMS ESTEN CORRECTAMENTE INGRESADOS
    For A = 1 To Fg1.Rows - 1
        If NulosC(Fg1.TextMatrix(A, 1)) = "" Then
            Fg1.RemoveItem (A)
        Else
            If NulosC(Fg1.TextMatrix(A, 4)) = "" Then
                MsgBox "No ha especificado cantidad para el producto " + Fg1.TextMatrix(A, 2), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Fg1.SetFocus
                SALIR = True
                Exit For
            End If
        End If
    Next A
    If SALIR = True Then
        Exit Function
    End If
    
    Dim xRs As New ADODB.Recordset
    Dim Rstnumgui As New ADODB.Recordset
    
    If QueHace = 1 Then
        ' CUANDO ES UN NUEVO REGISTRO
        ' Validar si el nro de documento existe solo en modo adicionar documento
        RST_Busq xRs, "SELECT * FROM VTA_Guia WHERE VTA_Guia.Numser  = '" & Trim(TxtNumSer.Text) & "' and VTA_Guia.numdoc  = '" & Trim(TxtNumDoc.Text) & "'", xCon
        
        If xRs.RecordCount > 0 Then
            RST_Busq Rstnumgui, "SELECT vta_guia.numser, vta_guia.numdoc From vta_guia Where (((vta_guia.numser) = '" & Format(TxtNumSer.Text, "0000") & "'))" _
                & " ORDER BY vta_guia.numdoc", xCon
    
            If Rstnumgui.RecordCount = 0 Then
                TxtNumDoc.Text = "0000000001"
            Else
                TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
                TxtNumDoc.Text = Format((Val(Rstnumgui("numdoc")) + 1), "0000000000")
            End If
        End If
    End If
    Set Rstnumgui = Nothing
    
    Dim RstDet As New ADODB.Recordset
    Dim RstCab As New ADODB.Recordset
    Dim xId As Double
    
   On Error GoTo LaCague
    
    xCon.BeginTrans
    
    swguiafact = 1
    
    If QueHace = 1 Then
        ' CUANDO ES UN NUEVO REGISTRO
        xId = HallaCodigoTabla("vta_guia", xCon, "id")
        RST_Busq RstCab, "SELECT TOP 1 * FROM vta_guia", xCon
        RST_Busq RstDet, "SELECT TOP 1 * FROM vta_guiadet", xCon
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstGui("id")
        RST_Busq RstCab, "SELECT * FROM vta_guia WHERE id = " & xId & "", xCon
        'Actualizamos  stock
        Call ActualizarStock("E", xId)
        xCon.Execute "DELETE * FROM vta_guiadet WHERE idgui =" & xId & ""
        RST_Busq RstDet, "SELECT * FROM vta_guiadet", xCon
        
    End If
    
    mIdRegistro = xId
    ' GRABAMOS LOS DATOS PRINCIPALES
    RstCab("numdoc") = TxtNumDoc
    RstCab("numser") = TxtNumSer.Text
    RstCab("idcli") = NulosN(LblIdCli.Caption)
    RstCab("tipdoc") = 9
    RstCab("fecgiro") = CDate(TxtFchEmi.Valor)
    RstCab("idmottra") = NulosN(TxtMotivo.Text)
    RstCab("tippro") = NulosN(TxtTipItem.Text)
    RstCab("numordcom") = NulosC(TxtNumOrd.Text)
    RstCab("idalm") = NulosN(TxtIdAlm.Text)
    
    If NulosC(TxtFchEmiPed.Valor) <> "" Then RstCab("FchEmiord") = TxtFchEmiPed.Valor
    If NulosC(TxtFchEnt.Valor) <> "" Then RstCab("FchEntord") = TxtFchEnt.Valor
    
    RstCab("idpunven") = NulosN(LblIdPunVen.Caption)
    RstCab("idemptra") = NulosN(LblIdTran.Caption)
    RstCab("idcho") = NulosN(LblIdCho.Caption)
    RstCab("idveh") = NulosN(LblIdNumPlaca.Caption)
    
    'grabamos datos adicionales
    If NulosC(TxtFchProd.Valor) <> "" Then RstCab("fchpro") = TxtFchProd.Valor
    If NulosC(TxtFchVen.Valor) <> "" Then RstCab("fchven") = TxtFchVen.Valor
    RstCab("numlote") = NulosC(TxtNumLote.Text)
    If NulosN(TxtIdTipDoc.Text) = 5 Then JALOPEDIDO = True Else JALOPEDIDO = False
    
    RstCab.Update
    
    ' GRABAMOS EL DETALLE DE LA GUIA
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idgui") = xId
        RstDet("iditem") = NulosC(Fg1.TextMatrix(A, 6))
        RstDet("canpro") = NulosN(Fg1.TextMatrix(A, 4))
        RstDet("idunimed") = NulosC(Fg1.TextMatrix(A, 5))
        RstDet("lote") = NulosC(Fg1.TextMatrix(A, 8))
        RstDet("preuni") = NulosN(Fg1.TextMatrix(A, Fg1.ColIndex("PREUNI")))
        If JALOPEDIDO Then RstDet("iddocref") = NulosN(Fg1.TextMatrix(A, 0))
        RstDet.Update
    Next A
    
'    '*************************************************
'    ' Modificado por Jose Chacon 20160511
'    ' Se comenta debido a que no se ve utilidad en el momento
'    ' Actualizamos stock
'    Call ActualizarStock("S", xId)
'
'    ' Actualizamos el numero de documento en la tabla Mae_series
'    Call ActualizaNroDocumento(NulosN(TxtNumDoc.Text), 9, NulosN(TxtNumSer))
'    '*************************************************
        
    Set RstCab = Nothing
    Set RstDet = Nothing
    
    ' Actualizamos el pedido
    If JALOPEDIDO = True Then
        ' Por defecto se agrega el tipo de documento de referencia 5 = orden de pedido
        'Dim NumGui As String
        'NumGui = NulosC(TxtNumSer.Text) & "-" & NulosC(TxtNumDoc.Text)
'        xCon.Execute "UPDATE vta_guia SET vta_guia.idtipdocref = 5, vta_guia.iddocref = " & VAR_IDPEDIDO & "" _
'            & " WHERE (((vta_guia.id)=" & xId & "))"
        
        xCon.Execute "UPDATE vta_guia SET vta_guia.idtipdocref = 5 WHERE (((vta_guia.id)=" & xId & "))"
        
        If VAR_IDPEDIDO <> 0 Then
            xCon.Execute "UPDATE ped_pedidodet SET ped_pedidodet.idtipdoc = 1, ped_pedidodet.iddocven = " & xId & "" _
                & " WHERE (((ped_pedidodet.idped)=" & VAR_IDPEDIDO & ") AND ((ped_pedidodet.fchent)=CDate('" & VAR_FECHAPEDIDO & "')))"
        End If
        
        For A = 1 To Fg1.Rows - 1
            If QueHace = 2 Then
                xCon.Execute "UPDATE ped_pedidodet SET ped_pedidodet.canproent = [ped_pedidodet]![canproent] - " & Fg1.TextMatrix(A, 7) & " + " & Fg1.TextMatrix(A, 4) & " " _
                & "WHERE (((ped_pedidodet.idpeddet)=" & NulosN(Fg1.TextMatrix(A, 0)) & "));"
            End If
            
            If QueHace = 1 Then
                xCon.Execute "UPDATE ped_pedidodet SET ped_pedidodet.canproent = [ped_pedidodet]![canproent] + " & Fg1.TextMatrix(A, 4) & " " _
                & "WHERE (((ped_pedidodet.idpeddet)=" & NulosN(Fg1.TextMatrix(A, 0)) & "))"
            End If
        Next A
        
    End If
    
    ' Creamos el movimiento automatico
    If F.NuloNumeric(F.KeyValue("CreacionMovimientoAutoGuiaRemision", xCon)) = -1 Then ' Se valida la fecha de cierre de mes
        If F.MesCerradoOpcion(F.RetornarMesFecha(CDate(TxtFchEmi.Valor)), CLng(F.KeyValue("IdOpcionSistemaMovimientoAlmacen", xCon)), xCon) Then
            Err.Raise &HFFFFFF01, , "El mes al que pertenece el documento se encuentra cerrado para la opcion: [Ingresos y Salidas de almacen] y no se pueden generar movimientos automaticos, modifique la fecha o aperture el mes "
        Else
            ' Verificamos si ya tiene registro en movimientos
            Dim database As New SistemaData.EDataBase
            Dim record As New ADODB.Recordset
            Dim Movimiento As New AlmacenEntidad.EMovimiento
            
            Set database.Connection = xCon
            database.CommandText = "SELECT alm_ingreso.id AS idmov " _
                        + vbCr + "FROM alm_ingreso " _
                        + vbCr + "WHERE (((alm_ingreso.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoGuiaRemision", xCon)) & ") AND ((alm_ingreso.iddocref)=" & xId & "))"
            Set record = database.GetRecordset
            ' Se cambia para que solo elimine movimientos de registros ya creados.
            If record.RecordCount > 0 And QueHace <> 1 Then
                ' Eliminamos todos los movimientos ya generados
                record.MoveFirst
                While Not record.EOF
                    Dim mMovAux As New AlmacenEntidad.EMovimiento
                    Set mMovAux.Conexion = xCon
                    mMovAux.IdMovimiento = F.NuloNumeric(record("idmov"))
                    mMovAux.Delete CLng(xIdUsuario), F.MachineName
                    record.MoveNext
                Wend
                Set record = Nothing
            End If
            ' Cabecera
            Movimiento.IdMovimiento = 0
            Movimiento.IdTipoMovimiento = 0
            Movimiento.FechaMovimiento = CDate(TxtFchEmi.Valor)
            Movimiento.NumeroSerie = F.NuloString(TxtNumSer.Text)
            Movimiento.NumeroDocumento = F.HallaNumeroDocumento("alm_ingreso", "'" & Movimiento.NumeroSerie & "'", "numser", xCon)
            Movimiento.IdEstado = F.NuloNumeric(F.KeyValue("EstadoAprobadoMovimiento", xCon))
            Movimiento.IdAlmacen = F.NuloNumeric(TxtIdAlm.Text)
            Movimiento.Glosa = F.NuloString(LblMotivo.Caption)
            Movimiento.IdTipoDocumentoReferencia = F.NuloNumeric(F.KeyValue("IdDocumentoGuiaRemision", xCon))
            Movimiento.IdDocumentoReferencia = xId
            Movimiento.DocumentoReferencia = F.NuloString(TxtNumSer.Text & " - " & TxtNumDoc)
            Movimiento.MesTrabajo = mMesActivo
            Movimiento.AnhoTrabajo = AnoTra
            ' Detalle
            For A = 1 To Fg1.Rows - 1
                Dim MovimientoDet As New AlmacenEntidad.EMovimientoDet
                MovimientoDet.IdItem = F.NuloNumeric(Fg1.TextMatrix(A, 6))
                MovimientoDet.Cantidad = NulosN(Fg1.TextMatrix(A, 4))
                MovimientoDet.CantidadTeorica = NulosN(Fg1.TextMatrix(A, 4))
                ' Se agrega al padre
                Movimiento.LMovimientoDet.Add MovimientoDet
                Set MovimientoDet = Nothing
            Next A
            ' Se graba el movimiento
            Set Movimiento.Conexion = xCon
            If Not Movimiento.Save(CLng(xIdUsuario), F.MachineName) Then Err.Raise &HFFFFFF01, , "No se puedo registrar el movimiento"
        End If
    End If
    
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    xCon.CommitTrans
    
    
    MsgBox "La guia de Remision se grabó con éxito", vbInformation + vbOKOnly + vbOKOnly, xTitulo
    Grabar = True
    CmdAddEntrega.Enabled = False
    Exit Function
    
LaCague:
''Resume
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set RstDet = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Exit Function
End Function

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        
       'MODIFICACION DE GUIAS
        If ButtonMenu.Index = 1 Then
            If RstGui.RecordCount = 0 Then
                MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
                Exit Sub
            End If
            If RstGui("anulado") = -1 Then
                MsgBox "No puede modificar una guia anulada proceda a restaurarla", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            ElseIf RstGui("estado") = -1 Then
                MsgBox "No puede modificar una guia facturada proceda a anular la factura", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                Exit Sub
            Else
                Modificar
            End If
        End If
        
        'RESTAURAR GUIAS
        If ButtonMenu.Index = 2 Then
            If RstGui.RecordCount = 0 Then
                MsgBox "La cuadricula de datos esta vacia", vbInformation, Me.Caption
                Exit Sub
            End If
                If RstGui("anulado") = -1 Then
                    RestaurarGuia
                Else
                    If RstGui("estado") = -1 Then 'si la guia esta facturada no se restaura
                        MsgBox "No puede aplicar esta opcion a una guia facturada ", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
                        Exit Sub
                    ElseIf RstGui("estado") = 0 And RstGui("anulado") = 0 Then
                        MsgBox "No puede aplicar esta opcion a una guia pendiente sin facturar", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
                        Exit Sub
                    End If
                End If
        End If
    
        If ButtonMenu.Index = 3 Then
            Frame7.Left = 1320
            Frame7.Top = 1200
            TxtCliente.Text = ""
            LblIdCli.Caption = ""
            TxtProducto.Text = ""
            LblIdProd2.Caption = ""
            Fg6.Rows = 1
            fg5.ColComboList(1) = "|..."
            fg5.Rows = 1
            fg5.Rows = fg5.Rows + 1
            fg5.Editable = flexEDKbdMouse
            Fg6.Editable = flexEDKbdMouse
            
            fg5.ColWidth(5) = 0
            Fg6.ColWidth(6) = 0
            Fg6.ColWidth(7) = 0
            
            Toolbar1.Enabled = False
            TabOne1.Enabled = False
            Frame7.Visible = True
        End If
    End If
    If ButtonMenu.Parent.Index = 3 Then
        If ButtonMenu.Index = 1 Then Anular
        If ButtonMenu.Index = 2 Then Eliminar
        If ButtonMenu.Index = 3 Then EmitirAnulada
    End If
    
    If ButtonMenu.Parent.Index = 8 Then
       FiltrarGuias (ButtonMenu.Index)
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : FiltrarGuias
'* Tipo             : PROCEDIMIENTO
'* Descripcion      :
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    numindex  |  INTEGER    |  ESPECIFICA EL NUMERO DE INDICE
'* Devuelve         :
'*****************************************************************************************************
Private Sub FiltrarGuias(numindex As Integer)
    Dim nFiltro As String
    
    If numindex = 1 Then 'Facturadas
        nFiltro = "numdocref <> '-' and numordcom<>'ANULADA'"
    ElseIf numindex = 2 Then 'No Facturadas
        nFiltro = "numdocref = '-' and numordcom<>'ANULADA'"
    ElseIf numindex = 3 Then 'Anuladas
        nFiltro = "numordcom='ANULADA'"
    Else 'Todas
        nFiltro = ""
    End If
    RstGui.Filter = nFiltro
End Sub

'*****************************************************************************************************
'* Nombre           : EmitirAnulada
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EMITE UNA GUIA COMO ANULADA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub EmitirAnulada()
    QueHace = 1
    xHorIni = Time
    
    TabOne1.CurrTab = 0
    ActivarEntorno
    
    Fraseldoc.Left = 3315
    Fraseldoc.Top = 2505
    
    TxtNumSer2.Text = ""
    TxtNumDocGen.Text = ""
    Fraseldoc.Visible = True
    TxtFchEmiAnul.SetFocus
End Sub

Sub ActivarEntorno()
    TabOne1.Enabled = Not TabOne1.Enabled
    Toolbar1.Enabled = Not Toolbar1.Enabled
End Sub

'*****************************************************************************************************
'* Nombre           : RestaurarGuia
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : RESTABLECE UNA GUIA ANULADA, PARA ELLO CAMBIA EL ESTADO DE LA GUIA 0
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub RestaurarGuia()
    Dim Rpta As Integer
    
    Rpta = MsgBox("Esta seguro de restaurar la guia Nº " + RellenaNumdoc(RstGui("numser"), RstGui("numdoc")), vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
'        xCon.Execute "UPDATE vta_guia SET vta_guia.Anulado = 0, vta_guia.Estado = 0, vta_guia.idcli = 1,  vta_guia.numordcom= ' ', idpunven = 0" _
'            & " WHERE vta_guia.id =" & RstGui("id") & ""
        
        xCon.Execute "UPDATE vta_guia SET vta_guia.Anulado = 0, vta_guia.idcli = 1,  vta_guia.numordcom= ' ', idpunven = 0" _
            & " WHERE vta_guia.id =" & RstGui("id") & ""
        
        xCon.Execute "DELETE * FROM vta_guiadet WHERE vta_guiadet.idgui =" & RstGui("id") & ""
        RstGui.Requery
        Dg1.Refresh
        MsgBox "La guia se restauro con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    End If
End Sub

Private Sub TxtChofer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtChofer_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdCho_Click
    End If
End Sub

Private Sub TxtCli_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtCli_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdCli_Click
    End If
End Sub

Private Sub TxtCli_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtCli.Text) = "" Then
        TxtIdTipDoc.Text = ""
        LblDescTipDocRef.Caption = ""
            
        TxtNumCotizacion.Text = ""
        LblIdDocRef.Caption = ""
        Fg1.Rows = 1

    End If
    

End Sub

Private Sub TxtCliente_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCli_Click
    End If
End Sub

Private Sub TxtDescTrans_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDescTrans_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdEmpTra_Click
    End If
End Sub

Private Sub TxtDirPunVen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdTipDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusIdTipDocRef_Click
    End If
End Sub

Private Sub TxtIdTipDoc_Validate(Cancel As Boolean)
    ' VALIDA EL TIPO DE DOCUMENTO
    If QueHace = 3 Then Exit Sub
    
'    If NulosN(TxtIdTipDoc.Text) = 0 Then
'        TxtIdTipDoc.Text = ""
'        LblDescTipDocRef.Caption = ""
'
'        TxtNumCotizacion.Text = ""
'        LblIdDocRef.Caption = ""
'        Fg1.Rows = 1
'        Exit Sub
'    End If
    
    Dim xRs1 As New ADODB.Recordset
    
    RST_Busq xRs1, "SELECT * FROM mae_docreferencia WHERE id = " & NulosN(TxtIdTipDoc.Text) & "", xCon
    
    If xRs1.RecordCount <> 0 Then
  
''        '--verificar si cambian de cliente => limpiar campos
        If xRs1("id") <> NulosN(TxtIdTipDoc.Tag) And NulosN(TxtIdTipDoc.Tag) <> 0 Then
            TxtNumCotizacion.Text = ""
            LblIdDocRef.Caption = ""
            Fg1.Rows = 1
        End If
    
        TxtIdTipDoc.Text = xRs1("id")
        TxtIdTipDoc.Tag = xRs1("id")
        LblDescTipDocRef.Caption = Trim(xRs1("descripcion"))
                
        Habilitar_Control_Pedido
        
    End If
    Set xRs1 = Nothing
End Sub

Private Sub TxtMotivo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtMotivo_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdMot_Click
    End If
End Sub

Private Sub TxtMotivo_Validate(Cancel As Boolean)
    ' VALIDA EL MOTIVO DE TRASLADO
    If TxtMotivo.Text <> "" Then
        LblMotivo.Caption = Busca_Codigo(TxtMotivo.Text, "id", "descripcion", "mae_mottra", "N", xCon)
        If LblMotivo.Caption = "" Then
            TxtMotivo.Text = ""
        End If
    End If
End Sub

Private Sub TxtNumBre_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

'Private Sub TxtNumCotizacion_Change()
'    TxtNumOrd.Text = TxtNumCotizacion.Text
'    Fg1.Rows = 1
'End Sub

Private Sub TxtNumCotizacion_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumCotizacion_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If KeyCode = 45 Then
        
        
        JALOPEDIDO = False
        
        
        '--------
        If NulosC(TxtNumCotizacion.Text) = "" Then
            LblIdDocRef.Caption = ""
            Fg1.Rows = 1
        End If
        '---
        
    End If
    If KeyCode = 116 Then
        cmdbuscotizacion_Click
    End If
End Sub

Private Sub TxtNumCotizacion_Validate(Cancel As Boolean)
    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtNumCotizacion.Text) = "" Then
        TxtNumOrd.Text = TxtNumCotizacion.Text
        LblIdDocRef.Caption = ""
        Fg1.Rows = 1
    End If
    
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
'    If NulosC(TxtNumSer.Text) = "" Then
'        MsgBox "No ha especificado el numero de serie del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        Exit Sub
'    End If
'
'    If TxtNumDoc.Text <> "" Then
'        If QueHace <> 1 Then Exit Sub
'        Dim xNumGui As String
'        Dim xBus As String
'        Dim Rst As New ADODB.Recordset
'
'        RST_Busq Rst, "SELECT vta_guia.numser, vta_guia.numdoc From vta_guia Where (((vta_guia.numser) = '" & Format(NulosC(TxtNumSer.Text), "0000") & "') " _
'            & " And ((vta_guia.NumDoc) = '" & Format(NulosC(TxtNumDoc.Text), "0000000000") & "')) ORDER BY vta_guia.numdoc", xCon
'
'        If Rst.RecordCount <> 0 Then
'            TxtNumDoc.Text = ""
'            TxtNumSer_Validate True
'            MsgBox "El numero de guia especificado ya existe, el numero se actualizo a " + Trim(TxtNumSer.Text) + "-" + Trim(TxtNumDoc.Text), vbInformation + vbOKCancel + vbDefaultButton1, xTitulo
'            Exit Sub
'        End If
'
'        TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
'        Set Rst = Nothing
'    End If
    
    Dim idDocumento As Long
    
    If NulosC(TxtNumDoc.Text) = "" Then Exit Sub
    TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
    
    If QueHace = 1 Then idDocumento = 0 Else idDocumento = F.NuloNumeric(RstGui("id"))
    If F.ExisteDocumento("vta_guia", "'" & F.NuloString(TxtNumDoc.Text) & "'", xCon, , "'" & F.NuloString(TxtNumSer.Text) & "'", , , , idDocumento, "id") Then
        MsgBox "El documento ingresado ya existe" & vbCr & "Corrija el numero de documento", vbInformation, xTitulo
        TxtNumDoc.Text = ""
        TxtNumDoc.SetFocus
        Exit Sub
    End If
End Sub

Private Sub TxtNumDocGen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumDocGen_Validate(Cancel As Boolean)
    If NulosC(TxtNumSer2.Text) = "" Then
        MsgBox "No ha especificado el numero de serie del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    If TxtNumDocGen.Text <> "" Then
        If QueHace <> 1 Then Exit Sub
        Dim xNumGui As String
        Dim xBus As String
        Dim Rst As New ADODB.Recordset
        
        RST_Busq Rst, "SELECT vta_guia.numser, vta_guia.numdoc From vta_guia Where (((vta_guia.numser) = '" & Format(NulosC(TxtNumSer2.Text), "0000") & "') " _
            & " And ((vta_guia.NumDoc) = '" & Format(NulosC(TxtNumDocGen.Text), "0000000000") & "')) ORDER BY vta_guia.numdoc", xCon
        
        If Rst.RecordCount <> 0 Then
            TxtNumDocGen.Text = ""
            TxtNumSer2_Validate True
            MsgBox "El numero de guia especificado ya existe, el numero se actualizo a " + Trim(TxtNumSer2.Text) + "-" + Trim(TxtNumDocGen.Text), vbInformation + vbOKCancel + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        
        TxtNumDocGen.Text = Format(TxtNumDocGen.Text, "0000000000")
        Set Rst = Nothing
    End If
End Sub

Private Sub TxtNumLote_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumOrd_Change()
    TxtNumCotizacion.Text = TxtNumOrd.Text
End Sub

Private Sub TxtNumOrd_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumPlaCar_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumPlaCar_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdPlaCar_Click
    End If
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumSer_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 116 Then
'        CmdBusNumSer_Click
'    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
'    If QueHace <> 1 Then Exit Sub
'    Dim Rstnumgui As New ADODB.Recordset
'    If NulosC(TxtNumSer.Text) <> "" Then
'        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
'        RST_Busq Rstnumgui, "SELECT vta_guia.numser, vta_guia.numdoc From vta_guia Where (((vta_guia.numser) = '" & Format(TxtNumSer.Text, "0000") & "'))" _
'            & " ORDER BY vta_guia.numdoc", xCon
'
'        If Rstnumgui.RecordCount = 0 Then
'            TxtNumDoc.Text = "0000000001"
'        Else
'            Rstnumgui.MoveLast
'            TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
'            TxtNumDoc.Text = Format((Val(Rstnumgui("numdoc")) + 1), "0000000000")
'        End If
'    End If
'    Set Rstnumgui = Nothing

    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtNumSer.Text) <> "" Then
        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
        TxtNumDoc.Text = F.HallaNumeroDocumento("vta_guia", "'" & NulosC(TxtNumSer.Text) & "'", "numser", xCon)
        If NulosC(TxtNumDoc.Text) = "" Then TxtNumSer.Text = ""
    End If
End Sub

Private Sub TxtNumSer2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumSer2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusNumSer2_Click
    End If
End Sub

Private Sub CmdBusNumSer2_Click()
    ' EJECUTA LA BUSQUEDA DE UN NUMERO DE SERIE
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Codigo":         xCampos(1, 1) = "iddoc":       xCampos(0, 2) = "1500":    xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":    xCampos(0, 1) = "descripcion": xCampos(1, 2) = "1500":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Serie":          xCampos(2, 1) = "numser":      xCampos(2, 2) = "1500":    xCampos(2, 3) = "C"
    xCampos(3, 0) = "Nro Documento":  xCampos(3, 1) = "numdoc":      xCampos(3, 2) = "1500":    xCampos(3, 3) = "C"
    
    xform.SQLCad = "SELECT mae_documento.descripcion, mae_series.iddoc, mae_series.numser, mae_series.numdoc " & _
                   " FROM mae_documento INNER JOIN mae_series ON mae_documento.id = mae_series.iddoc where iddoc = 9"

    xform.titulo = "Buscando Guias"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numser"
    xform.CampoBusca = "numser"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumSer2.Text = Format(NulosN(xRs("numser")), "0000")
            TxtNumDocGen.Text = HallaNumGuia(NulosC(TxtNumSer2.Text), xCon)
        End If
        TxtNumDocGen.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub TxtNumSer2_Validate(Cancel As Boolean)
    
    If QueHace = 3 Then Exit Sub
    Dim Rstnumgui As New ADODB.Recordset
    If NulosC(TxtNumSer2.Text) <> "" Then
        TxtNumSer2.Text = Format(TxtNumSer2.Text, "0000")
        RST_Busq Rstnumgui, "SELECT vta_guia.numser, vta_guia.numdoc From vta_guia Where (((vta_guia.numser) = '" & Format(TxtNumSer2.Text, "0000") & "'))" _
            & " ORDER BY vta_guia.numdoc", xCon

        If Rstnumgui.RecordCount = 0 Then
            TxtNumDocGen.Text = "0000000001"
        Else
            Rstnumgui.MoveLast
            TxtNumSer2.Text = Format(TxtNumSer2.Text, "0000")
            TxtNumDocGen.Text = Format((Val(Rstnumgui("numdoc")) + 1), "0000000000")
        End If
    End If
    Set Rstnumgui = Nothing
End Sub

Private Sub TxtPunVen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If NulosC(TxtPunVen.Text) = "" Then
            LblIdPunVen.Caption = ""
        End If
        SendKeys vbTab
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Nuevo()
    QueHace = 1
    Bloquea
    Blanquea
    Label1.Caption = "Agregando Guia de Remision"
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    
    Fg1.Rows = 1
    xHorIni = Time
    JALOPEDIDO = False
    VAR_IDPEDIDO = 0
    VAR_FECHAPEDIDO = ""
    
    TxtTipItem.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    QueHace = 2
    Bloquea
    Blanquea
    Label1.Caption = "Modificando Guia de Remision"
    TabOne1.CurrTab = 1
    ActivaTool
    TabOne1.TabEnabled(0) = False
    MuestraSegundoTab
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    xHorIni = Time
    If NulosN(TxtIdTipDoc.Text) = 5 Then
        CmdAddEntrega.Enabled = True
        CmdAddItem.Enabled = False
        CmdDelItem.Enabled = False
    End If
    TxtFchEmi.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELA EL PROCESO DE ADICIONAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    Dim X As Integer
    Label1.Caption = "Detalle Guia de Remision"
    ActivaTool
    Fg1.ColComboList(1) = ""
    Fg1.SelectionMode = flexSelectionByRow
    
    TabOne1.TabEnabled(0) = True
    Bloquea
    TabOne1.CurrTab = 0
    QueHace = 3
    
    'Colocamos en el campo idest 2  de la tabla cotizacion  que indica Aprobado
    If fgdocsproc.Rows - 1 > 0 Then
        If swguiafact = 0 Then
            For X = 1 To fgdocsproc.Rows - 1
                xCon.Execute " UPDATE vta_cotizacion SET vta_cotizacion.idEst = 2 WHERE vta_cotizacion.id = " & Val(fgdocsproc.TextMatrix(X, 1)) & ""
            Next
            fgdocsproc.Rows = 1
        End If
    End If
    CmdAddEntrega.Enabled = False
    CmdDatosAdicion.Enabled = False
    CmdAddItem.Enabled = False
    CmdDelItem.Enabled = False
    
    swguiafact = 0
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LA BARRA DE HERRAMIENTAS
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivaTool()
    Toolbar1.Buttons(1).Enabled = Not Toolbar1.Buttons(1).Enabled
    Toolbar1.Buttons(2).Enabled = Not Toolbar1.Buttons(2).Enabled
    Toolbar1.Buttons(3).Enabled = Not Toolbar1.Buttons(3).Enabled
    
    Toolbar1.Buttons(5).Enabled = Not Toolbar1.Buttons(5).Enabled
    Toolbar1.Buttons(6).Enabled = Not Toolbar1.Buttons(6).Enabled
    
    Toolbar1.Buttons(8).Enabled = Not Toolbar1.Buttons(8).Enabled
    Toolbar1.Buttons(9).Enabled = Not Toolbar1.Buttons(9).Enabled
    Toolbar1.Buttons(10).Enabled = Not Toolbar1.Buttons(10).Enabled
    Toolbar1.Buttons(11).Enabled = Not Toolbar1.Buttons(11).Enabled
    
    Toolbar1.Buttons(13).Enabled = Not Toolbar1.Buttons(13).Enabled
    Toolbar1.Buttons(15).Enabled = Not Toolbar1.Buttons(15).Enabled
End Sub

Sub AddItem()
    Fg1.Rows = Fg1.Rows + 1
    With Fg1
        .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, 1
    End With
    Fg1.SetFocus
    Fg1_CellButtonClick Fg1.Row - 1, 1
End Sub

Sub DelItem()
    If Fg1.Row < 1 Then
        MsgBox "Seleccione el reigistro a eliminar", vbExclamation, xTitulo
        Exit Sub
    End If
    If Fg1.Rows - 1 > 0 Then
        If Fg1.Rows - 1 = 1 Then
            Fg1.Rows = 1
            Exit Sub
        Else
            Fg1.RemoveItem Fg1.Row
        End If
    End If
    
End Sub

Private Sub TxtPunVen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        LblIdPunVen.Caption = ""
        TxtPunVen.Text = ""
    End If
    
    If KeyCode = 116 Then
        CmdBusPunVen_Click
    End If
End Sub

Private Sub TxtTipItem_KeyPress(KeyAscii As Integer)
    Dim RstTmp As New ADODB.Recordset
    If KeyAscii = 13 Then
        If NulosC(TxtTipItem.Text) <> "" Then
            Set RstTmp = BuscaConCriterio("SELECT * FROM mae_tipoproducto WHERE id = " & Val(TxtTipItem.Text) & "", xCon)
            If RstTmp.RecordCount <> 0 Then
                LblTipoItem.Caption = RstTmp("descripcion")
            Else
                TxtTipItem.Text = ""
                LblTipoItem.Caption = ""
            End If
        End If
        TxtFchEmi.SetFocus
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
    Set RstTmp = Nothing
End Sub

Private Sub TxtTipItem_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipItem_Click
    End If
End Sub

Private Sub TxtTipItem_Validate(Cancel As Boolean)
    ' EJECUTA LA BUSQUEDA DEL TIPO DE ITEM
    If QueHace = 3 Then Exit Sub
    If NulosN(TxtTipItem.Text) = 0 Then LblTipoItem.Caption = "": Exit Sub
    
    Dim xRs As New ADODB.Recordset
    
    RST_Busq xRs, "SELECT mae_tipoproducto.* FROM mae_tipoproducto WHERE id = " & TxtTipItem.Text & " ", xCon
    
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            LblTipoItem = xRs("descripcion")
        End If
    End If
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : Imprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME FISICAMENTE UN GUIA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Imprimir()
    Dim RsPDoc As New ADODB.Recordset
    Dim RsPCab As New ADODB.Recordset
    Dim RsPDet As New ADODB.Recordset
    Dim xRsDoc As New ADODB.Recordset
    Dim xRsDet As New ADODB.Recordset
    Dim F As New SistemaLogica.Funciones

    ' CARAGAMOS LOS DATOS DE LA GUIA
    RST_Busq xRsDoc, "SELECT Format(vta_guia.numser,'0000')+'-'+Format(vta_guia.numdoc,'0000000000') AS numguia, VTA_Guia.*, MAE_Cliente.nombre, MAE_Cliente.dir, " _
        & " VTA_PuntoVenta.descripcion AS despunven, VTA_PuntoVenta.dir AS Direccion, mae_emptra.nombre AS desemptra, mae_emptra.numruc, " _
        & " mae_mottra.descripcion AS descmotgui, MAE_Cliente.numruc AS RucCli, mae_chofer.numbre, mae_vehiculo.marca AS marcacar, mae_vehiculo.numpla, " _
        & " Format([vta_ventas].[numser],'0000')+'-'+Format([vta_ventas].[numdoc],'0000000000') AS numdocref, UCase([pla_empleados]![apepat])+' '+ UCase([pla_empleados]![apemat])+', '+ " _
        & " [pla_empleados]![nom] AS apenomcho, mae_emptra.direccion AS dirori" _
        & " FROM pla_empleados RIGHT JOIN (vta_ventas RIGHT JOIN ((((((VTA_Guia LEFT JOIN MAE_Cliente ON VTA_Guia.idcli = MAE_Cliente.id) LEFT JOIN VTA_PuntoVenta " _
        & " ON VTA_Guia.idpunven = VTA_PuntoVenta.id) LEFT JOIN mae_emptra ON VTA_Guia.idemptra = mae_emptra.id) LEFT JOIN mae_mottra ON VTA_Guia.idmottra = mae_mottra.id) " _
        & " LEFT JOIN mae_chofer ON VTA_Guia.idcho = mae_chofer.id) LEFT JOIN mae_vehiculo ON VTA_Guia.idveh = mae_vehiculo.id) ON vta_ventas.id = VTA_Guia.iddocven) " _
        & " ON pla_empleados.id = mae_chofer.idper Where (((VTA_Guia.id) = " & RstGui("id") & ")) ORDER BY VTA_Guia.fecgiro", xCon

    ' CARGAMOS EL DETALLE DE LA GUIA
    RST_Busq xRsDet, "SELECT DISTINCT vta_guiadet.*, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev FROM mae_unidades RIGHT JOIN ((alm_inventario " _
        & " RIGHT JOIN vta_guiadet ON alm_inventario.id = vta_guiadet.iditem) LEFT JOIN mae_productoscen ON vta_guiadet.iditem = mae_productoscen.iditem) " _
        & " ON mae_unidades.id = alm_inventario.idunimed WHERE (((vta_guiadet.idgui)=" & RstGui("id") & "))", xCon

    ' BUSCAMOS LA PLANTILLA DE IMPRESION DE LA GUIA
    RST_Busq RsPDoc, "SELECT * FROM var_plantilladoc WHERE tipdoc = " & xRsDoc("tipdoc") & " ", xCon

    If RsPDoc.RecordCount = 0 Then
        MsgBox "No se ha definido la plantilla de impresion para este tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set xRsDoc = Nothing
        Set xRsDet = Nothing
        Set RsPDoc = Nothing
        Exit Sub
    End If

    ' CARGAMOS LA PLANTILLA DE IMPRESION DE LA GUIA
    RST_Busq RsPCab, "SELECT var_plantillacab.* From var_plantillacab Where (((var_plantillacab.idplan) = " & RsPDoc("id") & ")) " _
        & " ORDER BY var_plantillacab.numord", xCon
    
    RST_Busq RsPDet, "SELECT var_plantilladet.* FROM var_plantilladet WHERE idplan = " & RsPDoc("id") & " ORDER BY var_plantilladet.numord", xCon

    ' CONFIGURAMOS LA IMPRESION
    'Printer.FontSize = RsPDoc("tamañoletra")
    'Printer.FontBold = True
    RsPDoc.MoveFirst
    Printer.Font = NulosC(RsPDoc("tipoletra"))
    Printer.ScaleMode = 6
    
    Dim xCam, xCam2, xFor As String

    ' imprime cabezera
    Do While RsPCab.EOF = False
        xCam = NulosC(RsPCab("campo"))
        xFor = NulosC(RsPCab("formato"))
        Printer.CurrentX = RsPCab("posx")
        Printer.CurrentY = RsPCab("posy")
        xCam2 = ""
        If UCase(xCam) = "NUMORDCOM" Then xCam2 = "Nº de Orden : "
        If UCase(xCam) = "FCHEMIORD" Then xCam2 = "Fch. Emi. : "
        If UCase(xCam) = "FCHENTORD" Then xCam2 = "Fch. Ent. : "
        If UCase(xCam) = "NUMLOTE" Then xCam2 = "Nº de Lote : "
        If UCase(xCam) = "FCHPRO" Then xCam2 = "Fch. Prod. : "
        If UCase(xCam) = "FCHVEN" Then xCam2 = "Fch. Ven. : "
        
        If NulosC(xFor) = "" Then
            If UCase(xCam) = "NUMLOTE" Or UCase(xCam) = "FCHPRO" Or UCase(xCam) = "FCHVEN" Then
                If xRsDoc(xCam) <> "" Then
                    Printer.Print xCam2 + NulosC(xRsDoc(xCam))
                End If
            Else
                
            Printer.FontSize = NulosN(RsPCab("tamanho"))
            Printer.FontBold = NulosN(RsPCab("negrita"))
            F.PrintText Printer, xCam2 + NulosC(xRsDoc(xCam)), NulosN(RsPCab("alineacion"))
            End If
        Else
            If UCase(xCam) = "NUMLOTE" Or UCase(xCam) = "FCHPRO" Or UCase(xCam) = "FCHVEN" Then
                If xRsDoc(xCam) <> "" Then
                    Printer.FontSize = NulosN(RsPCab("tamanho"))
                    Printer.FontBold = NulosN(RsPCab("negrita"))
                    F.PrintText Printer, xCam2 + Format((NulosC(xRsDoc(xCam))), xFor), NulosN(RsPCab("alineacion"))
                End If
            Else
                
            Printer.FontSize = NulosN(RsPCab("tamanho"))
            Printer.FontBold = NulosN(RsPCab("negrita"))
            F.PrintText Printer, xCam2 + Format((NulosC(xRsDoc(xCam))), xFor), NulosN(RsPCab("alineacion"))
            End If
        End If
        RsPCab.MoveNext
    Loop
   
    'imprime detalle
    Dim Fila As Integer
    
    Fila = RsPDet("posy")
    xRsDet.MoveFirst
    Dim xRs2 As New ADODB.Recordset
    Dim xCad As String
    
    Do While xRsDet.EOF = False
        RsPDet.MoveFirst
        Do While RsPDet.EOF = False
            xCam = NulosC(RsPDet("campo"))
            xFor = NulosC(RsPDet("formato"))
            Printer.CurrentX = RsPDet("posx")
            Printer.CurrentY = Fila
            If UCase(xCam) = "DESCRIPCION" Then
                xCad = "SELECT * FROM mae_productoscen WHERE iditem = " & xRsDet("iditem") & ""
                Set xRs2 = BuscaConCriterio(xCad, xCon)
                If xRs2.RecordCount <> 0 Then
                    Printer.FontSize = NulosN(RsPDet("tamanho"))
                    Printer.FontBold = NulosN(RsPDet("negrita"))
                    F.PrintText Printer, Format((NulosC(xRsDet(xCam))), xFor) + " - " + NulosC(xRs2("codewong")), NulosN(RsPDet("alineacion"))
                Else
                    Printer.FontSize = NulosN(RsPDet("tamanho"))
                    Printer.FontBold = NulosN(RsPDet("negrita"))
                    F.PrintText Printer, Format((NulosC(xRsDet(xCam))), xFor), NulosN(RsPDet("alineacion"))
                End If
            Else
                Printer.FontSize = NulosN(RsPDet("tamanho"))
                Printer.FontBold = NulosN(RsPDet("negrita"))
                F.PrintText Printer, Format((NulosC(xRsDet(xCam))), xFor), NulosN(RsPDet("alineacion"))
            End If
            RsPDet.MoveNext
        Loop
        Fila = Fila + 4
        xRsDet.MoveNext
    Loop
    
    ' MANDA A IMPRIMIR EL DOCUMENTO
    Printer.EndDoc
    
    MsgBox "El documento se imprimio con exito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Set RsPDoc = Nothing
    Set RsPCab = Nothing
    Set RsPDet = Nothing
    Set xRsDoc = Nothing
    Set xRsDet = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : CargarGuia
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS DATOS DE UNA GUIA PARA EFECTUAR UNA BUSQUEDA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargarGuia()
    Dim xfrm As New eps_librerias.FormSeleccion
    Dim xCampos(4, 5) As String
    Dim xRs As New ADODB.Recordset
    Dim xRs1 As New ADODB.Recordset
    
    xCampos(0, 0) = "Nº Documento":    xCampos(0, 1) = "nrodoc":        xCampos(0, 2) = "1500":   xCampos(0, 3) = "C":     xCampos(0, 4) = "S"
    xCampos(1, 0) = "Fch. Giro":       xCampos(1, 1) = "fecgiro":       xCampos(1, 2) = "1000":   xCampos(1, 3) = "C":     xCampos(1, 4) = "N"
    xCampos(2, 0) = "Cliente":         xCampos(2, 1) = "nombre":        xCampos(2, 2) = "2500":   xCampos(2, 3) = "C":     xCampos(2, 4) = "N"
    xCampos(3, 0) = "Motivo":          xCampos(3, 1) = "descripcion":   xCampos(3, 2) = "2000":   xCampos(3, 3) = "C":     xCampos(3, 4) = "N"

    xfrm.SQLCad = "SELECT vta_guia.id, vta_guia.fecgiro, [vta_guia]![numser]+'-'+[vta_guia]![numdoc] AS NroDoc, mae_cliente.numruc, mae_cliente.nombre, " _
        & " mae_mottra.descripcion, vta_guia.idcli, mae_documento.abrev, vta_guia.iddocven FROM mae_mottra RIGHT JOIN ((mae_cliente RIGHT JOIN vta_guia " _
        & " ON mae_cliente.id = vta_guia.idcli) LEFT JOIN mae_documento ON vta_guia.tipdoc = mae_documento.id) ON mae_mottra.id = vta_guia.idmottra " _
        & " Where (((vta_guia.idcli) = " & NulosN(LblIdcli2.Caption) & ") And ((vta_guia.Anulado) = 0) And ((vta_guia.iddocven) = 0) And ((vta_guia.iddocven) = 0 Or (vta_guia.iddocven) Is Null)) " _
        & " ORDER BY [vta_guia]![numser]+'-'+[vta_guia]![numdoc] DESC"
        
    xfrm.titulo = "Buscando Guias del Guias"
    
    Set xfrm.Coneccion = xCon
    Set xRs = xfrm.Seleccionar(xCampos)
    
    If xRs.State = 1 Then
        If xRs.RecordCount = 0 Then
            Set xRs = Nothing
            Exit Sub
        End If
        Dim xCadWhere As String
        Dim A As Integer
        Dim Rst As New ADODB.Recordset
        
        fg5.Rows = 1
        xRs.MoveFirst
        fg5.Rows = 1
        
        ' CARGAMOS LOS DOCUMENTOS ADJUNTOS Y LO MOSTRAMOS EN LA LISTA DE "DOCUMENTOS ADJUNTOS"
        For A = 1 To xRs.RecordCount
            fg5.Rows = fg5.Rows + 1
            fg5.TextMatrix(A, 1) = xRs("nrodoc")
            fg5.TextMatrix(A, 2) = xRs("fecgiro")
            fg5.TextMatrix(A, 3) = xRs("nombre")
            fg5.TextMatrix(A, 4) = ""
            fg5.TextMatrix(A, 5) = xRs("id")
            xRs.MoveNext
            If xRs.EOF = True Then Exit For
        Next A
        TxtProducto.Text = ""
        LblIdProd2.Caption = ""
        Fg6.Rows = 1
    End If
    
    Set xfrm = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS REGISTRAS DE LA TABLA vta_guia EN EL RECORDSET RstGui, ESTOS DATOS SE
'*                    VISUALIZARAN EN LA PESTAÑA CONSULTA DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarGrid()
    Dim nSQL  As String
    Dim Rpta As Integer
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    LblPeriodo2.Caption = LblMes.Caption

    TDB_FiltroLimpiar Dg1
    Set RstGui = Nothing
    
    ' bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
    
    nSQL = "SELECT Format(vta_guia.numser,'0000')+'-'+Format(vta_guia.numdoc,'0000000000') AS numguia, VTA_Guia.*, IIf([vta_guia].[anulado]=-1,'ANULADO',[mae_cliente].[nombre]) AS nombre, MAE_Cliente.Dir, VTA_PuntoVenta.descripcion AS despunven, IIf([vta_guia].[Anulado]=0,'','Anulado') AS Anulado2, VTA_PuntoVenta.Dir AS Direccion, mae_emptra.nombre AS desemptra, mae_emptra.numruc, mae_mottra.descripcion AS descmotgui, MAE_Cliente.numruc AS RucCli, UCase([pla_empleados].[apepat])+' '+UCase([pla_empleados].[apemat])+', '+[pla_empleados].[nom] AS apenomcho, mae_chofer.numbre, mae_vehiculo.marca AS marcacar, mae_vehiculo.numpla, Format([vta_ventas].[numser],'0000')+'-'+Format([vta_ventas].[numdoc],'0000000000') AS numdocref, VTA_Guia.fecgiro & '' AS fecgiro1, alm_almacenes.descripcion AS desalm " _
        + vbCr + "FROM (pla_empleados RIGHT JOIN (vta_ventas RIGHT JOIN ((((((VTA_Guia LEFT JOIN MAE_Cliente ON VTA_Guia.idcli = MAE_Cliente.id) LEFT JOIN VTA_PuntoVenta ON VTA_Guia.idpunven = VTA_PuntoVenta.id) LEFT JOIN mae_emptra ON VTA_Guia.idemptra = mae_emptra.id) LEFT JOIN mae_mottra ON VTA_Guia.idmottra = mae_mottra.id) LEFT JOIN mae_chofer ON VTA_Guia.idcho = mae_chofer.id) LEFT JOIN mae_vehiculo ON VTA_Guia.idveh = mae_vehiculo.id) ON vta_ventas.id = VTA_Guia.iddocven) ON pla_empleados.id = mae_chofer.idper) LEFT JOIN alm_almacenes ON VTA_Guia.idalm = alm_almacenes.id " _
        + vbCr + "WHERE (((Year([VTA_Guia].[fecgiro])) = " & AnoTra & ") And ((Month([VTA_Guia].[fecgiro])) = " & mMesActivo & ")) " _
        + vbCr + "ORDER BY VTA_Guia.numser, VTA_Guia.numdoc DESC"
        
    '--cargando datos
    Me.MousePointer = vbHourglass
    RST_Busq RstGui, nSQL, xCon

    Set Dg1.DataSource = RstGui
    
    Me.MousePointer = vbDefault
    
    
    TabOne1.CurrTab = 0
    If RstGui.State = 1 Then
        If RstGui.RecordCount = 0 Then
            Rpta = MsgBox("No se ha registrado ninguna Guia, ¿Desea agregar una ahora?", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
            If Rpta = vbYes Then
                Nuevo
            Else
            End If
        End If
    End If
End Sub


Private Sub Habilitar_Control_Pedido()
    '09/04/11
    '-Permitira habilitar/ deshabilitar los controles relacionados a la seleccion del pedido
    If NulosN(TxtIdTipDoc.Text) = 5 Then
            lblcotizacion.Visible = True
            TxtNumCotizacion.Visible = True
            cmdbuscotizacion.Visible = True
            CmdDatosAdicion.Enabled = True
            TxtNumOrd.Enabled = False
            TxtFchEmiPed.Enabled = False
            TxtFchEnt.Enabled = False
            CmdAddEntrega.Enabled = True
            CmdAddItem.Enabled = False
            CmdDelItem.Enabled = False
        Else
            lblcotizacion.Visible = False
            TxtNumCotizacion.Visible = False
            cmdbuscotizacion.Visible = False
            CmdDatosAdicion.Enabled = True
            TxtNumOrd.Enabled = True
            TxtFchEmiPed.Enabled = True
            TxtFchEnt.Enabled = True
            CmdAddEntrega.Enabled = False
            CmdAddItem.Enabled = True
            CmdDelItem.Enabled = True
        End If
End Sub
