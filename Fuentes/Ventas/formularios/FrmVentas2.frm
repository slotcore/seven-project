VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmVentas2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ventas - Ingreso de Ventas"
   ClientHeight    =   7845
   ClientLeft      =   45
   ClientTop       =   225
   ClientWidth     =   11925
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11925
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Command3"
      Height          =   390
      Left            =   6690
      TabIndex        =   137
      Top             =   345
      Visible         =   0   'False
      Width           =   2250
   End
   Begin VB.CommandButton CmdApertura 
      Caption         =   "&Apertura"
      Height          =   315
      Left            =   10560
      TabIndex        =   136
      Top             =   390
      Visible         =   0   'False
      Width           =   1320
   End
   Begin VB.Frame FrameDetalleItem 
      BorderStyle     =   0  'None
      Caption         =   "2"
      Height          =   4215
      Left            =   4440
      TabIndex        =   128
      Top             =   7860
      Visible         =   0   'False
      Width           =   9645
      Begin VSFlex7Ctl.VSFlexGrid Fg3 
         Height          =   3165
         Left            =   75
         TabIndex        =   129
         Top             =   450
         Width           =   9495
         _cx             =   16748
         _cy             =   5583
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
         GridColor       =   16777215
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
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmVentas2.frx":0000
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
      Begin VB.Frame Frame12 
         Height          =   630
         Left            =   75
         TabIndex        =   131
         Top             =   3540
         Width           =   9510
         Begin VB.CommandButton CmdDetalleCargarDefe 
            Caption         =   "&Cargar Descripcion"
            Height          =   390
            Left            =   4845
            TabIndex        =   133
            Top             =   165
            Width           =   1785
         End
         Begin VB.CommandButton CmdDetalleAceptar 
            Caption         =   "&Aceptar"
            Height          =   390
            Left            =   6660
            TabIndex        =   132
            Top             =   165
            Width           =   1785
         End
         Begin VB.Label LblNumLineas 
            Alignment       =   2  'Center
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNumLineas"
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
            Left            =   1890
            TabIndex        =   135
            Top             =   225
            Width           =   840
         End
         Begin VB.Label Label14 
            Caption         =   "Nº de Lineas Posibles"
            Height          =   210
            Left            =   150
            TabIndex        =   134
            Top             =   255
            Width           =   1560
         End
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Descripción Detallada del Item"
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
         TabIndex        =   130
         Top             =   105
         Width           =   2625
      End
      Begin VB.Shape Shape5 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   30
         Top             =   45
         Width           =   9555
      End
      Begin VB.Line Line11 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   9660
         Y1              =   4200
         Y2              =   4200
      End
      Begin VB.Line Line12 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   15
         Y2              =   4185
      End
      Begin VB.Line Line12 
         BorderColor     =   &H80000003&
         BorderWidth     =   2
         Index           =   0
         X1              =   9630
         X2              =   9630
         Y1              =   15
         Y2              =   4185
      End
      Begin VB.Line Line11 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -15
         X2              =   9630
         Y1              =   15
         Y2              =   15
      End
   End
   Begin VB.Frame Frame8 
      BorderStyle     =   0  'None
      Caption         =   "Frame8"
      Height          =   2025
      Left            =   12165
      TabIndex        =   91
      Top             =   5535
      Visible         =   0   'False
      Width           =   6705
      Begin VB.TextBox TxtNewSaldo2 
         Alignment       =   1  'Right Justify
         Height          =   300
         Left            =   1320
         TabIndex        =   99
         Text            =   "TxtNewSaldo2"
         Top             =   1515
         Width           =   1395
      End
      Begin VB.TextBox TxtSaldo2 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1320
         TabIndex        =   97
         Text            =   "TxtSaldo2"
         Top             =   1200
         Width           =   1395
      End
      Begin VB.TextBox TxtCliente2 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1320
         TabIndex        =   95
         Text            =   "TxtCliente2"
         Top             =   780
         Width           =   5280
      End
      Begin VB.TextBox TxtNumDoc2 
         BackColor       =   &H00E0E0E0&
         Height          =   300
         Left            =   1320
         TabIndex        =   93
         Text            =   "TxtNumDoc2"
         Top             =   465
         Width           =   2055
      End
      Begin VB.Frame Frame9 
         Height          =   870
         Left            =   3240
         TabIndex        =   101
         Top             =   1050
         Width           =   3375
         Begin VB.CommandButton Command2 
            Height          =   630
            Left            =   1710
            Picture         =   "FrmVentas2.frx":003D
            Style           =   1  'Graphical
            TabIndex        =   103
            Top             =   180
            Width           =   750
         End
         Begin VB.CommandButton Command1 
            Height          =   630
            Left            =   930
            Picture         =   "FrmVentas2.frx":0347
            Style           =   1  'Graphical
            TabIndex        =   102
            Top             =   180
            Width           =   750
         End
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nuevo Saldo"
         Height          =   195
         Index           =   3
         Left            =   165
         TabIndex        =   100
         Top             =   1560
         Width           =   930
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Saldo"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   98
         Top             =   1245
         Width           =   405
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Cliente"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   96
         Top             =   825
         Width           =   480
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   94
         Top             =   510
         Width           =   1050
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Actualizar Saldo del Documento"
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
         Left            =   225
         TabIndex        =   92
         Top             =   90
         Width           =   2730
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   300
         Left            =   30
         Top             =   45
         Width           =   6615
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   1995
      End
      Begin VB.Line Line7 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   6690
         X2              =   6690
         Y1              =   15
         Y2              =   2010
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   6690
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line6 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   6675
         Y1              =   2010
         Y2              =   2010
      End
   End
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
      Height          =   2190
      Left            =   12300
      TabIndex        =   59
      Top             =   660
      Visible         =   0   'False
      Width           =   5565
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchEmiAnul 
         Height          =   300
         Left            =   1425
         TabIndex        =   142
         Top             =   1125
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
      Begin VB.TextBox TxtNumDocGen 
         Height          =   300
         Left            =   1425
         MaxLength       =   10
         TabIndex        =   144
         Text            =   "TxtNumDocGen"
         Top             =   1755
         Width           =   1335
      End
      Begin VB.Frame Frame7 
         Height          =   930
         Left            =   3150
         TabIndex        =   89
         Top             =   1140
         Width           =   2280
         Begin VB.CommandButton cmdsalirseldoc 
            Height          =   510
            Left            =   1140
            Picture         =   "FrmVentas2.frx":0651
            Style           =   1  'Graphical
            TabIndex        =   146
            Top             =   240
            Width           =   1050
         End
         Begin VB.CommandButton cmdokseldoc 
            Height          =   510
            Left            =   90
            Picture         =   "FrmVentas2.frx":24A3
            Style           =   1  'Graphical
            TabIndex        =   145
            Top             =   240
            Width           =   1050
         End
      End
      Begin VB.CommandButton CmdBusAlm2 
         Height          =   240
         Left            =   2085
         Picture         =   "FrmVentas2.frx":4829
         Style           =   1  'Graphical
         TabIndex        =   138
         Top             =   480
         Width           =   240
      End
      Begin VB.TextBox TxtIdAlm2 
         Height          =   300
         Left            =   1440
         MaxLength       =   2
         TabIndex        =   139
         Text            =   "TxtIdAlm2"
         Top             =   450
         Width           =   915
      End
      Begin VB.CommandButton CmdBusTipDoc2 
         Height          =   240
         Left            =   2085
         Picture         =   "FrmVentas2.frx":495B
         Style           =   1  'Graphical
         TabIndex        =   147
         Top             =   810
         Width           =   240
      End
      Begin VB.TextBox TxtTipDoc2 
         Height          =   300
         Left            =   1440
         MaxLength       =   3
         TabIndex        =   140
         Text            =   "TxtTipDoc2"
         Top             =   780
         Width           =   915
      End
      Begin VB.CommandButton CmdBusNumSer2 
         Height          =   240
         Left            =   2085
         Picture         =   "FrmVentas2.frx":4A8D
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   1470
         Width           =   240
      End
      Begin VB.TextBox TxtNumSer2 
         Height          =   300
         Left            =   1440
         MaxLength       =   4
         TabIndex        =   143
         Text            =   "TxtNumSer2"
         Top             =   1440
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Almacén"
         Height          =   195
         Index           =   2
         Left            =   165
         TabIndex        =   90
         Top             =   510
         Width           =   615
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Documento"
         Height          =   195
         Index           =   1
         Left            =   165
         TabIndex        =   88
         Top             =   1170
         Width           =   1185
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Nº Documento"
         Height          =   195
         Left            =   165
         TabIndex        =   70
         Top             =   1785
         Width           =   1050
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   -30
         X2              =   7380
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Nº Serie"
         Height          =   195
         Left            =   165
         TabIndex        =   69
         Top             =   1470
         Width           =   585
      End
      Begin VB.Label Label6 
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
         TabIndex        =   68
         Top             =   105
         Width           =   2880
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   5550
         X2              =   5550
         Y1              =   0
         Y2              =   3360
      End
      Begin VB.Line Line9 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   2235
      End
      Begin VB.Line Line8 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   5550
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Tipo Documento"
         Height          =   195
         Index           =   0
         Left            =   165
         TabIndex        =   67
         Top             =   840
         Width           =   1185
      End
      Begin VB.Label LblAlmacen2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblAlmacen2"
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
         Left            =   2370
         TabIndex        =   141
         Top             =   450
         Width           =   3075
      End
      Begin VB.Label LblNomDoc2 
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LblNomDoc2"
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
         Left            =   2370
         TabIndex        =   148
         Top             =   780
         Width           =   3075
      End
      Begin VB.Shape Shape4 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Left            =   45
         Top             =   45
         Width           =   5475
      End
   End
   Begin VB.Frame Fradocsproc 
      BorderStyle     =   0  'None
      Caption         =   "2"
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
      Height          =   3285
      Left            =   12540
      TabIndex        =   23
      Top             =   1980
      Visible         =   0   'False
      Width           =   3705
      Begin VB.CommandButton cmdEliminarOKdocsproc 
         Height          =   630
         Left            =   1380
         Picture         =   "FrmVentas2.frx":4BBF
         Style           =   1  'Graphical
         TabIndex        =   27
         Top             =   2535
         Width           =   765
      End
      Begin VB.CommandButton cmdOKdocsproc 
         Height          =   630
         Left            =   600
         Picture         =   "FrmVentas2.frx":4CC1
         Style           =   1  'Graphical
         TabIndex        =   26
         Top             =   2535
         Width           =   750
      End
      Begin VB.CommandButton cmdSalirdocsproc 
         Height          =   630
         Left            =   2355
         Picture         =   "FrmVentas2.frx":4FCB
         Style           =   1  'Graphical
         TabIndex        =   25
         Top             =   2535
         Width           =   750
      End
      Begin VSFlex7Ctl.VSFlexGrid fgdocsproc 
         Height          =   1950
         Left            =   150
         TabIndex        =   24
         Top             =   450
         Width           =   3405
         _cx             =   6006
         _cy             =   3440
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
         FormatString    =   $"FrmVentas2.frx":52D5
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
      Begin VB.Line Line10 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   3660
         Y1              =   3270
         Y2              =   3270
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   3690
         X2              =   3690
         Y1              =   15
         Y2              =   3285
      End
      Begin VB.Line Line10 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   3675
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line5 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   -15
         Y2              =   3255
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Documentos Facturados en pantalla"
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
         TabIndex        =   72
         Top             =   90
         Width           =   3075
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   315
         Left            =   45
         Top             =   45
         Width           =   3615
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   7755
      Top             =   45
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
            Picture         =   "FrmVentas2.frx":535A
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas2.frx":589E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas2.frx":5C30
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas2.frx":5DB4
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas2.frx":6208
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas2.frx":6320
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas2.frx":6864
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas2.frx":6DA8
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas2.frx":6EBC
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas2.frx":6FD0
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas2.frx":7424
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas2.frx":7590
            Key             =   "IMG11"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas2.frx":7AD8
            Key             =   "IMG12"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmVentas2.frx":7DF2
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7470
      Left            =   15
      TabIndex        =   16
      Top             =   360
      Width           =   11895
      _cx             =   20981
      _cy             =   13176
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
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7050
         Left            =   12540
         TabIndex        =   31
         Top             =   375
         Width           =   11805
         Begin VB.Frame Frame3 
            Caption         =   "[ Opciones de Descuento]"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   645
            Left            =   7395
            TabIndex        =   71
            Top             =   2790
            Width           =   4395
            Begin VB.CheckBox Check1 
               Caption         =   "Ingresar Neto"
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
               Height          =   195
               Left            =   2640
               TabIndex        =   84
               Top             =   300
               Width           =   1500
            End
            Begin VB.OptionButton OptDes1 
               Caption         =   "Porcentaje"
               Height          =   195
               Left            =   150
               TabIndex        =   83
               Top             =   270
               Width           =   1215
            End
            Begin VB.OptionButton OptDes2 
               Caption         =   "Valor"
               Height          =   195
               Left            =   1440
               TabIndex        =   82
               Top             =   270
               Width           =   735
            End
            Begin VB.Line Line4 
               BorderColor     =   &H00FFFFFF&
               Index           =   1
               X1              =   2460
               X2              =   2460
               Y1              =   120
               Y2              =   570
            End
            Begin VB.Line Line4 
               BorderColor     =   &H80000003&
               Index           =   0
               X1              =   2490
               X2              =   2490
               Y1              =   120
               Y2              =   600
            End
         End
         Begin VB.TextBox TxtGlosa 
            Height          =   285
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   14
            Text            =   "TxtGlosa"
            Top             =   2490
            Width           =   10170
         End
         Begin VB.TextBox TxtTC 
            Alignment       =   1  'Right Justify
            BackColor       =   &H8000000F&
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
            Height          =   285
            Left            =   10680
            TabIndex        =   126
            Text            =   "TxtTC"
            Top             =   300
            Width           =   1065
         End
         Begin VB.CheckBox ChkTC 
            Caption         =   "Check2"
            Enabled         =   0   'False
            Height          =   195
            Left            =   10440
            TabIndex        =   125
            Top             =   360
            Width           =   195
         End
         Begin VB.Frame Frame6 
            Height          =   3000
            Left            =   9240
            TabIndex        =   73
            Top             =   3390
            Width           =   2550
            Begin VSFlex7Ctl.VSFlexGrid Fg4 
               Height          =   2445
               Left            =   60
               TabIndex        =   74
               Top             =   420
               Width           =   2430
               _cx             =   4286
               _cy             =   4313
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
               Cols            =   4
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmVentas2.frx":8184
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
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Documentos Adjuntos"
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
               Height          =   210
               Left            =   60
               TabIndex        =   75
               Top             =   180
               Width           =   1860
            End
         End
         Begin VB.CommandButton CmdBusDocRef2 
            Height          =   240
            Left            =   9405
            Picture         =   "FrmVentas2.frx":8209
            Style           =   1  'Graphical
            TabIndex        =   113
            Top             =   2220
            Width           =   240
         End
         Begin VB.CommandButton CmdBusDocRef 
            Height          =   240
            Left            =   8040
            Picture         =   "FrmVentas2.frx":833B
            Style           =   1  'Graphical
            TabIndex        =   110
            Top             =   1590
            Visible         =   0   'False
            Width           =   240
         End
         Begin VB.TextBox TxtNumDocRef 
            Height          =   300
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   13
            Text            =   "TxtNumDocRef"
            Top             =   2190
            Width           =   3390
         End
         Begin VB.CommandButton CmdBusIdTipDocRef 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmVentas2.frx":846D
            Style           =   1  'Graphical
            TabIndex        =   107
            Top             =   2220
            Width           =   240
         End
         Begin VB.Frame Frame10 
            Height          =   465
            Left            =   9630
            TabIndex        =   104
            Top             =   1290
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
               Left            =   120
               TabIndex        =   105
               Top             =   120
               Width           =   1860
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "[ Motivo de la Nota de Credito ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   645
            Left            =   2010
            TabIndex        =   86
            Top             =   4590
            Visible         =   0   'False
            Width           =   6855
            Begin VB.CommandButton cmdMotDev 
               Height          =   240
               Left            =   4140
               Picture         =   "FrmVentas2.frx":859F
               Style           =   1  'Graphical
               TabIndex        =   151
               Top             =   270
               Visible         =   0   'False
               Width           =   240
            End
            Begin VB.TextBox txtMotDevOtr 
               Height          =   315
               Left            =   4440
               TabIndex        =   150
               Text            =   "txtMotDevOtr"
               Top             =   240
               Visible         =   0   'False
               Width           =   2295
            End
            Begin VB.CommandButton CmdMotNotCre 
               Height          =   240
               Left            =   1650
               Picture         =   "FrmVentas2.frx":86D1
               Style           =   1  'Graphical
               TabIndex        =   123
               Top             =   285
               Width           =   240
            End
            Begin VB.TextBox TxtDocRef 
               Height          =   300
               Left            =   75
               MaxLength       =   50
               TabIndex        =   87
               Text            =   "TxtDocRef"
               Top             =   255
               Width           =   1830
            End
            Begin VB.TextBox txtMotDev 
               Height          =   300
               Left            =   1920
               MaxLength       =   50
               TabIndex        =   152
               Text            =   "txtMotDev"
               Top             =   240
               Visible         =   0   'False
               Width           =   2490
            End
            Begin VB.Label lblIdMotDev 
               Caption         =   "lblIdMotDev"
               ForeColor       =   &H000000C0&
               Height          =   210
               Left            =   4200
               TabIndex        =   153
               Top             =   30
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label LblIdConNC 
               Caption         =   "LblIdConNC"
               ForeColor       =   &H000000C0&
               Height          =   210
               Left            =   3060
               TabIndex        =   124
               Top             =   30
               Visible         =   0   'False
               Width           =   855
            End
         End
         Begin VB.CommandButton CmdBusAlm 
            Height          =   240
            Left            =   6720
            Picture         =   "FrmVentas2.frx":8803
            Style           =   1  'Graphical
            TabIndex        =   76
            Top             =   645
            Width           =   240
         End
         Begin VB.Frame FraRetencion 
            Caption         =   "[ Retención de 4ta Cat. ]"
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   645
            Left            =   4905
            TabIndex        =   63
            Top             =   2790
            Visible         =   0   'False
            Width           =   2460
            Begin VB.OptionButton OptSi 
               Caption         =   "Afecto"
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
               Height          =   195
               Left            =   90
               TabIndex        =   65
               Top             =   285
               Width           =   885
            End
            Begin VB.OptionButton OptNo 
               Caption         =   "No Afecto"
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
               Height          =   195
               Left            =   1080
               TabIndex        =   64
               Top             =   285
               Width           =   1200
            End
         End
         Begin VB.CommandButton CmdBusVen 
            Height          =   240
            Left            =   6720
            Picture         =   "FrmVentas2.frx":8935
            Style           =   1  'Graphical
            TabIndex        =   60
            Top             =   960
            Width           =   240
         End
         Begin VB.TextBox TxtIdVen 
            Height          =   300
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   5
            Text            =   "TxtIdVen"
            Top             =   930
            Width           =   705
         End
         Begin VB.CommandButton CmdBusTipItem 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmVentas2.frx":8A67
            Style           =   1  'Graphical
            TabIndex        =   18
            Top             =   645
            Width           =   240
         End
         Begin VB.CommandButton CmdBusCli 
            Height          =   240
            Left            =   3075
            Picture         =   "FrmVentas2.frx":8B99
            Style           =   1  'Graphical
            TabIndex        =   47
            Top             =   1275
            Width           =   240
         End
         Begin VB.CommandButton CmdBusTipDoc 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmVentas2.frx":8CCB
            Style           =   1  'Graphical
            TabIndex        =   46
            Top             =   960
            Width           =   240
         End
         Begin VB.TextBox TxtNumDoc 
            Height          =   300
            Left            =   2730
            Locked          =   -1  'True
            MaxLength       =   10
            TabIndex        =   8
            Text            =   "TxtNumDoc"
            Top             =   1560
            Width           =   1440
         End
         Begin VB.CommandButton CmdBusCondicion 
            Height          =   240
            Left            =   2220
            Picture         =   "FrmVentas2.frx":8DFD
            Style           =   1  'Graphical
            TabIndex        =   45
            Top             =   1905
            Width           =   240
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   6720
            Picture         =   "FrmVentas2.frx":8F2F
            Style           =   1  'Graphical
            TabIndex        =   44
            Top             =   330
            Width           =   240
         End
         Begin VB.TextBox TxtNumSer 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   7
            Text            =   "TxtNumSer"
            Top             =   1560
            Width           =   915
         End
         Begin VB.TextBox TxtTipDoc 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   4
            Text            =   "TxtTipDoc"
            Top             =   930
            Width           =   915
         End
         Begin VB.TextBox TxtNumRuc 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   6
            Text            =   "TxtNumRuc"
            Top             =   1245
            Width           =   1770
         End
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   1
            Text            =   "TxtIdMon"
            Top             =   300
            Width           =   705
         End
         Begin VB.TextBox TxtConPag 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   10
            Text            =   "TxtConPag"
            Top             =   1875
            Width           =   915
         End
         Begin VB.TextBox TxtTipItem 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   14
            TabIndex        =   2
            Text            =   "TxtTipItem"
            Top             =   615
            Width           =   915
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchDoc 
            Height          =   300
            Left            =   1575
            TabIndex        =   0
            Top             =   300
            Width           =   1260
            _ExtentX        =   2223
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
            Valor           =   "03/01/2004"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchVen 
            Height          =   300
            Left            =   6285
            TabIndex        =   11
            Top             =   1875
            Width           =   1215
            _ExtentX        =   2143
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
            Valor           =   "03/01/2004"
         End
         Begin VB.Frame Fratipven 
            Caption         =   "[ Tipo de Facturación ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00400000&
            Height          =   645
            Left            =   60
            TabIndex        =   48
            Top             =   2790
            Width           =   4815
            Begin VB.OptionButton optconcotizacion 
               Caption         =   "Orden Pedido"
               Enabled         =   0   'False
               Height          =   255
               Left            =   1950
               TabIndex        =   66
               Top             =   285
               Width           =   1290
            End
            Begin VB.CommandButton cmdagregardocs 
               Caption         =   "Adicionar"
               Enabled         =   0   'False
               Height          =   330
               Left            =   3465
               MaskColor       =   &H00E0E0E0&
               TabIndex        =   51
               Top             =   195
               Width           =   1200
            End
            Begin VB.OptionButton optsinguia 
               Caption         =   "Sin Guia"
               Enabled         =   0   'False
               Height          =   270
               Left            =   75
               TabIndex        =   50
               Top             =   285
               Width           =   945
            End
            Begin VB.OptionButton optconguia 
               Caption         =   "Guia"
               Enabled         =   0   'False
               Height          =   270
               Left            =   1140
               TabIndex        =   49
               Top             =   285
               Value           =   -1  'True
               Width           =   675
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   2910
            Left            =   60
            TabIndex        =   15
            Top             =   3465
            Width           =   11670
            _cx             =   20585
            _cy             =   5133
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
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   20
            Cols            =   16
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmVentas2.frx":9061
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
         Begin VB.Frame Frame4 
            Height          =   750
            Left            =   60
            TabIndex        =   32
            Top             =   6300
            Width           =   11700
            Begin VB.CommandButton CmdPreHist 
               Caption         =   "Ver His. Precios"
               Enabled         =   0   'False
               Height          =   495
               Left            =   3645
               Style           =   1  'Graphical
               TabIndex        =   106
               Top             =   165
               Width           =   1170
            End
            Begin VB.CommandButton CmdSel 
               Caption         =   "&Detallar Item"
               Height          =   495
               Left            =   2440
               Style           =   1  'Graphical
               TabIndex        =   85
               Top             =   165
               Width           =   1170
            End
            Begin VB.CommandButton CmdDelItem 
               Caption         =   "&Eliminar Item"
               Enabled         =   0   'False
               Height          =   495
               Left            =   1235
               TabIndex        =   81
               Top             =   165
               Width           =   1170
            End
            Begin VB.CommandButton CmdAddItem 
               Caption         =   "&Agregar Item"
               Enabled         =   0   'False
               Height          =   495
               Left            =   30
               TabIndex        =   80
               Top             =   165
               Width           =   1170
            End
            Begin VB.TextBox txtisc 
               Alignment       =   1  'Right Justify
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
               Left            =   9180
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   37
               TabStop         =   0   'False
               Text            =   "TxtIsc"
               Top             =   360
               Width           =   1200
            End
            Begin VB.TextBox txtinafecto 
               Alignment       =   1  'Right Justify
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
               Left            =   6285
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   36
               TabStop         =   0   'False
               Text            =   "TxtInafecto"
               Top             =   360
               Width           =   1200
            End
            Begin VB.TextBox TxtBruto 
               Alignment       =   1  'Right Justify
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
               Left            =   5055
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   35
               TabStop         =   0   'False
               Text            =   "TxtBruto"
               Top             =   360
               Width           =   1200
            End
            Begin VB.TextBox TxtIGV 
               Alignment       =   1  'Right Justify
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
               Left            =   7605
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   34
               TabStop         =   0   'False
               Text            =   "TxtIGV"
               Top             =   360
               Width           =   1200
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
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
               Left            =   10380
               Locked          =   -1  'True
               MaxLength       =   10
               TabIndex        =   33
               TabStop         =   0   'False
               Text            =   "TxtTotal"
               Top             =   360
               Width           =   1200
            End
            Begin VB.Line Line3 
               BorderColor     =   &H00FFFFFF&
               Index           =   1
               X1              =   4920
               X2              =   4920
               Y1              =   90
               Y2              =   810
            End
            Begin VB.Line Line3 
               BorderColor     =   &H80000003&
               Index           =   0
               X1              =   4905
               X2              =   4905
               Y1              =   105
               Y2              =   825
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "I.S.C."
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
               Index           =   3
               Left            =   9180
               TabIndex        =   43
               Top             =   120
               Width           =   495
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Inafecto"
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
               Index           =   1
               Left            =   6255
               TabIndex        =   42
               Top             =   120
               Width           =   1140
            End
            Begin VB.Label LblIgvTasa 
               Alignment       =   2  'Center
               AutoSize        =   -1  'True
               Caption         =   "LblIgvTasa"
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
               Height          =   195
               Left            =   7950
               TabIndex        =   41
               Top             =   120
               Width           =   960
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Imp. Bruto"
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
               Index           =   0
               Left            =   5070
               TabIndex        =   40
               Top             =   120
               Width           =   885
            End
            Begin VB.Label LblRotulo 
               AutoSize        =   -1  'True
               Caption         =   "I.G.V. (         )"
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
               Left            =   7560
               TabIndex        =   39
               Top             =   120
               Width           =   1230
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total"
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
               Index           =   2
               Left            =   10365
               TabIndex        =   38
               Top             =   120
               Width           =   450
            End
         End
         Begin VB.TextBox TxtIdAlm 
            Height          =   300
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   3
            Text            =   "TxtIdAlm"
            Top             =   615
            Width           =   705
         End
         Begin VB.TextBox TxtIdTipDoc 
            Height          =   300
            Left            =   1575
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   12
            Text            =   "TxtIdTipDoc"
            Top             =   2190
            Width           =   915
         End
         Begin VB.TextBox TxtDocRefCredi 
            Height          =   300
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   9
            Text            =   "TxtDocRefCredi"
            Top             =   1560
            Visible         =   0   'False
            Width           =   2025
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Index           =   10
            Left            =   75
            TabIndex        =   127
            Top             =   2520
            Width           =   405
         End
         Begin VB.Label lblReg 
            Alignment       =   1  'Right Justify
            Caption         =   "lblReg"
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
            Height          =   270
            Left            =   9555
            TabIndex        =   122
            Top             =   30
            Width           =   2190
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Condición de Pago"
            Height          =   195
            Index           =   4
            Left            =   75
            TabIndex        =   121
            Top             =   1920
            Width           =   1350
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Documento"
            Height          =   195
            Index           =   0
            Left            =   75
            TabIndex        =   120
            Top             =   1605
            Width           =   1275
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo de Documento"
            Height          =   195
            Index           =   1
            Left            =   75
            TabIndex        =   119
            Top             =   975
            Width           =   1410
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Index           =   7
            Left            =   75
            TabIndex        =   118
            Top             =   1290
            Width           =   480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Emisión"
            Height          =   195
            Index           =   2
            Left            =   75
            TabIndex        =   117
            Top             =   345
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Item"
            Height          =   195
            Index           =   6
            Left            =   75
            TabIndex        =   116
            Top             =   660
            Width           =   660
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tip. Doc. Ref."
            Height          =   195
            Index           =   8
            Left            =   75
            TabIndex        =   115
            ToolTipText     =   "Tipo de Documento de Referencia"
            Top             =   2235
            Width           =   1005
         End
         Begin VB.Label LblIdDocRef2 
            AutoSize        =   -1  'True
            Caption         =   "LblIdDocRef2"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   9735
            TabIndex        =   114
            Top             =   2235
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label33 
            AutoSize        =   -1  'True
            Caption         =   "Referente al Documento"
            Height          =   195
            Left            =   4455
            TabIndex        =   112
            Top             =   1605
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.Label LblIdDocRef 
            AutoSize        =   -1  'True
            Caption         =   "LblIdDocRef"
            ForeColor       =   &H00000080&
            Height          =   195
            Left            =   8340
            TabIndex        =   111
            Top             =   1605
            Visible         =   0   'False
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Doc. Ref."
            Height          =   195
            Index           =   9
            Left            =   5505
            TabIndex        =   109
            ToolTipText     =   "Documento de Referencia"
            Top             =   2235
            Width           =   690
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
            Left            =   2520
            TabIndex        =   108
            Top             =   2190
            Width           =   2655
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
            Left            =   7020
            TabIndex        =   79
            Top             =   615
            Width           =   2655
         End
         Begin VB.Label LblIdAlmacen 
            AutoSize        =   -1  'True
            Caption         =   "LblIdAlmacen"
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   3780
            TabIndex        =   78
            Top             =   390
            Visible         =   0   'False
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Almacén"
            Height          =   195
            Index           =   11
            Left            =   5580
            TabIndex        =   77
            Top             =   660
            Width           =   615
         End
         Begin VB.Label LblNomVen 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomVen"
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
            Left            =   7020
            TabIndex        =   62
            Top             =   930
            Width           =   4710
         End
         Begin VB.Label Lblvendedor 
            AutoSize        =   -1  'True
            Caption         =   "Vendedor"
            Height          =   195
            Left            =   5505
            TabIndex        =   61
            Top             =   975
            Width           =   690
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de Ventas"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Left            =   105
            TabIndex        =   58
            Top             =   45
            Width           =   11595
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
            Left            =   2520
            TabIndex        =   19
            Top             =   615
            Width           =   2715
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Ven."
            Height          =   195
            Index           =   3
            Left            =   5505
            TabIndex        =   57
            ToolTipText     =   "Fecha de Vencimiento"
            Top             =   1920
            Width           =   690
         End
         Begin VB.Label LblTipCam 
            AutoSize        =   -1  'True
            Caption         =   "T.C."
            Height          =   195
            Left            =   10125
            TabIndex        =   56
            Top             =   345
            Width           =   300
         End
         Begin VB.Label LblCondPag 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblCondPag"
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
            Left            =   2520
            TabIndex        =   55
            Top             =   1875
            Width           =   2655
         End
         Begin VB.Label LblMoneda 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblMoneda"
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
            Left            =   7020
            TabIndex        =   21
            Top             =   300
            Width           =   2655
         End
         Begin VB.Label LblNomCli 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomCli"
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
            Left            =   3375
            TabIndex        =   54
            Top             =   1245
            Width           =   4635
         End
         Begin VB.Label LblNomDoc 
            BorderStyle     =   1  'Fixed Single
            Caption         =   "LblNomDoc"
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
            Left            =   2520
            TabIndex        =   53
            Top             =   930
            Width           =   2715
         End
         Begin VB.Shape Shape1 
            BackColor       =   &H80000001&
            BackStyle       =   1  'Opaque
            Height          =   90
            Left            =   2550
            Top             =   1665
            Width           =   105
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   5610
            TabIndex        =   20
            Top             =   345
            Width           =   585
         End
         Begin VB.Label LblIdCliente 
            AutoSize        =   -1  'True
            Caption         =   "LblIdCliente"
            ForeColor       =   &H00000040&
            Height          =   195
            Left            =   2865
            TabIndex        =   52
            Top             =   360
            Visible         =   0   'False
            Width           =   825
         End
      End
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   7050
         Left            =   45
         TabIndex        =   28
         Top             =   375
         Width           =   11805
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6735
            Left            =   30
            TabIndex        =   17
            Top             =   300
            Width           =   11760
            _ExtentX        =   20743
            _ExtentY        =   11880
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
            Columns(1).Caption=   "Nº Reg"
            Columns(1).DataField=   "numreg1"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "TD"
            Columns(2).DataField=   "abrev"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Nº Documento"
            Columns(3).DataField=   "numerodoc"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Fch. Emi"
            Columns(4).DataField=   "fchdoc1"
            Columns(4).NumberFormat=   "Short Date"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Fch. Ven."
            Columns(5).DataField=   "fchven1"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Cliente"
            Columns(6).DataField=   "nombre"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "M."
            Columns(7).DataField=   "simbolo"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "T.C."
            Columns(8).DataField=   "impven1"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Imp. Bru."
            Columns(9).DataField=   "impbru1"
            Columns(9).NumberFormat=   "0.00"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "I.G.V."
            Columns(10).DataField=   "impigv1"
            Columns(10).NumberFormat=   "0.00"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Importe"
            Columns(11).DataField=   "imptotdoc1"
            Columns(11).NumberFormat=   "0.00"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "Saldo"
            Columns(12).DataField=   "impsal1"
            Columns(12).NumberFormat=   "0.00"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   13
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   503
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=13"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1588"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1508"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=661"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=582"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=512"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=2566"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=2487"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1455"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1376"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=1667"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=1588"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=513"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=4366"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=4286"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=512"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=688"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=609"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=513"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=953"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=873"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=514"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(56)=   "Column(9).Width=1508"
            Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=1429"
            Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=514"
            Splits(0)._ColumnProps(61)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(62)=   "Column(10).Width=1244"
            Splits(0)._ColumnProps(63)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(64)=   "Column(10)._WidthInPix=1164"
            Splits(0)._ColumnProps(65)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(66)=   "Column(10)._ColStyle=514"
            Splits(0)._ColumnProps(67)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(68)=   "Column(11).Width=1508"
            Splits(0)._ColumnProps(69)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(70)=   "Column(11)._WidthInPix=1429"
            Splits(0)._ColumnProps(71)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(72)=   "Column(11)._ColStyle=514"
            Splits(0)._ColumnProps(73)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(74)=   "Column(12).Width=1588"
            Splits(0)._ColumnProps(75)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(76)=   "Column(12)._WidthInPix=1508"
            Splits(0)._ColumnProps(77)=   "Column(12)._EditAlways=0"
            Splits(0)._ColumnProps(78)=   "Column(12)._ColStyle=514"
            Splits(0)._ColumnProps(79)=   "Column(12).Order=13"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=90,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=87,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=88,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=89,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=55,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=56,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=57,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=29,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=30,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=31,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=46,.parent=13,.alignment=2"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=43,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=44,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=45,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=2"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=67,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=68,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=69,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=82,.parent=13,.alignment=1"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=63,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=64,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=65,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=78,.parent=13,.alignment=1"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=75,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=76,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=77,.parent=17"
            _StyleDefs(80)  =   "Splits(0).Columns(11).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(81)  =   "Splits(0).Columns(11).HeadingStyle:id=51,.parent=14"
            _StyleDefs(82)  =   "Splits(0).Columns(11).FooterStyle:id=52,.parent=15"
            _StyleDefs(83)  =   "Splits(0).Columns(11).EditorStyle:id=53,.parent=17"
            _StyleDefs(84)  =   "Splits(0).Columns(12).Style:id=74,.parent=13,.alignment=1"
            _StyleDefs(85)  =   "Splits(0).Columns(12).HeadingStyle:id=71,.parent=14"
            _StyleDefs(86)  =   "Splits(0).Columns(12).FooterStyle:id=72,.parent=15"
            _StyleDefs(87)  =   "Splits(0).Columns(12).EditorStyle:id=73,.parent=17"
            _StyleDefs(88)  =   "Named:id=33:Normal"
            _StyleDefs(89)  =   ":id=33,.parent=0"
            _StyleDefs(90)  =   "Named:id=34:Heading"
            _StyleDefs(91)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(92)  =   ":id=34,.wraptext=-1"
            _StyleDefs(93)  =   "Named:id=35:Footing"
            _StyleDefs(94)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(95)  =   "Named:id=36:Selected"
            _StyleDefs(96)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(97)  =   "Named:id=37:Caption"
            _StyleDefs(98)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(99)  =   "Named:id=38:HighlightRow"
            _StyleDefs(100) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(101) =   "Named:id=39:EvenRow"
            _StyleDefs(102) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(103) =   "Named:id=40:OddRow"
            _StyleDefs(104) =   ":id=40,.parent=33"
            _StyleDefs(105) =   "Named:id=41:RecordSelector"
            _StyleDefs(106) =   ":id=41,.parent=34"
            _StyleDefs(107) =   "Named:id=42:FilterBar"
            _StyleDefs(108) =   ":id=42,.parent=33"
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
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
            TabIndex        =   29
            Top             =   30
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta Ventas"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   30
            Top             =   45
            Width           =   11565
         End
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000005&
         BorderWidth     =   2
         X1              =   12840
         X2              =   24645
         Y1              =   375
         Y2              =   7425
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   11925
      _ExtentX        =   21034
      _ExtentY        =   1058
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   16
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
                  Text            =   "Modificar Factura"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Restaurar Factura"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Modificar Saldo del Documento"
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
                  Text            =   "Anular Factura"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar Factura"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Emitir Factura Anulada"
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
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Object.ToolTipText     =   "Exportar MSExcel"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Imprimir Documento"
            ImageIndex      =   12
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Imprimir Documento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Exportar a Excel"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.Menu menu1 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu menu1_1 
         Caption         =   "Agregar Item            "
      End
      Begin VB.Menu menu1_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_3 
         Caption         =   "Eliminar Item                "
      End
      Begin VB.Menu menu1_4 
         Caption         =   "-"
      End
      Begin VB.Menu menu1_5 
         Caption         =   "Ver Historico de Precios"
      End
   End
   Begin VB.Menu menu2 
      Caption         =   "Menu2"
      Visible         =   0   'False
      Begin VB.Menu menu2_1 
         Caption         =   "Agregar Documento"
      End
      Begin VB.Menu menu2_2 
         Caption         =   "-"
      End
      Begin VB.Menu menu2_3 
         Caption         =   "Eliminar Documento"
      End
   End
End
Attribute VB_Name = "FrmVentas2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre            : FrmVentas2
'* Tipo              : FORMULARIO
'* Descripcion       : FORMULARIO DONDE SE REGISTRAN LAS VENTAS, PERMITIENDO DETALLAR LOS ITEMS DE LA
'*                     VENTA, ASI MISMO SE GENERA EL PROCESO CONTABLE PARA LA VENTA
'* DISEÑADO POR      : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION   : 24/09/09
'* VERSION           : 1.0
'*****************************************************************************************************

Option Explicit
Dim RstVent As New ADODB.Recordset   ' RECORDSET EN EL QUE SE CARGARAN LA VENTAS, ESTE RECORDSET SE VISUALIZARA EN LA PESTAÑA CONSULTA
Dim RstVentItemsDeta As New ADODB.Recordset  ' RECORDSET QUE ALMACENARA LA INFORMACION ADICIONAL DEL ITEM QUE SE IMPRIMIRA EN LA FACTURA
Dim QueHace As Integer               ' VARIABLE QUE INFORMA EN QUE MODO SE ENCUENTRA EL FORMULARIO
Dim TasaImpuesto As Double           ' ALMACENA LA TASA DEL IMPUESTO
Dim CaracteresNumericos As String    ' ALMACENA LOS CARACTERES NUMERICOS QUE SE UTILIZARAN EN LOS CONTROLES TextBox
Dim SeEjecuto As Boolean             ' VARIABLE QUE INDICA QUE EL EVENTO ACTIVATE YA SE EJECUTO
Dim ValTipCam As Double              ' VARIABLE QUE ALMACENA EL VALOR DEL TIPO DE CAMBIO
Dim xIdCuenTasa As Integer           ' codigo de la cuenta contable del impuesto
Dim xCuentaDoc As Integer            ' codigo de la cuenta contable del documento
Dim Mostrando As Boolean             ' VARIABLE QUE INFORMA A LOS CONTROLES FlexGrid QUE SE ESTA INSERTANDO UN FILA
Dim swguiafact                       ' 0 No se facturaron, 1 Se facturaron
Dim Agregando As Boolean             ' para saber cuando se este agregando datos en el grid de productos
Dim xHorIni As Date                  ' ESEPECIFICA LA HORA DE INICIO DE INGRESO DEL REGISTRO
Dim fOrdenLista As Boolean           ' especfica el orden de la lista de la consulta
Dim mIdRegistro&                     ' identificador del registro
Dim mMesActivo As Integer            ' VARIABLE QUE ALAMCENA EL ID DEL MES ACTIVO
Dim fCierrePeriodo As Boolean        ' indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)

Dim JALOPEDIDO As Boolean            ' INDICA SI LA VENTA ESTA HACIENDO REFERENCIA A UNA ORDEN DE PEDIDO
Dim VAR_IDPEDIDO As Integer          ' ESPECIFICA EL ID DEL PEDIDO
Dim VAR_FECHAPEDIDO As String        ' ESPECIFICA LA FECHA DEL PEDIDO

Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO
Dim F As New SistemaLogica.Funciones
'*****************************************************************************************************
'* Nombre           : ActivarEntorno
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES Tabone y Toolbar DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ActivarEntorno()
    TabOne1.Enabled = Not TabOne1.Enabled
    Toolbar1.Enabled = Not Toolbar1.Enabled
End Sub

'*****************************************************************************************************
'* Nombre           : Eliminar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ELIMINA UN REGISTRO DE LA TABLA vta_ventas, CUANDO SE ELIMINA UN REGISTRO LOS
'*                    ITEMS DEL ALMACEN RESTORNAN A SU STOCK ORIGINAL, TAMBIEN SE ELIMINAN LOS ASIENTOS
'*                    CONTABLES QUE SE HAYNA GENERADO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Eliminar()
    Dim Rpta As Integer
    Dim F As New SistemaLogica.Funciones
    
    If RstVent.RecordCount = 0 Then
        MsgBox "No hay documentos para eliminar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    TabOne1.CurrTab = 0
    Rpta = MsgBox("¿ Esta seguro de eliminar el documento seleccionado ?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    ' SI EL USUARIO CONFIRMA QUE SE VA A ELIMINAR EL REGISTRO
    If Rpta = vbYes Then
        ' ELIMINAMOS EL ASIENTO CONTABLE DE LA TABLA con_diario
        xCon.Execute "DELETE * FROM con_diario WHERE idmov = " & RstVent("id") & " AND idlib = 2 AND Iddoc = " & RstVent("tipdoc") & ""
               
        ' DETERMINAMOS SI EL REGISTRO ESTA VINCULADO A UN DOCUMENTO D INGRESO O A UN PEDIDO
        If RstVent("oriitem") = 1 Then
            ' si el origen del item es igual a 1 Actualizamos el saldo del stock
            Call ActualizarStock("E", RstVent("id"))
        End If
        If RstVent("oriitem") = 2 Then
            ' actualizamos a 0 el campo "iddocven" de la tabla vta_guia para poder facturarla con otro numero de factura
            xCon.Execute "UPDATE vta_guia SET vta_guia.iddocven = 0 WHERE (((vta_guia.iddocven)=" & RstVent("id") & "))"
        End If
        
        If RstVent("oriitem") = 3 Then
            ' eliminamos la referencia del documento a la orden de pedido
            xCon.Execute "UPDATE ped_pedidodetent SET ped_pedidodetent.idtipdoc = 0, ped_pedidodetent.iddocven = 0, ped_pedidodetent.estado = 2" _
                & " WHERE (((ped_pedidodetent.idtipdoc)=2) AND ((ped_pedidodetent.iddocven)=" & RstVent("id") & "))"
        End If
        
        ' ELIMINAMOS EL ASIENTO CONTABLE DE LA TABLA con_diario
        xCon.Execute "DELETE * FROM con_diario WHERE idlib = 2 AND idmov = " & RstVent("id") & ""
        
        ' ELIMINAMOS EL DETALLE DEL ITEM
        xCon.Execute "DELETE * FROM vta_ventasdetitems WHERE idventa = " & RstVent("id") & ""
        
        '--eliminamos el registro del analisis de cta cte
        xCon.Execute "DELETE * FROM var_analisisctacte WHERE idlib = 2 AND idope = " & RstVent("id") & ""
        
        ' ELIMINAMOS EL REGISTRO
        
        xCon.Execute "DELETE * FROM vta_ventasdetitems WHERE idventa = " & RstVent("id") & ""
        xCon.Execute "DELETE * FROM vta_ventasdet WHERE idvta = " & RstVent("id") & ""
        xCon.Execute "DELETE * FROM vta_ventas WHERE id = " & RstVent("id") & ""
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstVent("id") & " AND idform = " & IdMenuActivo
        
        '**********************************
        ' Eliminamos los Movimientos Generados
        If F.NuloNumeric(F.KeyValue("CreacionMovimientoAutoVenta", xCon)) = -1 Then
            ' Verificamos si ya tiene registro en movimientos
            Dim database As New SistemaData.EDataBase
            Dim record As New ADODB.Recordset
            Dim Movimiento As New AlmacenEntidad.EMovimiento
            
            Set database.Connection = xCon
            database.CommandText = "SELECT alm_ingreso.id AS idmov " _
                        + vbCr + "FROM alm_ingreso " _
                        + vbCr + "WHERE (((alm_ingreso.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoFactura", xCon)) & ") AND ((alm_ingreso.iddocref)=" & F.NuloNumeric(RstVent("id")) & "))"
            Set record = database.GetRecordset
            If record.RecordCount > 0 Then
                Movimiento.IdMovimiento = F.NuloNumeric(record("idmov"))
                Set Movimiento.Conexion = xCon
                Movimiento.Delete CLng(xIdUsuario), ""
            End If
        End If
        '**********************************
        
        MsgBox NulosC(RstVent("nomdoc")) & " se eliminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        RstVent.Requery
        Dg1.Refresh

    End If
End Sub

'*****************************************************************************************************
'* Nombre           : ActivaTool
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LA BARRA DE HERRAMIENTAS DEL SISTEMA
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

'*****************************************************************************************************
'* Nombre           : RestaurarFactura
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : RESTAURA UN REGISTRO ANULADO, PARA ELLO ACTUALIZA EL CAMPO anulado = 0 EN LA
'*                    TABLA vta_ventas
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub RestaurarFactura()
    'Se restaura una factura anulada
    Dim Rpta As Integer
    
    Rpta = MsgBox("Esta seguro de restaurar la factura Nº " + NulosC(RstVent("numser")) & "-" & NulosC(RstVent("numdoc")), vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption)
    If Rpta = vbYes Then
        
        mIdRegistro = RstVent("id")
        
        xCon.Execute "UPDATE vta_ventas SET vta_ventas.Anulado = 0, " _
            & " vta_ventas.idcli = 0  " _
            & " WHERE vta_ventas.id =" & RstVent("id") & ""
        
        xCon.Execute "DELETE * FROM vta_ventasdet WHERE vta_ventasdet.idvta =" & RstVent("id") & ""
        RstVent.Requery
        Dg1.Refresh
        
        '--posicionar en la posicion inicial
        If RstVent.RecordCount <> 0 Then
            RstVent.MoveFirst
            RstVent.Find "id=" & mIdRegistro
            If RstVent.EOF = True Then RstVent.MoveFirst
        End If
        
        MsgBox "La factura se restauró con éxito", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
    End If
    
    
End Sub

'*****************************************************************************************************
'* Nombre           : ActualizarStock
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTUALIZA EL SALDO DEL STOCK DE LOS ITEMS QUE PERTENESCAN AL DOCUMENTO DE VENTA
'* Paranetros       : NOMBRE    |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    TIPO      |  STRING     |  ESPECFICA EL TIPO DE OPERACION QUE REALIZARA EL
'*                                               PROCEDIMIENTO S = SALIDA; 3 = EXTONOR
'*                    numid     |  INTEGER    |  ESPECIFICA EL ID DEL DOCUMENTO DE VENTA
'* Devuelve         :
'*****************************************************************************************************
Sub ActualizarStock(TIPO As String, numid As Double)
    ' Tipo S = Salida
    ' Tipo E = Extorno por Anular Guia , Eliminar Guia
    Dim RstDet As New ADODB.Recordset
    Dim Rstitem As New ADODB.Recordset
    Dim xcant As Double

   ' SI NO ESTA EL ID DEl DOCUMENTO ES UN DOCUMENTO SIN GUIA DE REMISION
    RST_Busq RstDet, "SELECT vta_guia.* From vta_guia WHERE vta_guia.iddocven = " & numid & "", xCon

    If RstDet.RecordCount = 0 Then
        Set RstDet = Nothing
        ' OBTENEMOS LOS ITEMS DEL DOCUMENTO DE VENTA
        RST_Busq RstDet, "SELECT vta_ventasdet.* FROM vta_ventasdet WHERE idvta = " & numid & "", xCon
        Do While Not RstDet.EOF
            RST_Busq Rstitem, "SELECT Alm_Inventario.* FROM ALM_Inventario WHERE id = " & RstDet("iditem") & "", xCon
            'ACTUALIZAMOS LOS STOCK
            If Rstitem.RecordCount > 0 Then
                If TIPO = "S" Then
                    ' SI ES UNA SALIDA RESTAMOS DEL STOCK ACTUAL
                    Rstitem("stckact") = Rstitem("stckact") - RstDet("canpro")
                ElseIf TIPO = "E" Then
                    ' SI ES UNA ENTRADA AGREGAMOS AL STOCK ACTUAL
                    Rstitem("stckact") = Rstitem("stckact") + RstDet("canpro")
                End If
                Rstitem.Update
            End If
            
            RstDet.MoveNext
        Loop
    End If
    
    Set RstDet = Nothing
    Set Rstitem = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : Anular
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ANULA UN REGISTRO DE LA TABLA vta_ventas, PARA ELLO ACTUALIZA EL CAMPO
'*                    ANULADO = -1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Anular()
    Dim Rpta As Integer
    Dim A As Integer
    Dim Rst As New ADODB.Recordset
        
    If RstVent.RecordCount = 0 Or RstVent.EOF = True Then
        MsgBox "No hay registros para anularo", vbInformation, xTitulo
        Exit Sub
    End If
        
    If RstVent("anulado") = -1 Then
        MsgBox "El registro esta anulado", vbInformation, xTitulo
        Exit Sub
    End If
    
    Rpta = MsgBox("¿Esta seguro de anular " & RstVent("nomdoc") & " Nº " & RstVent("numser") & "-" & RstVent("numdoc") + "?", vbYesNo + vbDefaultButton1 + vbQuestion, Me.Caption)
    
    Dim xNumAsiento As String
    
    If Rpta = vbYes Then
    
        On Error GoTo LaCague
        
        xCon.BeginTrans
    
        mIdRegistro = RstVent("id")
    
        ' ANULAMOS EL REGISTRO ACTUALIZANDO EL CAMPO anulado = 0  Y ACTUALIZAMOS LOS VALORES DEL REGISTRO A O
        xCon.Execute "UPDATE vta_ventas SET vta_ventas.Anulado = -1, " _
            & " vta_ventas.impbru = 0, vta_ventas.impinaf = 0, vta_ventas.impigv = 0,  vta_ventas.impisc = 0,  " _
            & " vta_ventas.impotr = 0, vta_ventas.imptotdoc = 0,  vta_ventas.impsal = 0  " _
            & " WHERE vta_ventas.id = " & RstVent("id") & " "
        
        If RstVent("oriitem") = 1 Then
            ' si el origen del item es igual a 1 Actualizamos el saldo del stock
            Call ActualizarStock("E", RstVent("id"))
        End If
        If RstVent("oriitem") = 2 Then
            ' actualizamos a 0 el campo "iddocven" d ela tabla vta_guia para poder facturarla con otro numero de factura
            xCon.Execute "UPDATE vta_guia SET vta_guia.iddocven = 0 WHERE (((vta_guia.iddocven)=" & RstVent("id") & "))"
        End If
        
        If RstVent("oriitem") = 3 Then
            ' eliminamos la referencia del documento a la orden de pedido
            xCon.Execute "UPDATE ped_pedidodetent SET ped_pedidodetent.idtipdoc = 0, ped_pedidodetent.iddocven = 0, ped_pedidodetent.estado = 2" _
                & " WHERE (((ped_pedidodetent.idtipdoc)=2) AND ((ped_pedidodetent.iddocven)=" & RstVent("id") & "))"
        End If
        
        ' ELIMINAMOS EL DETALLE DEL REGISTRO
        xCon.Execute "DELETE * FROM vta_ventasdet WHERE vta_ventasdet.idvta = " & RstVent("id") & ""
                
        ' ELIMINAMOS EL DETALLE DEL ITEM
        xCon.Execute "DELETE * FROM vta_ventasdetitems WHERE idventa = " & RstVent("id") & ""
                
        ' ACTUALIZAMOS LOS REGISTRO DEL ASIENTO DEL DOCUMENTO A 0
   
        xNumAsiento = GenerarAsiento(xCon, 2, RstVent("id"), AnoTra, mMesActivo, 1)
        If xNumAsiento = "" Then GoTo LaCague

    
        ' Grabamos el movimiento en la tabla var_edicion
        GrabarOperacion xIdUsuario, IdMenuActivo, 2, Time, Time, Date, xCon, NulosN(RstVent("id"))
        
        xCon.CommitTrans
        
        MsgBox RstVent("nomdoc") & " se anuló con éxito" & vbCr & "Registro Nº: " & xNumAsiento, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        
        RstVent.Requery
        Dg1.Refresh
        '--posicionar en la posicion inicial
        If RstVent.RecordCount <> 0 Then
            RstVent.MoveFirst
            RstVent.Find "id=" & mIdRegistro
            If RstVent.EOF = True Then RstVent.MoveFirst
        End If
        
    End If
    Exit Sub
LaCague:
    xCon.RollbackTrans
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description), xTitulo
    
    
End Sub

'*****************************************************************************************************
'* Nombre           : Cancelar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CANCELAR EL PROCESO DE AGREGAR O MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Cancelar()
    Dim X As Integer
    Bloquea
    optsinguia.Enabled = False
    optconguia.Enabled = False
    Fg1.ColComboList(1) = ""
    Label5.Caption = "Detalle de Venta"
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
    Fg1.SelectionMode = flexSelectionByRow
     
    ActivaTool
    QueHace = 3
    Dg1.SetFocus
       
    ' Colocamos en el campo estado 0  de la tabla guia que indica no esta facturado
    If fgdocsproc.Rows - 1 > 0 Then
        If swguiafact = 0 Then
            For X = 1 To fgdocsproc.Rows - 1
                xCon.Execute " UPDATE vta_guia SET Vta_guia.Estado = 0 WHERE vta_guia.id = " & NulosN(fgdocsproc.TextMatrix(X, 1)) & ""
            Next
            fgdocsproc.Rows = 1
        End If
    End If
    swguiafact = 0
End Sub

'*****************************************************************************************************
'* Nombre           : Nuevo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA EL INGRESO DE UN NUEVO REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Nuevo()
    If PuedeAgregarRegistro("VENTAS", xCon) = False Then
        MsgBox "Esta utilizando una versión de prueba del maravilloso sistema SEVEN Soft, si desea la versión comercial contactese con el " & Chr(13) _
            & " extraordinario programador Enrique Pollongo a eps_76@hotmail.com y solicite un número de licencia para esta PC", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    QueHace = 1
    Blanquea
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    Label5.Caption = "Agregando Venta"
    Fg1.ColComboList(1) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    OptSi.Value = True
    Fg1.Rows = 1
    Fg4.Rows = 1
    optsinguia.Value = True
    optsinguia_Click
    OptDes1.Value = True

    TxtFchDoc.Valor = Format(Date, "dd/mm/yyyy")
    xHorIni = Time
    If Check1.Value = 1 Then Check1_Click
    JALOPEDIDO = False
    
    Set RstVentItemsDeta = Nothing
        
    RST_Busq RstVentItemsDeta, "SELECT vta_ventasdetitems.idventa, vta_ventasdetitems.iditem, vta_ventasdetitems.orden, vta_ventasdetitems.texto " _
        & " From vta_ventasdetitems Where (((vta_ventasdetitems.idventa) = " & 9999 & ")) ORDER BY vta_ventasdetitems.iditem, vta_ventasdetitems.orden", xCon
    
    Set RstVentItemsDeta.ActiveConnection = Nothing
    TxtFchDoc.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Modificar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA EL FORMULARIO PARA MODIFICAR UN REGISTRO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Modificar()
    If RstVent.RecordCount = 0 Then
        MsgBox "No hay Registros para Modificar", vbInformation, Me.Caption
        Exit Sub
    End If
    If NulosN(RstVent("anulado")) = -1 Then
        MsgBox "El Documento de Venta esta Anulado" & vbCr & "No se Puede Modificar", vbInformation, Me.Caption
        Exit Sub
    End If
   
    QueHace = 2
    Bloquea
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivaTool
    MuestraSegundoTab
    Label5.Caption = "Modificando Ventas"
    Fg1.ColComboList(1) = "|..."
    Fg1.ColComboList(0) = "|..."
    Fg1.SelectionMode = flexSelectionFree
    xHorIni = Time
    TxtFchDoc.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : MuestraSegundoTab
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA INFORMACION DETALLADA DEL REGISTRO EN LA PESTAÑA DETALLE
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MuestraSegundoTab()
    Dim Rst As New ADODB.Recordset
    
    If RstVent.RecordCount = 0 Then Exit Sub
    Blanquea
    lblReg.Caption = "Nº Reg. " & NulosC(RstVent("numreg1"))
    
    ' MOSTRAMOS LOS DATOS DEL REGISTRO
    TxtTipItem.Text = NulosN(RstVent("idtipo"))
    TxtTipDoc.Text = NulosN(RstVent("tipdoc"))
    
    TxtNumRuc.Text = NulosC(RstVent("numruc"))
    TxtNumSer.Text = NulosC(RstVent("numser"))
    TxtNumDoc.Text = NulosC(RstVent("numdoc"))
    If IsDate(RstVent("fchdoc")) = True Then TxtFchDoc.Valor = CDate(RstVent("fchdoc"))
    If IsDate(RstVent("fchven")) = True Then TxtFchVen.Valor = CDate(RstVent("fchven"))
    
    TxtConPag.Text = NulosN(RstVent("idconpag"))
    TxtIdMon.Text = NulosN(RstVent("idmon"))
    ' PREGUNTAMOS SI LA VENTA TIENE VENDEDOR ASIGNADO
    If RstVent("idven") <> 0 Then
        TxtIdVen.Text = NulosN(RstVent("idven"))
        ' MOSTRAMOS EL NOMBRE DEL VENDEDOR
        Set Rst = BuscaConCriterio("SELECT vta_vendedores.*, UCase(pla_empleados!apepat)+' '+ UCase(pla_empleados!apemat)+', '+pla_empleados!nom AS apenom " _
            & " FROM vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id WHERE vta_vendedores.id = " & NulosN(TxtIdVen.Text) & "", xCon)
        
        If Rst.RecordCount <> 0 Then
            LblNomVen.Caption = NulosC(Rst("apenom"))
        End If
    End If
    
    ' PREGUNTAMOS SI LA VENTA TIENE UN DOCUMENTO DE REFERENCIA ASIGNADO
    If NulosN(RstVent("idtipdocref")) <> 0 Then
        TxtIdTipDoc.Text = NulosC(RstVent("idtipdocref"))
        LblDescTipDocRef.Caption = Busca_Codigo(NulosC(RstVent("idtipdocref")), "id", "descripcion", "mae_docreferencia", "N", xCon)
        
        If NulosN(RstVent("idtipdocref")) = 4 Then ' buscamos el numero de documento de la orden de despacho
            RST_Busq Rst, "SELECT var_ordendespacho.id, var_ordendespacho.numerodoc AS numdoc " _
                & " From var_ordendespacho WHERE (((var_ordendespacho.id)=" & NulosN(RstVent("iddocref2")) & "))", xCon
        End If
        
        If NulosN(RstVent("idtipdocref")) = 5 Then ' buscamos el numero de documento del pedido
            RST_Busq Rst, "SELECT ped_pedido.id, [ped_pedido]![numser] & '-' & [ped_pedido]![numdoc] AS numdoc From ped_pedido " _
                & " WHERE (((ped_pedido.id)=" & NulosN(RstVent("iddocref2")) & "))", xCon
        End If
        If NulosN(RstVent("idtipdocref")) = 4 Or NulosN(RstVent("idtipdocref")) = 5 Then
            If Rst.State = 1 Then
                If Rst.RecordCount <> 0 Then
                    TxtNumDocRef.Text = NulosC(Rst("numdoc"))
                    LblIdDocRef2.Caption = Rst("id")
                End If
            End If
        End If
        Set Rst = Nothing
    End If
    
    '--uso temporal
    If Trim(TxtNumDocRef.Text) = "" Then TxtNumDocRef.Text = NulosC(RstVent("numerodocref"))
    
    ' MOSTRAMOS LOS TOTALES DEL DOCUMENTO
    TxtBruto.Text = Format(NulosN(RstVent("impbru")), FORMAT_MONTO)
    TxtIGV.Text = Format(NulosN(RstVent("impigv")), FORMAT_MONTO)
    TxtTotal.Text = Format(NulosN(RstVent("imptotdoc")), FORMAT_MONTO)
    txtinafecto.Text = Format(NulosN(RstVent("impinaf")), FORMAT_MONTO)
    txtisc.Text = Format(NulosN(RstVent("impisc")), FORMAT_MONTO)
    
    If NulosN(RstVent("idalm")) <> 0 Then
        TxtIdAlm.Text = Format(NulosN(RstVent("idalm")), "0")
    Else
        TxtIdAlm.Text = ""
    End If
    
    LblTipoItem.Caption = NulosC(RstVent("desctipcom"))
    LblNomDoc.Caption = NulosC(RstVent("nomdoc"))
    LblNomCli.Caption = NulosC(RstVent("nombre"))
    LblCondPag.Caption = NulosC(RstVent("desccond"))
    TxtNumRuc.Text = NulosC(RstVent("numruc"))
    LblMoneda.Caption = NulosC(RstVent("descmon"))
    LblIdAlmacen.Caption = NulosN(RstVent("idalm"))
    LblAlmacen.Caption = Busca_Codigo(RstVent("idalm"), "id", "descripcion", "alm_almacenes", "N", xCon)
                
    LblIdCliente.Caption = NulosN(RstVent("idcli"))
    xIdCuenTasa = NulosN(RstVent("idcuenvta"))


    ' Tipo de cambio
    If NulosN(RstVent("tc")) = 0 Then
        ChkTC.Value = 0
        TxtTC.Text = NulosN(RstVent("impven1"))
        TxtTC.BackColor = &H8000000F
        TxtTC.Enabled = False
    Else
        ChkTC.Value = 1
        TxtTC.Text = NulosN(RstVent("tc"))
        TxtTC.BackColor = vbWhite
        TxtTC.Enabled = True
    End If
    If QueHace = 3 Then TxtTC.BackColor = &H8000000F
    
    txtglosa.Text = NulosC(RstVent("glosa"))
    
    '---------------------------------------------------
    
    If RstVent("oriitem") = 1 Then optsinguia.Value = True: optsinguia_Click
    If RstVent("oriitem") = 2 Then optconguia.Value = True: optconguia_Click
    If RstVent("oriitem") = 3 Then optconcotizacion.Value = True: optconcotizacion_Click
    
    ' MOSTRAMOS EL TIPO DE DESCUENTO APLICADO
    If RstVent("tipdes") = 1 Or NulosN(RstVent("tipdes")) = 0 Then OptDes1.Value = True
    If RstVent("tipdes") = 2 Then OptDes2.Value = True
    
    ' cargamos las entregas del pedido
    If NulosN(TxtIdTipDoc.Text) = "5" Then
        ConfigurarParaPedido
        Dim Rst3 As New ADODB.Recordset
        RST_Busq Rst3, "SELECT DISTINCT ped_pedidodetent.idped, ped_pedidodetent.idtipdoc, ped_pedidodetent.iddocven, [ped_pedido]![numser] & '-' & [ped_pedido]![numdoc] AS numdoc, " _
            & " mae_documento.abrev, ped_pedidodetent.fchent FROM mae_documento RIGHT JOIN (ped_pedidodetent LEFT JOIN ped_pedido ON ped_pedidodetent.idped = ped_pedido.id) " _
            & " ON mae_documento.id = ped_pedido.tipdoc WHERE (((ped_pedidodetent.idtipdoc)=2) AND ((ped_pedidodetent.iddocven)=" & RstVent("id") & "))", xCon
        
        VAR_IDPEDIDO = Rst3("idped")
        VAR_FECHAPEDIDO = Rst3("fchent")

        If Rst3.RecordCount <> 0 Then
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(1, 1) = Rst3("numdoc")
            Fg4.TextMatrix(1, 2) = Rst3("abrev")
            Fg4.TextMatrix(1, 3) = Rst3("fchent")
        End If
        Set Rst3 = Nothing
    End If
    
    Frame5.Visible = False
    
    TxtIdTipDoc.Visible = True
    LblDescTipDocRef.Visible = True
    CmdBusIdTipDocRef.Visible = True
    
    TxtNumDocRef.Visible = True
    CmdBusDocRef2.Visible = True
    Label3(9).Visible = True
    Label3(8).Visible = True
    
    ' PREGUNTAMOS SI ES UNA NOTA DE CREDITO O UNA NOTA DE DEBITO
    'If (NulosN(RstVent("tipdoc")) = 7 And NulosN(RstVent("iddocref")) <> 0) Or (NulosN(RstVent("tipdoc")) = 8 And NulosN(RstVent("iddocref")) <> 0) Then
    If (NulosN(RstVent("tipdoc")) = 7) Or (NulosN(RstVent("tipdoc")) = 8) Then
        TxtIdTipDoc.Visible = False
        LblDescTipDocRef.Visible = False
        CmdBusIdTipDocRef.Visible = False
        
        TxtNumDocRef.Visible = False
        CmdBusDocRef2.Visible = False
        Label3(9).Visible = False
        Label3(8).Visible = False
        
        Frame3.Visible = False
        Frame5.Left = 4905
        Frame5.Top = 2790
        Frame5.Visible = True
        If NulosN(RstVent("idmotnotcre")) <> 0 Then
            LblIdConNC.Caption = NulosN(RstVent("idmotnotcre"))
            TxtDocRef.Text = Busca_Codigo(NulosN(RstVent("idmotnotcre")), "id", "descripcion", "vta_conceptonc", "N", xCon)
            
            If NulosN(RstVent("idmotnotcre")) = 4 Then ' Devolucion
                txtMotDev.Visible = True
                cmdMotDev.Visible = True
                txtMotDev.Text = Busca_Codigo(NulosN(RstVent("idmotdev")), "id", "descripcion", "mae_motivodevolucion", "N", xCon)
                lblIdMotDev.Caption = NulosN(RstVent("idmotdev"))
                
                If NulosN(RstVent("idmotdev")) = 16 Then ' Otros
                    txtMotDevOtr.Visible = True
                    txtMotDevOtr.Text = NulosC(RstVent("desmotdev"))
                Else
                    txtMotDevOtr.Visible = False
                End If
            Else
                txtMotDev.Visible = False
                cmdMotDev.Visible = False
                txtMotDevOtr.Visible = False
            End If
        End If
        
        ' LA NOTA DE CREDITO HACE REFERENCIA A UNA FACTURA
        TxtDocRefCredi.Visible = True
        Label33.Visible = True
        CmdBusDocRef.Visible = True
        
        LblIdDocRef.Caption = NulosN(RstVent("iddocref"))
        TxtDocRefCredi.Text = Busca_Codigo(NulosN(RstVent("iddocref")), "id", "numser", "vta_ventas", "N", xCon) + "-" + Busca_Codigo(RstVent("iddocref"), "id", "numdoc", "vta_ventas", "N", xCon)
    Else
        TxtDocRefCredi.Visible = False
        Label33.Visible = False
        CmdBusDocRef.Visible = False
    End If
       
    
    Dim RstDet As New ADODB.Recordset
    Dim A As Integer

    ' CARGAMOS LAS GUIAS DE LAS FACTURAS
    If optconguia.Value = True Then
        RST_Busq RstDet, "SELECT vta_guia.id, mae_documento.abrev, [vta_guia]![numser]+'-'+[vta_guia]![numdoc] AS numdoc" _
            & " FROM vta_guia LEFT JOIN mae_documento ON vta_guia.tipdoc = mae_documento.id " _
            & " WHERE (((vta_guia.iddocven)=" & RstVent("id") & "))", xCon
        If RstDet.RecordCount <> 0 Then
            
            RstDet.MoveFirst
            For A = 1 To RstDet.RecordCount
                Fg4.Rows = Fg4.Rows + 1
                Fg4.TextMatrix(A, 1) = NulosC(RstDet("numdoc"))
                Fg4.TextMatrix(A, 2) = NulosC(RstDet("abrev"))
                Fg4.TextMatrix(A, 3) = RstDet("id")
                
                RstDet.MoveNext
                If RstDet.EOF = True Then
                    Exit For
                End If
            Next A
        End If
        Set RstDet = Nothing
    End If
     
    ' CARGAMOS LOS ITEMS DE LA FACTURA
    Set RstDet = Nothing
    Mostrando = True

    RST_Busq RstDet, "SELECT vta_ventasdet.*, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.idcuentaven, " _
        & " alm_inventario.idtipven, alm_inventario.stckact FROM mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN vta_ventasdet " _
        & " ON alm_inventario.id = vta_ventasdet.iditem) ON mae_unidades.id = alm_inventario.idunimed " _
        & " WHERE (((vta_ventasdet.idvta)=" & RstVent("id") & "))", xCon
    
    If RstDet.RecordCount <> 0 Then
        Do While Not RstDet.EOF
            Fg1.Rows = Fg1.Rows + 1
            '*************************
            Dim RstItemDoc As New ADODB.Recordset
            Dim mDetalleItem As String
            Dim F As New SistemaLogica.Funciones
            ' Validamos si esta configurado el item para mostrar el nombre tecnico
            RST_Busq RstItemDoc, "SELECT mae_itemdocconfig.* From mae_itemdocconfig WHERE (((mae_itemdocconfig.iditem)=" & NulosN(RstDet("iditem")) & ") AND ((mae_itemdocconfig.iddoc)=" & NulosN(TxtTipDoc.Text) & "))", xCon
            If RstItemDoc.RecordCount > 0 Then
                mDetalleItem = F.BuscaCodigoTabla(NulosN(RstDet("iditem")), "id", "desctec", "alm_inventario", "N", xCon)
            Else
                mDetalleItem = NulosC(RstDet("descripcion"))
            End If
            '*************************
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = mDetalleItem
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(RstDet("abrev"))
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(RstDet("canpro"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(RstDet("preunibru"), "0.000000")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(RstDet("valdes"), "0.0000")
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Format(RstDet("preuni"), "0.000000")
            Fg1.TextMatrix(Fg1.Rows - 1, 7) = Format(RstDet("imptot"), "0.0000")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = NulosN(RstDet("iditem"))
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosN(RstDet("idunimed"))
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosN(RstDet("idcuentaven"))
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(RstDet("idtipven"))
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosN(RstDet("stckact") + RstDet("canpro"))
            RstDet.MoveNext
        Loop
    End If
    
    Set RstDet = Nothing
    Mostrando = False
    
    Set RstDet = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & NulosN(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & NulosN(TxtIdMon) & " and tipope = -1", xCon)
    If RstDet.RecordCount = 1 Then
        xCuentaDoc = RstDet("idcuen")
    End If
    
    Set RstDet = BuscaConCriterio("SELECT mae_impuestos.tasa from mae_impuestos WHERE mae_impuestos.id = 1 ", xCon)
    If RstDet.RecordCount = 1 Then
        TasaImpuesto = NulosN(RstDet("tasa"))
        LblIgvTasa.Caption = Format(Trim(Str(TasaImpuesto)), "0.00")
    End If
    
    ' CARGAMOS EL DETALLE DE CADA ITEMS
    Set RstVentItemsDeta = Nothing
    
    RST_Busq RstVentItemsDeta, "SELECT vta_ventasdetitems.idventa, vta_ventasdetitems.iditem, vta_ventasdetitems.orden, vta_ventasdetitems.texto " _
        & " From vta_ventasdetitems Where (((vta_ventasdetitems.idventa) = " & RstVent("id") & ")) ORDER BY vta_ventasdetitems.iditem, vta_ventasdetitems.orden", xCon
    
    Set RstVentItemsDeta.ActiveConnection = Nothing
    
    pGridConfigurar
    Set RstDet = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : Bloquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA O DESACTIVA LOS CONTROLES TextBox DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Bloquea()
    TxtTipItem.Locked = Not TxtTipItem.Locked
    TxtTipDoc.Locked = Not TxtTipDoc.Locked
    TxtNumRuc.Locked = Not TxtNumRuc.Locked
    TxtNumSer.Locked = Not TxtNumSer.Locked
    TxtNumDoc.Locked = Not TxtNumDoc.Locked
    ChkTC.Enabled = Not ChkTC.Enabled
    TxtFchDoc.Locked = Not TxtFchDoc.Locked
    TxtFchVen.Locked = Not TxtFchVen.Locked
    TxtConPag.Locked = Not TxtConPag.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtIdAlm.Locked = Not TxtIdAlm.Locked
    
    Frame3.Enabled = Not Frame3.Enabled
    CmdAddItem.Enabled = Not CmdAddItem.Enabled
    CmdDelItem.Enabled = Not CmdDelItem.Enabled
    'CmdSel.Enabled = Not CmdSel.Enabled
    CmdPreHist.Enabled = Not CmdPreHist.Enabled
    
    optsinguia.Enabled = Not optsinguia.Enabled
    optconguia.Enabled = Not optconguia.Enabled
    optconcotizacion.Enabled = Not optconcotizacion.Enabled
    
    TxtDocRef.Locked = Not TxtDocRef.Locked
    TxtIdTipDoc.Locked = Not TxtIdTipDoc.Locked
    TxtNumDocRef.Locked = Not TxtNumDocRef.Locked
    
    TxtIdVen.Locked = Not TxtIdVen.Locked
    
    TxtTC.BackColor = &H8000000F
    txtglosa.Locked = Not txtglosa.Locked
    
    TxtIdTipDoc.Visible = True
    LblDescTipDocRef.Visible = True
    CmdBusIdTipDocRef.Visible = True
    
    TxtNumDocRef.Visible = True
    CmdBusDocRef2.Visible = True
    Label3(9).Visible = True
    Label3(8).Visible = True
End Sub

'*****************************************************************************************************
'* Nombre           : Blanquea
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : BLANQUE LOS CONTROLES TextBox DEL FORMULARIO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Blanquea()
    lblReg.Caption = ""
    
    TxtTipItem.Text = ""
    TxtTipDoc.Text = ""
    TxtNumRuc.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtFchDoc.Valor = ""
    
    TxtFchVen.Valor = ""
    TxtConPag.Text = ""
    TxtIdMon.Text = ""
    TxtDocRef.Text = ""
    TxtDocRefCredi.Text = ""
    
    LblNomDoc.Caption = ""
    LblNomCli.Caption = ""
    LblCondPag.Caption = ""
    LblMoneda.Caption = ""
    LblIdCliente.Caption = ""
    LblTipoItem.Caption = ""
    TxtIdVen = ""
    LblNomVen = ""
    TxtIdAlm.Text = ""
    LblAlmacen.Caption = ""
    
    txtinafecto = ""
    txtisc = ""
    TxtBruto.Text = ""
    TxtIGV.Text = ""
    TxtTotal.Text = ""

    TxtDocRef.Text = ""
    TxtIdTipDoc.Text = ""
    TxtNumDocRef.Text = ""
        
    LblDescTipDocRef.Caption = ""
    LblIdDocRef.Caption = ""
    LblIdDocRef2.Caption = ""
    
    ChkTC.Value = 0
    TxtTC.Text = ""
    
    txtglosa.Text = ""
    
    Fg4.Rows = 1
    Fg1.Rows = 1
End Sub

Private Sub cbMotDev_Click()
End Sub

Private Sub Check1_Click()
    ' INDICA SI SE INGRESARA EL VALOR NETO DEL ITEM PARA ACONDICIONAR EL CONTROL FlexGrid Fg1
    If Check1.Value = 1 Then
        Fg1.ColWidth(14) = 1005
        Fg1.ColWidth(15) = 705
        If optconguia = True Then
            Fg1.ColWidth(1) = 2900 '5400 - 1710
        End If
        If optsinguia = True Then
            Fg1.ColWidth(1) = 5400 - 1710
        End If
    Else
        Fg1.ColWidth(14) = 0
        Fg1.ColWidth(15) = 0
        If optconguia = True Then
            Fg1.ColWidth(1) = 3500
        End If
        If optsinguia = True Then
            Fg1.ColWidth(1) = 5400
        End If
    End If
End Sub

Private Sub CmdAddItem_Click()
    ' PERMITE AGREGAR UNA FILA AL CONTROL FlexGrid Fg1
    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Sub
    End If
    
    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtTipItem.Text) = "" Then
        MsgBox "No ha especificado el tipo de item a buscar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipItem.SetFocus
        Exit Sub
    End If
    
    If NulosC(Fg1.TextMatrix(Fg1.Rows - 1, 1)) = "" Then
        Fg1.Col = 1
        Fg1.Row = Fg1.Rows - 1
        Fg1_CellButtonClick Fg1.Rows - 1, 1
        Fg1.SetFocus
        Exit Sub
    End If
    
    Fg1.Rows = Fg1.Rows + 1
    ' agregando cantidad por defecto a 1 cuando es servcio
    If NulosN(TxtTipItem.Text) = 5 Then
        Fg1.TextMatrix(Fg1.Rows - 1, 3) = 1
    End If
    
    With Fg1
        .Select Fg1.Rows - 1, 1, Fg1.Rows - 1, 1
    End With
    
    Fg1_CellButtonClick Fg1.Rows - 1, 1
    
    Fg1.SetFocus
End Sub

Private Sub cmdagregardocs_Click()
    ' PERMIRE AGREGAR DOCUMENTOS RELACIONADOS A LA VENTA
    
    ' VERIFICA QUE SE HAYAN INGRESADO ALGUNOS DATOS NECESARIO PARA EJECUTAR ESTE PROCESO
    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Sub
    End If
    
    If NulosC(LblIdCliente.Caption) = "" Then
        MsgBox "No ha especificado el cliente", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Sub
    End If
    
    If optconguia.Value = True Then
        ' SI SE SELECCIONA LA OPCION GUIA, SE CARGARAN LAS GUIAS EMITIDAS AL CLIENTE
        CargarGuia
    End If
    If optconcotizacion.Value = True Then
        ' SE SE SELECCIONA LA OPCION ORDEN DE PEDIDO, SE CARGARAN LAS ORDENES DE PEDIDO DEL CLIENTE
        If NulosC(TxtNumDocRef.Text) = "" Then
            MsgBox "No ha especificado el Numero de documento de referencia para la " & LblDescTipDocRef.Caption, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtNumDocRef.SetFocus
            Exit Sub
        End If
        CargarCotizacion
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : CargarCotizacion
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITE CARGAR LAS ORDENES DE PEDIDO DE UN CLIENTE
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub CargarCotizacion()
    Dim xfrm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(4, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Fc. Entrega":      xCampos(0, 1) = "fchent":       xCampos(0, 2) = "1200":         xCampos(0, 3) = "C":
    xCampos(1, 0) = "Descripcion":      xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "3500":         xCampos(1, 3) = "C":
    xCampos(2, 0) = "Uni. Med":         xCampos(2, 1) = "abreunimed":   xCampos(2, 2) = "1000":         xCampos(2, 3) = "C":
    xCampos(3, 0) = "Cantidad":         xCampos(3, 1) = "canpro":       xCampos(3, 2) = "1200":         xCampos(3, 3) = "N":
    
    
    'CARGAMOS LOS PEDIDOS PENDIENTES
    xfrm.SQLCad = "SELECT DISTINCT ped_pedido.id, ped_pedido.idcli, ped_pedidodetent.fchent, alm_inventario.descripcion, mae_unidades.abrev AS abreunimed, " _
        & " ped_pedidodetent.canpro, [ped_pedido]![numser] & '-' & [ped_pedido]![numdoc] AS numdoc, mae_documento.abrev AS abredoc" _
        & " FROM (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN ped_pedido ON mae_cliente.id = ped_pedido.idcli) " _
        & " ON mae_documento.id = ped_pedido.tipdoc) RIGHT JOIN (ped_pedidodet RIGHT JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN ped_pedidodetent " _
        & " ON alm_inventario.id = ped_pedidodetent.iditem) ON mae_unidades.id = ped_pedidodetent.idunimed) ON (ped_pedidodet.idped = ped_pedidodetent.idped) " _
        & " AND (ped_pedidodet.iditem = ped_pedidodetent.iditem)) ON ped_pedido.id = ped_pedidodet.idped " _
        & " WHERE (((ped_pedido.id)=" & NulosN(LblIdDocRef2.Caption) & ") AND ((ped_pedido.idcli)=" & NulosN(LblIdCliente.Caption) & ") AND ((ped_pedidodetent.estado)=2))"
    
    xfrm.titulo = "Entregas de la Orden de Pedido"
    xfrm.FormaBusca = Principio
    xfrm.Criterio = ""
    xfrm.Ordenado = "fchent"
    xfrm.CampoBusca = "fchent"
    Set xfrm.Coneccion = xCon
    Set xRs = xfrm.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        Dim Rst2 As New ADODB.Recordset
        Dim A As Integer
        ConfigurarParaPedido
        ' CARGAMOS LOS ITEMS DE LOS PEDIDOS
        RST_Busq Rst2, "SELECT ped_pedido.id, ped_pedido.idcli, mae_documento.abrev, ped_pedido!numser & '-' & ped_pedido!numdoc AS numdoc, mae_cliente.nombre, " _
            & " ped_pedidodetent.iditem, ped_pedidodetent.idunimed, alm_inventario.idcuentaven, alm_inventario.idtipven, alm_inventario.descripcion, " _
            & " ped_pedidodetent.fchent, mae_unidades.abrev AS unimed, ped_pedidodetent.canpro, 9999.99 AS stckact, " _
            & " (SELECT Max(vta_ventasdet.preunibru) AS MáxDepreunibru" _
            & " From vta_ventasdet GROUP BY vta_ventasdet.iditem HAVING (((vta_ventasdet.iditem)=ped_pedidodetent.iditem))) " _
            & " AS preven " _
            & " FROM (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN ped_pedido ON mae_cliente.id = ped_pedido.idcli) " _
            & " ON mae_documento.id = ped_pedido.tipdoc) INNER JOIN (ped_pedidodet INNER JOIN (mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN ped_pedidodetent " _
            & "  ON alm_inventario.id = ped_pedidodetent.iditem) ON mae_unidades.id = ped_pedidodetent.idunimed) ON (ped_pedidodet.idped = ped_pedidodetent.idped) " _
            & " AND (ped_pedidodet.iditem = ped_pedidodetent.iditem)) ON ped_pedido.id = ped_pedidodet.idped " _
            & " WHERE (((ped_pedido.id)=" & LblIdDocRef2.Caption & ") AND ((ped_pedido.idcli)=" & LblIdCliente.Caption & ") AND ((ped_pedidodetent.fchent)=CDate('" & xRs("fchent") & "')) " _
            & " AND ((ped_pedidodetent.estado)=2)) ORDER BY ped_pedidodetent.fchent", xCon
        If Rst2.RecordCount <> 0 Then
            VAR_IDPEDIDO = Rst2("id")
            VAR_FECHAPEDIDO = Rst2("fchent")
            
            Rst2.MoveFirst
            Fg1.Rows = 1
            
            ' MOSTRAMOS LOS ITEMS EN EL CONTROL FlexGrid Fg1
            For A = 1 To Rst2.RecordCount
                Fg1.Rows = Fg1.Rows + 1
                Fg1.TextMatrix(A, 1) = Rst2("descripcion")
                Fg1.TextMatrix(A, 2) = Rst2("unimed")
                Fg1.TextMatrix(A, 3) = Format(Rst2("canpro"), "0.00")
                
                Fg1.TextMatrix(A, 8) = Rst2("iditem")
                Fg1.TextMatrix(A, 9) = Rst2("idunimed")
                Fg1.TextMatrix(A, 10) = Rst2("idcuentaven")
                Fg1.TextMatrix(A, 11) = Rst2("idtipven")
                Fg1.TextMatrix(A, 13) = Rst2("stckact")
                Fg1.TextMatrix(A, 4) = Format(Rst2("preven"), "0.00")
                Rst2.MoveNext
                If Rst2.EOF = True Then Exit For
            Next A
            
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(1, 1) = xRs("numdoc")
            Fg4.TextMatrix(1, 2) = xRs("abredoc")
            Fg4.TextMatrix(1, 3) = xRs("fchent")
        End If
    End If
    Set xfrm = Nothing
    Set xRs = Nothing
        
End Sub

'*****************************************************************************************************
'* Nombre           : ConfigurarParaPedido
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CONFIGURA PARA PEDIDO EL FlexGrid Fg4
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ConfigurarParaPedido()
    Fg4.Rows = 1
    Fg4.ColWidth(3) = 1000
    Fg4.TextMatrix(0, 3) = "Fch. Ent."
End Sub

'*****************************************************************************************************
'* Nombre           : MostrarItemsCotizacion
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LOS ITEMS ASIGNADOS A UNA COTIZACION
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MostrarItemsCotizacion()
    Dim A As Integer
    Dim xCadWhere  As String
    Dim Rst As New ADODB.Recordset
    xCadWhere = ""
    
    If Fg4.Rows = 1 Then
        Fg1.Rows = 1
        Exit Sub
    End If
    'CREAMOS LA SENTENCIA WHERE PARA LA CONSULTA SQL
    For A = 1 To Fg4.Rows - 1
        xCadWhere = xCadWhere + "(vta_cotizaciondet.idvta = " & Fg4.TextMatrix(A, 3) & ")"
        If A = Fg4.Rows - 1 Then Exit For
        xCadWhere = xCadWhere + " OR "
    Next A
    
    RST_Busq Rst, "SELECT vta_cotizaciondet.iditem AS id, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, Sum(vta_cotizaciondet.canpro) AS campro, " _
        & " vta_cotizaciondet.iditem, alm_inventario.idtipven, alm_inventario.idcuentaven, vta_cotizaciondet.preuni " _
        & " FROM vta_cotizacion LEFT JOIN ((vta_cotizaciondet LEFT JOIN alm_inventario ON vta_cotizaciondet.iditem = alm_inventario.id) LEFT JOIN mae_unidades " _
        & " ON vta_cotizaciondet.idunimed = mae_unidades.id) ON vta_cotizacion.id = vta_cotizaciondet.idvta " _
        & " Where " + xCadWhere _
        & " GROUP BY vta_cotizaciondet.iditem, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.idtipven, " _
        & " alm_inventario.idcuentaven, vta_cotizaciondet.preuni", xCon

    Fg1.Rows = 1
    
    Agregando = True
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Busca_Codigo(Rst("id"), "id", "stckact", "alm_inventario", "N", xCon)
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("descripcion") 'DESCRIPCION DEL PRODUCTO
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("abrev")       'ABREVIATURA
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(Rst("preuni"), "0.00")   'PRECIO UNITARIO
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Rst("campro")
            Fg1.TextMatrix(Fg1.Rows - 1, 5) = Format(Rst("preuni") * Rst("campro"), "0.00")
            Fg1.TextMatrix(Fg1.Rows - 1, 6) = Rst("id")
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Rst("idcuentaven")
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Rst("idtipven")
            
            Rst.MoveNext
            If Rst.EOF = True Then
                Exit For
            End If
        Next A
    End If
    Agregando = False
    HallarTotal
End Sub

'*****************************************************************************************************
'* Nombre           : CargarGuia
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PERMITE ASIGNAR GUIAS AL DOCUMENTO DE VENTA
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

    xfrm.SQLCad = "SELECT 0 as xSel, vta_guia.id, vta_guia.fecgiro, [vta_guia]![numser]+'-'+[vta_guia]![numdoc] AS NroDoc, mae_cliente.numruc, mae_cliente.nombre, " _
        & " mae_mottra.descripcion, vta_guia.idcli, mae_documento.abrev FROM mae_mottra RIGHT JOIN ((mae_cliente RIGHT JOIN vta_guia ON " _
        & " mae_cliente.id = vta_guia.idcli) LEFT JOIN mae_documento ON vta_guia.tipdoc = mae_documento.id) ON mae_mottra.id = vta_guia.idmottra " _
        & " WHERE (((vta_guia.idcli)=" & NulosN(LblIdCliente.Caption) & ") AND ((vta_guia.Anulado)=0) AND ((vta_guia.iddocven)=0)) " _
        & " ORDER BY [vta_guia]![numser]+'-'+[vta_guia]![numdoc] DESC"
        
    xfrm.titulo = "Buscando Guias del Guias"
    
    Set xfrm.Coneccion = xCon
    Set xRs = xfrm.Seleccionar(xCampos)
    
    Fg4.Rows = 1
    If xRs.State = 1 Then
        Fg4.ColWidth(3) = 0
        Fg4.TextMatrix(0, 3) = "iddoc"
    
        If xRs.RecordCount = 0 Then
            Set xRs = Nothing
            Exit Sub
        End If
        Dim xCadWhere As String
        Dim A As Integer
        Dim Rst As New ADODB.Recordset
        
        xRs.MoveFirst
        
        'CARGAMOS LOS DOCUMENTOS ADJUNTOS Y LO MOSTRAMOS EN LA LISTA DE "DOCUMENTOS ADJUNTOS"
        For A = 1 To xRs.RecordCount
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(A, 1) = xRs("nrodoc")
            Fg4.TextMatrix(A, 2) = xRs("abrev")
            Fg4.TextMatrix(A, 3) = xRs("id")
            xRs.MoveNext
            If xRs.EOF = True Then Exit For
        Next A
        
        MostrarItems
        Agregando = False
    End If
    
    HallarTotal
    Set xfrm = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : MostrarItems
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS ITEMS DE LAS GUIAS ASIGNADAS AL DOCUMENTO DE VENTA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub MostrarItems()
    Dim A As Integer
    Dim xCadWhere  As String
    Dim Rst As New ADODB.Recordset
    xCadWhere = ""
    
    If Fg4.Rows = 1 Then
        Fg1.Rows = 1
        Exit Sub
    End If
    
    ' CREAMOS LA SENTENCIA WHERE PARA LA CONSULTA SQL
    For A = 1 To Fg4.Rows - 1
        xCadWhere = xCadWhere + "(vta_guiadet.idgui=" & Fg4.TextMatrix(A, 3) & ")"
        If A = Fg4.Rows - 1 Then Exit For
        xCadWhere = xCadWhere + " OR "
    Next A
    
    ' CARGAMOS LOS ITEMS
    RST_Busq Rst, "SELECT alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, Sum(vta_guiadet.canpro) AS SumaDecanpro, alm_inventario.id, " _
        & " alm_inventario.idtipven, alm_inventario.idcuentaven, alm_inventario.stckact, alm_inventario.idunimed FROM (vta_guiadet LEFT JOIN alm_inventario " _
        & " ON vta_guiadet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON vta_guiadet.idunimed = mae_unidades.id " _
        & " WHERE " & Trim(xCadWhere) _
        & " GROUP BY alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.id, alm_inventario.idtipven, alm_inventario.idcuentaven, " _
        & " alm_inventario.stckact, alm_inventario.idunimed ORDER BY alm_inventario.id", xCon
    
    Fg1.Rows = 1
    If Rst.RecordCount <> 0 Then
        Rst.MoveFirst
        Dim xPrecio As Double
        Agregando = True
        ' MOSTRAMOS LOS ITEMS CARGADOS EN EL CONTROL FlexGrid Fg1
        For A = 1 To Rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(Fg1.Rows - 1, 1) = Rst("descripcion")                  ' DESCRIPCION DEL PRODUCTO
            Fg1.TextMatrix(Fg1.Rows - 1, 2) = Rst("abrev")                        ' ABREVIATURA
            Fg1.TextMatrix(Fg1.Rows - 1, 3) = Format(Rst("SumaDecanpro"), "0.00") ' CANTIDAD
            Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(1, "0.00")                   ' PRECIO UNITARIO
            Fg1.TextMatrix(Fg1.Rows - 1, 8) = Rst("id")
            Fg1.TextMatrix(Fg1.Rows - 1, 9) = Rst("idunimed")
            Fg1.TextMatrix(Fg1.Rows - 1, 10) = Rst("idcuentaven")
            Fg1.TextMatrix(Fg1.Rows - 1, 11) = Rst("idtipven")
            Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosN(Rst("stckact"))
            
            If NulosN(LblIdCliente.Caption) <> 0 Then
                xPrecio = UltimoPrecio(NulosN(Rst("id")), NulosN(LblIdCliente.Caption))
            Else
                xPrecio = UltimoPrecio(NulosN(Rst("id")), 0)
            End If
            
            Fg1.TextMatrix(A, 4) = Format((xPrecio), "0.000000")
            Fg1.TextMatrix(A, 6) = Format((xPrecio), "0.000000")
            Fg1.TextMatrix(A, 7) = (xPrecio * NulosN(Fg1.TextMatrix(A, 3)))
            Fg1.TextMatrix(A, 7) = Format(Fg1.TextMatrix(A, 7), "0.00")
            
            Rst.MoveNext
            If Rst.EOF = True Then Exit For
        Next A
     Agregando = False
    End If
End Sub

Private Sub CmdApertura_Click()
    AperturaDocumento xCon, xIdUsuario, 2, IdMenuActivo
    ' refrescar la consulta
    RstVent.Filter = ""
    TDB_FiltroLimpiar Dg1
    RstVent.Requery
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
        LblAlmacen.Caption = xRs("descripcion")
        TxtIdAlm.Text = xRs("id")
        TxtTipDoc.SetFocus
        
        If TxtTipDoc.Text <> "" Then
            Dim Rst As New ADODB.Recordset
            Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc.Text) & " AND idalm = " & NulosN(TxtIdAlm.Text) & "", xCon)
            If Rst.RecordCount <> 0 Then
                TxtNumSer.Text = NulosC(Rst("numser"))
                TxtNumSer_Validate True
            End If
            
            Set Rst = Nothing
        Else
            TxtNumSer.Text = ""
            TxtNumDoc.Text = ""
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusCondicion_Click()
    ' EJECUTA LA BUSQUEDA DE CONDICIONES DE PAGO
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripción":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Código":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_condpago ORDER BY descripcion"
    
    xform.titulo = "Buscando Condición de Pago"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtConPag.Text = xRs("id")
            LblCondPag.Caption = NulosC(xRs("descripcion"))
            If NulosC(TxtFchDoc.Valor) <> "" Then
                TxtFchVen.Valor = CDate(TxtFchDoc.Valor) + NulosN(xRs("numdia"))
            End If
            TxtFchVen.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDocRef_Click()
    ' EJECUTA LA BUSQUEDA DEL DOCUMENTO DE REFERENCIA
    If QueHace = 3 Then Exit Sub

    If NulosN(LblIdCliente.Caption) = 0 Then
        MsgBox "No ha especificado el cliente para referenciar este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(8, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "T.D.":       xCampos(0, 1) = "abrev":                xCampos(0, 2) = "450":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Fch. Doc.":        xCampos(1, 1) = "fchdoc":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Nº Documento":     xCampos(2, 1) = "numdoc":               xCampos(2, 2) = "1500":         xCampos(2, 3) = "C"
    xCampos(3, 0) = "Fch. Ven.":        xCampos(3, 1) = "fchven":               xCampos(3, 2) = "1000":         xCampos(3, 3) = "C"
    xCampos(4, 0) = "M":                xCampos(4, 1) = "simbolo":              xCampos(4, 2) = "450":         xCampos(4, 3) = "C"
    xCampos(5, 0) = "T.C.":             xCampos(5, 1) = "tipcam":               xCampos(5, 2) = "700":         xCampos(5, 3) = "N"
    xCampos(6, 0) = "Total":            xCampos(6, 1) = "imptotdoc":            xCampos(6, 2) = "1000":         xCampos(6, 3) = "N"
    xCampos(7, 0) = "Condición":        xCampos(7, 1) = "descripcion":          xCampos(7, 2) = "1000":         xCampos(7, 3) = "C"
    
    xform.SQLCad = "SELECT mae_documento.abrev, vta_ventas.fchdoc, [vta_ventas]![numser] & '-' & [vta_ventas]![numdoc] AS numdoc, vta_ventas.fchven, " _
        & " mae_cliente.nombre, mae_condpago.descripcion, vta_ventas.id, vta_ventas.imptotdoc, vta_ventas.idcli, vta_ventas.tipdoc, " _
        & " IIf([vta_ventas].[tc]=0,[con_tc].[impven],[vta_ventas].[tc]) AS tipcam, mae_moneda.simbolo " _
        & " FROM ((((mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_condpago ON vta_ventas.idconpag = mae_condpago.id)  " _
        & " LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id Where (((vta_ventas.idcli) = " & NulosN(LblIdCliente.Caption) & ") And ((vta_ventas.tipdoc) <> 7)) " _
        & " ORDER BY vta_ventas.fchdoc DESC"
    
    xform.titulo = "Buscando Documentos del Cliente"
    xform.FormaBusca = CualquierParte
    xform.Criterio = ""
    xform.Ordenado = "numdoc"
    xform.CampoBusca = "numdoc"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtDocRefCredi.Text = NulosC(xRs("numdoc"))
            LblIdDocRef.Caption = xRs("id")
            ' actualizando el tipo de cambio al de la emision
            If NulosN(xRs("tipcam")) <> 0 Then
                ChkTC.Value = 1
                TxtTC.Text = NulosN(xRs("tipcam"))
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusDocRef2_Click()
    ' EJECUTA LA BUSQUEDA DEL DOCUMENTO DE REFERENCIA
    If QueHace = 3 Then Exit Sub
    
    If NulosN(TxtIdTipDoc.Text) = 0 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(4, 4) As String
    
    
    If NulosN(TxtIdTipDoc.Text) = 4 Then
        'Orden de Despacho
        xCampos(0, 0) = "Nº Documento":      xCampos(0, 1) = "numdoc":      xCampos(0, 2) = "1500":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Fch. Emi.":         xCampos(1, 1) = "fchemi":      xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Fch. Ven.":         xCampos(2, 1) = "fchven":      xCampos(2, 2) = "1000":         xCampos(2, 3) = "C"
        xCampos(3, 0) = "Cliente":           xCampos(3, 1) = "nombre":      xCampos(3, 2) = "4000":         xCampos(3, 3) = "C"
        
        xform.SQLCad = "SELECT var_ordendespacho.id, var_ordendespacho.numerodoc AS numdoc, " _
            & " mae_cliente.nombre, var_ordendespacho.idcli, var_ordendespacho.fchemi, var_ordendespacho.fchven FROM var_ordendespacho LEFT JOIN mae_cliente " _
            & " ON var_ordendespacho.idcli = mae_cliente.id"
        
        
        xform.titulo = "Orden de Despacho"
        'Set xRs = xform.BuscarReg(xCampos)
        
    End If
    
    If NulosN(TxtIdTipDoc.Text) = 5 Then
        Dim xCampos2(7, 4) As String
        
        xCampos2(0, 0) = "Tipo Documento":    xCampos2(0, 1) = "abredoc":     xCampos2(0, 2) = "1000":         xCampos2(0, 3) = "C":
        xCampos2(1, 0) = "Nº Documento":      xCampos2(1, 1) = "numdoc":      xCampos2(1, 2) = "1400":         xCampos2(1, 3) = "C":
        xCampos2(2, 0) = "Fc. Emi.":          xCampos2(2, 1) = "fchemi":      xCampos2(2, 2) = "950":          xCampos2(2, 3) = "C":
        xCampos2(3, 0) = "Nº OC Cliente":     xCampos2(3, 1) = "oc":          xCampos2(3, 2) = "1300":         xCampos2(3, 3) = "C":
        xCampos2(4, 0) = "Producto":          xCampos2(4, 1) = "descripcion": xCampos2(4, 2) = "2500":         xCampos2(4, 3) = "C":
        xCampos2(5, 0) = "Uni. Med.":         xCampos2(5, 1) = "abreunimed":  xCampos2(5, 2) = "800":          xCampos2(5, 3) = "C":
        xCampos2(6, 0) = "Can. Pro":          xCampos2(6, 1) = "canpro":      xCampos2(6, 2) = "1200":         xCampos2(6, 3) = "N":
        
        ' CARGAMOS LOS PEDIDOS
        xform.SQLCad = "SELECT DISTINCT ped_pedido.id, ped_pedido.idcli, ped_pedido.fchemi, mae_documento.abrev AS abredoc, ped_pedido!numser & '-' & ped_pedido!numdoc AS numdoc, " _
            & " alm_inventario.descripcion, mae_unidades.abrev AS abreunimed, ped_pedidodet.canpro, ped_pedido.oc FROM (((mae_documento RIGHT JOIN (mae_cliente " _
            & " RIGHT JOIN ped_pedido ON mae_cliente.id = ped_pedido.idcli) ON mae_documento.id = ped_pedido.tipdoc) RIGHT JOIN ped_pedidodet ON ped_pedido.id = ped_pedidodet.idped) " _
            & " LEFT JOIN alm_inventario ON ped_pedidodet.iditem = alm_inventario.id) LEFT JOIN mae_unidades ON ped_pedidodet.idunimed = mae_unidades.id " _
            & " Where (((ped_pedido.idcli) = " & NulosN(LblIdCliente.Caption) & ") And ((ped_pedidodet.estado) <> 1)) ORDER BY ped_pedido.fchemi DESC"
    End If
    
    xform.FormaBusca = CualquierParte
    
    xform.Criterio = ""
    xform.Ordenado = "numdoc"
    xform.CampoBusca = "numdoc"
    Set xform.Coneccion = xCon
    If NulosN(TxtIdTipDoc.Text) = 4 Then Set xRs = xform.BuscarReg(xCampos)
    If NulosN(TxtIdTipDoc.Text) = 5 Then Set xRs = xform.BuscarReg(xCampos2)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumDocRef.Text = NulosC(xRs("numdoc"))
            LblIdDocRef2.Caption = xRs("id")
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusIdTipDocRef_Click()
    ' BUSCA EL TIPO DE DOCUMENTO DE REFERENCIA
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_docreferencia ORDER BY descripcion"
    
    xform.titulo = "Buscando Tipo de Documento de Referencia"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdTipDoc.Text = xRs("id")
            LblDescTipDocRef.Caption = xRs("descripcion")
            TxtNumDocRef.Text = ""
            LblIdDocRef2.Caption = ""
            Fg1.Rows = 1
            Fg4.Rows = 1
            
            TxtNumDocRef.SetFocus
            If xRs("id") = 5 Then
                optconguia.Enabled = False
                optsinguia.Enabled = False
                optconcotizacion.Enabled = True
                optconcotizacion.Value = True
                JALOPEDIDO = True
            Else
                TxtNumDocRef.Text = ""
                LblIdDocRef2.Caption = ""
                Fg1.Rows = 1
                optconguia.Enabled = True
                optsinguia.Enabled = True
                optconcotizacion.Enabled = False
                JALOPEDIDO = False
            End If
        End If
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusNumSer_Click()
    ' BUSCA EL NUMERO DE SERIE PARA EL TIPO DE DOCUENTOA ACTUAL
    
    ' VERIFICAMOS QUE LOS DATOS NECESARIOS PARA EJECUTAR EL PROCESO ESTEN INGRESADOS
    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Sub
    End If

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "iddoc":       xCampos(0, 2) = "1500":    xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion": xCampos(1, 2) = "2500":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Serie":          xCampos(2, 1) = "numser":      xCampos(2, 2) = "1500":    xCampos(2, 3) = "C"
    
    xform.SQLCad = "SELECT mae_documento.descripcion, mae_series.iddoc, Format([mae_series].[numser],'0000') AS numser, " _
        & " mae_series.numdoc FROM mae_documento INNER JOIN mae_series ON mae_documento.id = mae_series.iddoc " _
        & " WHERE (((mae_series.iddoc)=1))"
    
    xform.titulo = "Buscando Series"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numser"
    xform.CampoBusca = "numser"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumSer.Text = Format(xRs("numser"), "0000")
            TxtNumDoc = HallaNumdocVenta(NulosN(TxtTipDoc.Text), TxtNumSer.Text, xCon)
        End If
        TxtConPag.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusCli_Click()
    ' EJECUTA LA BUSQUEDA DE UN CLIENTE
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Cliente":      xCampos(0, 1) = "nombre":      xCampos(0, 2) = "5000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":    xCampos(1, 1) = "numruc":      xCampos(1, 2) = "2000":         xCampos(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_cliente.nombre, mae_cliente.numruc, mae_cliente.id, mae_cliente.idven From mae_cliente where mae_cliente.id <>0"
    
    xform.titulo = "Buscando Cliente"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "nombre"
    xform.CampoBusca = "nombre"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            If xRs.RecordCount <> 0 Then
                TxtNumRuc.Text = NulosC(xRs("numruc"))
                LblNomCli.Caption = NulosC(xRs("nombre"))
                LblIdCliente.Caption = xRs("id")
                TxtIdVen.Text = NulosN(xRs("idven"))
                TxtIdVen_Validate True
                TxtNumSer.SetFocus
            End If
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusMon_Click()
    ' EJECUTA LA BUSQUEDA DE LA MONEDA
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripción":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Código":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT * FROM mae_moneda ORDER BY descripcion"
    
    xform.titulo = "Buscando Moneda"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtIdMon.Text = xRs("id")
            LblMoneda.Caption = xRs("descripcion")
            Fg1.SetFocus
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipDoc_Click()
    ' EJECUTA LA BUSQUEDA DEL TIPO DE DOCUMENTO
    If QueHace = 3 Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, mae_impuestos.Abrev AS abreimp, " _
        & " mae_impuestos.idcuenvta AS cuentaimp, alm_numseries.numser FROM alm_numseries LEFT JOIN (mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas " _
        & " ON mae_impuestos.idcuenvta = con_planctas.id) ON mae_documento.idimp = mae_impuestos.id) ON alm_numseries.idtipdoc = mae_documento.id " _
        & " WHERE (((alm_numseries.idalm)=" & NulosN(TxtIdAlm.Text) & "))"

    Dim xImpuesto As Double
    
    xform.titulo = "Buscando Tipo de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumSer.Text = xRs("numser")
            TxtNumSer_Validate True
            TxtTipDoc.Text = xRs("id")
            LblNomDoc.Caption = xRs("descripcion")
            TasaImpuesto = NulosN(xRs("tasa"))
            
            xIdCuenTasa = NulosN(xRs("cuentaimp"))
            LblRotulo.Caption = Trim(NulosC(xRs("abreimp"))) + " (         )"
            LblIgvTasa.Caption = Format(Trim(Str(TasaImpuesto)), "0.00")
            FraRetencion.Caption = "( Afecta : " + NulosC(xRs("descimp")) + ")"
            
            Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & NulosN(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & NulosN(TxtIdMon) & " and tipope = -1", xCon)
            If xRs2.RecordCount > 0 Then
                xCuentaDoc = NulosN(xRs2("idcuen"))
            End If
            Set xRs2 = Nothing
            
            Frame5.Visible = False
            
            ' Si es Recibo por honorarios
            If xRs("id") = 2 Then
                 FraRetencion.Enabled = True
                 Fratipven.Enabled = False
                 FraRetencion.Visible = True
                 FraRetencion.Caption = "Retención de 4ta Categoria" & "10%"
                 txtisc.Enabled = False
                 txtinafecto.Enabled = False
            Else
                 Fratipven.Enabled = True
                 FraRetencion.Enabled = False
                 FraRetencion.Visible = False
                 txtisc.Enabled = True
                 txtinafecto.Enabled = True
            End If
            
            If xRs("id") = 7 Or xRs("id") = 8 Then
                Label33.Visible = True
                TxtDocRefCredi.Visible = True
                CmdBusDocRef.Visible = True
            Else
                Label33.Visible = False
                TxtDocRefCredi.Visible = False
                CmdBusDocRef.Visible = False
            End If
            TxtNumRuc.SetFocus
        End If
    
        If TxtTipDoc.Text <> "" Then
            Dim Rst As New ADODB.Recordset
            Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc.Text) & " AND idalm = " & NulosN(LblIdAlmacen.Caption) & "", xCon)
            If Rst.RecordCount <> 0 Then
                TxtNumSer.Text = Rst("numser")
                TxtNumSer_Validate True
            End If
            Set Rst = Nothing
        Else
            TxtNumSer.Text = ""
            TxtNumDoc.Text = ""
        End If
        
        TxtTipDoc_Validate False
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusTipItem_Click()
    ' EJECUTA LA BUSQUEDA TIPO DE ITEM
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
            TxtTipItem.Text = NulosN(xRs("id"))
            LblTipoItem = NulosC(xRs("descripcion"))
            TxtIdAlm.SetFocus
        End If
        
        pGridConfigurar
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdBusVen_Click()
    ' EJECUTA LA BUSQUEDA DE VENDEDORES
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(4, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Codigo":    xCampos(0, 1) = "id":         xCampos(0, 2) = "800":          xCampos(0, 3) = "N"
    xCampos(1, 0) = "Vendedor":  xCampos(1, 1) = "apenom":     xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
    xCampos(2, 0) = "Basico":    xCampos(2, 1) = "basico":     xCampos(2, 2) = "1200":         xCampos(2, 3) = "N"
    xCampos(3, 0) = "Comision":  xCampos(3, 1) = "comision":   xCampos(3, 2) = "1200":         xCampos(3, 3) = "N"
    
    xform.SQLCad = "SELECT vta_vendedores.*, UCase([pla_empleados]![apepat]) & UCase([pla_empleados]![apemat]) & ', ' & [pla_empleados]![nom] AS apenom " _
                & " FROM vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id"
    
    xform.titulo = "Buscando Vendedores"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "apenom"
    xform.CampoBusca = "apenom"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    
    If xRs.State = 1 Then
        LblNomVen.Caption = xRs("apenom")
        TxtIdVen.Text = xRs("id")
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdDelItem_Click()
    ' ELIMINA UNA FILA DEL CONTROL FlexGrid Fg1
    If QueHace = 3 Then Exit Sub
    If Fg1.Row < 1 Or Fg1.Rows < 1 Then Exit Sub
    
    Fg1.RemoveItem Fg1.Row
    HallarTotal
    If Fg1.Rows <> 1 Then Fg1.Select Fg1.Rows - 1, 1
End Sub

Private Sub CmdDetalleAceptar_Click()
    If QueHace = 3 Then
        FrameDetalleItem.Visible = True
    End If
    Dim A As Integer
    ' QUITAMOS EL FILTRO A LA TABL DE DETALLES
    RstVentItemsDeta.Filter = adFilterNone
    
    ' FILTRAMOS EL RECORDSERD SEGUN EL ITEM ACTUAL
    RstVentItemsDeta.Filter = "iditem = " & Fg1.TextMatrix(Fg1.Row, 8) & ""

    If RstVentItemsDeta.RecordCount <> 0 Then
        RstVentItemsDeta.MoveFirst
        For A = 1 To RstVentItemsDeta.RecordCount
            
            RstVentItemsDeta.Delete
            RstVentItemsDeta.MoveNext
            If RstVentItemsDeta.EOF = True Then Exit For
        Next A
    End If
    
    ' AGREGAMOS EL DETALLE DEL ITEM ACTUALIZADO
    For A = 0 To Fg3.Rows - 1
        If NulosC(Fg3.TextMatrix(A, 1)) <> "" Then
            RstVentItemsDeta.AddNew
            RstVentItemsDeta("idventa") = 9999
            RstVentItemsDeta("iditem") = NulosN(Fg1.TextMatrix(Fg1.Row, 8))
            RstVentItemsDeta("orden") = A
            RstVentItemsDeta("texto") = NulosC(Fg3.TextMatrix(A, 1))
        End If
    Next A
    FrameDetalleItem.Visible = False
End Sub

Private Sub cmdEliminarOKdocsproc_Click()
    ' ELIMINA LOS DOCUMENTOS ASIGNADOS A LA VENTA
    Dim Rstguia As New ADODB.Recordset
    Dim X As Integer

    If fgdocsproc.Rows - 1 > 0 Then
        If fgdocsproc.Rows - 1 = 1 Then
            fgdocsproc.Rows = 1
            Fg1.Rows = 1
            HallarTotal
            Exit Sub
        Else
            With Me.Fg1
                For X = 1 To Me.Fg1.Rows - 1
                    RST_Busq Rstguia, "Select Vta_GuiaDet.* From Vta_GuiaDet where Vta_GuiaDet.IdGui = " & NulosN(fgdocsproc.TextMatrix(fgdocsproc.Row, 1)) & " and Vta_GuiaDet.IdItem = " & NulosN(Fg1.TextMatrix(X, 6)) & "", xCon
                    
                    If Rstguia.RecordCount > 0 Then
                        .TextMatrix(X, 4) = NulosN(.TextMatrix(X, 4)) - Rstguia("canpro")
                    End If
                Next
                ' Colocamos en el campo estado 0  de la tabla guia que indica que no esta facturado
                xCon.Execute " UPDATE vta_guia SET Vta_guia.Estado = 0 WHERE vta_guia.id = " & NulosN(fgdocsproc.TextMatrix(fgdocsproc.Row, 1)) & ""
            End With
                
            fgdocsproc.RemoveItem fgdocsproc.Row
            HallarTotal
        End If
    End If
    Set Rstguia = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : CargarRSTCom
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS REGISTRO DE LA TABLA vta_ventas EN EL RECORDSET RstVent
'* Paranetros       : NOMBRE         |  TIPO       |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    xFechaRegistro |  STRING     |  ESPECIFICA LA FECHA DE REGISTRO
'*                    Mes            |  INTEGER    |  ESPECIFICA EL ID DEL MES
'* Devuelve         :
'*****************************************************************************************************
Sub CargarRSTCom(xFechaRegistro As String, Mes As Integer)
    Dim DiaIniAño As String
    DiaIniAño = "01/01/" + Trim(AnoTra)
    
    If mMesActivo >= 1 And mMesActivo <= 12 Then
        ' SI SE HA SELECCIONADO UN MES
        RST_Busq RstVent, "SELECT vta_ventas.*, IIf(vta_ventas.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, " _
            & " IIf(vta_ventas.Anulado=0,'Facturado','Anulado') AS EstadoVenta, mae_documento.descripcion AS nomdoc, IIf(vta_ventas.anulado=-1,'', mae_condpago.descripcion) AS desccond, " _
            & " mae_documento.abrev, mae_cliente.numruc, mae_moneda.descripcion AS descmon, IIf(vta_ventas.anulado=-1,'',mae_moneda.simbolo) AS simbolo, mae_impuestos.idcuenvta, " _
            & " con_tc.impcom, mae_tipoproducto.descripcion AS desctipcom, IIf(vta_ventas.anulado=-1,'',mae_condpago.abrev) AS conpagabre, " _
            & " Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4) AS numreg1 FROM ((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN ((mae_documento LEFT JOIN mae_impuestos " _
            & " ON mae_documento.idimp = mae_impuestos.id) RIGHT JOIN (mae_condpago RIGHT JOIN (vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) " _
            & " ON mae_condpago.id = vta_ventas.idconpag) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) " _
            & " LEFT JOIN mae_tipoproducto ON vta_ventas.idtipo = mae_tipoproducto.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
            & " WHERE (((vta_ventas.fchreg)=CDate('" & xFechaRegistro & "')) AND ((vta_ventas.fchdoc)>=CDate('" & DiaIniAño & "'))) ORDER BY vta_ventas.numreg  DESC", xCon
    End If
    
    If mMesActivo = 0 Then
        ' SI SE HA SELECCIONADO APERTURA
        RST_Busq RstVent, "SELECT vta_ventas.*, mae_cliente.nombre, [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] AS numerodoc, IIf(vta_ventas.Anulado=0,'Facturado','Anulado') AS EstadoVenta, " _
            & " mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, mae_moneda.descripcion AS descmon, " _
            & " mae_moneda.simbolo, mae_impuestos.idcuenvta, con_tc.impcom, mae_tipoproducto.descripcion AS desctipcom, mae_condpago.abrev AS conpagabre, " _
            & " Mid([vta_ventas].[numreg],1,2)+[mae_libros].[codsun]+Mid([vta_ventas].[numreg],3,4) AS numreg1 FROM ((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN " _
            & " ((mae_documento LEFT JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) RIGHT JOIN (mae_condpago RIGHT JOIN (vta_ventas LEFT JOIN con_tc " _
            & " ON vta_ventas.fchdoc = con_tc.fecha) ON mae_condpago.id = vta_ventas.idconpag) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) " _
            & " ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_tipoproducto ON vta_ventas.idtipo = mae_tipoproducto.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
            & " WHERE (((vta_ventas.fchreg)=CDate('" & DiaIniAño & "')) AND ((vta_ventas.fchdoc)<CDate('" & DiaIniAño & "'))) ORDER BY [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] DESC", xCon
    End If
    
    If mMesActivo = 13 Then
        MsgBox "Ha selecionado el mes de Cierre, selecciones meses comprendidos entre Enero y Diciembre", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstVent = Nothing
        Set Dg1.DataSource = Nothing
        Dg1.Refresh
        Exit Sub
    End If
    RstVent.Requery
    Set Dg1.DataSource = RstVent
    Dg1.Refresh
End Sub

Private Sub cmdMotDev_Click()
    ReDim xCampos(2, 4) As String
    Dim xRs As New ADODB.Recordset
    Dim titulo As String
    Dim cSQL As String
    
    If QueHace = 3 Then Exit Sub

    'descripcion                     'campo                       'tamaño                         'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Id":           xCampos(0, 1) = "id":           xCampos(0, 2) = "1000":     xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":  xCampos(1, 2) = "5000":     xCampos(1, 3) = "C"
    
    cSQL = "SELECT mae_motivodevolucion.id, mae_motivodevolucion.descripcion " _
        + vbCr + "FROM mae_motivodevolucion"
        
    titulo = "Buscando Motivos"
    
    CARGAR_DLL_EPSBUSCAR xCon, xRs, cSQL, xCampos, titulo, "id", "id"
    
    If xRs.State = 0 Then Exit Sub
    If xRs.RecordCount = 0 Then Exit Sub
    
    txtMotDev.Text = NulosC(xRs("descripcion"))
    lblIdMotDev.Caption = NulosN(xRs("id"))
    
    If NulosN(xRs("id")) = 16 Then ' Otros
        txtMotDevOtr.Text = ""
        txtMotDevOtr.Visible = True
    Else
        txtMotDevOtr.Visible = False
    End If
End Sub

Private Sub CmdMotNotCre_Click()
    ' EJECUTA LA BUSQUEDA DE CONCEPTOS DE NOTA DE CREDITO
    If QueHace = 3 Then Exit Sub

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    If NulosN(TxtTipDoc.Text) = 7 Then
        xform.SQLCad = "SELECT * FROM vta_conceptonc WHERE tipo = 1 ORDER BY descripcion"
        xform.titulo = "Buscando Concepto Nota de Credito"
    End If
    
    If NulosN(TxtTipDoc.Text) = 8 Then
        xform.SQLCad = "SELECT * FROM vta_conceptonc WHERE tipo = 2 ORDER BY descripcion"
        xform.titulo = "Buscando Concepto Nota de Debito"
    End If
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtDocRef.Text = xRs("descripcion")
            LblIdConNC.Caption = xRs("id")
            
            '*******************************
            If NulosN(xRs("id")) = 4 Then ' Devolucion
                txtMotDev.Text = ""
                txtMotDev.Visible = True
                cmdMotDev.Visible = True
            Else
                txtMotDev.Visible = False
                cmdMotDev.Visible = False
            End If
            '*******************************
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub cmdokseldoc_Click()
    ' EMITE UN DOCUMENTO ANULADO
    
    ' VERIFICA QUE LOS DASTOS NECESARIOS SE HAYAN INGRESADO CORRECTAMENTE
    If NulosN(TxtIdAlm2.Text) = 0 Then
        MsgBox "No ha especificado el Almacén", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdAlm2.SetFocus
        Exit Sub
    End If
    
    
    If NulosN(TxtTipDoc2.Text) = 0 Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc2.SetFocus
        Exit Sub
    End If
    
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

    Dim xFecha As String
    xFecha = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    
    ' VERIFICAMOS QUE LA FECHA DE EMISION DEL DOCUMENTO CORRESPONDA AL PERIODO ESPECIFICADO
    If CDate(TxtFchEmiAnul.Valor) < CDate(xFecha) Then
        MsgBox "La fecha del documento no corresponde la periodo contable especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If

    Dim RstCab As New ADODB.Recordset
    'Dim RstDia As New ADODB.Recordset
    Dim xRs As New ADODB.Recordset
    Dim xId As Double
    Dim xNumAsiento As String

    'CONSULTAMOS SI EL DOCUMENTO GENERADO YA EXISTE
    RST_Busq xRs, "SELECT vta_ventas.tipdoc, vta_ventas.numser, vta_ventas.numdoc From vta_ventas " _
        & " WHERE (((vta_ventas.tipdoc)=" & NulosN(TxtTipDoc2.Text) & ") AND ((vta_ventas.numser)='" & NulosC(TxtNumSer2.Text) & "') " _
        & " AND ((vta_ventas.numdoc)='" & NulosC(TxtNumDocGen.Text) & "'))", xCon

    If xRs.RecordCount = 1 Then
        ' SI EXISTE AVISAMOS Y SALE DEL PROCEDIMIENTO
        Set xRs = Nothing
        MsgBox "El numero de documento que quiere emitir ya existe", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumDocGen.SetFocus
        Exit Sub
    End If
    
    ' GENERAMOS EL NUMERO DE ASIENTO
    'xNumAsiento = NuevoNumAsiento(2, mMesActivo, xCon)

On Error GoTo LaCague
    xCon.BeginTrans
    ' GRABAMOS EL DOCUMENTO ANULADO QUE SE ESTA GENERANDO
    ' Validar si el nro de documento existe solo en modo adicionar documento
    RST_Busq RstCab, "SELECT TOP 1 * FROM vta_ventas", xCon
    'RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    ' GRABAMOS LA CABECERA
    xId = HallaCodigoTabla("vta_ventas", xCon, "id")
    RstCab.AddNew
    RstCab("id") = xId
    RstCab("idlib") = 2
    RstCab("idtipo") = 1
    RstCab("tipdoc") = NulosN(TxtTipDoc2.Text)
    RstCab("idcli") = 0
    RstCab("numser") = TxtNumSer2.Text
    RstCab("numdoc") = TxtNumDocGen.Text
    RstCab("Fchdoc") = TxtFchEmiAnul.Valor
    RstCab("Fchven") = TxtFchEmiAnul.Valor
    RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    RstCab("idconpag") = 1
    RstCab("idmon") = 1
    RstCab("impbru") = 0
    RstCab("impinaf") = 0
    RstCab("impigv") = 0
    RstCab("impisc") = 0
    RstCab("impotr") = 0
    RstCab("imptotdoc") = 0
    RstCab("impsal") = 0
    RstCab("idmon") = 1
    'RstCab("numreg") = Format(mMesActivo, "00") + Trim(xNumAsiento)
    RstCab("anulado") = -1
    RstCab("idalm") = NulosN(TxtIdAlm2.Text)
    'Determinamos si es una exportacion
    RstCab("idtipven") = 0 ' en el cual puede ser venta afecta o inafecta para el registro de de ventas
                           ' se valida por programa ver tabla mae_tipoventa
    RstCab("tasaigv") = 0
    
    RstCab.Update
    
''    ' GENERAMOS EL ASIENTO CONTABLE DEL REGISTRO
''    RstDia.AddNew
''    'grabamos el documento de venta en la tabla diario
''    RstDia("año") = AnoTra
''    RstDia("idmes") = mMesActivo
''    RstDia("idlib") = 2
''    RstDia("iddoc") = NulosN(LblIdDocumentoGen.Caption)
''    RstDia("idmov") = xId
''    RstDia("numasi") = xNumAsiento
''    RstDia("tc") = 0
''    RstDia("idcue") = xCuentaDoc
''
''    If TxtIdMon.Text = "1" Then
''        RstDia("impdebsol") = 0
''        RstDia("impdebdol") = 0
''    Else
''        RstDia("impdebsol") = 0
''        RstDia("impdebdol") = 0
''    End If
''    RstDia.Update
''
''    RstDia.AddNew
''
''    ' grabamos el impuesto del documento de venta en la tabla diario
''    RstDia("año") = AnoTra
''    RstDia("idmes") = mMesActivo
''    RstDia("idlib") = 2
''    RstDia("iddoc") = NulosN(LblIdDocumentoGen.Caption)
''    RstDia("idmov") = xId
''    RstDia("numasi") = xNumAsiento
''    RstDia("tc") = 0
''    RstDia("idcue") = xIdCuenTasa
''
''    If TxtIdMon.Text = "1" Then
''        RstDia("impdebsol") = 0
''        RstDia("impdebdol") = 0
''    Else
''        RstDia("impdebsol") = 0
''        RstDia("impdebdol") = 0
''    End If
''    RstDia.Update
    
    
    '---generar asiento
    xNumAsiento = GenerarAsiento(xCon, 2, xId, AnoTra, mMesActivo, 1)
    If xNumAsiento = "" Then GoTo LaCague
    
    ' ----------------------------------------------
    ' Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId

    xCon.CommitTrans
                
    MsgBox "El documento anulado se registró con éxito" & vbCr & "Registro Nº: " & xNumAsiento, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Set RstCab = Nothing
    'Set RstDia = Nothing
    RstVent.Requery
    Dg1.Refresh
    cmdsalirseldoc_Click
    Exit Sub
    
LaCague:
    xCon.RollbackTrans
    Set RstCab = Nothing
    Set xRs = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    
End Sub

Private Sub CmdPreHist_Click()
    ' MUESTRA EL PRECIO HISTORICO DEL ITEM ESPECIFICADO
    If Fg1.Rows < 1 Then Exit Sub
    If Fg1.Row < 1 Then
        MsgBox "Seleccione un Registro para ver el Histórico de Precios", vbInformation, xTitulo
        Exit Sub
    End If
    
    Dim xfrm As New SGI2_funciones.formularios
    Me.MousePointer = vbHourglass
    xfrm.PreciosHistoricos xCon, Fg1.TextMatrix(Fg1.Row, 8), False, NulosC(TxtNumRuc.Text)
    Set xfrm = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub cmdSalirdocsproc_Click()
    ' SALE DEL FRAME Fradocsproc
    Fradocsproc.Visible = False
    Toolbar1.Enabled = True
    TabOne1.Enabled = True
End Sub

Private Sub cmdsalirseldoc_Click()
    ' SALE DEL FRAME Fraseldoc
    QueHace = 3
    ActivarEntorno
    Fraseldoc.Visible = False
End Sub

Private Sub Command1_Click()
    ' ACTUALIZA EL SALDO DEL DOCUMENTO SELECCIONADO
    Dim Rpta As Integer
    
    If NulosN(TxtNewSaldo2.Text) = 0 Then
        MsgBox "Falta especificar el nuevo saldo", vbInformation, xTitulo
        TxtNewSaldo2.SetFocus
        Exit Sub
    End If
    
    If NulosN(TxtNewSaldo2.Text) > NulosN(RstVent("imptotdoc")) Then
        MsgBox "El valor del saldo no puede ser mayor al importe del documento", vbInformation, xTitulo
        TxtNewSaldo2.SetFocus
        Exit Sub
    End If
    
    On Error GoTo LaCague
    
    Rpta = MsgBox("Esta seguro de modificar el saldo del documento", vbInformation + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then
        mIdRegistro = RstVent("id")
        
        'actualizamos el saldo del documento en vta_ventas
        xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = " & NulosN(TxtNewSaldo2.Text) & " WHERE (((vta_ventas.id)=" & RstVent("id") & "))"
        
        RstVent.Requery
        Dg1.Refresh
        
        '--posicionar en la posicion inicial
        If RstVent.RecordCount <> 0 Then
            RstVent.MoveFirst
            RstVent.Find "id=" & mIdRegistro
            If RstVent.EOF = True Then RstVent.MoveFirst
        End If
        
        '--salir
        Command2_Click
        
    End If
    Exit Sub
LaCague:
    
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description), xTitulo
    
End Sub

Private Sub Command2_Click()
    ' SALE DEL FORMULARIO Frame8
    ActivarEntorno
    Frame8.Visible = False
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstVent
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    ' ORDENA LA COLUMNAS DEL DtaGrid Dg1
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstVent.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstVent("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    ' PERMITE EJECUTAR LA BUSQUEDA DE ITEMS EN LA COLUMNA1 DEL CONTROL FlexGrid Fg1
    If NulosC(TxtTipItem.Text) = "" Then
        MsgBox "No ha especificado el tipo de item a buscar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipItem.SetFocus
        Exit Sub
    End If
    
    If optsinguia.Value <> True Then Exit Sub
    
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset

    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(4, 4) As String
    
    xCampos(0, 0) = "Código":       xCampos(0, 1) = "codpro":         xCampos(0, 2) = "2000":    xCampos(0, 3) = "C"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":    xCampos(1, 2) = "4800":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Unid.":        xCampos(2, 1) = "abrev":          xCampos(2, 2) = "500":     xCampos(2, 3) = "C"
    xCampos(3, 0) = "Stock":        xCampos(3, 1) = "stckact":        xCampos(3, 2) = "800":     xCampos(3, 3) = "N"
    
    Dim nSQLId As String
    
    ' obs. apareceran solo items de ventas que tengan cuenta contable
    xform.SQLCad = "SELECT alm_inventario.*, mae_unidades.descripcion AS descuni, mae_unidades.abrev, mae_percepcion.tasa " _
        & " FROM mae_unidades RIGHT JOIN (mae_percepcion RIGHT JOIN alm_inventario ON mae_percepcion.id = alm_inventario.idper) " _
        & " ON mae_unidades.id = alm_inventario.idunimed WHERE (alm_inventario.tippro = " & NulosN(TxtTipItem) & " ) " & nSQLId & " AND alm_inventario.tipo In (2,3) AND alm_inventario.activo = -1 ORDER BY alm_inventario.descripcion"
    
    xform.titulo = "Buscando Productos"
    xform.FormaBusca = OpcionBusquedaForm(1, 1, xCon)
    
    xform.Criterio = ""
    
    Dim RstCamBus As New ADODB.Recordset
    RST_Busq RstCamBus, "SELECT var_opcionesformulario.idform, var_opcionesformulario.campobus From var_opcionesformulario " _
        & " WHERE (((var_opcionesformulario.idform)=78))", xCon
    
    If RstCamBus.RecordCount <> 0 Then
        xform.Ordenado = "codpro" 'RstCamBus("campobus")
        xform.CampoBusca = "codpro" 'RstCamBus("campobus")
    Else
        xform.Ordenado = "codpro"
        xform.CampoBusca = "codpro"
    End If
    
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    Dim A As Integer
    If xRs.State = 1 Then
        If NulosN(xRs("idcuentaven")) = 0 Then
            MsgBox "El item seleccionado no tiene una cuenta contable asignada", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Set xform = Nothing
            Set xRs = Nothing
            Exit Sub
        End If
        
        If xRs.RecordCount <> 0 Then
            '*************************
            Dim RstItemDoc As New ADODB.Recordset
            Dim mDetalleItem As String
            Dim F As New SistemaLogica.Funciones
            ' Validamos si esta configurado el item para mostrar el nombre tecnico
            RST_Busq RstItemDoc, "SELECT mae_itemdocconfig.* From mae_itemdocconfig WHERE (((mae_itemdocconfig.iditem)=" & NulosN(xRs("id")) & ") AND ((mae_itemdocconfig.iddoc)=" & NulosN(TxtTipDoc.Text) & "))", xCon
            If RstItemDoc.RecordCount > 0 Then
                mDetalleItem = F.BuscaCodigoTabla(NulosN(xRs("id")), "id", "desctec", "alm_inventario", "N", xCon)
            Else
                mDetalleItem = NulosC(xRs("descripcion"))
            End If
            '*************************
            Fg1.TextMatrix(Fg1.Row, 1) = mDetalleItem
            Fg1.TextMatrix(Fg1.Row, 2) = NulosC(xRs("abrev"))
            If NulosN(TxtTipItem.Text) <> 5 Then
                If NulosN(LblIdCliente.Caption) <> 0 Then
                    Fg1.TextMatrix(Fg1.Row, 4) = UltimoPrecio(xRs("id"), NulosN(LblIdCliente.Caption)) 'Format(NulosN(xRs("preuni")), "0.0000")
                Else
                    Fg1.TextMatrix(Fg1.Row, 4) = UltimoPrecio(xRs("id"), 0)
                End If
            Else
                Fg1.TextMatrix(Fg1.Row, 4) = 0
            End If
            Fg1.TextMatrix(Fg1.Row, 8) = xRs("id")
            Fg1.TextMatrix(Fg1.Row, 9) = NulosN(xRs("idunimed"))
            Fg1.TextMatrix(Fg1.Row, 10) = NulosN(xRs("idcuentaven"))
            Fg1.TextMatrix(Fg1.Row, 11) = NulosN(xRs("idtipven"))
            Fg1.TextMatrix(Fg1.Row, 12) = NulosN(xRs("tasa"))
            Fg1.TextMatrix(Fg1.Row, 13) = NulosN(xRs("stckact"))
        End If
    End If
    
    If Fg1.Row >= 1 Then
        If NulosN(TxtTipItem.Text) = 5 Then
            Fg1.Col = 4
        Else
            Fg1.Col = 3
        End If
    End If
    
    Set xform = Nothing
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : UltimoPrecio
'* Tipo             : FUNCION
'* Descripcion      : DEVUELVE EL ULTIMO PRECIO ASIGNADO AL ITEM, DEVUELVE UN ENTERO DOBLE DE TENER
'*                    EXITO
'* Paranetros       : NOMBRE    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    IdItem    |  INTEGER   |  ESPECIFICA EL ID DEL ITEM
'*                    IdCliente |  INTEGER   |  ESPECIFICA EL ID DEL CLIENTE
'* Devuelve         : DOUBLE
'*****************************************************************************************************
Function UltimoPrecio(IdItem As Integer, IdCliente As Integer) As Double
    Dim Rst As New ADODB.Recordset
    If IdCliente <> 0 Then
        ' Si hay un clientes asignado buscamos el ultimo precio de venta al cliente
        RST_Busq Rst, "SELECT vta_ventas.fchdoc, vta_ventasdet.preuni, vta_ventasdet.iditem, vta_ventas.idcli FROM vta_ventas LEFT JOIN vta_ventasdet " _
            & " ON vta_ventas.id = vta_ventasdet.idvta Where (((vta_ventasdet.IdItem) = " & IdItem & ") And ((vta_ventas.idcli) = " & IdCliente & ")) " _
            & " ORDER BY vta_ventas.fchdoc, vta_ventasdet.preuni", xCon
    Else
        ' si no hay un cliente especificado buscamos el ultimo precio de venta del item
        RST_Busq Rst, "SELECT vta_ventas.fchdoc, vta_ventasdet.preuni, vta_ventasdet.iditem, vta_ventas.idcli FROM vta_ventas LEFT JOIN vta_ventasdet " _
            & " ON vta_ventas.id = vta_ventasdet.idvta Where (((vta_ventasdet.IdItem) = " & IdItem & ")) ORDER BY vta_ventas.fchdoc, vta_ventasdet.preuni", xCon
    End If
    
    If Rst.RecordCount <> 0 Then
        Rst.MoveLast
        UltimoPrecio = NulosN(Rst("preuni"))
    Else
        Set Rst = Nothing
        RST_Busq Rst, "SELECT * FROM alm_inventario WHERE (id = " & IdItem & ")", xCon
        If Rst.RecordCount = 0 Then
            UltimoPrecio = 0
        Else
            UltimoPrecio = NulosN(Rst("preuni"))
        End If
    End If
    Set Rst = Nothing
End Function

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim xTotPorDes As Double
    
    If Agregando = True Then Exit Sub
    If Mostrando = True Then Exit Sub
    If Fg1.Row < 0 Then Exit Sub
    If optsinguia.Value = True Then
        If Col = 3 Then
            Fg1.TextMatrix(Fg1.Row, 3) = Format(Fg1.TextMatrix(Fg1.Row, 3), "0.0000")
        End If
    End If
    
    If Col = 4 And NulosN(TxtTipItem.Text) <> 5 Then
        Dim xSaldo As Double
        xSaldo = NulosN(Fg1.TextMatrix(Fg1.Row, 13)) - NulosN(Fg1.TextMatrix(Fg1.Row, 3))
        
        If NulosN(Fg1.TextMatrix(Fg1.Row, 3)) > NulosN(Fg1.TextMatrix(Fg1.Row, 13)) And xValidarStckVenta = -1 Then
            MsgBox "No hay suficiente stock del producto : " + Fg1.TextMatrix(Fg1.Row, 1) & Chr(13) _
                & "Cantidad Solicitada : " + Trim(Fg1.TextMatrix(Fg1.Row, 3)) + Chr(13) _
                & "Stock Actual  : " + Trim(Format(Fg1.TextMatrix(Fg1.Row, 13), "0.00")) + Chr(13) _
                & "Faltante        : " + Trim(Str(Format(xSaldo, "0.00"))) + Chr(13), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Fg1.TextMatrix(Fg1.Row, 4) = ""
            Exit Sub
        End If
    End If
        
    If Col = 3 Or Col = 4 Or Col = 5 Or Col = 6 Then
        If NulosN(TxtConPag.Text) = 8 Then
            OptDes1.Value = True
            Fg1.TextMatrix(Fg1.Row, 5) = "100.00"
        End If
        
        If OptDes1.Value = True Then
            xTotPorDes = (NulosN(Fg1.TextMatrix(Fg1.Row, 5)) / 100)
        End If
        
        Fg1.TextMatrix(Fg1.Row, 4) = Format(Fg1.TextMatrix(Fg1.Row, 4), "0.000000")
        Fg1.TextMatrix(Fg1.Row, 5) = Format(Fg1.TextMatrix(Fg1.Row, 5), "0.0000")
        
        If OptDes1.Value = True Then
            Fg1.TextMatrix(Fg1.Row, 6) = NulosN(Fg1.TextMatrix(Fg1.Row, 4)) - (NulosN(Fg1.TextMatrix(Fg1.Row, 4)) * xTotPorDes)
        End If
        If OptDes2.Value = True Then
            Fg1.TextMatrix(Fg1.Row, 6) = NulosN(Fg1.TextMatrix(Fg1.Row, 4)) - NulosN(Fg1.TextMatrix(Fg1.Row, 5))
        End If
        
        Fg1.TextMatrix(Fg1.Row, 6) = Format(Fg1.TextMatrix(Fg1.Row, 6), "0.000000")
        Fg1.TextMatrix(Fg1.Row, 7) = NulosN(Fg1.TextMatrix(Fg1.Row, 6)) * NulosN(Fg1.TextMatrix(Fg1.Row, 3))
        Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), "0.00")
        HallarTotal
    End If
    
    If Col = 14 Or Col = 15 Then
        Dim xIgv As Double
        xIgv = (NulosN(LblIgvTasa.Caption) / 100) + 1
        Fg1.TextMatrix(Fg1.Row, 4) = NulosN(Fg1.TextMatrix(Fg1.Row, 14)) / xIgv
        Fg1.TextMatrix(Fg1.Row, 4) = Format(Fg1.TextMatrix(Fg1.Row, 4), "0.000000")
        
        If NulosN(Fg1.TextMatrix(Fg1.Row, 15)) <> 0 Then
            Dim xPreUni As Double
            xPreUni = NulosN(Fg1.TextMatrix(Fg1.Row, 14)) / NulosN(Fg1.TextMatrix(Fg1.Row, 15))
            Fg1.TextMatrix(Fg1.Row, 4) = xPreUni / xIgv
            Fg1.TextMatrix(Fg1.Row, 4) = Format(Fg1.TextMatrix(Fg1.Row, 4), "0.000000")
            
            Fg1.TextMatrix(Fg1.Row, 3) = Format(Fg1.TextMatrix(Fg1.Row, 15), "0.000000")
        End If
        
        'hallamos los totales
        If OptDes1.Value = True Then
            If xTotPorDes <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 6) = NulosN(Fg1.TextMatrix(Fg1.Row, 4)) / xTotPorDes
            Else
                Fg1.TextMatrix(Fg1.Row, 6) = NulosN(Fg1.TextMatrix(Fg1.Row, 4))
            End If
        End If
        If OptDes2.Value = True Then
            Fg1.TextMatrix(Fg1.Row, 6) = NulosN(Fg1.TextMatrix(Fg1.Row, 4)) - NulosN(Fg1.TextMatrix(Fg1.Row, 5))
        End If
        
        Fg1.TextMatrix(Fg1.Row, 6) = Format(Fg1.TextMatrix(Fg1.Row, 6), "0.000000")
        Fg1.TextMatrix(Fg1.Row, 7) = NulosN(Fg1.TextMatrix(Fg1.Row, 6)) * NulosN(Fg1.TextMatrix(Fg1.Row, 3))
        Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), "0.00")
        HallarTotal
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : HallarTotal
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : HALLA LOS TOTALES DEL DOCUMENTO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub HallarTotal()
    Dim A As Integer
    Dim totalafec As Double
    Dim totalinaf As Double
    
    txtinafecto.Text = "0.00"
    TxtIGV.Text = "0.00"
    txtisc = "0.00"
    TxtTotal.Text = "0.00"
    
    For A = 1 To Fg1.Rows - 1
        If NulosN(TxtTipDoc) = 2 Then ' SI ES RECIBO POR HONORARIOS
            totalafec = totalafec + NulosN(Fg1.TextMatrix(A, 7))     ' venta  gravada
        Else
            If Fg1.TextMatrix(A, 11) = "1" Then 'si es venta gravada
                totalafec = totalafec + NulosN(Fg1.TextMatrix(A, 7)) ' venta  gravada
            Else
                totalinaf = totalinaf + NulosN(Fg1.TextMatrix(A, 7)) ' venta no gravada
            End If
        End If
    Next A
        
    If NulosN(TxtTipDoc) = 1 Then
        TxtIGV.Text = (totalafec * ((TasaImpuesto / 100) + 1)) - totalafec
        TxtTotal.Text = (totalafec * ((TasaImpuesto / 100) + 1)) + totalinaf
    Else
        TxtTotal.Text = (totalafec * ((TasaImpuesto / 100) + 1)) + totalinaf
        If totalafec > 0 Then
            TxtIGV.Text = (totalafec * ((TasaImpuesto / 100) + 1)) - totalafec
        End If
        txtinafecto = totalinaf
    End If
    
    TxtBruto.Text = Format(totalafec, FORMAT_MONTO)
    txtinafecto.Text = Format(totalinaf, FORMAT_MONTO)
    TxtIGV.Text = Format(TxtIGV.Text, FORMAT_MONTO)
    TxtTotal.Text = Format(TxtTotal.Text, FORMAT_MONTO)
End Sub

Private Sub Fg1_EnterCell()
    If Agregando = True Then Exit Sub
    If QueHace = 3 Then Exit Sub
    
    If Fg1.Col = 2 Or Fg1.Col = 7 Then
        Fg1.Editable = flexEDNone
    Else
        If Fg1.Col = 1 Or Fg1.Col = 3 Or Fg1.Col = 5 Or Fg1.Col = 6 Or Fg1.Col = 14 Or Fg1.Col = 15 Then
            If optconguia.Value = True Then
                Fg1.Editable = flexEDNone
            Else
                Fg1.Editable = flexEDKbdMouse
            End If
        End If
        If Fg1.Col = 4 Or Fg1.Col = 5 Then
            Fg1.Editable = flexEDKbdMouse
        End If
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If QueHace = 3 Then KeyAscii = 0
    If KeyAscii = 13 Then Exit Sub
    ' validar los caracteres que se ingresan
    Select Case Col
        Case 1 ' descripcion
        Case 3, 4, 5, 6, 7 ' canpro,preunibru,valdes,preuni,imptot
            If validar_numero(KeyAscii) = False Then KeyAscii = 0
        Case 2 '--abrev
            KeyAscii = 0
        Case 8, 9, 10, 11 ' iditem,idunimed,idcuentaven,idtipven
            KeyAscii = 0
    End Select
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If QueHace = 3 Then Exit Sub
    If optsinguia.Value = True Then
        If KeyCode = 46 Then CmdDelItem_Click
        If KeyCode = 45 Then CmdAddItem_Click
    End If
End Sub

Private Sub Fg1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    
    If optsinguia.Value = True Then
        If Button = 2 Then PopupMenu menu1
    End If
End Sub

Private Sub Fg4_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        cmdagregardocs_Click
    End If
    
    If KeyCode = 46 Then
        If Fg4.Rows = 1 Then Exit Sub
        Fg4.RemoveItem Fg4.Row
        If optconguia.Value = True Then
            MostrarItems
        Else
            MostrarItemsCotizacion
        End If
        HallarTotal
    End If
End Sub

Private Sub Fg4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If QueHace = 3 Then Exit Sub
    
    If Button = 2 Then
        PopupMenu menu2
    End If
End Sub

Private Sub fgdocsproc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 46 Then cmdEliminarOKdocsproc_Click
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE AL CARGAR EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        mMesActivo = xMes
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        pCargarGrid
    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then '--F3 Nuevo
        If fCierrePeriodo = False Then Exit Sub
        If QueHace <> 3 Then Exit Sub
        Nuevo
    End If
    
    If KeyCode = 115 Then '--F4 Modificar
        If fCierrePeriodo = False Then Exit Sub
        If QueHace <> 3 Then Exit Sub
        Modificar
    End If
    
    
'    If KeyCode = 113 Then '--F2 Grabar
'        If fCierrePeriodo = False Then Exit Sub
'        If QueHace = 3 Then Exit Sub
'        If Grabar = True Then
'            QueHace = 3
'            Set RstVent = Nothing
'            Unload Me
'        End If
'    End If
    
    If KeyCode = 116 Then '--F5 actualizar
    End If
    If KeyCode = 117 Then '--F6 '--cancelar
        If fCierrePeriodo = False Then Exit Sub
        If QueHace = 3 Then Exit Sub
        Cancelar
    End If
End Sub

Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE AL CARGAR EL FORMULARIO
    QueHace = 3
    TabOne1.CurrTab = 0
    SeEjecuto = False
    
    Dg1.Columns("fchdoc1").NumberFormat = FORMAT_DATE
    Dg1.Columns("fchven1").NumberFormat = FORMAT_DATE
    
    Dg1.Columns("impbru1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("impigv1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("imptotdoc1").NumberFormat = FORMAT_MONTO
    Dg1.Columns("impsal1").NumberFormat = FORMAT_MONTO
            
    CaracteresNumericos = "0123456789." & Chr(8)
    Fg1.ColWidth(8) = 0
    Fg1.ColWidth(9) = 0
    Fg1.ColWidth(10) = 0
    Fg1.ColWidth(11) = 0
    Fg1.ColWidth(12) = 0
    Fg1.ColWidth(13) = 0
    Fg1.ColWidth(14) = 0
    Fg1.ColWidth(15) = 0
    '5400
    Fg4.ColWidth(3) = 0
    fgdocsproc.Rows = 1
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    
    Fg1.SelectionMode = flexSelectionByRow
    Fg1.ColComboList(1) = ""
    swguiafact = 0
    LblIgvTasa.Caption = ""
    TxtFchDoc.Valor = Date
    TxtFchVen.Valor = Date
    
    TxtFchDoc.Valor = ""
    TxtFchVen.Valor = ""
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando una compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub menu1_1_Click()
    CmdAddItem_Click
End Sub

Private Sub menu1_3_Click()
    CmdDelItem_Click
End Sub

Private Sub Menu1_5_Click()
    CmdPreHist_Click
End Sub

Private Sub menu2_1_Click()
    cmdagregardocs_Click
End Sub

Private Sub menu2_3_Click()
    If Fg4.Rows = 1 Then Exit Sub
    Fg4.RemoveItem Fg4.Row
    CargarGuia
End Sub

Private Sub optconcotizacion_Click()
    ' ESPECIFICA QUE SE ASIGNARAN PEDIDOS DE CLIENTES A LA VENTA
    If QueHace <> 3 Then
        cmdagregardocs.Enabled = True
        CmdAddItem.Enabled = False
        CmdDelItem.Enabled = False
    End If
    'TxtIdTipDoc.Text = "5"
    'TxtIdTipDoc_Validate True
    JALOPEDIDO = True
    Frame6.Left = 9195
    Frame6.Top = 3390
    Frame6.Visible = True
    
    Fg1.ColWidth(1) = 5400
    Fg1.Width = 9105
End Sub

Private Sub optconguia_Click()
    ' ESPECIFICA QUE SE ASIGNARAN GUIAS A LA VENTA
    If QueHace <> 3 Then
        cmdagregardocs.Enabled = True
        CmdAddItem.Enabled = False
        CmdDelItem.Enabled = False
    End If
    TxtIdTipDoc.Text = ""
    LblDescTipDocRef.Caption = ""
    TxtNumDocRef.Text = ""
    LblIdDocRef2.Caption = ""
    
    Frame6.Left = 9210
    Frame6.Top = 3390
    Frame6.Visible = True
    
    Fg1.ColWidth(1) = 3500
    Fg1.Width = 9105
End Sub

Private Sub OptDes1_Click()
    ' ESPECIFICA QUE SE APLICARA EL DESCUENTO POR PORCENTAJE
    If OptDes1.Value = True Then Fg1.TextMatrix(0, 5) = "   % Dscto."
End Sub

Private Sub OptDes2_Click()
    ' ESPECIFICA QUE SE APLICARA EL DESCUENTO EN VALOR
    If OptDes2.Value = True Then Fg1.TextMatrix(0, 5) = "Imp. Dscto."
End Sub

Private Sub OptNo_Click()
    ' ESPECIFICA QUE NO ESTA AFECTO A LA RENTA DE 4TA CATEGORIA
    If OptNo.Value = True Then HallarTotal
End Sub

Private Sub OptSi_Click()
    ' ESPECIFICA QUE ESTA AFECTO A LA RENTA DE 4TA CATEGORIA
    If OptSi.Value = True Then HallarTotal
End Sub

Private Sub optsinguia_Click()
    ' ESPECIFICA QUE ES UNA VENTA SOLA SIN RELACION A GUIAS O ORDENES DE PEDIDOS, QUIERE DECIR QUE LOS ITEMS SE INGRESARAN UNO A UNO
    If QueHace <> 3 Then
        cmdagregardocs.Enabled = False
        CmdAddItem.Enabled = True
        CmdDelItem.Enabled = True
    End If
    
    Fg1.ColWidth(1) = 5400
    Fg1.Width = 11670
    Frame6.Visible = False
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If mMesActivo = 0 Then Cancel = 1: Exit Sub
        ' Validamos si la cuadricula tiene datos
        If QueHace = 3 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No existe información para visualizar", vbInformation, Me.Caption
                Blanquea
                Exit Sub
            ElseIf NulosN(RstVent("anulado")) = -1 Then
                MsgBox "El Documento de Venta esta Anulado", vbInformation, Me.Caption
                Cancel = 1
                Exit Sub
            Else
                
                MuestraSegundoTab
            End If
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then
        If RstVent.RecordCount = 0 Then
            MsgBox "No se han registardos ventas para realizar esta opcion", vbInformation, Me.Caption
            Exit Sub
        End If
        Modificar
    End If
    
    If Button.Index = 3 Then
        If RstVent.RecordCount = 0 Then
            MsgBox "No se han registardos ventas para realizar esta opción", vbInformation, Me.Caption
            Exit Sub
        End If
    
        ' Validamos si el documento esta anulado
        If RstVent("Anulado") = -1 Then
            MsgBox RstVent("nomdoc") & " ya fue anulado, seleccione otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Anular
        
    End If
        
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            RstVent.Requery
            Dg1.Refresh
            
            If RstVent.RecordCount <> 0 Then
                RstVent.MoveFirst
                RstVent.Find "id=" & mIdRegistro
                If RstVent.EOF = True Then RstVent.MoveFirst
            End If
            
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    If Button.Index = 8 Then Filtrar
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstVent.Filter = ""
    End If
    If Button.Index = 10 Then Buscar
    If Button.Index = 11 Then CambiarMes
    
    If Button.Index = 13 Then
    '--cambiado el 14/10/10;
    '--deja de usar pues esta forma solo exporta datos en grilla
''        Dim xFun As New eps_librerias.FuncionesDGrid
''        xFun.xnomemp = NomEmp
''        xFun.xNumRuc = NumRUC
''        xFun.ExportarDGExcel RstVent, Dg1, "DOCUMENTOS DE VENTA DEL MES DE " & UCase(LblMes.Caption)
''        Set xFun = Nothing
        '--Exporta datos adicionales que no necesariamente esten en grilla
        pExportar
    End If
    
    If Button.Index = 14 Then Imprimir
    
    If Button.Index = 16 Then
        Set RstVent = Nothing
        Unload Me
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : OpcionesPeriodo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ESPECIFICA LAS OPCIONES DE EDICION DE REGISTRO PARA EL PERIODO ACTUAL
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub OpcionesPeriodo()
'Modificado 12/01/11 Johan Castro
'           Agregar envío de parametro xIdUsuario a procidimiento CierrePeriodo

    Dim NomMes As String
    Dim xFechaMes As String
    Dim xFchIni, xFchFin As Date
    Dim Rpta As Integer
    
    ' mostrar el boton para agregar apertura
    If mMesActivo = 0 Then CmdApertura.Visible = True Else CmdApertura.Visible = False
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    
    ' bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    
    If mMesActivo <> 0 And mMesActivo <> 13 Then
        xFechaMes = "01/" + Trim(Format(mMesActivo, "00")) + "/" + Trim(Format(Year(Date), "0000"))
        xFchIni = xFechaMes
        xFchFin = Format(HallaDiasMes(CDate(xFechaMes)), "00") + "/" + Mid(xFechaMes, 4, 7)
    End If
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 2 Then
        'MODIFICACION DE DOCUMENTOS
        If ButtonMenu.Index = 1 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados ventas para realizar esta opcion", vbInformation, Me.Caption
                Exit Sub
            End If
            
            If RstVent("anulado") = -1 Then
                MsgBox "No puede modificar " & RstVent("nomdoc") & " anulado proceda a restaurarlo", vbInformation + vbOKOnly + vbDefaultButton1, Me.Caption
                Exit Sub
            Else
                Modificar
            End If
        End If
        
        'RESTAURAR DOCUMENTOS
        If ButtonMenu.Index = 2 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados ventas para realizar esta opcion", vbInformation, Me.Caption
                Exit Sub
            End If
            
            If RstVent("anulado") = -1 Then ' SI EL DOCUMENTO ESTA ANULADO
                RestaurarFactura
            End If
        End If
        
        'ACTUALIZAR SALDO DOCUMENTOS
        If ButtonMenu.Index = 3 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados ventas para realizar esta opción", vbInformation, Me.Caption
                Exit Sub
            End If
            
            If RstVent("anulado") = -1 Then ' SI EL DOCUMENTO ESTA ANULADO
                MsgBox "El documento se encuentra anulado, no puede modificar el saldo", vbInformation, Me.Caption
                Exit Sub
            End If
            
            ModificarSaldo
        End If
        
    End If
     
    If ButtonMenu.Parent.Index = 3 Then
        If ButtonMenu.Index = 1 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados ventas para realizar esta opcion", vbInformation, Me.Caption
                Exit Sub
            End If
            Anular
        End If
        If ButtonMenu.Index = 2 Then
            If RstVent.RecordCount = 0 Then
                MsgBox "No se han registrados ventas para realizar esta opcion", vbInformation, Me.Caption
                Exit Sub
            End If
            
            Eliminar
        End If
        
        If ButtonMenu.Index = 3 Then EmitirAnulada
    End If
    
    If ButtonMenu.Parent.Index = 13 Then
        If ButtonMenu.Index = 1 Then Imprimir
        If ButtonMenu.Index = 2 Then
            'Exportar
        End If
    End If

End Sub

Private Sub TxtConPag_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtConPag_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCondicion_Click
    End If
End Sub

Private Sub TxtConPag_Validate(Cancel As Boolean)
    ' VALIDA LA CONDICION DE PAGO
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtConPag.Text) = "" Then
        'SendKeys vbTab
        Exit Sub
    End If
    Dim xRs1 As New ADODB.Recordset

    RST_Busq xRs1, "SELECT * FROM mae_condpago WHERE id = " & NulosN(TxtConPag.Text) & "", xCon

    If xRs1.RecordCount = 0 Then
        TxtConPag.Text = ""
        LblCondPag.Caption = ""
        TxtFchVen.Valor = ""
    Else
        If TxtFchDoc.Valor = "" Then
            MsgBox "No ha especificado la fecha de emision del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            TxtConPag.Text = ""
            LblCondPag.Caption = ""
            Exit Sub
        End If
        LblCondPag.Caption = Trim(xRs1("descripcion"))
        TxtFchVen.Valor = CDate(TxtFchDoc.Valor) + xRs1("numdia")
    End If
    Set xRs1 = Nothing
End Sub

Private Sub TxtDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdMotNotCre_Click
    End If
End Sub

Private Sub TxtDocRefCredi_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtDocRefCredi_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusDocRef_Click
    End If
    If KeyCode = 46 Then
        TxtDocRefCredi.Text = ""
        LblIdDocRef.Caption = ""
    End If
End Sub

Private Sub TxtFchDoc_Validate(Cancel As Boolean)
    ' VALIDA LA FECHA DEL DOCUMENTO
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtFchDoc.Valor) <> "" Then
        If ChkTC.Value = 0 Then TxtTC.Text = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
    Else
        If ChkTC.Value = 0 Then TxtTC.Text = "0.00"
    End If
End Sub

Private Sub txtglosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
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


Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    ' VALIDA EL ID DE LA MONEDA
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtIdMon.Text) = "" Then
        LblMoneda.Caption = ""
        Exit Sub
    End If
    
    Dim xRs1 As New ADODB.Recordset
    
    'buscamos el codigo de la moneda         digitada
    RST_Busq xRs1, "SELECT * FROM mae_moneda WHERE id = " & NulosN(TxtIdMon.Text) & "", xCon
    
    If xRs1.RecordCount = 0 Then
        TxtIdMon.Text = ""
        LblMoneda.Caption = ""
    Else
        LblMoneda.Caption = Trim(xRs1("descripcion"))
    End If
    Set xRs1 = Nothing

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
    If NulosN(TxtIdTipDoc.Text) = 0 Then
        TxtIdTipDoc.Text = ""
        LblDescTipDocRef.Caption = ""
        
        TxtNumDocRef.Text = ""
        LblIdDocRef2.Caption = ""
        Fg1.Rows = 1
        optconguia.Enabled = True
        optsinguia.Enabled = True
        optconcotizacion.Enabled = False
        Exit Sub
    End If
    
    Dim xRs1 As New ADODB.Recordset
    
    RST_Busq xRs1, "SELECT * FROM mae_docreferencia WHERE id = " & Val(TxtIdTipDoc.Text) & "", xCon
    
    If xRs1.RecordCount = 0 Then
        TxtIdTipDoc.Text = ""
        LblDescTipDocRef.Caption = ""
        TxtNumDocRef.Text = ""
        LblIdDocRef2.Caption = ""
    Else
        LblDescTipDocRef.Caption = Trim(xRs1("descripcion"))
        TxtNumDocRef.Text = ""
        LblIdDocRef2.Caption = ""
        Fg1.Rows = 1
        Fg4.Rows = 1
        If NulosN(TxtIdTipDoc.Text) = 5 Then
            optsinguia.Enabled = False
            optconguia.Enabled = False
            optconcotizacion.Enabled = True
            optconcotizacion.Value = True
            JALOPEDIDO = True
        Else
            optsinguia.Enabled = True
            optconguia.Enabled = True
            optconcotizacion.Enabled = False
            optsinguia.Value = True
            JALOPEDIDO = False
        End If
    End If
    Set xRs1 = Nothing
End Sub

Private Sub TxtIdVen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdVen_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusVen_Click
    End If
End Sub

Private Sub TxtIdVen_Validate(Cancel As Boolean)
    ' VALIDA EL ID DEL VENDEDOR
    If QueHace = 3 Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    
    If NulosC(TxtIdVen.Text) <> "" Then
        Set RstTmp = BuscaConCriterio("SELECT vta_vendedores.*, UCase(pla_empleados!apepat)+' '+ UCase(pla_empleados!apemat)+', '+pla_empleados!nom AS apenom " _
            & " FROM vta_vendedores LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id WHERE vta_vendedores.id = " & NulosN(TxtIdVen.Text) & "", xCon)
        
        If RstTmp.RecordCount <> 0 Then
            LblNomVen.Caption = RstTmp("apenom")
        Else
            TxtIdVen.Text = ""
            LblNomVen.Caption = ""
        End If
    End If

    Set RstTmp = Nothing
End Sub

Private Sub TxtNumDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumDoc_Validate(Cancel As Boolean)
'    ' VALIDA EL NUMERO DE DOCUMENTO
'    If QueHace = 3 Then Exit Sub
'    If NulosC(TxtNumDoc.Text) <> "" Then
'
'        TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
'
'        Dim Rst As New ADODB.Recordset
'        Dim nSQL As String
'        '--ver si existe el numero de doc
'        If QueHace <> 1 Then nSQL = " and vta_ventas.id <> " & NulosN(RstVent("id"))
'
'        RST_Busq Rst, "SELECT vta_ventas.numser, vta_ventas.numdoc, vta_ventas.fchdoc, mae_cliente.nombre, Left([vta_ventas].[numreg],2) & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & Right([vta_ventas].[numreg],4) AS registro " _
'            & " FROM (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
'            & " WHERE (((vta_ventas.numser)='" & Trim(TxtNumSer.Text) & "') AND ((vta_ventas.numdoc)='" & TxtNumDoc.Text & "')) and vta_ventas.tipdoc = " & NulosN(TxtTipDoc.Text) & nSQL, xCon
'
'
'        If Rst.RecordCount <> 0 Then
'            '--poner el nuevo numero doc
'            TxtNumSer_Validate True
'            MsgBox "El número de documento de venta ya existe " & vbCr & "Nº Registro: " & NulosC(Rst("registro")) & vbCr & "Fecha Doc.   " & NulosC(Rst("fchdoc")) & vbCr & "Cliente:         " & NulosC(Rst("nombre")) & vbCr & "Será reemplazado por " + Trim(TxtNumSer.Text) + "-" + Trim(TxtNumDoc.Text), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        End If
'        Set Rst = Nothing
'
'    End If

    Dim idDocumento As Long
    
    If NulosC(TxtNumDoc.Text) = "" Then Exit Sub
    TxtNumDoc.Text = Format(TxtNumDoc.Text, "0000000000")
    
    If QueHace = 1 Then idDocumento = 0 Else idDocumento = F.NuloNumeric(RstVent("id"))
    If F.ExisteDocumento("vta_ventas", "'" & F.NuloString(TxtNumDoc.Text) & "'", xCon, , "'" & F.NuloString(TxtNumSer.Text) & "'", , , , idDocumento, "id") Then
        MsgBox "El documento ingresado ya existe" & vbCr & "Corrija el numero de documento", vbInformation, xTitulo
        TxtNumDoc.Text = ""
        TxtNumDoc.SetFocus
        Exit Sub
    End If
End Sub

Private Sub TxtNumDocGen_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
        TxtNumDocGen_Validate True
    End If
End Sub

Private Sub TxtNumDocGen_Validate(Cancel As Boolean)
    If TxtNumDocGen.Text <> "" Then
        TxtNumDocGen.Text = Format(TxtNumDocGen.Text, "0000000000")
    End If
End Sub

Private Sub TxtNumDocRef_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Fg1.Rows > 1 Then
            Fg1.Col = 1
            Fg1.SetFocus
        Else
            If CmdAddItem.Enabled = True Then CmdAddItem.SetFocus
        End If
    End If
End Sub

Private Sub TxtNumDocRef_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusDocRef2_Click
    End If
    If KeyCode = 46 Then
        TxtDocRefCredi.Text = ""
        LblIdDocRef.Caption = ""
    End If
End Sub

Private Sub TxtNumRuc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumRuc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCli_Click
    End If
End Sub

Private Sub TxtNumRuc_Validate(Cancel As Boolean)
    ' VALIDA EL NUMERO DE R.U.C.
    If NulosC(TxtNumRuc.Text) = "" Then
        Exit Sub
    End If
    
    Dim xRs1 As New ADODB.Recordset
    RST_Busq xRs1, "SELECT * FROM mae_cliente WHERE numruc like '" & TxtNumRuc.Text & "%' ORDER BY numruc", xCon
    If xRs1.RecordCount <> 0 Then
        TxtNumRuc.Text = xRs1("numruc")
        LblNomCli.Caption = xRs1("nombre")
        LblIdCliente.Caption = xRs1("id")
    Else
        TxtNumRuc.Text = ""
        LblNomCli.Caption = ""
        LblIdCliente.Caption = ""
        TxtIdVen.Text = ""
        Lblvendedor.Caption = ""
    End If
    Set xRs1 = Nothing
End Sub

Private Sub TxtNumSer_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If validar_numero(KeyAscii) = False Then KeyAscii = 0
    End If
End Sub

Private Sub TxtNumSer_KeyUp(KeyCode As Integer, Shift As Integer)
'    If KeyCode = 116 Then
'        CmdBusNumSer_Click
'    End If
End Sub

Private Sub TxtNumSer_Validate(Cancel As Boolean)
'    ' VALIDA EL NUMERO DE SERIE
'    If QueHace = 3 Then Exit Sub
'    Dim Rstdoc As New ADODB.Recordset
'    If NulosC(TxtNumSer.Text) = "" Then
'        Exit Sub
'    Else
'        If QueHace <> 1 Then Exit Sub
'
'        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
'        Dim Rst As New ADODB.Recordset
'
'        RST_Busq Rst, "SELECT top 1 vta_ventas.numdoc AS numero from vta_ventas WHERE (((vta_ventas.numser)='" & NulosC(TxtNumSer.Text) & "') AND ((vta_ventas.tipdoc)=" & NulosN(TxtTipDoc.Text) & ")) and IsNumeric(vta_ventas.numdoc)=True " _
'            & " ORDER BY vta_ventas.numdoc DESC ", xCon
'
'        If Rst.RecordCount = 0 Then
'            TxtNumDoc.Text = "0000000001"
'        Else
'            Rst.MoveFirst
'            TxtNumDoc.Text = Format(NulosN(Rst("numero")) + 1, "0000000000")
'        End If
'        Set Rst = Nothing
'    End If

    If QueHace = 3 Then Exit Sub
    
    If NulosC(TxtNumSer.Text) <> "" Then
        TxtNumSer.Text = Format(TxtNumSer.Text, "0000")
        TxtNumDoc.Text = F.HallaNumeroDocumento("vta_ventas", "'" & NulosC(TxtNumSer.Text) & "'", "numser", xCon)
        If NulosC(TxtNumDoc.Text) = "" Then TxtNumSer.Text = ""
    End If
End Sub


Private Sub TxtTipDoc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipDoc_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDoc_Click
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : EmitirAnulada
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA EL FRAM Fraseldoc
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
    
    TxtIdAlm2.Text = ""
    LblAlmacen2.Caption = ""
    TxtTipDoc2.Text = ""
    LblNomDoc2.Caption = ""
    TxtNumSer2.Text = ""
    
    
    TxtNumDocGen.Text = ""
    Fraseldoc.Visible = True
    TxtIdAlm2.SetFocus
End Sub

'*****************************************************************************************************
'* Nombre           : Grabar
'* Tipo             : FUNCION
'* Descripcion      : GRABA UN REGISTRO EN TABLA vta_ventas, ADEMAS ESCRIBE EL ASIENTO CONTABLE DEL
'*                    REGISTRO
'* Paranetros       :
'* Devuelve         : BOOLEAN
'*****************************************************************************************************
Function Grabar() As Boolean
    Dim F As New SistemaLogica.Funciones
    Dim A As Integer
    Dim xFchReg As String
    Dim xFchFin As String
    
    xFchReg = "01/" + Format(mMesActivo, "00") + "/" + Trim(AnoTra)
    A = HallaDiasMes(CDate(xFchReg))
    xFchFin = Trim(Str(A)) + "/" + Format(mMesActivo, "00") + "/" + Trim(AnoTra)
    
    ' VERIFICAMOS QUE LOS DATOS INGRESADOS SEAN LOS CORRECTOS
    If NulosN(TxtTipDoc.Text) <> 0 Then
        If xCuentaDoc = 0 Then
            MsgBox "No se ha asignado una cuenta contable al documento " + LblNomDoc.Caption & Chr(13) _
                & "Asignele una cuenta en el menu Contabilidad opcion Asignar Ctas. Contables a documentos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    End If
    
    If NulosN(TxtTipDoc.Text) <> 0 Then
        If xIdCuenTasa = 0 Then
            MsgBox "El impuesto asignado al documento " + LblNomDoc.Caption & Chr(13) & " no tiene cuenta contable" & Chr(13) _
                & "Asignele una cuenta en el menu Contabilidad opcion Maestro de Impuestos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    End If
    
    For A = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(A, 10)) = 0 Then
            MsgBox "No se le ha asignado una cuenta contable para venta al item : " & Chr(13) _
                & Fg1.TextMatrix(A, 1) & Chr(13) _
                & "Asignele una cuenta en el menu Almacen opcion Mantenimiento Items de Compra y Venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Function
        End If
    Next A
    
    If TxtTipItem.Text = "" Then
        MsgBox "No ha especificado el tipo de item", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipItem.SetFocus
        Exit Function
    End If
    
    If TxtTipDoc.Text = "" Then
        MsgBox "No ha especificado el tipo de documento de venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc.SetFocus
        Exit Function
    End If
    
    If TxtNumRuc.Text = "" Then
        MsgBox "No ha especificado cliente de la venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumRuc.SetFocus
        Exit Function
    End If
    
    If TxtNumSer.Text = "" Or TxtNumDoc.Text = "" Then
        MsgBox "No ha especificado el numero del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer.SetFocus
        Exit Function
    End If
    
    If TxtFchDoc.Valor = "" Then
        MsgBox "No ha especificado la fecha de emision del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchDoc.SetFocus
        Exit Function
    End If
    
    If CDate(TxtFchDoc.Valor) > CDate(xFchFin) Then
        MsgBox "No se puede grabar este documento en el periodo actual la fecha de emision es mayor a : " + xFchFin, vbInformation + vbOKOnly + vbDefaultButton1
        TxtFchDoc.SetFocus
        Exit Function
    End If
    
    
    If TxtFchVen.Valor = "" Then
        MsgBox "No ha especificado la fecha de vencimiento del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchVen.SetFocus
        Exit Function
    End If
    
    If CDate(TxtFchVen.Valor) < CVDate(TxtFchDoc.Valor) Then
        MsgBox " La fecha de vencimiento del documento no puede ser menor a la fecha de emision", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchVen.SetFocus
        Exit Function
    End If
    
    If TxtConPag.Text = "" Then
        MsgBox "No ha especificado la condicion de pago del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtConPag.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "No ha especificado la moneda del documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Function
    End If
    
    If NulosN(TxtIdAlm.Text) = 0 Then
        MsgBox "No ha especificado el nombre del almacen", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdAlm.SetFocus
        Exit Function
    End If
    
    If ChkTC.Value = 1 And NulosN(TxtTC.Text) = 0 Then
        MsgBox "No ha especificado el Tipo de Cambio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTC.SetFocus
        Exit Function
    End If
    
       
    If NulosN(TxtIdTipDoc.Text) <> 0 Then
        If NulosN(LblIdDocRef2.Caption) = 0 Then
            MsgBox "No ha especificado el documento de referencia para este documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            'TxtNumDocRef2.SetFocus
            Exit Function
        End If
    End If
       
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado items para la venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Exit Function
    End If
    
    ' Validamos si existe el numero del documento en modo QUEHACE = 1
    If QueHace = 1 Then
        Dim RstCab As New ADODB.Recordset
    
        RST_Busq RstCab, "SELECT * FROM vta_ventas WHERE tipdoc =" & NulosN(TxtTipDoc.Text) & " AND numser ='" & TxtNumSer.Text & "' AND numdoc = '" & TxtNumDoc.Text & "' ", xCon
    
        If RstCab.RecordCount > 0 Then
            ' SI EXISTE EL NUMERO DE DOCUMENTO, GENERAMOS UN NUEVO NUMERO DE DOCUMENTO
            MsgBox "El Nro de documento ha sido registrado por otro usuario se grabara con otro numero", vbInformation, Me.Caption
            TxtNumDoc.Text = HallaNumdocVenta(NulosN(TxtTipDoc.Text), TxtNumSer.Text, xCon)
        End If
        Set RstCab = Nothing
    End If
    
    ' SI ES NOTA DE CREDITO
    If NulosN(TxtTipDoc.Text) = 7 Then
        ' VERIFICAMOS QUE SE HAYA INGRESADO EL DOCUMENTO DE REFERENCIA PARA LA NOTA DE CREDITO
        If NulosC(TxtDocRefCredi.Text) = "" Then
            MsgBox "No ha especificado el documento al que hace referencia la nota de crédito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
    End If
    
    Dim RstDeta2 As New ADODB.Recordset
    Dim RstActPro As New ADODB.Recordset
    
    Dim RstDet As New ADODB.Recordset
'    Dim RstDia As New ADODB.Recordset
    Dim xIdCuen As Integer
    Dim xTotal As Double
    Dim xSaldo As Double
    
    Dim xidtipven As String 'Determina si la venta es de tipo exportacion
    Dim xNumAsiento As String
    
    Dim xId As Double
    Dim X As Integer
    Dim P As Integer
    
On Error GoTo LaCague
    swguiafact = 1
    xCon.BeginTrans
    
    If QueHace = 1 Then
        ' SI SE ESTA AGREGANDO UN REGISTRO
        xId = HallaCodigoTabla("vta_ventas", xCon, "id")         ' OBTENEMOS EL ID PARA EL REGISTRO ACTUAL
        
        xNumAsiento = NuevoNumAsiento(2, mMesActivo, xCon)       ' OBTENEMOS EL NUMERO DE ASIENTO PARA EL REGISTRO ACTUAL
        
        RST_Busq RstCab, "SELECT TOP 1 * FROM vta_ventas", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
        
        If NulosN(TxtTipDoc.Text) = 7 Then
            xSaldo = 0
        Else
            xSaldo = NulosN(TxtTotal.Text)
        End If
    Else
        xId = RstVent("id")                                      ' ASIGNAMOS EL ID DEL REGISTRO ACTUAL
        RST_Busq RstCab, "SELECT * FROM vta_ventas WHERE id = " & xId & "", xCon
        
        ' Eliminamos el stock agregado con la venta
        RST_Busq RstDeta2, "SELECT vta_ventasdet.* From vta_ventasdet WHERE (((vta_ventasdet.idvta)= " & xId & "))", xCon

        If RstDeta2.RecordCount <> 0 Then
            RstDeta2.MoveFirst
            For A = 1 To RstDeta2.RecordCount
                RST_Busq RstActPro, "SELECT alm_inventario.id, alm_inventario.stckact  From alm_inventario WHERE ((alm_inventario.id=" & RstDeta2("iditem") & "))", xCon
                If RstActPro.RecordCount = 1 Then
                    RstActPro("stckact") = RstActPro("stckact") + RstDeta2("canpro")
                    RstActPro.Update
                End If
                Set RstActPro = Nothing
            Next A
        End If
        Set RstDeta2 = Nothing
        
        ' eliminamos la referencia del documento a la orden de pedido
        xCon.Execute "UPDATE ped_pedidodetent SET ped_pedidodetent.idtipdoc = 0, ped_pedidodetent.iddocven = 0, ped_pedidodetent.estado = 1" _
            & " WHERE (((ped_pedidodetent.idtipdoc)=2) AND ((ped_pedidodetent.iddocven)=" & xId & "))"
        
        ' Eliminamos el detalle de la venta
        xCon.Execute "DELETE * FROM vta_ventasdet WHERE idvta = " & xId & ""
        ' Obtenemos el numero de asiento
        xNumAsiento = Mid(RstVent("numreg"), 3, 4)
                    
        If NulosN(TxtTipDoc.Text) = 7 Or NulosN(TxtTipDoc.Text) = 8 Then
            xSaldo = 0
        Else
            xSaldo = NulosN(TxtTotal.Text)
        End If
    End If
    
    RST_Busq RstDet, "SELECT TOP 1 * FROM vta_ventasdet", xCon
'    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    
    mIdRegistro = xId
    ' ESCRIBIMOS LA CABECERA DEL REGISTRO
    RstCab("idlib") = 2
    RstCab("idtipo") = NulosN(TxtTipItem.Text)
    RstCab("tipdoc") = NulosN(TxtTipDoc.Text)
    RstCab("idcli") = NulosN(LblIdCliente.Caption)
    RstCab("numser") = TxtNumSer.Text
    RstCab("numdoc") = TxtNumDoc.Text
    RstCab("fchdoc") = TxtFchDoc.Valor
    RstCab("fchven") = TxtFchVen.Valor
    RstCab("idconpag") = NulosN(TxtConPag.Text)
    RstCab("idmon") = NulosN(TxtIdMon.Text)
    RstCab("impbru") = NulosN(TxtBruto.Text)
    RstCab("impinaf") = NulosN(txtinafecto.Text)
    RstCab("impigv") = NulosN(TxtIGV.Text)
    RstCab("impisc") = NulosN(txtisc.Text)
    RstCab("impotr") = 0
    RstCab("imptotdoc") = NulosN(TxtTotal.Text)
    
    If QueHace = 1 Then
        RstCab("impsal") = NulosN(TxtTotal.Text)
    End If
    
    RstCab("idalm") = NulosN(TxtIdAlm.Text)
    
    RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + AnoTra)
    
    If OptDes1.Value = True Then RstCab("tipdes") = 1
    If OptDes2.Value = True Then RstCab("tipdes") = 2
    
    If CONTABILIZAR = True Then 'se actualiza internamente
        'RstCab("numreg") = Trim(Format(Str(mMesActivo), "00")) + xNumAsiento
    End If
    
    RstCab("idven") = NulosN(TxtIdVen.Text)
    
    If NulosN(TxtTipDoc.Text) = 7 Or NulosN(TxtTipDoc.Text) = 8 Then
        RstCab("iddocref") = NulosN(LblIdDocRef.Caption)
        RstCab("idmotnotcre") = NulosN(LblIdConNC.Caption)
        
        '*************************************************************************
        If txtMotDev.Visible Then RstCab("idmotdev") = NulosN(lblIdMotDev.Caption)
        If txtMotDevOtr.Visible Then RstCab("desmotdev") = NulosC(txtMotDevOtr.Text)
        '*************************************************************************
    End If
    
    RstCab("anulado") = 0
    ' grabamos el documento de referencia para la venta (orden venta, orden de despacho etc)
    RstCab("idtipdocref") = NulosN(TxtIdTipDoc.Text)
    RstCab("iddocref2") = NulosN(LblIdDocRef2.Caption)
    '--uso temporal
    RstCab("numerodocref") = NulosC(TxtNumDocRef.Text)
    
    RstCab("glosa") = NulosC(txtglosa.Text)
    
    ' Determinamos si es una exportacion
    For A = 1 To Fg1.Rows - 1
        xidtipven = NulosN(Fg1.TextMatrix(A, 11))
    Next A
    
    If xidtipven = 2 And NulosN(TxtIGV.Text) = 0 Then 'si esta venta exportacion
        RstCab("idtipven") = 2
    Else
        RstCab("idtipven") = 0 ' en el cual puede ser venta afecta o inafecta para el registro de
                               ' de ventas se valida por programa ver tabla mae_tipoventa
    End If
    
    If optsinguia.Value = True Then RstCab("oriitem") = 1
    If optconguia.Value = True Then RstCab("oriitem") = 2
    If optconcotizacion.Value = True Then RstCab("oriitem") = 3
    
    If ChkTC.Value = 1 Then
        RstCab("tc") = NulosN(TxtTC.Text)
    Else
        RstCab("tc") = 0
    End If
    
    '--grabar la tasa del igv aplicada; solo si hay impuesto
    If NulosN(TxtIGV.Text) <> 0 Then
        RstCab("tasaigv") = TasaImpuesto
    Else
        RstCab("tasaigv") = 0
    End If
        
    ' Actualizamos el saldo del documento
    If (NulosN(LblIdDocRef.Caption) <> 0) Then
        ActualizaSaldoDoc NulosN(LblIdDocRef.Caption), 2, NulosN(TxtTotal.Text)
    End If

    RstCab.Update
    ' GRABAMOS EL DETALLE DEL REGISTRO
    For A = 1 To Fg1.Rows - 1
        RstDet.AddNew
        RstDet("idvta") = xId
        RstDet("iditem") = NulosN(Fg1.TextMatrix(A, 8))
        RstDet("idunimed") = NulosN(Fg1.TextMatrix(A, 9))
        
        If NulosN(Fg1.TextMatrix(A, 6)) <> 0 Then
            RstDet("preuni") = NulosN(Fg1.TextMatrix(A, 6))
        Else
            RstDet("preuni") = NulosN(Fg1.TextMatrix(A, 4))
        End If
        
        RstDet("valdes") = NulosN(Fg1.TextMatrix(A, 5))
        RstDet("canpro") = NulosN(Fg1.TextMatrix(A, 3))
        RstDet("imptot") = NulosN(Fg1.TextMatrix(A, 7))
        RstDet("tasaper") = NulosN(Fg1.TextMatrix(A, 12))
        RstDet("preunibru") = NulosN(Fg1.TextMatrix(A, 4))
        RstDet.Update
    Next A
   
    ' ACTUALIZAMOS EL STOCK
'    If optsinguia.Value = True Or optconcotizacion.Value = True Then
'        For A = 1 To Fg1.Rows - 1
'            xCon.Execute "UPDATE alm_inventario SET alm_inventario.stckact = ( [alm_inventario]![stckact]-" & NulosN(Fg1.TextMatrix(A, 3)) & ")" _
'                & " WHERE (((alm_inventario.id)=" & NulosN(Fg1.TextMatrix(A, 8)) & "))"
'        Next A
'    End If
        
    ' ACTUALIZAMOS EL PEDIDO CON LA VENTA EFECTUADA
    If NulosN(TxtIdTipDoc.Text) = 5 Then
        'actualizamos idtipdoc  = 1 porque es una guia
        xCon.Execute "UPDATE ped_pedidodetent SET ped_pedidodetent.estado = 1, ped_pedidodetent.idtipdoc = 2, ped_pedidodetent.iddocven = " & xId & "" _
            & " WHERE (((ped_pedidodetent.idped)=" & VAR_IDPEDIDO & ") AND ((ped_pedidodetent.fchent)=CDate('" & VAR_FECHAPEDIDO & "')))"
    End If
    
    ' GRABAMOS EL DETALLE DEL ITEM
    Dim B As Integer
    
    ' BORRAMOS EL DETALLE DEL ITEM EN CASO DE QUE SEA UNA MODIFICACION
    If QueHace = 2 Then
        xCon.Execute "DELETE * FROM vta_ventasdetitems WHERE idventa = " & xId & ""
    End If
    
    For A = 1 To Fg1.Rows - 1
        RstVentItemsDeta.Filter = adFilterNone
        RstVentItemsDeta.Filter = "iditem = " & NulosN(Fg1.TextMatrix(A, 8)) & ""

        If RstVentItemsDeta.RecordCount <> 0 Then
            RstVentItemsDeta.Sort = "orden"
            RstVentItemsDeta.MoveFirst
            B = 0
            For B = 1 To RstVentItemsDeta.RecordCount
                xCon.Execute "INSERT INTO vta_ventasdetitems ( idventa, iditem, orden, texto )" _
                    & " SELECT " & xId & " AS Expr1, " & RstVentItemsDeta("iditem") & " AS Expr2, " & RstVentItemsDeta("orden") & " AS Expr3, '" & RstVentItemsDeta("texto") & "' AS Expr4"

                RstVentItemsDeta.MoveNext
                If RstVentItemsDeta.EOF = True Then Exit For
            Next B
        End If
    Next A
    
    If CONTABILIZAR = True Then ' SI ESTA EN MODO CONTABILIZAR
    End If
     
    ' Actualizamos en el campo Iddocven de la tabla Guias el Id del Documento de Venta para relacionarlo Factura -Guia
    If optconguia.Value = True Then
        For X = 1 To Fg4.Rows - 1
            xCon.Execute " UPDATE vta_guia SET vta_guia.iddocven = " & xId & " WHERE vta_guia.id = " & NulosN(Fg4.TextMatrix(X, 3)) & ""
        Next X
    End If
    
    ' Actualizamos en el campo Iddocven de la tabla cotizaciones el Id del Documento de Venta para relacionarlo Factura - Cotizacion
    If optconcotizacion.Value = True Then
        For X = 1 To Fg4.Rows - 1
            xCon.Execute " UPDATE vta_cotizacion SET vta_cotizacion.iddocven = " & xId & ", vta_cotizacion.idest = 3 WHERE vta_cotizacion.id = " & NulosN(Fg4.TextMatrix(X, 3)) & ""
        Next X
    End If
    
    ' si esta afecto a la detraccion
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT mae_detraccion.id, mae_detraccion.descripcion, mae_detraccion.tasa, alm_inventario.iddet " _
        & " FROM alm_inventario LEFT JOIN mae_detraccion ON alm_inventario.iddet = mae_detraccion.id " _
        & " WHERE ((alm_inventario.id= " & Val(Fg1.TextMatrix(Fg1.Row, 8)) & "))", xCon

    If Rst.RecordCount <> 0 Then
        If Rst("iddet") <> 0 Then
            Dim RstDeta As New ADODB.Recordset
            Dim xId2 As Integer
            
            If QueHace = 1 Then
                xId2 = HallaCodigoTabla("con_detraccion", xCon, "id")
                RST_Busq RstDeta, "SELECT * FROM con_detraccion", xCon
                RstDeta.AddNew
                RstDeta("id") = xId2
            Else
                RST_Busq RstDeta, "SELECT con_detraccion.* From con_detraccion " _
                    & " WHERE (((con_detraccion.iddoc)=" & xId & "))", xCon
            End If
            
            If RstDeta.RecordCount = 0 Then
                ' este procedimiento es solo para cuando se este modificando una compra afecta a la detraccion y no se le haya hecho la detraccion a la hora de ingresar la compra
                xId2 = HallaCodigoTabla("con_detraccion", xCon, "id")
                RstDeta.AddNew
                RstDeta("id") = xId2
            End If
            
            RstDeta("iddet") = Rst("iddet")
            RstDeta("por") = Rst("tasa")
            RstDeta("iddoc") = xId
            RstDeta("idmon") = NulosN(TxtIdMon.Text)
            RstDeta("tipo") = 2   'especificamos que es una venta
            RstDeta("fchmov") = Date
            RstDeta("Glosa") = ""
            RstDeta("imp") = Format((NulosN(TxtTotal.Text) * (Rst("tasa") / 100)), "0.00")
            RstDeta("numdet") = "SIN NUMERO"
            RstDeta.Update
        End If
    End If
    Dim nSQL As String
    
    
    '-------------------------------------------------------------------------------------
    ' Creamos el movimiento automatico
    If F.NuloNumeric(F.KeyValue("CreacionMovimientoAutoVenta", xCon)) = -1 Then
        ' Verificamos si ya tiene registro en movimientos
        Dim database As New SistemaData.EDataBase
        Dim record As New ADODB.Recordset
        Dim Movimiento As New AlmacenEntidad.EMovimiento
        
        Set database.Connection = xCon
        database.CommandText = "SELECT alm_ingreso.id AS idmov " _
                    + vbCr + "FROM alm_ingreso " _
                    + vbCr + "WHERE (((alm_ingreso.idtipdocref)=" & F.NuloNumeric(F.KeyValue("IdDocumentoFactura", xCon)) & ") AND ((alm_ingreso.iddocref)=" & xId & "))"
        Set record = database.GetRecordset
        If record.RecordCount > 0 Then
            ' Eliminamos todos los movimientos ya generados
            record.MoveFirst
            While Not record.EOF
                Dim mMovAux As New AlmacenEntidad.EMovimiento
                Set mMovAux.Conexion = xCon
                mMovAux.IdMovimiento = F.NuloNumeric(record("idmov"))
                mMovAux.Delete CLng(xIdUsuario), F.MachineName
                record.MoveNext
            Wend
        End If
        
        ' Si no tiene Guias se crea el movimiento
        If optsinguia.Value = True Then
            ' Se valida la fecha de cierre de mes
            If F.MesCerradoOpcion(F.RetornarMesFecha(CDate(TxtFchDoc.Valor)), CLng(F.KeyValue("IdOpcionSistemaMovimientoAlmacen", xCon)), xCon) Then
                Err.Raise &HFFFFFF01, , "El mes al que pertenece el documento se encuentra cerrado para la opcion: [Ingresos y Salidas de almacen] y no se pueden generar movimientos automaticos, modifique la fecha o aperture el mes "
            Else
                ' Cabecera
                ' Si es NC se crea un ingreso sino una salida
                If F.NuloNumeric(TxtTipDoc.Text) = F.NuloNumeric(F.KeyValue("IdDocumentoNotaCredito", xCon)) Then
                    Movimiento.IdTipoMovimiento = -1
                Else
                    Movimiento.IdTipoMovimiento = 0
                End If
                Movimiento.FechaMovimiento = CDate(TxtFchDoc.Valor)
                Movimiento.NumeroSerie = F.NuloString(TxtNumSer.Text)
                Movimiento.NumeroDocumento = F.HallaNumeroDocumento("alm_ingreso", "'" & Movimiento.NumeroSerie & "'", "numser", xCon)
                Movimiento.IdEstado = F.NuloNumeric(F.KeyValue("EstadoAprobadoMovimiento", xCon))
                Movimiento.IdAlmacen = F.NuloNumeric(TxtIdAlm.Text)
                Movimiento.Glosa = F.NuloString(txtglosa.Text)
                Movimiento.IdTipoDocumentoReferencia = F.NuloNumeric(F.KeyValue("IdDocumentoFactura", xCon))
                Movimiento.IdDocumentoReferencia = xId
                Movimiento.DocumentoReferencia = F.NuloString(TxtNumSer.Text & " - " & TxtNumDoc)
                Movimiento.MesTrabajo = mMesActivo
                Movimiento.AnhoTrabajo = AnoTra
                ' Detalle
                For A = 1 To Fg1.Rows - 1
                    Dim MovimientoDet As New AlmacenEntidad.EMovimientoDet
                    MovimientoDet.IdItem = F.NuloNumeric(Fg1.TextMatrix(A, 8))
                    MovimientoDet.Cantidad = F.NuloNumeric(Fg1.TextMatrix(A, 3))
                    MovimientoDet.CantidadTeorica = F.NuloNumeric(Fg1.TextMatrix(A, 3))
                    ' Se agrega al padre
                    Movimiento.LMovimientoDet.Add MovimientoDet
                    Set MovimientoDet = Nothing
                Next A
                ' Se graba el movimiento
                Set Movimiento.Conexion = xCon
                If Not Movimiento.Save(CLng(xIdUsuario), F.MachineName) Then Err.Raise &HFFFFFF01, , F.ErrorDescriptionDLL(Err.LastDllError)
            End If
        End If
        
    End If
    '-------------------------------------------------------------------------------------
    
    '----------------------------------------------------------------------------------
    
    '---generar asiento
    xNumAsiento = GenerarAsiento(xCon, 2, xId, AnoTra, mMesActivo, 1)
    If xNumAsiento = "" Then GoTo LaCague
    '----------------------------------------------------------------------------------
    
    ' ----------------------------------------------
    ' Grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    ' ----------------------------------------------------------------------------
    ' grabamos la operacion para el analisis de cuenta por documento de referencia
    GrabarOperacionCtaCte 2, xId, xCon
     
    xCon.CommitTrans
    
    MsgBox "La " & Trim(LblNomDoc) & " se registró con éxito" & vbCr & "Registro Nº: " & xNumAsiento, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    Set RstCab = Nothing
    Set RstDet = Nothing
'    Set RstDia = Nothing
    Grabar = True
    Exit Function
    
LaCague:
   'Resume
    xCon.RollbackTrans
'    Set rstdocus = Nothing
    Set RstCab = Nothing
    Set RstDet = Nothing
'    Set RstDia = Nothing
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)
    Grabar = False
    Exit Function
End Function

'*****************************************************************************************************
'* Nombre           : HallaNumAsiento
'* Tipo             : FUNCION
'* Descripcion      : HALLA EL NUMERO DE ASIENTO ACTUAL, ESTA FUNCION DEVUELVE UNA CADENA QUE CONTIENE
'*                    EL NUMERO DE ASIENTO
'* Paranetros       : NOMBRE    |  TIPO             |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Mes       |  INTEGER          |  ESPECIFICA EL ID DEL MES ACTUAL
'* Devuelve         : STRINF
'*****************************************************************************************************
Function HallaNumAsiento(Mes As Integer) As String
    Dim Rst As New ADODB.Recordset
    RST_Busq Rst, "SELECT con_diario.idmes, con_diario.idlib, con_diario.numasi From con_diario " _
        & " WHERE (((con_diario.idmes)=" & Mes & ") AND ((con_diario.idlib)=2)) ORDER BY numasi", xCon
    
    If Rst.RecordCount = 0 Then
        HallaNumAsiento = "0001"
    Else
        Rst.MoveLast
        HallaNumAsiento = Format(NulosN(Rst("numasi")) + 1, "0000")
    End If
    Exit Function
End Function

Private Sub TxtTipDoc_Validate(Cancel As Boolean)
    ' VALIDA EL TIPO DE DOCUMENTO
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtTipDoc.Text) = "" Then
        LblNomDoc.Caption = ""
        Exit Sub
    End If
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
       
    RST_Busq xRs, "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, mae_impuestos.Abrev AS abreimp, " _
        & " mae_impuestos.idcuenvta AS cuentaimp, alm_numseries.numser, alm_numseries.idtipdoc " _
        & " FROM alm_numseries LEFT JOIN (mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuenvta = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id) ON alm_numseries.idtipdoc = mae_documento.id " _
        & " WHERE (((alm_numseries.idalm)=" & NulosN(TxtIdAlm.Text) & ") AND ((alm_numseries.idtipdoc)=" & NulosN(TxtTipDoc.Text) & "))", xCon
        
    If xRs.RecordCount = 0 Then
        TxtTipDoc.Text = ""
        LblNomDoc.Caption = ""
    Else
        TxtTipDoc.Text = NulosN(xRs("id"))
        LblNomDoc.Caption = NulosC(xRs("descripcion"))
        TasaImpuesto = NulosN(xRs("tasa"))
        
        xIdCuenTasa = NulosN(xRs("cuentaimp"))
        
        Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & NulosN(TxtTipDoc.Text) & " and mae_documentocta.idmon =" & NulosN(TxtIdMon) & " and tipope = -1", xCon)
        If xRs2.RecordCount > 0 Then
            xCuentaDoc = NulosN(xRs2("idcuen"))
        End If
        
        Set xRs2 = Nothing
        
        LblRotulo.Caption = Trim(NulosC(xRs("abreimp"))) + " (         )"
        LblIgvTasa.Caption = Format(Trim(Str(TasaImpuesto)), "0.00")
        
        If xRs("id") = 7 Or xRs("id") = 8 Then ' Nota de Credito/Nota de Debito
            Label33.Visible = True
            TxtDocRefCredi.Visible = True
            CmdBusDocRef.Visible = True
            FraRetencion.Visible = False
            Frame5.Left = 4905
            Frame5.Top = 2790
            Frame5.Visible = True
            Frame3.Visible = False
            txtMotDev.Visible = False
            cmdMotDev.Visible = False
            txtMotDevOtr.Visible = False
        Else
            Frame5.Visible = False
            Label33.Visible = False
            Frame3.Visible = True
            TxtDocRefCredi.Visible = False
            CmdBusDocRef.Visible = False
        End If
    End If
    
    ' si es Recibo por honorarios
    If NulosN(TxtTipDoc) = 2 Then
         FraRetencion.Enabled = True
         FraRetencion.Visible = True
         Fratipven.Enabled = False
         FraRetencion.Caption = "Retención de 4ta Categoria " & Trim(Str(TasaImpuesto)) + "%"
         txtisc.Enabled = False
         txtinafecto.Enabled = False
    Else
         Fratipven.Enabled = True
         FraRetencion.Enabled = False
         FraRetencion.Visible = False
         txtisc.Enabled = True
         txtinafecto.Enabled = True
    End If
    
    ' buscamos para hallar el numero de serie asignado al almacen
    If TxtTipDoc.Text <> "" Then
        Dim Rst As New ADODB.Recordset
        Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc.Text) & " AND idalm = " & NulosN(LblIdAlmacen.Caption) & "", xCon)
        If Rst.RecordCount <> 0 Then
            TxtNumSer.Text = Rst("numser")
            TxtNumSer_Validate True
        End If
        Set Rst = Nothing
    Else
        TxtNumSer.Text = ""
        TxtNumDoc.Text = ""
    End If
    
    Set xRs = Nothing
End Sub

Private Sub TxtTipItem_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipItem_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipItem_Click
    End If
End Sub

'*****************************************************************************************************
'* Nombre           : Filtrar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA LA ACCION DE FILTRAR SOBRE EL RECORDSET RstVent
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Filtrar()
    TabOne1.CurrTab = 0
    Dim xform As New eps_librerias.FormFiltrar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(7, 4) As String
    
    xCampos(0, 0) = "Cliente":         xCampos(0, 1) = "nombre":        xCampos(0, 2) = "C":         xCampos(0, 3) = "1500"
    xCampos(1, 0) = "Fch. Emision":    xCampos(1, 1) = "fchdoc":        xCampos(1, 2) = "F":         xCampos(1, 3) = "1500"
    xCampos(2, 0) = "Nº Documento":    xCampos(2, 1) = "numerodoc":     xCampos(2, 2) = "C":         xCampos(2, 3) = "1500"
    xCampos(3, 0) = "Tipo Documento":  xCampos(3, 1) = "abrev":         xCampos(3, 2) = "C":         xCampos(3, 3) = "1500"
    xCampos(4, 0) = "Forma de Pago":   xCampos(4, 1) = "desccond":      xCampos(4, 2) = "C":         xCampos(4, 3) = "1500"
    xCampos(5, 0) = "Moneda":          xCampos(5, 1) = "simbolo":       xCampos(5, 2) = "C":         xCampos(5, 3) = "1500"
    xCampos(6, 0) = "Estado":          xCampos(6, 1) = "estadoventa":   xCampos(6, 2) = "C":         xCampos(6, 3) = "1500"
    xCampos(7, 0) = "importe":         xCampos(7, 1) = "imptotdoc":     xCampos(7, 2) = "N":         xCampos(7, 3) = "1500"
    
    Set xform.Coneccion = xCon
    Set xform.Rst = RstVent
    Set RstVent = xform.FiltrarReg(xCampos)
    Set Dg1.DataSource = RstVent
    Dg1.Refresh
End Sub

Private Sub TxtTipItem_Validate(Cancel As Boolean)
    ' VALIDA EL TIPO DE ITEM SELECCIONADO
    If QueHace = 3 Then Exit Sub
    Dim RstTmp As New ADODB.Recordset
    If NulosC(TxtTipItem.Text) <> "" Then
        Set RstTmp = BuscaConCriterio("SELECT * FROM mae_tipoproducto WHERE id = " & NulosN(TxtTipItem.Text) & "", xCon)
        If RstTmp.RecordCount <> 0 Then
           LblTipoItem.Caption = RstTmp("descripcion")
        Else
            TxtTipItem.Text = ""
            LblTipoItem.Caption = ""
        End If
    Else
        LblTipoItem.Caption = ""
    End If
    Set RstTmp = Nothing
    
    pGridConfigurar
    
End Sub

'*****************************************************************************************************
'* Nombre           : Imprimir
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : IMPRIME EL DOCUMENTO FISICO DE LA VENTA
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Imprimir()
    Dim RsPDoc As New ADODB.Recordset
    Dim RsPCab As New ADODB.Recordset
    Dim RsPDet As New ADODB.Recordset
    Dim xRsDoc As New ADODB.Recordset
    Dim xRsDet As New ADODB.Recordset
    Dim RstGui As New ADODB.Recordset
    Dim A As Integer
    Dim xCadGuias As String
    Dim F As New SistemaLogica.Funciones
    
    ' BUSCAMOS LA VENTA
    Dim mSQL As String
    
    mSQL = "SELECT vta_ventas.*, mae_cliente.nombre, mae_cliente.Dir, mae_cliente.numruc, mae_moneda.descripcion AS mon, [vta_conceptonc]![descripcion] & ' : REF A => ' & [mae_documento]![abrev] & ' ' & [vta_ventas_1]![numser] & '-' & [vta_ventas_1]![numdoc] AS docref2, pla_empleados.nombre AS nomven " _
        + vbCr + "FROM (((((mae_moneda RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON mae_moneda.id = vta_ventas.idmon) LEFT JOIN vta_conceptonc ON vta_ventas.idmotnotcre = vta_conceptonc.id) LEFT JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id) LEFT JOIN mae_documento ON vta_ventas_1.tipdoc = mae_documento.id) LEFT JOIN vta_vendedores ON vta_ventas.idven = vta_vendedores.id) LEFT JOIN pla_empleados ON vta_vendedores.idper = pla_empleados.id " _
        + vbCr + "WHERE (((vta_ventas.id) = " & RstVent("id") & ")) " _
        + vbCr + "ORDER BY vta_ventas.fchdoc"
    
    RST_Busq xRsDoc, mSQL, xCon
    
'
'    RST_Busq xRsDoc, "SELECT vta_ventas.*, mae_cliente.nombre, mae_cliente.dir, mae_cliente.numruc, mae_moneda.descripcion AS mon, " _
'        & " [vta_conceptonc]![descripcion] & ' : REF A => ' & [mae_documento]![abrev] & ' ' & [vta_ventas_1]![numser] & '-' & [vta_ventas_1]![numdoc] AS docref2 " _
'        & " FROM (((mae_moneda RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON mae_moneda.id = vta_ventas.idmon) " _
'        & " LEFT JOIN vta_conceptonc ON vta_ventas.idmotnotcre = vta_conceptonc.id) LEFT JOIN vta_ventas AS vta_ventas_1 ON vta_ventas.iddocref = vta_ventas_1.id) " _
'        & " LEFT JOIN mae_documento ON vta_ventas_1.tipdoc = mae_documento.id Where (((vta_ventas.id) = " & RstVent("id") & ")) ORDER BY vta_ventas.fchdoc", xCon
'
    ' CARGAMOS EL DETALLE DE LA VENTA
    RST_Busq xRsDet, "SELECT vta_ventasdet.*, alm_inventario.codpro, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.idcuentaven, " _
        & " alm_inventario.idtipven FROM mae_unidades RIGHT JOIN (alm_inventario RIGHT JOIN vta_ventasdet " _
        & " ON alm_inventario.id = vta_ventasdet.iditem) ON mae_unidades.id = alm_inventario.idunimed " _
        & " WHERE (((vta_ventasdet.idvta)=" & RstVent("id") & "))", xCon

    ' BUSCAMOS SI EL DOCUMENTO TIENES UNA PLANTILLA DE IMPRESION ASIGNADA
    RST_Busq RsPDoc, "SELECT * FROM var_plantilladoc WHERE tipdoc = " & xRsDoc("tipdoc") & " ", xCon
    
    ' OBTENEMOS LOS NUMEROS DE LAS GUIAS DE EMISION, EN CASO DE TENER ASIGNADA UNA O VARIAS GUIAS
    RST_Busq RstGui, "SELECT vta_guia.numser, vta_guia.numdoc From vta_guia WHERE (((vta_guia.iddocven)=" & RstVent("id") & "))" _
        & " ORDER BY [vta_guia]![numser]+'-'+[vta_guia]![numdoc]", xCon

    If RstGui.RecordCount <> 0 Then
        ' SI EXISTEN GUIAS ASIGNADAS
        RstGui.MoveFirst
        xCadGuias = ""
                
        ' ALMACENAMOS LOS NUMEROS DE GUIA EN LA VARIABLE xCadGuias
        For A = 1 To RstGui.RecordCount
            xCadGuias = xCadGuias + Format(NulosN(RstGui("numser")), "0000") + "-" + Format(NulosN(RstGui("numdoc")), "0000000000")
            'xCadGuias = xCadGuias + Trim(Str(NulosN(RstGui("numdoc"))))
            RstGui.MoveNext
            If RstGui.EOF = True Then
                Exit For
            End If
            xCadGuias = xCadGuias + ", "
        Next A
    End If
    Set RstGui = Nothing
    
    ' SI NO EXISTE PLANTILLA DE IMPRESION EMITIMOS UNA ALERTA
    If RsPDoc.RecordCount = 0 Then
        MsgBox "No se ha definido la plantilla de impresion para este tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set xRsDoc = Nothing
        Set xRsDet = Nothing
        Set RsPDoc = Nothing
        Exit Sub
    End If
    
    RST_Busq RsPCab, "SELECT * FROM var_plantilladoc WHERE tipdoc = " & RsPDoc("tipdoc") & " ", xCon
    ' CARGAMOS LAS PLANTILLAS DE IMPRESION DEL DOCUMENTO
    If RsPCab.RecordCount <> 0 Then
        A = RsPCab("id")
        RST_Busq RsPCab, "SELECT * FROM var_plantillacab WHERE idplan = " & A & " ORDER BY item", xCon
        RST_Busq RsPDet, "SELECT * FROM var_plantilladet WHERE idplan = " & A & " ORDER BY item", xCon
    End If
   
    ' CONFIGURAMOS TIPO Y TAMAÑO DE LETRA
'    Printer.Font = "Super Draft 15cpi"
'    Printer.FontBold = False
'    Printer.FontSize = 11
'    Printer.ScaleMode = 6
    
    RsPDoc.MoveFirst
    Printer.Font = NulosC(RsPDoc("tipoletra"))
    Printer.ScaleMode = 6
    
    Dim xCam, xFor As String
    Dim RstAux As New ADODB.Recordset
    Dim TasaImpuesto As String
    
    Set RstAux = BuscaConCriterio("SELECT mae_impuestos.tasa from mae_impuestos WHERE mae_impuestos.id = 1 ", xCon)
    If RstAux.RecordCount = 1 Then
        TasaImpuesto = NulosN(RstAux("tasa"))
    End If

    ' imprime LA cabezera DEL DOCUMENTO
    Do While RsPCab.EOF = False
        xCam = RsPCab("campo")
        xFor = NulosC(RsPCab("formato"))
        
        Printer.CurrentX = RsPCab("posx")
        Printer.CurrentY = RsPCab("posy")
        
        If NulosC(UCase(xCam)) <> UCase("x-numeletra") _
                And NulosC(UCase(xCam)) <> UCase("x-idconpag") _
                And NulosC(UCase(xCam)) <> UCase("x-numguia") _
                And NulosC(UCase(xCam)) <> UCase("x-docref") Then
                
            'Printer.Print Format((NulosC(xRsDoc(xCam))), xFor)
            
            Printer.FontSize = NulosN(RsPCab("tamanho"))
            Printer.FontBold = NulosN(RsPCab("negrita"))
            F.PrintText Printer, Format(NulosC(xRsDoc(xCam)), NulosC(RsPCab("formato"))), NulosN(RsPCab("alineacion"))
            
            ' Valor de Venta
            If NulosC(UCase(xCam)) = UCase("impbru") Then
                Printer.CurrentX = RsPCab("posx") - 4
                Printer.CurrentY = RsPCab("posy") - 5
                Printer.Print "VENTA"
            End If
            ' IGV
            If NulosC(UCase(xCam)) = UCase("impigv") Then
                Printer.CurrentX = RsPCab("posx") - 4
                Printer.CurrentY = RsPCab("posy") - 5
                Printer.Print "I.G.V.(" + TasaImpuesto + "%)"
            End If
        Else
            ' Condicion de Pago
            If NulosC(UCase(xCam)) = UCase("x-idconpag") Then
                'Printer.Print Busca_Codigo(xRsDoc("idconpag"), "id", "descripcion", "mae_condpago", "N", xCon)
                
                Printer.FontSize = NulosN(RsPCab("tamanho"))
                Printer.FontBold = NulosN(RsPCab("negrita"))
                F.PrintText Printer, Busca_Codigo(xRsDoc("idconpag"), "id", "descripcion", "mae_condpago", "N", xCon), NulosN(RsPCab("alineacion"))
            
            End If
            If NulosC(UCase(xCam)) = UCase("x-numeletra") Then
                'Printer.Print "Son : "; NumeroLetra(xRsDoc("imptotdoc"), xRsDoc("idmon"))
                
                Printer.FontSize = NulosN(RsPCab("tamanho"))
                Printer.FontBold = NulosN(RsPCab("negrita"))
                F.PrintText Printer, "Son : " & NumeroLetra(xRsDoc("imptotdoc"), xRsDoc("idmon")), NulosN(RsPCab("alineacion"))
            
            End If
            If NulosC(UCase(xCam)) = UCase("x-numguia") Then
                'Printer.Print xCadGuias
                
                Printer.FontSize = NulosN(RsPCab("tamanho"))
                Printer.FontBold = NulosN(RsPCab("negrita"))
                F.PrintText Printer, xCadGuias, NulosN(RsPCab("alineacion"))
            
            End If
            If NulosC(UCase(xCam)) = UCase("x-docref") Then
                'Printer.Print "MOTIVO ==> "; xRsDoc("docref2")
                
                Printer.FontSize = NulosN(RsPCab("tamanho"))
                Printer.FontBold = NulosN(RsPCab("negrita"))
                F.PrintText Printer, "MOTIVO ==> " & xRsDoc("docref2"), NulosN(RsPCab("alineacion"))
            
            End If
        End If
        
        RsPCab.MoveNext
    Loop

    ' imprime EL detalle DEL DOCUMENTO
    Dim Fila As Integer
    Dim RstDetItem As New ADODB.Recordset
    Dim J As Integer
    
    RST_Busq RstDetItem, "SELECT vta_ventasdetitems.* From vta_ventasdetitems WHERE (((vta_ventasdetitems.idventa)=" & RstVent("id") & "))", xCon
    
    Fila = RsPDet("posy")
    xRsDet.MoveFirst
    Do While xRsDet.EOF = False
        
        RstDetItem.Filter = adFilterNone
        RstDetItem.Filter = "iditem = " & xRsDet("iditem") & ""
        
        Dim RstItemDoc As New ADODB.Recordset
        If RstDetItem.RecordCount <> 0 Then
            ' IMPRIMIMOS EL ITEM
            RsPDet.MoveFirst
            Do While RsPDet.EOF = False
                xCam = RsPDet("campo")
                xFor = NulosC(RsPDet("formato"))
                Printer.CurrentX = RsPDet("posx")
                Printer.CurrentY = Fila
                If xFor = "" Then
                    'Printer.Print NulosC(xRsDet(xCam))
                    
                    Printer.FontSize = NulosN(RsPDet("tamanho"))
                    Printer.FontBold = NulosN(RsPDet("negrita"))
                    F.PrintText Printer, NulosC(xRsDet(xCam)), NulosN(RsPDet("alineacion"))
            
                Else
                    'Printer.Print Format((NulosC(xRsDet(xCam))), xFor)
                    
                    Printer.FontSize = NulosN(RsPDet("tamanho"))
                    Printer.FontBold = NulosN(RsPDet("negrita"))
                    F.PrintText Printer, Format((NulosC(xRsDet(xCam))), xFor), NulosN(RsPDet("alineacion"))
            
                End If
                RsPDet.MoveNext
            Loop
            Fila = Fila + 4
            
            ' IMPRIMIMOS EL DETALLE DEL ITEM
            RsPDet.MoveFirst
            Do While RsPDet.EOF = False
                ' Si es el campo descripcion
                If RsPDet("campo") = "descripcion" Then
                    Dim mDetalleItem As String
                    
                    xCam = "texto"
                    xFor = ""
                    Printer.CurrentX = RsPDet("posx")
                    
                    '*************************
                    ' Si esta configurado el item para mostrar el nombre tecnico
                    RST_Busq RstItemDoc, "SELECT mae_itemdocconfig.* From mae_itemdocconfig WHERE (((mae_itemdocconfig.iditem)=" & NulosN(xRsDet("iditem")) & ") AND ((mae_itemdocconfig.iddoc)=" & NulosN(RstVent("tipdoc")) & "))", xCon
                    If RstItemDoc.RecordCount > 0 Then
                        mDetalleItem = F.BuscaCodigoTabla(NulosN(xRsDet("iditem")), "id", "desctec", "alm_inventario", "N", xCon)
                    Else
                        mDetalleItem = NulosC(RstDetItem(xCam))
                    End If
                    '*************************
                    
                    For J = 1 To RstDetItem.RecordCount
                        Printer.CurrentX = RsPDet("posx")
                        Printer.CurrentY = Fila
                        If xFor = "" Then
                            'Printer.Print NulosC(RstDetItem(xCam))
                            
                            Printer.FontSize = NulosN(RsPDet("tamanho"))
                            Printer.FontBold = NulosN(RsPDet("negrita"))
                            F.PrintText Printer, mDetalleItem, NulosN(RsPDet("alineacion"))
            
                        Else
                            'Printer.Print Format((NulosC(RstDetItem(xCam))), xFor)
                            
                            Printer.FontSize = NulosN(RsPDet("tamanho"))
                            Printer.FontBold = NulosN(RsPDet("negrita"))
                            F.PrintText Printer, Format((mDetalleItem), xFor), NulosN(RsPDet("alineacion"))
            
                        End If
                        
                        RstDetItem.MoveNext
                        If RstDetItem.EOF = True Then
                            Fila = Fila + 4
                            Exit For
                        End If
                        
                        Fila = Fila + 4
                    Next J
                End If
                RsPDet.MoveNext
            Loop
        Else
            RsPDet.MoveFirst
            Do While RsPDet.EOF = False
                xCam = RsPDet("campo")
                
                xFor = NulosC(RsPDet("formato"))
                Printer.CurrentX = RsPDet("posx")
                Printer.CurrentY = Fila
                
                '*************************
                ' Si es el campo descripcion
                If RsPDet("campo") = "descripcion" Then
                    ' Si esta configurado el item para mostrar el nombre tecnico
                    RST_Busq RstItemDoc, "SELECT mae_itemdocconfig.* From mae_itemdocconfig WHERE (((mae_itemdocconfig.iditem)=" & NulosN(xRsDet("iditem")) & ") AND ((mae_itemdocconfig.iddoc)=" & NulosN(RstVent("tipdoc")) & "))", xCon
                    If RstItemDoc.RecordCount > 0 Then
                        mDetalleItem = F.BuscaCodigoTabla(NulosN(xRsDet("iditem")), "id", "desctec", "alm_inventario", "N", xCon)
                    Else
                        mDetalleItem = NulosC(xRsDet(xCam))
                    End If
                    If xFor = "" Then
                        Printer.FontSize = NulosN(RsPDet("tamanho"))
                        Printer.FontBold = NulosN(RsPDet("negrita"))
                        F.PrintText Printer, mDetalleItem, NulosN(RsPDet("alineacion"))
                
                    Else
                        Printer.FontSize = NulosN(RsPDet("tamanho"))
                        Printer.FontBold = NulosN(RsPDet("negrita"))
                        F.PrintText Printer, Format((mDetalleItem), xFor), NulosN(RsPDet("alineacion"))
                    End If
                '*************************
                Else
                    If xFor = "" Then
                        'Printer.Print NulosC(xRsDet(xCam))
                        
                        Printer.FontSize = NulosN(RsPDet("tamanho"))
                        Printer.FontBold = NulosN(RsPDet("negrita"))
                        F.PrintText Printer, NulosC(xRsDet(xCam)), NulosN(RsPDet("alineacion"))
                
                    Else
                        'Printer.Print Format((NulosC(xRsDet(xCam))), xFor)
                        
                        Printer.FontSize = NulosN(RsPDet("tamanho"))
                        Printer.FontBold = NulosN(RsPDet("negrita"))
                        F.PrintText Printer, Format((NulosC(xRsDet(xCam))), xFor), NulosN(RsPDet("alineacion"))
                
                    End If
                End If
                RsPDet.MoveNext
            Loop
            Fila = Fila + 4
        End If
        
        xRsDet.MoveNext
    Loop
    ' ENVIA LOS DATOS A LA IMPRESORA
    Printer.EndDoc
End Sub

Private Sub CmdSel_Click()
    Fg3.Rows = 0
    Dim xRst As New ADODB.Recordset
    Dim A As Integer
    
    RST_Busq xRst, "SELECT * FROM var_plantilladoc WHERE tipdoc = " & NulosN(TxtTipDoc.Text) & "", xCon
    If xRst.RecordCount = 0 Then
        MsgBox "No se ha encontrado la plantilla de impresion para el documento " & LblNomDoc, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set xRst = Nothing
        Exit Sub
    End If
    
    LblNumLineas.Caption = (xRst("numitem") - 1)
    Fg3.Rows = (xRst("numitem") - 1)
    Set xRst = Nothing
    RstVentItemsDeta.Filter = adFilterNone
    If RstVentItemsDeta.RecordCount <> 0 Then
        RstVentItemsDeta.Filter = "iditem = " & Fg1.TextMatrix(Fg1.Row, 8) & ""
        
        If RstVentItemsDeta.RecordCount <> 0 Then
            RstVentItemsDeta.Sort = "orden"
            RstVentItemsDeta.MoveFirst
            For A = 0 To RstVentItemsDeta.RecordCount
                Fg3.TextMatrix(A, 1) = NulosC(RstVentItemsDeta("texto"))
                RstVentItemsDeta.MoveNext
                If RstVentItemsDeta.EOF = True Then Exit For
            Next A
        End If
    End If
    
    FrameDetalleItem.Left = 1080
    FrameDetalleItem.Top = 1695
    
    If QueHace = 3 Then
        Fg3.Editable = flexEDNone
        Fg3.SelectionMode = flexSelectionByRow
    Else
        Fg3.Editable = flexEDKbdMouse
        Fg3.SelectionMode = flexSelectionFree
    End If
    
    FrameDetalleItem.Visible = True
    
    
    
    
'    ' PERMITE SELECCIONAR UNO O MAS ITEMS
'    If NulosC(TxtTipItem.Text) = "" Then
'        MsgBox "No ha especificado el tipo de venta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'        TxtTipItem.SetFocus
'        Exit Sub
'    End If
'
'    Dim xfrm As New eps_librerias.FormSeleccion
'    Dim xCampos(3, 5) As String
'    Dim xRs As New ADODB.Recordset
'    Dim A As Integer
'
'    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "4000":         xCampos(0, 3) = "C":    xCampos(0, 4) = "N"
'    xCampos(1, 0) = "Uni. Med":       xCampos(1, 1) = "abrev":          xCampos(1, 2) = "1000":        xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
'    xCampos(2, 0) = "Codigo":         xCampos(2, 1) = "codpro":        xCampos(2, 2) = "1200":         xCampos(2, 3) = "C":    xCampos(2, 4) = "S"
'
'    Dim nSQLId As String
'    nSQLId = GRID_GENERAR_SQL_ID(Fg1, 8, "alm_inventario.id", " NOT IN ", True)
'    If nSQLId <> "" Then nSQLId = " AND " & nSQLId
'
'    xfrm.SQLCad = "SELECT alm_inventario.*, mae_unidades.descripcion AS descuni, mae_unidades.abrev, mae_percepcion.tasa " _
'        & " FROM mae_unidades RIGHT JOIN (mae_percepcion RIGHT JOIN alm_inventario ON mae_percepcion.id = alm_inventario.idper) " _
'        & " ON mae_unidades.id = alm_inventario.idunimed Where (((alm_inventario.tippro) = " & NulosN(TxtTipItem) & " )) " & nSQLId & " ORDER BY alm_inventario.descripcion"
'
'    xfrm.Titulo = "Buscando Productos"
'
'    Set xfrm.Coneccion = xCon
'    Set xRs = xfrm.Seleccionar(xCampos)
'    If xRs.State = 1 Then
'        If xRs.RecordCount <> 0 Then
'            Dim X As Integer
'            Dim Agregar As Boolean
'            Dim xPrecio As Double
'            Agregar = True
'            Mostrando = True
'            xRs.MoveFirst
'
'            ' AGREGAMOS LOS ITEMS AL CONTROL FlexGrid Fg1
'            For X = 1 To xRs.RecordCount
'                For A = 1 To Fg1.Rows - 1
'                    If Fg1.TextMatrix(A, 6) = xRs("id") Then
'                        Agregar = False
'                    End If
'                Next A
'                If Agregar = True Then
'                    Fg1.Rows = Fg1.Rows + 1
'                    xPrecio = 0
'                    Fg1.TextMatrix(Fg1.Rows - 1, 1) = NulosC(xRs("descripcion"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 2) = NulosC(xRs("abrev"))
'
'                    If NulosN(LblIdCliente.Caption) <> 0 Then
'                        xPrecio = UltimoPrecio(NulosN(xRs("id")), NulosN(LblIdCliente.Caption))
'                    Else
'                        xPrecio = UltimoPrecio(NulosN(xRs("id")), 0)
'                    End If
'                    Fg1.TextMatrix(Fg1.Rows - 1, 4) = Format(xPrecio, "0.0000")
'
'                    Fg1.TextMatrix(Fg1.Rows - 1, 8) = xRs("id")
'                    Fg1.TextMatrix(Fg1.Rows - 1, 9) = NulosN(xRs("idunimed"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 10) = NulosN(xRs("idcuentaven"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 11) = NulosN(xRs("idtipven"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 12) = NulosN(xRs("tasa"))
'                    Fg1.TextMatrix(Fg1.Rows - 1, 13) = NulosN(xRs("stckact"))
'
'                End If
'                xRs.MoveNext
'                If xRs.EOF = True Then
'                    Exit For
'                End If
'                Agregar = True
'            Next X
'            Mostrando = False
'        End If
'    End If
'    Set xfrm = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : ModificarSaldo
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTIVA EL FRAME Frame8
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub ModificarSaldo()
    ActivarEntorno
    Frame8.Top = 2580
    Frame8.Left = 2760
    Frame8.Visible = True
    
    TxtNumDoc2.Text = ""
    TxtCliente2.Text = ""
    TxtSaldo2.Text = ""
    TxtNewSaldo2.Text = ""
    
    TxtNumDoc2.Text = NulosC(RstVent("numerodoc"))
    TxtCliente2.Text = NulosC(RstVent("nombre"))
    TxtSaldo2.Text = Format(NulosN(RstVent("impsal")), FORMAT_MONTO)
    TxtNewSaldo2.SetFocus
    
End Sub

'*****************************************************************************************************
'* Nombre           : CambiarMes
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CAMBIA EL MES ACTUAL DE TRABAJO
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub CambiarMes()
    TabOne1.CurrTab = 0
    mMesActivo = SeleccionaMes(xCon)
    pCargarGrid
End Sub

'*****************************************************************************************************
'* Nombre           : pCargarGrid
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LOS REGISTROS DE LA TABLA vta_ventas EN EL RECORDSET RstVent, ESTOS DATOS
'*                    SE VISUALIZARAN EN LA PESTAÑA DETALLE
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pCargarGrid()
    Dim nSQL  As String
    Dim Rpta As Integer
    Dim DiaIniAño  As String
    Dim xFechaRegistro As String
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    LblPeriodo2.Caption = LblMes.Caption
    DiaIniAño = "01/01/" + Trim(AnoTra)
    xFechaRegistro = "01/" + Format(mMesActivo, "00") + "/" + Trim(AnoTra)
    
'    If mMesActivo = 0 Then
'        ' SI ES EL MES DE APERTURA
'        nSQL = "SELECT vta_ventas.*, mae_cliente.nombre, IIf(IsNull([vta_ventas]![numser])=-1,[vta_ventas]![numdoc],[vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]) AS numerodoc, IIf(vta_ventas.Anulado=0,'Facturado','Anulado') AS EstadoVenta, " _
'            & " mae_documento.descripcion AS nomdoc, mae_condpago.descripcion AS desccond, mae_documento.abrev, mae_cliente.numruc, mae_moneda.descripcion AS descmon, mae_moneda.simbolo, " _
'            & " mae_impuestos.idcuenvta, con_tc.impcom, mae_tipoproducto.descripcion AS desctipcom, mae_condpago.abrev AS conpagabre, Mid([vta_ventas].[numreg],1,2)+[mae_libros].[codsun]+Mid([vta_ventas].[numreg],3,4) AS numreg1, " _
'            & " vta_ventas.fchdoc & '' as fchdoc1,vta_ventas.fchven & '' as fchven1,vta_ventas.impbru & '' as impbru1, vta_ventas.impigv & '' as impigv1 ,vta_ventas.imptotdoc & '' as imptotdoc1, vta_ventas.impsal & '' as impsal1, " _
'            & " IIF(vta_ventas.anulado=-1,0,IIf([vta_ventas].[tc]=0,[con_tc].[impven],[vta_ventas].[tc])) & '' AS impven1 " _
'            & " FROM ((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN ((mae_documento LEFT JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) RIGHT JOIN (mae_condpago RIGHT " _
'            & " JOIN (vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_condpago.id = vta_ventas.idconpag) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) " _
'            & " ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_tipoproducto ON vta_ventas.idtipo = mae_tipoproducto.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
'            & " WHERE vta_ventas.numreg LIKE '" & Format(mMesActivo, "00") & "%' ORDER BY [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc] DESC"
            '(((vta_ventas.fchdoc)<CDate('" & DiaIniAño & "')))
'    ElseIf mMesActivo < 13 Then
        ' SI ES UN MES ENTRE ENERO Y DICIEMBRE
        'nSQL = "SELECT vta_ventas.*, IIf(vta_ventas.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, IIf(IsNull([vta_ventas]![numser])=-1,[vta_ventas]![numdoc], " _
            & " [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]) AS numerodoc, IIf(vta_ventas.Anulado=0,'Facturado','Anulado') AS EstadoVenta, " _
            & " mae_documento.descripcion AS nomdoc, IIf(vta_ventas.anulado=-1,'',mae_condpago.descripcion) AS desccond, mae_documento.abrev, " _
            & " mae_cliente.numruc, mae_moneda.descripcion AS descmon, IIf(vta_ventas.anulado=-1,'',mae_moneda.simbolo) AS simbolo, mae_impuestos.idcuenvta, " _
            & " con_tc.impcom, mae_tipoproducto.descripcion AS desctipcom, IIf(vta_ventas.anulado=-1,'',mae_condpago.abrev) AS conpagabre, " _
            & " Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4) AS numreg1, vta_ventas.fchdoc AS fchdoc1, vta_ventas.fchven AS fchven1, " _
            & " vta_ventas.impbru AS impbru1, vta_ventas.impigv AS impigv1, vta_ventas.imptotdoc AS imptotdoc1, vta_ventas.impsal AS impsal1, " _
            & " IIf([vta_ventas].[anulado]=-1,0,IIf([vta_ventas].[tc]=0,[con_tc].[impven],[vta_ventas].[tc])) AS impven1 " _
            & " FROM ((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN ((mae_documento LEFT JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) " _
            & " RIGHT JOIN (mae_condpago RIGHT JOIN (vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_condpago.id = vta_ventas.idconpag) " _
            & " ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_tipoproducto " _
            & " ON vta_ventas.idtipo = mae_tipoproducto.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
            & " WHERE (((vta_ventas.fchreg)=CDate('" & xFechaRegistro & "')) AND ((vta_ventas.fchdoc)>=CDate('" & DiaIniAño & "'))) " _
            & " ORDER BY vta_ventas!numser+'-'+vta_ventas!numdoc DESC"
    
    If mMesActivo <= 12 Then
        
        nSQL = "SELECT vta_ventas.*, IIf(vta_ventas.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, IIf(IsNull([vta_ventas]![numser])=-1,[vta_ventas]![numdoc], " _
            & " [vta_ventas]![numser]+'-'+[vta_ventas]![numdoc]) AS numerodoc, IIf(vta_ventas.Anulado=0,'Facturado','Anulado') AS EstadoVenta,  " _
            & " mae_documento.descripcion AS nomdoc, IIf(vta_ventas.anulado=-1,'', mae_condpago.descripcion) AS desccond,mae_documento.abrev, " _
            & " mae_cliente.numruc, mae_moneda.descripcion AS descmon, IIf(vta_ventas.anulado=-1,'',mae_moneda.simbolo) AS simbolo, mae_impuestos.idcuenvta, " _
            & " con_tc.impcom, mae_tipoproducto.descripcion AS desctipcom, IIf(vta_ventas.anulado=-1,'',mae_condpago.abrev) AS conpagabre, " _
            & " Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4) AS numreg1,vta_ventas.fchdoc & '' as fchdoc1,vta_ventas.fchven & '' as fchven1, " _
            & " vta_ventas.impbru & '' as impbru1, vta_ventas.impigv & '' as impigv1 ,vta_ventas.imptotdoc & '' as imptotdoc1, vta_ventas.impsal & '' as impsal1, " _
            & " IIF(vta_ventas.anulado=-1,0,IIf([vta_ventas].[tc]=0,[con_tc].[impven],[vta_ventas].[tc])) & '' AS impven1 " _
            & " FROM ((mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN ((mae_documento LEFT JOIN mae_impuestos ON mae_documento.idimp = mae_impuestos.id) " _
            & " RIGHT JOIN (mae_condpago RIGHT JOIN (vta_ventas LEFT JOIN con_tc ON vta_ventas.fchdoc = con_tc.fecha) ON mae_condpago.id = vta_ventas.idconpag) " _
            & " ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_tipoproducto " _
            & " ON vta_ventas.idtipo = mae_tipoproducto.id) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
            & " WHERE vta_ventas.numreg LIKE '" & Format(mMesActivo, "00") & "%'  " _
            & " ORDER BY vta_ventas!numser+'-'+vta_ventas!numdoc DESC"
            
            '(((vta_ventas.fchreg)=CDate('" & xFechaRegistro & "')) AND ((vta_ventas.fchdoc)>=CDate('" & DiaIniAño & "')))
            
    Else
        MsgBox "Ha selecionado el mes de Cierre, selecciones meses comprendidos entre Enero y Diciembre", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Set RstVent = Nothing
        Set Dg1.DataSource = Nothing
        Dg1.Refresh
        Exit Sub
    End If
    
    TDB_FiltroLimpiar Dg1
    Set RstVent = Nothing
    
    ' cargando datos
    Me.MousePointer = vbHourglass
    RST_Busq RstVent, nSQL, xCon

    Set Dg1.DataSource = RstVent
    
    Me.MousePointer = vbDefault
    
    OpcionesPeriodo
    TabOne1.CurrTab = 0
    
End Sub

'*****************************************************************************************************
'* Nombre           : Buscar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : EJECUTA UNA BUSQUEDA EN EL RECORDSET RstVent
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Sub Buscar()
    TabOne1.CurrTab = 0
    Dim xRs As New ADODB.Recordset
    
    Dim nSQL As String
    Dim xCampos(8, 4) As String
    
    xCampos(0, 0) = "NumReg":        xCampos(0, 1) = "registro":     xCampos(0, 2) = "820":   xCampos(0, 3) = "C"
    xCampos(1, 0) = "T.D.":          xCampos(1, 1) = "abrev":      xCampos(1, 2) = "400":   xCampos(1, 3) = "C"
    xCampos(2, 0) = "N°. Documento": xCampos(2, 1) = "numerodoc":  xCampos(2, 2) = "1400":  xCampos(2, 3) = "C"
    xCampos(3, 0) = "FchEmi":        xCampos(3, 1) = "fchdoc":     xCampos(3, 2) = "830":   xCampos(3, 3) = "C"
    xCampos(4, 0) = "FchVenc":       xCampos(4, 1) = "fchven":     xCampos(4, 2) = "830":   xCampos(4, 3) = "C"
    xCampos(5, 0) = "Cliente":       xCampos(5, 1) = "nombre":     xCampos(5, 2) = "2600":  xCampos(5, 3) = "C"
    
    xCampos(6, 0) = "M":             xCampos(6, 1) = "simbolo":    xCampos(6, 2) = "450":    xCampos(6, 3) = "C"
    xCampos(7, 0) = "Importe":         xCampos(7, 1) = "imptotdoc":     xCampos(7, 2) = "850":    xCampos(7, 3) = "N"
    
    nSQL = "SELECT vta_ventas.id,Mid([numreg],1,2)+[mae_libros].[codsun]+Mid([numreg],3,4) AS registro, IIf(vta_ventas.anulado=-1,'ANULADO',mae_cliente.nombre) AS nombre, vta_ventas!numser+'-'+vta_ventas!numdoc AS numerodoc, mae_documento.abrev, IIf(vta_ventas.anulado=-1,'',mae_moneda.simbolo) AS simbolo, format(vta_ventas.fchdoc,'dd/mm/yy') as fchdoc, format(vta_ventas.fchven,'dd/mm/yy') as fchven, vta_ventas.imptotdoc " _
        + vbCr + " FROM (mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN vta_ventas ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
        + vbCr + " WHERE (((vta_ventas.numreg) Like '" & Format(mMesActivo, "00") & "%')) " _

    CARGAR_DLL_EPSBUSCAR xCon, xRs, nSQL, xCampos(), "Buscando Compras", "nombre", "nombre", Principio

    If xRs.State = 1 Then
        RstVent.MoveFirst
        RstVent.Find "id = " & xRs("id") & ""
    End If
    Set xRs = Nothing
End Sub

'*****************************************************************************************************
'* Nombre           : ActualizaSaldoDoc
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : ACTUALIZA EL IMPORTE DEL DOCUMENTO ESPECIFICADO
'* Paranetros       : NOMBRE       |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    idDocumento  |  INTEGER   |  ID DEL DOCUMENTO
'*                    Tabla        |  INTEGER   |  ID DE LA TABLA
'*                    ImporteRestar|  SOUBLE    |  IMPORTE A RESTAR
'* Devuelve         :
'*****************************************************************************************************
Sub ActualizaSaldoDoc(idDocumento As Double, Tabla As Integer, ImporteRestar As Double)
    '1 = compras
    '2 = Ventas
    '3 = honorarios
    
    Dim Rst As New ADODB.Recordset
    Dim Total As Double
    
    If Tabla = 2 Then
        RST_Busq Rst, "SELECT Sum(tes_cajadestinodet.acuenta) AS total FROM tes_caja LEFT JOIN tes_cajadestinodet ON tes_caja.id = tes_cajadestinodet.idtes " _
            & " GROUP BY tes_cajadestinodet.iddoc, tes_caja.tipmov HAVING (((tes_cajadestinodet.iddoc)=" & idDocumento & ") AND ((tes_caja.tipmov)=1))", xCon
            
        Total = BuscaImporteDocumento(idDocumento, 1)
        
    End If
    
    If Rst.RecordCount <> 0 Then
        Total = ((Total - Rst("total")) - ImporteRestar)
    Else
        Total = (Total - ImporteRestar)
    End If
    
    xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = " & Total & " WHERE (((vta_ventas.id)=" & idDocumento & "))"
    
    Set Rst = Nothing
        
End Sub

'*****************************************************************************************************
'* Nombre           : BuscaImporteDocumento
'* Tipo             : FUNCION
'* Descripcion      : DEVUELVE EL IMPORTE DEL DOCUMENTO ESPECIFICADO, DEVUELVE UN ENTERO DOBLE
'* Paranetros       : NOMBRE      |  TIPO     |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    idDocumento |  INTEGER  |  ESPECIFICA EL ID DEL DOCUMENTO
'*                    Tabla       |  TABLA    |  ESPECIFICA EL ID DE LA TABLA
'* Devuelve         :
'*****************************************************************************************************
Function BuscaImporteDocumento(idDocumento As Double, Tabla As Integer) As Double
    '1 = compras
    '2 = Ventas
    '3 = honorarios
    Dim Rst As New ADODB.Recordset
    
    'compras
    If Tabla = 1 Then RST_Busq Rst, "SELECT * FROM vta_ventas WHERE id = " & idDocumento & "", xCon
    
    If Rst.RecordCount <> 0 Then
        BuscaImporteDocumento = Rst("imptotdoc")
    Else
        BuscaImporteDocumento = 0
    End If
    
    Set Rst = Nothing
End Function

'*****************************************************************************************************
'* Nombre           : pGridConfigurar
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CONFIGURA EL CONTROL FlexGrid Fg1
'* Paranetros       :
'* Devuelve         :
'*****************************************************************************************************
Private Sub pGridConfigurar()
    If NulosN(TxtTipItem.Text) = 5 Then
        Fg1.ColWidth(2) = 0
        Fg1.ColWidth(3) = 0
        Fg1.ColWidth(4) = 1100
        Fg1.ColWidth(5) = 1100
        Fg1.ColWidth(7) = 1200
    Else
        Fg1.ColWidth(2) = 435
        Fg1.ColWidth(3) = 855
        Fg1.ColWidth(4) = 930
        Fg1.ColWidth(5) = 960
        Fg1.ColWidth(7) = 1020
    End If
End Sub

Private Sub ChkTC_Click()
    ' ACTIVA O DESACTIVA EL INGRESO DEL TIPO DE CAMBIO
    If QueHace = 3 Then Exit Sub
    If ChkTC.Value = 0 Then
        TxtTC.BackColor = &H8000000F
        TxtTC.Enabled = False
        If IsDate(TxtFchDoc.Valor) = True Then
            TxtTC.Text = HallaTipoCambio(TxtFchDoc.Valor, 2, Venta, xCon)
        Else
            Exit Sub
        End If
    Else
        TxtTC.Enabled = True
        TxtTC.BackColor = vbWhite
        TxtTC.SetFocus
    End If
End Sub

Private Sub Command3_Click()
    '--grabar operacion en analisis cta cte savar
    Dim A As Integer
    TabOne1.CurrTab = 0
    RstVent.MoveFirst
'    Dim xCodSunLib  As String
'    Dim xTc As Double
'    xCodSunLib = Busca_Codigo(1 , "id", "codsun", "mae_libros", "N", xCon)
    For A = 1 To RstVent.RecordCount
        
'        If NulosN(RstVent("tc")) = 0 Then
'            xTc = HallaTipoCambio(RstVent("fchdoc"), 2, Venta, xCon)
'        Else
'            xTc = RstVent("tc")
'        End If
'        If RstVent("tipdoc") <> 7 Then
'            If NulosN(RstVent("idmon")) = 1 Then
'                GrabarOperacionCtaCteDocRef 1, RstVent("id"), NulosC(RstVent("numerodocref")), RstVent("idcli"), RstVent("tipdoc"), RstVent("numerodoc"), _
'                    RstVent("fchdoc"), RstVent("idmon"), xTc, 0, RstVent("imptotdoc"), 0, 0, Format(xCodSunLib, "00") & RstVent("numreg"), xCon
'            Else
'                GrabarOperacionCtaCteDocRef 1, RstVent("id"), NulosC(RstVent("numerodocref")), RstVent("idcli"), RstVent("tipdoc"), RstVent("numerodoc"), _
'                    RstVent("fchdoc"), RstVent("idmon"), xTc, 0, 0, 0, RstVent("imptotdoc"), Format(xCodSunLib, "00") & RstVent("numreg"), xCon
'            End If
'        Else
'            If NulosN(RstVent("idmon")) = 1 Then
'                GrabarOperacionCtaCteDocRef 1, RstVent("id"), NulosC(RstVent("numerodocref")), RstVent("idcli"), RstVent("tipdoc"), RstVent("numerodoc"), _
'                    RstVent("fchdoc"), RstVent("idmon"), xTc, RstVent("imptotdoc"), 0, 0, 0, Format(xCodSunLib, "00") & RstVent("numreg"), xCon
'            Else
'                GrabarOperacionCtaCteDocRef 1, RstVent("id"), NulosC(RstVent("numerodocref")), RstVent("idcli"), RstVent("tipdoc"), RstVent("numerodoc"), _
'                    RstVent("fchdoc"), RstVent("idmon"), xTc, 0, 0, RstVent("imptotdoc"), 0, Format(xCodSunLib, "00") & RstVent("numreg"), xCon
'            End If
'        End If
'
        GrabarOperacionCtaCte 2, RstVent("id"), xCon
         
        RstVent.MoveNext
        If RstVent.EOF = True Then Exit For
    Next A
    MsgBox "se termino con exito"
End Sub

Private Sub pExportar()
    
    
    Dim nSQL As String
    Dim oExport As New SGI2_funciones.formularios
    Dim RstTmp  As New ADODB.Recordset
    Dim xCampos(16, 3) As String
    
    TabOne1.CurrTab = 0
    
    '0::Nombre a Mostrar;
    '1::nombre de Campo del Rst;
    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
    '3::ancho de columna
    '--obs: el rst puede tener mas columnas solo se consideran los campos del array
    xCampos(0, 0) = "Id":           xCampos(0, 1) = "id":           xCampos(0, 2) = 2:   xCampos(0, 3) = "500"
    xCampos(1, 0) = "Nº Reg":       xCampos(1, 1) = "numreg1":      xCampos(1, 2) = 0:   xCampos(1, 3) = "900"
    xCampos(2, 0) = "R.U.C.":       xCampos(2, 1) = "numruc":       xCampos(2, 2) = 0:   xCampos(2, 3) = "1200"
    xCampos(3, 0) = "Cliente":      xCampos(3, 1) = "nombre":       xCampos(3, 2) = 0:   xCampos(3, 3) = "3290"
    xCampos(4, 0) = "T.D.":         xCampos(4, 1) = "abrev":        xCampos(4, 2) = 0:   xCampos(4, 3) = "350"
    xCampos(5, 0) = "Num. Doc":     xCampos(5, 1) = "numerodoc":    xCampos(5, 2) = 0:   xCampos(5, 3) = "1600"
    xCampos(6, 0) = "Fch.Emi":      xCampos(6, 1) = "fchdoc1":      xCampos(6, 2) = 1:   xCampos(6, 3) = "900"
    xCampos(7, 0) = "Fch. Venc":    xCampos(7, 1) = "fchven1":      xCampos(7, 2) = 1:   xCampos(7, 3) = "900"
    xCampos(8, 0) = "Glosa":        xCampos(8, 1) = "glosa":        xCampos(8, 2) = 0:   xCampos(8, 3) = "2000"
    xCampos(9, 0) = "M":            xCampos(9, 1) = "simbolo":      xCampos(9, 2) = 1:   xCampos(9, 3) = "500"
    xCampos(10, 0) = "T.C.":        xCampos(10, 1) = "impven1":     xCampos(10, 2) = 2:  xCampos(10, 3) = "700"
    xCampos(11, 0) = "Imp Afec":    xCampos(11, 1) = "impbru":      xCampos(11, 2) = 2:  xCampos(11, 3) = "900"
    xCampos(12, 0) = "Imp Inaf":    xCampos(12, 1) = "impinaf":     xCampos(12, 2) = 2:  xCampos(12, 3) = "900"
    xCampos(13, 0) = "Imp Igv":     xCampos(13, 1) = "impigv":      xCampos(13, 2) = 2:  xCampos(13, 3) = "900"
    xCampos(14, 0) = "Imp ISC":     xCampos(14, 1) = "impisc":      xCampos(14, 2) = 2:  xCampos(14, 3) = "1057"
    xCampos(15, 0) = "Imp Otros":   xCampos(15, 1) = "impotr":      xCampos(15, 2) = 2:  xCampos(15, 3) = "900"
    xCampos(16, 0) = "Imp Total":   xCampos(16, 1) = "imptotdoc":   xCampos(16, 2) = 2:  xCampos(16, 3) = "1000"
    
    Set RstTmp = RstVent.Clone
    
    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "LISTADO DE VENTAS", "Periodo " & LblMes.Caption, "", "Listado de Ventas - " & LblMes.Caption, RstTmp, xCampos
    Set oExport = Nothing
    Set RstTmp = Nothing
    
End Sub





'*---------------------------------------------------------------------------------------
'*--------------
'*---------------------------------------------------------------------------------------

Private Sub CmdBusAlm2_Click()
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
        LblAlmacen2.Caption = xRs("descripcion")
        TxtIdAlm2.Text = xRs("id")
        TxtTipDoc2.SetFocus
        
        If TxtTipDoc2.Text <> "" Then
            Dim Rst As New ADODB.Recordset
            Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc2.Text) & " AND idalm = " & NulosN(TxtIdAlm2.Text) & "", xCon)
            If Rst.RecordCount <> 0 Then
                TxtNumSer2.Text = NulosC(Rst("numser"))
                TxtNumSer2_Validate True
            End If
            
            Set Rst = Nothing
        Else
            TxtNumSer2.Text = ""
            TxtNumDoc2.Text = ""
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub TxtIdAlm2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtIdAlm2_Validate(Cancel As Boolean)
    ' VALIDA EL ID DEL ALMACEN
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtIdAlm2.Text) <> "" Then
        LblAlmacen2.Caption = Busca_Codigo(NulosN(TxtIdAlm2.Text), "id", "descripcion", "alm_almacenes", "N", xCon)
        If LblAlmacen2.Caption = "" Then
            TxtIdAlm2.Text = ""
        End If
    Else
        LblAlmacen2.Caption = ""
    End If
End Sub


Private Sub TxtIdAlm2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusAlm2_Click
    End If
End Sub


Private Sub CmdBusTipDoc2_Click()
    ' EJECUTA LA BUSQUEDA DE TIPO DE DOCUMENTO
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
    Dim xCampos(2, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Descripcion":    xCampos(0, 1) = "descripcion":      xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Codigo":         xCampos(1, 1) = "id":               xCampos(1, 2) = "1000":         xCampos(1, 3) = "N"
    
    xform.SQLCad = "SELECT DISTINCT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, " _
        & " mae_impuestos.Abrev AS abreimp, mae_impuestos.idcuenvta AS cuentaimp " _
        & " FROM alm_numseries LEFT JOIN (mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuenvta = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id) ON alm_numseries.idtipdoc = mae_documento.id WHERE (((alm_numseries.idalm)=" & NulosN(TxtIdAlm2.Text) & "))"

    Dim xImpuesto As Double
    
    xform.titulo = "Buscando Tipo de Documento"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtTipDoc2.Text = xRs("id")
            LblNomDoc2.Caption = NulosC(xRs("descripcion"))
            xIdCuenTasa = NulosN(xRs("cuentaimp"))
            
            TxtFchEmiAnul.SetFocus
            
            Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & NulosN(TxtTipDoc2.Text) & " and mae_documentocta.idmon =" & 1 & " and tipope = -1", xCon)
            If xRs2.RecordCount > 0 Then
                xCuentaDoc = NulosN(xRs2("idcuen"))
            Else
                MsgBox "No se ha encontrado cuenta contable para el documento especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                '--salir de la opcion de anular documentos
                cmdsalirseldoc_Click
            End If
            Set xRs2 = Nothing
            
        End If
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub


Private Sub TxtTipDoc2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    Else
        If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub TxtTipDoc2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusTipDoc2_Click
    End If
End Sub


Private Sub TxtTipDoc2_Validate(Cancel As Boolean)
    ' VALIDA EL TIPO DE DOCUMENTO
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtTipDoc2.Text) = "" Then
        LblNomDoc2.Caption = ""
        Exit Sub
    End If
    Dim xRs As New ADODB.Recordset
    Dim xRs2 As New ADODB.Recordset
       
    RST_Busq xRs, "SELECT mae_documento.*, mae_impuestos.tasa, mae_impuestos.descripcion AS descimp, con_planctas.cuenta, mae_impuestos.Abrev AS abreimp, " _
        & " mae_impuestos.idcuenvta AS cuentaimp, alm_numseries.numser, alm_numseries.idtipdoc " _
        & " FROM alm_numseries LEFT JOIN (mae_documento LEFT JOIN (mae_impuestos LEFT JOIN con_planctas ON mae_impuestos.idcuenvta = con_planctas.id) " _
        & " ON mae_documento.idimp = mae_impuestos.id) ON alm_numseries.idtipdoc = mae_documento.id " _
        & " WHERE (((alm_numseries.idalm)=" & NulosN(TxtIdAlm2.Text) & ") AND ((alm_numseries.idtipdoc)=" & NulosN(TxtTipDoc2.Text) & "))", xCon
        
    If xRs.RecordCount = 0 Then
        TxtTipDoc2.Text = ""
        LblNomDoc2.Caption = ""
        TxtNumDocGen.Text = ""
    Else
        TxtTipDoc2.Text = NulosN(xRs("id"))
        LblNomDoc2.Caption = NulosC(xRs("descripcion"))

        
        xIdCuenTasa = NulosN(xRs("cuentaimp"))
        

        Set xRs2 = BuscaConCriterio("SELECT idcuen from mae_documentocta  WHERE mae_documentocta.iddoc  = " & NulosN(TxtTipDoc2.Text) & " and mae_documentocta.idmon =" & 1 & " and tipope = -1", xCon)
        If xRs2.RecordCount > 0 Then
             xCuentaDoc = NulosN(xRs2("idcuen"))
        Else
             MsgBox "No se ha encontrado cuenta contable para el documento especificado", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
             
        End If

        
        Set xRs2 = Nothing
        
    End If
    
    
    ' buscamos para hallar el numero de serie asignado al almacen
    If TxtTipDoc2.Text <> "" Then
        Dim Rst As New ADODB.Recordset
        Set Rst = BuscaConCriterio("SELECT * FROM alm_numseries WHERE idtipdoc = " & NulosN(TxtTipDoc2.Text) & " AND idalm = " & NulosN(TxtIdAlm2.Text) & "", xCon)
        If Rst.RecordCount <> 0 Then
            TxtNumSer2.Text = NulosC(Rst("numser"))
            TxtNumSer2_Validate True
        End If
        Set Rst = Nothing
    Else
        TxtNumSer2.Text = ""
        TxtNumDoc2.Text = ""
    End If
    
    Set xRs = Nothing
End Sub


Private Sub CmdBusNumSer2_Click()
    ' BUSCA EL NUMERO DE SERIE PARA EL TIPO DE DOCUENTOA ACTUAL
    
    ' VERIFICAMOS QUE LOS DATOS NECESARIOS PARA EJECUTAR EL PROCESO ESTEN INGRESADOS
    If TxtTipDoc2.Text = "" Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtTipDoc2.SetFocus
        Exit Sub
    End If

    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos(3, 4) As String
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    xCampos(0, 0) = "Codigo":         xCampos(0, 1) = "iddoc":       xCampos(0, 2) = "1500":    xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":    xCampos(1, 1) = "descripcion": xCampos(1, 2) = "2500":    xCampos(1, 3) = "C"
    xCampos(2, 0) = "Serie":          xCampos(2, 1) = "numser":      xCampos(2, 2) = "1500":    xCampos(2, 3) = "C"
    
    xform.SQLCad = "SELECT mae_documento.descripcion, mae_series.iddoc, Format([mae_series].[numser],'0000') AS numser, " _
        & " mae_series.numdoc FROM mae_documento INNER JOIN mae_series ON mae_documento.id = mae_series.iddoc " _
        & " WHERE (((mae_series.iddoc)=1))"
    
    xform.titulo = "Buscando Series"
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "numser"
    xform.CampoBusca = "numser"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos)
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            TxtNumSer2.Text = Format(NulosN(xRs("numser")), "0000")
            TxtNumDoc2.Text = HallaNumdocVenta(NulosN(TxtTipDoc2.Text), TxtNumSer.Text, xCon)
        End If
        TxtNumDocGen.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub



Private Sub TxtNumSer2_KeyPress(KeyAscii As Integer)
    If NulosN(TxtTipDoc2.Text) = 0 Then
        MsgBox "No ha especificado el tipo de documento", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtNumSer2.Text = ""
        TxtTipDoc2.SetFocus
        Exit Sub
    End If
        
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumSer2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusNumSer2_Click
    End If
End Sub



Private Sub TxtNumSer2_Validate(Cancel As Boolean)
    ' VALIDA EL NUMERO DE SERIE
    If QueHace = 3 Then Exit Sub
    Dim Rstdoc As New ADODB.Recordset
    If NulosC(TxtNumSer2.Text) = "" Then
        Exit Sub
    Else
        If QueHace <> 1 Then Exit Sub
        
        TxtNumSer2.Text = Format(TxtNumSer2.Text, "0000")
        Dim Rst As New ADODB.Recordset
        
        RST_Busq Rst, "SELECT top 1 vta_ventas.numdoc AS numero from vta_ventas WHERE (((vta_ventas.numser)='" & NulosC(TxtNumSer2.Text) & "') AND ((vta_ventas.tipdoc)=" & NulosN(TxtTipDoc2.Text) & ")) and IsNumeric(vta_ventas.numdoc)=True " _
            & " ORDER BY vta_ventas.numdoc DESC ", xCon

        If Rst.RecordCount = 0 Then
            TxtNumDocGen.Text = "0000000001"
        Else
            Rst.MoveFirst
            TxtNumDocGen.Text = Format(NulosN(Rst("numero")) + 1, "0000000000")
        End If
        Set Rst = Nothing
    End If
End Sub



Private Sub TxtNumDoc2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtNumDoc2_Validate(Cancel As Boolean)
    ' VALIDA EL NUMERO DE DOCUMENTO
    If QueHace = 3 Then Exit Sub
    If NulosC(TxtNumDoc2.Text) <> "" Then
    
        TxtNumDoc2.Text = Format(TxtNumDoc2.Text, "0000000000")
        
        Dim Rst As New ADODB.Recordset
        Dim nSQL As String
        '--ver si existe el numero de doc
        
        RST_Busq Rst, "SELECT vta_ventas.numser, vta_ventas.numdoc, vta_ventas.fchdoc, mae_cliente.nombre, Left([vta_ventas].[numreg],2) & IIf([mae_libros].[codsun] Is Null,'',[mae_libros].[codsun]) & Right([vta_ventas].[numreg],4) AS registro " _
            & " FROM (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) LEFT JOIN mae_libros ON vta_ventas.idlib = mae_libros.id " _
            & " WHERE (((vta_ventas.numser)='" & Trim(TxtNumSer2.Text) & "') AND ((vta_ventas.numdoc)='" & TxtNumDoc2.Text & "')) and vta_ventas.tipdoc = " & NulosN(TxtTipDoc2.Text) & nSQL, xCon
                
                
        If Rst.RecordCount <> 0 Then
            '--poner el nuevo numero doc
            TxtNumSer2_Validate True
            MsgBox "El número de documento de venta ya existe " & vbCr & "Nº Registro: " & NulosC(Rst("registro")) & vbCr & "Fecha Doc.   " & NulosC(Rst("fchdoc")) & vbCr & "Cliente:         " & NulosC(Rst("nombre")) & vbCr & "Será reemplazado por " + Trim(TxtNumSer2.Text) + "-" + Trim(TxtNumDoc2.Text), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        End If
        Set Rst = Nothing
        
    End If
End Sub


'*---------------------------------------------------------------------------------------
'*--------------
'*---------------------------------------------------------------------------------------






