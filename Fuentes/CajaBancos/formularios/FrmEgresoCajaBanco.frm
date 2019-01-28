VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form FrmEgresoCajaBanco 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Caja y Bancos - Egresos"
   ClientHeight    =   8265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   14880
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   8265
   ScaleWidth      =   14880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame9 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   5475
      Left            =   1140
      TabIndex        =   54
      Top             =   7110
      Visible         =   0   'False
      Width           =   11355
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   0
         Left            =   11100
         Picture         =   "FrmEgresoCajaBanco.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   104
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VB.TextBox TxtSaldoCambio 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   10305
         Locked          =   -1  'True
         TabIndex        =   68
         Text            =   "TxtSaldoCambio"
         Top             =   2100
         Width           =   975
      End
      Begin VB.TextBox TxtHabDol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   10035
         Locked          =   -1  'True
         TabIndex        =   58
         Text            =   "TxtHabDol"
         Top             =   4905
         Width           =   975
      End
      Begin VB.TextBox TxtDebDol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9075
         Locked          =   -1  'True
         TabIndex        =   57
         Text            =   "TxtDebDol"
         Top             =   4905
         Width           =   975
      End
      Begin VB.TextBox TxtHabSol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8115
         Locked          =   -1  'True
         TabIndex        =   56
         Text            =   "TxtHabSol"
         Top             =   4905
         Width           =   975
      End
      Begin VB.TextBox TxtDebSol 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7155
         Locked          =   -1  'True
         TabIndex        =   55
         Text            =   "TxtDebSol"
         Top             =   4905
         Width           =   975
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg4 
         Height          =   2430
         Left            =   90
         TabIndex        =   59
         Top             =   2460
         Width           =   11190
         _cx             =   19738
         _cy             =   4286
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   14
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmEgresoCajaBanco.frx":02EC
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
      Begin VB.Frame Frame11 
         Height          =   570
         Left            =   90
         TabIndex        =   60
         Top             =   4830
         Width           =   5400
         Begin VB.CommandButton CmdExportar1 
            Caption         =   "E&xportar"
            Height          =   345
            Left            =   3330
            TabIndex        =   109
            Top             =   165
            Width           =   1545
         End
         Begin VB.CommandButton Command5 
            Caption         =   "&Aceptar"
            Height          =   345
            Left            =   660
            TabIndex        =   61
            Top             =   165
            Width           =   1695
         End
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg5 
         Height          =   1470
         Left            =   90
         TabIndex        =   65
         Top             =   615
         Width           =   11190
         _cx             =   19738
         _cy             =   2593
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
         Rows            =   50
         Cols            =   13
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmEgresoCajaBanco.frx":048B
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
      Begin VB.Label Label9 
         Caption         =   "Total ==>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   9135
         TabIndex        =   69
         Top             =   2130
         Width           =   990
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Asiento por la Cancelacion"
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
         TabIndex        =   67
         Top             =   2235
         Width           =   2295
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Documentos Provicionados en Dolares"
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
         Left            =   120
         TabIndex        =   66
         Top             =   390
         Width           =   3300
      End
      Begin VB.Label Label6 
         Caption         =   "Total ==>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   6000
         TabIndex        =   63
         Top             =   4935
         Width           =   990
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Asiento Contable de la Operación"
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
         TabIndex        =   62
         Top             =   75
         Width           =   2865
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   3
         X1              =   11340
         X2              =   11325
         Y1              =   15
         Y2              =   5445
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   3
         X1              =   15
         X2              =   11580
         Y1              =   5460
         Y2              =   5460
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000E&
         BorderWidth     =   2
         Index           =   2
         X1              =   15
         X2              =   11595
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000E&
         BorderWidth     =   2
         Index           =   2
         X1              =   15
         X2              =   15
         Y1              =   -30
         Y2              =   5430
      End
      Begin VB.Shape Shape2 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   45
         Top             =   45
         Width           =   11265
      End
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3810
      Left            =   12840
      TabIndex        =   34
      Top             =   1380
      Visible         =   0   'False
      Width           =   11355
      Begin VB.CheckBox ChkVerSaldo 
         Caption         =   "Documentos con &Saldo"
         Height          =   255
         Left            =   7710
         TabIndex        =   110
         Top             =   420
         Value           =   1  'Checked
         Width           =   2265
      End
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   1
         Left            =   11100
         Picture         =   "FrmEgresoCajaBanco.frx":0616
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   103
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VB.TextBox TxtTotal1 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "TxtTotal1"
         Top             =   2940
         Width           =   975
      End
      Begin VB.TextBox TxtTotal2 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7110
         Locked          =   -1  'True
         TabIndex        =   44
         Text            =   "TxtTotal2"
         Top             =   2940
         Width           =   975
      End
      Begin VB.TextBox TxtTotal3 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8070
         Locked          =   -1  'True
         TabIndex        =   43
         Text            =   "TxtTotal3"
         Top             =   2940
         Width           =   975
      End
      Begin VB.TextBox TxtTotal4 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9030
         Locked          =   -1  'True
         TabIndex        =   42
         Text            =   "TxtTotal4"
         Top             =   2940
         Width           =   975
      End
      Begin VB.TextBox TxtTotal5 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9990
         Locked          =   -1  'True
         TabIndex        =   41
         Text            =   "TxtTotal5"
         Top             =   2940
         Width           =   1035
      End
      Begin VB.CommandButton CmdBusCliente 
         Enabled         =   0   'False
         Height          =   240
         Left            =   5955
         Picture         =   "FrmEgresoCajaBanco.frx":0902
         Style           =   1  'Graphical
         TabIndex        =   35
         Top             =   435
         Width           =   240
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg3 
         Height          =   2190
         Left            =   90
         TabIndex        =   38
         Top             =   735
         Width           =   11190
         _cx             =   19738
         _cy             =   3863
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
         Rows            =   50
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmEgresoCajaBanco.frx":0A34
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
      Begin VB.Frame Frame7 
         Height          =   600
         Left            =   90
         TabIndex        =   46
         Top             =   3180
         Width           =   6090
         Begin VB.CommandButton CmdAgregar 
            Caption         =   "&Agregar Documento"
            Height          =   345
            Left            =   60
            TabIndex        =   36
            Top             =   180
            Width           =   1815
         End
         Begin VB.CommandButton CmdEliminar 
            Caption         =   "&Eliminar Documento"
            Height          =   345
            Left            =   2010
            TabIndex        =   37
            Top             =   180
            Width           =   1815
         End
         Begin VB.CommandButton CmdExportar 
            Caption         =   "E&xportar"
            Height          =   345
            Left            =   4410
            TabIndex        =   108
            Top             =   180
            Width           =   1545
         End
      End
      Begin VB.TextBox TxtProv 
         Height          =   300
         Left            =   1215
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   40
         Text            =   "TxtProv"
         Top             =   405
         Width           =   5010
      End
      Begin VB.Frame Frame4 
         Height          =   600
         Left            =   6945
         TabIndex        =   51
         Top             =   3180
         Width           =   4350
         Begin VB.CommandButton Command11 
            Caption         =   "Cancelar"
            Height          =   345
            Left            =   2355
            TabIndex        =   98
            Top             =   180
            Width           =   1695
         End
         Begin VB.CommandButton Command1 
            Caption         =   "&Aceptar"
            Height          =   345
            Left            =   300
            TabIndex        =   39
            Top             =   180
            Width           =   1695
         End
      End
      Begin VB.Label LblTc 
         Alignment       =   1  'Right Justify
         Caption         =   "LblTc"
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
         Height          =   225
         Left            =   9675
         TabIndex        =   106
         Top             =   450
         Width           =   1605
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000E&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   -30
         Y2              =   4035
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000E&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   11595
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   11580
         Y1              =   3795
         Y2              =   3795
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   0
         X1              =   11340
         X2              =   11325
         Y1              =   15
         Y2              =   4050
      End
      Begin VB.Label LblTitulo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LblTitulo"
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
         TabIndex        =   50
         Top             =   90
         Width           =   750
      End
      Begin VB.Label Label2 
         Caption         =   "Total ==>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5130
         TabIndex        =   49
         Top             =   2970
         Width           =   990
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   48
         Top             =   435
         Width           =   735
      End
      Begin VB.Label LblIdCliente 
         AutoSize        =   -1  'True
         Caption         =   "LblIdCliente"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6360
         TabIndex        =   47
         Top             =   450
         Visible         =   0   'False
         Width           =   825
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   45
         Top             =   45
         Width           =   11265
      End
   End
   Begin VB.Frame Frame12 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   3810
      Left            =   13440
      TabIndex        =   77
      Top             =   4380
      Visible         =   0   'False
      Width           =   11355
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Index           =   2
         Left            =   11100
         Picture         =   "FrmEgresoCajaBanco.frx":0C2E
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   105
         ToolTipText     =   "Cerrar"
         Top             =   90
         Width           =   195
      End
      Begin VB.Frame Frame13 
         Height          =   600
         Left            =   90
         TabIndex        =   85
         Top             =   3180
         Width           =   4080
         Begin VB.CommandButton Command8 
            Caption         =   "&Eliminar Documento"
            Height          =   345
            Left            =   2025
            TabIndex        =   87
            Top             =   180
            Width           =   1695
         End
         Begin VB.CommandButton Command7 
            Caption         =   "&Agregar Documento"
            Height          =   345
            Left            =   315
            TabIndex        =   86
            Top             =   180
            Width           =   1695
         End
      End
      Begin VB.CommandButton Command6 
         Height          =   240
         Left            =   5955
         Picture         =   "FrmEgresoCajaBanco.frx":0F1A
         Style           =   1  'Graphical
         TabIndex        =   83
         Top             =   435
         Width           =   240
      End
      Begin VB.TextBox TxtTotal5A 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9990
         Locked          =   -1  'True
         TabIndex        =   82
         Text            =   "TxtTotal5A"
         Top             =   2940
         Width           =   1035
      End
      Begin VB.TextBox TxtTotal4A 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   9030
         Locked          =   -1  'True
         TabIndex        =   81
         Text            =   "TxtTotal4A"
         Top             =   2940
         Width           =   975
      End
      Begin VB.TextBox TxtTotal3A 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   8070
         Locked          =   -1  'True
         TabIndex        =   80
         Text            =   "TxtTotal3A"
         Top             =   2940
         Width           =   975
      End
      Begin VB.TextBox TxtTotal2A 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   7110
         Locked          =   -1  'True
         TabIndex        =   79
         Text            =   "TxtTotal2A"
         Top             =   2940
         Width           =   975
      End
      Begin VB.TextBox TxtTotal1A 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         ForeColor       =   &H00C00000&
         Height          =   285
         Left            =   6150
         Locked          =   -1  'True
         TabIndex        =   78
         Text            =   "TxtTotal1A"
         Top             =   2940
         Width           =   975
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg6 
         Height          =   2190
         Left            =   120
         TabIndex        =   84
         Top             =   735
         Width           =   11190
         _cx             =   19738
         _cy             =   3863
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
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   50
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmEgresoCajaBanco.frx":104C
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
      Begin VB.TextBox TxtProvA 
         Height          =   300
         Left            =   1215
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   88
         Text            =   "TxtProvA"
         Top             =   405
         Width           =   5010
      End
      Begin VB.Frame Frame14 
         Height          =   600
         Left            =   6135
         TabIndex        =   89
         Top             =   3180
         Width           =   5130
         Begin VB.CommandButton Command10 
            Caption         =   "Cancelar"
            Height          =   345
            Left            =   2600
            TabIndex        =   97
            Top             =   180
            Width           =   1695
         End
         Begin VB.CommandButton Command9 
            Caption         =   "&Aceptar"
            Height          =   345
            Left            =   540
            TabIndex        =   90
            Top             =   180
            Width           =   1695
         End
      End
      Begin VB.Label LblIdClienteA 
         AutoSize        =   -1  'True
         Caption         =   "LblIdClienteA"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   6360
         TabIndex        =   94
         Top             =   450
         Visible         =   0   'False
         Width           =   930
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Proveedor"
         Height          =   195
         Index           =   9
         Left            =   120
         TabIndex        =   93
         Top             =   435
         Width           =   735
      End
      Begin VB.Label Label11 
         Caption         =   "Total ==>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   180
         Left            =   5130
         TabIndex        =   92
         Top             =   2970
         Width           =   990
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "LblTitulo"
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
         TabIndex        =   91
         Top             =   75
         Width           =   750
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   5
         X1              =   11340
         X2              =   11325
         Y1              =   15
         Y2              =   4050
      End
      Begin VB.Line Line3 
         BorderColor     =   &H00404040&
         BorderWidth     =   2
         Index           =   5
         X1              =   15
         X2              =   11580
         Y1              =   3795
         Y2              =   3795
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000E&
         BorderWidth     =   2
         Index           =   4
         X1              =   15
         X2              =   11595
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000E&
         BorderWidth     =   2
         Index           =   4
         X1              =   15
         X2              =   15
         Y1              =   -30
         Y2              =   4035
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         Height          =   285
         Left            =   45
         Top             =   45
         Width           =   11265
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
            Picture         =   "FrmEgresoCajaBanco.frx":1217
            Key             =   "IMG1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgresoCajaBanco.frx":175B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgresoCajaBanco.frx":1AED
            Key             =   "IMG2"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgresoCajaBanco.frx":1C71
            Key             =   "IMG3"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgresoCajaBanco.frx":20C5
            Key             =   "IMG4"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgresoCajaBanco.frx":21DD
            Key             =   "IMG5"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgresoCajaBanco.frx":2721
            Key             =   "IMG6"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgresoCajaBanco.frx":2C65
            Key             =   "IMG7"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgresoCajaBanco.frx":2D79
            Key             =   "IMG8"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgresoCajaBanco.frx":2E8D
            Key             =   "IMG9"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgresoCajaBanco.frx":32E1
            Key             =   "IMG10"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmEgresoCajaBanco.frx":344D
            Key             =   "IMG11"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   30
      Top             =   0
      Width           =   14880
      _ExtentX        =   26247
      _ExtentY        =   609
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
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
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
            Object.ToolTipText     =   "Imprimir "
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
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7230
      Left            =   15
      TabIndex        =   7
      Top             =   375
      Width           =   11880
      _cx             =   20955
      _cy             =   12753
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
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6810
         Left            =   45
         TabIndex        =   25
         Top             =   375
         Width           =   11790
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   6420
            Left            =   30
            TabIndex        =   26
            Top             =   360
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
            Columns(1).Caption=   "Nº Reg."
            Columns(1).DataField=   "registro"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Fch. Mov."
            Columns(2).DataField=   "fchope"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Importe"
            Columns(3).DataField=   "importe"
            Columns(3).NumberFormat=   "0.00"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "M"
            Columns(4).DataField=   "simbolo"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Origen"
            Columns(5).DataField=   "descori"
            Columns(5).NumberFormat=   "Short Date"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "T.D."
            Columns(6).DataField=   "desdocabre"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Nº Documento"
            Columns(7).DataField=   "numdoc"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Glosa"
            Columns(8).DataField=   "glosa"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Nº Cuenta"
            Columns(9).DataField=   "numcue"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Banco"
            Columns(10).DataField=   "descban"
            Columns(10).NumberFormat=   "0.00"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   11
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=11"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1455"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1376"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=512"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1535"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1455"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=513"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1826"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1746"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=514"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=635"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=556"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=512"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=4207"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=4128"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=1005"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=926"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(43)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(44)=   "Column(7).Width=2355"
            Splits(0)._ColumnProps(45)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(46)=   "Column(7)._WidthInPix=2275"
            Splits(0)._ColumnProps(47)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(48)=   "Column(7)._ColStyle=512"
            Splits(0)._ColumnProps(49)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(50)=   "Column(8).Width=6773"
            Splits(0)._ColumnProps(51)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(52)=   "Column(8)._WidthInPix=6694"
            Splits(0)._ColumnProps(53)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(54)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(55)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(56)=   "Column(9).Width=953"
            Splits(0)._ColumnProps(57)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(58)=   "Column(9)._WidthInPix=873"
            Splits(0)._ColumnProps(59)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(60)=   "Column(9)._ColStyle=516"
            Splits(0)._ColumnProps(61)=   "Column(9).Visible=0"
            Splits(0)._ColumnProps(62)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(63)=   "Column(10).Width=318"
            Splits(0)._ColumnProps(64)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(65)=   "Column(10)._WidthInPix=238"
            Splits(0)._ColumnProps(66)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(67)=   "Column(10)._ColStyle=514"
            Splits(0)._ColumnProps(68)=   "Column(10).Visible=0"
            Splits(0)._ColumnProps(69)=   "Column(10).Order=11"
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
            HeadLines       =   1.5
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
            _StyleDefs(11)  =   ":id=2,.bold=0,.fontsize=825,.italic=0,.underline=0,.strikethrough=0,.charset=0"
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
            _StyleDefs(36)  =   "Splits(0).Columns(0).Style:id=78,.parent=13"
            _StyleDefs(37)  =   "Splits(0).Columns(0).HeadingStyle:id=75,.parent=14"
            _StyleDefs(38)  =   "Splits(0).Columns(0).FooterStyle:id=76,.parent=15"
            _StyleDefs(39)  =   "Splits(0).Columns(0).EditorStyle:id=77,.parent=17"
            _StyleDefs(40)  =   "Splits(0).Columns(1).Style:id=32,.parent=13,.alignment=0"
            _StyleDefs(41)  =   "Splits(0).Columns(1).HeadingStyle:id=29,.parent=14"
            _StyleDefs(42)  =   "Splits(0).Columns(1).FooterStyle:id=30,.parent=15"
            _StyleDefs(43)  =   "Splits(0).Columns(1).EditorStyle:id=31,.parent=17"
            _StyleDefs(44)  =   "Splits(0).Columns(2).Style:id=70,.parent=13,.alignment=2"
            _StyleDefs(45)  =   "Splits(0).Columns(2).HeadingStyle:id=67,.parent=14"
            _StyleDefs(46)  =   "Splits(0).Columns(2).FooterStyle:id=68,.parent=15"
            _StyleDefs(47)  =   "Splits(0).Columns(2).EditorStyle:id=69,.parent=17"
            _StyleDefs(48)  =   "Splits(0).Columns(3).Style:id=28,.parent=13,.alignment=1"
            _StyleDefs(49)  =   "Splits(0).Columns(3).HeadingStyle:id=25,.parent=14"
            _StyleDefs(50)  =   "Splits(0).Columns(3).FooterStyle:id=26,.parent=15"
            _StyleDefs(51)  =   "Splits(0).Columns(3).EditorStyle:id=27,.parent=17"
            _StyleDefs(52)  =   "Splits(0).Columns(4).Style:id=58,.parent=13,.alignment=0"
            _StyleDefs(53)  =   "Splits(0).Columns(4).HeadingStyle:id=55,.parent=14"
            _StyleDefs(54)  =   "Splits(0).Columns(4).FooterStyle:id=56,.parent=15"
            _StyleDefs(55)  =   "Splits(0).Columns(4).EditorStyle:id=57,.parent=17"
            _StyleDefs(56)  =   "Splits(0).Columns(5).Style:id=46,.parent=13,.alignment=0"
            _StyleDefs(57)  =   "Splits(0).Columns(5).HeadingStyle:id=43,.parent=14"
            _StyleDefs(58)  =   "Splits(0).Columns(5).FooterStyle:id=44,.parent=15"
            _StyleDefs(59)  =   "Splits(0).Columns(5).EditorStyle:id=45,.parent=17"
            _StyleDefs(60)  =   "Splits(0).Columns(6).Style:id=74,.parent=13"
            _StyleDefs(61)  =   "Splits(0).Columns(6).HeadingStyle:id=71,.parent=14"
            _StyleDefs(62)  =   "Splits(0).Columns(6).FooterStyle:id=72,.parent=15"
            _StyleDefs(63)  =   "Splits(0).Columns(6).EditorStyle:id=73,.parent=17"
            _StyleDefs(64)  =   "Splits(0).Columns(7).Style:id=50,.parent=13,.alignment=0"
            _StyleDefs(65)  =   "Splits(0).Columns(7).HeadingStyle:id=47,.parent=14"
            _StyleDefs(66)  =   "Splits(0).Columns(7).FooterStyle:id=48,.parent=15"
            _StyleDefs(67)  =   "Splits(0).Columns(7).EditorStyle:id=49,.parent=17"
            _StyleDefs(68)  =   "Splits(0).Columns(8).Style:id=82,.parent=13"
            _StyleDefs(69)  =   "Splits(0).Columns(8).HeadingStyle:id=79,.parent=14"
            _StyleDefs(70)  =   "Splits(0).Columns(8).FooterStyle:id=80,.parent=15"
            _StyleDefs(71)  =   "Splits(0).Columns(8).EditorStyle:id=81,.parent=17"
            _StyleDefs(72)  =   "Splits(0).Columns(9).Style:id=62,.parent=13"
            _StyleDefs(73)  =   "Splits(0).Columns(9).HeadingStyle:id=59,.parent=14"
            _StyleDefs(74)  =   "Splits(0).Columns(9).FooterStyle:id=60,.parent=15"
            _StyleDefs(75)  =   "Splits(0).Columns(9).EditorStyle:id=61,.parent=17"
            _StyleDefs(76)  =   "Splits(0).Columns(10).Style:id=66,.parent=13,.alignment=1"
            _StyleDefs(77)  =   "Splits(0).Columns(10).HeadingStyle:id=63,.parent=14"
            _StyleDefs(78)  =   "Splits(0).Columns(10).FooterStyle:id=64,.parent=15"
            _StyleDefs(79)  =   "Splits(0).Columns(10).EditorStyle:id=65,.parent=17"
            _StyleDefs(80)  =   "Named:id=33:Normal"
            _StyleDefs(81)  =   ":id=33,.parent=0"
            _StyleDefs(82)  =   "Named:id=34:Heading"
            _StyleDefs(83)  =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(84)  =   ":id=34,.wraptext=-1"
            _StyleDefs(85)  =   "Named:id=35:Footing"
            _StyleDefs(86)  =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(87)  =   "Named:id=36:Selected"
            _StyleDefs(88)  =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(89)  =   "Named:id=37:Caption"
            _StyleDefs(90)  =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(91)  =   "Named:id=38:HighlightRow"
            _StyleDefs(92)  =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(93)  =   "Named:id=39:EvenRow"
            _StyleDefs(94)  =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(95)  =   "Named:id=40:OddRow"
            _StyleDefs(96)  =   ":id=40,.parent=33"
            _StyleDefs(97)  =   "Named:id=41:RecordSelector"
            _StyleDefs(98)  =   ":id=41,.parent=34"
            _StyleDefs(99)  =   "Named:id=42:FilterBar"
            _StyleDefs(100) =   ":id=42,.parent=33"
         End
         Begin VB.Label LblPeriodo 
            Alignment       =   2  'Center
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
            Height          =   300
            Left            =   9810
            TabIndex        =   29
            Top             =   0
            Width           =   1860
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Consulta de Egresos"
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
            Left            =   105
            TabIndex        =   28
            Top             =   30
            Width           =   11610
         End
         Begin VB.Line Line1 
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label LblMes1 
            AutoSize        =   -1  'True
            Caption         =   "LblMes1"
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
            Left            =   8235
            TabIndex        =   27
            Top             =   30
            Width           =   885
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "LblidDocumento"
         Height          =   6810
         Left            =   12525
         TabIndex        =   8
         Top             =   375
         Width           =   11790
         Begin VB.TextBox TxtImpDifDol 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   7935
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   101
            Text            =   "TxtImpDifDol"
            Top             =   5910
            Width           =   1095
         End
         Begin VB.TextBox TxtImpDifSol 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   6855
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   100
            Text            =   "TxtImpDifSol"
            Top             =   5910
            Width           =   1095
         End
         Begin VB.OptionButton OptCanje 
            Caption         =   "Canje"
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
            Height          =   225
            Left            =   3585
            TabIndex        =   76
            Top             =   600
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.TextBox TxtImpDebDol 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   7935
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   75
            Text            =   "TxtImpDebDol"
            Top             =   3690
            Width           =   1095
         End
         Begin VB.TextBox TxtImpHabDol 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   7935
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   74
            Text            =   "TxtImpHabDol"
            Top             =   5460
            Width           =   1095
         End
         Begin VB.Frame Frame10 
            Enabled         =   0   'False
            Height          =   1335
            Left            =   9180
            TabIndex        =   71
            Top             =   4125
            Width           =   2490
            Begin VB.CheckBox ChkChequeAnulado 
               Caption         =   "Cheque Extraviado / Anulado"
               Height          =   405
               Left            =   300
               TabIndex        =   107
               Top             =   870
               Width           =   2205
            End
            Begin VB.CommandButton Command4 
               Caption         =   "Eliminar Destino"
               Height          =   285
               Left            =   315
               TabIndex        =   73
               Top             =   540
               Width           =   1860
            End
            Begin VB.CommandButton Command3 
               Caption         =   "&Agregar Destino"
               Height          =   285
               Left            =   315
               TabIndex        =   72
               Top             =   240
               Width           =   1860
            End
         End
         Begin VB.TextBox TxtGlosa 
            Height          =   945
            Left            =   1470
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Text            =   "FrmEgresoCajaBanco.frx":3995
            Top             =   1185
            Width           =   7605
         End
         Begin VB.CommandButton Command2 
            Caption         =   "Ver Asiento Contable"
            Height          =   615
            Left            =   9735
            TabIndex        =   64
            Top             =   5865
            Width           =   1305
         End
         Begin VB.Frame Frame8 
            Caption         =   "[ Tipo de Cambio ]"
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
            Height          =   900
            Left            =   9180
            TabIndex        =   52
            Top             =   1425
            Width           =   2490
            Begin VB.Label lblTipCambio 
               Alignment       =   2  'Center
               Caption         =   "lblTipCambio"
               BeginProperty Font 
                  Name            =   "Courier New"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00C00000&
               Height          =   270
               Left            =   135
               TabIndex        =   53
               Top             =   345
               Width           =   2220
            End
         End
         Begin VB.TextBox TxtImpHabSol 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   6855
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   31
            Text            =   "TxtImpHabSol"
            Top             =   5460
            Width           =   1095
         End
         Begin VB.Frame Frame5 
            Caption         =   "[ Periodo ]"
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
            Height          =   630
            Left            =   9180
            TabIndex        =   16
            Top             =   765
            Width           =   2490
            Begin VB.Label LblMes 
               Alignment       =   2  'Center
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
               Height          =   240
               Left            =   315
               TabIndex        =   17
               Top             =   240
               Width           =   1860
            End
         End
         Begin VB.CommandButton CmdBusMon 
            Height          =   240
            Left            =   6270
            Picture         =   "FrmEgresoCajaBanco.frx":399E
            Style           =   1  'Graphical
            TabIndex        =   15
            Top             =   885
            Width           =   240
         End
         Begin VB.TextBox TxtImpDebSol 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   6855
            Locked          =   -1  'True
            MaxLength       =   20
            TabIndex        =   14
            Text            =   "TxtImpDebSol"
            Top             =   3690
            Width           =   1095
         End
         Begin VB.OptionButton OptCaja 
            Caption         =   "Caja"
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
            Height          =   225
            Left            =   1470
            TabIndex        =   0
            Top             =   600
            Visible         =   0   'False
            Width           =   765
         End
         Begin VB.OptionButton OptBanco 
            Caption         =   "Banco"
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
            Height          =   225
            Left            =   2520
            TabIndex        =   1
            Top             =   600
            Visible         =   0   'False
            Width           =   945
         End
         Begin VB.Frame Frame6 
            Enabled         =   0   'False
            Height          =   1335
            Left            =   9180
            TabIndex        =   9
            Top             =   2355
            Width           =   2490
            Begin VB.CommandButton CmdAddCon 
               Caption         =   "&Agregar Destino"
               Height          =   285
               Left            =   315
               TabIndex        =   13
               Top             =   285
               Width           =   1860
            End
            Begin VB.CommandButton CmdDelCon 
               Caption         =   "Eliminar Destino"
               Height          =   285
               Left            =   315
               TabIndex        =   12
               Top             =   585
               Width           =   1860
            End
            Begin VB.OptionButton OptDe1 
               Caption         =   "x Descipción"
               Height          =   195
               Left            =   135
               TabIndex        =   11
               Top             =   945
               Width           =   1230
            End
            Begin VB.OptionButton OptDe2 
               Caption         =   "x Cuenta"
               Height          =   195
               Left            =   1455
               TabIndex        =   10
               Top             =   945
               Value           =   -1  'True
               Width           =   945
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg2 
            Height          =   1230
            Left            =   105
            TabIndex        =   5
            Top             =   2445
            Width           =   9030
            _cx             =   15928
            _cy             =   2170
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
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEgresoCajaBanco.frx":3AD0
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchMov 
            Height          =   300
            Left            =   1470
            TabIndex        =   2
            Top             =   855
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
            Valor           =   "07/12/2007"
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   1230
            Left            =   105
            TabIndex        =   6
            Top             =   4215
            Width           =   9030
            _cx             =   15928
            _cy             =   2170
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
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmEgresoCajaBanco.frx":3BF8
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
         Begin VB.TextBox TxtIdMon 
            Height          =   300
            Left            =   5565
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   3
            Text            =   "TxtIdMon"
            Top             =   855
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Diferencia"
            Height          =   195
            Index           =   0
            Left            =   5730
            TabIndex        =   102
            Top             =   6000
            Width           =   720
         End
         Begin VB.Label lblReg 
            Caption         =   "lblReg"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   270
            Left            =   9240
            TabIndex        =   99
            Top             =   315
            Width           =   2415
         End
         Begin VB.Label Label13 
            AutoSize        =   -1  'True
            Caption         =   "( Origen = Haber )"
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
            Left            =   1635
            TabIndex        =   96
            Top             =   3990
            Width           =   1545
         End
         Begin VB.Label Label12 
            AutoSize        =   -1  'True
            Caption         =   "( Destino = Debe )"
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
            Left            =   1635
            TabIndex        =   95
            Top             =   2220
            Width           =   1575
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Index           =   7
            Left            =   135
            TabIndex        =   70
            Top             =   1215
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Origen del Egreso"
            Height          =   195
            Index           =   6
            Left            =   105
            TabIndex        =   33
            Top             =   3990
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Importe Haber"
            Height          =   195
            Index           =   1
            Left            =   5730
            TabIndex        =   32
            Top             =   5505
            Width           =   1005
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Operación"
            Height          =   195
            Index           =   10
            Left            =   120
            TabIndex        =   24
            Top             =   615
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fecha de Emisión"
            Height          =   195
            Index           =   8
            Left            =   120
            TabIndex        =   23
            Top             =   915
            Width           =   1260
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
            Left            =   6540
            TabIndex        =   22
            Top             =   855
            Width           =   2535
         End
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle de la Operación"
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
            Top             =   30
            Width           =   11610
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Importe Debe"
            Height          =   195
            Index           =   4
            Left            =   5730
            TabIndex        =   20
            Top             =   3720
            Width           =   960
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   4800
            TabIndex        =   19
            Top             =   915
            Width           =   690
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Destino del Egreso"
            Height          =   195
            Index           =   3
            Left            =   105
            TabIndex        =   18
            Top             =   2220
            Width           =   1335
         End
      End
   End
End
Attribute VB_Name = "FrmEgresoCajaBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim RstMov As New ADODB.Recordset
Dim RstTMPDoc As New ADODB.Recordset
Dim RstTmpDocOri As New ADODB.Recordset
Dim Agregando As Boolean
Dim xFchPer As String
Dim CaracteresNumericos As String
Dim xHorIni As Date

Dim mIdRegistro& '--identificador del registro
Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim mMesActivo As Integer '--indica el mes activo

Dim mCorrelativo1 As Long '--diferencia de cada item origen
Dim mCorrelativo2 As Long '--diferencia de cada item destino
Dim fCierrePeriodo As Boolean '--indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)

Dim mVentanaActiva As Integer '--INDICA QUE PESTAÑA ESTA ACTIVA 0::ORIGEN; 1::DESTINO

'--ESTA EN FUNCION DEL REGISTRO SELECCIONADO (ORIGEN O DESTINO)
Dim IdConcepto As Integer  '--CODIGO DEL CONCEPTO ORIGEN O DESTINO
Dim IdModulo As Integer '--CODIGO DEL MODULO
Dim TipCam As Double    '--INDICA EL TIPO DE CAMBIO
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO

Dim OrigFX As Long '--para mover el frame posicion horizontal
Dim OrigFY As Long '--para mover el frame posicion vertical

Sub Cancelar()
    QueHace = 3
    ActivaTool
    Label5.Caption = "Detalle de la operación"
    Bloquea
    TabOne1.TabEnabled(0) = True
    TabOne1.CurrTab = 0
End Sub

Private Sub ChkChequeAnulado_Click()
    If ChkChequeAnulado.Value = 1 Then
        Fg2.Rows = 1
        TxtImpDebSol.Text = "0.00"
        TxtImpDebDol.Text = "0.00"
    End If
End Sub

Private Sub CmdAddCon_Click()
    If Fg2.Rows = 1 Then
        Fg2.Rows = Fg2.Rows + 1
        Fg2.Row = Fg2.Rows - 1
        Fg2_CellButtonClick Fg2.Rows - 1, 1
    Else
        If NulosN(Fg2.TextMatrix(Fg2.Rows - 1, 3)) = 0 Then
            MsgBox "No ha especificado un concepto para la ultima fila del destino de egresos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Fg2.Rows = Fg2.Rows + 1
        Fg2.Row = Fg2.Rows - 1
        Fg2_CellButtonClick Fg2.Rows - 1, 1
    End If
    
    Fg2.Select Fg2.Rows - 1, 1
    Fg2.SetFocus
End Sub

Private Sub CmdAgregar_Click()
    If QueHace = 3 Then Exit Sub
    
    If Fg2.Row < 1 Then Exit Sub
    If Fg2.Rows = 1 Then Exit Sub
    
    If TxtProv.Enabled = False Then Exit Sub
    
    Select Case IdModulo
        Case 3, 5
            '--3::Entregas a rendir
            '--5::Anticipos
            RstTMPDoc.MoveLast
            RstTMPDoc.AddNew
            RstTMPDoc("idconc") = IdConcepto
            RstTMPDoc("iddocu") = 0
            RstTMPDoc("corr") = mCorrelativo2
            RstTMPDoc("xtipo") = mVentanaActiva
            
            mCorrelativo2 = mCorrelativo2 + 1
            Fg3.Rows = Fg3.Rows + 1
            
            Agregando = True
            Fg3.TextMatrix(Fg3.Rows - 1, 11) = NulosN(RstTMPDoc("idconc"))
            Fg3.TextMatrix(Fg3.Rows - 1, 12) = NulosN(RstTMPDoc("iddocu"))
            Fg3.TextMatrix(Fg3.Rows - 1, 15) = NulosN(RstTMPDoc("corr"))
            Agregando = False
        
        Case 6 '--Bancos
            If RstTMPDoc.BOF = False And RstTMPDoc.EOF = False Then RstTMPDoc.MoveLast
            RstTMPDoc.AddNew
            RstTMPDoc("idconc") = IdConcepto
            RstTMPDoc("iddocu") = 0
            RstTMPDoc("corr") = mCorrelativo2
            RstTMPDoc("xtipo") = mVentanaActiva
            
            mCorrelativo2 = mCorrelativo2 + 1
            Fg3.Rows = Fg3.Rows + 1
            
            Agregando = True
            Fg3.TextMatrix(Fg3.Rows - 1, 11) = RstTMPDoc("idconc")
            Fg3.TextMatrix(Fg3.Rows - 1, 12) = RstTMPDoc("iddocu")
            Fg3.TextMatrix(Fg3.Rows - 1, 15) = RstTMPDoc("corr")
            Agregando = False
            
        Case 1, 8, 9, 10, 2, 4, 11, 19, 20
            '--1=Compras; 8=Boleta de Pago; 9=Honorarios; 10=Reembolsables
            '--2=Ventas; 4=Letras; 11=Liquidacion Gasto Debito; 19=Planilla de Letras
            '--20 Orden Desopacho
            
            CargarDocumentos NulosN(LblIdCliente.Caption)

    End Select
    
    Fg3.SetFocus
    
End Sub

Private Sub CmdBusCliente_Click()
    If QueHace = 3 Then Exit Sub

    Dim xCadWhere1, xCadWhere2 As String
    Dim nSQL As String
    
''    xCadWhere1 = CadWhere(IdConcepto, 1, 2)
''    xCadWhere2 = CadWhere(IdConcepto, 2, 2)
'
'    '--no tomar en cuenta cuando sea letras y orden de despacho
''    If NulosC(xCadWhere1) = "" And (IdModulo <> 2 And IdModulo <> 4 And IdModulo <> 19 And IdModulo <> 20) Then
''        MsgBox "El destino seleccionado no tiene documentos de compra asignado para su cancelacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
''        Exit Sub
''    End If
    
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Proveedor":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":     xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"
    
    
    'buscamos los proveedores que tengan el documento especificado
    Select Case IdModulo
    
        Case 1 '--Compras
            nSQL = "SELECT mae_prov.id, mae_prov.numruc, mae_prov.nombre " _
                & " FROM mae_prov WHERE (((mae_prov.id)<>0));"
            
            xForm.Titulo = "Buscando Proveedores"
    
        Case 2, 11 '--Ventas, Liquidación Gasto Debito
            nSQL = "SELECT mae_cliente.id, mae_cliente.numruc, mae_cliente.nombre " _
                + vbCr + " FROM mae_cliente " _
                + vbCr + " Where (((mae_cliente.id) <> 0)) "
            
            xForm.Titulo = "Buscando Clientes"
            xCampos(0, 0) = "Cliente":
        
        Case 4 '--Letras
            nSQL = "SELECT mae_cliente.id,mae_cliente.numruc, mae_cliente.nombre " _
                + vbCr + " FROM (let_letra LEFT JOIN mae_cliente ON let_letra.idclipro = mae_cliente.id) RIGHT JOIN let_letradet ON let_letra.id = let_letradet.idlet " _
                + vbCr + " GROUP BY mae_cliente.id,let_letra.idclipro, mae_cliente.numruc, mae_cliente.nombre; "
            
            xForm.Titulo = "Buscando Clientes"
            xCampos(0, 0) = "Cliente":
            
        Case 9  ' Honorarios
            nSQL = " SELECT DISTINCT mae_prov.id, mae_prov.numruc, mae_prov.nombre " _
                & " FROM mae_prov RIGHT JOIN com_honorarios ON mae_prov.id = com_honorarios.idpro  " _
                & " WHERE (com_honorarios.impsal<>0) "
            
            xForm.Titulo = "Buscando Prestadores de Servicio"
            xCampos(0, 0) = "Prestador de Servicio":
            
        Case 10 ' Reembolsables
            nSQL = " SELECT DISTINCT mae_prov.id, mae_prov.numruc, mae_prov.nombre " _
                & " FROM mae_prov RIGHT JOIN com_reembolsables ON mae_prov.id = com_reembolsables.idpro  " _
                & " WHERE (com_reembolsables.impsal<>0) "
                
            xForm.Titulo = "Buscando Proveedores"
        
        Case 19 '--Planilla de Letras
            xCampos(0, 0) = "Banco":          xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Nº Cta Cte":     xCampos(1, 1) = "numcue":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"
            
            nSQL = "SELECT mae_banconumcta.id, mae_banconumcta.numcue , Trim([mae_bancos].[descripcion]) AS nombre1,Trim([mae_bancos].[descripcion]) & ' Nro Cta. ' & trim(mae_banconumcta.numcue) as nombre " _
                + vbCr + " FROM mae_bancos RIGHT JOIN (mae_banconumcta RIGHT JOIN let_planilla ON mae_banconumcta.id = let_planilla.idbcocta) ON mae_bancos.id = mae_banconumcta.idban " _
                + vbCr + " Where (((let_planilla.impsal) <> 0)) " _
                + vbCr + " GROUP BY mae_banconumcta.id, mae_banconumcta.numcue, Trim([mae_bancos].[descripcion]); "
            
            xForm.Titulo = "Buscando Cta Cte"
        
        Case 20 '--orden de despacho
            nSQL = "SELECT var_ordendespacho.idcli as id, mae_cliente.numruc, mae_cliente.nombre " _
                + vbCr + " FROM var_ordendespacho INNER JOIN mae_cliente ON var_ordendespacho.idcli = mae_cliente.id " _
                + vbCr + " GROUP BY var_ordendespacho.idcli, mae_cliente.numruc, mae_cliente.nombre;"

            xForm.Titulo = "Buscando Clientes"
            xCampos(0, 0) = "Cliente":
            
        Case Else
        
            Exit Sub
    End Select

    xForm.SQLCad = nSQL
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "nombre"
    xForm.CampoBusca = "nombre"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtProv.Text = NulosC(xRs("nombre"))
        LblIdCliente.Caption = xRs("id")
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Function CadWhere(idDestino As Integer, Tabla As Integer, tipo As Integer) As String
    'tabla = 1 tabla de compras
    'tabla = 2 tabla de recibos
    
    'Tipo = 1 Origen
    'Tipo = 2 Destino
    'esta funcion permite filtrar a los proveedores cuyos documentos esten en la lista de documentos del destino del egreso
    Dim Rst2 As New ADODB.Recordset
    Dim A As Integer
    Dim xCadWhere As String
    'preparamos la linea WHERE de la consulta para ver los documentos que tenga asignado el destino del egreso
    If tipo = 1 Then RST_Busq Rst2, "SELECT * FROM tes_origendoc WHERE id = " & idDestino & "", xCon
    If tipo = 2 Then RST_Busq Rst2, "SELECT * FROM tes_destinodoc WHERE id = " & idDestino & "", xCon
    
    If Rst2.RecordCount <> 0 Then
        Rst2.MoveFirst
        For A = 1 To Rst2.RecordCount
            If Tabla = 1 Then xCadWhere = xCadWhere + "(com_compras.tipdoc=" & Rst2("iddoc") & ")"
            If Tabla = 2 Then xCadWhere = xCadWhere + "(con_recibos.tipdoc=" & Rst2("iddoc") & ")"
            Rst2.MoveNext
            If Rst2.EOF = True Then Exit For
            xCadWhere = xCadWhere + " OR "
        Next A
    End If
    Set Rst2 = Nothing
    CadWhere = xCadWhere
End Function

Private Sub CmdBusMon_Click()
    If QueHace = 3 Then Exit Sub

    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Codigo":       xCampos(0, 1) = "id":            xCampos(0, 2) = "1200":         xCampos(0, 3) = "N"
    xCampos(1, 0) = "Descripcion":  xCampos(1, 1) = "descripcion":   xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
    
    'filtramos por tipo de movimiento  = 1 (Ingreso)
    xForm.SQLCad = "SELECT * FROM  mae_moneda ORDER BY descripcion"

    xForm.Titulo = "Buscando Moneda"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "id"
    xForm.CampoBusca = "id"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtIdMon.Text = xRs("id")
        LblMoneda.Caption = xRs("descripcion")
        Fg1.SetFocus
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Sub ActualizarImportesRstTmp()
    Dim A As Integer
    If RstTMPDoc.State = 0 Then Exit Sub
    RstTMPDoc.Filter = adFilterNone
    
    If RstTMPDoc.RecordCount <> 0 Then
        RstTMPDoc.MoveFirst
        For A = 1 To RstTMPDoc.RecordCount
            If NulosN(TxtIdMon.Text) = RstTMPDoc("idmone") Then
                RstTMPDoc("impsal2") = RstTMPDoc("impsal")
            Else
                If RstTMPDoc("idmone") = 1 Then
                    If NulosN(TxtIdMon.Text) = 2 Then
                        RstTMPDoc("impsal2") = NulosN(RstTMPDoc("impsal")) * NulosN(lblTipCambio.Caption)
                    End If
                End If
                If RstTMPDoc("idmone") = 2 Then
                    If NulosN(TxtIdMon.Text) = 1 Then
                        RstTMPDoc("impsal2") = NulosN(RstTMPDoc("impsal")) * NulosN(lblTipCambio.Caption)
                    End If
                End If
            End If
            If NulosN(TxtIdMon.Text) = 1 Then
                RstTMPDoc("acuent") = Format(RstTMPDoc("acuent") * NulosN(lblTipCambio.Caption), "0.00")
            Else
                RstTMPDoc("acuent") = Format(RstTMPDoc("acuent") / NulosN(lblTipCambio.Caption), FORMAT_MONTO)
            End If
            RstTMPDoc("newsal") = Format(RstTMPDoc("impsal2") - RstTMPDoc("acuent"), FORMAT_MONTO)
            
            RstTMPDoc.MoveNext
            If RstTMPDoc.EOF = True Then Exit For
        Next A
    End If
End Sub

Private Sub CmdDelCon_Click()
    If Fg2.Row < 1 Then Exit Sub
    If Fg2.Rows = 1 Then Exit Sub
    
    RstTMPDoc.Filter = adFilterNone
    
    RstTMPDoc.Filter = "xtipo=1 AND idconc = " & NulosN(Fg2.TextMatrix(Fg2.Row, 3)) & " "
    
    'eliminamos los documentos del concepto si es que los tuviera
    If RstTMPDoc.RecordCount <> 0 Then
        RstTMPDoc.MoveFirst
        Dim A As Integer
        
        For A = 1 To RstTMPDoc.RecordCount
            RstTMPDoc.Delete
            RstTMPDoc.MoveNext
            If RstTMPDoc.EOF = True Then Exit For
        Next A
    End If
    Fg2.RemoveItem Fg2.Row
    TotalizarFG2
End Sub

Private Sub CmdEliminar_Click()
    If QueHace = 3 Then Exit Sub
    If Fg3.Row < 1 Then
        MsgBox "Seleccione una fila para seleccionar", vbExclamation, xTitulo
        Exit Sub
    End If
    If Fg3.Rows = 1 Then Exit Sub
    
    RstTMPDoc.Filter = adFilterNone
        
    RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & NulosN(Fg3.TextMatrix(Fg3.Row, 11)) & " AND corr= " & NulosN(Fg3.TextMatrix(Fg3.Row, 15)) & ""
    
    If RstTMPDoc.RecordCount = 1 Then
        RstTMPDoc.Delete
    End If
    
    RstTMPDoc.Filter = adFilterNone
    RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & Fg3.TextMatrix(Fg3.Row, 11) & ""
    Fg3.RemoveItem Fg3.Row
    TotalizarFG3
End Sub

Private Sub Command1_Click()
'--1=Compras; 8=Boleta de Pago; 9=Honorarios; 10=Reembolsables;
'--2=Ventas; 4=Letras; 19=Planilla del Letras
'--6=Bancos
'--20=Orden de Despacho
    
    Agregando = True
    
    'si es origen
    If mVentanaActiva = 0 Then
        Select Case IdModulo
        
            Case 1, 8, 9, 10, 4, 19, 2, 11
                
                If TxtIdMon.Text = "1" Then
                    Fg1.TextMatrix(Fg1.Row, 7) = Format(TxtTotal4.Text, FORMAT_MONTO)
                    If TipCam <> 0 Then Fg1.TextMatrix(Fg1.Row, 8) = NulosN(TxtTotal4.Text) / TipCam
                    Fg1.TextMatrix(Fg1.Row, 8) = Format(Fg1.TextMatrix(Fg1.Row, 8), FORMAT_MONTO)
                    
                Else
                    Fg1.TextMatrix(Fg1.Row, 8) = Format(TxtTotal4.Text, FORMAT_MONTO)
                    Fg1.TextMatrix(Fg1.Row, 7) = NulosN(TxtTotal4.Text) * TipCam
                    Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), FORMAT_MONTO)
                End If
        
            Case Else
    
                
        End Select
    
    Else
        '--colocar el total en vista principal (destino del egreso)
        Select Case IdModulo
            Case 7 'fondo fijo
                If TxtIdMon.Text = "1" Then
                    Fg2.TextMatrix(Fg2.Row, 7) = Format(NulosN(Fg3.TextMatrix(1, 6)), FORMAT_MONTO)
                    If TipCam <> 0 Then Fg2.TextMatrix(Fg2.Row, 8) = NulosN(Fg3.TextMatrix(1, 6)) / TipCam
                    Fg2.TextMatrix(Fg2.Row, 8) = Format(Fg2.TextMatrix(Fg2.Row, 8), FORMAT_MONTO)
                    
                Else
                    Fg2.TextMatrix(Fg2.Row, 8) = Format(NulosN(Fg3.TextMatrix(1, 6)), FORMAT_MONTO)
                    Fg2.TextMatrix(Fg2.Row, 7) = NulosN(Fg3.TextMatrix(1, 6)) * TipCam
                    Fg2.TextMatrix(Fg2.Row, 7) = Format(Fg2.TextMatrix(Fg2.Row, 7), FORMAT_MONTO)
                    
                End If
            
            
            Case 3, 5
                '--3::Entregas a rengir
                '--5::Anticipos
                If TxtIdMon.Text = "1" Then
                    Fg2.TextMatrix(Fg2.Row, 7) = Format(NulosN(TxtTotal1.Text), FORMAT_MONTO)
                    If TipCam <> 0 Then Fg2.TextMatrix(Fg2.Row, 8) = NulosN(TxtTotal1.Text) / TipCam
                    Fg2.TextMatrix(Fg2.Row, 8) = Format(Fg2.TextMatrix(Fg2.Row, 8), FORMAT_MONTO)
                    
                Else
                    Fg2.TextMatrix(Fg2.Row, 8) = Format(NulosN(TxtTotal1.Text), FORMAT_MONTO)
                    Fg2.TextMatrix(Fg2.Row, 7) = NulosN(TxtTotal1.Text) * TipCam
                    Fg2.TextMatrix(Fg2.Row, 7) = Format(Fg2.TextMatrix(Fg2.Row, 7), FORMAT_MONTO)
                End If
            
            Case 1, 8, 9, 10, 6, 4, 19, 20, 2, 11
                    
                If TxtIdMon.Text = "1" Then
                    Fg2.TextMatrix(Fg2.Row, 7) = Format(TxtTotal4.Text, FORMAT_MONTO)
                    If TipCam <> 0 Then Fg2.TextMatrix(Fg2.Row, 8) = NulosN(TxtTotal4.Text) / TipCam
                    Fg2.TextMatrix(Fg2.Row, 8) = Format(Fg2.TextMatrix(Fg2.Row, 8), FORMAT_MONTO)
                    
                Else
                    Fg2.TextMatrix(Fg2.Row, 8) = Format(TxtTotal4.Text, FORMAT_MONTO)
                    Fg2.TextMatrix(Fg2.Row, 7) = NulosN(TxtTotal4.Text) * TipCam
                    Fg2.TextMatrix(Fg2.Row, 7) = Format(Fg2.TextMatrix(Fg2.Row, 7), FORMAT_MONTO)
                End If
            
            Case Else
        
        End Select
    
    End If
    
    
    Agregando = False
    ActivarEntorno
    Frame3.Visible = False
    
    If mVentanaActiva = 0 Then
        '--origen
        TotalizarFG1
        Fg1.SetFocus
    Else
        '--destino
        TotalizarFG2
        Fg2.SetFocus
    End If
End Sub

Private Sub Command10_Click()
    ActivarEntorno
    'TotalizarFG1
    Frame12.Visible = False
    Fg1.SetFocus
End Sub

Private Sub Command11_Click()

    ActivarEntorno
    
    Frame3.Visible = False

    If mVentanaActiva = 0 Then
        TotalizarFG1
        Fg1.SetFocus
    Else
        TotalizarFG2
        Fg2.SetFocus
    End If


End Sub

Private Sub Command2_Click()
    ActivarEntorno
    MostrarAsiento True
    If Command5.Visible = True Then Command5.SetFocus
End Sub

Sub MostrarAsiento(VerVentana As Boolean)
    '--ventana: indica si se mostrara en el formulario
    '-- true: mostrar ventana; false: no mostrar

    Frame9.Left = 270
    Frame9.Top = 1500
    Frame9.Visible = VerVentana
    
    Dim A As Integer
    Dim TotDebSol, TotHabSol, TotDebDol, TotHabDol, TotalCambio As Double
    Fg4.Rows = 1
    Fg5.Rows = 1
    TxtSaldoCambio = 0
    
    Dim TotCuenta As Double
    Dim RstTmp As New ADODB.Recordset
    
    Set RstTmp = PreparaRST2
       
    Dim B As Integer
    
    TxtSaldoCambio.Text = Format(TotalCambio, FORMAT_MONTO)
    
    'mostramos el debe
    For A = 1 To Fg2.Rows - 1
        RstTMPDoc.Filter = adFilterNone
        RstTMPDoc.Sort = "idconc"
        RstTMPDoc.Filter = "xtipo=1 AND idconc = " & NulosN(Fg2.TextMatrix(A, 3)) & ""
        
        If RstTMPDoc.RecordCount <> 0 Then
            'si el concepto tiene detalle los mostramos
            RstTMPDoc.MoveFirst
            For B = 1 To RstTMPDoc.RecordCount
                Fg4.Rows = Fg4.Rows + 1
                Fg4.TextMatrix(Fg4.Rows - 1, 1) = Busca_Codigo(NulosN(Fg2.TextMatrix(A, 4)), "id", "cuenta", "con_planctas", "N", xCon)
                Fg4.TextMatrix(Fg4.Rows - 1, 2) = Busca_Codigo(NulosN(Fg2.TextMatrix(A, 4)), "id", "descripcion", "con_planctas", "N", xCon)
                Fg4.TextMatrix(Fg4.Rows - 1, 3) = NulosC(RstTMPDoc("tipdoc"))
                Fg4.TextMatrix(Fg4.Rows - 1, 4) = NulosC(RstTMPDoc("numdoc"))
                Fg4.TextMatrix(Fg4.Rows - 1, 5) = Format(RstTMPDoc("fchemi"), "dd/mm/yy")
'                Fg4.TextMatrix(Fg4.Rows - 1, 6) = Format(NulosC(RstTMPDoc("acuent")), "0.00")
                Fg4.TextMatrix(Fg4.Rows - 1, 10) = NulosN(Fg2.TextMatrix(A, 3)) ' id del origen o destino
                Fg4.TextMatrix(Fg4.Rows - 1, 11) = NulosN(Fg2.TextMatrix(A, 5)) ' idmodulo
                
                    
                If TxtIdMon.Text = "1" Then
                    If Fg2.TextMatrix(A, 5) = 7 Or Fg2.TextMatrix(A, 5) = 5 Then
                        Fg4.TextMatrix(Fg4.Rows - 1, 6) = Format(RstTMPDoc("imptot"), "0.00")
                        If NulosN(Fg2.TextMatrix(A, 2)) <> 0 Then
                            Fg4.TextMatrix(Fg4.Rows - 1, 8) = RstTMPDoc("imptot") / NulosN(Fg2.TextMatrix(A, 2))
                        Else
                            Fg4.TextMatrix(Fg4.Rows - 1, 8) = 0
                        End If
                        Fg4.TextMatrix(Fg4.Rows - 1, 12) = 0
                        
                    ElseIf Fg2.TextMatrix(A, 5) = 6 Then
                        Fg4.TextMatrix(Fg4.Rows - 1, 6) = Format(RstTMPDoc("imptot"), "0.00")
                        Fg4.TextMatrix(Fg4.Rows - 1, 8) = RstTMPDoc("imptot") / NulosN(Fg2.TextMatrix(A, 2))
                        
                     Else
                        Fg4.TextMatrix(Fg4.Rows - 1, 12) = NulosN(RstTMPDoc("iddocu"))
                        Fg4.TextMatrix(Fg4.Rows - 1, 6) = Format(RstTMPDoc("acuent"), "0.00")
                        If NulosN(Fg2.TextMatrix(A, 2)) <> 0 Then
                            Fg4.TextMatrix(Fg4.Rows - 1, 8) = RstTMPDoc("acuent") / NulosN(Fg2.TextMatrix(A, 2))
                        Else
                            Fg4.TextMatrix(Fg4.Rows - 1, 8) = 0
                        End If
                        
                   End If
                    
                    
                    Fg4.TextMatrix(Fg4.Rows - 1, 8) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 8), FORMAT_MONTO)
                Else
                    'aqui preguntamos si es Anticipos, Entregas a rendir o fondo fijo
                    If Fg2.TextMatrix(A, 5) = 7 Or Fg2.TextMatrix(A, 5) = 5 Or Fg2.TextMatrix(A, 5) = 3 Then
                        Fg4.TextMatrix(Fg4.Rows - 1, 12) = NulosN(RstTMPDoc("iddocu"))
                        
                        Fg4.TextMatrix(Fg4.Rows - 1, 8) = Format(RstTMPDoc("imptot"), "0.00")
                        Fg4.TextMatrix(Fg4.Rows - 1, 6) = RstTMPDoc("imptot") * NulosN(Fg2.TextMatrix(A, 2))
                        
                    ElseIf Fg2.TextMatrix(A, 5) = 1 Or Fg2.TextMatrix(A, 5) = 10 Then '--compras u honorarios
                        Fg4.TextMatrix(Fg4.Rows - 1, 12) = NulosN(RstTMPDoc("iddocu"))
                        
                        Fg4.TextMatrix(Fg4.Rows - 1, 8) = Format(RstTMPDoc("acuent"), "0.00")
                        Fg4.TextMatrix(Fg4.Rows - 1, 6) = RstTMPDoc("acuent") * NulosN(Fg2.TextMatrix(A, 2))
                        
                    ElseIf Fg2.TextMatrix(A, 5) = 6 Then
                        Fg4.TextMatrix(Fg4.Rows - 1, 8) = Format(RstTMPDoc("imptot"), "0.00")
                        Fg4.TextMatrix(Fg4.Rows - 1, 6) = RstTMPDoc("imptot") * NulosN(Fg2.TextMatrix(A, 2))
                        
                    Else
                       ''Fg4.TextMatrix(Fg4.Rows - 1, 12) = NulosN(RstTMPDoc("iddocu"))
                       Fg4.TextMatrix(Fg4.Rows - 1, 6) = Format(RstTMPDoc("acuent"), "0.00")
                       If NulosN(Fg2.TextMatrix(A, 2)) <> 0 Then
                           Fg4.TextMatrix(Fg4.Rows - 1, 8) = RstTMPDoc("acuent") * NulosN(Fg2.TextMatrix(A, 2))
                       Else
                           Fg4.TextMatrix(Fg4.Rows - 1, 8) = 0
                       End If
                        
                        
                    End If
                    
                    Fg4.TextMatrix(Fg4.Rows - 1, 6) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 6), FORMAT_MONTO)
                End If
                
                RstTMPDoc.MoveNext
                If RstTMPDoc.EOF = True Then Exit For
            Next B
        Else
            'si no tiene detalle mostramos los datos del concepto
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(Fg4.Rows - 1, 1) = Busca_Codigo(NulosN(Fg2.TextMatrix(A, 4)), "id", "cuenta", "con_planctas", "N", xCon)
            Fg4.TextMatrix(Fg4.Rows - 1, 2) = Busca_Codigo(NulosN(Fg2.TextMatrix(A, 4)), "id", "descripcion", "con_planctas", "N", xCon)
            Fg4.TextMatrix(Fg4.Rows - 1, 3) = ""
            Fg4.TextMatrix(Fg4.Rows - 1, 4) = ""
            Fg4.TextMatrix(Fg4.Rows - 1, 5) = ""
                
            Fg4.TextMatrix(Fg4.Rows - 1, 10) = NulosN(Fg2.TextMatrix(A, 3)) ' id del origen o destino
            Fg4.TextMatrix(Fg4.Rows - 1, 11) = NulosN(Fg2.TextMatrix(A, 5)) ' idmodulo
            
            If NulosN(Fg2.TextMatrix(A, 5)) = 6 Then Fg4.TextMatrix(Fg4.Rows - 1, 12) = 0
                        
            If NulosN(Fg2.TextMatrix(A, 5)) = 1 Then
                If RstTMPDoc.RecordCount <> 0 Then Fg4.TextMatrix(Fg4.Rows - 1, 12) = NulosN(RstTMPDoc("iddocu"))
            End If
                
            If TxtIdMon.Text = "1" Then
                Fg4.TextMatrix(Fg4.Rows - 1, 6) = Format(Fg2.TextMatrix(A, 7), "0.00")
                If NulosN(Fg2.TextMatrix(A, 2)) <> 0 Then
                    Fg4.TextMatrix(Fg4.Rows - 1, 8) = NulosN(Fg2.TextMatrix(A, 7)) / NulosN(Fg2.TextMatrix(A, 2))
                Else
                    Fg4.TextMatrix(Fg4.Rows - 1, 8) = 0
                End If
                Fg4.TextMatrix(Fg4.Rows - 1, 8) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 8), FORMAT_MONTO)
            Else
                Fg4.TextMatrix(Fg4.Rows - 1, 8) = Format(Fg2.TextMatrix(A, 8), "0.00")
                Fg4.TextMatrix(Fg4.Rows - 1, 6) = NulosN(Fg2.TextMatrix(A, 8)) * NulosN(Fg2.TextMatrix(A, 2))
                Fg4.TextMatrix(Fg4.Rows - 1, 6) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 6), FORMAT_MONTO)
            End If
            
        End If
    Next A

    'mostramos el Haber
    For A = 1 To Fg1.Rows - 1
        
        '--si el modulo de compras, ventas, letras, lgd, planilla letras
        If NulosN(Fg1.TextMatrix(A, 5)) = 1 Or NulosN(Fg1.TextMatrix(A, 5)) = 2 Or NulosN(Fg1.TextMatrix(A, 5)) = 4 Or NulosN(Fg1.TextMatrix(A, 5)) = 11 Or NulosN(Fg1.TextMatrix(A, 5)) = 19 Then
        
            RstTMPDoc.Filter = adFilterNone
            RstTMPDoc.Sort = "idconc"
            RstTMPDoc.Filter = "xtipo=0 AND idconc = " & NulosN(Fg1.TextMatrix(A, 3)) & ""
            
            If RstTMPDoc.RecordCount <> 0 Then
                'si el concepto tiene detalle los mostramos
                RstTMPDoc.MoveFirst
                For B = 1 To RstTMPDoc.RecordCount
                    Fg4.Rows = Fg4.Rows + 1
                    Fg4.TextMatrix(Fg4.Rows - 1, 1) = Busca_Codigo(NulosN(Fg1.TextMatrix(A, 4)), "id", "cuenta", "con_planctas", "N", xCon)
                    Fg4.TextMatrix(Fg4.Rows - 1, 2) = Busca_Codigo(NulosN(Fg1.TextMatrix(A, 4)), "id", "descripcion", "con_planctas", "N", xCon)
                    Fg4.TextMatrix(Fg4.Rows - 1, 3) = NulosC(RstTMPDoc("tipdoc"))
                    Fg4.TextMatrix(Fg4.Rows - 1, 4) = NulosC(RstTMPDoc("numdoc"))
                    Fg4.TextMatrix(Fg4.Rows - 1, 5) = Format(RstTMPDoc("fchemi"), "dd/mm/yy")
    '                Fg4.TextMatrix(Fg4.Rows - 1, 6) = Format(NulosC(RstTMPDoc("acuent")), "0.00")
                    Fg4.TextMatrix(Fg4.Rows - 1, 10) = NulosN(Fg1.TextMatrix(A, 3)) ' id del origen o destino
                    Fg4.TextMatrix(Fg4.Rows - 1, 11) = NulosN(Fg1.TextMatrix(A, 5)) ' idmodulo
                    
                    If TxtIdMon.Text = "1" Then

                        Fg4.TextMatrix(Fg4.Rows - 1, 12) = NulosN(RstTMPDoc("iddocu"))
                        Fg4.TextMatrix(Fg4.Rows - 1, 7) = Format(RstTMPDoc("acuent"), "0.00")
                        If NulosN(Fg1.TextMatrix(A, 2)) <> 0 Then
                            Fg4.TextMatrix(Fg4.Rows - 1, 9) = RstTMPDoc("acuent") / NulosN(Fg1.TextMatrix(A, 2))
                        Else
                            Fg4.TextMatrix(Fg4.Rows - 1, 9) = 0
                        End If

                        Fg4.TextMatrix(Fg4.Rows - 1, 9) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 9), FORMAT_MONTO)
                    Else

                        Fg4.TextMatrix(Fg4.Rows - 1, 12) = NulosN(RstTMPDoc("iddocu"))
                        
                        Fg4.TextMatrix(Fg4.Rows - 1, 9) = Format(RstTMPDoc("acuent"), "0.00")
                        Fg4.TextMatrix(Fg4.Rows - 1, 7) = RstTMPDoc("acuent") * NulosN(Fg1.TextMatrix(A, 2))
                        
                        Fg4.TextMatrix(Fg4.Rows - 1, 7) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 7), FORMAT_MONTO)
                    End If
                    
                    RstTMPDoc.MoveNext
                    If RstTMPDoc.EOF = True Then Exit For
                Next B
            Else
                
                
            End If
            
        Else
    
            RstTmpDocOri.Filter = adFilterNone
            RstTmpDocOri.Filter = "idconc = " & NulosN(Fg1.TextMatrix(A, 3)) & ""
            
            If RstTmpDocOri.RecordCount <> 0 Then
                RstTmpDocOri.MoveFirst
                For B = 1 To RstTmpDocOri.RecordCount
                    Fg4.Rows = Fg4.Rows + 1
                    Fg4.TextMatrix(Fg4.Rows - 1, 1) = Busca_Codigo(NulosN(Fg1.TextMatrix(A, 4)), "id", "cuenta", "con_planctas", "N", xCon)
                    Fg4.TextMatrix(Fg4.Rows - 1, 2) = Busca_Codigo(NulosN(Fg1.TextMatrix(A, 4)), "id", "descripcion", "con_planctas", "N", xCon)
                    Fg4.TextMatrix(Fg4.Rows - 1, 3) = NulosC(RstTmpDocOri("tipdoc"))
                    Fg4.TextMatrix(Fg4.Rows - 1, 4) = NulosC(RstTmpDocOri("numdoc"))
                    Fg4.TextMatrix(Fg4.Rows - 1, 5) = Format(RstTmpDocOri("fchemi"), "dd/mm/yy")
                    
                    Fg4.TextMatrix(Fg4.Rows - 1, 10) = NulosN(Fg1.TextMatrix(A, 3)) ' id del origen o destino
                    Fg4.TextMatrix(Fg4.Rows - 1, 11) = NulosN(Fg1.TextMatrix(A, 5)) ' idmodulo
                    If NulosN(Fg1.TextMatrix(A, 5)) = 6 Then Fg4.TextMatrix(Fg4.Rows - 1, 12) = 0
                    If NulosN(Fg1.TextMatrix(A, 5)) = 1 Then Fg4.TextMatrix(Fg4.Rows - 1, 12) = NulosN(RstTmpDocOri("iddocu"))
                    
                    If TxtIdMon.Text = "1" Then
                        Fg4.TextMatrix(Fg4.Rows - 1, 7) = Format(RstTmpDocOri("imptot"), "0.00")
                        If NulosN(Fg1.TextMatrix(A, 2)) <> 0 Then
                            Fg4.TextMatrix(Fg4.Rows - 1, 9) = NulosN(RstTmpDocOri("imptot")) / NulosN(Fg1.TextMatrix(A, 2))
                        Else
                            Fg4.TextMatrix(Fg4.Rows - 1, 9) = 0
                        End If
                        
                        Fg4.TextMatrix(Fg4.Rows - 1, 9) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 9), FORMAT_MONTO)
                    Else
                        Fg4.TextMatrix(Fg4.Rows - 1, 9) = Format(RstTmpDocOri("imptot"), "0.00")
                        Fg4.TextMatrix(Fg4.Rows - 1, 7) = RstTmpDocOri("imptot") * NulosN(Fg1.TextMatrix(A, 2))
                        Fg4.TextMatrix(Fg4.Rows - 1, 7) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 7), FORMAT_MONTO)
                    End If
                    
                    RstTmpDocOri.MoveNext
                    If RstTmpDocOri.EOF = True Then Exit For
                Next B
            Else
                Fg4.Rows = Fg4.Rows + 1
                Fg4.TextMatrix(Fg4.Rows - 1, 1) = Busca_Codigo(NulosN(Fg1.TextMatrix(A, 4)), "id", "cuenta", "con_planctas", "N", xCon)
                Fg4.TextMatrix(Fg4.Rows - 1, 2) = Busca_Codigo(NulosN(Fg1.TextMatrix(A, 4)), "id", "descripcion", "con_planctas", "N", xCon)
                Fg4.TextMatrix(Fg4.Rows - 1, 3) = ""  'NulosC(RstTmpDocOri("tipdoc"))
                Fg4.TextMatrix(Fg4.Rows - 1, 4) = "" 'NulosC(RstTmpDocOri("numdoc"))
                Fg4.TextMatrix(Fg4.Rows - 1, 5) = "" 'Format(RstTmpDocOri("fchemi"), "dd/mm/yy")
                
                If TxtIdMon.Text = "1" Then
                    Fg4.TextMatrix(Fg4.Rows - 1, 7) = Format(NulosN(Fg1.TextMatrix(A, 7)), "0.00")
                    If NulosN(Fg1.TextMatrix(A, 2)) <> 0 Then
                        Fg4.TextMatrix(Fg4.Rows - 1, 9) = NulosN(Fg1.TextMatrix(A, 7)) / NulosN(Fg1.TextMatrix(A, 2))
                    Else
                        Fg4.TextMatrix(Fg4.Rows - 1, 9) = 0
                    End If
                    Fg4.TextMatrix(Fg4.Rows - 1, 9) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 9), FORMAT_MONTO)
                Else
                    Fg4.TextMatrix(Fg4.Rows - 1, 9) = Format(NulosN(Fg1.TextMatrix(A, 8)), "0.00")
                    Fg4.TextMatrix(Fg4.Rows - 1, 7) = NulosN(Fg1.TextMatrix(A, 8)) * NulosN(Fg1.TextMatrix(A, 2))
                    Fg4.TextMatrix(Fg4.Rows - 1, 7) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 7), FORMAT_MONTO)
                End If
            End If
        
        End If


    Next A
    
    'CARGAMOS LOS DESTINOS
    For A = 1 To Fg2.Rows - 1
        RstTMPDoc.Filter = adFilterNone
        RstTMPDoc.Sort = "idconc"
        RstTMPDoc.Filter = "xtipo=1 AND idconc = " & NulosN(Fg2.TextMatrix(A, 3)) & ""
        
'            'agregamos las cuentas de destino
            
            Dim xIdCtaDes As Integer
            'DEBE Destino
            If NulosN(Busca_Codigo(NulosN(Fg2.TextMatrix(A, 4)), "id", "ctadesdeb", "con_planctas", "N", xCon)) <> 0 Then
                xIdCtaDes = Busca_Codigo(NulosN(Fg2.TextMatrix(A, 4)), "id", "ctadesdeb", "con_planctas", "N", xCon)
                Fg4.Rows = Fg4.Rows + 1
                Fg4.TextMatrix(Fg4.Rows - 1, 1) = Busca_Codigo(NulosN(xIdCtaDes), "id", "cuenta", "con_planctas", "N", xCon)
                Fg4.TextMatrix(Fg4.Rows - 1, 2) = Busca_Codigo(NulosN(xIdCtaDes), "id", "descripcion", "con_planctas", "N", xCon)
                Fg4.TextMatrix(Fg4.Rows - 1, 3) = ""
                Fg4.TextMatrix(Fg4.Rows - 1, 4) = ""
                Fg4.TextMatrix(Fg4.Rows - 1, 5) = ""
                    
                Fg4.TextMatrix(Fg4.Rows - 1, 10) = NulosN(Fg2.TextMatrix(A, 3)) ' id del origen o destino
                Fg4.TextMatrix(Fg4.Rows - 1, 11) = NulosN(Fg2.TextMatrix(A, 5)) ' idmodulo
                Fg4.TextMatrix(Fg4.Rows - 1, 13) = -1  'con esta columna sabemos que el asiento es automatico
                
                If NulosN(Fg2.TextMatrix(A, 5)) = 6 Then Fg4.TextMatrix(Fg4.Rows - 1, 12) = 0
                            
                If NulosN(Fg2.TextMatrix(A, 5)) = 1 Then Fg4.TextMatrix(Fg4.Rows - 1, 12) = NulosN(RstTMPDoc("iddocu"))
                    
                If TxtIdMon.Text = "1" Then
                    Fg4.TextMatrix(Fg4.Rows - 1, 6) = Format(Fg2.TextMatrix(A, 7), "0.00")
                    If NulosN(Fg2.TextMatrix(A, 2)) <> 0 Then
                        Fg4.TextMatrix(Fg4.Rows - 1, 8) = NulosN(Fg2.TextMatrix(A, 7)) / NulosN(Fg2.TextMatrix(A, 2))
                    Else
                        Fg4.TextMatrix(Fg4.Rows - 1, 8) = 0
                    End If
                    Fg4.TextMatrix(Fg4.Rows - 1, 8) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 8), FORMAT_MONTO)
                Else
                    Fg4.TextMatrix(Fg4.Rows - 1, 8) = Format(Fg2.TextMatrix(A, 8), "0.00")
                    Fg4.TextMatrix(Fg4.Rows - 1, 6) = NulosN(Fg2.TextMatrix(A, 8)) * NulosN(Fg2.TextMatrix(A, 2))
                    Fg4.TextMatrix(Fg4.Rows - 1, 6) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 6), FORMAT_MONTO)
                End If
            End If
            
            'HABER Destino
            If NulosN(Busca_Codigo(NulosN(Fg2.TextMatrix(A, 4)), "id", "ctadeshab", "con_planctas", "N", xCon)) <> 0 Then
                xIdCtaDes = Busca_Codigo(NulosN(Fg2.TextMatrix(A, 4)), "id", "ctadeshab", "con_planctas", "N", xCon)
                Fg4.Rows = Fg4.Rows + 1
                Fg4.TextMatrix(Fg4.Rows - 1, 1) = Busca_Codigo(NulosN(xIdCtaDes), "id", "cuenta", "con_planctas", "N", xCon)
                Fg4.TextMatrix(Fg4.Rows - 1, 2) = Busca_Codigo(NulosN(xIdCtaDes), "id", "descripcion", "con_planctas", "N", xCon)
                Fg4.TextMatrix(Fg4.Rows - 1, 3) = ""
                Fg4.TextMatrix(Fg4.Rows - 1, 4) = ""
                Fg4.TextMatrix(Fg4.Rows - 1, 5) = ""
                    
                Fg4.TextMatrix(Fg4.Rows - 1, 10) = NulosN(Fg2.TextMatrix(A, 3)) ' id del origen o destino
                Fg4.TextMatrix(Fg4.Rows - 1, 11) = NulosN(Fg2.TextMatrix(A, 5)) ' idmodulo
                Fg4.TextMatrix(Fg4.Rows - 1, 13) = -1  'con esta columna sabemos que el asiento es automatico
                
                If NulosN(Fg2.TextMatrix(A, 5)) = 6 Then Fg4.TextMatrix(Fg4.Rows - 1, 12) = 0
                            
                If NulosN(Fg2.TextMatrix(A, 5)) = 1 Then Fg4.TextMatrix(Fg4.Rows - 1, 12) = NulosN(RstTMPDoc("iddocu"))
                    
                If TxtIdMon.Text = "1" Then
                    Fg4.TextMatrix(Fg4.Rows - 1, 7) = Format(Fg2.TextMatrix(A, 7), "0.00")
                    If NulosN(Fg2.TextMatrix(A, 2)) <> 0 Then
                        Fg4.TextMatrix(Fg4.Rows - 1, 9) = NulosN(Fg2.TextMatrix(A, 7)) / NulosN(Fg2.TextMatrix(A, 2))
                    Else
                        Fg4.TextMatrix(Fg4.Rows - 1, 9) = 0
                    End If
                    Fg4.TextMatrix(Fg4.Rows - 1, 9) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 9), FORMAT_MONTO)
                Else
                    Fg4.TextMatrix(Fg4.Rows - 1, 9) = Format(Fg2.TextMatrix(A, 8), "0.00")
                    Fg4.TextMatrix(Fg4.Rows - 1, 7) = NulosN(Fg2.TextMatrix(A, 8)) * NulosN(Fg2.TextMatrix(A, 2))
                    Fg4.TextMatrix(Fg4.Rows - 1, 7) = Format(Fg4.TextMatrix(Fg4.Rows - 1, 7), FORMAT_MONTO)
                End If
            End If
         
    Next A
    
    'mostramos el total
    For A = 1 To Fg4.Rows - 1
        TotDebSol = TotDebSol + NulosN(Fg4.TextMatrix(A, 6))
        TotHabSol = TotHabSol + NulosN(Fg4.TextMatrix(A, 7))
        TotDebDol = TotDebDol + NulosN(Fg4.TextMatrix(A, 8))
        TotHabDol = TotHabDol + NulosN(Fg4.TextMatrix(A, 9))
    Next A
    
    TxtDebSol.Text = Format(TotDebSol, FORMAT_MONTO)
    TxtHabSol.Text = Format(TotHabSol, FORMAT_MONTO)
    TxtDebDol.Text = Format(TotDebDol, FORMAT_MONTO)
    TxtHabDol.Text = Format(TotHabDol, FORMAT_MONTO)
    
End Sub

Private Sub Command3_Click()
    If Fg1.Rows = 1 Then
        Fg1.Rows = Fg1.Rows + 1
        Fg1.Row = Fg1.Rows - 1
        Fg1_CellButtonClick Fg1.Rows - 1, 1
    Else
        If NulosN(Fg1.TextMatrix(Fg1.Rows - 1, 3)) = 0 Then
            MsgBox "No ha especificado un concepto para la ultima fila del origen de egresos", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Fg1.Rows = Fg1.Rows + 1
        Fg1.Row = Fg1.Rows - 1
        Fg1_CellButtonClick Fg1.Rows - 1, 1
    End If
    Fg1.Select Fg1.Rows - 1, 1
    Fg1.SetFocus
End Sub

Private Sub Command4_Click()

    If Fg1.Row < 1 Then Exit Sub
    If Fg1.Rows = 1 Then Exit Sub

    RstTmpDocOri.Filter = adFilterNone
    RstTmpDocOri.Filter = "idconc = " & NulosN(Fg1.TextMatrix(Fg1.Row, 3)) & " "

    'eliminamos los documentos del concepto si es que los tuviera
    If RstTmpDocOri.RecordCount <> 0 Then
        RstTmpDocOri.MoveFirst
        Dim A As Integer

        For A = 1 To RstTmpDocOri.RecordCount
            RstTmpDocOri.Delete
            RstTmpDocOri.MoveNext
            If RstTmpDocOri.EOF = True Then Exit For
        Next A
    End If

    '--------
    RstTMPDoc.Filter = adFilterNone
    
    RstTMPDoc.Filter = "xtipo=0 AND idconc = " & NulosN(Fg1.TextMatrix(Fg1.Row, 3)) & " "
    
    'eliminamos los documentos del concepto si es que los tuviera
    If RstTMPDoc.RecordCount <> 0 Then
        RstTMPDoc.MoveFirst
        
        For A = 1 To RstTMPDoc.RecordCount
            RstTMPDoc.Delete
            RstTMPDoc.MoveNext
            If RstTMPDoc.EOF = True Then Exit For
        Next A
    End If
    '--------

    Fg1.RemoveItem Fg1.Row
    TotalizarFG1
End Sub

Private Sub Command5_Click()
    ActivarEntorno
    Frame9.Visible = False
End Sub

Private Sub Command6_Click()
    If QueHace = 3 Then Exit Sub

    Dim xCadWhere1, xCadWhere2 As String
    
    xCadWhere1 = CadWhere(NulosN(Fg1.TextMatrix(Fg1.Row, 3)), 1, 1)
    xCadWhere2 = CadWhere(NulosN(Fg1.TextMatrix(Fg1.Row, 2)), 2, 1)
    
    If NulosC(xCadWhere1) = "" Then
        MsgBox "El origen seleccionado no tiene documentos de compra asignado para su cancelacion", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    
    Dim xCampos(2, 4) As String
    
    xCampos(0, 0) = "Proveedor":     xCampos(0, 1) = "nombre":        xCampos(0, 2) = "4500":         xCampos(0, 3) = "C"
    xCampos(1, 0) = "Nº R.U.C.":     xCampos(1, 1) = "numruc":        xCampos(1, 2) = "1400":         xCampos(1, 3) = "C"
    
    'buscamos los proveedores que tengan el documento especificado
    xForm.SQLCad = "SELECT con_recibos.id, mae_prov.numruc, mae_prov.nombre FROM con_recibos LEFT JOIN mae_prov ON con_recibos.idcli = mae_prov.id " _
        & " WHERE ((con_recibos.impsal<>0) AND " & xCadWhere2 & ")" _
        & " UNION " _
        & " SELECT DISTINCT mae_prov.id, mae_prov.numruc, mae_prov.nombre FROM mae_prov RIGHT JOIN com_compras ON mae_prov.id = com_compras.idpro  " _
        & " WHERE ((com_compras.impsal<>0) AND " & xCadWhere1 & ")"

    xForm.Titulo = "Buscando Proveedores"
    xForm.FormaBusca = Principio
    xForm.Criterio = ""
    xForm.Ordenado = "nombre"
    xForm.CampoBusca = "nombre"
    Set xForm.Coneccion = xCon
    Set xRs = xForm.BuscarReg(xCampos)
    If xRs.State = 1 Then
        TxtProvA.Text = xRs("nombre")
        LblIdClienteA.Caption = xRs("id")
    End If
    Set xForm = Nothing
    Set xRs = Nothing
End Sub

Private Sub Command7_Click()
    If Fg1.Row < 1 Then Exit Sub
    If Fg1.Rows = 1 Then Exit Sub
    If Fg1.TextMatrix(Fg1.Row, 1) = "" Then
        MsgBox "Seleccione un origen para el egreso", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    If TxtProvA.Enabled = False Then Exit Sub
    CargarFacturasPorCanjear NulosN(LblIdCliente.Caption)
End Sub

Private Sub Command8_Click()
    '--eliminar detalle banco
    If Fg3.Rows = 1 Then Exit Sub
    
    RstTMPDoc.Filter = adFilterNone
    
    RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & Fg3.TextMatrix(Fg3.Row, 11) & " AND iddocu = " & Fg3.TextMatrix(Fg3.Row, 12) & ""
    
    If RstTMPDoc.RecordCount = 1 Then
        RstTMPDoc.Delete
    End If
    
    RstTMPDoc.Filter = adFilterNone
    RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & Fg3.TextMatrix(Fg3.Row, 11) & ""
    Fg3.RemoveItem Fg3.Row
    TotalizarFG3
End Sub

Private Sub Command9_Click()
    Agregando = True
    If NulosN(Fg1.TextMatrix(Fg1.Row, 5)) = 6 Then
        If NulosC(Fg6.TextMatrix(1, 12)) = "" Then
            MsgBox "No ha especificado el medio de pago", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        If NulosC(Fg6.TextMatrix(1, 14)) = "" Then
            MsgBox "No ha especificado el tipo de documento para la operación", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        If NulosC(Fg6.TextMatrix(1, 5)) = "" Then
            MsgBox "No ha especificado el numero de documento para la operación", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        If NulosN(Fg6.TextMatrix(1, 6)) = 0 And ChkChequeAnulado.Value = 0 Then
            MsgBox "No ha especificado el importe para la operación", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
    End If
    
    If NulosN(Fg1.TextMatrix(Fg1.Row, 5)) = 6 Then
        If TxtIdMon.Text = "1" Then
            Fg1.TextMatrix(Fg1.Row, 7) = Format(NulosN(Fg6.TextMatrix(1, 6)), FORMAT_MONTO)
            If NulosN(Fg1.TextMatrix(Fg1.Row, 2)) <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 8) = NulosN(Fg6.TextMatrix(1, 6)) / NulosN(Fg1.TextMatrix(Fg1.Row, 2))
            End If
            Fg1.TextMatrix(Fg1.Row, 8) = Format(Fg1.TextMatrix(Fg1.Row, 8), FORMAT_MONTO)
        Else
            Fg1.TextMatrix(Fg1.Row, 8) = Format(NulosN(Fg6.TextMatrix(1, 6)), FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Row, 7) = NulosN(Fg6.TextMatrix(1, 6)) * NulosN(Fg1.TextMatrix(Fg1.Row, 2))
            Fg1.TextMatrix(Fg1.Row, 7) = Format(Fg1.TextMatrix(Fg1.Row, 7), FORMAT_MONTO)
        End If
    
    Else
        If TxtIdMon.Text = "1" Then
            Fg1.TextMatrix(Fg1.Row, 7) = Format(TxtTotal4A.Text, FORMAT_MONTO)
            If NulosN(Fg1.TextMatrix(Fg1.Row, 2)) <> 0 Then
                Fg1.TextMatrix(Fg1.Row, 8) = NulosN(TxtTotal4A.Text) / NulosN(Fg1.TextMatrix(Fg1.Row, 2))
            End If
        Else
            Fg1.TextMatrix(Fg1.Row, 8) = Format(TxtTotal4A.Text, FORMAT_MONTO)
            Fg1.TextMatrix(Fg1.Row, 7) = NulosN(TxtTotal4A.Text) * NulosN(Fg1.TextMatrix(Fg1.Row, 2))
        End If
    End If
    Agregando = False
    
    ActivarEntorno
    
    
    
    TotalizarFG1
    Frame12.Visible = False
End Sub

Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstMov
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstMov.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear

End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstMov("id")), xCon
    End If
End Sub

Private Sub Fg1_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    mVentanaActiva = 0 '--ORIGEN
    
    '--CARGANDO VALORES TEMPORALES
    IdConcepto = NulosN(Fg1.TextMatrix(Fg1.Row, 3))
    IdModulo = NulosN(Fg1.TextMatrix(Fg1.Row, 5))
    TipCam = NulosN(Fg1.TextMatrix(Fg1.Row, 2))
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "No ha especificado la moneda para realizar la operación", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If
    
    If NulosN(lblTipCambio.Caption) = 0 Then
        MsgBox "No ha especificado el tipo de cambio, ingrese la fecha de movimiento de operación para mostrar el tipo de cambio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchMov.SetFocus
        Exit Sub
    End If
    
    Agregando = True
    
    If Col = 1 Then
        If QueHace = 3 Then Exit Sub
        
        Dim xForm As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
        Dim xCampos(4, 4) As String
        Dim nSQLId As String
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        '--tipo de busqueda
        If OptDe1.Value = True Then
            xCampos(0, 0) = "Concepto":         xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "3450":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Nombre Cuenta":    xCampos(1, 1) = "descuen":       xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Nº Cuenta":        xCampos(2, 1) = "cuenta":        xCampos(2, 2) = "1100":         xCampos(2, 3) = "C"
            xForm.Ordenado = "descripcion"
            xForm.CampoBusca = "descripcion"
        Else
            xCampos(0, 0) = "Nº Cuenta":        xCampos(0, 1) = "cuenta":        xCampos(0, 2) = "1100":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Nombre Cuenta":    xCampos(1, 1) = "descuen":       xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Concepto":         xCampos(2, 1) = "descripcion":   xCampos(2, 2) = "3450":         xCampos(2, 3) = "C"
            xForm.Ordenado = "cuenta"
            xForm.CampoBusca = "cuenta"
        End If
        xCampos(3, 0) = "Módulo":  xCampos(3, 1) = "modulo":   xCampos(3, 2) = "650":         xCampos(3, 3) = "C"
        
'        xForm.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion AS descuen, tes_origen.id, tes_origen.idmon, tes_origen.descripcion," _
            & " tes_origen.idcuen, tes_origen.tipmov, tes_origen.entgen, tes_origen.idmod, (SELECT Count([iddoc]) AS numdocs From tes_origendoc " _
            & " WHERE (((tes_origendoc.id)=tes_origen.id))) AS numdocasi, tes_origen.activo, mae_banconumcta.id AS idbcocta " _
            & " FROM (tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) LEFT JOIN mae_banconumcta " _
            & " ON con_planctas.id = mae_banconumcta.idcuen WHERE (((tes_origen.idmon)=" & NulosN(TxtIdMon.Text) & ") AND ((tes_origen.tipmov)=2) AND ((tes_origen.activo)=-1))"
        
        
        '--sentencia sql para no mostrar documentos seleccionados
        nSQLId = GRID_GENERAR_SQL_ID(Fg1, 3, " AND tes_origen.id", " NOT IN", True)

        xForm.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion AS descuen, tes_origen.id, tes_origen.idmon, tes_origen.descripcion, tes_origen.idcuen, tes_origen.tipmov, " _
        & " tes_origen.entgen, tes_origen.idmod, 0 AS numdocasi, tes_origen.activo, mae_banconumcta.id AS idbcocta, tes_modulos.descripcion AS modulo " _
        & " FROM ((tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) LEFT JOIN mae_banconumcta ON con_planctas.id = mae_banconumcta.idcuen) LEFT JOIN tes_modulos ON tes_origen.idmod = tes_modulos.id " _
        & " WHERE (((tes_origen.idmon)=" & NulosN(TxtIdMon.Text) & ") AND ((tes_origen.tipmov)=2) AND ((tes_origen.activo)=-1)) " & nSQLId
        
        xForm.Titulo = "Buscando Origen del Egreso"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        Dim A As Integer
        Agregando = True
        If xRs.State = 1 Then
            For A = 1 To Fg1.Rows - 1
                If Fg1.TextMatrix(A, 3) = xRs("id") Then
                    MsgBox "El concepto seleccionado ya fue agregado, seleccione otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Set xRs = Nothing
                    Agregando = False
                    Exit Sub
                End If
            Next A
        
            Fg1.TextMatrix(Row, 1) = NulosC(xRs("descripcion"))
            Fg1.TextMatrix(Row, 3) = xRs("id")
            
            Fg1.TextMatrix(Row, 4) = xRs("idcuen")
            Fg1.TextMatrix(Row, 5) = NulosN(xRs("idmod"))
            Fg1.TextMatrix(Row, 6) = NulosN(xRs("numdocasi"))   'especifica el numero de documentos asignado al destino
            
            If NulosN(xRs("entgen")) = 5 Then '--proveedores
                CmdBusCliente.Enabled = True
            Else
                CmdBusCliente.Enabled = False
                TxtProv.Text = ""
                LblIdCliente.Caption = ""
            End If
            
            Fg1.TextMatrix(Row, 9) = NulosN(xRs("idbcocta"))
            
        End If
        Set xForm = Nothing
        Set xRs = Nothing
        
        If NulosN(Fg1.TextMatrix(Row, 2)) = 0 Then
            Fg1.TextMatrix(Row, 2) = NulosN(lblTipCambio.Caption)
        End If
        
        Agregando = False
    End If
    
    If Col = 7 Or Col = 8 Then
        If NulosN(Fg1.TextMatrix(Fg1.Row, 3)) < 0 Then
            MsgBox "Seleccione  el concepto del destino", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Agregando = False
            Exit Sub
        End If
        If NulosN(Fg1.TextMatrix(Fg1.Row, 3)) = 0 Then
            MsgBox "No ha especificado el concepto del destino", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Agregando = False
            Exit Sub
        End If
        
        '***************************************************************************************
        If IdModulo = 1 Or IdModulo = 8 Or IdModulo = 9 Or IdModulo = 10 Or _
           IdModulo = 2 Or IdModulo = 4 Or IdModulo = 11 Or IdModulo = 19 Then
           
            '--1=Compras
            '--8=Boleta de Pago
            '--9=Honorarios
            '--10=Reembolsables
            '--2=Ventas
            '--4=Letras
            '--11=Liquidacion de Gasto Debito
            '--19=Planilla de Letras
            '--'--'--'--'--'--'--'--'--'--'--'--'--'--'--'--'--'--
            Frame3.Left = 150
            Frame3.Top = 2400
            Frame3.Visible = True
    
            LblTitulo.Caption = Fg1.TextMatrix(Row, 1)
            LblTc.Caption = "T.C. " & Fg1.TextMatrix(Row, 2)
            
            CmdAgregar.Enabled = True:         CmdEliminar.Enabled = True
            
            Fg3.Rows = 1
            TxtProv.Text = ""
            LblIdCliente.Caption = ""
            TxtTotal1.Text = "": TxtTotal2.Text = "": TxtTotal3.Text = "": TxtTotal4.Text = "": TxtTotal5.Text = ""
            ActivarEntorno
            Fg3.Editable = flexEDKbdMouse
            '--'--'--'--'--'--'--'--'--'--'--'--'--'--'--'--'--'--
       
            CmdBusCliente.Enabled = True: TxtProv.Enabled = True
            
            Fg3.ColWidth(1) = 2160: Fg3.ColWidth(7) = 960: Fg3.ColWidth(8) = 960: Fg3.ColWidth(9) = 960: Fg3.ColWidth(10) = 960
            TxtTotal1.Visible = True: TxtTotal2.Visible = True: TxtTotal3.Visible = True: TxtTotal4.Visible = True: TxtTotal5.Visible = True
            
            Select Case IdModulo
                Case 1, 9, 10
                    Label3(2).Caption = "Proveedor"
                    Fg3.TextMatrix(0, 1) = "Proveedor"
                Case 2, 4, 11
                    Label3(2).Caption = "Cliente"
                    Fg3.TextMatrix(0, 1) = "Cliente"
                Case 8
                    Label3(2).Caption = "Personal"
                    Fg3.TextMatrix(0, 1) = "Personal"
                Case 19
                    Label3(2).Caption = "Ent. Bancaria"
                    Fg3.TextMatrix(0, 1) = "Entidad Bancaria"
            End Select
            
            Label2.Visible = True
            Frame7.Enabled = True
            
            TxtProv.Enabled = True
            Command6.Enabled = True
            
            CargaRstTmp IdConcepto
            
            If CmdBusCliente.Visible = True Then CmdBusCliente.SetFocus
            TotalizarFG3
            Exit Sub
        End If
        
        '***************************************************************************************
        ActivarEntorno
        
        Frame12.Left = 150
        Frame12.Top = 2400
        LblTitulo.Caption = Fg1.TextMatrix(Fg1.Row, 1)
        Fg6.Rows = 1
        TxtProvA.Text = ""
        LblIdClienteA.Caption = ""
        
        Frame12.Visible = True
        
        If NulosN(Fg1.TextMatrix(Fg1.Row, 5)) = 6 Then ' bancos
            Fg6.TextMatrix(0, 1) = "Medio Pago"
            Label10.Caption = Fg1.TextMatrix(Fg1.Row, 1) '--titulo
            Fg6.ColWidth(1) = 5000
            Command6.Enabled = False: TxtProvA.Enabled = False
            Fg6.ColWidth(3) = 0: Fg6.ColWidth(7) = 0: Fg6.ColWidth(8) = 0: Fg6.ColWidth(9) = 0: Fg6.ColWidth(10) = 0
            TxtTotal1A.Visible = False: TxtTotal2A.Visible = False: TxtTotal3A.Visible = False: TxtTotal4A.Visible = False: TxtTotal5A.Visible = False
            Label11.Visible = False
            Frame13.Enabled = False
                
            Fg6.Rows = 1
            
            RstTmpDocOri.Filter = adFilterNone
            RstTmpDocOri.Filter = "idconc = " & NulosN(Fg1.TextMatrix(Fg1.Row, 3)) & ""
            
            If RstTmpDocOri.RecordCount = 0 Then
                RstTmpDocOri.AddNew
                RstTmpDocOri("idconc") = NulosN(Fg1.TextMatrix(Fg1.Row, 3))
            End If
            
            RstTmpDocOri("fchemi") = NulosC(TxtFchMov.Valor)
            RstTmpDocOri("moneda") = Busca_Codigo(NulosN(TxtIdMon.Text), "id", "simbolo", "mae_moneda", "N", xCon)
            RstTmpDocOri("idmone") = NulosN(TxtIdMon.Text)
            
            CargaRstTmpOri Fg1.TextMatrix(Fg1.Row, 3)
            
            TotalizarFG6
            
            
            Fg6.Select Fg6.Rows - 1, 1, Fg6.Rows - 1, 1
            Fg6.SetFocus
        End If
        
        If NulosN(Fg1.TextMatrix(Fg1.Row, 5)) = 5 Then ' anticipos a proveedores
            Command6.Enabled = False: TxtProvA.Enabled = False
            Label11.Visible = False
            Frame13.Enabled = False
            
            Fg6.ColWidth(3) = 0: Fg6.ColWidth(7) = 0: Fg6.ColWidth(8) = 0: Fg6.ColWidth(9) = 0: Fg6.ColWidth(10) = 0
            TxtTotal1A.Visible = False: TxtTotal2A.Visible = False: TxtTotal3A.Visible = False: TxtTotal4A.Visible = False: TxtTotal5A.Visible = False
            
            
            Frame13.Enabled = True
            Fg6.TextMatrix(0, 1) = "Proveedor"
            
            
            RstTmpDocOri.Filter = adFilterNone
            RstTmpDocOri.Filter = "idconc = " & NulosN(Fg1.TextMatrix(Fg1.Row, 3)) & ""
            
            If RstTmpDocOri.RecordCount = 0 Then
                RstTmpDocOri.AddNew
                RstTmpDocOri("idconc") = NulosN(Fg1.TextMatrix(Fg1.Row, 3))
            End If
            
            CargaRstTmpOri Fg1.TextMatrix(Fg1.Row, 3)
            TotalizarFG6
        End If
        
    End If
    Agregando = False
    
End Sub

Private Sub Fg1_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    
    If Col = 2 Then
        If NulosN(TxtIdMon.Text) = 1 Then
            Fg1_CellChanged Row, 7
        Else
            Fg1_CellChanged Row, 8
        End If
        Exit Sub
    End If
    
    If Col = 7 Or Col = 8 Then
        If IsNumeric(Fg1.TextMatrix(Fg1.Row, Col)) = False Then
            Fg1.TextMatrix(Fg1.Row, Col) = ""
            Exit Sub
        End If
    End If
    
    If Col = 7 Then
        Fg1.TextMatrix(Row, 7) = Format(Fg1.TextMatrix(Row, 7), FORMAT_MONTO)
        If NulosN(Fg1.TextMatrix(Row, 2)) <> 0 Then
            Fg1.TextMatrix(Row, 8) = NulosN(Fg1.TextMatrix(Row, 7)) / NulosN(Fg1.TextMatrix(Row, 2))
        Else
            Fg1.TextMatrix(Row, 8) = 0
        End If
        Fg1.TextMatrix(Row, 8) = Format(Fg1.TextMatrix(Row, 8), "0.0000")
    End If
    If Col = 8 Then
        Fg1.TextMatrix(Row, 8) = Format(Fg1.TextMatrix(Row, 8), FORMAT_MONTO)
        Fg1.TextMatrix(Row, 7) = NulosN(Fg1.TextMatrix(Row, 8)) * NulosN(Fg1.TextMatrix(Row, 2))
        Fg1.TextMatrix(Row, 7) = Format(Fg1.TextMatrix(Row, 7), FORMAT_MONTO)
    End If
    TotalizarFG1
End Sub

Private Sub Fg1_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Fg1_CellButtonClick Row, Col

End Sub

Private Sub Fg1_EnterCell()
    If QueHace = 3 Then
        If Fg1.Col = 1 Then
            Fg1.Editable = flexEDNone: Exit Sub
        End If
    End If

    If TxtIdMon.Text = "1" Then
        If Fg1.Col = 7 Or Fg1.Col = 1 Or Fg1.Col = 2 Then
            Fg1.Editable = flexEDKbdMouse
        Else
            Fg1.Editable = flexEDNone
        End If
    Else
        If Fg1.Col = 8 Or Fg1.Col = 1 Or Fg1.Col = 2 Then
            Fg1.Editable = flexEDKbdMouse
        Else
            Fg1.Editable = flexEDNone
        End If
    End If
    
    If DetallarModulo(NulosN(Fg1.TextMatrix(Fg1.Row, 3)), origen, xCon) = True Then
        Fg1.ColComboList(7) = "|..."
        Fg1.ColComboList(8) = "|..."
    Else
        Fg1.ColComboList(7) = ""
        Fg1.ColComboList(8) = ""
    End If
End Sub

Private Sub Fg1_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If Col = 1 Then
            KeyAscii = 0
        ElseIf Col = 2 Then
        Else
            If DetallarModulo(NulosN(Fg1.TextMatrix(Row, 3)), origen, xCon) = True Then
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    
    If KeyAscii = 13 Then
        
    Else
        If Col = 2 Or Col = 6 Or Col = 7 Then
            If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        Command3_Click
    End If
    
    If KeyCode = 46 Then
        Command4_Click
    End If
    
    If KeyCode = 121 Then
        If Fg1.Row < Fg1.FixedRows Or Fg1.Col < Fg1.FixedCols Then Exit Sub
        Fg1_CellButtonClick Fg1.Row, Fg1.Col
    
    End If
    
End Sub

Private Sub Fg2_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    mVentanaActiva = 1 '--DESTINO

    IdConcepto = NulosN(Fg2.TextMatrix(Fg2.Row, 3))
    IdModulo = NulosN(Fg2.TextMatrix(Fg2.Row, 5))
    TipCam = NulosN(Fg2.TextMatrix(Fg2.Row, 2))
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "No ha especificado la moneda para realizar la operación", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If
    If NulosN(lblTipCambio.Caption) = 0 Then
        MsgBox "No ha especificado el tipo de cambio, ingrese la fecha de movimiento de operación para mostrar el tipo de cambio", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchMov.SetFocus
        Exit Sub
    End If
    
    If Col = 1 Then
        Dim xForm As New eps_librerias.FormBuscar
        Dim xRs As New ADODB.Recordset
        Dim xCampos(4, 4) As String
        Dim A As Integer
        Dim nSQLId As String

        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        '--ordernar por
        If OptDe1.Value = True Then
            xCampos(0, 0) = "Concepto":         xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "3450":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Nombre Cuenta":    xCampos(1, 1) = "descuen":       xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Nº Cuenta":        xCampos(2, 1) = "cuenta":        xCampos(2, 2) = "1100":         xCampos(2, 3) = "C"
            xForm.Ordenado = "descripcion"
            xForm.CampoBusca = "descripcion"
        Else
            xCampos(0, 0) = "Nº Cuenta":        xCampos(0, 1) = "cuenta":        xCampos(0, 2) = "1100":         xCampos(0, 3) = "C"
            xCampos(1, 0) = "Nombre Cuenta":    xCampos(1, 1) = "descuen":       xCampos(1, 2) = "3000":         xCampos(1, 3) = "C"
            xCampos(2, 0) = "Concepto":         xCampos(2, 1) = "descripcion":   xCampos(2, 2) = "3450":         xCampos(2, 3) = "C"
            xForm.Ordenado = "cuenta"
            xForm.CampoBusca = "cuenta"
        End If
        xCampos(3, 0) = "Módulo":  xCampos(3, 1) = "modulo":   xCampos(3, 2) = "650":         xCampos(3, 3) = "C"
        
'        xForm.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion AS descuen, tes_destino.id, tes_destino.idmon, tes_destino.descripcion, " _
            & " tes_destino.idcuen, tes_destino.tipmov, tes_destino.idmod, (SELECT Count([iddoc]) AS numdocs From tes_destinodoc " _
            & " WHERE (((tes_destinodoc.id)=tes_destino.id))) AS numdocasi, tes_destino.idbcocta FROM tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id " _
            & "WHERE (((tes_destino.idmon)=" & NulosN(TxtIdMon.Text) & ") AND ((tes_destino.tipmov)=2) AND ((tes_destino.activo)=-1))"
        
        '--sentencia sql para no mostrar documentos seleccionados
        nSQLId = GRID_GENERAR_SQL_ID(Fg1, 3, " AND tes_destino.id", " NOT IN", True)

        xForm.SQLCad = "SELECT con_planctas.cuenta, con_planctas.descripcion AS descuen, tes_destino.id, tes_destino.idmon, tes_destino.descripcion, tes_destino.idcuen, tes_destino.tipmov, " _
        & " tes_destino.entgen, tes_destino.idmod, 0 AS numdocasi, tes_destino.activo, mae_banconumcta.id AS idbcocta, tes_modulos.descripcion AS modulo " _
        & " FROM ((tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) LEFT JOIN mae_banconumcta ON con_planctas.id = mae_banconumcta.idcuen) LEFT JOIN tes_modulos ON tes_destino.idmod = tes_modulos.id " _
        & " WHERE (((tes_destino.idmon)=" & NulosN(TxtIdMon.Text) & ") AND ((tes_destino.tipmov)=2) AND ((tes_destino.activo)=-1)) " & nSQLId

        xForm.Titulo = "Buscando Destino del Egreso"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        Agregando = True
        If xRs.State = 1 Then
            'buscamos si el concepto ya fue agregado
            For A = 1 To Fg2.Rows - 1
                If Fg2.TextMatrix(A, 3) = xRs("id") Then
                    MsgBox "El concepto seleccionado ya fue agregado, seleccione otro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Agregando = False
                    Set xRs = Nothing
                    Exit Sub
                End If
            Next A
            
            Fg2.TextMatrix(Row, 1) = xRs("descripcion")
            
            Fg2.TextMatrix(Row, 3) = xRs("id")
            Fg2.TextMatrix(Row, 4) = xRs("idcuen")
            Fg2.TextMatrix(Row, 5) = NulosN(xRs("idmod"))
            Fg2.TextMatrix(Row, 6) = NulosN(xRs("numdocasi"))   'especifica el numero de documentos asignado al destino
            
            Fg2.TextMatrix(Row, 9) = NulosN(xRs("idbcocta"))
            
        End If
        Set xForm = Nothing
        Set xRs = Nothing
        
        If NulosN(Fg2.TextMatrix(Row, 2)) = 0 Then
            Fg2.TextMatrix(Row, 2) = NulosN(lblTipCambio.Caption)
        End If
        
        Agregando = False
        
    End If
    
    
    If Col = 7 Or Col = 8 Then
        If IdConcepto = 0 Then
            MsgBox "No ha especificado el concepto del destino", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
            Exit Sub
        End If
        Frame3.Left = 150
        Frame3.Top = 2400
        LblTitulo.Caption = Fg2.TextMatrix(Row, 1)
        LblTc.Caption = "T.C. " & Fg2.TextMatrix(Row, 2)
        
        CmdAgregar.Enabled = True:         CmdEliminar.Enabled = True
        
        Fg3.Rows = 1
        TxtProv.Text = ""
        LblIdCliente.Caption = ""
        TxtTotal1.Text = "": TxtTotal2.Text = "": TxtTotal3.Text = "": TxtTotal4.Text = "": TxtTotal5.Text = ""
        ActivarEntorno
        Fg3.Editable = flexEDKbdMouse
        
        Frame3.Visible = True

        If IdModulo = 7 Then   ' fondo fijo
            CmdBusCliente.Enabled = False: TxtProv.Enabled = False
            Label2.Visible = False
            Frame7.Enabled = False
            
            CmdAgregar.Enabled = False:         CmdEliminar.Enabled = False
            TxtTotal1.Visible = False: TxtTotal2.Visible = False: TxtTotal3.Visible = False: TxtTotal4.Visible = False: TxtTotal5.Visible = False
            Fg3.ColWidth(7) = 0: Fg3.ColWidth(8) = 0: Fg3.ColWidth(9) = 0: Fg3.ColWidth(10) = 0
            Fg3.TextMatrix(0, 1) = "Empleado"
            'agregamos el detalle del concepto
            RstTMPDoc.Filter = adFilterNone
            
            RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & NulosN(Fg2.TextMatrix(Row, 3)) & ""
            
            If RstTMPDoc.RecordCount = 0 Then
                RstTMPDoc.AddNew
            End If
            
            RstTMPDoc("idconc") = NulosN(Fg2.TextMatrix(Row, 3))
            CargaRstTmp Fg2.TextMatrix(Row, 3)
            Fg3.Select Fg3.Rows - 1, 1, Fg3.Rows - 1, 1
            Fg3.SetFocus
        End If
        
        If IdModulo = 3 Then  ' Entregas a rendir
            CmdBusCliente.Enabled = False: TxtProv.Enabled = False
            Label2.Visible = False
            Frame7.Enabled = False
            
            Fg3.ColWidth(7) = 0: Fg3.ColWidth(8) = 0: Fg3.ColWidth(9) = 0: Fg3.ColWidth(10) = 0
            TxtTotal1.Visible = True: TxtTotal2.Visible = False: TxtTotal3.Visible = False: TxtTotal4.Visible = False: TxtTotal5.Visible = False
            Frame7.Enabled = True
            Fg3.TextMatrix(0, 1) = "Personal/Empleado"
            
            'agregamos el detalle del concepto
            RstTMPDoc.Filter = adFilterNone
            
            RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & NulosN(Fg2.TextMatrix(Row, 3)) & ""
            
            If RstTMPDoc.RecordCount = 0 Then
                RstTMPDoc.AddNew
                RstTMPDoc("idconc") = IdConcepto
                RstTMPDoc("iddocu") = 0
            End If
            
            CargaRstTmp Fg2.TextMatrix(Row, 3)
            Fg3.Select Fg3.Rows - 1, 1, Fg3.Rows - 1, 1
            Fg3.SetFocus
            
        End If
        
        If IdModulo = 5 Then  ' Anticipos a proveedores
            CmdBusCliente.Enabled = False: TxtProv.Enabled = False
            Label2.Visible = False
            Frame7.Enabled = False
            
            Fg3.ColWidth(7) = 0: Fg3.ColWidth(8) = 0: Fg3.ColWidth(9) = 0: Fg3.ColWidth(10) = 0
            TxtTotal1.Visible = True: TxtTotal2.Visible = False: TxtTotal3.Visible = False: TxtTotal4.Visible = False: TxtTotal5.Visible = False
            Frame7.Enabled = True
            
            Label3(2).Caption = "Proveedor"
            Fg3.TextMatrix(0, 1) = "Proveedor"
            
            'agregamos el detalle del concepto
            RstTMPDoc.Filter = adFilterNone
            
            RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & NulosN(Fg2.TextMatrix(Row, 3)) & ""
            
            If RstTMPDoc.RecordCount = 0 Then
                RstTMPDoc.AddNew
                RstTMPDoc("idconc") = IdConcepto
                RstTMPDoc("iddocu") = 0
            End If
            CargaRstTmp IdConcepto
            If CmdAgregar.Visible = True Then CmdAgregar.SetFocus
        End If
    
        '***************************************************************************************
        If IdModulo = 1 Or IdModulo = 8 Or IdModulo = 9 Or IdModulo = 10 Or _
           IdModulo = 2 Or IdModulo = 4 Or IdModulo = 11 Or IdModulo = 19 Or _
           IdModulo = 20 Then
           
            '--1=Compras
            '--8=Boleta de Pago
            '--9=Honorarios
            '--10=Reembolsables
            
            '--En caso de devoluciones por cobranza en exceso o pago de nc, lgc
            '--2=Ventas
            '--4=Letras
            '--11=Liquidacion de Gasto Debito
            '--19=Planilla de Letras
            
            '--20=orden despacho
            '--'--'--'--'--'--'--'--'--'--'--'--'--'--'--'--'--'--
            CmdBusCliente.Enabled = True: TxtProv.Enabled = True
            
            Fg3.ColWidth(1) = 2160: Fg3.ColWidth(7) = 960: Fg3.ColWidth(8) = 960: Fg3.ColWidth(9) = 960: Fg3.ColWidth(10) = 960
            TxtTotal1.Visible = True: TxtTotal2.Visible = True: TxtTotal3.Visible = True: TxtTotal4.Visible = True: TxtTotal5.Visible = True
            
            Label2.Visible = True
            Frame7.Enabled = True
            
            TxtProv.Enabled = True
            Command6.Enabled = True
            CargaRstTmp IdConcepto
            '--'--'--'--'--'--'--'--'--'--'--'--'--'--'--'--'--'--
            
            Select Case IdModulo
                Case 1, 9, 10
                    Label3(2).Caption = "Proveedor"
                    Fg3.TextMatrix(0, 1) = "Proveedor"
                Case 2, 4, 11, 20
                    Label3(2).Caption = "Cliente"
                    Fg3.TextMatrix(0, 1) = "Cliente"
                Case 8
                    Label3(2).Caption = "Personal"
                    Fg3.TextMatrix(0, 1) = "Personal"
                Case 19
                    Label3(2).Caption = "Ent. Bancaria"
                    Fg3.TextMatrix(0, 1) = "Entidad Bancaria"
            End Select
            
            If CmdBusCliente.Visible = True Then CmdBusCliente.SetFocus
        End If
        
        '***************************************************************************************
        If IdModulo = 6 Then  ' bancos
            CmdBusCliente.Enabled = True: TxtProv.Enabled = True
            
            Fg3.ColWidth(1) = 2160: Fg3.ColWidth(7) = 960: Fg3.ColWidth(8) = 960: Fg3.ColWidth(9) = 960: Fg3.ColWidth(10) = 960
            TxtTotal1.Visible = True: TxtTotal2.Visible = True: TxtTotal3.Visible = True: TxtTotal4.Visible = True: TxtTotal5.Visible = True
            
            Label2.Visible = True
            Frame7.Enabled = True
            
            TxtProv.Enabled = True
            Command6.Enabled = True
            CargaRstTmp IdConcepto
            
            If CmdBusCliente.Visible = True Then CmdBusCliente.SetFocus
        End If
               
        TotalizarFG3
        
    End If
    
    '***************************************************************************************
End Sub

Sub CargaRstTmpOri(IdConcepto As Integer)
    RstTmpDocOri.Filter = adFilterNone
    RstTmpDocOri.Filter = "idconc = " & IdConcepto & ""
    Dim A As Integer
    
    If RstTmpDocOri.RecordCount <> 0 Then
        RstTmpDocOri.MoveFirst
        Agregando = True
        For A = 1 To RstTmpDocOri.RecordCount
            Fg6.Rows = Fg6.Rows + 1
            Fg6.TextMatrix(A, 1) = RstTmpDocOri("cliente")
            Fg6.TextMatrix(A, 2) = RstTmpDocOri("tipdoc")
            Fg6.TextMatrix(A, 3) = RstTmpDocOri("fchemi")
            Fg6.TextMatrix(A, 4) = RstTmpDocOri("moneda")
            Fg6.TextMatrix(A, 5) = RstTmpDocOri("numdoc")
            Fg6.TextMatrix(A, 6) = Format(RstTmpDocOri("imptot"), "0.00")
            Fg6.TextMatrix(A, 7) = Format(RstTmpDocOri("impsal"), "0.00")
            Fg6.TextMatrix(A, 8) = Format(RstTmpDocOri("impsal2"), "0.00")
            Fg6.TextMatrix(A, 9) = Format(RstTmpDocOri("acuent"), "0.00")
            Fg6.TextMatrix(A, 10) = Format(RstTmpDocOri("newsal"), "0.00")
            Fg6.TextMatrix(A, 11) = RstTmpDocOri("idconc")
            Fg6.TextMatrix(A, 12) = RstTmpDocOri("iddocu")
            Fg6.TextMatrix(A, 13) = RstTmpDocOri("idmone")
            Fg6.TextMatrix(A, 14) = RstTmpDocOri("idtipd")
            
            RstTmpDocOri.MoveNext
            If RstTmpDocOri.EOF = True Then Exit For
        Next A
        Agregando = False
    End If
End Sub

Sub CargaRstTmp(IdConcepto As Integer)
    RstTMPDoc.Filter = adFilterNone
    RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & IdConcepto & ""
    Dim A As Integer
    Fg3.Rows = 1
    If RstTMPDoc.RecordCount <> 0 Then
        RstTMPDoc.MoveFirst
        Agregando = True
        For A = 1 To RstTMPDoc.RecordCount
            Fg3.Rows = Fg3.Rows + 1
            Fg3.TextMatrix(A, 1) = NulosC(RstTMPDoc("cliente"))
            Fg3.TextMatrix(A, 2) = NulosC(RstTMPDoc("tipdoc"))
            Fg3.TextMatrix(A, 3) = NulosC(RstTMPDoc("fchemi"))
            Fg3.TextMatrix(A, 4) = NulosC(RstTMPDoc("moneda"))
            Fg3.TextMatrix(A, 5) = NulosC(RstTMPDoc("numdoc"))
            Fg3.TextMatrix(A, 6) = Format(RstTMPDoc("imptot"), FORMAT_MONTO)
            Fg3.TextMatrix(A, 7) = Format(RstTMPDoc("impsal"), FORMAT_MONTO)
            Fg3.TextMatrix(A, 8) = Format(RstTMPDoc("impsal2"), FORMAT_MONTO)
            Fg3.TextMatrix(A, 9) = Format(RstTMPDoc("acuent"), FORMAT_MONTO)
            Fg3.TextMatrix(A, 10) = Format(RstTMPDoc("newsal"), FORMAT_MONTO)
            Fg3.TextMatrix(A, 11) = RstTMPDoc("idconc")
            Fg3.TextMatrix(A, 12) = RstTMPDoc("iddocu")
            Fg3.TextMatrix(A, 13) = RstTMPDoc("idmone")
            Fg3.TextMatrix(A, 14) = RstTMPDoc("idtipd")
            
            Fg3.TextMatrix(A, 15) = RstTMPDoc("corr")
            
            '--calculo de porcentaje
            If NulosN(Fg3.TextMatrix(A, 8)) <> 0 And NulosN(Fg3.TextMatrix(A, 9)) Then
                Fg3.TextMatrix(A, 16) = Format(NulosN(Fg3.TextMatrix(A, 9) / Fg3.TextMatrix(A, 8)) * 100, "##.00")
            End If
            
            RstTMPDoc.MoveNext
            If RstTMPDoc.EOF = True Then Exit For
        Next A
        Agregando = False
    End If
End Sub

Sub ActivarEntorno()
    Toolbar1.Enabled = Not Toolbar1.Enabled
    TabOne1.Enabled = Not TabOne1.Enabled
End Sub

Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    
    If Col = 2 Then
        If NulosN(TxtIdMon.Text) = 1 Then
            Fg2_CellChanged Row, 7
        Else
            Fg2_CellChanged Row, 8
        End If
        Exit Sub
    End If
    
    If Col = 7 Or Col = 8 Then
        If IsNumeric(Fg2.TextMatrix(Row, Col)) = False Then
            Fg2.TextMatrix(Row, Col) = ""
            Exit Sub
        End If
    End If
    
    If Col = 7 Then
        Fg2.TextMatrix(Row, 7) = Format(Fg2.TextMatrix(Row, 7), FORMAT_MONTO)
        If NulosN(Fg2.TextMatrix(Row, 2)) <> 0 Then
            Fg2.TextMatrix(Row, 8) = NulosN(Fg2.TextMatrix(Row, 7)) / NulosN(Fg2.TextMatrix(Row, 2))
        Else
            Fg2.TextMatrix(Row, 8) = 0
        End If
        Fg2.TextMatrix(Row, 8) = Format(Fg2.TextMatrix(Row, 8), FORMAT_MONTO)
    End If
    
    If Col = 8 Then
        Fg2.TextMatrix(Row, 8) = Format(Fg2.TextMatrix(Row, 8), FORMAT_MONTO)
        Fg2.TextMatrix(Row, 7) = NulosN(Fg2.TextMatrix(Row, 8)) * NulosN(Fg2.TextMatrix(Row, 2))
        Fg2.TextMatrix(Row, 7) = Format(Fg2.TextMatrix(Row, 7), FORMAT_MONTO)
    End If
    
    TotalizarFG2
    
End Sub

Private Sub Fg2_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Fg2_CellButtonClick Row, Col
End Sub

Private Sub Fg2_EnterCell()

    If QueHace = 3 Then
        If Fg2.Col = 1 Then
            Fg2.Editable = flexEDNone: Exit Sub
        End If
    End If
    
    If TxtIdMon.Text = "1" Then
        If Fg2.Col = 7 Or Fg2.Col = 1 Or Fg2.Col = 2 Then
            Fg2.Editable = flexEDKbdMouse
        Else
            Fg2.Editable = flexEDNone
        End If
    Else
        If Fg2.Col = 8 Or Fg2.Col = 1 Or Fg2.Col = 2 Then
            Fg2.Editable = flexEDKbdMouse
        Else
            Fg2.Editable = flexEDNone
        End If
    End If
    
    If DetallarModulo(NulosN(Fg2.TextMatrix(Fg2.Row, 3)), destino, xCon) = True Then
        Fg2.ColComboList(7) = "|..."
        Fg2.ColComboList(8) = "|..."
    Else
        Fg2.ColComboList(7) = ""
        Fg2.ColComboList(8) = ""
    End If
    
End Sub

Private Sub Fg2_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii <> 13 Then
        If Col = 1 Then
            KeyAscii = 0
        ElseIf Col = 2 Then
        Else
            If DetallarModulo(NulosN(Fg2.TextMatrix(Row, 3)), destino, xCon) = True Then
                KeyAscii = 0
                Exit Sub
            End If
        End If
    End If
    
    'If Fg2.Col = 6 Or Fg2.Col = 1 Then
    '    If Fg2.TextMatrix(Fg2.Row, 5) <> 0 Then If KeyAscii <> 13 Then KeyAscii = 0
    'End If
    If KeyAscii = 13 Then
        
    Else
        If Col = 2 Or Col = 7 Or Col = 8 Then
            If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End If
End Sub



Private Sub Fg2_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        CmdAddCon_Click
    End If
    
    If KeyCode = 46 Then
        CmdDelCon_Click
    End If
    
    If KeyCode = 121 Then
        If Fg2.Row < Fg2.FixedRows Or Fg2.Col < Fg2.FixedCols Then Exit Sub
        Fg2_CellButtonClick Fg2.Row, Fg2.Col
    End If
    
End Sub

Function GeneraNumDocumento(Serie As String, TipoDocumento As Integer) As String
    Dim rst As New ADODB.Recordset

    RST_Busq rst, "SELECT tes_cajadestinodet.tipdoc, tes_cajadestinodet.numser, tes_cajadestinodet.numdoc From tes_cajadestinodet " _
        & " Where (((tes_cajadestinodet.tipdoc) = " & TipoDocumento & ")) ORDER BY tes_cajadestinodet.numser, tes_cajadestinodet.numdoc", xCon

    If rst.RecordCount <> 0 Then
        rst.MoveLast
        GeneraNumDocumento = Format(Serie, "0000") + "-" + Format(rst("numdoc") + 1, "0000000000")
    Else
        GeneraNumDocumento = Format(Serie, "0000") + "-" + "0000000001"
    End If
    Set rst = Nothing
End Function

Private Sub Fg3_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xForm As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    
    If Col = 1 Then
        Dim xCampos2(3, 4) As String
        
        If IdModulo = 5 Then
            xCampos2(0, 0) = "Proveedor":      xCampos2(0, 1) = "nombre":     xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
            xCampos2(1, 0) = "Nº RUC":         xCampos2(1, 1) = "numruc":     xCampos2(1, 2) = "1200":         xCampos2(1, 3) = "C"
            xCampos2(2, 0) = "Codigo":         xCampos2(2, 1) = "id":         xCampos2(2, 2) = "1000":         xCampos2(2, 3) = "N"
        
            xForm.SQLCad = "SELECT mae_prov.nombre, mae_prov.numruc, mae_prov.id From mae_prov Where (((mae_prov.activo) = -1)) ORDER BY mae_prov.nombre"
            xForm.Titulo = "Proveedores"
            xForm.Ordenado = "nombre"
            xForm.CampoBusca = "nombre"
        
        ElseIf IdModulo = 6 Then '--bancos
            xCampos2(0, 0) = "Proveedor":      xCampos2(0, 1) = "nombre":     xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
            xCampos2(1, 0) = "Nº RUC":         xCampos2(1, 1) = "numruc":     xCampos2(1, 2) = "1200":         xCampos2(1, 3) = "C"
            xCampos2(2, 0) = "Codigo":         xCampos2(2, 1) = "id":         xCampos2(2, 2) = "1000":         xCampos2(2, 3) = "N"
        
            xForm.SQLCad = "SELECT mae_prov.nombre, mae_prov.numruc, mae_prov.id From mae_prov Where (((mae_prov.activo) = -1)) ORDER BY mae_prov.nombre"
            xForm.Titulo = "Proveedores"
            xForm.Ordenado = "nombre"
            xForm.CampoBusca = "nombre"
        
        Else
            xCampos2(0, 0) = "Empleado":      xCampos2(0, 1) = "apenom":     xCampos2(0, 2) = "4000":         xCampos2(0, 3) = "C"
            xCampos2(1, 0) = "Tipo Doc.":     xCampos2(1, 1) = "apenom":     xCampos2(1, 2) = "1200":         xCampos2(1, 3) = "C"
            xCampos2(2, 0) = "NºDocumento.":  xCampos2(2, 1) = "numdoc":     xCampos2(2, 2) = "1000":         xCampos2(2, 3) = "C"
        
            xForm.SQLCad = "SELECT UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom] AS apenom, mae_dociden.descripcion, " _
                & " pla_empleados.numdoc, tes_usuarios.id FROM (pla_empleados RIGHT JOIN tes_usuarios ON pla_empleados.id = tes_usuarios.idper) " _
                & " LEFT JOIN mae_dociden ON pla_empleados.idtipdoc = mae_dociden.id"

            xForm.Titulo = "Usuarios de Tesoreria"
            xForm.Ordenado = "apenom"
            xForm.CampoBusca = "apenom"
        
        End If
        
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos2)
        If xRs.State = 1 Then
            
            RstTMPDoc.Filter = adFilterNone
            RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & IdConcepto & " AND corr = " & NulosN(Fg3.TextMatrix(Fg3.Row, 15)) & ""
            
            'RstTMPDoc.Filter = "idconc = " & NulosN(Fg2.TextMatrix(Fg2.Row, 3)) & " AND iddocu = " & NulosN(Fg3.TextMatrix(Fg3.Row, 12)) & ""
            
            If RstTMPDoc.RecordCount = 0 Then
                RstTMPDoc.AddNew
            End If
            Agregando = True
            If IdModulo = 3 Or IdModulo = 7 Then
                Fg3.TextMatrix(Fg3.Row, 12) = NulosN(xRs("id"))
                Fg3.TextMatrix(Row, 1) = NulosC(xRs("apenom"))
                RstTMPDoc("cliente") = NulosC(xRs("apenom"))
                RstTMPDoc("iddocu") = NulosN(xRs("id"))           'id de la persona que se le esta asignando el fondo fijo
            Else
                Fg3.TextMatrix(Fg3.Row, 12) = NulosN(xRs("id"))
                Fg3.TextMatrix(Row, 1) = NulosC(xRs("nombre"))
                RstTMPDoc("cliente") = NulosC(xRs("nombre"))
                RstTMPDoc("iddocu") = NulosN(xRs("id"))           'id de la persona que se le esta asignando el fondo fijo
            End If
            Agregando = False
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    End If
    
    If Col = 2 Then
        ReDim xCampos(3, 4) As String
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion":     xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Abreviatura":     xCampos(1, 1) = "abrev":         xCampos(1, 2) = "1000":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Codigo":          xCampos(2, 1) = "id":            xCampos(2, 2) = "1000":         xCampos(2, 3) = "N"
        
        
        If IdModulo <> 6 Then
            xForm.SQLCad = "SELECT tes_documentos.id, tes_documentos.descripcion, tes_documentos.abrev From tes_documentos WHERE (((tes_documentos.tipo)=1))"
        Else
            xForm.SQLCad = "SELECT mae_documento.id, mae_documento.descripcion, mae_documento.abrev From mae_documento where mae_documento.id in ( select tes_destinodoc.iddoc from tes_destinodoc where tes_destinodoc.id = " & IdConcepto & ") "
        End If
        
        xForm.Titulo = "Documentos Asignados"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        
        If xRs.State = 1 Then
            Fg3.TextMatrix(Row, 2) = NulosC(xRs("abrev"))
            Fg3.TextMatrix(Row, 3) = NulosC(TxtFchMov.Valor)
            Fg3.TextMatrix(Row, 4) = Busca_Codigo(NulosN(TxtIdMon.Text), "id", "simbolo", "mae_moneda", "N", xCon)
            Fg3.TextMatrix(Row, 5) = GeneraNumDocumento(0, xRs("id"))
            
            RstTMPDoc.Filter = adFilterNone
            'RstTMPDoc.Filter = "idconc = " & NulosN(Fg2.TextMatrix(Fg2.Row, 3)) & ""
            'RstTMPDoc.Filter = "idconc = " & NulosN(Fg2.TextMatrix(Fg2.Row, 3)) & " AND iddocu = " & NulosN(Fg3.TextMatrix(Fg3.Row, 12)) & ""
            RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & IdConcepto & " AND corr= " & NulosN(Fg3.TextMatrix(Fg3.Row, 15)) & ""
            
            If RstTMPDoc.RecordCount = 0 Then
                RstTMPDoc.AddNew
            End If
            
            RstTMPDoc("tipdoc") = NulosC(xRs("abrev"))
            RstTMPDoc("fchemi") = NulosC(TxtFchMov.Valor)
            RstTMPDoc("moneda") = NulosC(Busca_Codigo(NulosN(TxtIdMon.Text), "id", "simbolo", "mae_moneda", "N", xCon))
            RstTMPDoc("numdoc") = Fg3.TextMatrix(Row, 5)
            RstTMPDoc("idtipd") = xRs("id")
            RstTMPDoc("idmone") = NulosN(TxtIdMon.Text)
            
        End If
        Set xForm = Nothing
        Set xRs = Nothing
    
        Dim rst As New ADODB.Recordset
        RST_Busq rst, "SELECT * FROM tes_destinodoc WHERE id = " & IdConcepto & "", xCon
        
    End If
    
    If Col = 4 Then '--moneda
        
        ReDim xCampos(3, 4) As String
        
        'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
        xCampos(0, 0) = "Descripcion": xCampos(0, 1) = "descripcion":   xCampos(0, 2) = "4000":         xCampos(0, 3) = "C"
        xCampos(1, 0) = "Simbolo":     xCampos(1, 1) = "simbolo":         xCampos(1, 2) = "800":         xCampos(1, 3) = "C"
        xCampos(2, 0) = "Id":          xCampos(2, 1) = "id":            xCampos(2, 2) = "450":         xCampos(2, 3) = "N"

        xForm.SQLCad = "SELECT mae_moneda.id, mae_moneda.descripcion, mae_moneda.simbolo From mae_moneda "
        
        xForm.Titulo = "Seleccionar Moneda"
        xForm.FormaBusca = Principio
        xForm.Criterio = ""
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos)
        
        If xRs.State = 1 Then
            Fg3.TextMatrix(Row, 4) = NulosC(xRs("simbolo"))
            
            RstTMPDoc.Filter = adFilterNone
            
            RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & IdConcepto & " AND corr = " & NulosN(Fg3.TextMatrix(Fg3.Row, 15)) & " "
            
            If RstTMPDoc.RecordCount = 0 Then
                RstTMPDoc.AddNew
                RstTMPDoc("corr") = mCorrelativo2
                Fg3.TextMatrix(Row, 15) = mCorrelativo2
                mCorrelativo2 = mCorrelativo2 + 1
            End If
                        
            RstTMPDoc("idmone") = NulosN(xRs("id"))
            
        End If
        Set xForm = Nothing
        Set xRs = Nothing
        
    End If
    
End Sub

Private Sub Fg3_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    
    If Col = 9 Or Col = 6 Or Col = 7 Or Col = 8 Then
        If IsNumeric(Fg3.TextMatrix(Row, Col)) = False Then
            Fg3.TextMatrix(Fg3.Row, Col) = ""
            Exit Sub
        End If
        
        If Col = 7 Then

            If NulosN(Fg3.TextMatrix(Row, 13)) <> NulosN(TxtIdMon.Text) Then
            
                If NulosN(TxtIdMon.Text) = 1 Then
                    Fg3.TextMatrix(Row, 8) = NulosN(Fg3.TextMatrix(Row, 7)) * TipCam
                Else
                    If TipCam = 0 Then
                        Fg3.TextMatrix(Row, 8) = 0
                    Else
                        Fg3.TextMatrix(Row, 8) = NulosN(Fg3.TextMatrix(Row, 7)) / TipCam
                    End If
                End If
                Fg3.TextMatrix(Row, 8) = Format(Fg3.TextMatrix(Row, 8), FORMAT_MONTO)
            
            Else
                Fg3.TextMatrix(Row, 8) = Format(NulosN(Fg3.TextMatrix(Row, 7)), FORMAT_MONTO)
            End If
            
            Fg3.TextMatrix(Row, 9) = Format(NulosN(Fg3.TextMatrix(Row, 8)), FORMAT_MONTO)
            
            Fg3.TextMatrix(Fg3.Row, 10) = NulosN(Fg3.TextMatrix(Fg3.Row, 8)) - NulosN(Fg3.TextMatrix(Fg3.Row, 9))
            Fg3.TextMatrix(Fg3.Row, 10) = Format(Fg3.TextMatrix(Fg3.Row, 10), FORMAT_MONTO)
            
        
            RstTMPDoc.Filter = adFilterNone
            
            If RstTMPDoc.RecordCount <> 0 Then
                RstTMPDoc.MoveFirst
                RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & IdConcepto & " AND iddocu = " & NulosN(Fg3.TextMatrix(Row, 12)) & ""
                
                If RstTMPDoc.RecordCount <> 0 Then
                
                    RstTMPDoc("impsal") = NulosN(Fg3.TextMatrix(Row, 7))
                    RstTMPDoc("impsal2") = NulosN(Fg3.TextMatrix(Row, 8))
                    RstTMPDoc("acuent") = NulosN(Fg3.TextMatrix(Row, 9))
                    RstTMPDoc("newsal") = NulosN(Fg3.TextMatrix(Fg3.Row, 10))
                
                End If
            End If
        ElseIf Col = 6 Then
        
            RstTMPDoc.Filter = adFilterNone
            
            If RstTMPDoc.RecordCount <> 0 Then
                RstTMPDoc.MoveFirst
                RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & IdConcepto & " AND iddocu = " & NulosN(Fg3.TextMatrix(Row, 12)) & ""
                
                If RstTMPDoc.RecordCount <> 0 Then
                
                    RstTMPDoc("imptot") = NulosN(Fg3.TextMatrix(Row, 6))
        
                End If
            End If
        
        End If
        
        
    End If
    
    If IdModulo = 7 Or IdModulo = 3 Or IdModulo = 5 Then
        If Col = 6 Or Col = 5 Then
            RstTMPDoc.Filter = adFilterNone
            RstTMPDoc.MoveFirst
            RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & IdConcepto & " AND iddocu = " & NulosN(Fg3.TextMatrix(Fg3.Row, 12)) & ""
            
            If RstTMPDoc.RecordCount <> 0 Then
                RstTMPDoc("imptot") = NulosN(Fg3.TextMatrix(Fg3.Row, 6))
                RstTMPDoc("numdoc") = NulosC(Fg3.TextMatrix(Fg3.Row, 5))
            End If
        End If
        
    End If
    
    If Col = 9 Then
    
        Fg3.TextMatrix(Fg3.Row, 9) = Format(Fg3.TextMatrix(Fg3.Row, 9), FORMAT_MONTO)
        Fg3.TextMatrix(Fg3.Row, 10) = NulosN(Fg3.TextMatrix(Fg3.Row, 8)) - NulosN(Fg3.TextMatrix(Fg3.Row, 9))
        Fg3.TextMatrix(Fg3.Row, 10) = Format(Fg3.TextMatrix(Fg3.Row, 10), FORMAT_MONTO)
    
        If NulosN(Fg3.TextMatrix(Fg3.Row, 10)) < 0 Then
            MsgBox "El Importe ingresado es superior al Saldo del Documento", vbExclamation, xTitulo
            Fg3.TextMatrix(Row, Col) = 0
            Fg3.TextMatrix(Row, 10) = 0
            Fg3.Row = Row
            Fg3.SetFocus
            Exit Sub
        End If
        'actualizamos el acuenta y el saldo en el recorset temporal
        RstTMPDoc.Filter = adFilterNone
        RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & IdConcepto & " AND iddocu = " & Fg3.TextMatrix(Fg3.Row, 12) & ""
        
        If RstTMPDoc.RecordCount <> 0 Then
            RstTMPDoc("acuent") = Fg3.TextMatrix(Fg3.Row, 9)
            RstTMPDoc("newsal") = Fg3.TextMatrix(Fg3.Row, 10)
        End If
        
        
        If NulosN(Fg3.TextMatrix(Row, 9)) <> 0 Then
            Fg3.TextMatrix(Row, 16) = Format(NulosN(Fg3.TextMatrix(Row, 9) / Fg3.TextMatrix(Row, 8)) * 100, "##.00")
        End If

    End If
    
    If Col = 10 Then
        '--nuevo saldo
        Fg3.TextMatrix(Fg3.Row, 10) = Format(Fg3.TextMatrix(Fg3.Row, 10), FORMAT_MONTO)
        Fg3.TextMatrix(Fg3.Row, 9) = NulosN(Fg3.TextMatrix(Fg3.Row, 8)) - NulosN(Fg3.TextMatrix(Fg3.Row, 10))
        Fg3.TextMatrix(Fg3.Row, 9) = Format(Fg3.TextMatrix(Fg3.Row, 9), FORMAT_MONTO)
    
        If NulosN(Fg3.TextMatrix(Fg3.Row, 9)) < 0 Then
            MsgBox "El Importe ingresado es superior al Saldo del Documento", vbExclamation, xTitulo
            Fg3.TextMatrix(Row, 9) = 0
            Fg3.TextMatrix(Row, 10) = 0
            Fg3.TextMatrix(Row, 16) = 0
            Fg3.Row = Row
            Fg3.SetFocus
            Exit Sub
        End If
        
        Fg3_CellChanged Row, 9

    End If
    
    If Col = 16 Then
        If IsNumeric(Fg3.TextMatrix(Row, 16)) = False Then
            MsgBox "El valor ingresado no es numérico", vbInformation, xTitulo
            Fg3.TextMatrix(Row, 16) = "0"
        ElseIf NulosN(Fg3.TextMatrix(Row, 16)) > 100 Then
            MsgBox "El % no debe superar a 100", vbInformation, xTitulo
            Fg3.TextMatrix(Row, 16) = "0"
        Else
            If NulosN(Fg3.TextMatrix(Row, 8)) = 0 Then
                MsgBox "No hay importe para hacer cálculo", vbInformation, xTitulo
            Else
                Fg3.TextMatrix(Row, 9) = Format(NulosN(Fg3.TextMatrix(Row, 8) * Fg3.TextMatrix(Row, 16)) / 100, FORMAT_MONTO)
                Fg3_CellChanged Row, 9
            End If
        End If
        
    End If
    
    TotalizarFG3
End Sub

Sub TotalizarFG1()
''    Dim A As Integer
''    Dim xTotalSol, xTotalDol As Double
''
''    xTotalSol = 0
''    For A = 1 To Fg1.Rows - 1
''        xTotalSol = NulosN(Fg1.TextMatrix(A, 7)) + xTotalSol
''        xTotalDol = NulosN(Fg1.TextMatrix(A, 8)) + xTotalDol
''    Next A
''    TxtImpHabSol.Text = Format(xTotalSol, FORMAT_MONTO)
''    TxtImpHabDol.Text = Format(xTotalDol, FORMAT_MONTO)
'
    TxtImpHabSol.Text = Format(GRID_SUMAR_COL(Fg1, 7), FORMAT_MONTO)
    TxtImpHabDol.Text = Format(GRID_SUMAR_COL(Fg1, 8), FORMAT_MONTO)
    
    
    TotalizarDif
End Sub

Sub TotalizarFG2()
''    Dim A As Integer
''    Dim xTotal As Double
''    Dim xTotal2 As Double
''    xTotal = 0
''    xTotal2 = 0
''
''    Agregando = True
''    For A = 1 To Fg2.Rows - 1
''        xTotal = NulosN(Fg2.TextMatrix(A, 7)) + xTotal
''        xTotal2 = NulosN(Fg2.TextMatrix(A, 8)) + xTotal2
''    Next A
''    Agregando = False
''    TxtImpDebSol.Text = Format(xTotal, FORMAT_MONTO)
''    TxtImpDebDol.Text = Format(xTotal2, FORMAT_MONTO)
''
    TxtImpDebSol.Text = Format(GRID_SUMAR_COL(Fg2, 7), FORMAT_MONTO)
    TxtImpDebDol.Text = Format(GRID_SUMAR_COL(Fg2, 8), FORMAT_MONTO)
    
    TotalizarDif
End Sub

Private Sub TotalizarDif()
    TxtImpDifSol = Format(NulosN(TxtImpDebSol.Text) - NulosN(TxtImpHabSol.Text), FORMAT_MONTO)
    TxtImpDifDol = Format(NulosN(TxtImpDebDol.Text) - NulosN(TxtImpHabDol.Text), FORMAT_MONTO)
    '--mostrando alertas
    If NulosN(TxtImpDifSol.Text) <> 0 Then
        TxtImpDifSol.BackColor = vbYellow
    Else
        TxtImpDifSol.BackColor = vbWhite
    End If
    If NulosN(TxtImpDifDol.Text) <> 0 Then
        TxtImpDifDol.BackColor = vbYellow
    Else
        TxtImpDifDol.BackColor = vbWhite
    End If
End Sub

Sub TotalizarFG3()
'    Dim A As Integer
'    Dim xTotal1, xTotal2, xTotal3, xTotal4, xTotal5 As Double
'
'    For A = 1 To Fg3.Rows - 1
'        xTotal1 = xTotal1 + NulosN(Fg3.TextMatrix(A, 6))
'        xTotal2 = xTotal2 + NulosN(Fg3.TextMatrix(A, 7))
'        xTotal3 = xTotal3 + NulosN(Fg3.TextMatrix(A, 8))
'        xTotal4 = xTotal4 + NulosN(Fg3.TextMatrix(A, 9))
'        xTotal5 = xTotal5 + NulosN(Fg3.TextMatrix(A, 10))
'    Next A
'
    TxtTotal1.Text = Format(GRID_SUMAR_COL(Fg3, 6), FORMAT_MONTO)
    TxtTotal2.Text = Format(GRID_SUMAR_COL(Fg3, 7), FORMAT_MONTO)
    TxtTotal3.Text = Format(GRID_SUMAR_COL(Fg3, 8), FORMAT_MONTO)
    TxtTotal4.Text = Format(GRID_SUMAR_COL(Fg3, 9), FORMAT_MONTO)
    TxtTotal5.Text = Format(GRID_SUMAR_COL(Fg3, 10), FORMAT_MONTO)
    
End Sub

Sub TotalizarFG6()
'    Dim A As Integer
'    Dim xTotal1, xTotal2, xTotal3, xTotal4, xTotal5 As Double
'    For A = 1 To Fg6.Rows - 1
'        xTotal1 = xTotal1 + NulosN(Fg6.TextMatrix(A, 6))
'        xTotal2 = xTotal2 + NulosN(Fg6.TextMatrix(A, 7))
'        xTotal3 = xTotal3 + NulosN(Fg6.TextMatrix(A, 8))
'        xTotal4 = xTotal4 + NulosN(Fg6.TextMatrix(A, 9))
'        xTotal5 = xTotal5 + NulosN(Fg6.TextMatrix(A, 10))
'    Next A
'
'    TxtTotal1A.Text = Format(xTotal1, FORMAT_MONTO)
'    TxtTotal2A.Text = Format(xTotal2, FORMAT_MONTO)
'    TxtTotal3A.Text = Format(xTotal3, FORMAT_MONTO)
'    TxtTotal4A.Text = Format(xTotal4, FORMAT_MONTO)
'    TxtTotal5A.Text = Format(xTotal5, FORMAT_MONTO)
    
    TxtTotal1A.Text = Format(GRID_SUMAR_COL(Fg6, 6), FORMAT_MONTO)
    TxtTotal2A.Text = Format(GRID_SUMAR_COL(Fg6, 7), FORMAT_MONTO)
    TxtTotal3A.Text = Format(GRID_SUMAR_COL(Fg6, 8), FORMAT_MONTO)
    TxtTotal4A.Text = Format(GRID_SUMAR_COL(Fg6, 9), FORMAT_MONTO)
    TxtTotal5A.Text = Format(GRID_SUMAR_COL(Fg6, 10), FORMAT_MONTO)
    
End Sub

Sub PreparaRST()
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(21, 3) As String

    xCampos(0, 0) = "cliente":        xCampos(0, 1) = "C":      xCampos(0, 2) = "100" ' cliente
    xCampos(1, 0) = "tipdoc":         xCampos(1, 1) = "C":      xCampos(1, 2) = "10" ' tipo de documento
    xCampos(2, 0) = "fchemi":         xCampos(2, 1) = "C":      xCampos(2, 2) = "10" ' fecha de emision
    xCampos(3, 0) = "moneda":         xCampos(3, 1) = "C":      xCampos(3, 2) = "30" ' moneda del documento
    xCampos(4, 0) = "numdoc":         xCampos(4, 1) = "C":      xCampos(4, 2) = "50" ' numero de documento
    xCampos(5, 0) = "imptot":         xCampos(5, 1) = "D":      xCampos(5, 2) = "2" ' importe total del documento
    xCampos(6, 0) = "impsal":         xCampos(6, 1) = "D":      xCampos(6, 2) = "2" ' saldo del documento
    xCampos(7, 0) = "impsal2":        xCampos(7, 1) = "D":      xCampos(7, 2) = "2" ' saldo del documento en la moneda de trabajo
    xCampos(8, 0) = "acuent":         xCampos(8, 1) = "D":      xCampos(8, 2) = "2" ' importe acuenta
    xCampos(9, 0) = "newsal":         xCampos(9, 1) = "D":      xCampos(9, 2) = "2" ' nuevo saldo del documento
    xCampos(10, 0) = "idconc":        xCampos(10, 1) = "N":     xCampos(10, 2) = "2" ' id del cocepto
    xCampos(11, 0) = "iddocu":        xCampos(11, 1) = "N":     xCampos(11, 2) = "2" ' id del documento
    xCampos(12, 0) = "idmone":        xCampos(12, 1) = "N":     xCampos(12, 2) = "2" ' id del al moneda del documento
    xCampos(13, 0) = "idtipd":        xCampos(13, 1) = "N":     xCampos(13, 2) = "2" ' id del tipo del documento
    xCampos(14, 0) = "idori":         xCampos(14, 1) = "N":     xCampos(14, 2) = "2" ' codigo del origen del documento
    
    xCampos(15, 0) = "corr":          xCampos(15, 1) = "N":     xCampos(15, 2) = "2" ' correlativo es unico
    xCampos(16, 0) = "glosa":         xCampos(16, 1) = "C":     xCampos(16, 2) = "240" ' glosa
    xCampos(17, 0) = "registro":      xCampos(17, 1) = "C":     xCampos(17, 2) = "10" ' registro
    xCampos(18, 0) = "idtipper":      xCampos(18, 1) = "N":     xCampos(18, 2) = "2" ' codigo del tipo de persona 1 proveedor, 2 cliente, 3 empleado 4 otros 5 banco
    xCampos(19, 0) = "idper":         xCampos(19, 1) = "N":     xCampos(19, 2) = "3" ' codigo de la entidad
    xCampos(20, 0) = "xtipo":         xCampos(20, 1) = "N":     xCampos(20, 2) = "1" ' indica la venta activa, origen o destino igual a varaible mVentanaActiva(0::ORIGEN; 1::DESTINO)
    
    
    Set RstTMPDoc = xFun.CrearRstTMP(xCampos)
    RstTMPDoc.Open
End Sub

Sub PreparaRSTOri()
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(16, 3) As String

    xCampos(0, 0) = "cliente":        xCampos(0, 1) = "C":      xCampos(0, 2) = "100" ' cliente
    xCampos(1, 0) = "tipdoc":         xCampos(1, 1) = "C":      xCampos(1, 2) = "10" ' tipo de documento
    xCampos(2, 0) = "fchemi":         xCampos(2, 1) = "C":      xCampos(2, 2) = "10" ' fecha de emision
    xCampos(3, 0) = "moneda":         xCampos(3, 1) = "C":      xCampos(3, 2) = "30" ' moneda del documento
    xCampos(4, 0) = "numdoc":         xCampos(4, 1) = "C":      xCampos(4, 2) = "50" ' numero de documento
    xCampos(5, 0) = "imptot":         xCampos(5, 1) = "D":      xCampos(5, 2) = "2" ' importe total del documento
    xCampos(6, 0) = "impsal":         xCampos(6, 1) = "D":      xCampos(6, 2) = "2" ' saldo del documento
    xCampos(7, 0) = "impsal2":        xCampos(7, 1) = "D":      xCampos(7, 2) = "2" ' saldo del documento en la moneda de trabajo
    xCampos(8, 0) = "acuent":         xCampos(8, 1) = "D":      xCampos(8, 2) = "2" ' importe acuenta
    xCampos(9, 0) = "newsal":         xCampos(9, 1) = "D":      xCampos(9, 2) = "2" ' nuevo saldo del documento
    xCampos(10, 0) = "idconc":         xCampos(10, 1) = "N":      xCampos(10, 2) = "2" ' id del cocepto
    xCampos(11, 0) = "iddocu":         xCampos(11, 1) = "N":      xCampos(11, 2) = "2" ' id del documento
    xCampos(12, 0) = "idmone":         xCampos(12, 1) = "N":      xCampos(12, 2) = "2" ' id del al moneda del documento
    xCampos(13, 0) = "idtipd":         xCampos(13, 1) = "N":      xCampos(13, 2) = "2" ' id del tipo del documento
    
    xCampos(14, 0) = "corr":          xCampos(14, 1) = "N":      xCampos(14, 2) = "2" ' correlativo es unico
    xCampos(15, 0) = "idbcocta":      xCampos(15, 1) = "N":      xCampos(15, 2) = "2" ' id del banco
    
    Set RstTmpDocOri = xFun.CrearRstTMP(xCampos)
    RstTmpDocOri.Open
End Sub

Function PreparaRST2() As ADODB.Recordset
    'Dim xFun As New Eps_DataAcces.FuncionesData
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(4, 3) As String

    xCampos(0, 0) = "idconc":        xCampos(0, 1) = "N":      xCampos(0, 2) = "2" ' codigo de la cuenta
    xCampos(1, 0) = "importe":       xCampos(1, 1) = "D":      xCampos(1, 2) = "2" ' importe de la cuenta
    xCampos(2, 0) = "cuenta":        xCampos(2, 1) = "C":      xCampos(2, 2) = "10" ' numero de la cuenta
    xCampos(3, 0) = "descta":        xCampos(3, 1) = "C":      xCampos(3, 2) = "100" ' descripcion de la cuenta
    
    Set PreparaRST2 = xFun.CrearRstTMP(xCampos)
    PreparaRST2.Open
End Function

Private Sub Fg3_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Fg3_CellButtonClick Row, Col

End Sub

Private Sub Fg3_EnterCell()
    If QueHace = 3 Then
        Fg3.SelectionMode = flexSelectionByRow
        Fg3.Editable = flexEDNone
    Else
        Fg3.SelectionMode = flexSelectionFree
        Fg3.Editable = flexEDKbdMouse
    End If
    
    If IdModulo = 7 Or IdModulo = 3 Or IdModulo = 5 Or IdModulo = 6 Then
        If Fg3.Col = 2 Then Fg3.Editable = flexEDKbdMouse
        Exit Sub
    End If
    
    If Fg3.Col = 10 Or Fg3.Col = 9 Or Fg3.Col = 7 Or Fg3.Col = 16 Then
        Fg3.Editable = flexEDKbdMouse
    Else
        Fg3.Editable = flexEDNone
    End If
    
End Sub

Private Sub Fg3_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        
    Else
        If Col = 9 Or Col = 6 Or Col = 7 Or Col = 8 Then
            If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End If

End Sub

Private Sub Fg3_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
        CmdAgregar_Click
    End If
    
    If KeyCode = 46 Then
        CmdEliminar_Click
    End If
    
    If KeyCode = 121 Then
        If Fg3.Row < Fg3.FixedRows Or Fg3.Col < Fg3.FixedCols Then Exit Sub
        Fg3_CellButtonClick Fg3.Row, Fg3.Col
    End If
    
End Sub



Private Sub Fg6_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim xRs As New ADODB.Recordset
    Dim xForm As New eps_librerias.FormBuscar
    Dim xCampos2(3, 4) As String
    
    If Col = 1 Then
    
        If Fg1.TextMatrix(Fg1.Row, 5) = 5 Then
            xCampos2(0, 0) = "Proveedor":      xCampos2(0, 1) = "nombre":     xCampos2(0, 2) = "5000":         xCampos2(0, 3) = "C"
            xCampos2(1, 0) = "Nº RUC":         xCampos2(1, 1) = "numruc":     xCampos2(1, 2) = "1200":         xCampos2(1, 3) = "C"
            xCampos2(2, 0) = "Codigo":         xCampos2(2, 1) = "id":         xCampos2(2, 2) = "1000":         xCampos2(2, 3) = "N"
        
            xForm.SQLCad = "SELECT mae_prov.nombre, mae_prov.numruc, mae_prov.id From mae_prov Where (((mae_prov.activo) = -1)) ORDER BY mae_prov.nombre"
            xForm.Titulo = "Proveedores"
            xForm.Ordenado = "nombre"
            xForm.CampoBusca = "nombre"
        End If
        
        If Fg1.TextMatrix(Fg1.Row, 5) = 6 Then
            xCampos2(0, 0) = "Documento":      xCampos2(0, 1) = "nombre":    xCampos2(0, 2) = "4000":         xCampos2(0, 3) = "C"
            xCampos2(1, 0) = "Cod. Sunat.":    xCampos2(1, 1) = "codsun":         xCampos2(1, 2) = "1200":         xCampos2(1, 3) = "C"
            xCampos2(2, 0) = "Codigo":         xCampos2(2, 1) = "id":             xCampos2(2, 2) = "1000":         xCampos2(2, 3) = "C"
            
            xForm.SQLCad = "SELECT tes_mediopago.id, tes_mediopago.descripcion as nombre, tes_mediopago.codsun From tes_mediopago ORDER BY tes_mediopago.descripcion"
            xForm.Titulo = "Busqueda de Medio de Pago"
            xForm.FormaBusca = Principio
            xForm.Criterio = ""
            xForm.Ordenado = "nombre"
            xForm.CampoBusca = "nombre"
        End If
        
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos2)
        
        If xRs.State = 1 Then
            RstTmpDocOri.Filter = adFilterNone
            RstTmpDocOri.Filter = "idconc = " & NulosN(Fg1.TextMatrix(Fg1.Row, 3)) ' & " AND iddocu = " & NulosN(Fg6.TextMatrix(Fg6.Row, 12)) & ""
            
            If RstTmpDocOri.RecordCount = 0 Then
                RstTmpDocOri.AddNew
            End If
            Agregando = True
            If Fg1.TextMatrix(Fg1.Row, 5) = 3 Or Fg1.TextMatrix(Fg1.Row, 5) = 7 Then
                Fg6.TextMatrix(Fg6.Row, 12) = NulosN(xRs("id"))
                Fg6.TextMatrix(Row, 1) = xRs("apenom")
                RstTmpDocOri("cliente") = NulosC(xRs("apenom"))
                RstTmpDocOri("iddocu") = NulosN(xRs("id"))           'id de la persona que se le esta asignando el fondo fijo
            Else
                Fg6.TextMatrix(Fg6.Row, 12) = NulosN(xRs("id"))
                Fg6.TextMatrix(Row, 1) = xRs("nombre")
                RstTmpDocOri("cliente") = NulosC(xRs("nombre"))
                RstTmpDocOri("iddocu") = NulosN(xRs("id"))
            End If
            Agregando = False
        
        End If
        
        Set xForm = Nothing
        Set xRs = Nothing
    End If
    
    If Col = 2 Then
        
        xCampos2(0, 0) = "Documento":      xCampos2(0, 1) = "descripcion":    xCampos2(0, 2) = "4000":         xCampos2(0, 3) = "C"
        xCampos2(1, 0) = "Abrev.":         xCampos2(1, 1) = "abrev":          xCampos2(1, 2) = "1200":         xCampos2(1, 3) = "C"
        xCampos2(2, 0) = "Codigo":         xCampos2(2, 1) = "id":             xCampos2(2, 2) = "1000":         xCampos2(2, 3) = "C"
        
        If Fg1.TextMatrix(Fg1.Row, 5) = 6 Then
            'si el modulo que genera es BANCO
            xForm.SQLCad = "SELECT tes_documentos.id, tes_documentos.descripcion, tes_documentos.abrev, tes_documentos.tipo From tes_documentos " _
                & " Where (((tes_documentos.Tipo) = 2)) ORDER BY tes_documentos.descripcion"
        Else
            xForm.SQLCad = "SELECT tes_documentos.id, tes_documentos.descripcion, tes_documentos.abrev, tes_documentos.tipo From tes_documentos " _
                & " Where (((tes_documentos.Tipo) = 1)) ORDER BY tes_documentos.descripcion"
        End If
        xForm.Titulo = "Busqueda de Tipo de Documento"
        xForm.FormaBusca = Principio
        xForm.Ordenado = "descripcion"
        xForm.CampoBusca = "descripcion"
        xForm.Criterio = ""
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos2)
        
        If xRs.State = 1 Then
            Fg6.TextMatrix(Row, 2) = NulosC(xRs("abrev"))
            Fg6.TextMatrix(Row, 14) = xRs("id")
            RstTmpDocOri.Filter = "idconc = " & NulosN(Fg1.TextMatrix(Fg1.Row, 3)) & ""
            
            If RstTmpDocOri.RecordCount = 0 Then
                RstTmpDocOri.AddNew
            End If
            
            RstTmpDocOri("tipdoc") = NulosC(xRs("abrev"))
            RstTmpDocOri("idtipd") = NulosC(xRs("id"))  'el iddocu tambien alamcenarar el codigo del medio de pago
        End If
        
        Set xForm = Nothing
        Set xRs = Nothing
    End If
    
    If Col = 5 Then
        Dim xCampos22(3, 4) As String
    
        xCampos22(0, 0) = "Proveedor":      xCampos22(0, 1) = "nombre":     xCampos22(0, 2) = "5000":         xCampos22(0, 3) = "C"
        xCampos22(1, 0) = "Nº RUC":         xCampos22(1, 1) = "numruc":     xCampos22(1, 2) = "1200":         xCampos22(1, 3) = "C"
        xCampos22(2, 0) = "Codigo":         xCampos22(2, 1) = "id":         xCampos22(2, 2) = "1000":         xCampos22(2, 3) = "N"
    
        xForm.SQLCad = "SELECT tes_cajadestinodet.idmod, tes_cajadestinodet.idper, tes_cajadestinodet.tipdoc, tes_documentos.abrev, " _
            & " [tes_cajadestinodet]![numser]+'-'+[tes_cajadestinodet]![numdoc] AS numdoc, tes_cajadestinodet.importe, tes_caja.fchope, mae_prov.nombre, " _
            & " tes_caja.idmon, mae_moneda.simbolo FROM (((tes_cajadestinodet LEFT JOIN tes_documentos ON tes_cajadestinodet.tipdoc = tes_documentos.id) " _
            & " LEFT JOIN tes_caja ON tes_cajadestinodet.idtes = tes_caja.id) LEFT JOIN mae_prov ON tes_cajadestinodet.idper = mae_prov.id) LEFT JOIN " _
            & " mae_moneda ON tes_caja.idmon = mae_moneda.id Where (((tes_cajadestinodet.idmod) = 5) And ((tes_cajadestinodet.idper) = 1646)) " _
            & " ORDER BY tes_caja.fchope"

        xForm.Titulo = "Anticipos Emitidos al Proveedor"
        xForm.Ordenado = "nombre"
        xForm.CampoBusca = "nombre"
    
        Set xForm.Coneccion = xCon
        Set xRs = xForm.BuscarReg(xCampos22)
        
        If xRs.State = 1 Then
            Fg6.TextMatrix(Fg6.Row, 2) = xRs("abrev")
            Fg6.TextMatrix(Fg6.Row, 3) = xRs("fchope")
            Fg6.TextMatrix(Fg6.Row, 4) = xRs("simbolo")
            Fg6.TextMatrix(Fg6.Row, 5) = xRs("numdoc")
            Fg6.TextMatrix(Fg6.Row, 6) = xRs("importe")
            'Fg6.TextMatrix(Fg6.Row, 7) = xRs("importe") aqui mostrara el saldo del documento
            Fg6.TextMatrix(Fg6.Row, 11) = xRs("") 'idconcep
            Fg6.TextMatrix(Fg6.Row, 12) = xRs("idper") 'iddocumento
            Fg6.TextMatrix(Fg6.Row, 13) = xRs("idmon") 'idmoneda
            Fg6.TextMatrix(Fg6.Row, 14) = xRs("tipdoc") 'idtipdoc
        End If
    End If
End Sub

Private Sub Fg6_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Agregando = True Then Exit Sub
    
    If Col = 9 Or Col = 6 Or Col = 7 Or Col = 8 Then
        If IsNumeric(Fg6.TextMatrix(Fg6.Row, Col)) = False Then
            Fg6.TextMatrix(Fg6.Row, Col) = ""
            Exit Sub
        End If
    End If
    
    If Col = 5 Or Col = 6 Then
    
        RstTmpDocOri.Filter = "idconc = " & NulosN(Fg1.TextMatrix(Fg1.Row, 3)) & ""
        
        If RstTmpDocOri.RecordCount = 0 Then
            RstTmpDocOri.AddNew
        End If
        
        RstTmpDocOri("numdoc") = Trim(Fg6.TextMatrix(Fg6.Row, 5))
        RstTmpDocOri("imptot") = Format(NulosN(Fg6.TextMatrix(Fg6.Row, 6)), "0.00")
    End If
    
    If Col = 9 Then
        Fg6.TextMatrix(Fg6.Row, 10) = NulosN(Fg6.TextMatrix(Fg6.Row, 8)) - NulosN(Fg3.TextMatrix(Fg6.Row, 9))
        Fg6.TextMatrix(Fg6.Row, 10) = Format(Fg6.TextMatrix(Fg6.Row, 10), FORMAT_MONTO)
    
        'actualizamos el acuenta y el saldo en el recorset temporal
        RstTmpDocOri.Filter = adFilterNone
        RstTmpDocOri.Filter = "idconc = " & Fg1.TextMatrix(Fg1.Row, 3) & " AND iddocu = " & Fg6.TextMatrix(Fg6.Row, 12) & ""
        
        If RstTmpDocOri.RecordCount <> 0 Then
            RstTmpDocOri("acuent") = Fg6.TextMatrix(Fg6.Row, 9)
            RstTmpDocOri("newsal") = Fg6.TextMatrix(Fg6.Row, 10)
        End If
    End If
    TotalizarFG6
End Sub

Private Sub Fg6_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Fg6_CellButtonClick Row, Col

End Sub

Private Sub Fg6_EnterCell()
    If QueHace = 3 Then
        Fg6.Editable = flexEDNone
        Fg6.SelectionMode = flexSelectionByRow
        Exit Sub
    Else
        Fg6.SelectionMode = flexSelectionFree
        Fg6.Editable = flexEDKbdMouse
    End If
    If NulosN(Fg1.TextMatrix(Fg1.Row, 5)) = 6 Then
        If Fg6.Col = 1 Or Fg6.Col = 5 Or Fg6.Col = 6 Then Fg6.Editable = flexEDKbdMouse
        Exit Sub
    End If
    
    If Fg6.Col = 9 Then
        Fg6.Editable = flexEDKbdMouse
    Else
        Fg6.Editable = flexEDNone
    End If
End Sub

Private Sub Fg6_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        
    Else
        If Col = 9 Or Col = 6 Or Col = 7 Or Col = 8 Then
            If InStr(CaracteresNumericos, Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub Fg6_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 121 Then
        If Fg6.Row < Fg6.FixedRows Or Fg6.Col < Fg6.FixedCols Then Exit Sub
        Fg6_CellButtonClick Fg6.Row, Fg6.Col
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
    
        Dim Rpta As Integer
        SeEjecuto = True
        
        mMesActivo = xMes
        
        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
        LblPeriodo.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
        
        OpcionesPeriodo
        
        CargarRSTCom
        
        Set Dg1.DataSource = RstMov

    End If
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        If fCierrePeriodo = True Then Exit Sub
        If QueHace <> 3 Then Exit Sub
        Nuevo
    End If

    If KeyCode = 115 Then
        If fCierrePeriodo = False Then Exit Sub
        If QueHace <> 3 Then Exit Sub
        Modificar
    End If

    If KeyCode = 113 Then
        If fCierrePeriodo = False Then Exit Sub
        If QueHace = 3 Then Exit Sub
        If Grabar = True Then
            Cancelar
            RstMov.Requery
            Dg1.Refresh
        End If
    End If
    
    'If KeyCode = 116 Then
    '    'Buscar
    'End If
End Sub

Private Sub Form_Load()
    TabOne1.CurrTab = 0
    QueHace = 3
    SeEjecuto = False
    
    Dg1.Columns("fchope").NumberFormat = FORMAT_DATE
    Dg1.Columns("importe").NumberFormat = FORMAT_MONTO
    
    lblTipCambio.Caption = ""
    Fg1.ColWidth(3) = 0
    Fg1.ColWidth(4) = 0
    Fg1.ColWidth(5) = 0
    Fg1.ColWidth(6) = 0
    Fg1.ColWidth(9) = 0 '--idcuenta de banco
   
    Fg2.ColWidth(3) = 0
    Fg2.ColWidth(4) = 0
    Fg2.ColWidth(5) = 0
    Fg2.ColWidth(6) = 0
    Fg2.ColWidth(9) = 0 '--idcuenta de banco
    
    Fg3.ColWidth(11) = 0
    Fg3.ColWidth(12) = 0
    Fg3.ColWidth(13) = 0
    Fg3.ColWidth(14) = 0
    
    Fg3.ColWidth(15) = 0 '--correlativo
    
    Fg6.ColWidth(11) = 0
    Fg6.ColWidth(12) = 0
    Fg6.ColWidth(13) = 0
    Fg6.ColWidth(14) = 0
    
    Fg4.ColWidth(10) = 0
    Fg4.ColWidth(11) = 0
    Fg4.ColWidth(12) = 0
    Fg4.ColWidth(13) = 0
    
    Frame2.BackColor = &H8000000F
    Frame1.BackColor = &H8000000F
    
    Fg1.ColComboList(1) = "|..."
    Fg1.ColComboList(7) = "|..."
    
     

    Fg2.ColComboList(1) = "|..."
    Fg2.ColComboList(7) = "|..."

    Fg3.ColComboList(1) = "|..."
    Fg3.ColComboList(2) = "|..."
    
    Fg6.ColComboList(1) = "|..."
    Fg6.ColComboList(2) = "|..."
    
    Fg2.SelectionMode = flexSelectionByRow
    Fg1.SelectionMode = flexSelectionByRow
    
    Fg4.SelectionMode = flexSelectionByRow
    
    CaracteresNumericos = "0123456789." & Chr(8)
    
'    Dim xArchivos(16) As String
'
'    Dim A As Integer
'    For A = 1 To 17
'        xArchivos(A) = "D:\Proyectos\contabilidad\bmps\toolbar\" & Trim(Str(A)) & ".ico"
'    Next A
'    'Call AgregarToImageList(App.Path & "\iconos-32-bits\" & I & ".ico", Ancho, Alto, ImageList1)
'
'    CargarToolbar ImageList2, Toolbar1, Me, 20, xArchivos, T_ToolTipText

    '---------------
    Me.Width = 12000
    Me.Height = 8000
    '---------------
    
End Sub

Private Sub OptBanco_Click()
    If OptBanco.Value = True Then
        Fg1.Rows = 1
        Fg1.Rows = Fg1.Rows + 1
        'TxtMedPag.Enabled = True
        'CmdMP.Enabled = True
    End If
End Sub

Private Sub OptCaja_Click()
    If OptCaja.Value = True Then
        Fg1.Rows = 1
        Fg1.Rows = Fg1.Rows + 1
        'TxtMedPag.Text = ""
        'LblMedPag.Caption = ""
        'TxtMedPag.Enabled = False
        'CmdMP.Enabled = False
    End If
End Sub

Sub MuestraSegundoTab()
    Dim rst As New ADODB.Recordset
    Dim Rst2 As New ADODB.Recordset
    Dim A, x As Integer
    Dim xId As Double
    Dim nSQL As String
    Dim xModulo As Integer '--Codigo del modulo
    Dim xConcepto  As Integer '--Codigo del concepto
    
    Blanquea
    
    If RstMov.BOF = True Or RstMov.EOF = True Or RstMov.RecordCount = 0 Then Exit Sub
    
    lblReg.Caption = "Nº Reg. " & NulosC(RstMov("registro"))
    
    TxtFchMov.Valor = RstMov("fchope")
    lblTipCambio.Caption = Format(BuscaTC(CDate(TxtFchMov.Valor), 2), "0.000")
    
    TxtIdMon.Text = RstMov("idmon")
    TxtIdMon_Validate True
    TxtGlosa.Text = NulosC(RstMov("glosa"))
    xId = RstMov("id")
    
    If NulosN(RstMov("importe")) = 0 Then
        ChkChequeAnulado.Value = 1
    Else
        ChkChequeAnulado.Value = 0
    End If
    
    '--definir el rst temporal para detallar las operaciones
    PreparaRST
    PreparaRSTOri
    
    '---------
    'mostramos el destino del movimiento "DEBE"
    
    nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion AS descuen, tes_cajadestino.iddes AS id, tes_destino.idmon, " _
        & " tes_destino.descripcion, tes_destino.idcuen, tes_destino.tipmov, tes_cajadestino.importe, 0 AS numdocasi, tes_destino.idmod, tes_cajadestino.tc, tes_destino.idbcocta " _
        & " FROM (tes_destino LEFT JOIN con_planctas ON tes_destino.idcuen = con_planctas.id) RIGHT JOIN tes_cajadestino ON tes_destino.id = tes_cajadestino.iddes " _
        & " WHERE (((tes_cajadestino.idtes)=" & xId & "))"
        
    RST_Busq rst, nSQL, xCon
    
    If rst.RecordCount <> 0 Then
        
        rst.MoveFirst
        
'        Agregando = True
        
        For A = 1 To rst.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            
            Fg2.TextMatrix(A, 1) = NulosC(rst("descripcion"))
            Fg2.TextMatrix(A, 2) = NulosN(rst("tc"))
            Fg2.TextMatrix(A, 3) = NulosN(rst("id"))
            Fg2.TextMatrix(A, 4) = NulosN(rst("idcuen"))
            Fg2.TextMatrix(A, 5) = NulosN(rst("idmod"))
            Fg2.TextMatrix(A, 6) = NulosN(rst("numdocasi"))   'especifica el numero de documentos asignado al destino
            If TxtIdMon.Text = 1 Then
                Fg2.TextMatrix(A, 7) = NulosN(rst("importe"))
            Else
                Fg2.TextMatrix(A, 8) = NulosN(rst("importe"))
            End If
            Fg2.TextMatrix(A, 9) = NulosN(rst("idbcocta"))
            
            Set Rst2 = Nothing
            
            xModulo = NulosN(rst("idmod"))
            xConcepto = NulosN(rst("id"))
            
            nSQL = ""
            Select Case xModulo
                Case 1 '--Compras
                    '--comprobante de percepcion unido a documentos de compras (mod. compras)
                    nSQL = "SELECT tes_cajadestinodet.idtes, con_percepcion.fchdoc, tes_cajadestinodet.iddes, tes_cajadestinodet.idmod, " _
                        & " tes_cajadestinodet.iddoc, con_percepcion.tipdoc, mae_documento.abrev, [con_percepcion]![numser]+'-'+[con_percepcion]![numdoc] AS numdoc, " _
                        & " mae_prov.nombre, con_percepcion.imptotper as imptot, con_percepcion.idmon, mae_moneda.simbolo, tes_cajadestinodet.saldo, " _
                        & " tes_cajadestinodet.acuenta, tes_cajadestinodet.idori FROM (((tes_cajadestinodet LEFT JOIN con_percepcion ON tes_cajadestinodet.iddoc = con_percepcion.id) " _
                        & " LEFT JOIN mae_documento ON con_percepcion.tipdoc = mae_documento.id) LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id) " _
                        & " LEFT JOIN mae_moneda ON con_percepcion.idmon = mae_moneda.id WHERE (((tes_cajadestinodet.idtes)=" & xId & ") " _
                        & " AND ((tes_cajadestinodet.iddes)=" & xConcepto & ") AND ((tes_cajadestinodet.idori)=2))" _
                        & " Union " _
                        & " SELECT tes_cajadestinodet.idtes, com_compras.fchdoc, tes_cajadestinodet.iddes, tes_cajadestinodet.idmod, tes_cajadestinodet.iddoc, " _
                        & " com_compras.tipdoc, mae_documento.abrev, iif(com_compras!numser is null or com_compras!numser ='','', com_compras!numser & '-')  & com_compras!numdoc AS numdoc, mae_prov.nombre, com_compras.imptot, " _
                        & " com_compras.idmon, mae_moneda.simbolo, tes_cajadestinodet.saldo, tes_cajadestinodet.acuenta, tes_cajadestinodet.idori " _
                        & " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (tes_cajadestinodet LEFT JOIN com_compras " _
                        & " ON tes_cajadestinodet.iddoc = com_compras.id) ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) " _
                        & " ON mae_prov.id = com_compras.idpro WHERE tes_cajadestinodet.idtes=" & xId & " AND tes_cajadestinodet.iddes=" & xConcepto & " " _
                        & " AND (tes_cajadestinodet.idori in (0,1) or tes_cajadestinodet.idori is null) "
                
                Case 8 '-- boleta pago
                    nSQL = "SELECT [pla_boleta]![numser]+'-'+[pla_boleta]![numdoc] AS numdoc, UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom] AS nombre, " _
                        & " mae_moneda.simbolo, mae_documento.abrev, tes_cajadestinodet.idtes, tes_cajadestinodet.iddes, tes_cajadestinodet.idtipper, tes_cajadestinodet.idmod, tes_cajadestinodet.iddoc, " _
                        & " tes_cajadestinodet.idper, tes_cajadestinodet.saldo, tes_cajadestinodet.acuenta, pla_boleta.fchdoc, pla_boleta.idmon, pla_boleta.id, pla_boleta.iddoc as tipdoc, pla_boleta.imptot, tes_cajadestinodet.idori " _
                        & " FROM ((pla_empleados RIGHT JOIN (tes_cajadestinodet LEFT JOIN pla_boleta ON tes_cajadestinodet.iddoc = pla_boleta.id) ON pla_empleados.id = pla_boleta.idemp) " _
                        & " LEFT JOIN mae_moneda ON pla_boleta.idmon = mae_moneda.id) LEFT JOIN mae_documento ON pla_boleta.iddoc = mae_documento.id " _
                        & " WHERE (((tes_cajadestinodet.idtes)=" & xId & ") AND ((tes_cajadestinodet.iddes)=" & xConcepto & "))"
                
                Case 9 '--honorarios
                    nSQL = "SELECT tes_cajadestinodet.idtes, com_honorarios.fchdoc, tes_cajadestinodet.iddes, tes_cajadestinodet.idmod, tes_cajadestinodet.iddoc, com_honorarios.tipdoc , mae_documento.abrev, iif( com_honorarios.numser is null or  com_honorarios.numser='', com_honorarios.numser & '-' ) & com_honorarios.numdoc AS numdoc, mae_prov.nombre, com_honorarios.imptot, com_honorarios.idmon, mae_moneda.simbolo, tes_cajadestinodet.saldo, tes_cajadestinodet.acuenta, com_honorarios.glosa ,tes_cajadestinodet.idori " _
                        + vbCr + " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (tes_cajadestinodet LEFT JOIN com_honorarios ON tes_cajadestinodet.iddoc = com_honorarios.id) ON mae_documento.id = com_honorarios.tipdoc) ON mae_moneda.id = com_honorarios.idmon) ON mae_prov.id = com_honorarios.idpro " _
                        + vbCr + " WHERE (((tes_cajadestinodet.idtes)=" & xId & ") AND ((tes_cajadestinodet.iddes)=" & xConcepto & "));"
                
                Case 10 '--reembolsables
                    nSQL = "SELECT tes_cajadestinodet.idtes, com_reembolsables.fchdoc, tes_cajadestinodet.iddes, tes_cajadestinodet.idmod, tes_cajadestinodet.iddoc, com_reembolsables.tipdoc as tipdoc, mae_documento.abrev, com_reembolsables.numser+'-'+com_reembolsables.numdoc AS numdoc, mae_prov.nombre, com_reembolsables.imptot, com_reembolsables.idmon, mae_moneda.simbolo, tes_cajadestinodet.saldo, tes_cajadestinodet.acuenta, com_reembolsables.glosa ,tes_cajadestinodet.idori " _
                        + vbCr + " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (tes_cajadestinodet LEFT JOIN com_reembolsables ON tes_cajadestinodet.iddoc = com_reembolsables.id) ON mae_documento.id = com_reembolsables.tipdoc) ON mae_moneda.id = com_reembolsables.idmon) ON mae_prov.id = com_reembolsables.idpro " _
                        + vbCr + " WHERE (((tes_cajadestinodet.idtes)=" & xId & ") AND ((tes_cajadestinodet.iddes)=" & xConcepto & "));"
                
                Case 2 '--ventas
                    nSQL = "SELECT tes_cajadestinodet.idtes, vta_ventas.fchdoc, tes_cajadestinodet.iddes, tes_cajadestinodet.idmod, tes_cajadestinodet.iddoc, mae_documento.abrev, [vta_ventas].[numser] & '-' & [vta_ventas].[numdoc] AS numdoc, mae_cliente.nombre, vta_ventas.imptotdoc AS imptot, vta_ventas.idmon, mae_moneda.simbolo, tes_cajadestinodet.saldo, tes_cajadestinodet.acuenta, vta_ventas.tipdoc, tes_cajadestinodet.idori " _
                        + vbCr + " FROM tes_cajadestinodet LEFT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN (mae_cliente RIGHT JOIN vta_ventas ON mae_cliente.id = vta_ventas.idcli) ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON tes_cajadestinodet.iddoc = vta_ventas.id " _
                        + vbCr + " WHERE (((tes_cajadestinodet.idtes)=" & xId & ") AND ((tes_cajadestinodet.iddes)=" & xConcepto & ")) "
    
                Case 4 '--letras
                    nSQL = "SELECT tes_cajadestinodet.idtes, let_letradet.fchemi AS fchdoc, tes_cajadestinodet.iddes, tes_cajadestinodet.idmod, tes_cajadestinodet.iddoc, " _
                        & " mae_documento.abrev, [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser] AS numdoc, mae_cliente.nombre, " _
                        & " let_letradet.implet AS imptot, let_letra.idmon, mae_moneda.simbolo, tes_cajadestinodet.saldo, tes_cajadestinodet.acuenta, let_letra.tipdoc, tes_cajadestinodet.idori " _
                        & " FROM tes_cajadestinodet LEFT JOIN (((mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN let_letra ON mae_documento.id = let_letra.tipdoc) ON mae_moneda.id = let_letra.idmon) " _
                        & " RIGHT JOIN let_letradet ON let_letra.id = let_letradet.idlet) LEFT JOIN mae_cliente ON let_letra.idclipro = mae_cliente.id) ON tes_cajadestinodet.iddoc = let_letradet.corr" _
                        & " WHERE (((tes_cajadestinodet.idtes)=" & xId & ") AND ((tes_cajadestinodet.iddes)=" & xConcepto & "))"
                
                Case 11 '--lgd
                    nSQL = "SELECT tes_cajadestinodet.idtes, vta_gastodebito.fchemi AS fchdoc, tes_cajadestinodet.iddes, tes_cajadestinodet.idmod, tes_cajadestinodet.iddoc, vta_gastodebito.tipdoc, mae_documento.abrev, " _
                        & " [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc] AS numdoc, mae_cliente.nombre, vta_gastodebito.imptot, vta_gastodebito.idmon, mae_moneda.simbolo, tes_cajadestinodet.saldo, tes_cajadestinodet.acuenta, tes_cajadestinodet.idori " _
                        & " FROM (((tes_cajadestinodet LEFT JOIN vta_gastodebito ON tes_cajadestinodet.iddoc = vta_gastodebito.id) LEFT JOIN mae_cliente ON vta_gastodebito.idcli = mae_cliente.id) LEFT JOIN mae_moneda ON " _
                        & " vta_gastodebito.idmon = mae_moneda.id) LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id " _
                        & " WHERE (((tes_cajadestinodet.idtes)=" & xId & ") AND ((tes_cajadestinodet.iddes)=" & xConcepto & "));"
                
                Case 19 '--planilla letras
                    nSQL = "SELECT tes_cajadestinodet.idtes, let_planilla.fchemi AS fchdoc, tes_cajadestinodet.iddes, tes_cajadestinodet.idmod, tes_cajadestinodet.iddoc, mae_documento.abrev, let_planilla.numdoc, Trim([mae_bancos].[descripcion]) & ' N° Cta. ' & Trim([mae_banconumcta].[numcue]) AS nombre, let_planilla.imptot, let_planilla.idmon, mae_moneda.simbolo, tes_cajadestinodet.saldo, tes_cajadestinodet.acuenta, let_planilla.tipdoc, tes_cajadestinodet.idori " _
                        + vbCr + " FROM mae_bancos INNER JOIN (mae_banconumcta INNER JOIN (tes_cajadestinodet INNER JOIN (mae_moneda INNER JOIN (mae_documento INNER JOIN let_planilla ON mae_documento.id = let_planilla.tipdoc) ON mae_moneda.id = let_planilla.idmon) ON tes_cajadestinodet.iddoc = let_planilla.id) ON mae_banconumcta.id = let_planilla.idbcocta) ON mae_bancos.id = mae_banconumcta.idban " _
                        + vbCr + " WHERE (((tes_cajadestinodet.idtes)=" & xId & ") AND ((tes_cajadestinodet.iddes)=" & xConcepto & ")) "
                
                Case 20 '--Orden Despacho
                    nSQL = "SELECT tes_cajadestinodet.idtes, var_ordendespacho.fchdoc, tes_cajadestinodet.iddes, tes_cajadestinodet.idmod, tes_cajadestinodet.iddoc, var_ordendespacho.tipdoc, mae_documento.abrev, var_ordendespacho.numdoc, mae_cliente.nombre, 0 AS imptot, tes_caja.idmon, mae_moneda.simbolo, tes_cajadestinodet.saldo, tes_cajadestinodet.acuenta, '' AS glosa, tes_cajadestinodet.idori " _
                        + vbCr + " FROM ((var_ordendespacho LEFT JOIN mae_cliente ON var_ordendespacho.idcli = mae_cliente.id) LEFT JOIN mae_documento ON var_ordendespacho.tipdoc = mae_documento.id) INNER JOIN (mae_moneda INNER JOIN (tes_cajadestinodet INNER JOIN tes_caja ON tes_cajadestinodet.idtes = tes_caja.id) ON mae_moneda.id = tes_caja.idmon) ON var_ordendespacho.id = tes_cajadestinodet.iddoc " _
                        + vbCr + " WHERE (((tes_cajadestinodet.idtes)=" & xId & ") AND ((tes_cajadestinodet.iddes)=" & xConcepto & ")); "
    
                Case 6 '--Bancos
                    nSQL = "SELECT tes_cajadestinodet.*, mae_documento.abrev, IIf([tes_cajadestinodet].[numser] Is Null,'',[tes_cajadestinodet].[numser] & '-') & [tes_cajadestinodet].[numdoc] AS numdoc, IIf([tes_cajadestinodet].[idtipper]=1 Or [tes_cajadestinodet].[idtipper] Is Null,[mae_prov].[nombre],IIf([tes_cajadestinodet].[idtipper]=1,[mae_cliente].[nombre],IIf([tes_cajadestinodet].[idtipper]=3,[pla_empleados].[nombre],IIf([tes_cajadestinodet].[idtipper]=5,[mae_bancos].[descripcion],'')))) AS nombre,mae_moneda.simbolo, tes_cajadestinodet.importe AS imptot  " _
                        + vbCr + " FROM (((((tes_cajadestinodet LEFT JOIN mae_documento ON tes_cajadestinodet.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON tes_cajadestinodet.idmon = mae_moneda.id) LEFT JOIN pla_empleados ON tes_cajadestinodet.idper = pla_empleados.id) LEFT JOIN mae_cliente ON tes_cajadestinodet.idper = mae_cliente.id) LEFT JOIN mae_prov ON tes_cajadestinodet.idper = mae_prov.id) LEFT JOIN mae_bancos ON tes_cajadestinodet.idper = mae_bancos.id " _
                        + vbCr + " WHERE (((tes_cajadestinodet.idtes)=" & xId & ") AND ((tes_cajadestinodet.iddes)=" & xConcepto & "));"

                Case Else
                    nSQL = ""
                
            End Select
            
            If nSQL <> "" Then '--Si hay consulta se procedera a detallar
                
                RST_Busq Rst2, nSQL, xCon

                If Rst2.RecordCount <> 0 Then
                    '--aplicar orden
                    Rst2.Sort = "nombre,numdoc"
                    Rst2.MoveFirst
                    For x = 1 To Rst2.RecordCount
                        RstTMPDoc.AddNew
                        RstTMPDoc("cliente") = NulosC(Rst2("nombre"))            'descripcion del medio de pago
                        RstTMPDoc("tipdoc") = NulosC(Rst2("abrev"))              'abreviatura del tipo de documento
                        RstTMPDoc("fchemi") = NulosC(Rst2("fchdoc"))             'fecha de emision del documento
                        RstTMPDoc("moneda") = NulosC(Rst2("simbolo"))    'descripcion de la moneda
                        RstTMPDoc("numdoc") = NulosC(Rst2("numdoc"))
                        RstTMPDoc("imptot") = NulosN(Rst2("imptot"))
                        RstTMPDoc("impsal") = NulosN(Rst2("saldo"))
                        
                        If Rst2("idmon") = NulosN(TxtIdMon.Text) Then
                            RstTMPDoc("impsal2") = NulosN(Rst2("saldo"))
                        Else
                            If NulosN(TxtIdMon.Text) = 1 Then
                                If Rst2("idmon") = 2 Then
                                    RstTMPDoc("impsal2") = NulosN(Rst2("saldo")) * NulosN(Fg2.TextMatrix(A, 2))
                                End If
                            Else
                                If Rst2("idmon") = 1 Then
                                    RstTMPDoc("impsal2") = NulosN(Rst2("saldo")) / NulosN(Fg2.TextMatrix(A, 2))
                                End If
                            End If
                        End If
                        
                        RstTMPDoc("acuent") = NulosN(Rst2("acuenta"))
                        
                        RstTMPDoc("newsal") = 0
                        RstTMPDoc("idconc") = NulosN(Rst2("iddes"))
                        RstTMPDoc("iddocu") = NulosN(Rst2("iddoc"))
                        RstTMPDoc("idmone") = NulosN(Rst2("idmon"))
                        RstTMPDoc("idtipd") = NulosN(Rst2("tipdoc"))           'codigo del medio de pago
                        RstTMPDoc("idori") = NulosN(Rst2("idori"))             ' codigo del origen del documento 1 = compras: 2 = percepcion
                        
                        RstTMPDoc("corr") = mCorrelativo2
                        RstTMPDoc("xtipo") = 1
                        mCorrelativo2 = mCorrelativo2 + 1
                                                
                        Rst2.MoveNext
                        If Rst2.EOF = True Then Exit For
                    Next x
                End If
            End If
            
            '****************************************************************************
            nSQL = ""
            Select Case IdModulo
                Case 7 '--fondo fijo
                    nSQL = "SELECT tes_cajadestinodet.*, tes_documentos.abrev, UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom] AS apenom" _
                        & " FROM ((tes_cajadestinodet LEFT JOIN tes_documentos ON tes_cajadestinodet.tipdoc = tes_documentos.id) LEFT JOIN tes_usuarios " _
                        & " ON tes_cajadestinodet.idper = tes_usuarios.id) LEFT JOIN pla_empleados ON tes_usuarios.idper = pla_empleados.id " _
                        & " WHERE (((tes_cajadestinodet.idtes)=" & xId & ") AND ((tes_cajadestinodet.iddes)=" & xConcepto & "))"
                Case 5 '--Anticipos a Proveedores
                    nSQL = "SELECT tes_cajadestinodet.*, mae_prov.nombre, tes_documentos.abrev FROM (tes_cajadestinodet LEFT JOIN mae_prov " _
                        & " ON tes_cajadestinodet.idper = mae_prov.id) LEFT JOIN tes_documentos ON tes_cajadestinodet.tipdoc = tes_documentos.id " _
                        & " WHERE (((tes_cajadestinodet.idtes)=" & xId & ") AND ((tes_cajadestinodet.iddes)=" & rst("id") & "))"
                
                Case 3 '--Entregas a Rendir
                    nSQL = "SELECT tes_cajadestinodet.*, tes_documentos.abrev, UCase([pla_empleados]![apepat])+' '+UCase([pla_empleados]![apemat])+', '+[pla_empleados]![nom] AS apenom" _
                        & " FROM ((tes_cajadestinodet LEFT JOIN tes_documentos ON tes_cajadestinodet.tipdoc = tes_documentos.id) LEFT JOIN tes_usuarios " _
                        & " ON tes_cajadestinodet.idper = tes_usuarios.id) LEFT JOIN pla_empleados ON tes_usuarios.idper = pla_empleados.id " _
                        & " WHERE (((tes_cajadestinodet.idtes)=" & xId & ") AND ((tes_cajadestinodet.iddes)=" & xConcepto & "))"
                
                Case Else
                    nSQL = ""
            
            End Select
            
            If nSQL <> "" Then
                 
                 RST_Busq Rst2, nSQL, xCon

                If Rst2.RecordCount <> 0 Then
                    Rst2.MoveFirst
                    For x = 1 To Rst2.RecordCount
                        RstTMPDoc.AddNew
                        RstTMPDoc("idconc") = NulosN(Rst2("iddes"))
                        RstTMPDoc("cliente") = NulosC(Rst2("apenom"))
                        RstTMPDoc("tipdoc") = NulosC(Rst2("abrev"))
                        RstTMPDoc("fchemi") = NulosC(Rst2("fchdoc"))
                        RstTMPDoc("numdoc") = NulosC(Rst2("numser")) + "-" + NulosC(Rst2("numdoc"))
                        RstTMPDoc("imptot") = NulosN(Rst2("importe"))
                        RstTMPDoc("idtipd") = NulosC(Rst2("tipdoc"))
                        RstTMPDoc("iddocu") = NulosN(Rst2("idper"))
                        
                        RstTMPDoc("corr") = mCorrelativo2
                        RstTMPDoc("xtipo") = 1
                        mCorrelativo2 = mCorrelativo2 + 1
                        
                    Next x
                End If
            End If
                        
'****************************************************************************

            rst.MoveNext
            If rst.EOF = True Then Exit For
        Next A
        TotalizarFG2
        Agregando = False
        
    End If
            
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'mostramos el origen del movimiento "HABER"
    Set rst = Nothing
    
    nSQL = "SELECT con_planctas.cuenta, con_planctas.descripcion AS descuen, tes_origen.id, tes_origen.idmon, tes_origen.descripcion, tes_origen.idcuen, tes_origen.tipmov, " _
        + vbCr + " tes_origen.idmod, tes_origen.entgen, 0 AS numdocasi, tes_cajaori.importe, tes_cajaori.tc, tes_origen.idbcocta " _
        + vbCr + " FROM (tes_origen LEFT JOIN con_planctas ON tes_origen.idcuen = con_planctas.id) RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori " _
        + vbCr + " WHERE (((tes_cajaori.idtes)=" & xId & ")) "
    
    RST_Busq rst, nSQL, xCon

    If rst.RecordCount <> 0 Then
        rst.MoveFirst
        
        For A = 1 To rst.RecordCount
            Fg1.Rows = Fg1.Rows + 1
            Fg1.TextMatrix(A, 1) = NulosC(rst("descripcion"))
            Fg1.TextMatrix(A, 2) = NulosN(rst("tc"))
            Fg1.TextMatrix(A, 3) = NulosN(rst("id"))
            Fg1.TextMatrix(A, 4) = NulosN(rst("idcuen"))
            Fg1.TextMatrix(A, 5) = NulosN(rst("idmod"))
            Fg1.TextMatrix(A, 6) = NulosN(rst("numdocasi"))   'especifica el numero de documentos asignado al destino
            If TxtIdMon.Text = 1 Then
                Fg1.TextMatrix(A, 7) = NulosN(rst("importe"))
            Else
                Fg1.TextMatrix(A, 8) = NulosN(rst("importe"))
            End If
            Fg1.TextMatrix(A, 9) = NulosN(rst("idbcocta"))
            
            xModulo = NulosN(rst("idmod"))
            xConcepto = NulosN(rst("id"))
            nSQL = ""
            
            Select Case xModulo
                Case 1 '--compras
                    nSQL = "SELECT tes_cajaorigendet.idtes, com_compras.fchdoc, tes_cajaorigendet.idori, tes_cajaorigendet.idmod, tes_cajaorigendet.iddoc, com_compras.tipdoc, mae_documento.abrev, IIf(com_compras!numser Is Null Or com_compras!numser='','',com_compras!numser & '-') & com_compras!numdoc AS numdoc, mae_prov.nombre, com_compras.imptot, com_compras.idmon, mae_moneda.simbolo, tes_cajaorigendet.saldo, tes_cajaorigendet.acuenta " _
                                + vbCr + " FROM tes_cajaorigendet LEFT JOIN (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro) ON tes_cajaorigendet.iddoc = com_compras.id " _
                                + vbCr + " WHERE (((tes_cajaorigendet.idtes)=" & xId & ") AND ((tes_cajaorigendet.idori)=" & xConcepto & " )); "
                Case 10 '--reembolsables
                    nSQL = "SELECT tes_cajaorigendet.idtes, com_reembolsables.fchdoc, tes_cajaorigendet.idori, tes_cajaorigendet.idmod, tes_cajaorigendet.iddoc, com_reembolsables.tipdoc, mae_documento.abrev, IIf(com_reembolsables!numser Is Null Or com_reembolsables!numser='','',com_reembolsables!numser & '-') & com_reembolsables!numdoc AS numdoc, mae_prov.nombre, com_reembolsables.imptot, com_reembolsables.idmon, mae_moneda.simbolo, tes_cajaorigendet.saldo, tes_cajaorigendet.acuenta " _
                                + vbCr + " FROM tes_cajaorigendet LEFT JOIN (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN com_reembolsables ON mae_documento.id = com_reembolsables.tipdoc) ON mae_moneda.id = com_reembolsables.idmon) ON mae_prov.id = com_reembolsables.idpro) ON tes_cajaorigendet.iddoc = com_reembolsables.id " _
                                + vbCr + " WHERE (((tes_cajaorigendet.idtes)=" & xId & ") AND ((tes_cajaorigendet.idori)=" & xConcepto & " )); "
                
                Case 2 '--ventas
                    nSQL = "SELECT tes_cajaorigendet.idtes, vta_ventas.fchdoc, tes_cajaorigendet.idori, tes_cajaorigendet.idmod, tes_cajaorigendet.iddoc, vta_ventas.tipdoc, mae_documento.abrev, IIf(vta_ventas!numser Is Null Or vta_ventas!numser='','',vta_ventas!numser & '-') & vta_ventas!numdoc AS numdoc, mae_cliente.nombre, vta_ventas.imptotdoc as imptot, vta_ventas.idmon, mae_moneda.simbolo, tes_cajaorigendet.saldo, tes_cajaorigendet.acuenta " _
                                + vbCr + " FROM tes_cajaorigendet LEFT JOIN (mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN vta_ventas ON mae_documento.id = vta_ventas.tipdoc) ON mae_moneda.id = vta_ventas.idmon) ON mae_cliente.id = vta_ventas.idcli) ON tes_cajaorigendet.iddoc = vta_ventas.id " _
                                + vbCr + " WHERE (((tes_cajaorigendet.idtes)=" & xId & ") AND ((tes_cajaorigendet.idori)=" & xConcepto & " )) "
                
                Case 4 '--letras
                    nSQL = "SELECT tes_cajaorigendet.idtes, let_letradet.fchemi AS fchdoc, tes_cajaorigendet.idori, tes_cajaorigendet.idmod, tes_cajaorigendet.iddoc, mae_documento.abrev, [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser] AS numdoc, mae_cliente.nombre, let_letradet.implet AS imptot, let_letra.idmon, mae_moneda.simbolo, tes_cajaorigendet.saldo, tes_cajaorigendet.acuenta, let_letra.tipdoc " _
                                + vbCr + " FROM tes_cajaorigendet LEFT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((let_letra LEFT JOIN mae_cliente ON let_letra.idclipro = mae_cliente.id) RIGHT JOIN let_letradet ON let_letra.id = let_letradet.idlet) ON mae_documento.id = let_letra.tipdoc) ON mae_moneda.id = let_letra.idmon) ON tes_cajaorigendet.iddoc = let_letradet.corr " _
                                + vbCr + " WHERE (((tes_cajaorigendet.idtes)=" & xId & ") AND ((tes_cajaorigendet.idori)=" & xConcepto & " )) "
    
                Case 11 '--lgd
                    nSQL = "SELECT tes_cajaorigendet.idtes, vta_gastodebito.fchemi as fchdoc, tes_cajaorigendet.idori, tes_cajaorigendet.idmod, tes_cajaorigendet.iddoc, vta_gastodebito.tipdoc, mae_documento.abrev, IIf(vta_gastodebito!numser Is Null Or vta_gastodebito!numser='','',vta_gastodebito!numser & '-') & vta_gastodebito!numdoc AS numdoc, mae_cliente.nombre, vta_gastodebito.imptotdoc as imptot, vta_gastodebito.idmon, mae_moneda.simbolo, tes_cajaorigendet.saldo, tes_cajaorigendet.acuenta " _
                                + vbCr + " FROM tes_cajaorigendet LEFT JOIN (mae_cliente RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN vta_gastodebito ON mae_documento.id = vta_gastodebito.tipdoc) ON mae_moneda.id = vta_gastodebito.idmon) ON mae_cliente.id = vta_gastodebito.idcli) ON tes_cajaorigendet.iddoc = vta_gastodebito.id " _
                                + vbCr + " WHERE (((tes_cajaorigendet.idtes)=" & xId & ") AND ((tes_cajaorigendet.idori)=" & xConcepto & " )) "
                Case 19 '--planilla letras
                    nSQL = "SELECT tes_cajaorigendet.idtes, let_planilla.fchemi AS fchdoc, tes_cajaorigendet.idori, tes_cajaorigendet.idmod, tes_cajaorigendet.iddoc, mae_documento.abrev, let_planilla.numdoc, Trim([mae_bancos].[descripcion]) & ' N° Cta. ' & Trim([mae_banconumcta].[numcue]) AS nombre, let_planilla.imptot, let_planilla.idmon, mae_moneda.simbolo, tes_cajaorigendet.saldo, tes_cajaorigendet.acuenta, let_planilla.tipdoc " _
                                + vbCr + " FROM tes_cajaorigendet INNER JOIN (mae_moneda INNER JOIN (mae_documento INNER JOIN (mae_bancos INNER JOIN (mae_banconumcta INNER JOIN let_planilla ON mae_banconumcta.id = let_planilla.idbcocta) ON mae_bancos.id = mae_banconumcta.idban) ON mae_documento.id = let_planilla.tipdoc) ON mae_moneda.id = let_planilla.idmon) ON tes_cajaorigendet.iddoc = let_planilla.id " _
                                + vbCr + " WHERE (((tes_cajaorigendet.idtes)=" & xId & ") AND ((tes_cajaorigendet.idori)=" & xConcepto & " )) "

                Case Else
                    nSQL = ""
            End Select
            
            If nSQL <> "" Then
                
                RST_Busq Rst2, nSQL, xCon
                
                If Rst2.RecordCount <> 0 Then
                    '--aplicar orden
                    Rst2.Sort = "nombre,numdoc"
                    Rst2.MoveFirst
                    For x = 1 To Rst2.RecordCount
                        RstTMPDoc.AddNew
                        RstTMPDoc("cliente") = NulosC(Rst2("nombre"))            'descripcion del medio de pago
                        RstTMPDoc("tipdoc") = NulosC(Rst2("abrev"))              'abreviatura del tipo de documento
                        RstTMPDoc("fchemi") = NulosC(Rst2("fchdoc"))             'fecha de emision del documento
                        RstTMPDoc("moneda") = NulosC(Rst2("simbolo"))    'descripcion de la moneda
                        RstTMPDoc("numdoc") = NulosC(Rst2("numdoc"))
                        RstTMPDoc("imptot") = NulosN(Rst2("imptot"))
                        RstTMPDoc("impsal") = NulosN(Rst2("saldo"))
                        If Rst2("idmon") = NulosN(TxtIdMon.Text) Then
                            RstTMPDoc("impsal2") = NulosN(Rst2("saldo"))
                        Else
                            If NulosN(TxtIdMon.Text) = 1 Then
                                If Rst2("idmon") = 2 Then
                                    RstTMPDoc("impsal2") = NulosN(Rst2("saldo")) * NulosN(Fg2.TextMatrix(A, 2))
                                End If
                            Else
                                If Rst2("idmon") = 1 Then
                                    RstTMPDoc("impsal2") = NulosN(Rst2("saldo")) / NulosN(Fg2.TextMatrix(A, 2))
                                End If
                            End If
                        End If
                        
                        RstTMPDoc("acuent") = NulosN(Rst2("acuenta"))
                        RstTMPDoc("newsal") = 0
                        RstTMPDoc("idconc") = NulosN(Rst2("idori"))
                        RstTMPDoc("iddocu") = NulosN(Rst2("iddoc"))
                        RstTMPDoc("idmone") = NulosN(Rst2("idmon"))
                        RstTMPDoc("idtipd") = NulosN(Rst2("tipdoc"))           'codigo del medio de pago
                        RstTMPDoc("idori") = 1             ' codigo del origen del documento 1 = compras: 2 = percepcion
                        RstTMPDoc("xtipo") = 0
                        RstTMPDoc("corr") = mCorrelativo2
                        mCorrelativo2 = mCorrelativo2 + 1
                                                
                        Rst2.MoveNext
                        If Rst2.EOF = True Then Exit For
                    Next x
                End If
           Else
           
                '--ver si tiene detalle
                Set Rst2 = Nothing
                RST_Busq Rst2, "SELECT tes_cajaorigendet.*, tes_documentos.abrev, tes_mediopago.descripcion FROM (tes_cajaorigendet LEFT JOIN tes_documentos " _
                    & " ON tes_cajaorigendet.tipdoc = tes_documentos.id) LEFT JOIN tes_mediopago ON tes_cajaorigendet.idmedpag = tes_mediopago.id " _
                    & " WHERE (((tes_cajaorigendet.idtes)=" & xId & ")) and tes_cajaorigendet.idori = " & NulosN(rst("id")), xCon
                
                If Rst2.State = 1 Then
                    If Rst2.RecordCount <> 0 Then
                        Rst2.MoveFirst
                        For x = 1 To Rst2.RecordCount
                            RstTmpDocOri.AddNew
                            RstTmpDocOri("cliente") = NulosC(Rst2("descripcion"))       'descripcion del medio de pago
                            RstTmpDocOri("tipdoc") = NulosC(Rst2("abrev"))              'abreviatura del tipo de documento
                            RstTmpDocOri("fchemi") = ""                        'fecha de emision del documento
                            RstTmpDocOri("moneda") = NulosC(LblMoneda.Caption) 'descripcion de la moneda
                            RstTmpDocOri("numdoc") = NulosC(Rst2("numdoc"))
                            RstTmpDocOri("imptot") = NulosN(Rst2("importe"))
                            RstTmpDocOri("idtipd") = NulosC(Rst2("tipdoc"))           'codigo del medio de pago
                            RstTmpDocOri("idconc") = NulosC(Rst2("idori"))
                            RstTmpDocOri("iddocu") = NulosC(Rst2("idmedpag"))
                            'RstTmpDocOri("idbcocta") = NulosC(Rst2("idmedpag"))
                            Rst2.MoveNext
                            If Rst2.EOF = True Then Exit For
                        Next x
                    End If
                End If

            End If
            
            rst.MoveNext
            
            If rst.EOF = True Then Exit For
        Next A
    End If
    
    '--colocando el cursor en la primera columna
    If Fg1.Rows > Fg1.FixedRows Then
        Fg1.Row = 1
        Fg1.Col = 1
    End If
    
    If Fg2.Rows > Fg2.FixedRows Then
        Fg2.Row = 1
        Fg2.Col = 1
    End If
    
    Set rst = Nothing
    Set Rst2 = Nothing
    
End Sub

Private Sub pic_Click(Index As Integer)
    Select Case Index
    Case 0 '--ver asiento
        Command5_Click
    Case 1 '--destnino
        Command11_Click
    Case 2 '--origen
        Command9_Click
    End Select
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then MuestraSegundoTab
    End If
End Sub

Sub Eliminar()
    
    If RstMov.BOF = True Or RstMov.EOF = True Or RstMov.RecordCount = 0 Then Exit Sub
    
    TabOne1.CurrTab = 0
    
    Dim Rpta As Integer
    Rpta = MsgBox("¿Esta seguro de eliminar el movimiento de egreso?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    
    If Rpta = vbYes Then
        xCon.Execute "DELETE * FROM tes_cajadestinodet WHERE idtes = " & RstMov("id") & " "
        xCon.Execute "DELETE * FROM tes_cajadestino WHERE idtes = " & RstMov("id") & " "
        
        xCon.Execute "DELETE * FROM tes_cajaorigendet WHERE idtes = " & RstMov("id") & " "
        xCon.Execute "DELETE * FROM tes_cajaori WHERE idtes = " & RstMov("id") & " "
        
        
        xCon.Execute "DELETE * FROM tes_caja WHERE id = " & RstMov("id") & " "
        
        xCon.Execute "DELETE * FROM con_diario WHERE idmov = " & RstMov("id") & " and idlib = 6 "
        
        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstMov("id") & " AND idform = " & IdMenuActivo
        
        MsgBox "El movimiento de egreso se eliminó correctamente", vbInformation, xTitulo
        
        Dg1.Refresh
        If RstMov.RecordCount = 0 Then
            Rpta = MsgBox("No se ha registrado movimientos, ¿Desea agregar uno ahora?", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
            If Rpta = vbYes Then
                Nuevo
            Else
                Set RstMov = Nothing
                Unload Me
            End If
        Else
            RstMov.Requery
            Dg1.Refresh
        End If
    End If
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then Nuevo
    
    If Button.Index = 2 Then Modificar
    
    If Button.Index = 3 Then Eliminar
    
    If Button.Index = 5 Then
        If Grabar = True Then
            Cancelar
            TDB_FiltroLimpiar Dg1
            RstMov.Requery
            Dg1.Refresh
            
            If RstMov.RecordCount <> 0 Then
                RstMov.MoveFirst
                RstMov.Find "id=" & mIdRegistro
                If RstMov.EOF = True Then RstMov.MoveFirst
            End If
            
        End If
    End If
    
    If Button.Index = 6 Then Cancelar
    
    
    If Button.Index = 9 Then
        TabOne1.CurrTab = 0
        TDB_FiltroLimpiar Dg1
        RstMov.Filter = ""
    End If
       
    
    If Button.Index = 11 Then CambiarMes
    
    If Button.Index = 13 Then
        '27/01/11 Johan Castro
        'Muestra seleccion de registro cuando esta en consulta
        'Imprimir directamente cuando esta en vista previa del registro
        If TabOne1.CurrTab = 1 And QueHace = 3 Then
            ImprimirOperacion 2, RstMov("id"), RstMov("fchreg"), xCon
        Else
            ImprimirOperacion 2, 0, RstMov("fchreg"), xCon
        End If
    End If
    
    If Button.Index = 15 Then
        Set RstMov = Nothing
        Unload Me
        Exit Sub
    End If
End Sub

Sub CambiarMes()
    TabOne1.CurrTab = 0
    mMesActivo = SeleccionaMes(xCon)
    OpcionesPeriodo
    If mMesActivo = 0 Or mMesActivo = 13 Then
        MsgBox "Selecione un Periodo Correcto", vbExclamation, xTitulo
        CambiarMes
        Exit Sub
    End If
    
    CargarRSTCom
End Sub

Sub CargarRSTCom()
    Dim Rpta As Integer
    Dim nSQL As String
    
    If mMesActivo = 0 Then
        Set RstMov = Nothing
        Exit Sub
    End If
    '--limpiar los filtros
    TDB_FiltroLimpiar Dg1
    Set RstMov = Nothing
    Set Dg1.DataSource = Nothing
    DoEvents
    Me.MousePointer = vbHourglass
    
    xFchPer = "01/" + Format(Trim(Str(mMesActivo)), "00") + "/" + Trim(Str(AnoTra))
    
    'RST_Busq RstMov, "SELECT con_cajabanco.*, mae_moneda.simbolo, mae_doccajaban.descripcion AS descdoc, con_bancocuenta.numcue, mae_bancos.descripcion AS descban, " _
        & " con_origen.descripcion AS descori, con_origen.idcue AS idcueori, mae_doccajaban.abrev AS abredoc, IIf(con_cajabanco!tipope=1,'Caja','Banco') AS motmov, " _
        & " con_cajabanco.tipmov, Mid([con_cajabanco]![numreg],1,2)+[mae_libros]![codsun]+Mid([con_cajabanco]![numreg],3,4) AS numregi FROM (mae_moneda RIGHT JOIN " _
        & " (mae_bancos RIGHT JOIN (((con_cajabanco LEFT JOIN con_bancocuenta ON con_cajabanco.idcueban = con_bancocuenta.id) LEFT JOIN mae_doccajaban " _
        & " ON con_cajabanco.iddoc = mae_doccajaban.id) LEFT JOIN con_origen ON con_cajabanco.idori = con_origen.id) ON mae_bancos.id = con_bancocuenta.idban) " _
        & " ON mae_moneda.id = con_cajabanco.idmon) LEFT JOIN mae_libros ON con_cajabanco.idlib = mae_libros.id WHERE (((con_cajabanco.tipmov)=2) " _
        & " AND ((con_cajabanco.fchreg)=CDate('" & xFchPer & "'))) ORDER BY con_cajabanco.id DESC", xCon
''12/03/10
''    nSQL = "SELECT tes_caja.id, tes_caja.fchreg, tes_caja.tipmov, tes_caja.fchope & '' AS fchope, tes_caja.numreg,  tes_caja.glosa, mae_moneda.simbolo,  tes_documentos.abrev, tes_documentos.descripcion AS descdoc, tes_origen.descripcion AS descori, IIf(IsNull(tes_cajaorigendet!numser)=-1,tes_cajaorigendet!numdoc,tes_cajaorigendet!numser & '-' & tes_cajaorigendet!numdoc) AS numdoc, 'Egreso' AS tipo, tes_cajaori.importe & '' AS importe, tes_caja.idmon, tes_documentos.abrev AS desdocabre, IIf([tes_caja].[numreg] Is Null,'',Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) AS registro " _
''        + vbCr + " FROM (((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN (tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (tes_cajaorigendet LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id " _
''        + vbCr + " GROUP BY tes_caja.id, tes_caja.fchreg, tes_caja.tipmov,  tes_caja.glosa,tes_caja.fchope & '', tes_caja.numreg, tes_caja.glosa, mae_moneda.simbolo, tes_documentos.abrev, tes_documentos.descripcion, tes_origen.descripcion, IIf(IsNull(tes_cajaorigendet!numser)=-1,tes_cajaorigendet!numdoc,tes_cajaorigendet!numser & '-' & tes_cajaorigendet!numdoc), 'Egreso', tes_cajaori.importe & '', tes_caja.idmon, tes_documentos.abrev, IIf([tes_caja].[numreg] Is Null,'',Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) " _
''        + vbCr + " HAVING (((tes_caja.fchreg)=CDate('" & xFchPer & "')) AND ((tes_caja.tipmov)=2))   " _
''        + vbCr + " ORDER BY tes_caja.numreg DESC "
''


nSQL = "SELECT cab.*, gl.glosa " _
    + vbCr + " FROM " _
    + vbCr + " ( SELECT tes_caja.id, tes_caja.fchreg, tes_caja.tipmov, tes_caja.fchope & '' AS fchope, tes_caja.numreg, mae_moneda.simbolo, tes_documentos.abrev, tes_documentos.descripcion AS descdoc, tes_origen.descripcion AS descori, IIf(IsNull(tes_cajaorigendet!numser)=-1,tes_cajaorigendet!numdoc,tes_cajaorigendet!numser & '-' & tes_cajaorigendet!numdoc) AS numdoc, 'Egreso' AS tipo, tes_cajaori.importe & '' AS importe, tes_caja.idmon, tes_documentos.abrev AS desdocabre, IIf([tes_caja].[numreg] Is Null,'',Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) AS registro " _
    + vbCr + " FROM (((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN (tes_origen RIGHT JOIN tes_cajaori ON tes_origen.id = tes_cajaori.idori) ON tes_caja.id = tes_cajaori.idtes) LEFT JOIN (tes_cajaorigendet LEFT JOIN tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) ON (tes_cajaori.idtes = tes_cajaorigendet.idtes) AND (tes_cajaori.idori = tes_cajaorigendet.idori)) LEFT JOIN mae_libros ON tes_caja.idlib = mae_libros.id " _
    + vbCr + " GROUP BY tes_caja.id, tes_caja.fchreg, tes_caja.tipmov, tes_caja.fchope & '', tes_caja.numreg, tes_caja.glosa, mae_moneda.simbolo, tes_documentos.abrev, tes_documentos.descripcion, tes_origen.descripcion, IIf(IsNull(tes_cajaorigendet!numser)=-1,tes_cajaorigendet!numdoc,tes_cajaorigendet!numser & '-' & tes_cajaorigendet!numdoc), 'Egreso', tes_cajaori.importe & '', tes_caja.idmon, tes_documentos.abrev, IIf([tes_caja].[numreg] Is Null,'',Left([tes_caja].[numreg],2) & [mae_libros].[codsun] & Right([tes_caja].[numreg],4)) " _
    + vbCr + " HAVING (((tes_caja.fchreg)=CDate('" & xFchPer & "')) AND ((tes_caja.tipmov)=2))  )  AS cab " _
    + vbCr + " LEFT JOIN " _
    + vbCr + " (SELECT tes_caja.id, tes_caja.glosa FROM tes_caja " _
    + vbCr + " WHERE (((tes_caja.fchreg)=CDate('" & xFchPer & "')) AND ((tes_caja.tipmov)=2)) )  AS gl ON cab.id = gl.id;"
    
    RST_Busq RstMov, nSQL, xCon
   
    
    'SELECT tes_caja.id, tes_caja.fchreg, tes_caja.tipmov, tes_caja.fchope, tes_caja.numreg, tes_caja.glosa, mae_moneda.simbolo, tes_cajaorigendet.iddoc, " _
        & " tes_documentos.abrev, tes_documentos.descripcion AS descdoc, tes_origen.descripcion AS descori, IIf(IsNull(tes_cajaorigendet!numser)=-1,tes_cajaorigendet!numdoc, " _
        & " tes_cajaorigendet!numser & '-' & tes_cajaorigendet!numdoc) AS numdoc, 'Egreso' AS tipo, tes_cajaori.importe, tes_caja.idmon, tes_documentos.abrev AS desdocabre " _
        & " FROM tes_origen RIGHT JOIN ((tes_caja LEFT JOIN mae_moneda ON tes_caja.idmon = mae_moneda.id) LEFT JOIN (tes_cajaori LEFT JOIN (tes_cajaorigendet LEFT JOIN " _
        & " tes_documentos ON tes_cajaorigendet.tipdoc = tes_documentos.id) ON (tes_cajaori.idori = tes_cajaorigendet.idori) AND (tes_cajaori.idtes = tes_cajaorigendet.idtes)) " _
        & " ON tes_caja.id = tes_cajaori.idtes) ON tes_origen.id = tes_cajaori.idori WHERE (((tes_caja.fchreg)=CDate('" & xFchPer & "'))) ORDER BY numreg DESC", xCon
    
    Set Dg1.DataSource = RstMov
    Me.MousePointer = vbDefault
    
End Sub

Sub OpcionesPeriodo()
    Dim NomMes As String
    Dim Cerrado As Boolean
    Dim xFechaMes As String
    Dim xFchIni, xFchFin As Date
    Dim Rpta As Integer
    
    LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
    
    '------------------------------------------------------------------------------------------
    '--bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------


    If mMesActivo <> 0 Then
        xFechaMes = "01/" + Trim(Format(mMesActivo, "00")) + "/" + Trim(Format(Year(Date), "0000"))
        xFchIni = xFechaMes
        xFchFin = Format(HallaDiasMes(CDate(xFechaMes)), "00") + "/" + Mid(xFechaMes, 4, 7)
        LblMes.Caption = Busca_Codigo(mMesActivo, "id", "descripcion", "con_meses", "N", xCon)
        LblPeriodo.Caption = LblMes.Caption
    End If
    
    DoEvents
    
End Sub


Sub Nuevo()
    QueHace = 1
    ActivaTool
    Label5.Caption = "Agregando Egreso"
    Blanquea
    Bloquea
    Agregando = False
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Fg1.Rows = Fg1.Rows + 1
    Fg2.Rows = Fg2.Rows + 1
        
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDKbdMouse
    Fg2.SelectionMode = flexSelectionFree
    Fg1.SelectionMode = flexSelectionFree
    
    lblTipCambio.Caption = ""
    PreparaRST
    PreparaRSTOri
    OptDe2.Value = True
    xHorIni = Time
    TxtFchMov.SetFocus
End Sub

Sub Modificar()
    QueHace = 2
    ActivaTool
    Label5.Caption = "Modificando Egreso"
    Blanquea
    Bloquea
    
    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    Fg1.Rows = Fg1.Rows + 1
    Fg2.Rows = Fg2.Rows + 1
        
    Fg1.Editable = flexEDKbdMouse
    Fg2.Editable = flexEDKbdMouse
    Fg2.SelectionMode = flexSelectionFree
    Fg1.SelectionMode = flexSelectionFree
    
    lblTipCambio.Caption = ""
    'PreparaRST
    'PreparaRSTOri
    MuestraSegundoTab
    OptDe2.Value = True
    xHorIni = Time
    TxtFchMov.SetFocus
End Sub

Sub Bloquea()
    TxtFchMov.Locked = Not TxtFchMov.Locked
    TxtIdMon.Locked = Not TxtIdMon.Locked
    TxtImpDebSol.Locked = Not TxtImpDebSol.Locked
    TxtImpDebDol.Locked = Not TxtImpDebDol.Locked
    TxtImpHabSol.Locked = Not TxtImpHabSol.Locked
    TxtImpHabDol.Locked = Not TxtImpHabDol.Locked
    
    Frame6.Enabled = Not Frame6.Enabled
    Frame10.Enabled = Not Frame10.Enabled
End Sub

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

Sub Blanquea()
    lblReg.Caption = ""
    
    Fg1.Rows = 1
    Fg2.Rows = 1
    TxtFchMov.Valor = ""
    TxtIdMon.Text = ""
    TxtImpDebSol.Text = ""
    TxtImpDebDol.Text = ""
    
    TxtImpDifSol.BackColor = vbWhite
    TxtImpDifDol.BackColor = vbWhite
    
    TxtImpHabSol.Text = ""
    TxtImpHabDol.Text = ""
    
    TxtImpDifSol.Text = ""
    TxtImpDifDol.Text = ""

    TxtGlosa.Text = ""
    
    LblMoneda.Caption = ""
    
    ChkChequeAnulado.Value = 0
    
    mCorrelativo1 = 1
    mCorrelativo2 = 1
    
    ChkVerSaldo.Value = 1
    
End Sub

Private Sub TxtFchMov_Validate(Cancel As Boolean)
    If NulosC(TxtFchMov.Valor) <> "" Then
        lblTipCambio.Caption = BuscaTC(CDate(TxtFchMov.Valor), 2)
        lblTipCambio.Caption = Format(lblTipCambio.Caption, "0.0000")
    End If
End Sub

Function BuscaTC(Fecha As Date, tipo As Integer) As Double
    Dim xRs As New ADODB.Recordset
    'Tipo = 1 compras
    'Tipo = 2 Venta
    
    RST_Busq xRs, "SELECT * FROM con_tc WHERE fecha = CDate('" & Fecha & "') and idmon = 2", xCon
    If xRs.RecordCount <> 0 Then
        xRs.MoveLast
        If tipo = 1 Then BuscaTC = xRs("impcom")
        If tipo = 2 Then BuscaTC = xRs("impven")
    Else
        BuscaTC = 0
    End If
End Function

Private Sub TxtGlosa_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
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
    If TxtIdMon.Text <> "" Then
        LblMoneda.Caption = Busca_Codigo(Val(TxtIdMon.Text), "id", "descripcion", "mae_moneda", "N", xCon)
        If LblMoneda.Caption = "" Then
            TxtIdMon.Text = ""
        Else
''            ActualizarImportesRstTmp
        End If
    End If
End Sub




Sub CargarFacturasPorCanjear(IdProveedor As Integer)
    Dim xForm As New eps_librerias.FormSeleccion
    Dim xRs As New ADODB.Recordset
    Dim xCadWhere1, xCadWhere2 As String
    
    xCadWhere1 = CadWhere(NulosN(Fg1.TextMatrix(Fg1.Row, 3)), 1, 1)
    xCadWhere2 = CadWhere(NulosN(Fg1.TextMatrix(Fg1.Row, 3)), 2, 1)
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    Dim xCampos(7, 5) As String
    
    xCampos(0, 0) = "Nº Documento":  xCampos(0, 1) = "numdoc":         xCampos(0, 2) = "1500":    xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
    xCampos(1, 0) = "T.D.":          xCampos(1, 1) = "codsun":         xCampos(1, 2) = "600":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    xCampos(2, 0) = "Fch. Emi.":     xCampos(2, 1) = "fchdoc":         xCampos(2, 2) = "1000":    xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
    xCampos(3, 0) = "Proveedor":     xCampos(3, 1) = "nombre":         xCampos(3, 2) = "4000":    xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
    xCampos(4, 0) = "Moneda":        xCampos(4, 1) = "simbolo":        xCampos(4, 2) = "800":     xCampos(4, 3) = "C":    xCampos(4, 4) = "N"
    xCampos(5, 0) = "Importe":       xCampos(5, 1) = "imptot":         xCampos(5, 2) = "1200":    xCampos(5, 3) = "N":    xCampos(5, 4) = "N"
    xCampos(6, 0) = "Saldo":         xCampos(6, 1) = "impsal":         xCampos(6, 2) = "1200":    xCampos(6, 3) = "N":    xCampos(6, 4) = "N"
    
    If TxtProv.Text = "" Then
        xForm.SQLCad = "SELECT 0 As xSel, com_compras.id, mae_prov.nombre, mae_documento.codsun, com_compras.fchdoc, com_compras.fchven, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, " _
            & " mae_moneda.simbolo, com_compras.imptot, 'Compras' AS origen, 1 AS idori, com_compras.impsal, com_compras.idmon, com_compras.tipdoc  FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento " _
            & " RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
            & " WHERE (((com_compras.impsal)<>0) AND " & " ( " & xCadWhere1 & "))" _
            & " Union " _
            & " SELECT con_percepcion.id, mae_prov.nombre, mae_documento.codsun, con_percepcion.fchdoc, '' AS fchven, [con_percepcion]![numser]+'-'+[con_percepcion]![numdoc] AS numdoc, " _
            & " mae_moneda.simbolo, con_percepcion.imptotper AS imptot, 'Percepcion' AS origen, 2 AS idori, con_percepcion.impsal, con_percepcion.idmon, con_percepcion.tipdoc FROM ((con_percepcion LEFT JOIN mae_documento " _
            & " ON con_percepcion.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON con_percepcion.idmon = mae_moneda.id) LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id " _
            & " WHERE (((con_percepcion.impsal)<>0))" _
            & " UNION " _
            & " SELECT con_recibos.id, mae_prov.nombre, mae_documento.codsun, con_recibos.fchemi, '' AS fchven, [con_recibos]![serdoc]+'-'+[con_recibos]![numdoc] AS numdoc, " _
            & " mae_moneda.simbolo, con_recibos.impdoc, 'documentos caja' AS origen, 3 AS idori, con_recibos.impdoc AS impsal, con_recibos.idmon, con_recibos.tipdoc " _
            & " FROM ((con_recibos LEFT JOIN mae_prov ON con_recibos.idcli = mae_prov.id) LEFT JOIN mae_documento ON con_recibos.tipdoc = mae_documento.id) " _
            & " LEFT JOIN mae_moneda ON con_recibos.idmon = mae_moneda.id WHERE (((con_recibos.tipmov)=2))"

    Else
        xForm.SQLCad = "SELECT 0 As xSel, com_compras.id, mae_prov.nombre, mae_documento.codsun, com_compras.fchdoc, com_compras.fchven, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, " _
            & " mae_moneda.simbolo, com_compras.imptot, 'Compras' AS origen, 1 AS idori, com_compras.impsal, com_compras.idmon, com_compras.tipdoc FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN " _
            & " (mae_documento RIGHT JOIN com_compras ON mae_documento.id = com_compras.tipdoc) ON mae_moneda.id = com_compras.idmon) ON mae_prov.id = com_compras.idpro " _
            & " WHERE (((com_compras.impsal)<>0) AND ((com_compras.idpro)=" & NulosN(LblIdCliente.Caption) & ") AND " & " ( " & xCadWhere1 & "))" _
            & " Union " _
            & " SELECT con_percepcion.id, mae_prov.nombre, mae_documento.codsun, con_percepcion.fchdoc, '' AS fchven, [con_percepcion]![numser]+'-'+[con_percepcion]![numdoc] AS numdoc, " _
            & " mae_moneda.simbolo, con_percepcion.imptotper AS imptot, 'Percepcion' AS origen, 2 AS idori, con_percepcion.impsal, con_percepcion.idmon, con_percepcion.tipdoc FROM ((con_percepcion LEFT JOIN " _
            & " mae_documento ON con_percepcion.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON con_percepcion.idmon = mae_moneda.id) LEFT JOIN mae_prov " _
            & " ON con_percepcion.idcli = mae_prov.id Where (((con_percepcion.impsal) <> 0) And ((con_percepcion.idcli) = " & NulosN(LblIdCliente.Caption) & "))" _
            & " UNION " _
            & " SELECT con_recibos.id, mae_prov.nombre, mae_documento.codsun, con_recibos.fchemi, '' AS fchven, [con_recibos]![serdoc]+'-'+[con_recibos]![numdoc] AS numdoc, " _
            & " mae_moneda.simbolo, con_recibos.impdoc AS imptot, 'Recibo Caja' AS origen, 3 AS idori, con_recibos.impsal, con_recibos.idmon, con_recibos.tipdoc " _
            & " FROM ((con_recibos LEFT JOIN mae_documento ON con_recibos.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON con_recibos.idmon = mae_moneda.id) " _
            & " LEFT JOIN mae_prov ON con_recibos.idcli = mae_prov.id WHERE (((con_recibos.impsal)<>0))"
    End If
    
    xForm.Titulo = "Buscando Documentos de Proveedores"
    Set xForm.Coneccion = xCon
    Set xRs = Nothing
    Set xRs = xForm.Seleccionar(xCampos)
    If xRs.State = 1 Then
        Dim A As Integer
        Dim xFila As Integer
        xFila = Fg6.Rows - 1
        
        If xRs.RecordCount <> 0 Then
            xRs.MoveFirst
            For A = 1 To xRs.RecordCount
            
                If ExisteDocumento(Fg1.TextMatrix(Fg1.Row, 3), xRs("id")) = False Then
                    Fg6.Rows = Fg6.Rows + 1
                    xFila = xFila + 1
                    Fg6.TextMatrix(xFila, 1) = xRs("nombre")
                    Fg6.TextMatrix(xFila, 2) = xRs("codsun")
                    Fg6.TextMatrix(xFila, 3) = xRs("fchdoc")
                    Fg6.TextMatrix(xFila, 4) = xRs("simbolo")
                    Fg6.TextMatrix(xFila, 5) = xRs("numdoc")
                    Fg6.TextMatrix(xFila, 6) = Format(xRs("imptot"), "0.00")
                    Fg6.TextMatrix(xFila, 7) = Format(xRs("impsal"), "0.00")
                    
                    Fg6.TextMatrix(xFila, 11) = Fg1.TextMatrix(Fg1.Row, 3)
                    Fg6.TextMatrix(xFila, 12) = xRs("id")
                    Fg6.TextMatrix(xFila, 13) = xRs("idmon")
                    Fg6.TextMatrix(xFila, 14) = xRs("tipdoc")
                    
                    If NulosN(xRs("idmon")) <> NulosN(TxtIdMon.Text) Then
                        If NulosN(TxtIdMon.Text) = 1 Then
                            Fg6.TextMatrix(xFila, 8) = xRs("impsal") * NulosN(Fg1.TextMatrix(Fg1.Row, 2))
                        Else
                            If NulosN(Fg1.TextMatrix(Fg1.Row, 2)) <> 0 Then
                                Fg6.TextMatrix(xFila, 8) = xRs("impsal") / NulosN(Fg1.TextMatrix(Fg1.Row, 2))
                            End If
                        End If
                        Fg6.TextMatrix(xFila, 8) = Format(Fg6.TextMatrix(xFila, 8), FORMAT_MONTO)
                    Else
                        Fg6.TextMatrix(xFila, 8) = Format(xRs("impsal"), FORMAT_MONTO)
                    End If
                    
                    RstTmpDocOri.AddNew
                    'agregamos las facturas al recorser temporal
                    RstTmpDocOri("cliente") = xRs("nombre")
                    RstTmpDocOri("tipdoc") = xRs("codsun")
                    RstTmpDocOri("fchemi") = xRs("fchdoc")
                    RstTmpDocOri("moneda") = xRs("simbolo")
                    RstTmpDocOri("numdoc") = xRs("numdoc")
                    RstTmpDocOri("imptot") = xRs("imptot")
                    RstTmpDocOri("impsal") = xRs("impsal")
                    RstTmpDocOri("impsal2") = NulosN(Fg6.TextMatrix(xFila, 8))
                    RstTmpDocOri("idconc") = NulosN(Fg1.TextMatrix(Fg2.Row, 4))
                    RstTmpDocOri("iddocu") = xRs("id")
                    RstTmpDocOri("idmone") = xRs("idmon")
                    RstTmpDocOri("idtipd") = NulosN(xRs("tipdoc"))
                End If
                
                xRs.MoveNext
                If xRs.EOF = True Then Exit For
            Next A
        End If
    End If
End Sub

Function ExisteDocumento(IdConcepto As Integer, IdDocumento As Integer)
    Dim A As Integer
    ExisteDocumento = False
    If RstTMPDoc.RecordCount <> 0 Then
        RstTMPDoc.MoveFirst
        RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND  idconc = " & IdConcepto & " AND iddocu = " & IdDocumento & ""
        
        If RstTMPDoc.RecordCount = 0 Then
            ExisteDocumento = False
        Else
            ExisteDocumento = True
        End If
        RstTMPDoc.Filter = adFilterNone
        RstTMPDoc.Filter = "xtipo=" & mVentanaActiva & " AND idconc = " & IdConcepto & ""
    End If
End Function


Private Sub TxtProv_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusCliente.SetFocus
        CmdBusCliente_Click
    End If
End Sub

Function Grabar() As Boolean
    Dim A As Integer
    
    If IsDate(TxtFchMov.Valor) = False Then
        MsgBox "No ha especificado la fecha de emisión del egreso", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchMov.SetFocus
        Grabar = False
        Exit Function
    End If
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "No ha especificado la moneda de la operación", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Grabar = False
        Exit Function
    End If
    
    'verificamos que al menos haya un concepto en el origen y el egreso
    If Fg1.Rows = 1 Then
        MsgBox "No ha especificado origen para el egreso", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Grabar = False
        Exit Function
    End If

    If Fg2.Rows = 1 And ChkChequeAnulado.Value = 0 Then
        MsgBox "No ha especificado origen para el egreso", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Fg1.SetFocus
        Grabar = False
        Exit Function
    End If
    
    
'    'verificamos que todos los conceptos del origen y destino tengan los datos minimos
    For A = 1 To Fg1.Rows - 1
        If NulosN(Fg1.TextMatrix(A, 3)) = 0 And ChkChequeAnulado.Value = 0 Then
            MsgBox "No ha especificado el origen del egreso", vbInformation + vbOKOnly + vbDefaultButton1
            Grabar = False
            Exit Function
        End If
        If ((NulosN(Fg1.TextMatrix(A, 7) = 0 And NulosN(TxtIdMon.Text) = 1) Or (NulosN(Fg1.TextMatrix(A, 8)) = 0 And NulosN(TxtIdMon.Text) = 2))) And ChkChequeAnulado.Value = 0 Then
            MsgBox "No ha especificado el importe para el origen del egreso", vbInformation + vbOKOnly + vbDefaultButton1
            Grabar = False
            Exit Function
        End If


    Next A

    For A = 1 To Fg2.Rows - 1
        If NulosN(Fg2.TextMatrix(A, 3)) = 0 Then
            MsgBox "No ha especificado el origen del egreso", vbInformation + vbOKOnly + vbDefaultButton1
            Grabar = False
            Exit Function
        End If
        If (NulosN(Fg2.TextMatrix(A, 7) = 0 And NulosN(TxtIdMon.Text) = 1) Or NulosN(Fg2.TextMatrix(A, 8)) = 0 And NulosN(TxtIdMon.Text) = 2) Then
            MsgBox "No ha especificado el importe para el origen del egreso", vbInformation + vbOKOnly + vbDefaultButton1
            Grabar = False
            Exit Function
        End If
    Next A
    
    Dim RstCab As New ADODB.Recordset
    Dim RstDet1 As New ADODB.Recordset
    Dim RstDet1_1 As New ADODB.Recordset
    Dim RstDet2 As New ADODB.Recordset
    Dim RstDet2_2 As New ADODB.Recordset
    'Dim RstDia As New ADODB.Recordset
    Dim rst As New ADODB.Recordset
    Dim xId As Double
    Dim A1&, x&, mCorr&
    Dim xNumAsiento As String
       
    Dim xAcuenta As Double
    Dim xModulo As Integer
    
On Error GoTo LaCague

    xCon.BeginTrans
    Me.MousePointer = vbHourglass
    If QueHace = 1 Then
        'xNumAsiento = NuevoNumAsiento(6, mMesActivo, xCon)
        xId = HallaCodigoTabla("tes_caja", xCon, "id")
        
        RST_Busq RstCab, "SELECT TOP 1 * FROM tes_caja", xCon
        
        RstCab.AddNew
        RstCab("id") = xId
    Else
        xId = RstMov("id")
        'xNumAsiento = DevuelveNumAsiento(6, RstMov("id"), mMesActivo, xCon)
        
        'If xNumAsiento = "" Then xNumAsiento = NuevoNumAsiento(6, mMesActivo, xCon)
        
        '--Actualizar el saldo del documento - origen
        RST_Busq rst, "SELECT * FROM tes_cajaorigendet WHERE idtes = " & xId & " ", xCon
        If rst.RecordCount <> 0 Then
            rst.MoveFirst
            Do While Not rst.EOF
                Select Case NulosN(rst("idmod"))
                    Case 1 '--Compras
                        xCon.Execute "UPDATE com_compras SET com_compras.impsal = [com_compras].[impsal]+ " & NulosN(rst("acuenta")) & " WHERE (((com_compras.id)=" & rst("iddoc") & "))"
                    Case 10 '--reembolsables
                        xCon.Execute "UPDATE com_reembolsables SET com_reembolsables.impsal = [com_reembolsables].[impsal]+ " & NulosN(rst("acuenta")) & " WHERE (((com_reembolsables.id)=" & rst("iddoc") & "))"
                    
                    Case 2 '--ventas
                        xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = [vta_ventas].[impsal]+ " & NulosN(rst("acuenta")) & " WHERE (((vta_ventas.id)=" & rst("iddoc") & "))"
                    Case 4 '--letras
                        xCon.Execute "UPDATE let_letrasdet SET let_letrasdet.impsal = [let_letrasdet].[impsal]+ " & NulosN(rst("acuenta")) & " WHERE (((let_letrasdet.corr)=" & rst("iddoc") & "))"
                    Case 11 '--lgd
                        xCon.Execute "UPDATE vta_gastodebito SET vta_gastodebito.impsal = [vta_gastodebito].[impsal]+ " & NulosN(rst("acuenta")) & " WHERE (((vta_gastodebito.id)=" & rst("iddoc") & "))"
                    Case 19 '--planillas
                        xCon.Execute "UPDATE let_planilla SET let_planilla.impsal = [let_planilla].[impsal]+ " & NulosN(rst("acuenta")) & " WHERE (((let_planilla.id)=" & rst("iddoc") & "))"
                
                End Select
                rst.MoveNext
            Loop
        End If
        xCon.Execute "DELETE * FROM tes_cajaorigendet WHERE idtes = " & xId & " "
        xCon.Execute "DELETE * FROM tes_cajaori WHERE idtes = " & xId & " "
        
        Set rst = Nothing
        
        '--Actualizar el saldo del documento - destino
        RST_Busq rst, "SELECT * FROM tes_cajadestinodet WHERE idtes = " & xId & " ", xCon
        If rst.RecordCount <> 0 Then
            rst.MoveFirst
            Do While Not rst.EOF
                Select Case NulosN(rst("idmod"))
                    Case 1 '--Compras
                        xCon.Execute "UPDATE com_compras SET com_compras.impsal = [com_compras].[impsal]+ " & NulosN(rst("acuenta")) & " WHERE (((com_compras.id)=" & rst("iddoc") & "))"
                    Case 8 '--Boleta de pago
                        xCon.Execute "UPDATE pla_boleta SET pla_boleta.impsal = [pla_boleta].[impsal]+ " & NulosN(rst("acuenta")) & " WHERE (((pla_boleta.id)=" & rst("iddoc") & "))"
                    Case 9 '--Honorarios
                        xCon.Execute "UPDATE com_honorarios SET com_honorarios.impsal = [com_honorarios].[impsal]+ " & NulosN(rst("acuenta")) & " WHERE (((com_honorarios.id)=" & rst("iddoc") & "))"
                    Case 10 '--reembolsables
                        xCon.Execute "UPDATE com_reembolsables SET com_reembolsables.impsal = [com_reembolsables].[impsal]+ " & NulosN(rst("acuenta")) & " WHERE (((com_reembolsables.id)=" & rst("iddoc") & "))"
                    
                    Case 2 '--ventas
                        xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = [vta_ventas].[impsal]- " & NulosN(rst("acuenta")) & " WHERE (((vta_ventas.id)=" & rst("iddoc") & "))"
                    Case 4 '--letras
                        xCon.Execute "UPDATE let_letrasdet SET let_letrasdet.impsal = [let_letrasdet].[impsal]- " & NulosN(rst("acuenta")) & " WHERE (((let_letrasdet.corr)=" & rst("iddoc") & "))"
                    Case 11 '--lgd
                        xCon.Execute "UPDATE vta_gastodebito SET vta_gastodebito.impsal = [vta_gastodebito].[impsal]- " & NulosN(rst("acuenta")) & " WHERE (((vta_gastodebito.id)=" & rst("iddoc") & "))"
                    Case 19 '--planillas
                        xCon.Execute "UPDATE let_planilla SET let_planilla.impsal = [let_planilla].[impsal]- " & NulosN(rst("acuenta")) & " WHERE (((let_planilla.id)=" & rst("iddoc") & "))"
                
                End Select
                rst.MoveNext
            Loop
        End If
        
        xCon.Execute "DELETE * FROM tes_cajadestinodet WHERE idtes = " & xId & " "
        xCon.Execute "DELETE * FROM tes_cajadestino WHERE idtes = " & xId & " "
        
'''        xCon.Execute "DELETE con_diario.* From con_diario WHERE (((con_diario.idmes)=" & mMesActivo & ") AND ((con_diario.idlib)=6) AND ((con_diario.idmov)=" & xId & "))"
        
        RST_Busq RstCab, "SELECT * FROM tes_caja WHERE id = " & xId & "", xCon
    
    End If
    
    '-----------------------------------------------------------------
    RST_Busq RstDet1, "SELECT TOP 1 * FROM tes_cajaori", xCon
    RST_Busq RstDet1_1, "SELECT TOP 1 * FROM tes_cajaorigendet", xCon
    RST_Busq RstDet2, "SELECT TOP 1 * FROM tes_cajadestino", xCon
    RST_Busq RstDet2_2, "SELECT TOP 1 * FROM tes_cajadestinodet", xCon
'    RST_Busq RstDia, "SELECT TOP 1 * FROM con_diario", xCon
    '------------------------------------------------------------------
    mIdRegistro = xId
    '------------------------------------------------------------------
    RstCab("tipmov") = 2
    RstCab("idlib") = 6
    RstCab("fchope") = TxtFchMov.Valor
    RstCab("fchreg") = CDate("01/" + Format(mMesActivo, "00") + "/" + Format(AnoTra, "0000"))
    
'''    RstCab("numreg") = Format(mMesActivo, "00") + xNumAsiento

    RstCab("idmon") = NulosN(TxtIdMon.Text)
    If NulosN(TxtIdMon.Text) = 1 Then
        RstCab("importe") = NulosN(TxtImpDebSol.Text)
    Else
        RstCab("importe") = NulosN(TxtImpDebDol.Text)
    End If
    RstCab("glosa") = NulosC(TxtGlosa.Text)
    
    RstCab.Update

    'grabamos el debe del movimiento =destino del movimiento
    
    mCorr = 1
    For A = 1 To Fg2.Rows - 1
        RstDet2.AddNew
        RstDet2("idtes") = xId
        RstDet2("iddes") = NulosN(Fg2.TextMatrix(A, 3))
        If NulosN(TxtIdMon.Text) = 1 Then
            RstDet2("importe") = NulosN(Fg2.TextMatrix(A, 7))
        Else
            RstDet2("importe") = NulosN(Fg2.TextMatrix(A, 8))
        End If
        RstDet2("tc") = NulosN(Fg2.TextMatrix(A, 2))
        RstDet2("idbcocta") = NulosN(Fg2.TextMatrix(A, 9))
        RstDet2("idmod") = NulosN(Fg2.TextMatrix(A, 5))
        RstDet2.Update
        
        xModulo = NulosN(Fg2.TextMatrix(A, 5))
        
        '--verificando si concepto tiene detalle
        RstTMPDoc.Filter = adFilterNone
        RstTMPDoc.Filter = " xtipo=1 AND idconc = " & NulosN(Fg2.TextMatrix(A, 3)) & ""
        
        Select Case xModulo
            Case 7, 3, 5
                '7=Fondo Fijo; 3=Entregas a Rendir; 5=Anticipos a proveedores
    
                If RstTMPDoc.RecordCount <> 0 Then
                    RstDet2_2.AddNew
                    RstDet2_2("idtes") = xId
                    RstDet2_2("iddes") = NulosN(Fg2.TextMatrix(A, 3))
                    RstDet2_2("idmod") = NulosN(Fg2.TextMatrix(A, 5))
                    RstDet2_2("idper") = NulosN(RstTMPDoc("iddocu"))
                    RstDet2_2("tipdoc") = NulosN(RstTMPDoc("idtipd"))
                    RstDet2_2("numser") = Mid(RstTMPDoc("numdoc"), 1, 4)
                    RstDet2_2("numdoc") = Mid(RstTMPDoc("numdoc"), 6, 10)
                    RstDet2_2("importe") = NulosN(RstTMPDoc("imptot"))
                    '-------------------------------------
                    RstDet2_2("corr") = mCorr
                    mCorr = mCorr + 1
                    '-------------------------------------
                    RstDet2_2.Update
                    
                End If
                
            Case 6 '--Bancos
                If RstTMPDoc.RecordCount <> 0 Then
                    RstDet2_2.AddNew
                    RstDet2_2("idtes") = xId
                    RstDet2_2("iddes") = NulosN(Fg2.TextMatrix(A, 3))
                    RstDet2_2("idmod") = NulosN(Fg2.TextMatrix(A, 5))
                    RstDet2_2("idtipper") = NulosN(RstTMPDoc("idtipper"))
                    RstDet2_2("idper") = NulosN(RstTMPDoc("idper"))
                    RstDet2_2("tipdoc") = NulosN(RstTMPDoc("idtipd"))
                    RstDet2_2("numser") = Mid(RstTMPDoc("numdoc"), 1, 4)
                    RstDet2_2("numdoc") = Mid(RstTMPDoc("numdoc"), 6, 10)
                    RstDet2_2("importe") = NulosN(RstTMPDoc("imptot"))
                    RstDet2_2("acuenta") = NulosN(RstTMPDoc("acuent"))
                    RstDet2_2("glosa") = NulosC(RstTMPDoc("glosa"))
                    If NulosC(RstTMPDoc("fchemi")) = "" Then
                        RstDet2_2("fchdoc") = Null
                    Else
                        RstDet2_2("fchdoc") = NulosC(RstTMPDoc("fchemi"))
                    End If
                    RstDet2_2("idmon") = NulosC(RstTMPDoc("idmone"))
                    '-------------------------------------
                    RstDet2_2("corr") = mCorr
                    mCorr = mCorr + 1
                    '-------------------------------------
                    RstDet2_2.Update
                End If
            
            Case 1, 8, 9, 10, 2, 4, 11, 19
                '--Documentos x Pagar
                '--1=Compras; 8=Boleta Pago; 9=Honorarios; 10=Reembolsables
                
                '--Devoluciones x cobranza en exceso
                '--2=Ventas; 4=Letras; 11=Lgd, 19=Planilla de Letras
            
                If RstTMPDoc.RecordCount <> 0 Then
                    RstTMPDoc.MoveFirst
                    Do While Not RstTMPDoc.EOF
                        RstDet2_2.AddNew
                        RstDet2_2("idtes") = xId
                        RstDet2_2("iddes") = NulosN(RstTMPDoc("idconc"))
                        RstDet2_2("idmod") = NulosN(Fg2.TextMatrix(A, 5))
                        RstDet2_2("iddoc") = NulosN(RstTMPDoc("iddocu"))
                        RstDet2_2("acuenta") = NulosN(RstTMPDoc("acuent"))
                        RstDet2_2("saldo") = NulosN(RstTMPDoc("impsal"))
                        RstDet2_2("idori") = NulosN(RstTMPDoc("idori"))
                        RstDet2_2("corr") = mCorr
                        '--tipo persona
                        If xModulo = 1 Or xModulo = 9 Or xModulo = 10 Then RstDet2_2("idtipper") = 1
                        If xModulo = 2 Or xModulo = 4 Or xModulo = 11 Then RstDet2_2("idtipper") = 2
                        If xModulo = 8 Then RstDet2_2("idtipper") = 3
                        If xModulo = 19 Then RstDet2_2("idtipper") = 2
                        
                        RstDet2_2.Update
                        '-------------------------------------
                        '--Actualizando saldos al documento
                        
                        '--Obteniendo el importe Acuenta
                        If NulosN(TxtIdMon.Text) = NulosN(RstTMPDoc("idmone")) Then
                            xAcuenta = NulosN(RstTMPDoc("acuent"))
                        ElseIf NulosN(TxtIdMon.Text) = 1 And NulosN(RstTMPDoc("idmone")) = 2 Then
                            xAcuenta = NulosN(RstTMPDoc("acuent")) / NulosN(Fg2.TextMatrix(A, 2))
                        ElseIf NulosN(TxtIdMon.Text) = 2 And NulosN(RstTMPDoc("idmone")) = 1 Then
                            xAcuenta = NulosN(RstTMPDoc("acuent")) * NulosN(Fg2.TextMatrix(A, 2))
                        End If
                        
                        Select Case xModulo
                            Case 1 '--compras
                                If RstDet2_2("idori") = 1 Then
                                    xCon.Execute "UPDATE com_compras SET com_compras.impsal = (com_compras.imptot- " & xAcuenta & " ) WHERE (((com_compras.id)=  " & RstTMPDoc("iddocu") & "));"
                                ElseIf RstDet2_2("idori") = 2 Then
                                    xCon.Execute "UPDATE con_percepcion SET con_percepcion.impsal = (con_percepcion.imptotper- " & xAcuenta & " ) WHERE (((con_percepcion.id)=  " & RstTMPDoc("iddocu") & "));"
                                End If
                                
                            Case 8 '--boleta pago
                                xCon.Execute "UPDATE pla_boleta SET pla_boleta.impsal = (pla_boleta.imptot- " & xAcuenta & " ) WHERE (((pla_boleta.id)=  " & RstTMPDoc("iddocu") & "));"
                            
                            Case 9 '--honorarios
                                xCon.Execute "UPDATE com_honorarios SET com_honorarios.impsal = (com_honorarios.imptot- " & xAcuenta & " ) WHERE (((com_honorarios.id)=  " & RstTMPDoc("iddocu") & "));"
                            
                            Case 10 '--Reembolsables
                                xCon.Execute "UPDATE com_reembolsables SET com_reembolsables.impsal = (com_reembolsables.imptot- " & xAcuenta & " ) WHERE (((com_reembolsables.id)=  " & RstTMPDoc("iddocu") & "));"
                            '---------------------------------------------------------------------
                            Case 2 '--ventas
                                If NulosC(RstTMPDoc("tipdoc")) = "NC" Then
                                    xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = (vta_ventas.imptotdoc- " & xAcuenta & " ) WHERE (((vta_ventas.id)=  " & RstTMPDoc("iddocu") & "));"
                                Else
                                    xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = (vta_ventas.imptotdoc+ " & xAcuenta & " ) WHERE (((vta_ventas.id)=  " & RstTMPDoc("iddocu") & "));"
                                End If
                            
                            Case 4 '--Letras
                                xCon.Execute "UPDATE let_letradet SET let_letradet.impsal = (let_letradet.implet+ " & xAcuenta & " ) WHERE (((let_letradet.corr)=  " & RstTMPDoc("iddocu") & "));"
                            
                            Case 11 '--lgd
                                If NulosC(RstTMPDoc("tipdoc")) = "LGC" Then
                                    xCon.Execute "UPDATE vta_gastodebito SET vta_gastodebito.impsal = (vta_gastodebito.imptot- " & xAcuenta & " ) WHERE (((vta_gastodebito.id)=  " & RstTMPDoc("iddocu") & "));"
                                Else
                                    xCon.Execute "UPDATE vta_gastodebito SET vta_gastodebito.impsal = (vta_gastodebito.imptot+ " & xAcuenta & " ) WHERE (((vta_gastodebito.id)=  " & RstTMPDoc("iddocu") & "));"
                                End If
                            
                            Case 19 '--Planilla de Letras
                                xCon.Execute "UPDATE let_planilla SET let_planilla.impsal = (let_planilla.imptot+ " & xAcuenta & " ) WHERE (((let_planilla.id)=  " & RstTMPDoc("iddocu") & "));"
                        
                        End Select
      
                        mCorr = mCorr + 1
                        
                        RstTMPDoc.MoveNext
                    Loop
                
                End If
            
            Case Else
        End Select
        '********************************************************

    Next A
    
    '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
    
    'grabamos el haber del movimiento =origen del movimiento
    mCorr = 1
    
    For A = 1 To Fg1.Rows - 1
        RstDet1.AddNew
        RstDet1("idtes") = xId
        RstDet1("idori") = NulosN(Fg1.TextMatrix(A, 3))
        If NulosN(TxtIdMon.Text) = 1 Then
            RstDet1("importe") = NulosN(Fg1.TextMatrix(A, 7))
        Else
            RstDet1("importe") = NulosN(Fg1.TextMatrix(A, 8))
        End If
        RstDet1("idmod") = NulosN(Fg1.TextMatrix(A, 5))
        RstDet1("tc") = NulosN(Fg1.TextMatrix(A, 2))
        RstDet1("idbcocta") = NulosN(Fg1.TextMatrix(A, 9))
        RstDet1.Update
        
        xModulo = NulosN(Fg1.TextMatrix(A, 5))
        
        If xModulo = 6 Then   'grabamos los datos del cheque
            RstTmpDocOri.Filter = adFilterNone
            RstTmpDocOri.Filter = "idconc = " & NulosN(Fg1.TextMatrix(A, 3)) & ""
            If RstTmpDocOri.RecordCount <> 0 Then
                RstDet1_1.AddNew
                RstDet1_1("idtes") = xId
                RstDet1_1("idori") = NulosN(RstTmpDocOri("idconc"))
                RstDet1_1("idmedpag") = NulosN(RstTmpDocOri("iddocu"))
                RstDet1_1("tipdoc") = NulosN(RstTmpDocOri("idtipd"))
                RstDet1_1("numdoc") = Trim(NulosC(RstTmpDocOri("numdoc")))
                RstDet1_1("importe") = NulosN(RstTmpDocOri("imptot"))
                RstDet1_1("idmod") = 6
                '-------------------------------------
                RstDet1_1("corr") = mCorr
                mCorr = mCorr + 1
                '-------------------------------------
                
                RstDet1_1.Update
            End If
        End If
        
        
        '*********************************************************************************************************************************
        RstTMPDoc.Filter = adFilterNone
        RstTMPDoc.Filter = "xtipo=0 AND idconc = " & NulosN(Fg1.TextMatrix(A, 3)) & ""
        
        Select Case xModulo
        
            Case 1, 2, 4, 10, 11, 19
                '--1=Compras; 10=Reembolsables
                '--Canjes de Documentos
                '--2=Ventas; 4=Letras; 11=Lgd; 19=Planilla de Letras
                If RstTMPDoc.RecordCount <> 0 Then
                    RstTMPDoc.MoveFirst
                    Do While Not RstTMPDoc.EOF
                        RstDet1_1.AddNew
                        RstDet1_1("idtes") = xId
                        RstDet1_1("idori") = NulosN(RstTMPDoc("idconc"))
                        RstDet1_1("idmod") = NulosN(Fg1.TextMatrix(A, 5))
                        RstDet1_1("iddoc") = NulosN(RstTMPDoc("iddocu"))
                        RstDet1_1("acuenta") = NulosN(RstTMPDoc("acuent"))
                        RstDet1_1("saldo") = NulosN(RstTMPDoc("impsal"))
                        RstDet1_1("corr") = mCorr
                        If NulosN(TxtIdMon.Text) = NulosN(RstTMPDoc("idmone")) Then
                            xAcuenta = NulosN(RstTMPDoc("acuent"))
                        ElseIf NulosN(TxtIdMon.Text) = 1 And NulosN(RstTMPDoc("idmone")) = 2 Then
                            xAcuenta = NulosN(RstTMPDoc("acuent")) / NulosN(Fg1.TextMatrix(A, 2))
                        ElseIf NulosN(TxtIdMon.Text) = 2 And NulosN(RstTMPDoc("idmone")) = 1 Then
                            xAcuenta = NulosN(RstTMPDoc("acuent")) * NulosN(Fg1.TextMatrix(A, 2))
                        End If
                        '--tipo persona
                        If xModulo = 1 Or xModulo = 9 Or xModulo = 10 Then RstDet1_1("idtipper") = 1
                        If xModulo = 2 Or xModulo = 4 Or xModulo = 11 Then RstDet1_1("idtipper") = 2
                        If xModulo = 8 Then RstDet1_1("idtipper") = 3
                        If xModulo = 19 Then RstDet1_1("idtipper") = 2
                        RstDet1_1.Update
                        '-------------------------------------
                        Select Case xModulo
                            Case 1 '--compras
                                xCon.Execute "UPDATE com_compras SET com_compras.impsal = ([com_compras].[imptot]- " & xAcuenta & " ) WHERE (((com_compras.id)=  " & RstTMPDoc("iddocu") & "));"
                            Case 10 '--Reembolsables
                                xCon.Execute "UPDATE com_reembolsables SET com_reembolsables.impsal = ([com_reembolsables].[imptot]- " & xAcuenta & " ) WHERE (((com_reembolsables.id)=  " & RstTMPDoc("iddocu") & "));"
                            
                            Case 2 '--Ventas
                                xCon.Execute "UPDATE vta_ventas SET vta_ventas.impsal = (vta_ventas.imptotdoc - " & xAcuenta & " ) WHERE (((vta_ventas.id)=  " & RstTMPDoc("iddocu") & "));"
                            Case 4 '--Letras
                                xCon.Execute "UPDATE let_letradet SET let_letradet.impsal = ([let_letradet].[implet]- " & xAcuenta & " ) WHERE (((let_letradet.corr)=  " & RstTMPDoc("iddocu") & "));"
                            Case 11 '--Lgd
                                xCon.Execute "UPDATE vta_gastodebito SET vta_gastodebito.impsal = (vta_gastodebito.imptot - " & xAcuenta & " ) WHERE (((vta_gastodebito.id)=  " & RstTMPDoc("iddocu") & "));"
                            Case 19 '--Planilla de Letras
                                xCon.Execute "UPDATE let_planilla SET let_planilla.impsal = (let_planilla.imptot - " & xAcuenta & " ) WHERE (((let_planilla.id)=  " & RstTMPDoc("iddocu") & "));"
                        End Select
                        '-------------------------------------
                        mCorr = mCorr + 1
                        RstTMPDoc.MoveNext
                        
                    Loop
                    
                End If
            
            Case 6 '--bancos
                If RstTMPDoc.RecordCount <> 0 Then
                    RstDet1_1.AddNew
                    RstDet1_1("idtes") = xId
                    RstDet1_1("iddes") = NulosN(Fg2.TextMatrix(A, 3))
                    RstDet1_1("idmod") = NulosN(Fg2.TextMatrix(A, 5))
                    RstDet1_1("idtipper") = NulosN(RstTMPDoc("idtipper"))
                    RstDet1_1("idper") = NulosN(RstTMPDoc("idper"))
                    RstDet1_1("tipdoc") = NulosN(RstTMPDoc("idtipd"))
                    RstDet1_1("numser") = Mid(RstTMPDoc("numdoc"), 1, 4)
                    RstDet1_1("numdoc") = Mid(RstTMPDoc("numdoc"), 6, 10)
                    RstDet1_1("importe") = NulosN(RstTMPDoc("imptot"))
                    RstDet1_1("acuenta") = NulosN(RstTMPDoc("acuent"))
                    RstDet1_1("glosa") = NulosC(RstTMPDoc("glosa"))
                    RstDet1_1("corr") = mCorr
                    If NulosC(RstTMPDoc("fchemi")) = "" Then
                        RstDet1_1("fchdoc") = Null
                    Else
                        RstDet1_1("fchdoc") = NulosC(RstTMPDoc("fchemi"))
                    End If
                    RstDet1_1("idmon") = NulosC(RstTMPDoc("idmone"))
                    '-------------------------------------
                    mCorr = mCorr + 1
                    RstDet1_1.Update
                End If
    
            Case Else
        
        End Select
       
    Next A
    
    'grabamos el libro diario
''''    MostrarAsiento False
      
''''    For A = 1 To Fg4.Rows - 1
''''        RstDia.AddNew
''''
''''        RstDia("año") = AnoTra
''''        RstDia("idmes") = mMesActivo
''''        RstDia("idlib") = 6
''''        RstDia("idmov") = xId
''''        RstDia("idcue") = Busca_Codigo(Fg4.TextMatrix(A, 1), "cuenta", "id", "con_planctas", "C", xCon)
''''        RstDia("numasi") = Format(xNumAsiento, "0000")
''''
''''        RstDia("impdebsol") = NulosN(Fg4.TextMatrix(A, 6))
''''        RstDia("imphabsol") = NulosN(Fg4.TextMatrix(A, 7))
''''        RstDia("impdebdol") = NulosN(Fg4.TextMatrix(A, 8))
''''        RstDia("imphabdol") = NulosN(Fg4.TextMatrix(A, 9))
''''        RstDia("idorides") = NulosN(Fg4.TextMatrix(A, 10))
''''        RstDia("idmod") = NulosN(Fg4.TextMatrix(A, 11))
''''        RstDia("iddocpro") = NulosN(Fg4.TextMatrix(A, 12))
''''        RstDia("fchasi") = CDate("01/" + Format(mMesActivo, "00") + "/" + Format(AnoTra, "0000"))
''''        RstDia("fchdoc") = TxtFchMov.Valor
''''        If NulosN(Fg4.TextMatrix(A, 13)) = 0 Then
''''            If NulosN(Fg4.TextMatrix(A, 6)) <> 0 Or NulosN(Fg4.TextMatrix(A, 8)) <> 0 Then
''''                RstDia("tipo") = 2
''''            Else
''''                RstDia("tipo") = 1
''''            End If
''''        Else
''''           RstDia("tipo") = 2
''''        End If
''''        RstDia.Update
''''    Next A
        
    '--generamos es asiento
    xNumAsiento = GenerarAsiento(xCon, 6, CDbl(xId), AnoTra, mMesActivo, 1, 2)
    If xNumAsiento = "" Then GoTo LaCague
    
    '---------------------------------------------------------------------------
    'grabamos el movimiento en la tabla var_edicion
    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId

    xCon.CommitTrans
    Me.MousePointer = vbDefault
    '--
    MsgBox "El movimiento se grabó con éxito" & vbCr & "Registro Nº: " & xNumAsiento, vbInformation, xTitulo
    
    Grabar = True
    
    Set RstCab = Nothing
    Set RstDet1 = Nothing
    Set RstDet1_1 = Nothing
    Set RstDet2 = Nothing
    Set RstDet2_2 = Nothing
'    Set RstDia = Nothing
    Set rst = Nothing
    
    Exit Function

LaCague:
''''    Resume
    Me.MousePointer = vbDefault
    xCon.RollbackTrans
    
    Set RstCab = Nothing
    Set RstDet1 = Nothing
    Set RstDet1_1 = Nothing
    Set RstDet2 = Nothing
    Set RstDet2_2 = Nothing
'    Set RstDia = Nothing
    Set rst = Nothing
    
    Grabar = False
    MsgBox "No se pudo guardar el registro por el siguiente motivo :" + Trim(Err.Description)

End Function

Sub CargarDocumentos(IdProveedor As Integer)
    '===================================================================================================
    'Creado :  xx/xx/xx Por: Enrique Pollongo
    'Propósito: Mostrar listado de documentos para detallar el destino
    '
    'Entradas:  IdProveedor = codigo del proveedor,cliente
    '
    'Resultados:Listado de Documentos para seleccionar
    
    'Modificado : 19/02/10 Por: Johan Castro
    '             Mostrar lista letras para protestar
    '             13/09/11
    '             Agregar campo xsel a consulta de seleccion de documentos reembolsables
    '             17/11/11
    '             Agregar codigo para mostrar selección de registro (ventas, Liquidacion de Gasto Debito, Planilla de Letras)
    '             01/02/12
    '             Agregar variable nSQLWhere,nSQLWhere1 para aplicar filtro a las consultas
    '             Se considera el filtro para mostrar documentos con/sin saldo
    '===================================================================================================
    
    Dim xRs As New ADODB.Recordset
    Dim xCadWhere1, xCadWhere2 As String
    Dim nSQLNotIn As String
    Dim nSQLIdPer As String '--sentencia SQL del proveedor/cliente seleccionado
    Dim nSQL As String
    
    Dim nTitulo As String '--indicar el titulo al seleccionar los documentos
    
    Dim nSQLWhere As String '--aplica filtro a la consulta
    Dim nSQLWhere1 As String '--aplica filtro a la consulta de percepcion
    
''    If mVentanaActiva = 1 Then
''        xCadWhere1 = CadWhere(IdConcepto, 1, 2)
''    Else
''        xCadWhere1 = CadWhere(IdConcepto, 1, 1)
''    End If
    
    'descripcion     'campo     'tamaño     'tipo = Numerico, caracter, fecha
    ReDim xCampos(10, 5) As String
    
  
    xCampos(0, 0) = "Nº Documento":  xCampos(0, 1) = "numdoc":         xCampos(0, 2) = "1500":    xCampos(0, 3) = "C":    xCampos(0, 4) = "S"
    xCampos(1, 0) = "T.D.":          xCampos(1, 1) = "abrev":          xCampos(1, 2) = "600":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
    xCampos(2, 0) = "Dia":           xCampos(2, 1) = "dia":            xCampos(2, 2) = "350":     xCampos(2, 3) = "C":    xCampos(2, 4) = "N"
    xCampos(3, 0) = "Mes":           xCampos(3, 1) = "mes":            xCampos(3, 2) = "400":     xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
    xCampos(4, 0) = "Año":           xCampos(4, 1) = "ano":            xCampos(4, 2) = "400":     xCampos(4, 3) = "C":    xCampos(4, 4) = "N"
    xCampos(5, 0) = "Proveedor":     xCampos(5, 1) = "nombre":         xCampos(5, 2) = "3000":    xCampos(5, 3) = "C":    xCampos(5, 4) = "N"
    xCampos(6, 0) = "Cond. Pago":    xCampos(6, 1) = "desccond":       xCampos(6, 2) = "1000":    xCampos(6, 3) = "C":    xCampos(6, 4) = "N"
    xCampos(7, 0) = "M":             xCampos(7, 1) = "simbolo":        xCampos(7, 2) = "500":     xCampos(7, 3) = "C":    xCampos(7, 4) = "N"
    xCampos(8, 0) = "Importe":       xCampos(8, 1) = "imptot":         xCampos(8, 2) = "1200":    xCampos(8, 3) = "N":    xCampos(8, 4) = "N"
    xCampos(9, 0) = "Saldo":         xCampos(9, 1) = "impsal":         xCampos(9, 2) = "1200":    xCampos(9, 3) = "N":    xCampos(9, 4) = "N"
    
    Select Case IdModulo
    
        Case 1 '--compras
            '------------------------------------------------------------------------
            '--verificar si se mostraran los documentos con saldo
            If ChkVerSaldo.Value = 1 Then
                nSQLWhere = " AND com_compras.impsal<>0 "
                nSQLWhere1 = " AND con_percepcion.impsal<>0"
            End If
            '--verificar si hay proveedor seleccionado
            If IdProveedor <> 0 Then
                nSQLWhere = nSQLWhere & " AND com_compras.idpro = " & IdProveedor
                nSQLWhere1 = nSQLWhere1 & " AND con_percepcion.idcli = " & IdProveedor
            End If
            '--sentencia sql para no mostrar documentos seleccionados
            nSQLWhere = nSQLWhere & GRID_GENERAR_SQL_ID(Fg3, 12, " AND com_compras.id", " NOT IN", True)
            '------------------------------------------------------------------------
            nSQL = "SELECT 0 AS xSel, com_compras.id, mae_prov.nombre, mae_documento.abrev, mae_documento.codsun, com_compras.fchdoc, " _
                & " com_compras.fchdoc & '' AS fchdoc1, com_compras.fchven, [com_compras]![numser]+'-'+[com_compras]![numdoc] AS numdoc, mae_moneda.simbolo, " _
                & " com_compras.imptot, 'Compras' AS origen, 1 AS idori, com_compras.impsal, com_compras.idmon, com_compras.tipdoc, mae_condpago.descripcion AS desccond, " _
                & " Format([fchdoc],'dd') AS dia, Format([fchdoc],'mm') AS mes, Format([fchdoc],'yy') AS ano " _
                & " FROM mae_condpago INNER JOIN (mae_documento RIGHT JOIN (mae_moneda RIGHT JOIN (mae_prov RIGHT JOIN com_compras ON mae_prov.id = com_compras.idpro) " _
                & " ON mae_moneda.id = com_compras.idmon) ON mae_documento.id = com_compras.tipdoc) ON mae_condpago.id = com_compras.idconpag " _
                & " WHERE com_compras.imptot<>0 " & nSQLWhere _
                & " Union " _
                & " SELECT 0 AS xSel, con_percepcion.id, mae_prov.nombre, mae_documento.abrev, mae_documento.codsun, con_percepcion.fchdoc, " _
                & " con_percepcion.fchdoc & '' AS fchdoc1, '' AS fchven, [con_percepcion]![numser]+'-'+[con_percepcion]![numdoc] AS numdoc, " _
                & " mae_moneda.simbolo, con_percepcion.imptotper AS imptot, 'Percepcion' AS origen, 2 AS idori, con_percepcion.impsal, con_percepcion.idmon, " _
                & " con_percepcion.tipdoc, '' AS desccond, Format([fchdoc],'dd') AS dia, Format([fchdoc],'mm') AS mes, Format([fchdoc],'yy') AS ano " _
                & " FROM ((con_percepcion LEFT JOIN mae_documento ON con_percepcion.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON con_percepcion.idmon = mae_moneda.id) " _
                & " LEFT JOIN mae_prov ON con_percepcion.idcli = mae_prov.id " _
                & " Where con_percepcion.imptotper<>0 " & nSQLWhere1
                    
            '--aplicar orden
            nSQL = "Select vw.* from (" & nSQL & ") as vw order by vw.nombre, vw.numdoc"
            
            nTitulo = "Buscando Documentos de Proveedores"
            
        '-------
        Case 2  '--ventas
            '------------------------------------------------------------------------
            '--verificar si se mostraran los documentos con saldo
            If ChkVerSaldo.Value = 1 Then nSQLWhere = " AND vta_ventas.impsal<>0 "
            '--verificar si hay cliente seleccionado
            If IdProveedor <> 0 Then nSQLWhere = nSQLWhere & " AND vta_ventas.idcli = " & IdProveedor
            '--sentencia sql para no mostrar documentos seleccionados
            nSQLWhere = nSQLWhere & GRID_GENERAR_SQL_ID(Fg3, 12, " AND vta_ventas.id", " NOT IN", True)
            '------------------------------------------------------------------------
        
            nSQL = "SELECT 0 AS xSel, vta_ventas.id, mae_cliente.nombre, mae_documento.abrev, mae_documento.codsun, vta_ventas.fchdoc,vta_ventas.fchdoc & '' as fchdoc1, vta_ventas.fchven, " _
                & " vta_ventas.numser+'-'+vta_ventas.numdoc AS numdoc, mae_moneda.simbolo, vta_ventas.imptotdoc AS imptot, 'Ventas' AS origen, 2 AS idori, " _
                & " vta_ventas.impsal, vta_ventas.tipdoc, vta_ventas.idmon, mae_condpago.abrev AS desccond, Format(fchdoc,'dd') AS dia, Format(fchdoc,'mm') AS mes, Format(fchdoc,'yy') AS ano" _
                & " FROM (((vta_ventas LEFT JOIN mae_documento ON vta_ventas.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON vta_ventas.idmon = mae_moneda.id) " _
                & " LEFT JOIN mae_cliente ON vta_ventas.idcli = mae_cliente.id) LEFT JOIN mae_condpago ON vta_ventas.idconpag = mae_condpago.id " _
                & " Where vta_ventas.imptotdoc <> 0 " & nSQLWhere _
                & " ORDER BY vta_ventas.numser+'-'+vta_ventas.numdoc"
            
            nTitulo = "Buscando Documentos del Ventas"
            xCampos(5, 0) = "Cliente":
            
            
        Case 4 '--letras
            '------------------------------------------------------------------------
            '--verificar si se mostraran los documentos con saldo
            If ChkVerSaldo.Value = 1 Then nSQLWhere = " AND let_letradet.impsal<>0 "
            '--verficar si hay cliente seleccionado
            If IdProveedor <> 0 Then nSQLWhere = nSQLWhere & " AND let_letra.idclipro = " & IdProveedor
            '--sentencia sql para no mostrar documentos seleccionados
            nSQLWhere = nSQLWhere & GRID_GENERAR_SQL_ID(Fg3, 12, " AND let_letradet.corr", " NOT IN", True)
            '------------------------------------------------------------------------
            
            '--se mostraran todas las letras
                
            nSQL = "SELECT 0 as xSel, let_letradet.corr AS id, mae_cliente.nombre, mae_documento.abrev, mae_documento.codsun, let_letradet.fchemi AS fchdoc, " _
                + vbCr + " let_letradet.fchemi & '' as fchdoc1, let_letradet.fchven, [let_letra].[ano] & ' ' & [let_letradet].[numdoc] & ' ' & [let_letradet].[numser] AS numdoc, " _
                + vbCr + " mae_moneda.simbolo, let_letradet.implet AS imptot, 'LETRA' AS origen, 0 AS idori, let_letradet.impsal, let_letra.tipdoc, let_letra.idmon, '' AS desccond, Format([fchdoc],'dd') AS dia, Format([fchdoc],'mm') AS mes, Format([fchdoc],'yy') AS ano  " _
                + vbCr + " FROM mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN ((let_letra LEFT JOIN mae_cliente ON let_letra.idclipro = mae_cliente.id) RIGHT JOIN let_letradet ON let_letra.id = let_letradet.idlet) ON mae_documento.id = let_letra.tipdoc) ON mae_moneda.id = let_letra.idmon " _
                + vbCr + " WHERE let_letradet.impsal<>0 " & nSQLWhere
                        
            nTitulo = "Buscando Letras del Cliente"
            
            '--dar ancho a la presentacion
            xCampos(5, 0) = "Cliente":
            xCampos(0, 2) = "1800": '--numerodoc
            xCampos(5, 2) = "2500": '--cliente
                        
        Case 8 '--planillas
            '------------------------------------------------------------------------
            '--verificar si se mostraran los documentos con saldo
            If ChkVerSaldo.Value = 1 Then nSQLWhere = " AND pla_boleta.impsal<>0 "
            '--verficar si hay empleado seleccionado
            If IdProveedor <> 0 Then nSQLWhere = nSQLWhere & " AND pla_boleta.idemp = " & IdProveedor
            '--sentencia sql para no mostrar documentos seleccionados
            nSQLWhere = nSQLWhere & GRID_GENERAR_SQL_ID(Fg3, 12, " AND pla_boleta.id", " NOT IN", True)
            '------------------------------------------------------------------------

            nSQL = "SELECT 0 as xSel, pla_boleta.id, con_meses.descripcion, pla_empleados.nombre, " _
                & " pla_boleta.ano, mae_documento.abrev, mae_moneda.simbolo, pla_boleta.imptot, pla_boleta.impsal, [pla_boleta]![numser]+'-'+[pla_boleta]![numdoc] AS numdoc, " _
                & " pla_boleta.fchdoc,pla_boleta.fchdoc & '' as fchdoc1, pla_boleta.iddoc as tipdoc, pla_boleta.idmon,0 as idori, '' AS desccond, Format([fchdoc],'dd') AS dia, Format([fchdoc],'mm') AS mes, Format([fchdoc],'yy') AS ano  " _
                & " FROM pla_empleados RIGHT JOIN (((pla_boleta LEFT JOIN con_meses ON pla_boleta.idmes = con_meses.id) " _
                & " LEFT JOIN mae_documento ON pla_boleta.iddoc = mae_documento.id) LEFT JOIN mae_moneda ON pla_boleta.idmon = mae_moneda.id) ON pla_empleados.id = pla_boleta.idemp " _
                & " Where pla_boleta.imptot <> 0 " & nSQLWhere _
                & " ORDER BY con_meses.descripcion, pla_empleados.nombre"
            
            nTitulo = "Buscando Documentos Boletas de Pago"
            
            xCampos(3, 0) = "Trabajador / Empleado":
            
        Case 9 '--honorarios
            '------------------------------------------------------------------------
            '--verificar si se mostraran los documentos con saldo
            If ChkVerSaldo.Value = 1 Then nSQLWhere = " AND com_honorarios.impsal<>0 "
            '--verificar si hay proveedor seleccionado
            If IdProveedor <> 0 Then nSQLWhere = nSQLWhere & " AND com_honorarios.idpro = " & IdProveedor
            '--sentencia sql para no mostrar documentos seleccionados
            nSQLWhere = nSQLWhere & GRID_GENERAR_SQL_ID(Fg3, 12, " AND com_honorarios.id", " NOT IN", True)
            '------------------------------------------------------------------------
            
            nSQL = "SELECT 0 as xSel, com_honorarios.id, mae_prov.nombre, mae_documento.abrev,mae_documento.codsun, com_honorarios.fchdoc, " _
                + vbCr + " com_honorarios.fchdoc & '' as fchdoc1, com_honorarios.fchven, IIf([com_honorarios]![numser] Is Null,'', " _
                + vbCr + " [com_honorarios].[numser] & '-') & [com_honorarios].[numdoc]  AS numdoc, mae_moneda.simbolo, com_honorarios.imptot, " _
                + vbCr + " 'Compras' AS origen, 0 AS idori, com_honorarios.impsal, com_honorarios.idmon, com_honorarios.tipdoc, '' AS desccond, Format([fchdoc],'dd') AS dia, Format([fchdoc],'mm') AS mes, Format([fchdoc],'yy') AS ano  " _
                + vbCr + " FROM mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN (mae_documento RIGHT JOIN com_honorarios ON mae_documento.id = com_honorarios.tipdoc) ON mae_moneda.id = com_honorarios.idmon) ON mae_prov.id = com_honorarios.idpro " _
                + vbCr + " WHERE com_honorarios.imptot<>0 " & nSQLWhere
            
            '--aplicar orden
            nSQL = "Select vw.* from (" & nSQL & ") as vw order by vw.nombre, vw.numdoc "
            
            nTitulo = "Buscando Documentos de Prestadores de Servicio"
            
            '--
            xCampos(3, 0) = "Prestador de Servicio":
            
        Case 10 '--reembolsables
            '------------------------------------------------------------------------
            '--verificar si se mostraran los documentos con saldo
            If ChkVerSaldo.Value = 1 Then nSQLWhere = " AND com_reembolsables.impsal<>0 "
            '--verificar si hay proveedor seleccionado
            If IdProveedor <> 0 Then nSQLWhere = nSQLWhere & " AND com_reembolsables.idpro = " & IdProveedor
            '--sentencia sql para no mostrar documentos seleccionados
            nSQLWhere = nSQLWhere & GRID_GENERAR_SQL_ID(Fg3, 12, " AND com_reembolsables.id", " NOT IN", True)
            '------------------------------------------------------------------------
            
            nSQL = "SELECT 0 as xsel ,com_reembolsables.id, mae_prov.nombre,mae_documento.abrev, mae_documento.codsun, com_reembolsables.fchdoc, " _
                + vbCr + " com_reembolsables.fchdoc & '' as fchdoc1, com_reembolsables.fchven, IIf([com_reembolsables]![numser] Is Null,'', " _
                + vbCr + " [com_reembolsables]![numser] & '-') & [com_reembolsables]![numdoc] AS numdoc, mae_moneda.simbolo, com_reembolsables.imptot, " _
                + vbCr + " 'Reembolsables' AS origen, 0 AS idori, com_reembolsables.impsal, com_reembolsables.idmon, com_reembolsables.tipdoc AS tipdoc, '' AS desccond, Format([fchdoc],'dd') AS dia, Format([fchdoc],'mm') AS mes, Format([fchdoc],'yy') AS ano  " _
                + vbCr + " FROM (mae_prov RIGHT JOIN (mae_moneda RIGHT JOIN com_reembolsables ON mae_moneda.id = com_reembolsables.idmon) ON mae_prov.id = com_reembolsables.idpro) LEFT JOIN mae_documento ON com_reembolsables.tipdoc = mae_documento.id " _
                + vbCr + " WHERE com_reembolsables.imptot<>0 " & nSQLWhere
                    
            '--aplicar orden
            nSQL = "Select vw.* from (" & nSQL & ") as vw order by vw.nombre, vw.numdoc "
            
            nTitulo = "Buscando Documentos de Proveedores"
            
    
        Case 11  '--Liquidacion de Gasto Debito
            '------------------------------------------------------------------------
            '--verificar si se mostraran los documentos con saldo
            If ChkVerSaldo.Value = 1 Then nSQLWhere = " AND vta_gastodebito.impsal<>0 "
            '--verificar si hay cliente seleccionado
            If IdProveedor <> 0 Then nSQLWhere = nSQLWhere & " AND vta_gastodebito.idcli = " & IdProveedor
            '--sentencia sql para no mostrar documentos seleccionados
            nSQLWhere = nSQLWhere & GRID_GENERAR_SQL_ID(Fg3, 12, " AND vta_gastodebito.id", " NOT IN", True)
            '------------------------------------------------------------------------
        
            nSQL = "SELECT 0 AS xSel,  vta_gastodebito.id, mae_cliente.nombre,  mae_documento.abrev,mae_documento.codsun, vta_gastodebito.fchemi AS fchdoc, vta_gastodebito.fchven, [vta_gastodebito]![numser]+'-'+[vta_gastodebito]![numdoc] AS numdoc, " _
                & " mae_moneda.simbolo, vta_gastodebito.imptot, 'LGD' AS origen, 2 AS idori, vta_gastodebito.impsal, vta_gastodebito.tipdoc, vta_gastodebito.idmon, '' AS desccond, Format([fchemi],'dd') AS dia, Format([fchemi],'mm') AS mes, Format([fchemi],'yy') AS ano " _
                & " FROM ((vta_gastodebito LEFT JOIN mae_documento ON vta_gastodebito.tipdoc = mae_documento.id) LEFT JOIN mae_cliente ON vta_gastodebito.idcli = mae_cliente.id) " _
                & " LEFT JOIN mae_moneda ON vta_gastodebito.idmon = mae_moneda.id " _
                & " Where vta_gastodebito.imptot <> 0 " & nSQLWhere _
                & " ORDER BY [vta_gastodebito]![numser] +'-'+ [vta_gastodebito]![numdoc];"
            
            nTitulo = "Buscando Documentos LGD"
        
       Case 19  '--Planilla de letras
        
            xCampos(0, 0) = "Fch. Emi":       xCampos(0, 1) = "fchdoc":        xCampos(0, 2) = "1000":     xCampos(0, 3) = "C":    xCampos(0, 4) = "N"
            xCampos(1, 0) = "T.D.":           xCampos(1, 1) = "abrev":         xCampos(1, 2) = "900":     xCampos(1, 3) = "C":    xCampos(1, 4) = "N"
            xCampos(2, 0) = "Nº Planilla":    xCampos(2, 1) = "numdoc":        xCampos(2, 2) = "1000":    xCampos(2, 3) = "C":    xCampos(2, 4) = "S"
            xCampos(3, 0) = "Moneda":         xCampos(3, 1) = "simbolo":       xCampos(3, 2) = "900":     xCampos(3, 3) = "C":    xCampos(3, 4) = "N"
            xCampos(4, 0) = "Nº Letras Abo.": xCampos(4, 1) = "numlet":        xCampos(4, 2) = "1200":    xCampos(4, 3) = "C":    xCampos(4, 4) = "N"
            xCampos(5, 0) = "Imp. Planilla":  xCampos(5, 1) = "imptot":        xCampos(5, 2) = "1200":    xCampos(5, 3) = "N":    xCampos(5, 4) = "N"
            
            '------------------------------------------------------------------------
            '--verificar si se mostraran los documentos con saldo
            If ChkVerSaldo.Value = 1 Then nSQLWhere = " AND let_planilla.impsal<>0 "
            '--verificar si hay cuenta de banco seleccionado
            If IdProveedor <> 0 Then nSQLWhere = nSQLWhere & " AND let_planilla.idbcocta = " & IdProveedor
            '--sentencia sql para no mostrar documentos seleccionados
            nSQLWhere = nSQLWhere & GRID_GENERAR_SQL_ID(Fg3, 12, " AND let_planilla.id", " NOT IN", True)
            '------------------------------------------------------------------------
            nSQL = "SELECT 0 AS xSel, let_planilla.id, Trim([mae_bancos].[descripcion]) & ' N° Cta. ' & Trim([mae_banconumcta].[numcue]) AS nombre, mae_documento.abrev, mae_documento.codsun, let_planilla.fchemi & '' AS fchdoc , Null AS fchvenc, let_planilla.numdoc, mae_moneda.simbolo, let_planilla.numlet, let_planilla.imptot, 'PLA LET' AS origen, 19 AS idori, let_planilla.impsal, let_planilla.tipdoc, let_planilla.idmon, '' AS desccond, Format([fchemi],'dd') AS dia, Format([fchemi],'mm') AS mes, Format([fchemi],'yy') AS ano  " _
                + vbCr + " FROM mae_bancos RIGHT JOIN (mae_banconumcta RIGHT JOIN ((let_planilla LEFT JOIN mae_moneda ON let_planilla.idmon = mae_moneda.id) LEFT JOIN mae_documento ON let_planilla.tipdoc = mae_documento.id) ON mae_banconumcta.id = let_planilla.idbcocta) ON mae_bancos.id = mae_banconumcta.idban " _
                + vbCr + " WHERE let_planilla.imptot<>0 " & nSQLWhere
            
            nTitulo = "Buscando Planilla de Letras"
            
            
        Case 20 '--orden de despacho
            '------------------------------------------------------------------------
            '--verificar si se mostraran los documentos con saldo
'            If ChkVerSaldo.Value = 1 Then nSQLWhere = " AND var_ordendespacho.impsal<>0 "
            '--verificar si hay cliente seleccionado
            If IdProveedor <> 0 Then nSQLWhere = nSQLWhere & " AND var_ordendespacho.idcli = " & IdProveedor
            '--sentencia sql para no mostrar documentos seleccionados
            nSQLWhere = nSQLWhere & GRID_GENERAR_SQL_ID(Fg3, 12, " AND var_ordendespacho.id", " NOT IN", True)
            '------------------------------------------------------------------------
                
            nSQL = "SELECT 0 as xsel, var_ordendespacho.id, mae_cliente.nombre, mae_documento.abrev, var_ordendespacho.fchdoc, [var_ordendespacho].[fchdoc] & '' AS fchdoc1, " _
                + vbCr + " var_ordendespacho.fchven, var_ordendespacho.numdoc, '' AS simbolo, 0 AS imptot, 'Orden Despacho' AS origen, 0 AS idori, 0 AS impsal, " _
                + vbCr + " 0 AS idmon, var_ordendespacho.tipdoc, var_ordendespacho.imptot1, var_ordendespacho.imptot2, '' AS desccond , Format([fchdoc],'dd') AS dia, Format([fchdoc],'mm') AS mes, Format([fchdoc],'yy') AS ano  " _
                + vbCr + " FROM (var_ordendespacho INNER JOIN mae_cliente ON var_ordendespacho.idcli = mae_cliente.id) INNER JOIN mae_documento ON var_ordendespacho.tipdoc = mae_documento.id " _
                + vbCr + " WHERE var_ordendespacho.numerodoc is not null " & nSQLWhere _
                + vbCr + " ORDER BY var_ordendespacho.numdoc  "

            '--ocultando columnas
            xCampos(4, 2) = "0":
            xCampos(5, 2) = "0":
            xCampos(6, 2) = "0":


        Case Else
        
            Exit Sub
    End Select
    
    
    CARGAR_DLL_EPSBUSCAR_SEL xCon, xRs, nSQL, xCampos(), nTitulo
    
    Agregando = True

    Dim A As Integer
    Dim xFila As Integer
    xFila = Fg3.Rows - 1
    If xRs.State = 1 Then
        If xRs.RecordCount <> 0 Then
            xRs.MoveFirst
            For A = 1 To xRs.RecordCount
    
                Fg3.Rows = Fg3.Rows + 1
                Fg3.Row = Fg3.Rows - 1
                xFila = xFila + 1
                
                Fg3.TextMatrix(xFila, 1) = NulosC(xRs("nombre"))
                Fg3.TextMatrix(xFila, 2) = NulosC(xRs("abrev"))
                Fg3.TextMatrix(xFila, 3) = NulosC(xRs("fchdoc"))
                Fg3.TextMatrix(xFila, 4) = NulosC(xRs("simbolo"))
                Fg3.TextMatrix(xFila, 5) = NulosC(xRs("numdoc"))
                
                Fg3.TextMatrix(xFila, 6) = Format(NulosN(xRs("imptot")), "0.00")
                Fg3.TextMatrix(xFila, 7) = Format(NulosN(xRs("impsal")), "0.00")
                
                Fg3.TextMatrix(xFila, 11) = IdConcepto
                Fg3.TextMatrix(xFila, 12) = NulosN(xRs("id"))
                Fg3.TextMatrix(xFila, 13) = NulosN(xRs("idmon"))
                Fg3.TextMatrix(xFila, 14) = NulosN(xRs("tipdoc"))
                
                Fg3.TextMatrix(xFila, 15) = mCorrelativo2
                
                If NulosN(xRs("idmon")) <> NulosN(TxtIdMon.Text) Then
                    If NulosN(TxtIdMon.Text) = 1 Then
                        Fg3.TextMatrix(xFila, 8) = NulosN(xRs("impsal")) * TipCam
                    Else
                        If TipCam <> 0 Then
                            Fg3.TextMatrix(xFila, 8) = NulosN(xRs("impsal")) / TipCam
                        End If
                    End If
                    Fg3.TextMatrix(xFila, 8) = Format(Fg3.TextMatrix(xFila, 8), FORMAT_MONTO)
                Else
                    Fg3.TextMatrix(xFila, 8) = Format(NulosN(xRs("impsal")), FORMAT_MONTO)
                End If
                
                '---
                Fg3.TextMatrix(xFila, 9) = Format(Fg3.TextMatrix(xFila, 8), "0.00")
                
                Fg3.TextMatrix(xFila, 10) = NulosN(Fg3.TextMatrix(xFila, 8)) - NulosN(Fg3.TextMatrix(xFila, 9))
                
                '---
                RstTMPDoc.AddNew
                'agregamos las facturas al recorser temporal
                RstTMPDoc("cliente") = NulosC(xRs("nombre"))
                RstTMPDoc("tipdoc") = NulosC(xRs("abrev"))
                RstTMPDoc("fchemi") = NulosC(xRs("fchdoc"))
                RstTMPDoc("moneda") = NulosC(xRs("simbolo"))
                RstTMPDoc("numdoc") = NulosC(xRs("numdoc"))
                RstTMPDoc("imptot") = NulosN(xRs("imptot"))
                RstTMPDoc("impsal") = NulosN(xRs("impsal"))
                RstTMPDoc("impsal2") = NulosN(Fg3.TextMatrix(xFila, 8))
                RstTMPDoc("idconc") = IdConcepto
                RstTMPDoc("iddocu") = NulosN(xRs("id"))
                RstTMPDoc("idmone") = NulosN(xRs("idmon"))
                RstTMPDoc("idtipd") = NulosN(xRs("tipdoc"))
                RstTMPDoc("idori") = NulosN(xRs("idori"))
                
                RstTMPDoc("acuent") = Fg3.TextMatrix(xFila, 9)
                RstTMPDoc("newsal") = Fg3.TextMatrix(xFila, 10)
        
                RstTMPDoc("corr") = mCorrelativo2
                
                RstTMPDoc("xtipo") = mVentanaActiva
                
                mCorrelativo2 = mCorrelativo2 + 1
                    
                xRs.MoveNext
                If xRs.EOF = True Then Exit For
            Next A
        End If
    End If
        
    Agregando = False
    
    Set xRs = Nothing
    TotalizarFG3
    
End Sub


Private Sub Frame3_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    OrigFX = x
    OrigFY = y
    Frame3.ZOrder 0
End Sub

Private Sub Frame3_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then
        With Frame3
            .Move .Left + x - OrigFX, .Top + y - OrigFY
        End With
    End If
End Sub

Private Sub Frame9_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    OrigFX = x
    OrigFY = y
    Frame9.ZOrder 0
End Sub

Private Sub Frame9_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then
        With Frame9
            .Move .Left + x - OrigFX, .Top + y - OrigFY
        End With
    End If
End Sub

Private Sub Frame12_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    OrigFX = x
    OrigFY = y
    Frame12.ZOrder 0
End Sub

Private Sub Frame12_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 0 Then
        With Frame12
            .Move .Left + x - OrigFX, .Top + y - OrigFY
        End With
    End If
End Sub


Private Sub CmdExportar_Click()

    If Fg3.Rows = Fg3.FixedRows Then Exit Sub
    Dim xfrm As New SGI2_funciones.formularios
    xfrm.VSFlexGrid_Exportar_MSExcel xCon, Fg3, lblReg.Caption, Fg2.TextMatrix(Fg2.Row, 1) & " - " & LblTc.Caption, TxtGlosa.Text, "Detalle de Cargos Bancarios"
    Set xfrm = Nothing

End Sub


Private Sub CmdExportar1_Click()

    If Fg4.Rows = Fg4.FixedRows Then Exit Sub
    Dim xfrm As New SGI2_funciones.formularios
    xfrm.VSFlexGrid_Exportar_MSExcel xCon, Fg4, lblReg.Caption, LblMoneda.Caption, TxtGlosa.Text, "Asiento Contable"
    Set xfrm = Nothing

End Sub

