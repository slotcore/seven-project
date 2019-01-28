VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Object = "{DEF7CADD-83C0-11D0-A0F1-00A024703500}#7.0#0"; "todg7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form FrmDocumentos 
   Caption         =   "Transferencia - Documentos"
   ClientHeight    =   8145
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13875
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   13875
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame FraDetalle 
      BorderStyle     =   0  'None
      Caption         =   "[ Depurar Datos ]"
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
      Height          =   4845
      Left            =   12960
      TabIndex        =   114
      Top             =   330
      Visible         =   0   'False
      Width           =   9495
      Begin VB.PictureBox pic1 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   9270
         Picture         =   "FrmDocumentos.frx":0000
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   117
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   4710
         TabIndex        =   116
         Top             =   4350
         Width           =   1755
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exportar MsExcel"
         Height          =   375
         Left            =   2460
         TabIndex        =   115
         Top             =   4350
         Width           =   1755
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg7 
         Height          =   3885
         Left            =   60
         TabIndex        =   119
         Top             =   360
         Width           =   9345
         _cx             =   16484
         _cy             =   6853
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
         BackColorSel    =   4210816
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483627
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
         Rows            =   50
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmDocumentos.frx":02EC
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
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   6
         X1              =   -30
         X2              =   11970
         Y1              =   0
         Y2              =   15
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   5
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   6500
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   4
         X1              =   9480
         X2              =   9480
         Y1              =   30
         Y2              =   7000
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   5
         X1              =   0
         X2              =   10140
         Y1              =   4830
         Y2              =   4830
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Detalle de Documentos Observados"
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
         Left            =   60
         TabIndex        =   118
         Top             =   60
         Width           =   3060
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   420
         Index           =   1
         Left            =   -690
         Top             =   -90
         Width           =   10800
      End
   End
   Begin VB.Frame FraDepura 
      BorderStyle     =   0  'None
      Caption         =   "[ Depurar Datos ]"
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
      Height          =   5745
      Left            =   2280
      TabIndex        =   9
      Top             =   1980
      Visible         =   0   'False
      Width           =   10155
      Begin VB.CommandButton CmdDetalle 
         Caption         =   "Ver Detalle"
         Height          =   375
         Left            =   3250
         TabIndex        =   120
         Top             =   5310
         Width           =   1755
      End
      Begin VB.CommandButton CmdExportar 
         Caption         =   "Exportar MsExcel"
         Height          =   375
         Left            =   5240
         TabIndex        =   18
         Top             =   5310
         Width           =   1755
      End
      Begin VB.CommandButton CmdSalir 
         Caption         =   "&Salir"
         Height          =   375
         Left            =   7230
         TabIndex        =   16
         Top             =   5310
         Width           =   1755
      End
      Begin VB.PictureBox pic 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   9900
         Picture         =   "FrmDocumentos.frx":03EC
         ScaleHeight     =   210
         ScaleWidth      =   195
         TabIndex        =   14
         ToolTipText     =   "Cerrar"
         Top             =   60
         Width           =   195
      End
      Begin VB.CommandButton CmdDepura 
         Caption         =   "Depurar"
         Height          =   375
         Left            =   1260
         TabIndex        =   13
         Top             =   5310
         Width           =   1755
      End
      Begin SizerOneLibCtl.TabOne TabOne3 
         Height          =   4845
         Left            =   90
         TabIndex        =   10
         Top             =   360
         Width           =   9975
         _cx             =   17595
         _cy             =   8546
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
         Appearance      =   0
         MousePointer    =   0
         _ConvInfo       =   1
         Version         =   700
         BackColor       =   6579300
         ForeColor       =   -2147483634
         FrontTabColor   =   -2147483635
         BackTabColor    =   6579300
         TabOutlineColor =   0
         FrontTabForeColor=   -2147483634
         Caption         =   " &Proveedor|&Items|&T. Documento"
         Align           =   0
         CurrTab         =   1
         FirstTab        =   0
         Style           =   6
         Position        =   5
         AutoSwitch      =   -1  'True
         AutoScroll      =   -1  'True
         TabPreview      =   -1  'True
         ShowFocusRect   =   0   'False
         TabsPerPage     =   0
         BorderWidth     =   0
         BoldCurrent     =   -1  'True
         DogEars         =   -1  'True
         MultiRow        =   0   'False
         MultiRowOffset  =   200
         CaptionStyle    =   0
         TabHeight       =   0
         TabCaptionPos   =   1
         TabPicturePos   =   0
         CaptionEmpty    =   ""
         Separators      =   0   'False
         Begin VSFlex7Ctl.VSFlexGrid Fg3 
            Height          =   4815
            Left            =   -9240
            TabIndex        =   11
            Top             =   15
            Width           =   8625
            _cx             =   15214
            _cy             =   8493
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
            BackColorSel    =   4210816
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483627
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
            Rows            =   50
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmDocumentos.frx":06D8
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
         Begin VSFlex7Ctl.VSFlexGrid Fg5 
            Height          =   4815
            Left            =   1335
            TabIndex        =   12
            Top             =   15
            Width           =   8625
            _cx             =   15214
            _cy             =   8493
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
            BackColorSel    =   4210816
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483627
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
            Rows            =   50
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmDocumentos.frx":0773
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
            Begin VB.CommandButton CmdActualizar1 
               Caption         =   "2.- Todos x Gasto Administrativo"
               Height          =   585
               Left            =   5910
               TabIndex        =   122
               Top             =   4050
               Visible         =   0   'False
               Width           =   2295
            End
            Begin VB.CommandButton CmdActualizar 
               Caption         =   "1.- Distribuir x Tipos de Gasto"
               Height          =   585
               Left            =   3180
               TabIndex        =   121
               Top             =   4050
               Visible         =   0   'False
               Width           =   2295
            End
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg4 
            Height          =   4815
            Left            =   11910
            TabIndex        =   17
            Top             =   15
            Width           =   8625
            _cx             =   15214
            _cy             =   8493
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
            BackColorSel    =   4210816
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483627
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
            Rows            =   50
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmDocumentos.frx":0845
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
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   7
         X1              =   0
         X2              =   10140
         Y1              =   5250
         Y2              =   5250
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00400000&
         BackStyle       =   0  'Transparent
         Caption         =   "Depurar Datos"
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
         Left            =   60
         TabIndex        =   15
         Top             =   60
         Width           =   1245
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00400000&
         Height          =   270
         Index           =   0
         Left            =   30
         Top             =   30
         Width           =   10080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   3
         X1              =   0
         X2              =   10140
         Y1              =   5730
         Y2              =   5730
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   3
         X1              =   10140
         X2              =   10140
         Y1              =   30
         Y2              =   7000
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   2
         X1              =   0
         X2              =   0
         Y1              =   0
         Y2              =   6500
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   2
         X1              =   -30
         X2              =   11970
         Y1              =   0
         Y2              =   15
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   12240
      Top             =   390
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame5 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   1200
      Left            =   12300
      TabIndex        =   6
      Top             =   1950
      Visible         =   0   'False
      Width           =   5805
      Begin MSComctlLib.ProgressBar ProgressBar2 
         Height          =   300
         Left            =   135
         TabIndex        =   7
         Top             =   615
         Width           =   5565
         _ExtentX        =   9816
         _ExtentY        =   529
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   5820
         Y1              =   15
         Y2              =   0
      End
      Begin VB.Line Line3 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   30
         Y2              =   1170
      End
      Begin VB.Line Line3 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   1
         X1              =   5790
         X2              =   5790
         Y1              =   15
         Y2              =   1155
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   5835
         Y1              =   1185
         Y2              =   1170
      End
      Begin VB.Label LblBarra 
         AutoSize        =   -1  'True
         Caption         =   "Procesando Documentos : "
         Height          =   195
         Left            =   150
         TabIndex        =   8
         Top             =   300
         Width           =   1935
      End
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   13875
      _ExtentX        =   24474
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   18
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Agregar "
            ImageIndex      =   1
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Importación"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Transferencia"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Modificar "
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   1
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Restaurar Documento"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Eliminar "
            ImageIndex      =   5
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Eliminar Documento"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Retirar Documento"
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
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Filtrar"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Quitar Filtro"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Depurar Datos"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Buscar"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Cambiar Mes"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Exportar"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Object.ToolTipText     =   "Imprimir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   7410
         Top             =   45
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483637
         ImageWidth      =   16
         ImageHeight     =   15
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   16
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":08A9
               Key             =   "IMG1"
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":0DED
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":117F
               Key             =   "IMG2"
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":1303
               Key             =   "IMG3"
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":1757
               Key             =   "IMG4"
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":186F
               Key             =   "IMG5"
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":1DB3
               Key             =   "IMG6"
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":22F7
               Key             =   "IMG7"
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":240B
               Key             =   "IMG8"
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":251F
               Key             =   "IMG9"
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":2973
               Key             =   "IMG10"
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":2ADF
               Key             =   "IMG11"
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":3027
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":3341
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":379B
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDocumentos.frx":3CDD
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7410
      Left            =   0
      TabIndex        =   0
      Top             =   360
      Width           =   12090
      _cx             =   21325
      _cy             =   13070
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
      Caption         =   "  &Consulta  |   &Detalle   |   &Detalle - Importación  |  Detalle - Transferencia  "
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
      Begin VB.Frame Frame8 
         BorderStyle     =   0  'None
         Caption         =   "Frame8"
         Height          =   6990
         Left            =   13335
         TabIndex        =   29
         Top             =   375
         Width           =   12000
         Begin VB.CommandButton CmdConsultarTr 
            Caption         =   "&Consultar"
            Height          =   360
            Left            =   6690
            TabIndex        =   59
            Top             =   780
            Width           =   1470
         End
         Begin VB.Frame Frame10 
            Caption         =   "[ Seleccionar ]"
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
            Left            =   30
            TabIndex        =   51
            Top             =   270
            Width           =   4365
            Begin VB.OptionButton Opt4Tr 
               Caption         =   "Fch. Imp."
               Height          =   195
               Left            =   1290
               TabIndex        =   60
               Top             =   600
               Width           =   1035
            End
            Begin VB.OptionButton Opt1Tr 
               Caption         =   "Fch. Doc."
               Height          =   195
               Left            =   60
               TabIndex        =   54
               Top             =   330
               Value           =   -1  'True
               Width           =   1035
            End
            Begin VB.OptionButton Opt2Tr 
               Caption         =   "Fch. Recep."
               Height          =   195
               Left            =   60
               TabIndex        =   53
               Top             =   600
               Width           =   1185
            End
            Begin VB.OptionButton Opt3Tr 
               Caption         =   "Fch. Venc"
               Height          =   195
               Left            =   1290
               TabIndex        =   52
               Top             =   330
               Width           =   1035
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIniTr 
               Height          =   300
               Left            =   3045
               TabIndex        =   55
               Top             =   180
               Width           =   1245
               _ExtentX        =   2196
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
               Valor           =   "25/10/2011"
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchFinTr 
               Height          =   300
               Left            =   3045
               TabIndex        =   56
               Top             =   540
               Width           =   1245
               _ExtentX        =   2196
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
               Valor           =   "25/10/2011"
            End
            Begin VB.Label Label10 
               AutoSize        =   -1  'True
               Caption         =   "Hasta:"
               Height          =   195
               Left            =   2490
               TabIndex        =   58
               Top             =   645
               Width           =   465
            End
            Begin VB.Label Label9 
               AutoSize        =   -1  'True
               Caption         =   "Desde:"
               Height          =   195
               Left            =   2490
               TabIndex        =   57
               Top             =   285
               Width           =   510
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "[ Seleccionar Módulo ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   915
            Left            =   4410
            TabIndex        =   47
            Top             =   270
            Width           =   2175
            Begin VB.OptionButton OptHonorario 
               Caption         =   "Honorarios"
               Height          =   195
               Left            =   300
               TabIndex        =   49
               Top             =   600
               Width           =   1125
            End
            Begin VB.OptionButton OptCompra 
               Caption         =   "Compras"
               Height          =   195
               Left            =   300
               TabIndex        =   48
               Top             =   300
               Value           =   -1  'True
               Width           =   1035
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "[ Seleccionar Periodo a Transferir]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00000000&
            Height          =   915
            Left            =   8460
            TabIndex        =   31
            Top             =   240
            Width           =   3255
            Begin VB.CommandButton cmd_periodo1 
               Height          =   240
               Left            =   1320
               Picture         =   "FrmDocumentos.frx":406F
               Style           =   1  'Graphical
               TabIndex        =   32
               Top             =   390
               Width           =   270
            End
            Begin VB.Label LblPerIni 
               BackColor       =   &H00FFFFFF&
               BorderStyle     =   1  'Fixed Single
               Caption         =   "LblPerIni"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   270
               TabIndex        =   33
               Top             =   360
               Width           =   1365
            End
         End
         Begin TrueOleDBGrid70.TDBGrid Dg2 
            Height          =   5460
            Left            =   30
            TabIndex        =   30
            Top             =   1230
            Width           =   11670
            _ExtentX        =   20585
            _ExtentY        =   9631
            _LayoutType     =   4
            _RowHeight      =   -2147483647
            _WasPersistedAsPixels=   0
            Columns(0)._VlistStyle=   0
            Columns(0)._MaxComboItems=   5
            Columns(0).Caption=   "Id"
            Columns(0).DataField=   "id"
            Columns(0)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(1)._VlistStyle=   4
            Columns(1)._MaxComboItems=   5
            Columns(1).Caption=   "Sel"
            Columns(1).DataField=   "xsel"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Nº Reg"
            Columns(2).DataField=   "numreg"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Corr"
            Columns(3).DataField=   "corr"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Ruc Prov."
            Columns(4).DataField=   "rucprov"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Proveedor"
            Columns(5).DataField=   "proveedor"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "T.D."
            Columns(6).DataField=   "abrev"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Nª Documento"
            Columns(7).DataField=   "numerodoc"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "Fch. Emi."
            Columns(8).DataField=   "fchdoc1"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Fch. Recep."
            Columns(9).DataField=   "fchrecep1"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Fch. Venc."
            Columns(10).DataField=   "Fchvenc1"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Fch. Imp."
            Columns(11).DataField=   "fchproc1"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "M"
            Columns(12).DataField=   "simbolo"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "T.C."
            Columns(13).DataField=   "tipcam1"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "Total"
            Columns(14).DataField=   "imptot1"
            Columns(14).NumberFormat=   "0.00"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "Glosa"
            Columns(15).DataField=   "glosa"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   16
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).AllowColMove=   -1  'True
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=16"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=635"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=556"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=513"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1402"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1323"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=900"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=820"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=1879"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=1799"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=513"
            Splits(0)._ColumnProps(31)=   "Column(4).Visible=0"
            Splits(0)._ColumnProps(32)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(33)=   "Column(5).Width=4710"
            Splits(0)._ColumnProps(34)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(35)=   "Column(5)._WidthInPix=4630"
            Splits(0)._ColumnProps(36)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(37)=   "Column(5)._ColStyle=512"
            Splits(0)._ColumnProps(38)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(39)=   "Column(6).Width=820"
            Splits(0)._ColumnProps(40)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(41)=   "Column(6)._WidthInPix=741"
            Splits(0)._ColumnProps(42)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(43)=   "Column(6)._ColStyle=516"
            Splits(0)._ColumnProps(44)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(45)=   "Column(7).Width=2328"
            Splits(0)._ColumnProps(46)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._WidthInPix=2249"
            Splits(0)._ColumnProps(48)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(49)=   "Column(7)._ColStyle=516"
            Splits(0)._ColumnProps(50)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(51)=   "Column(8).Width=1640"
            Splits(0)._ColumnProps(52)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._WidthInPix=1561"
            Splits(0)._ColumnProps(54)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(55)=   "Column(8)._ColStyle=513"
            Splits(0)._ColumnProps(56)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(57)=   "Column(9).Width=2011"
            Splits(0)._ColumnProps(58)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(59)=   "Column(9)._WidthInPix=1931"
            Splits(0)._ColumnProps(60)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(61)=   "Column(9)._ColStyle=513"
            Splits(0)._ColumnProps(62)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(63)=   "Column(10).Width=1799"
            Splits(0)._ColumnProps(64)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(65)=   "Column(10)._WidthInPix=1720"
            Splits(0)._ColumnProps(66)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(67)=   "Column(10)._ColStyle=516"
            Splits(0)._ColumnProps(68)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(69)=   "Column(11).Width=1693"
            Splits(0)._ColumnProps(70)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(71)=   "Column(11)._WidthInPix=1614"
            Splits(0)._ColumnProps(72)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(73)=   "Column(11)._ColStyle=516"
            Splits(0)._ColumnProps(74)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(75)=   "Column(12).Width=794"
            Splits(0)._ColumnProps(76)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(77)=   "Column(12)._WidthInPix=714"
            Splits(0)._ColumnProps(78)=   "Column(12)._EditAlways=0"
            Splits(0)._ColumnProps(79)=   "Column(12)._ColStyle=513"
            Splits(0)._ColumnProps(80)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(81)=   "Column(13).Width=820"
            Splits(0)._ColumnProps(82)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(83)=   "Column(13)._WidthInPix=741"
            Splits(0)._ColumnProps(84)=   "Column(13)._EditAlways=0"
            Splits(0)._ColumnProps(85)=   "Column(13)._ColStyle=514"
            Splits(0)._ColumnProps(86)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(87)=   "Column(14).Width=1244"
            Splits(0)._ColumnProps(88)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(89)=   "Column(14)._WidthInPix=1164"
            Splits(0)._ColumnProps(90)=   "Column(14)._EditAlways=0"
            Splits(0)._ColumnProps(91)=   "Column(14)._ColStyle=514"
            Splits(0)._ColumnProps(92)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(93)=   "Column(15).Width=7990"
            Splits(0)._ColumnProps(94)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(95)=   "Column(15)._WidthInPix=7911"
            Splits(0)._ColumnProps(96)=   "Column(15)._EditAlways=0"
            Splits(0)._ColumnProps(97)=   "Column(15)._ColStyle=516"
            Splits(0)._ColumnProps(98)=   "Column(15).Order=16"
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
            ColumnFooters   =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1.5
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   0
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
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(27)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(28)  =   ":id=14,.fontname=MS Sans Serif"
            _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=74,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=90,.parent=13,.alignment=2"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=87,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=88,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=89,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=98,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=95,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=96,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=97,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=86,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=83,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=84,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=85,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=47,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=48,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=49,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=59,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=60,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=61,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=78,.parent=13"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=75,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=76,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=77,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=94,.parent=13"
            _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=91,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=92,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=93,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=25,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=26,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=27,.parent=17"
            _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=55,.parent=14"
            _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=56,.parent=15"
            _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=57,.parent=17"
            _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=70,.parent=13"
            _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=67,.parent=14"
            _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=68,.parent=15"
            _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=69,.parent=17"
            _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=82,.parent=13"
            _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=79,.parent=14"
            _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=80,.parent=15"
            _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=81,.parent=17"
            _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=29,.parent=14"
            _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=30,.parent=15"
            _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=31,.parent=17"
            _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=43,.parent=14"
            _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=44,.parent=15"
            _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=45,.parent=17"
            _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=51,.parent=14"
            _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=52,.parent=15"
            _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=53,.parent=17"
            _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=66,.parent=13"
            _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=63,.parent=14"
            _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=64,.parent=15"
            _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=65,.parent=17"
            _StyleDefs(102) =   "Named:id=33:Normal"
            _StyleDefs(103) =   ":id=33,.parent=0"
            _StyleDefs(104) =   "Named:id=34:Heading"
            _StyleDefs(105) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(106) =   ":id=34,.wraptext=-1"
            _StyleDefs(107) =   "Named:id=35:Footing"
            _StyleDefs(108) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(109) =   "Named:id=36:Selected"
            _StyleDefs(110) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(111) =   "Named:id=37:Caption"
            _StyleDefs(112) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(113) =   "Named:id=38:HighlightRow"
            _StyleDefs(114) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(115) =   "Named:id=39:EvenRow"
            _StyleDefs(116) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(117) =   "Named:id=40:OddRow"
            _StyleDefs(118) =   ":id=40,.parent=33"
            _StyleDefs(119) =   "Named:id=41:RecordSelector"
            _StyleDefs(120) =   ":id=41,.parent=34"
            _StyleDefs(121) =   "Named:id=42:FilterBar"
            _StyleDefs(122) =   ":id=42,.parent=33"
         End
         Begin VB.Line Line5 
            BorderWidth     =   2
            X1              =   8310
            X2              =   8310
            Y1              =   330
            Y2              =   1200
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            Caption         =   "Detalle - Transferencia"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   120
            TabIndex        =   46
            Top             =   30
            Width           =   11550
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  'None
         Caption         =   "Frame7"
         Height          =   6990
         Left            =   13035
         TabIndex        =   28
         Top             =   375
         Width           =   12000
         Begin VB.Frame Frame4 
            Height          =   960
            Left            =   60
            TabIndex        =   37
            Top             =   240
            Width           =   11625
            Begin VB.CommandButton CmdBusArch 
               Height          =   240
               Left            =   9030
               Picture         =   "FrmDocumentos.frx":43F1
               Style           =   1  'Graphical
               TabIndex        =   40
               Top             =   220
               Width           =   240
            End
            Begin VB.CommandButton CmdCargar 
               Caption         =   "Cargar"
               Height          =   615
               Left            =   10290
               TabIndex        =   39
               Top             =   210
               Width           =   1125
            End
            Begin VB.CommandButton CmdBusArch2 
               Height          =   240
               Left            =   9030
               Picture         =   "FrmDocumentos.frx":4523
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   540
               Width           =   240
            End
            Begin VB.TextBox TxtArchivo 
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   1305
               TabIndex        =   42
               Text            =   "TxtArchivo"
               Top             =   195
               Width           =   8000
            End
            Begin VB.TextBox TxtArchivo2 
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   1305
               TabIndex        =   41
               Text            =   "TxtArchivo2"
               Top             =   510
               Width           =   8000
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Arch. Cabecera"
               Height          =   195
               Index           =   0
               Left            =   120
               TabIndex        =   44
               Top             =   225
               Width           =   1110
            End
            Begin VB.Label Label13 
               AutoSize        =   -1  'True
               Caption         =   "Arch. Detalle"
               Height          =   195
               Left            =   120
               TabIndex        =   43
               Top             =   540
               Width           =   915
            End
         End
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   5445
            Left            =   60
            TabIndex        =   34
            Top             =   1230
            Width           =   11625
            _cx             =   20505
            _cy             =   9604
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
            Appearance      =   1
            MousePointer    =   0
            _ConvInfo       =   1
            Version         =   700
            BackColor       =   -2147483633
            ForeColor       =   -2147483633
            FrontTabColor   =   -2147483633
            BackTabColor    =   -2147483632
            TabOutlineColor =   -2147483633
            FrontTabForeColor=   -2147483630
            Caption         =   " &Cabecera |     &Detalle   "
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   0
            Position        =   1
            AutoSwitch      =   -1  'True
            AutoScroll      =   -1  'True
            TabPreview      =   -1  'True
            ShowFocusRect   =   0   'False
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
            Begin VSFlex7Ctl.VSFlexGrid Fg1 
               Height          =   5070
               Left            =   45
               TabIndex        =   35
               Top             =   45
               Width           =   11535
               _cx             =   20346
               _cy             =   8943
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
               BackColorSel    =   4210816
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483627
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
               Rows            =   50
               Cols            =   24
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmDocumentos.frx":4655
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
               Height          =   5070
               Left            =   12270
               TabIndex        =   36
               Top             =   45
               Width           =   11535
               _cx             =   20346
               _cy             =   8943
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
               BackColorSel    =   4210816
               ForeColorSel    =   -2147483634
               BackColorBkg    =   -2147483636
               BackColorAlternate=   -2147483643
               GridColor       =   -2147483627
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
               Rows            =   50
               Cols            =   15
               FixedRows       =   1
               FixedCols       =   1
               RowHeightMin    =   0
               RowHeightMax    =   0
               ColWidthMin     =   0
               ColWidthMax     =   0
               ExtendLastCol   =   0   'False
               FormatString    =   $"FrmDocumentos.frx":4911
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
         Begin VB.Label Label5 
            Alignment       =   2  'Center
            Caption         =   "Detalle - Importación"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   60
            TabIndex        =   45
            Top             =   30
            Width           =   11610
         End
      End
      Begin VB.Frame Frame2 
         BorderStyle     =   0  'None
         Height          =   6990
         Left            =   12735
         TabIndex        =   4
         Top             =   375
         Width           =   12000
         Begin VB.Frame Frame12 
            Caption         =   "[ Transferencia ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1455
            Left            =   9270
            TabIndex        =   106
            Top             =   300
            Width           =   2385
            Begin VB.TextBox TxtEstado 
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   780
               Locked          =   -1  'True
               TabIndex        =   112
               Text            =   "TxtEstado"
               Top             =   240
               Width           =   1530
            End
            Begin VB.TextBox TxtNumReg 
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   780
               Locked          =   -1  'True
               TabIndex        =   110
               Text            =   "TxtNumReg"
               Top             =   960
               Width           =   1530
            End
            Begin VB.TextBox TxtModulo 
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   780
               Locked          =   -1  'True
               TabIndex        =   108
               Text            =   "TxtModulo"
               Top             =   600
               Width           =   1530
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Estado"
               Height          =   195
               Index           =   18
               Left            =   90
               TabIndex        =   113
               Top             =   330
               Width           =   495
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "N°. Reg. "
               Height          =   195
               Index           =   17
               Left            =   90
               TabIndex        =   109
               Top             =   1020
               Width           =   660
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Módulo"
               Height          =   195
               Index           =   14
               Left            =   90
               TabIndex        =   107
               Top             =   690
               Width           =   525
            End
         End
         Begin VB.TextBox TxtTipoDoc 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   105
            Text            =   "TxtTipoDoc"
            Top             =   1410
            Width           =   2730
         End
         Begin VB.TextBox TxtTC 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   7860
            Locked          =   -1  'True
            TabIndex        =   104
            Text            =   "TxtTC"
            Top             =   1410
            Width           =   1260
         End
         Begin VB.TextBox TxtMoneda 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   5130
            Locked          =   -1  'True
            MaxLength       =   11
            TabIndex        =   103
            Text            =   "TxtMoneda"
            Top             =   1410
            Width           =   1260
         End
         Begin VB.TextBox TxtNomCli 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   3450
            Locked          =   -1  'True
            TabIndex        =   102
            Text            =   "TxtNomCli"
            Top             =   2730
            Width           =   5670
         End
         Begin VB.TextBox TxtNomProv 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   3450
            Locked          =   -1  'True
            TabIndex        =   101
            Text            =   "TxtNomProv"
            Top             =   690
            Width           =   5670
         End
         Begin VB.TextBox TxtPeriodo 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   3210
            Locked          =   -1  'True
            TabIndex        =   100
            Text            =   "TxtPeriodo"
            Top             =   330
            Width           =   1320
         End
         Begin VB.TextBox TxtAnno 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   98
            Text            =   "TxtAnno"
            Top             =   330
            Width           =   780
         End
         Begin VB.TextBox TxtCorr 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   5130
            Locked          =   -1  'True
            TabIndex        =   96
            Text            =   "TxtCorr"
            Top             =   330
            Width           =   780
         End
         Begin VB.TextBox TxtRucCli 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   90
            Text            =   "TxtRucCli"
            Top             =   2730
            Width           =   1770
         End
         Begin VB.Frame Frame11 
            Height          =   735
            Left            =   60
            TabIndex        =   66
            Top             =   5985
            Width           =   11610
            Begin VB.TextBox TxtOtros 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               Left            =   8955
               Locked          =   -1  'True
               TabIndex        =   85
               TabStop         =   0   'False
               Text            =   "TxtOtros"
               Top             =   345
               Width           =   1230
            End
            Begin VB.TextBox TxtInafecto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               Left            =   6360
               Locked          =   -1  'True
               TabIndex        =   70
               TabStop         =   0   'False
               Text            =   "TxtInafect"
               Top             =   345
               Width           =   1230
            End
            Begin VB.TextBox TxtBruto 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               TabIndex        =   69
               TabStop         =   0   'False
               Text            =   "TxtBruto"
               Top             =   345
               Width           =   1230
            End
            Begin VB.TextBox TxtIGV 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               Left            =   7650
               Locked          =   -1  'True
               TabIndex        =   68
               TabStop         =   0   'False
               Text            =   "TxtIGV"
               Top             =   345
               Width           =   1230
            End
            Begin VB.TextBox TxtTotal 
               Alignment       =   1  'Right Justify
               Appearance      =   0  'Flat
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
               Left            =   10260
               Locked          =   -1  'True
               TabIndex        =   67
               TabStop         =   0   'False
               Text            =   "TxtTotal"
               Top             =   330
               Width           =   1230
            End
            Begin VB.Label Label11 
               AutoSize        =   -1  'True
               Caption         =   "Otros Cargos"
               Height          =   195
               Left            =   8955
               TabIndex        =   86
               Top             =   150
               Width           =   915
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total Inafecto"
               Height          =   195
               Index           =   3
               Left            =   6360
               TabIndex        =   74
               Top             =   150
               Width           =   990
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total Afecto"
               Height          =   195
               Index           =   1
               Left            =   5055
               TabIndex        =   73
               Top             =   150
               Width           =   870
            End
            Begin VB.Label LblRotulo 
               AutoSize        =   -1  'True
               Caption         =   "I.G.V. "
               Height          =   195
               Left            =   7650
               TabIndex        =   72
               Top             =   150
               Width           =   450
            End
            Begin VB.Label Label1 
               AutoSize        =   -1  'True
               Caption         =   "Total"
               Height          =   195
               Index           =   2
               Left            =   10260
               TabIndex        =   71
               Top             =   150
               Width           =   360
            End
         End
         Begin VB.TextBox TxtNumSer 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1620
            Locked          =   -1  'True
            MaxLength       =   4
            TabIndex        =   65
            Text            =   "TxtNumSer"
            Top             =   1770
            Width           =   915
         End
         Begin VB.TextBox TxtNumDoc 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2640
            Locked          =   -1  'True
            TabIndex        =   64
            Text            =   "TxtNumDoc"
            Top             =   1770
            Width           =   1710
         End
         Begin VB.TextBox TxtGlosa 
            Appearance      =   0  'Flat
            Height          =   540
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   61
            Text            =   "TxtGlosa"
            Top             =   2130
            Width           =   9990
         End
         Begin VSFlex7Ctl.VSFlexGrid Fg6 
            Height          =   2430
            Left            =   60
            TabIndex        =   62
            Top             =   3510
            Width           =   11610
            _cx             =   20479
            _cy             =   4286
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
            GridColor       =   -2147483627
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
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"FrmDocumentos.frx":4AE6
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
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchVen 
            Height          =   300
            Left            =   7860
            TabIndex        =   63
            Top             =   1050
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
            Valor           =   "22/05/2008"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchDoc 
            Height          =   300
            Left            =   1620
            TabIndex        =   84
            Top             =   1050
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
            Valor           =   "22/05/2008"
         End
         Begin VB.TextBox TxtNumRuc 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   75
            Text            =   "TxtNumRuc"
            Top             =   690
            Width           =   1770
         End
         Begin VB.TextBox TxtNumOrden 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   1620
            Locked          =   -1  'True
            TabIndex        =   87
            Text            =   "TxtDocRef2"
            Top             =   3090
            Width           =   2025
         End
         Begin AspaTextBoxFecha.TextBoxFecha TctFchRecep 
            Height          =   300
            Left            =   5130
            TabIndex        =   89
            Top             =   1050
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
            Valor           =   "22/05/2008"
         End
         Begin AspaTextBoxFecha.TextBoxFecha TxtFchSist 
            Height          =   300
            Left            =   7860
            TabIndex        =   93
            Top             =   330
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
            Valor           =   "22/05/2008"
         End
         Begin VB.Line Line2 
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            Index           =   4
            X1              =   0
            X2              =   11790
            Y1              =   6810
            Y2              =   6810
         End
         Begin VB.Line Line7 
            BorderColor     =   &H00000000&
            BorderWidth     =   2
            X1              =   11790
            X2              =   11790
            Y1              =   -60
            Y2              =   6810
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Periodo"
            Height          =   195
            Index           =   16
            Left            =   2520
            TabIndex        =   99
            Top             =   450
            Width           =   540
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Año"
            Height          =   195
            Index           =   15
            Left            =   180
            TabIndex        =   97
            Top             =   450
            Width           =   285
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Corr."
            Height          =   195
            Index           =   12
            Left            =   4710
            TabIndex        =   95
            Top             =   450
            Width           =   330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Sist"
            Height          =   195
            Index           =   11
            Left            =   7125
            TabIndex        =   94
            Top             =   450
            Width           =   615
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "T.C."
            Height          =   195
            Index           =   9
            Left            =   7440
            TabIndex        =   92
            Top             =   1515
            Width           =   300
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Cliente"
            Height          =   195
            Index           =   8
            Left            =   180
            TabIndex        =   91
            Top             =   2835
            Width           =   480
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Recepc."
            Height          =   195
            Index           =   4
            Left            =   4065
            TabIndex        =   88
            Top             =   1155
            Width           =   975
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Venc."
            Height          =   195
            Index           =   6
            Left            =   6960
            TabIndex        =   83
            Top             =   1155
            Width           =   780
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Fch. Doc."
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   82
            Top             =   1155
            Width           =   705
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº de Documento"
            Height          =   195
            Index           =   2
            Left            =   180
            TabIndex        =   81
            Top             =   1875
            Width           =   1275
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Moneda"
            Height          =   195
            Index           =   5
            Left            =   4455
            TabIndex        =   80
            Top             =   1515
            Width           =   585
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Glosa"
            Height          =   195
            Index           =   10
            Left            =   180
            TabIndex        =   79
            Top             =   2175
            Width           =   405
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Proveedor"
            Height          =   195
            Index           =   7
            Left            =   180
            TabIndex        =   78
            Top             =   795
            Width           =   735
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Tipo Documento"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   77
            Top             =   1515
            Width           =   1185
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Nº Orden"
            Height          =   195
            Index           =   13
            Left            =   180
            TabIndex        =   76
            Top             =   3165
            Width           =   660
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            Caption         =   "Detalle"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   255
            Left            =   90
            TabIndex        =   50
            Top             =   30
            Width           =   11550
         End
      End
      Begin VB.Frame Frame1 
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6990
         Left            =   45
         TabIndex        =   1
         Top             =   375
         Width           =   12000
         Begin VB.Frame Frame6 
            Caption         =   "[ Seleccionar ]"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   645
            Left            =   30
            TabIndex        =   19
            Top             =   210
            Width           =   11655
            Begin VB.OptionButton Opt4 
               Caption         =   "Fch. Imp."
               Height          =   195
               Left            =   3690
               TabIndex        =   111
               Top             =   300
               Width           =   1035
            End
            Begin VB.CommandButton CmdConsultar 
               Caption         =   "Consultar"
               Height          =   360
               Left            =   10050
               TabIndex        =   27
               Top             =   210
               Width           =   1470
            End
            Begin VB.OptionButton Opt3 
               Caption         =   "Fch. Venc"
               Height          =   195
               Left            =   2550
               TabIndex        =   26
               Top             =   300
               Width           =   1035
            End
            Begin VB.OptionButton Opt2 
               Caption         =   "Fch. Recep."
               Height          =   195
               Left            =   1260
               TabIndex        =   25
               Top             =   300
               Width           =   1185
            End
            Begin VB.OptionButton Opt1 
               Caption         =   "Fch. Doc."
               Height          =   195
               Left            =   120
               TabIndex        =   24
               Top             =   300
               Value           =   -1  'True
               Width           =   1035
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
               Height          =   300
               Left            =   5625
               TabIndex        =   20
               Top             =   210
               Width           =   1245
               _ExtentX        =   2196
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
               Valor           =   "25/10/2011"
            End
            Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
               Height          =   300
               Left            =   7665
               TabIndex        =   21
               Top             =   210
               Width           =   1245
               _ExtentX        =   2196
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
               Valor           =   "25/10/2011"
            End
            Begin VB.Line Line4 
               Index           =   1
               X1              =   4890
               X2              =   4890
               Y1              =   210
               Y2              =   510
            End
            Begin VB.Label Label6 
               AutoSize        =   -1  'True
               Caption         =   "Desde:"
               Height          =   195
               Left            =   5070
               TabIndex        =   23
               Top             =   315
               Width           =   510
            End
            Begin VB.Label Label3 
               AutoSize        =   -1  'True
               Caption         =   "Hasta:"
               Height          =   195
               Index           =   0
               Left            =   7110
               TabIndex        =   22
               Top             =   315
               Width           =   465
            End
         End
         Begin TrueOleDBGrid70.TDBGrid Dg1 
            Height          =   5730
            Left            =   0
            TabIndex        =   2
            Top             =   960
            Width           =   11730
            _ExtentX        =   20690
            _ExtentY        =   10107
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
            Columns(1).Caption=   "Estado"
            Columns(1).DataField=   "estado1"
            Columns(1)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(2)._VlistStyle=   0
            Columns(2)._MaxComboItems=   5
            Columns(2).Caption=   "Módulo"
            Columns(2).DataField=   "modulo"
            Columns(2)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(3)._VlistStyle=   0
            Columns(3)._MaxComboItems=   5
            Columns(3).Caption=   "Nº Reg"
            Columns(3).DataField=   "numreg"
            Columns(3)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(4)._VlistStyle=   0
            Columns(4)._MaxComboItems=   5
            Columns(4).Caption=   "Mes"
            Columns(4).DataField=   "messist"
            Columns(4)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(5)._VlistStyle=   0
            Columns(5)._MaxComboItems=   5
            Columns(5).Caption=   "Corr"
            Columns(5).DataField=   "corr"
            Columns(5)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(6)._VlistStyle=   0
            Columns(6)._MaxComboItems=   5
            Columns(6).Caption=   "Ruc Prov."
            Columns(6).DataField=   "rucprov"
            Columns(6)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(7)._VlistStyle=   0
            Columns(7)._MaxComboItems=   5
            Columns(7).Caption=   "Proveedor"
            Columns(7).DataField=   "proveedor"
            Columns(7)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(8)._VlistStyle=   0
            Columns(8)._MaxComboItems=   5
            Columns(8).Caption=   "T.D."
            Columns(8).DataField=   "abrev"
            Columns(8)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(9)._VlistStyle=   0
            Columns(9)._MaxComboItems=   5
            Columns(9).Caption=   "Nª Documento"
            Columns(9).DataField=   "numerodoc"
            Columns(9)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(10)._VlistStyle=   0
            Columns(10)._MaxComboItems=   5
            Columns(10).Caption=   "Fch. Emi"
            Columns(10).DataField=   "fchdoc"
            Columns(10)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(11)._VlistStyle=   0
            Columns(11)._MaxComboItems=   5
            Columns(11).Caption=   "Fch. Recep."
            Columns(11).DataField=   "fchrecep"
            Columns(11)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(12)._VlistStyle=   0
            Columns(12)._MaxComboItems=   5
            Columns(12).Caption=   "M"
            Columns(12).DataField=   "simbolo"
            Columns(12)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(13)._VlistStyle=   0
            Columns(13)._MaxComboItems=   5
            Columns(13).Caption=   "T.C."
            Columns(13).DataField=   "tipcam"
            Columns(13)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(14)._VlistStyle=   0
            Columns(14)._MaxComboItems=   5
            Columns(14).Caption=   "Total"
            Columns(14).DataField=   "imptot"
            Columns(14).NumberFormat=   "0.00"
            Columns(14)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns(15)._VlistStyle=   0
            Columns(15)._MaxComboItems=   5
            Columns(15).Caption=   "Glosa"
            Columns(15).DataField=   "glosa"
            Columns(15)._PropDict=   "_MaxComboItems,516,2;_VlistStyle,514,3"
            Columns.Count   =   16
            Splits(0)._UserFlags=   0
            Splits(0).MarqueeStyle=   3
            Splits(0).RecordSelectorWidth=   344
            Splits(0)._SavedRecordSelectors=   0   'False
            Splits(0).DividerColor=   12632256
            Splits(0).FilterBar=   -1  'True
            Splits(0).SpringMode=   0   'False
            Splits(0)._PropDict=   "_ColumnProps,515,0;_UserFlags,518,3"
            Splits(0)._ColumnProps(0)=   "Columns.Count=16"
            Splits(0)._ColumnProps(1)=   "Column(0).Width=2725"
            Splits(0)._ColumnProps(2)=   "Column(0).DividerColor=0"
            Splits(0)._ColumnProps(3)=   "Column(0)._WidthInPix=2646"
            Splits(0)._ColumnProps(4)=   "Column(0)._EditAlways=0"
            Splits(0)._ColumnProps(5)=   "Column(0)._ColStyle=516"
            Splits(0)._ColumnProps(6)=   "Column(0).Visible=0"
            Splits(0)._ColumnProps(7)=   "Column(0).Order=1"
            Splits(0)._ColumnProps(8)=   "Column(1).Width=1535"
            Splits(0)._ColumnProps(9)=   "Column(1).DividerColor=0"
            Splits(0)._ColumnProps(10)=   "Column(1)._WidthInPix=1455"
            Splits(0)._ColumnProps(11)=   "Column(1)._EditAlways=0"
            Splits(0)._ColumnProps(12)=   "Column(1)._ColStyle=516"
            Splits(0)._ColumnProps(13)=   "Column(1).Order=2"
            Splits(0)._ColumnProps(14)=   "Column(2).Width=1561"
            Splits(0)._ColumnProps(15)=   "Column(2).DividerColor=0"
            Splits(0)._ColumnProps(16)=   "Column(2)._WidthInPix=1482"
            Splits(0)._ColumnProps(17)=   "Column(2)._EditAlways=0"
            Splits(0)._ColumnProps(18)=   "Column(2)._ColStyle=516"
            Splits(0)._ColumnProps(19)=   "Column(2).Order=3"
            Splits(0)._ColumnProps(20)=   "Column(3).Width=1667"
            Splits(0)._ColumnProps(21)=   "Column(3).DividerColor=0"
            Splits(0)._ColumnProps(22)=   "Column(3)._WidthInPix=1588"
            Splits(0)._ColumnProps(23)=   "Column(3)._EditAlways=0"
            Splits(0)._ColumnProps(24)=   "Column(3)._ColStyle=516"
            Splits(0)._ColumnProps(25)=   "Column(3).Order=4"
            Splits(0)._ColumnProps(26)=   "Column(4).Width=900"
            Splits(0)._ColumnProps(27)=   "Column(4).DividerColor=0"
            Splits(0)._ColumnProps(28)=   "Column(4)._WidthInPix=820"
            Splits(0)._ColumnProps(29)=   "Column(4)._EditAlways=0"
            Splits(0)._ColumnProps(30)=   "Column(4)._ColStyle=516"
            Splits(0)._ColumnProps(31)=   "Column(4).Order=5"
            Splits(0)._ColumnProps(32)=   "Column(5).Width=900"
            Splits(0)._ColumnProps(33)=   "Column(5).DividerColor=0"
            Splits(0)._ColumnProps(34)=   "Column(5)._WidthInPix=820"
            Splits(0)._ColumnProps(35)=   "Column(5)._EditAlways=0"
            Splits(0)._ColumnProps(36)=   "Column(5)._ColStyle=516"
            Splits(0)._ColumnProps(37)=   "Column(5).Order=6"
            Splits(0)._ColumnProps(38)=   "Column(6).Width=1879"
            Splits(0)._ColumnProps(39)=   "Column(6).DividerColor=0"
            Splits(0)._ColumnProps(40)=   "Column(6)._WidthInPix=1799"
            Splits(0)._ColumnProps(41)=   "Column(6)._EditAlways=0"
            Splits(0)._ColumnProps(42)=   "Column(6)._ColStyle=513"
            Splits(0)._ColumnProps(43)=   "Column(6).Visible=0"
            Splits(0)._ColumnProps(44)=   "Column(6).Order=7"
            Splits(0)._ColumnProps(45)=   "Column(7).Width=4260"
            Splits(0)._ColumnProps(46)=   "Column(7).DividerColor=0"
            Splits(0)._ColumnProps(47)=   "Column(7)._WidthInPix=4180"
            Splits(0)._ColumnProps(48)=   "Column(7)._EditAlways=0"
            Splits(0)._ColumnProps(49)=   "Column(7)._ColStyle=512"
            Splits(0)._ColumnProps(50)=   "Column(7).Order=8"
            Splits(0)._ColumnProps(51)=   "Column(8).Width=926"
            Splits(0)._ColumnProps(52)=   "Column(8).DividerColor=0"
            Splits(0)._ColumnProps(53)=   "Column(8)._WidthInPix=847"
            Splits(0)._ColumnProps(54)=   "Column(8)._EditAlways=0"
            Splits(0)._ColumnProps(55)=   "Column(8)._ColStyle=516"
            Splits(0)._ColumnProps(56)=   "Column(8).Order=9"
            Splits(0)._ColumnProps(57)=   "Column(9).Width=2593"
            Splits(0)._ColumnProps(58)=   "Column(9).DividerColor=0"
            Splits(0)._ColumnProps(59)=   "Column(9)._WidthInPix=2514"
            Splits(0)._ColumnProps(60)=   "Column(9)._EditAlways=0"
            Splits(0)._ColumnProps(61)=   "Column(9)._ColStyle=516"
            Splits(0)._ColumnProps(62)=   "Column(9).Order=10"
            Splits(0)._ColumnProps(63)=   "Column(10).Width=1879"
            Splits(0)._ColumnProps(64)=   "Column(10).DividerColor=0"
            Splits(0)._ColumnProps(65)=   "Column(10)._WidthInPix=1799"
            Splits(0)._ColumnProps(66)=   "Column(10)._EditAlways=0"
            Splits(0)._ColumnProps(67)=   "Column(10)._ColStyle=513"
            Splits(0)._ColumnProps(68)=   "Column(10).Order=11"
            Splits(0)._ColumnProps(69)=   "Column(11).Width=2037"
            Splits(0)._ColumnProps(70)=   "Column(11).DividerColor=0"
            Splits(0)._ColumnProps(71)=   "Column(11)._WidthInPix=1958"
            Splits(0)._ColumnProps(72)=   "Column(11)._EditAlways=0"
            Splits(0)._ColumnProps(73)=   "Column(11)._ColStyle=513"
            Splits(0)._ColumnProps(74)=   "Column(11).Order=12"
            Splits(0)._ColumnProps(75)=   "Column(12).Width=794"
            Splits(0)._ColumnProps(76)=   "Column(12).DividerColor=0"
            Splits(0)._ColumnProps(77)=   "Column(12)._WidthInPix=714"
            Splits(0)._ColumnProps(78)=   "Column(12)._EditAlways=0"
            Splits(0)._ColumnProps(79)=   "Column(12)._ColStyle=513"
            Splits(0)._ColumnProps(80)=   "Column(12).Order=13"
            Splits(0)._ColumnProps(81)=   "Column(13).Width=926"
            Splits(0)._ColumnProps(82)=   "Column(13).DividerColor=0"
            Splits(0)._ColumnProps(83)=   "Column(13)._WidthInPix=847"
            Splits(0)._ColumnProps(84)=   "Column(13)._EditAlways=0"
            Splits(0)._ColumnProps(85)=   "Column(13)._ColStyle=514"
            Splits(0)._ColumnProps(86)=   "Column(13).Order=14"
            Splits(0)._ColumnProps(87)=   "Column(14).Width=1270"
            Splits(0)._ColumnProps(88)=   "Column(14).DividerColor=0"
            Splits(0)._ColumnProps(89)=   "Column(14)._WidthInPix=1191"
            Splits(0)._ColumnProps(90)=   "Column(14)._EditAlways=0"
            Splits(0)._ColumnProps(91)=   "Column(14)._ColStyle=514"
            Splits(0)._ColumnProps(92)=   "Column(14).Order=15"
            Splits(0)._ColumnProps(93)=   "Column(15).Width=6853"
            Splits(0)._ColumnProps(94)=   "Column(15).DividerColor=0"
            Splits(0)._ColumnProps(95)=   "Column(15)._WidthInPix=6773"
            Splits(0)._ColumnProps(96)=   "Column(15)._EditAlways=0"
            Splits(0)._ColumnProps(97)=   "Column(15)._ColStyle=516"
            Splits(0)._ColumnProps(98)=   "Column(15).Order=16"
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
            ColumnFooters   =   -1  'True
            DefColWidth     =   0
            HeadLines       =   1.5
            FootLines       =   1
            MultipleLines   =   0
            CellTipsWidth   =   0
            DeadAreaBackColor=   12632256
            RowDividerColor =   12632256
            RowSubDividerColor=   12632256
            DirectionAfterEnter=   0
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
            _StyleDefs(26)  =   "Splits(0).HeadingStyle:id=14,.parent=2,.bold=-1,.fontsize=825,.italic=0"
            _StyleDefs(27)  =   ":id=14,.underline=0,.strikethrough=0,.charset=0"
            _StyleDefs(28)  =   ":id=14,.fontname=MS Sans Serif"
            _StyleDefs(29)  =   "Splits(0).FooterStyle:id=15,.parent=3"
            _StyleDefs(30)  =   "Splits(0).InactiveStyle:id=16,.parent=5"
            _StyleDefs(31)  =   "Splits(0).SelectedStyle:id=18,.parent=6"
            _StyleDefs(32)  =   "Splits(0).EditorStyle:id=17,.parent=7"
            _StyleDefs(33)  =   "Splits(0).HighlightRowStyle:id=19,.parent=8,.bgcolor=&H80&"
            _StyleDefs(34)  =   "Splits(0).EvenRowStyle:id=20,.parent=9"
            _StyleDefs(35)  =   "Splits(0).OddRowStyle:id=21,.parent=10"
            _StyleDefs(36)  =   "Splits(0).RecordSelectorStyle:id=23,.parent=11"
            _StyleDefs(37)  =   "Splits(0).FilterBarStyle:id=24,.parent=12"
            _StyleDefs(38)  =   "Splits(0).Columns(0).Style:id=74,.parent=13"
            _StyleDefs(39)  =   "Splits(0).Columns(0).HeadingStyle:id=71,.parent=14"
            _StyleDefs(40)  =   "Splits(0).Columns(0).FooterStyle:id=72,.parent=15"
            _StyleDefs(41)  =   "Splits(0).Columns(0).EditorStyle:id=73,.parent=17"
            _StyleDefs(42)  =   "Splits(0).Columns(1).Style:id=102,.parent=13"
            _StyleDefs(43)  =   "Splits(0).Columns(1).HeadingStyle:id=99,.parent=14"
            _StyleDefs(44)  =   "Splits(0).Columns(1).FooterStyle:id=100,.parent=15"
            _StyleDefs(45)  =   "Splits(0).Columns(1).EditorStyle:id=101,.parent=17"
            _StyleDefs(46)  =   "Splits(0).Columns(2).Style:id=82,.parent=13"
            _StyleDefs(47)  =   "Splits(0).Columns(2).HeadingStyle:id=79,.parent=14"
            _StyleDefs(48)  =   "Splits(0).Columns(2).FooterStyle:id=80,.parent=15"
            _StyleDefs(49)  =   "Splits(0).Columns(2).EditorStyle:id=81,.parent=17"
            _StyleDefs(50)  =   "Splits(0).Columns(3).Style:id=70,.parent=13"
            _StyleDefs(51)  =   "Splits(0).Columns(3).HeadingStyle:id=67,.parent=14"
            _StyleDefs(52)  =   "Splits(0).Columns(3).FooterStyle:id=68,.parent=15"
            _StyleDefs(53)  =   "Splits(0).Columns(3).EditorStyle:id=69,.parent=17"
            _StyleDefs(54)  =   "Splits(0).Columns(4).Style:id=94,.parent=13"
            _StyleDefs(55)  =   "Splits(0).Columns(4).HeadingStyle:id=91,.parent=14"
            _StyleDefs(56)  =   "Splits(0).Columns(4).FooterStyle:id=92,.parent=15"
            _StyleDefs(57)  =   "Splits(0).Columns(4).EditorStyle:id=93,.parent=17"
            _StyleDefs(58)  =   "Splits(0).Columns(5).Style:id=86,.parent=13"
            _StyleDefs(59)  =   "Splits(0).Columns(5).HeadingStyle:id=83,.parent=14"
            _StyleDefs(60)  =   "Splits(0).Columns(5).FooterStyle:id=84,.parent=15"
            _StyleDefs(61)  =   "Splits(0).Columns(5).EditorStyle:id=85,.parent=17"
            _StyleDefs(62)  =   "Splits(0).Columns(6).Style:id=50,.parent=13,.alignment=2"
            _StyleDefs(63)  =   "Splits(0).Columns(6).HeadingStyle:id=47,.parent=14"
            _StyleDefs(64)  =   "Splits(0).Columns(6).FooterStyle:id=48,.parent=15"
            _StyleDefs(65)  =   "Splits(0).Columns(6).EditorStyle:id=49,.parent=17"
            _StyleDefs(66)  =   "Splits(0).Columns(7).Style:id=62,.parent=13,.alignment=0"
            _StyleDefs(67)  =   "Splits(0).Columns(7).HeadingStyle:id=59,.parent=14"
            _StyleDefs(68)  =   "Splits(0).Columns(7).FooterStyle:id=60,.parent=15"
            _StyleDefs(69)  =   "Splits(0).Columns(7).EditorStyle:id=61,.parent=17"
            _StyleDefs(70)  =   "Splits(0).Columns(8).Style:id=78,.parent=13"
            _StyleDefs(71)  =   "Splits(0).Columns(8).HeadingStyle:id=75,.parent=14"
            _StyleDefs(72)  =   "Splits(0).Columns(8).FooterStyle:id=76,.parent=15"
            _StyleDefs(73)  =   "Splits(0).Columns(8).EditorStyle:id=77,.parent=17"
            _StyleDefs(74)  =   "Splits(0).Columns(9).Style:id=90,.parent=13"
            _StyleDefs(75)  =   "Splits(0).Columns(9).HeadingStyle:id=87,.parent=14"
            _StyleDefs(76)  =   "Splits(0).Columns(9).FooterStyle:id=88,.parent=15"
            _StyleDefs(77)  =   "Splits(0).Columns(9).EditorStyle:id=89,.parent=17"
            _StyleDefs(78)  =   "Splits(0).Columns(10).Style:id=28,.parent=13,.alignment=2"
            _StyleDefs(79)  =   "Splits(0).Columns(10).HeadingStyle:id=25,.parent=14"
            _StyleDefs(80)  =   "Splits(0).Columns(10).FooterStyle:id=26,.parent=15"
            _StyleDefs(81)  =   "Splits(0).Columns(10).EditorStyle:id=27,.parent=17"
            _StyleDefs(82)  =   "Splits(0).Columns(11).Style:id=58,.parent=13,.alignment=2"
            _StyleDefs(83)  =   "Splits(0).Columns(11).HeadingStyle:id=55,.parent=14"
            _StyleDefs(84)  =   "Splits(0).Columns(11).FooterStyle:id=56,.parent=15"
            _StyleDefs(85)  =   "Splits(0).Columns(11).EditorStyle:id=57,.parent=17"
            _StyleDefs(86)  =   "Splits(0).Columns(12).Style:id=32,.parent=13,.alignment=2"
            _StyleDefs(87)  =   "Splits(0).Columns(12).HeadingStyle:id=29,.parent=14"
            _StyleDefs(88)  =   "Splits(0).Columns(12).FooterStyle:id=30,.parent=15"
            _StyleDefs(89)  =   "Splits(0).Columns(12).EditorStyle:id=31,.parent=17"
            _StyleDefs(90)  =   "Splits(0).Columns(13).Style:id=46,.parent=13,.alignment=1"
            _StyleDefs(91)  =   "Splits(0).Columns(13).HeadingStyle:id=43,.parent=14"
            _StyleDefs(92)  =   "Splits(0).Columns(13).FooterStyle:id=44,.parent=15"
            _StyleDefs(93)  =   "Splits(0).Columns(13).EditorStyle:id=45,.parent=17"
            _StyleDefs(94)  =   "Splits(0).Columns(14).Style:id=54,.parent=13,.alignment=1"
            _StyleDefs(95)  =   "Splits(0).Columns(14).HeadingStyle:id=51,.parent=14"
            _StyleDefs(96)  =   "Splits(0).Columns(14).FooterStyle:id=52,.parent=15"
            _StyleDefs(97)  =   "Splits(0).Columns(14).EditorStyle:id=53,.parent=17"
            _StyleDefs(98)  =   "Splits(0).Columns(15).Style:id=66,.parent=13"
            _StyleDefs(99)  =   "Splits(0).Columns(15).HeadingStyle:id=63,.parent=14"
            _StyleDefs(100) =   "Splits(0).Columns(15).FooterStyle:id=64,.parent=15"
            _StyleDefs(101) =   "Splits(0).Columns(15).EditorStyle:id=65,.parent=17"
            _StyleDefs(102) =   "Named:id=33:Normal"
            _StyleDefs(103) =   ":id=33,.parent=0"
            _StyleDefs(104) =   "Named:id=34:Heading"
            _StyleDefs(105) =   ":id=34,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(106) =   ":id=34,.wraptext=-1"
            _StyleDefs(107) =   "Named:id=35:Footing"
            _StyleDefs(108) =   ":id=35,.parent=33,.valignment=2,.bgcolor=&H8000000F&,.fgcolor=&H80000012&"
            _StyleDefs(109) =   "Named:id=36:Selected"
            _StyleDefs(110) =   ":id=36,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(111) =   "Named:id=37:Caption"
            _StyleDefs(112) =   ":id=37,.parent=34,.alignment=2"
            _StyleDefs(113) =   "Named:id=38:HighlightRow"
            _StyleDefs(114) =   ":id=38,.parent=33,.bgcolor=&H8000000D&,.fgcolor=&H8000000E&"
            _StyleDefs(115) =   "Named:id=39:EvenRow"
            _StyleDefs(116) =   ":id=39,.parent=33,.bgcolor=&HFFFF00&"
            _StyleDefs(117) =   "Named:id=40:OddRow"
            _StyleDefs(118) =   ":id=40,.parent=33"
            _StyleDefs(119) =   "Named:id=41:RecordSelector"
            _StyleDefs(120) =   ":id=41,.parent=34"
            _StyleDefs(121) =   "Named:id=42:FilterBar"
            _StyleDefs(122) =   ":id=42,.parent=33"
         End
         Begin VB.Line Line1 
            Index           =   0
            X1              =   3135
            X2              =   6945
            Y1              =   1455
            Y2              =   1470
         End
         Begin VB.Label Label4 
            Alignment       =   2  'Center
            Caption         =   "Documentos"
            BeginProperty Font 
               Name            =   "Courier New"
               Size            =   11.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00800000&
            Height          =   315
            Index           =   0
            Left            =   90
            TabIndex        =   3
            Top             =   30
            Width           =   11550
         End
      End
   End
   Begin VB.Menu Menu1 
      Caption         =   "Menu1"
      Visible         =   0   'False
      Begin VB.Menu Menu1_1 
         Caption         =   "&Activar"
      End
      Begin VB.Menu Menu1_2 
         Caption         =   "&Desactivar"
      End
      Begin VB.Menu Menu1_3 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_4 
         Caption         =   "Activar Todos Registros"
      End
      Begin VB.Menu Menu1_5 
         Caption         =   "Desactivar Todos Registros"
      End
      Begin VB.Menu Menu1_6 
         Caption         =   "-"
      End
      Begin VB.Menu Menu1_7 
         Caption         =   "Limpiar Filtro"
      End
   End
End
Attribute VB_Name = "FrmDocumentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RstDoc As New ADODB.Recordset
Dim RstDocTr As New ADODB.Recordset '--Para la transferencia
Dim QueHace As Integer
Dim SeEjecuto As Boolean
Dim Agregando As Boolean

Dim fOrdenLista As Boolean ''--especfica el orden de la lista de la consulta
Dim mIdRegistro& '--identificador del registro
Dim mMesActivo As Integer '--indica el mes activo
Dim fCierrePeriodo As Boolean '--indica si el periodo seleccionado esta cerrado o abierto (0 cerrado, -1 abierto)

Dim xHorIni As Date
Dim IdMenuActivo As Integer         'INDICA EL CODIGO DEL MENU ACTIVO

Dim mMesTrans As Integer
Dim RstTmpDet As New ADODB.Recordset
'Para mover el frame
Dim OrigFX As Long
Dim OrigFY As Long
Dim xOrden As Long '--indica el ordenamiento de los registros para transferir
                   '--segun seleccione los registros se asignara un numero


Sub CambiarMes()
    mMesActivo = SeleccionaMes(xCon)
    CargarGrid
    TabOne1.CurrTab = 0
End Sub

Sub CargarGrid()

    Dim nSQL As String
    Dim xCampo As String
    Dim nSQLFiltro As String

        '--limpiar los filtros
    TDB_FiltroLimpiar Dg1
    Set RstDoc = Nothing
    Set Dg1.DataSource = Nothing
    '----------------------
    OpcionesPeriodo
    '----------------------

    If TxtFchIni.Valor = "" Then
        MsgBox "Falta especificar la fecha de Inicio", vbInformation, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    If TxtFchFin.Valor = "" Then
        MsgBox "Falta especificar la fecha final", vbInformation, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        MsgBox "La fecha incial es superior a la fecha final", vbInformation, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If

    If Opt1.Value = True Then
        xCampo = "tra_documento.fchdoc "
    ElseIf Opt2.Value = True Then
        xCampo = "tra_documento.fchrecep "
    ElseIf Opt3.Value = True Then
        xCampo = "tra_documento.fchvenc "
    Else
        xCampo = "tra_documento.fchproc "
    End If

'    If ChkIniTransferencia.Value = 1 Then
'        nSQLFiltro = " and tra_documento.estado=0 and tra_documento.iddoc=0"
'    End If

    Set RstDoc = Nothing

    nSQL = "SELECT tra_documento.*, mae_documento.abrev, [tra_documento].[numser] & '-' & [tra_documento].[numdoc] AS numerodoc, mae_moneda.simbolo, tes_modulos.descripcion AS modulo,format(tra_documento.periodo,'00') as mes, IIf([tra_documento].[estado]=0,'PENDIENTE',IIf([tra_documento].[estado]=1,'OBSERVADO',IIf([tra_documento].[estado]=2,'TRANSFERIDO','RETIRADO'))) AS estado1, format(month(tra_documento.fchsist),'00') as messist " _
        + vbCr + " FROM ((tra_documento LEFT JOIN tes_modulos ON tra_documento.idmod = tes_modulos.id) LEFT JOIN mae_documento ON tra_documento.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON tra_documento.idmon = mae_moneda.id " _
        + vbCr + " WHERE (((" & xCampo & ") Between CDate('" & TxtFchIni.Valor & "') And CDate('" & TxtFchFin.Valor & "'))) " & nSQLFiltro _
        + vbCr + " ORDER BY tra_documento.fchdoc asc "
        'and tra_documento.estado<>3
    RST_Busq RstDoc, nSQL, xCon

    Dg1.DataSource = RstDoc
    
    TotalRegistros Dg1, RstDoc

End Sub



Private Sub cmd_periodo1_Click()
    mMesTrans = SeleccionaMes(xCon)
    LblPerIni.Caption = Busca_Codigo(mMesTrans, "id", "descripcion", "con_meses", "N", xCon)
    
    '--Si es apertura o Cierre => salir
    If mMesTrans = 0 Or mMesTrans = 13 Then
        MsgBox "Seleccione periodo nuevamente", vbInformation, xTitulo
        mMesTrans = 0
        LblPerIni.Caption = ""
        cmd_periodo1.SetFocus
        Exit Sub
    End If
    
        
    'Verificar si periodo esta cerrado
    Dim xMenu As Integer '--Utilizado para identificar el modulos para transferir (ver tabla var_menu)
    Dim xRst As New ADODB.Recordset
    Dim nSQL As String
    
    If OptCompra.Value = True Then
        xMenu = 218
    ElseIf OptHonorario.Value = True Then
        xMenu = 219
    Else
        xMenu = 0
    End If
    
    nSQL = "SELECT var_cierre.idform, var_cierre.idmes, var_cierre.estado " _
            & " From var_cierre " _
            & " WHERE (((var_cierre.idform)=" & xMenu & ") AND ((var_cierre.idmes)=" & mMesTrans & ")); "

    RST_Busq xRst, nSQL, xCon
    If xRst.RecordCount <> 0 Then
        If NulosN(xRst("estado")) = 0 Then
            MsgBox "El periodo para transferir está cerrado" & vbCr & "Cambie de periodo o Solicite a encargado habilite el periodo seleccionado", vbInformation, xTitulo
            mMesTrans = 0
            LblPerIni.Caption = ""
        End If
    End If
    
    Set xRst = Nothing
    
End Sub



Private Sub CmdConsultar_Click()
    CargarGrid
End Sub

Private Sub CmdConsultarTr_Click()

    Dim nSQL As String
    Dim xCampo As String
    Dim nSQLFiltro As String

        '--limpiar los filtros
    TDB_FiltroLimpiar Dg2
    Set RstDocTr = Nothing
    Set Dg2.DataSource = Nothing

    If TxtFchIniTr.Valor = "" Then
        MsgBox "Falta especificar la fecha de Inicio", vbInformation, xTitulo
        TxtFchIniTr.SetFocus
        Exit Sub
    End If
    If TxtFchFinTr.Valor = "" Then
        MsgBox "Falta especificar la fecha final", vbInformation, xTitulo
        TxtFchFinTr.SetFocus
        Exit Sub
    End If
    If CDate(TxtFchIniTr.Valor) > CDate(TxtFchFinTr.Valor) Then
        MsgBox "La fecha incial es superior a la fecha final", vbInformation, xTitulo
        TxtFchIniTr.SetFocus
        Exit Sub
    End If

    If Opt1Tr.Value = True Then
        xCampo = "tra_documento.fchdoc "
    ElseIf Opt2Tr.Value = True Then
        xCampo = "tra_documento.fchrecep "
    ElseIf Opt3Tr.Value = True Then
        xCampo = "tra_documento.fchvenc "
    Else
        xCampo = "tra_documento.fchproc "
    End If

    '--filtrando el tipo de documento x módulo
    If OptCompra.Value = True Then
        nSQLFiltro = " and tra_documento.tipdoc <>2 "
    ElseIf OptHonorario.Value = True Then
        nSQLFiltro = " and tra_documento.tipdoc =2 "
    End If

    Set RstDocTr = Nothing
    '--se muestran los documentos listos a transferir, no se muestran los documentos observados
    Dim xRs As New ADODB.Recordset
    nSQL = "SELECT '0' AS xsel, 0 as xorden, tra_documento.id,tra_documento.corr, tra_documento.rucprov, tra_documento.proveedor, mae_documento.abrev, [tra_documento].[numser] & '-' & [tra_documento].[numdoc] AS numerodoc, " _
            + vbCr + " tra_documento.fchdoc & '' as fchdoc1, tra_documento.fchrecep & '' as fchrecep1, tra_documento.fchvenc & '' as fchvenc1, tra_documento.fchproc & '' as fchproc1, mae_moneda.simbolo, tra_documento.tipcam & '' as tipcam1, tra_documento.imptot & '' as imptot1, tra_documento.glosa, '' AS numreg, '' AS modulo,tra_documento.anno,format(tra_documento.periodo,'00') as mes " _
            + vbCr + " FROM (tra_documento LEFT JOIN mae_documento ON tra_documento.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON tra_documento.idmon = mae_moneda.id " _
            + vbCr + " WHERE (((" & xCampo & ") Between CDate('" & TxtFchIniTr.Valor & "') And CDate('" & TxtFchFinTr.Valor & "')) AND ((tra_documento.estado)=0) AND ((tra_documento.iddoc)=0))  " & nSQLFiltro _
            + vbCr + " ORDER BY tra_documento.fchdoc asc"

    RST_Busq xRs, nSQL, xCon

    DEFINIR_RST_TMP RstDocTr, xRs

    CARGAR_RST_TMP RstDocTr, xRs

    Dg2.DataSource = RstDocTr


End Sub

Private Sub CmdDepura_Click()
    DepurarDatos
    VerDatosxDepurar
End Sub

Private Sub CmdDetalle_Click()
    VerDatosxDepurarDet
End Sub

Private Sub CmdExportar_Click()
    Dim oExport As New SGI2_funciones.formularios
    Dim xFg As Object
    Dim nTitulo As String

    If TabOne3.CurrTab = 0 Then
        nTitulo = "Datos Observados de Proveedores"
        Set xFg = Fg3
    ElseIf TabOne3.CurrTab = 0 Then
        nTitulo = "Datos Observados de Item's"
        Set xFg = Fg4
    Else
        nTitulo = "Datos Observados de Tipos de Documentos"
        Set xFg = Fg5
    End If

    Me.MousePointer = vbHourglass
    oExport.VSFlexGrid_Exportar_MSExcel xCon, xFg, "Transferencia de Documentos", nTitulo, "", nTitulo
    Set oExport = Nothing
    Me.MousePointer = vbDefault

End Sub

Private Sub CmdSalir_Click()
    FraDepura.Visible = False
End Sub

Private Sub Command1_Click()
    Dim oExport As New SGI2_funciones.formularios
    Dim xFg As Object
    Dim nTitulo As String

    nTitulo = "Detalle de documentos observados"

    Me.MousePointer = vbHourglass
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg7, "Transferencia de Documentos", nTitulo, "", nTitulo
    Set oExport = Nothing
    Me.MousePointer = vbDefault
End Sub

Private Sub Command2_Click()
    FraDetalle.Visible = False
End Sub


Private Sub Dg1_DblClick()
    TabOne1.CurrTab = 1
End Sub

Private Sub Dg1_FilterChange()
    TDB_FiltroGenerar Dg1, RstDoc
    
    TotalRegistros Dg1, RstDoc
    
End Sub

Private Sub Dg1_HeadClick(ByVal ColIndex As Integer)

    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstDoc.Sort = CStr(Dg1.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear

End Sub

Private Sub Dg1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        TabOne1.CurrTab = 1
    End If
End Sub

Private Sub Dg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 123 And TabOne1.CurrTab = 0 Then
        VerMovimientos1 IdMenuActivo, NulosN(RstDoc("id")), xCon
    End If
End Sub


Private Sub Dg2_HeadClick(ByVal ColIndex As Integer)
    On Error Resume Next
    Dim nOrden As String
    If fOrdenLista = False Then nOrden = "ASC"
    If fOrdenLista = True Then nOrden = "DESC"
    RstDocTr.Sort = CStr(Dg2.Columns(ColIndex).DataField) & " " & nOrden
    fOrdenLista = Not fOrdenLista
    Err.Clear

End Sub

Private Sub Fg1_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 45 Then
'        CmdAddDoc_Click
    End If

    If KeyCode = 46 Then
'        CmdDelDoc_Click
    End If
End Sub


Private Sub Fg2_CellChanged(ByVal Row As Long, ByVal Col As Long)

    '---------------------------------

    If Agregando = True Then Exit Sub
'
'    If Col = 3 Then
'        Dim Factor As Double
'
'        '--verificar que el plazo no supere el limite
'        If NulosN(Fg2.TextMatrix(Fg2.Row, 3)) > NulosN(TxtDiasPlazo.Text) Then
'            MsgBox "La plazo ingresado supera el plazo limite", vbInformation, xTitulo
'            Fg2.TextMatrix(Fg2.Row, 3) = "0"
'
'        End If
'
'
'        Factor = HallarFactor(NulosN(TxtImpTasa.Text), NulosN(Fg2.TextMatrix(Fg2.Row, 3)), NulosN(TxtTipInt.Text))
'        '--interes
'        Fg2.TextMatrix(Fg2.Row, 5) = Format(Fg2.TextMatrix(Fg2.Row, 4) * Factor, FORMAT_MONTO)
'        '--porte
'        Fg2.TextMatrix(Fg2.Row, 6) = Format(NulosN(TxtPortes.Text), FORMAT_MONTO)
'        '--igv
'        Fg2.TextMatrix(Fg2.Row, 7) = ((NulosN(Fg2.TextMatrix(Fg2.Row, 5)) + NulosN(Fg2.TextMatrix(Fg2.Row, 6))) * xIGV)
'        Fg2.TextMatrix(Fg2.Row, 7) = Format(Fg2.TextMatrix(Fg2.Row, 7), FORMAT_MONTO)
'        '--importe total
'        If ChkRetencion.value = 1 Then
'            'capital+interes+porte+igv
'            '--cuando es agente retenedor se le tendra que aplicar menos el % al (interes+porte+igv)
'            Fg2.TextMatrix(Fg2.Row, 8) = NulosN(Fg2.TextMatrix(Fg2.Row, 4)) + (NulosN(Fg2.TextMatrix(Fg2.Row, 5)) + NulosN(Fg2.TextMatrix(Fg2.Row, 6)) + NulosN(Fg2.TextMatrix(Fg2.Row, 7))) * (1 - xRetencion)
'        Else
'            Fg2.TextMatrix(Fg2.Row, 8) = NulosN(Fg2.TextMatrix(Fg2.Row, 4)) + NulosN(Fg2.TextMatrix(Fg2.Row, 5)) + NulosN(Fg2.TextMatrix(Fg2.Row, 6)) + NulosN(Fg2.TextMatrix(Fg2.Row, 7))
'        End If
'        Fg2.TextMatrix(Fg2.Row, 8) = Format(Fg2.TextMatrix(Fg2.Row, 8), FORMAT_MONTO)
'
'
'        If NulosN(Fg2.TextMatrix(Fg2.Row, 3)) <> 0 Then
'            If NulosC(TxtFchIni.Valor) <> "" Then
'                Fg2.TextMatrix(Fg2.Row, 9) = CDate(TxtFchIni.Valor) + NulosN(Fg2.TextMatrix(Fg2.Row, 3))
'                Fg2.TextMatrix(Fg2.Row, 9) = Format(Fg2.TextMatrix(Fg2.Row, 9), "dd/mm/yy")
'            End If
'        Else
'            Fg2.TextMatrix(Fg2.Row, 9) = ""
'        End If
'
'        SumarColumnas
'
'    ElseIf Col = 9 Then
'        '--verificar si la fecha es correcta
'        If IsDate(Fg2.TextMatrix(Fg2.Row, 9)) = False Then
'            MsgBox "Fecha incorrecta, ingrese ne nuevo la Fecha", vbInformation, xTitulo
'            Fg2.TextMatrix(Fg2.Row, 9) = ""
'            Fg2.SetFocus
'            Exit Sub
'        End If
'        '--verificar si la fecha de vencimiento es menor a la fecha de inicio
'        If CDate(Fg2.TextMatrix(Fg2.Row, 9)) <= CDate(TxtFchIni.Valor) Then
'            MsgBox "La fecha de Vencimiento debe ser mayor a la fecha de inicio", vbInformation, xTitulo
'            Fg2.TextMatrix(Fg2.Row, 9) = ""
'            Fg2.SetFocus
'        End If
'        '--verificar que la diferencia de fecha de vencimiento con fecha de inicio no supere el plazo limite
'        If CDate(Fg2.TextMatrix(Fg2.Row, 9)) - CDate(TxtFchIni.Valor) - 1 > NulosN(TxtDiasPlazo.Text) Then
'            MsgBox "La fecha de Vencimiento supera el plazo limite", vbInformation, xTitulo
'            Fg2.TextMatrix(Fg2.Row, 3) = ""
'            Fg2.SetFocus
'        End If
'
'        Fg2.TextMatrix(Fg2.Row, 3) = CDate(Fg2.TextMatrix(Fg2.Row, 9)) - CDate(TxtFchIni.Valor)
'        Fg2_CellChanged Row, 3
'
'        SumarColumnas
'
'    End If
End Sub

Private Sub Fg2_EnterCell()
    If QueHace = 3 Then
        Fg2.Editable = flexEDNone
        Exit Sub
    End If

     If Fg2.Col = 3 Or Fg2.Col = 9 Then
        Fg2.SelectionMode = flexSelectionFree

        Fg2.Editable = flexEDKbdMouse
    Else
        Fg2.SelectionMode = flexSelectionByRow
        Fg2.Editable = flexEDNone
     End If
End Sub

Private Sub Fg3_EnterCell()
    If Fg3.Col = 2 Or Fg3.Col = 3 Then
        Fg3.Editable = flexEDKbdMouse
    Else
        Fg3.Editable = flexEDNone
    End If
End Sub

Private Sub Fg3_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Fg5_EnterCell()
    If Fg5.Col > 1 And Fg5.Col < 6 Then
        Fg5.Editable = flexEDKbdMouse
    Else
        Fg5.Editable = flexEDNone
    End If
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        
        mMesActivo = xMes

        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu

        TxtFchIni.Valor = Date
        TxtFchFin.Valor = Date
        TxtFchIniTr.Valor = Date
        TxtFchFinTr.Valor = Date
        
        TabOne1.TabVisible(2) = False
        TabOne1.TabVisible(3) = False

        '--Almacenar temporalmente el codigo del menu
        IdMenuActivo = xIdMenu
        
        '--bloquear accesos
        OpcionesUsuario xIdUsuario, IdMenuActivo, Toolbar1, xCon

        SeEjecuto = True

        '--mostrar datos de los registros observados
        VerDatosxDepurar
        
        '--provisional, habilitar opcion para desagregar de glosa tipos de gasto reembolsable.
        If xIdUsuario = 1 Or xIdUsuario = 4 Then
            CmdActualizar.Visible = True
            CmdActualizar1.Visible = True
        Else
            CmdActualizar.Visible = False
            CmdActualizar1.Visible = False
        End If
        
        '--aplicando formato para mostrar el total de registros encontrados
        Dg1.Splits(0).Columns("numerodoc").FooterFont.Bold = True
        Dg1.Splits(0).Columns("fchdoc").FooterFont.Bold = True
        Dg1.Splits(0).Columns("fchdoc").FooterAlignment = dbgRight
        Dg1.Splits(0).Columns("numerodoc").FooterText = "Cant. Registros"
        Dg1.Splits(0).Columns("fchdoc").FooterText = "0"
        '---
        
    End If
End Sub

Sub Nuevo()
    QueHace = 1
'    TabOne1.CurrTab = 1
    TabOne1.TabEnabled(0) = False
    ActivarTool
    Blanquea
    Bloquea

    Label5.Caption = "Agregando Importación"
    Label7.Caption = "Agregando Transferencia"
    Label8.Caption = "Agregando Documento"

    xHorIni = Time
    
    If TabOne1.CurrTab = 1 Then

    ElseIf TabOne1.CurrTab = 2 Then
        TxtArchivo.SetFocus
    Else
        cmd_periodo1.SetFocus
    End If
    
    xOrden = 1
    
End Sub

Sub Modificar()
'    QueHace = 2
'    TabOne1.CurrTab = 1
'    TabOne1.TabEnabled(0) = False
'    ActivarTool
'    Blanquea
'    Bloquea
'    Label5.Caption = "Modificando Letra"
'    Fg2.Rows = 1
'    MuestraSegundoTab
'    xHorIni = Time
'
'    TxtRucPro.SetFocus
End Sub

Sub ActivarTool()
    Dim A&
    For A = 1 To Toolbar1.Buttons.Count
        Toolbar1.Buttons(A).Enabled = Not Toolbar1.Buttons(A).Enabled
    Next A
End Sub

Private Sub Form_Load()
    SeEjecuto = False
    QueHace = 3
    Fg2.ColWidth(10) = 0
    TabOne1.CurrTab = 0

    Dg2.Columns("fchdoc1").NumberFormat = FORMAT_DATE
    Dg2.Columns("fchrecep1").NumberFormat = FORMAT_DATE
    Dg2.Columns("fchvenc1").NumberFormat = FORMAT_DATE
    Dg2.Columns("fchproc1").NumberFormat = FORMAT_DATE
    Dg2.Columns("imptot1").NumberFormat = FORMAT_MONTO

    
    Fg1.ColWidth(1) = 0
    Fg1.ColWidth(2) = 0

    Fg6.ColWidth(7) = 0
    
    
    '---------------
    Me.WindowState = 2
    Me.Width = 12000
    Me.Height = 8200
    '---------------
    

End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub

    If Me.Height > 3000 Then
        TabOne1.Top = 345
        TabOne1.Width = Me.Width - 100
        TabOne1.Height = Me.Height - 500
        '--consulta
        Dg1.Top = 900
        Dg1.Width = Me.Width - 220
        Dg1.Height = Me.Height - 2180
        
        '--importacion
        TabOne2.Top = 1230
        TabOne2.Width = Me.Width - 220
        TabOne2.Height = Me.Height - 2500
        
        '--transferencia
        Dg2.Top = 1230
        Dg2.Width = Me.Width - 220
        Dg2.Height = Me.Height - 2500
        
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If QueHace <> 3 Then
        MsgBox "No puede salir del formulario miestras este agregando o editando un registro", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub pic_Click()
    CmdSalir_Click
End Sub

Private Sub pic1_Click()
    Command2_Click
End Sub

Private Sub TabOne1_Switch(OldTab As Integer, NewTab As Integer, Cancel As Integer)
    If OldTab = 0 Then
        If QueHace <> 1 Then
            If RstDoc.State = 0 Then
                Cancel = 1
                Exit Sub
            End If
            If RstDoc.RecordCount = 0 Then
                Cancel = 1
                Exit Sub
            End If

            MuestraSegundoTab
        End If
    End If
End Sub

Sub Bloquea()
'    TxtArchivo.Locked = Not TxtArchivo.Locked
'    TxtArchivo2.Locked = Not TxtArchivo2.Locked
End Sub

Sub MuestraSegundoTab()

    Blanquea
    If RstDoc.EOF = True Or RstDoc.BOF = True Or RstDoc.RecordCount = 0 Then Exit Sub
    
    TxtCorr.Text = NulosC(RstDoc("corr"))
    TxtAnno.Text = NulosN(RstDoc("anno"))
    TxtPeriodo.Text = NomMes(NulosN(RstDoc("periodo")))
    
    TxtFchSist.Valor = Format(RstDoc("fchsist"), "dd/mm/yyyy")
    TxtNumRuc.Text = NulosC(RstDoc("rucprov"))
    TxtNomProv.Text = NulosC(RstDoc("proveedor"))
    TxtFchDoc.Valor = Format(RstDoc("fchdoc"), "dd/mm/yyyy")
    TctFchRecep.Valor = Format(RstDoc("fchdoc"), "dd/mm/yyyy")
    TxtFchVen.Valor = Format(RstDoc("fchvenc"), "dd/mm/yyyy")
    TxtTipoDoc.Text = NulosC(RstDoc("documento"))
    TxtMoneda.Text = NulosC(RstDoc("moneda"))
    TxtTC.Text = NulosN(RstDoc("tipcam"))
    TxtNumSer.Text = NulosC(RstDoc("numser"))
    TxtNumDoc.Text = NulosC(RstDoc("numdoc"))
    TxtGlosa.Text = NulosC(RstDoc("glosa"))
    TxtRucCli.Text = NulosC(RstDoc("ruccliente"))
    TxtNomCli.Text = NulosC(RstDoc("cliente"))
    TxtNumOrden.Text = NulosC(RstDoc("numorden"))
    
    '--datos de la transferencia
    TxtEstado.Text = NulosC(RstDoc("estado1"))
    TxtModulo.Text = NulosC(RstDoc("modulo"))
    TxtNumReg.Text = NulosC(RstDoc("numreg"))
    
    TxtBruto.Text = Format(NulosN(RstDoc("impafec")), FORMAT_MONTO)
    TxtInafecto.Text = Format(NulosN(RstDoc("impexon")), FORMAT_MONTO)
    TxtIGV.Text = Format(NulosN(RstDoc("impigv")), FORMAT_MONTO)
    TxtOtros.Text = Format(NulosN(RstDoc("impret")), FORMAT_MONTO)
    TxtTotal.Text = Format(NulosN(RstDoc("imptot")), FORMAT_MONTO)
    
    '--ver si proveedor es observado
    If NulosN(RstDoc("idprov")) = 0 Then
        TxtNumRuc.BackColor = vbRed
    Else
        TxtNumRuc.BackColor = vbWhite
    End If
    '--ver si proveedor es observado
    If NulosN(RstDoc("tipdoc")) = 0 Then
        TxtTipoDoc.BackColor = vbRed
    Else
        TxtTipoDoc.BackColor = vbWhite
    End If
    
    
'    'CARGAMOS LOS DOCUMENTOS
    Dim xRst As New ADODB.Recordset
    RST_Busq xRst, "SELECT tra_documentodet.*, alm_inventario.codpro, alm_inventario.descripcion " _
                + vbCr + " FROM tra_documentodet LEFT JOIN alm_inventario ON tra_documentodet.iditem = alm_inventario.id " _
                + vbCr + " WHERE (((tra_documentodet.iddet)=" & RstDoc("id") & ")); ", xCon

    Fg6.Rows = 1

    If xRst.RecordCount <> 0 Then xRst.MoveFirst

    Do While Not xRst.EOF
        Fg6.Rows = Fg6.Rows + 1
        Fg6.TextMatrix(Fg6.Rows - 1, 1) = NulosC(xRst("vinc1"))
        Fg6.TextMatrix(Fg6.Rows - 1, 2) = NulosC(xRst("vinc2"))
        Fg6.TextMatrix(Fg6.Rows - 1, 3) = NulosC(xRst("vinc3"))
        Fg6.TextMatrix(Fg6.Rows - 1, 4) = NulosC(xRst("vinc4"))
        Fg6.TextMatrix(Fg6.Rows - 1, 5) = Format(NulosN(xRst("impafec")), FORMAT_MONTO)
        Fg6.TextMatrix(Fg6.Rows - 1, 6) = Format(NulosN(xRst("impexon")), FORMAT_MONTO)
        Fg6.TextMatrix(Fg6.Rows - 1, 7) = Format(NulosN(xRst("imptot")), FORMAT_MONTO)
        Fg6.TextMatrix(Fg6.Rows - 1, 8) = NulosC(xRst("codpro"))
        Fg6.TextMatrix(Fg6.Rows - 1, 9) = NulosC(xRst("descripcion"))
        
        If NulosN(xRst("iditem")) = 0 Then
            GRID_COLOR_FONDO Fg6, Fg6.Rows - 1, 8, Fg6.Rows - 1, 9, vbRed
        End If
        
        xRst.MoveNext
    Loop
    
    Set xRst = Nothing
    
    
    
End Sub

Sub Blanquea()

    '--Importación
    TxtArchivo.Text = ""
    TxtArchivo2.Text = ""

    Fg1.Rows = 1
    Fg2.Rows = 1

    
    
    '--Transferencia
    Set RstDocTr = Nothing
    Set Dg2.DataSource = Nothing
    TDB_FiltroLimpiar Dg2
    
    LblPerIni.Caption = ""
    mMesTrans = 0
    
    
    TxtAnno.Text = ""
    TxtPeriodo.Text = ""
    TxtFchSist.Valor = ""
    TxtNumRuc.Text = ""
    TxtNomProv.Text = ""
    TxtFchDoc.Valor = ""
    TctFchRecep.Valor = ""
    TxtFchVen.Valor = ""
    TxtTipoDoc.Text = ""
    TxtMoneda.Text = ""
    TxtTC.Text = ""
    TxtNumSer.Text = ""
    TxtNumDoc.Text = ""
    TxtGlosa.Text = ""
    TxtRucCli.Text = ""
    TxtNomCli.Text = ""
    TxtNumOrden.Text = ""
    TxtEstado.Text = ""
    TxtModulo.Text = ""
    TxtNumReg.Text = ""
    TxtBruto.Text = ""
    TxtInafecto.Text = ""
    TxtIGV.Text = ""
    TxtOtros.Text = ""
    TxtTotal.Text = ""
    
    Fg6.Rows = 1

End Sub

Sub Eliminar()
    Dim Rpta As Integer
    If RstDoc.State = 0 Then Exit Sub
    If RstDoc.EOF = True Or RstDoc.BOF = True Or RstDoc.RecordCount = 0 Then
        MsgBox "No hay registro para eliminar", vbExclamation, xTitulo
        Exit Sub
    End If
    
    If RstDoc("estado") = 2 Then
        MsgBox "El documento no se puede eliminar, fue transferido" & vbCr & "Módulo    :  " & NulosC(RstDoc("modulo")) & vbCr & "Nª. Reg. : " & NulosC(RstDoc("numreg")), vbInformation, xTitulo
        Exit Sub
    End If

    Rpta = MsgBox("Esta seguro de eliminar el documento seleccionado", vbQuestion + vbYesNo + vbDefaultButton1, xTitulo)
    If Rpta = vbYes Then

        xCon.Execute "DELETE * FROM tra_documentodet WHERE iddet = " & RstDoc("id") & ""
        xCon.Execute "DELETE * FROM tra_documento WHERE id = " & RstDoc("id") & ""

        'Eliminar historial del registro
        xCon.Execute "DELETE * FROM var_edicion WHERE idmov = " & RstDoc("id") & " AND idform = " & IdMenuActivo

        RstDoc.Requery
        Dg1.Refresh
    End If
End Sub


Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
'    If Button.Index = 1 Then Nuevo

    If Button.Index = 2 Then Modificar

    If Button.Index = 3 Then Eliminar

    If Button.Index = 5 Then
        If Grabar = True Then

            Cancelar
            If RstDoc.State = 0 Then Exit Sub
            RstDoc.Requery
            Dg1.Refresh
            If RstDoc.RecordCount <> 0 Then RstDoc.MoveFirst
            RstDoc.Find "id = " & mIdRegistro & ""
            If RstDoc.EOF = True And RstDoc.RecordCount <> 0 Then RstDoc.MoveFirst
        
        End If
    End If

    If Button.Index = 6 Then Cancelar

    If Button.Index = 9 Then '--actualizar
        TDB_Actualizar Me, TabOne1, Dg1, RstDoc
        TotalRegistros Dg1, RstDoc
    End If
    
    If Button.Index = 10 Then
        DepurarDatos
        VerDatosxDepurar
    End If

    If Button.Index = 12 Then
        CambiarMes
    End If
    
    If Button.Index = 14 Then pExportar
    
    If Button.Index = 18 Then
        Set RstDoc = Nothing
        Set RstDocTr = Nothing
        Unload Me
    End If
End Sub

Function Grabar() As Boolean

    On Error GoTo LaCague

    Dim xId As Double
    Dim nSQL As String
    Dim xRst As New ADODB.Recordset
    Dim xRstCab As New ADODB.Recordset
    Dim xRstDet As New ADODB.Recordset
    Dim xIgvTasa As Double '--indica la tasa aplicada
    Dim xNumAsiento As String '--Numero de registro contable
    Dim A, B As Long
    Dim xCorr As Integer '--Utilizado para diferenciar el detalle del documento, se rinicia cuando cambia el documento

    '--mostrar barra de tareas
    Frame5.Left = 3090
    Frame5.Top = 2910
    Frame5.Visible = True
    

    '--grabar importacion de datos
    If TabOne1.CurrTab = 2 Then
        
        ReDim xCampos(26, 4) As String
        ReDim xCampos2(10, 4) As String
        Dim xIdProc As Double '--Codigo del proceso de importacion
        Dim xRuc, xTipoDoc, xNumSer, xNumDoc As String
        
        If Fg1.Rows = Fg1.FixedRows Or Fg2.Rows = Fg2.FixedRows Then
            MsgBox "No hay registros para importar", vbInformation, xTitulo
            Frame5.Visible = False
            Exit Function
        End If

        ProgressBar2.Max = Fg1.Rows - 1
        LblBarra.Caption = "Importando Documentos"
        DoEvents
        
        xCon.BeginTrans
        
        RST_Busq xRstCab, "Select TOP 1 * from tra_documento", xCon
        RST_Busq xRstDet, "Select TOP 1 * from tra_documentodet", xCon
        
        '--buscar codigo del proceso
        xIdProc = NulosN(HallaCodigoTabla("tra_documento", xCon, "idproc"))
        
        For A = 1 To Fg1.Rows - 1

            ProgressBar2.Value = A
            
            DoEvents

            xRuc = NulosC(Fg1.TextMatrix(A, 4))
            xTipoDoc = NulosC(Fg1.TextMatrix(A, 6))
            xNumSer = NulosC(Fg1.TextMatrix(A, 7))
            xNumDoc = NulosC(Fg1.TextMatrix(A, 8))

            '--verificar si documento esta registrado en seven
            nSQL = "SELECT tra_documento.id From tra_documento " _
               + vbCr + " WHERE (((tra_documento.rucprov)='" & xRuc & "') AND " _
                            & " ((tra_documento.documento)='" & xTipoDoc & "') AND " _
                            & " ((tra_documento.numser)='" & xNumSer & "') AND " _
                            & " ((tra_documento.numdoc)='" & xNumDoc & "')); "
            Set xRst = Nothing

            RST_Busq xRst, nSQL, xCon

            
            If xRst.RecordCount = 0 Then
                
                '--Filtro del detalle
                RstTmpDet.Filter = ""
                RstTmpDet.Filter = "numruc='" & xRuc & "' and documento='" & xTipoDoc & "' and numser='" & xNumSer & "' and numdoc='" & xNumDoc & "'"
                
                If RstTmpDet.RecordCount > 0 Then
                    xId = HallaCodigoTabla("tra_documento", xCon, "id")
    
                    xRstCab.AddNew
                    xRstCab("id") = Str(xId)
                    xRstCab("anno") = NulosN(Fg1.TextMatrix(A, 1))
                    xRstCab("periodo") = NulosN(Fg1.TextMatrix(A, 2))
                    xRstCab("corr") = NulosC(Fg1.TextMatrix(A, 3))
                    xRstCab("rucprov") = xRuc
                    xRstCab("proveedor") = NulosC(Fg1.TextMatrix(A, 5))
                    xRstCab("documento") = xTipoDoc
                    xRstCab("numser") = xNumSer
                    xRstCab("numdoc") = xNumDoc
                    xRstCab("fchdoc") = NulosC(Fg1.TextMatrix(A, 9))
                    xRstCab("fchrecep") = NulosC(Fg1.TextMatrix(A, 10))
                    xRstCab("fchvenc") = NulosC(Fg1.TextMatrix(A, 11))
                    xRstCab("fchsist") = NulosC(Fg1.TextMatrix(A, 12))
                    xRstCab("moneda") = NulosC(Fg1.TextMatrix(A, 13))
                    xRstCab("tipcam") = NulosN(Fg1.TextMatrix(A, 14))
                    xRstCab("impafec") = NulosN(Fg1.TextMatrix(A, 15))
                    xRstCab("impexon") = NulosN(Fg1.TextMatrix(A, 16))
                    xRstCab("impigv") = NulosN(Fg1.TextMatrix(A, 18))
                    xRstCab("impret") = NulosN(Fg1.TextMatrix(A, 17))
                    xRstCab("imptot") = NulosN(Fg1.TextMatrix(A, 19))
                    xRstCab("glosa") = NulosC(Fg1.TextMatrix(A, 20))
                    xRstCab("ruccliente") = NulosC(Fg1.TextMatrix(A, 21))
                    xRstCab("cliente") = NulosC(Fg1.TextMatrix(A, 22))
                    xRstCab("numorden") = NulosC(Fg1.TextMatrix(A, 23))
                    xRstCab("idproc") = xIdProc
                    xRstCab("fchproc") = Date
                    xRstCab("estado") = 1
                    xRstCab.Update
                    '---------------------------------
                    'Grabar el detalle del documento
                    xCorr = 1
                    
                    RstTmpDet.MoveFirst
                    
                    Do While Not RstTmpDet.EOF
                        
                        xRstDet.AddNew
                        xRstDet("iddet") = Str(xId)
                        xRstDet("corr") = Str(xCorr)
                        xRstDet("impafec") = NulosN(RstTmpDet("impafec"))
                        xRstDet("impexon") = NulosN(RstTmpDet("impexon"))
                        xRstDet("impigv") = NulosN(RstTmpDet("impigv"))
                        xRstDet("impret") = NulosN(RstTmpDet("impret"))
                        xRstDet("imptot") = NulosN(RstTmpDet("imptot"))
                        xRstDet("vinc1") = NulosC(RstTmpDet("vinc1"))
                        xRstDet("vinc2") = NulosC(RstTmpDet("vinc2"))
                        xRstDet("vinc3") = NulosC(RstTmpDet("vinc3"))
                        xRstDet("vinc4") = NulosC(RstTmpDet("vinc4"))
                        xRstDet("iditem") = 0
                        xRstDet.Update
                        
                        xCorr = xCorr + 1

                        RstTmpDet.MoveNext

                    Loop
                Else
                
'                MsgBox "no hay registro detalle"

                End If
                
                
                
                

            End If

        Next A
        
        xCon.CommitTrans

'''    'grabamos el movimiento en la tabla var_edicion
'''    GrabarOperacion xIdUsuario, IdMenuActivo, QueHace, xHorIni, Time, Date, xCon, xId
    
    '--Depurar datos
            
        DepurarDatos
    
        VerDatosxDepurar

        MsgBox "La Importación se completó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    
    
    '--Grabar transferencia
    ElseIf TabOne1.CurrTab = 3 Then

                Dim xFechaReg  As Date '--Fecha de registro del sistema

                If mMesTrans = 0 Then
                    MsgBox "Falta especificar el periodo de transferencia", vbInformation, xTitulo
                    cmd_periodo1.SetFocus
                    Frame5.Visible = False
                    Exit Function
                End If
                
                '--Filtrando los documentos para transferir
                RstDocTr.Filter = "xsel=-1"
                
                If RstDocTr.RecordCount = 0 Then
                    MsgBox "Para transferir necesita seleccionar documentos", vbInformation, xTitulo
                    RstDocTr.Filter = "" '--Quitar filtro
                    Frame5.Visible = False
                    Exit Function
                End If
                
                
                
                If MsgBox("Seguro desea transferir los Documentos", vbYesNo + vbQuestion, xTitulo) = vbNo Then
                    RstDocTr.Filter = "" '--Quitar filtro
                    Frame5.Visible = False
                    Exit Function
                End If
                
                RstDocTr.Sort = "xorden asc"
                
                ProgressBar2.Max = RstDocTr.RecordCount
                
                '--Determinar fecha de registro
                xFechaReg = CDate("01/" & Format(mMesTrans, "00") & "/" & AnoTra)

                ProgressBar2.Value = 0
                DoEvents

                Dim xRstCab1 As New ADODB.Recordset '--Rst TEmporal para Cabecera
                Dim xRstDet1 As New ADODB.Recordset '--Rst TEmporal para Detalle

                xCon.BeginTrans

                If OptCompra.Value = True Then
                    '--Grabando Registro de Compra
                    LblBarra.Caption = "Transfiriendo Compras"
                    
                    DoEvents
                    
                    RST_Busq xRstCab, "Select TOP 1 * from com_compras", xCon
                    RST_Busq xRstDet, "Select TOP 1 * from com_comprasdet", xCon
                                       
                    Do While Not RstDocTr.EOF
                        Set xRstCab1 = Nothing
                        Set xRstDet1 = Nothing
                        
                        ProgressBar2.Value = ProgressBar2.Value + 1
                        
                        '--Cargar datos del documento
                        RST_Busq xRstCab1, "Select * from tra_documento where id = " & RstDocTr("id") & "", xCon
                        
                        nSQL = "SELECT tra_documentodet.*, alm_inventario.idcuenta AS idcue " _
                                    + vbCr + " FROM tra_documentodet INNER JOIN alm_inventario ON tra_documentodet.iditem = alm_inventario.id " _
                                    + vbCr + " WHERE (((tra_documentodet.iddet)=" & RstDocTr("id") & ")); "
                        RST_Busq xRstDet1, nSQL, xCon

                        xIgvTasa = 0

                        xId = HallaCodigoTabla("com_compras", xCon, "id")
                        
                        xRstCab.AddNew
                        xRstCab("id") = Str(xId):
                        xRstCab("idlib") = 1:
                        xRstCab("idtipo") = 5:
                        xRstCab("tipdoc") = NulosN(xRstCab1("tipdoc")):
                        xRstCab("idpro") = NulosN(xRstCab1("idprov")):
                        xRstCab("numser") = Format(NulosC(xRstCab1("numser")), "0000")
                        xRstCab("numdoc") = Format(NulosC(xRstCab1("numdoc")), "0000000000")
                        xRstCab("fchreg") = xFechaReg:
                        xRstCab("fchdoc") = NulosC(xRstCab1("fchdoc")):
                        xRstCab("fchven") = NulosC(xRstCab1("fchvenc")):
                        xRstCab("fchpag") = NulosC(xRstCab1("fchdoc")):
                        xRstCab("idconpag") = 1:
                        xRstCab("idmon") = NulosC(xRstCab1("idmon")):
                        xRstCab("impbru") = NulosN(xRstCab1("impafec")):
                        xRstCab("impina") = NulosN(xRstCab1("impexon")):
                        xRstCab("impigv") = NulosN(xRstCab1("impigv")):
                        xRstCab("imptot") = NulosN(xRstCab1("imptot")):
                        xRstCab("impsal") = NulosN(xRstCab1("imptot")):
                        xRstCab("importado") = -1:
                        xRstCab("idalm") = 1:
                        xRstCab("tipcom") = 1:
                        xRstCab("glosa") = NulosC(xRstCab1("glosa")):
                        xRstCab("tc") = NulosN(xRstCab1("tipcam")):
                        xRstCab("numerodocref") = NulosC(xRstCab1("numorden")):
                        xRstCab("fchrecep") = NulosC(xRstCab1("fchrecep")):
                        xRstCab("idcli") = 0:

                        If NulosN(xRstCab1("impigv")) <> 0 Then
                            xIgvTasa = (NulosN(xRstCab1("impigv")) / NulosN(xRstCab1("impafec"))) * 100
                        End If
                        xRstCab("tasaigv") = xIgvTasa:
                        xRstCab.Update
                        '--------------------------------
                        'Grabar el detalle del documento

                        If xRstDet1.RecordCount > 0 Then
                            Do While Not xRstDet1.EOF
                            
                                xRstDet.AddNew
                                xRstDet("idcom") = Str(xId)
                                xRstDet("iditem") = NulosN(xRstDet1("iditem"))
                                xRstDet("idunimed") = 1
                                xRstDet("canpro") = 1
                                xRstDet("preuni") = NulosN(xRstDet1("impafec")) + NulosC(xRstDet1("impexon"))
                                xRstDet("imptot") = NulosN(xRstDet1("impafec")) + NulosC(xRstDet1("impexon"))
                                xRstDet("preunibru") = NulosN(xRstDet1("impafec"))
                                xRstDet("preunibruina") = NulosC(xRstDet1("impexon"))
                                xRstDet("idcue") = NulosN(xRstDet1("idcue"))
                                xRstDet.Update

                                xRstDet1.MoveNext

                            Loop

                        Else

                        End If
                        
                        '--Generamos el asiento
                        xNumAsiento = GenerarAsiento(xCon, 1, CDbl(xId), AnoTra, mMesTrans, 1, 0)
                        If xNumAsiento = "" Then GoTo LaCague
                        RstDocTr("numreg") = xNumAsiento
                                                        
                        '--Actualizando datos al documento para indicar que fue transferido
                        '--Modulo = 1 = Compras
                        nSQL = "UPDATE tra_documento SET tra_documento.estado = 2, tra_documento.idmod = 1, tra_documento.iddoc = " & xId & ", tra_documento.numreg = '" & xNumAsiento & "'  " _
                            + vbCr + " WHERE (((tra_documento.id)=" & RstDocTr("id") & ")); "
                        xCon.Execute nSQL
                    
                        '---------------------------------------------------------------------------
                        'grabamos el movimiento en la tabla var_edicion
                        '--IdOperacion = 6 = Importado
                        '--IdMenu = 218 = Compras
                        GrabarOperacion xIdUsuario, 218, 6, xHorIni, Time, Date, xCon, xId
                        
                        RstDocTr.MoveNext
                    Loop
                    
                    xCon.CommitTrans
                    
                    If RstDocTr.RecordCount = 1 Then
                        MsgBox "Se transfirió un registro de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Else
                        MsgBox "Se transfirieron " & RstDocTr.RecordCount & " registros de compra", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    End If
                    
                End If
                
                If OptHonorario.Value = True Then
                
                    '--Grabando Registro de Compra
                    LblBarra.Caption = "Transfiriendo Honorarios"
                    DoEvents
                    
                    RST_Busq xRstCab, "Select TOP 1 * from com_honorarios", xCon
                    RST_Busq xRstDet, "Select TOP 1 * from com_honorariosdet", xCon
                    
                    Do While Not RstDocTr.EOF
                        Set xRstCab1 = Nothing
                        Set xRstDet1 = Nothing
                        
                        ProgressBar2.Value = ProgressBar2.Value + 1
                                                
                        '--Cargar datos del documento
                        RST_Busq xRstCab1, "Select * from tra_documento where id = " & RstDocTr("id") & "", xCon
                        
                        RST_Busq xRstDet1, "SELECT * FROM tra_documentodet WHERE tra_documentodet.iddet=" & RstDocTr("id") & " ", xCon

                        '--------------------------------
                        'Grabaos cabecera de honorarios
                        xId = HallaCodigoTabla("com_honorarios", xCon, "id")
                        
                        xRstCab.AddNew
                        xRstCab("id") = Str(xId):
                        xRstCab("idlib") = 40:
                        xRstCab("idtipo") = 5:
                        xRstCab("tipdoc") = NulosN(xRstCab1("tipdoc")):
                        xRstCab("idpro") = NulosN(xRstCab1("idprov")):
                        xRstCab("numser") = Format(NulosC(xRstCab1("numser")), "0000")
                        xRstCab("numdoc") = Format(NulosC(xRstCab1("numdoc")), "0000000000")
                        xRstCab("fchreg") = xFechaReg:
                        xRstCab("fchdoc") = NulosC(xRstCab1("fchdoc")):
                        xRstCab("fchven") = NulosC(xRstCab1("fchvenc")):
                        xRstCab("fchpag") = NulosC(xRstCab1("fchdoc")):
                        xRstCab("idconpag") = 1:
                        xRstCab("idmon") = NulosC(xRstCab1("idmon")):
                        xRstCab("impbru") = NulosN(xRstCab1("impafec")) + NulosN(xRstCab1("impexon")):
                        xRstCab("impigv") = NulosN(xRstCab1("impret")):
                        xRstCab("imptot") = NulosN(xRstCab1("imptot")):
                        xRstCab("impsal") = NulosN(xRstCab1("imptot")):
                        xRstCab("importado") = -1:
                        xRstCab("tipcom") = 1:
                        xRstCab("glosa") = NulosC(xRstCab1("glosa")):
                        xRstCab("tc") = NulosN(xRstCab1("tipcam")):
                        xRstCab("fchrecep") = NulosC(xRstCab1("fchrecep")):

                        xRstCab.Update
                        '---------------------------------
                        'Grabar el detalle del documento

                        If xRstDet1.RecordCount > 0 Then
                            Do While Not xRstDet1.EOF
                            
                                xRstDet.AddNew
                                xRstDet("idhon") = Str(xId)
                                xRstDet("iditem") = NulosN(xRstDet1("iditem"))
                                xRstDet("idunimed") = 1
                                xRstDet("canpro") = 1
                                xRstDet("preuni") = NulosN(xRstDet1("impafec")) + NulosN(xRstDet1("impexon"))
                                xRstDet("imptot") = xRstDet("preuni")
                                xRstDet("preunibru") = xRstDet("preuni")
                                xRstDet.Update

                                xRstDet1.MoveNext

                            Loop

                        Else

                        End If
                        
                        '--Generamos el asiento
                        '--IdLib=40=Recibo por Honorarios
                        xNumAsiento = GenerarAsiento(xCon, 40, CDbl(xId), AnoTra, mMesTrans, 1, 0)
                        If xNumAsiento = "" Then GoTo LaCague
                        RstDocTr("numreg") = xNumAsiento
                        Dg2.Refresh
                        '--Actualizando datos al documento para indicar que fue transferido
                        '--Modulo = 9 = Honorarios
                        nSQL = "UPDATE tra_documento SET tra_documento.estado = 2, tra_documento.idmod = 9, tra_documento.iddoc = " & xId & ", tra_documento.numreg = '" & xNumAsiento & "' " _
                            + vbCr + " WHERE (((tra_documento.id)=" & RstDocTr("id") & ")); "
                        xCon.Execute nSQL
                    
                        '---------------------------------------------------------------------------
                        'grabamos el movimiento en la tabla var_edicion
                        '--IdOperacion = 6 = Importado
                        '--IdMenu = 219 = Honorarios
                        GrabarOperacion xIdUsuario, 219, 6, xHorIni, Time, Date, xCon, xId

                        RstDocTr.MoveNext
                    Loop
                    
                    xCon.CommitTrans
                    
                    If RstDocTr.RecordCount = 1 Then
                        MsgBox "Se transfirió un registro de honorario", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    Else
                        MsgBox "Se transfirieron " & RstDocTr.RecordCount & " registros de honorarios", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    End If
                
                
                End If
                
                Set xRstCab1 = Nothing
                Set xRstDet1 = Nothing
    End If

    
    Set xRstCab = Nothing
    Set xRstDet = Nothing


    Grabar = True

    
    Frame5.Visible = False
    
    Grabar = True
    Exit Function

LaCague:
''Resume
    xCon.RollbackTrans
    MsgBox "No se pudo guardar la importación por el siguiente motivo : " & Err.Description, vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    Err.Clear
    Frame5.Visible = False
    Grabar = False
    
End Function

Sub Cancelar()
    ActivarTool
    Bloquea
    TabOne1.TabEnabled(0) = True
    TabOne1.TabVisible(1) = True
    TabOne1.TabVisible(2) = False
    TabOne1.TabVisible(3) = False

    Label8.Caption = "Detalle de Documento"
    TabOne1.CurrTab = 0
    QueHace = 3
End Sub

Private Sub OpcionesPeriodo()


    '------------------------------------------------------------------------------------------
    '--bloqueamos los botones del toolbar
    CierrePeriodo Toolbar1, IdMenuActivo, mMesActivo, fCierrePeriodo, xCon, xIdUsuario
    '------------------------------------------------------------------------------------------
End Sub



'------------------------
Private Sub CmdBusArch_Click()
    Err.Clear
    'CommonDialog1.CancelError = True
    'Especificar las extensiones a usar
    CommonDialog1.DefaultExt = "*.xls"
    CommonDialog1.FileName = ""
    'CommonDialog1.Filter = "Cardfile (*.crd)|*.crd|Textos (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
    CommonDialog1.Filter = "Documentos de Excel (*.xls)|*.xls"
    CommonDialog1.ShowOpen
    If Err Then
        'Cancelada la operación de abrir
        MsgBox Err.Description & vbCr & Err.Source & vbCr & CommonDialog1.FileName, vbInformation, xTitulo
        Err.Clear
        
    Else
        TxtArchivo.Text = CommonDialog1.FileName
        TxtArchivo2.SetFocus
    End If
    Err.Clear
End Sub

Private Sub CmdBusArch2_Click()
    'CommonDialog1.CancelError = True
    'Especificar las extensiones a usar
    CommonDialog1.DefaultExt = "*.xls"
    CommonDialog1.FileName = ""
    'CommonDialog1.Filter = "Cardfile (*.crd)|*.crd|Textos (*.txt)|*.txt|Todos los archivos (*.*)|*.*"
    CommonDialog1.Filter = "Documentos de Excel (*.xls)|*.xls"
    CommonDialog1.ShowOpen
    If Err Then
        'Cancelada la operación de abrir
    Else
        TxtArchivo2.Text = CommonDialog1.FileName
'        CmdCargar.SetFocus

    End If
End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    If ButtonMenu.Parent.Index = 1 Then
        TabOne1.TabEnabled(0) = False
        QueHace = 1
        'Nuevo importacion
        If ButtonMenu.Index = 1 Then

            TabOne1.TabVisible(1) = False
            TabOne1.TabVisible(2) = True
            TabOne1.TabVisible(3) = False
            TabOne1.CurrTab = 2

        End If

        'Nueva Transferencia
        If ButtonMenu.Index = 2 Then
            TabOne1.TabVisible(1) = False
            TabOne1.TabVisible(2) = False
            TabOne1.TabVisible(3) = True
            TabOne1.CurrTab = 3
        End If

        Nuevo

    End If
    
    '--Modificar
    If ButtonMenu.Parent.Index = 2 Then
        '--Activar documento
        If ButtonMenu.Index = 1 Then RestaurarDocumento

    End If
    
    '--Eliminar
    If ButtonMenu.Parent.Index = 3 Then
        '--Eliminar Documento
        If ButtonMenu.Index = 1 Then Eliminar
        '--Retirar Documento
        If ButtonMenu.Index = 2 Then RetirarDocumento

    End If

End Sub

Private Sub TxtArchivo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then CmdBusArch_Click
End Sub

Private Sub TxtArchivo2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF5 Then CmdBusArch2_Click
End Sub


Private Sub CmdCargar_Click()
    If QueHace = 1 Then
        If TxtArchivo.Text = "" Then
            MsgBox "No ha especificado el nombre del archivo cabecera ", vbInformation, xTitulo
            TxtArchivo.SetFocus
            Exit Sub
        End If

        If TxtArchivo2.Text = "" Then
            MsgBox "No ha especificado el nombre del archivo detalle ", vbInformation, xTitulo
            TxtArchivo2.SetFocus
            Exit Sub
        End If

        CargaDocumentos2  'estamos cargando datos de un archivo de excel
    End If

'    If QueHace = 3 Then
'        If TxtMes.Text = "" Then
'            MsgBox "No ha especificado el mes a consultar", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'            CmdBusMes.SetFocus
'            Exit Sub
'        End If
'        MostrarDocumentos  'mostramos los registros ya guardados
'    End If

End Sub

Sub CargaDocumentos1()
    Dim xNumFilas As Integer
    Dim A&
    Dim B As Integer
    Dim xFilas As Integer

    '-------

    Dim nSQL As String

    On Error GoTo error:

    '---------------------------------------------------------------------

    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    'Dim objExcel As New Excel.Application

    objExcel.Visible = True
    objExcel.SheetsInNewWorkbook = 1

    'abre el Libro
    objExcel.WindowState = 2
    objExcel.Workbooks.Open Trim(TxtArchivo.Text)

    Frame5.Left = 3090
    Frame5.Top = 2910
    Frame5.Visible = True
    '--definiendo la estructura del rst detalle para almacenar el detalle de los documentos
    PreparaRstTmp

    xFilas = 3

    xNumFilas = 1

    Fg1.Rows = 1
    Fg2.Rows = 1

    Fg1.Rows = 1

    With objExcel.ActiveSheet
        'DETERMINAMOS EL NUMERO DE FILAS CON DATOS
        LblBarra.Caption = "Calculando número de registros"
        DoEvents
        ProgressBar2.Max = 32000
        For A = 2 To 32000
            ProgressBar2.Value = A
            '--verificar campo proveeodor
            If NulosC(.Cells(A, 5)) <> "" Then
                xNumFilas = xNumFilas + 1
            Else
                Exit For
            End If
        Next A

        xNumFilas = xNumFilas + 1
        LblBarra.Caption = "Cargando registros - Cabecera de Documentos"
        DoEvents
        ProgressBar2.Max = xNumFilas

        For A = 2 To xNumFilas
            ProgressBar2.Value = A
            
            DoEvents
            
            If NulosC(.Cells(A, 1)) = "" Then Exit For
            Fg1.Rows = Fg1.Rows + 1

            Fg1.TextMatrix(A - 1, 1) = NulosC(.Cells(A, 1)) '--Año registro segun cliente
            Fg1.TextMatrix(A - 1, 2) = NulosC(.Cells(A, 2)) '--Periodo registro segun cliente
            Fg1.TextMatrix(A - 1, 3) = Format(NulosC(.Cells(A, 3)), "0000") '--Correlativo
            Fg1.TextMatrix(A - 1, 4) = NulosC(.Cells(A, 4)) '--Ruc proveedor
            Fg1.TextMatrix(A - 1, 5) = NulosC(.Cells(A, 5)) '--Razon social del proveedor
            Fg1.TextMatrix(A - 1, 6) = NulosC(.Cells(A, 6)) '--Tipo de Documento
            Fg1.TextMatrix(A - 1, 7) = NulosC(.Cells(A, 7)) '--N° Serie
            Fg1.TextMatrix(A - 1, 8) = NulosC(.Cells(A, 8)) '--N° Documento

            If IsDate(CDate(.Cells(A, 9))) = True Then Fg1.TextMatrix(A - 1, 9) = Format(CDate(.Cells(A, 9)), FORMAT_DATE) '--Fecha Documento
            If IsDate(CDate(.Cells(A, 10))) = True Then Fg1.TextMatrix(A - 1, 10) = Format(CDate(.Cells(A, 10)), FORMAT_DATE)  '--Fecha Recepción
            If IsDate(CDate(.Cells(A, 11))) = True Then Fg1.TextMatrix(A - 1, 11) = Format(CDate(.Cells(A, 11)), FORMAT_DATE)  '--Fecha Vencimiento
            If IsDate(CDate(.Cells(A, 12))) = True Then Fg1.TextMatrix(A - 1, 12) = Format(CDate(.Cells(A, 12)), FORMAT_DATE)  '--Fecha Sistema

            Fg1.TextMatrix(A - 1, 13) = NulosC(.Cells(A, 13)) '--Moneda
            Fg1.TextMatrix(A - 1, 14) = NulosN(.Cells(A, 14)) '--Tipo de Cambio
            
            '--Si documento es nota de credito, mostrar los importes en positivo
            Fg1.TextMatrix(A - 1, 15) = Abs(NulosN(.Cells(A, 15))) '--Imp. Afecto
            Fg1.TextMatrix(A - 1, 16) = Abs(NulosN(.Cells(A, 16))) '--Imp Inafecto
            Fg1.TextMatrix(A - 1, 17) = Abs(NulosN(.Cells(A, 17))) '--Imp. Retencion(para honorarios)
            Fg1.TextMatrix(A - 1, 18) = Abs(NulosN(.Cells(A, 18))) '--Imp Igv
            Fg1.TextMatrix(A - 1, 19) = Abs(NulosN(.Cells(A, 19))) '--Imp. Total

            Fg1.TextMatrix(A - 1, 20) = NulosC(.Cells(A, 20)) '--Glosa
            Fg1.TextMatrix(A - 1, 21) = NulosC(.Cells(A, 21)) '--Ruc cliente
            Fg1.TextMatrix(A - 1, 22) = NulosC(.Cells(A, 22)) '--Razón Social del Cliente
            Fg1.TextMatrix(A - 1, 23) = NulosC(.Cells(A, 23)) '--Orden de Despacho


        Next A
    End With

    DoEvents

    '*******************************************************************************************************************
    'CARGAMOS EL DETALLE DE LA VENTA
    objExcel.WindowState = 2
    objExcel.Workbooks.Open Trim(TxtArchivo2.Text)

    Fg2.Rows = 1
    With objExcel.ActiveSheet
        'DETERMINAMOS EL NUMERO DE FILAS CON DATOS
        LblBarra.Caption = "Calculando número de registros"
        ProgressBar2.Max = 32000
        ProgressBar2.Value = 1
        For A = 2 To 32000
            ProgressBar2.Value = A
            '--verificar campo proveeodor
            If NulosC(.Cells(A, 2)) <> "" Then
                xNumFilas = xNumFilas + 1
            Else
                Exit For
            End If
        Next A

        xNumFilas = xNumFilas + 1
        ProgressBar2.Max = xNumFilas
        LblBarra.Caption = "Cargando registros - Detalle de Documentos"
        DoEvents
        
        For A = 2 To xNumFilas
            ProgressBar2.Value = A
            DoEvents

            If NulosC(.Cells(A, 1)) = "" Then Exit For

            Fg2.Rows = Fg2.Rows + 1

            Fg2.TextMatrix(A - 1, 1) = NulosC(.Cells(A, 1)) '--Ruc proveedor
            Fg2.TextMatrix(A - 1, 2) = NulosC(.Cells(A, 2)) '--Razón social del proveedor
            Fg2.TextMatrix(A - 1, 3) = NulosC(.Cells(A, 3)) '--Tipo de Documento
            Fg2.TextMatrix(A - 1, 4) = NulosC(.Cells(A, 4)) '--N° Serie
            Fg2.TextMatrix(A - 1, 5) = NulosC(.Cells(A, 5)) '--N° Documento
            '--Si documento es nota de credito, mostrar los importes en positivo
            Fg2.TextMatrix(A - 1, 6) = Abs(NulosN(.Cells(A, 6))) '--Imp. Afecto
            Fg2.TextMatrix(A - 1, 7) = Abs(NulosN(.Cells(A, 7))) '--Imp Inafecto
            Fg2.TextMatrix(A - 1, 8) = Abs(NulosN(.Cells(A, 8))) '--Imp. Retencion(para honorarios)
            Fg2.TextMatrix(A - 1, 9) = Abs(NulosN(.Cells(A, 9))) '--Imp Igv
            
            Fg2.TextMatrix(A - 1, 11) = NulosC(.Cells(A, 10)) '--vinc1=Empresa
            Fg2.TextMatrix(A - 1, 12) = NulosC(.Cells(A, 11)) '--vinc2=Centro Costo
            Fg2.TextMatrix(A - 1, 13) = NulosC(.Cells(A, 12)) '--vinc3=Cuenta
            Fg2.TextMatrix(A - 1, 14) = NulosC(.Cells(A, 13)) '--vinc4=SubCuenta

            '--agregando registros al detalle
            RstTmpDet.AddNew
            RstTmpDet("numruc") = Trim(.Cells(A, 1))
            RstTmpDet("documento") = NulosC(.Cells(A, 3))
            RstTmpDet("numser") = NulosC(.Cells(A, 4))
            RstTmpDet("numdoc") = NulosC(.Cells(A, 5))
            RstTmpDet("impafec") = Abs(NulosN(.Cells(A, 6)))
            RstTmpDet("impexon") = Abs(NulosN(.Cells(A, 7)))
            RstTmpDet("impret") = Abs(NulosN(.Cells(A, 8)))
            RstTmpDet("impigv") = Abs(NulosN(.Cells(A, 9)))
            
            If InStr(RstTmpDet("documento"), "hon") = 0 Then
                RstTmpDet("imptot") = RstTmpDet("impafec") + RstTmpDet("impexon") + RstTmpDet("impigv")
            Else
                '--Si documento es recibo por honorario y se aplica retencion, este se se aplica al total del documento
                RstTmpDet("imptot") = RstTmpDet("impafec") - RstTmpDet("impret")
            End If
            RstTmpDet("vinc1") = NulosC(.Cells(A, 10))
            RstTmpDet("vinc2") = NulosC(.Cells(A, 11))
            RstTmpDet("vinc3") = NulosC(.Cells(A, 12))
            RstTmpDet("vinc4") = NulosC(.Cells(A, 13))

            RstTmpDet.Update
        Next A
    End With
    '---------------------------------------------

    Frame5.Visible = False
    MsgBox "El proceso terminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, "Registro de Compras y Ventas"
    objExcel.WindowState = 2
    objExcel.Workbooks.Close

    Set objExcel = Nothing
    Exit Sub
error:
'Resume
    Frame5.Visible = False
    objExcel.Workbooks.Close
    If Err.Number = 424 Then
        MsgBox Err.Description & vbCr & "El archivo fue cerrado antes de terminar de importar, vuelva a importar nuevamente.", vbCritical, xTitulo
    Else
        MsgBox Err.Description & vbCr & Err.Source, vbCritical, xTitulo
    End If
    Fg1.Rows = 1
    Fg2.Rows = 1
    Set objExcel = Nothing

End Sub

Sub PreparaRstTmp()
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(13, 3) As String

    xCampos(0, 0) = "numruc":       xCampos(0, 1) = "C":      xCampos(0, 2) = "20"
    xCampos(1, 0) = "documento":    xCampos(1, 1) = "C":      xCampos(1, 2) = "50"
    xCampos(2, 0) = "numser":       xCampos(2, 1) = "C":      xCampos(2, 2) = "10"
    xCampos(3, 0) = "numdoc":       xCampos(3, 1) = "C":      xCampos(3, 2) = "50"
    xCampos(4, 0) = "impafec":      xCampos(4, 1) = "D":      xCampos(4, 2) = "8"
    xCampos(5, 0) = "impexon":      xCampos(5, 1) = "D":      xCampos(5, 2) = "8"
    xCampos(6, 0) = "impigv":       xCampos(6, 1) = "D":      xCampos(6, 2) = "8"
    xCampos(7, 0) = "impret":       xCampos(7, 1) = "D":      xCampos(7, 2) = "8"
    xCampos(8, 0) = "imptot":       xCampos(8, 1) = "D":      xCampos(8, 2) = "8"
    xCampos(9, 0) = "vinc1":        xCampos(9, 1) = "C":      xCampos(9, 2) = "200"
    xCampos(10, 0) = "vinc2":       xCampos(10, 1) = "C":     xCampos(10, 2) = "200"
    xCampos(11, 0) = "vinc3":       xCampos(11, 1) = "C":     xCampos(11, 2) = "200"
    xCampos(12, 0) = "vinc4":       xCampos(12, 1) = "C":     xCampos(12, 2) = "200"

    Set RstTmpDet = xFun.CrearRstTMP(xCampos)

    RstTmpDet.Open

End Sub



'------------------------
Private Sub DepurarDatos()
    Dim nSQL As String
    '--------------------------------------------------------
    MousePointer = vbHourglass
    DoEvents
    
    '--------------------------------------------------------
    '--Restableciendo a valores inciales solo a documentos que no fueron transferidos  estado(pendiente, observado)
    nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet " _
        + vbCr + " SET tra_documento.idprov = 0, tra_documento.tipdoc = 0, tra_documento.idmon = 0, tra_documentodet.iditem = 0 " _
        + vbCr + " WHERE (((tra_documento.estado) In (0,1)));"
    xCon.Execute nSQL
    
    '--------------------------------------------------------
    '--Depurando proveedores - x Nombre(tabla equivalencia ruc)
    nSQL = "UPDATE tra_documento INNER JOIN tra_ruc ON tra_documento.proveedor = tra_ruc.descripcion SET tra_documento.idprov = [tra_ruc].[idref] " _
        + vbCr + " WHERE (((tra_ruc.tipo)=1) AND ((tra_documento.idprov)=0));"
    xCon.Execute nSQL

    '--Depurando proveedores - x Ruc(tabla mae_prov)
    nSQL = "UPDATE tra_documento INNER JOIN mae_prov ON tra_documento.rucprov = mae_prov.numruc SET tra_documento.idprov = [mae_prov].[id] " _
        + vbCr + " WHERE (((tra_documento.idprov)=0)); "
    xCon.Execute nSQL

    '--Depurando proveedores - x Nombre(tabla mae_prov)
    nSQL = "UPDATE tra_documento INNER JOIN mae_prov ON tra_documento.proveedor = mae_prov.nombre SET tra_documento.idprov = [mae_prov].[id] " _
        + vbCr + " WHERE (((tra_documento.idprov)=0)); "
    xCon.Execute nSQL

    '--Depurando Tipo de Documento(tabla equivalencia tipo doc)
    nSQL = "UPDATE tra_documento INNER JOIN tra_tipodoc ON tra_documento.documento = tra_tipodoc.descripcion2 SET tra_documento.tipdoc = [tra_tipodoc].[tipdoc] " _
        + vbCr + " WHERE (((tra_documento.tipdoc)=0)); "
    xCon.Execute nSQL

    '--Depurando la Moneda(Tabla mae_moneda.simbolo)
    nSQL = "UPDATE tra_documento INNER JOIN mae_moneda ON tra_documento.moneda = mae_moneda.simbolo SET tra_documento.idmon = [mae_moneda].[id] " _
        + vbCr + " WHERE (((tra_documento.idmon)=0)); "
    xCon.Execute nSQL

    '--Depurando la Moneda(Tabla mae_moneda.descripcion)
    nSQL = "UPDATE tra_documento INNER JOIN mae_moneda ON tra_documento.moneda = mae_moneda.descripcion SET tra_documento.idmon = [mae_moneda].[id] " _
        + vbCr + " WHERE (((tra_documento.idmon)=0)); "
    xCon.Execute nSQL

    '--Depurando Item del detalle de Documento(tabla equivalencia items)
    nSQL = "UPDATE (tra_documentodet INNER JOIN tra_item ON (tra_documentodet.vinc4 = tra_item.vinc4) AND (tra_documentodet.vinc3 = tra_item.vinc3) AND (tra_documentodet.vinc2 = tra_item.vinc2) AND (tra_documentodet.vinc1 = tra_item.vinc1)) INNER JOIN tra_documento ON tra_documentodet.iddet = tra_documento.id SET tra_documentodet.iditem = iif([tra_item].[iditem] is null,0,[tra_item].[iditem]) " _
        + vbCr + " WHERE (((tra_documento.estado)<>2)); "
        
    xCon.Execute nSQL
    '--------------------------------------------------------
    '--Actualizando campo estado a 0=listos para transferir, solo los observados
    nSQL = "UPDATE tra_documento SET tra_documento.estado = 0 " _
        + vbCr + " WHERE (((tra_documento.estado)=1)); "
    xCon.Execute nSQL

    '--Actualizando campo estado a 1=Documentos x Depurar
    nSQL = "UPDATE tra_documento LEFT JOIN (SELECT tra_documentodet.iddet " _
        + vbCr + " FROM tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet " _
        + vbCr + " WHERE (((tra_documento.estado) In (0,1)) AND ((tra_documentodet.iditem)=0)) " _
        + vbCr + " ) AS det ON tra_documento.id = det.iddet SET tra_documento.estado = 1 " _
        + vbCr + " WHERE (((tra_documento.estado)=0) AND ((tra_documento.idprov)=0)) OR (((tra_documento.estado)=0) AND ((tra_documento.tipdoc)=0)) OR (((tra_documento.estado)=0) AND ((tra_documento.idmon)=0)) OR (((det.iddet) Is Not Null)); "
    xCon.Execute nSQL


    MousePointer = vbDefault
    DoEvents

End Sub

Private Sub VerDatosxDepurar()
    Dim nSQL As String
    Dim xRs As New ADODB.Recordset
    '--------------------------------------------------------
    Fg3.Rows = 1
    Fg4.Rows = 1
    Fg5.Rows = 1
    DoEvents
    '--------------------------------------------------------
    '--Cargando proveedores Observados
    nSQL = "SELECT tra_documento.rucprov, tra_documento.proveedor, Count(tra_documento.id) AS candoc " _
    + vbCr + " FROM tra_documento " _
    + vbCr + " WHERE (((tra_documento.idprov) = 0)) and tra_documento.estado=1 " _
    + vbCr + " GROUP BY tra_documento.rucprov, tra_documento.proveedor " _
    + vbCr + " ORDER BY tra_documento.proveedor ;"
    RST_Busq xRs, nSQL, xCon
    If xRs.RecordCount <> 0 Then
        Do While Not xRs.EOF
            Fg3.Rows = Fg3.Rows + 1
            Fg3.TextMatrix(Fg3.Rows - 1, 1) = Fg3.Rows - 1
            Fg3.TextMatrix(Fg3.Rows - 1, 2) = NulosC(xRs("rucprov"))
            Fg3.TextMatrix(Fg3.Rows - 1, 3) = NulosC(xRs("proveedor"))
            Fg3.TextMatrix(Fg3.Rows - 1, 4) = NulosN(xRs("candoc"))
            xRs.MoveNext
        Loop
    End If
    Set xRs = Nothing

    '--Cargando tipos de documentos observados
    nSQL = "SELECT tra_documento.documento " _
        + vbCr + " From tra_documento " _
        + vbCr + " Where (((tra_documento.tipdoc) = 0)) " _
        + vbCr + " GROUP BY tra_documento.documento " _
        + vbCr + " ORDER BY tra_documento.documento; "
    
    RST_Busq xRs, nSQL, xCon
    
    If xRs.RecordCount <> 0 Then
        Do While Not xRs.EOF
            Fg4.Rows = Fg4.Rows + 1
            Fg4.TextMatrix(Fg4.Rows - 1, 1) = Fg4.Rows - 1
            Fg4.TextMatrix(Fg4.Rows - 1, 2) = NulosC(xRs("documento"))
            xRs.MoveNext
        Loop
    End If
    Set xRs = Nothing

    '--Cargando Items observados
    nSQL = "SELECT tra_documentodet.vinc1, tra_documentodet.vinc2, tra_documentodet.vinc3, tra_documentodet.vinc4, Count(tra_documento.id) AS candoc " _
        + vbCr + " FROM tra_documentodet INNER JOIN tra_documento ON tra_documentodet.iddet = tra_documento.id " _
        + vbCr + " WHERE (((tra_documentodet.iditem)=0 Or (tra_documentodet.iditem) Is Null) AND ((tra_documento.estado)=1)) " _
        + vbCr + " GROUP BY tra_documentodet.vinc1, tra_documentodet.vinc2, tra_documentodet.vinc3, tra_documentodet.vinc4 " _
        + vbCr + " ORDER BY tra_documentodet.vinc1, tra_documentodet.vinc2, tra_documentodet.vinc3, tra_documentodet.vinc4;"
    
    RST_Busq xRs, nSQL, xCon
    
    If xRs.RecordCount <> 0 Then
        Do While Not xRs.EOF
            Fg5.Rows = Fg5.Rows + 1
            Fg5.TextMatrix(Fg5.Rows - 1, 1) = Fg5.Rows - 1
            Fg5.TextMatrix(Fg5.Rows - 1, 2) = NulosC(xRs("vinc1"))
            Fg5.TextMatrix(Fg5.Rows - 1, 3) = NulosC(xRs("vinc2"))
            Fg5.TextMatrix(Fg5.Rows - 1, 4) = NulosC(xRs("vinc3"))
            Fg5.TextMatrix(Fg5.Rows - 1, 5) = NulosC(xRs("vinc4"))
            Fg5.TextMatrix(Fg5.Rows - 1, 6) = NulosN(xRs("candoc"))
            xRs.MoveNext
        Loop
    End If
    Set xRs = Nothing

    '--verificar si hay datos observados para mostrar ventana
    If Fg3.Rows <> Fg3.FixedRows Or Fg4.Rows <> Fg4.FixedRows Or Fg5.Rows <> Fg4.FixedRows Then
        FraDepura.Visible = True
        FraDepura.Left = 540
        FraDepura.Top = 1170
        '--Posicionando el TabOne en la pestaña de observados
        If Fg3.Rows <> Fg3.FixedRows Then
            TabOne3.CurrTab = 0
        ElseIf Fg4.Rows <> Fg4.FixedRows Then
            TabOne3.CurrTab = 1
        Else
            TabOne3.CurrTab = 2
        End If
    End If
    
    If Fg3.Rows = Fg3.FixedRows And Fg4.Rows = Fg4.FixedRows And Fg5.Rows = Fg4.FixedRows Then
    
        MsgBox "No hay datos por depurar", vbInformation, xTitulo
        
    End If
    
    
'--Documentos Observados cuyo suma de detalle no coincide con el total
'SELECT tra_documento.id, tra_documento.rucprov, tra_documento.proveedor, mae_documento.abrev, [tra_documento].[numser] & '-' & [tra_documento].[numdoc] AS numerodoc, tra_documento.fchdoc, mae_moneda.simbolo, tra_documento.imptot, Sum(tra_documentodet.imptot) AS totdet, tra_documento.imptot-Sum(tra_documentodet.imptot) AS dif
'FROM ((tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet) LEFT JOIN mae_documento ON tra_documento.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON tra_documento.idmon = mae_moneda.id
'GROUP BY tra_documento.id, tra_documento.rucprov, tra_documento.proveedor, mae_documento.abrev, [tra_documento].[numser] & '-' & [tra_documento].[numdoc], tra_documento.fchdoc, mae_moneda.simbolo, tra_documento.imptot, tra_documento.tipdoc
'HAVING (((tra_documento.imptot-Sum(tra_documentodet.imptot)) Not Between 1 And -1) AND ((tra_documento.tipdoc)<>2));
    
'--revisar si estan en regisro de savar
'--actualizar estado a pendiente u observado
    
    
    DoEvents
End Sub



Private Sub FraDepura_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    FraDepura.ZOrder 0
End Sub

Private Sub FraDepura_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With FraDepura
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub

Private Sub FraDetalle_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    OrigFX = X
    OrigFY = Y
    FraDetalle.ZOrder 0
End Sub

Private Sub FraDetalle_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 0 Then
        With FraDetalle
            .Move .Left + X - OrigFX, .Top + Y - OrigFY
        End With
    End If
End Sub
'--------------------------
Private Sub Dg2_DblClick()
    On Error Resume Next
    If RstDocTr.State = 0 Then Exit Sub
    If RstDocTr.RecordCount < 1 Then Exit Sub
    RstDocTr.Fields("xsel") = Not RstDocTr.Fields("xsel")
    If NulosN(RstDocTr.Fields("xsel")) = -1 Then
        RstDocTr.Fields("xorden") = Format(xOrden, "000")
        xOrden = xOrden + 1
    Else
        RstDocTr.Fields("xorden") = ""
    End If
    
    Err.Clear
End Sub

Private Sub Dg2_FilterChange()
    TDB_FiltroGenerar Dg2, RstDocTr
End Sub

Private Sub Dg2_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Dg2_DblClick
End Sub

Private Sub dg2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu Menu1
End Sub

Private Sub Menu1_1_Click()
    If RstDocTr.State = 0 Then Exit Sub
    TDB_SelDesActCheck Dg2, RstDocTr, "xsel", "-1"
End Sub

Private Sub Menu1_2_Click()
    If RstDocTr.State = 0 Then Exit Sub
    TDB_SelDesActCheck Dg2, RstDocTr, "xsel", "0"
End Sub

Private Sub Menu1_4_Click()
    If RstDocTr.State = 0 Then Exit Sub
    TDB_TodosDesActCheck Dg2, RstDocTr, "xsel", "-1"
End Sub

Private Sub Menu1_5_Click()
    If RstDocTr.State = 0 Then Exit Sub
    TDB_TodosDesActCheck Dg2, RstDocTr, "xsel", "0"
End Sub

Private Sub Menu1_7_Click()
    '--limpiar los filtros
    
    TDB_FiltroLimpiar Dg2
    
    If RstDocTr.State = 0 Then Exit Sub
    
    RstDocTr.Filter = "xsel=-1"
    If RstDocTr.RecordCount <> 0 Then
        RstDocTr.MoveFirst
        Do While Not RstDocTr.EOF
            RstDocTr("xsel") = "0"
            RstDocTr.MoveNext
        Loop
    End If
    RstDocTr.Filter = ""

    If RstDocTr.RecordCount <> 0 Then RstDocTr.MoveFirst


End Sub
'--------------------------

Private Sub RetirarDocumento()
'--30/10/11
    Dim xRst As New Recordset
    Dim nSQL As String

    ReDim xCampos(8, 4) As String
    
    xCampos(0, 0) = "Ruc":          xCampos(0, 1) = "rucprov":     xCampos(0, 2) = "1100":    xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
    xCampos(1, 0) = "Proveedor":    xCampos(1, 1) = "proveedor":   xCampos(1, 2) = "2500":    xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
    xCampos(2, 0) = "T.D.":         xCampos(2, 1) = "abrev":       xCampos(2, 2) = "500":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
    xCampos(3, 0) = "Nª Documento": xCampos(3, 1) = "numerodoc":   xCampos(3, 2) = "1200":    xCampos(3, 3) = "C":    xCampos(3, 4) = "C"
    xCampos(4, 0) = "Fch. Doc":     xCampos(4, 1) = "fchdoc1":     xCampos(4, 2) = "1000":    xCampos(4, 3) = "C":    xCampos(4, 4) = "C"
    xCampos(5, 0) = "M":            xCampos(5, 1) = "simbolo":     xCampos(5, 2) = "500":     xCampos(5, 3) = "C":    xCampos(5, 4) = "C"
    xCampos(6, 0) = "Importe":      xCampos(6, 1) = "imptot1":     xCampos(6, 2) = "800":    xCampos(6, 3) = "N":    xCampos(6, 4) = "C"
    xCampos(7, 0) = "Glosa":        xCampos(7, 1) = "glosa":       xCampos(7, 2) = "2500":    xCampos(7, 3) = "C":    xCampos(7, 4) = "C"
    
    '--tra_documento.estado 2=tTransferido, 3=Retirado
    nSQL = "SELECT 0 as xsel, tra_documento.id, tra_documento.rucprov, tra_documento.proveedor, mae_documento.abrev, [tra_documento].[numser] & '-' & [tra_documento].[numdoc] AS numerodoc, tra_documento.fchdoc & '' as fchdoc1, mae_moneda.simbolo, tra_documento.imptot & '' as imptot1, tra_documento.glosa " _
        + vbCr + " FROM (tra_documento LEFT JOIN mae_documento ON tra_documento.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON tra_documento.idmon = mae_moneda.id " _
        + vbCr + " WHERE (((tra_documento.estado) Not In (2,3))); "
    
    CARGAR_DLL_EPSBUSCAR_SEL xCon, xRst, nSQL, xCampos(), "Retirar Documentos"

    If xRst.State = 1 Then
        If xRst.RecordCount <> 0 Then
            While Not xRst.EOF
                xCon.Execute "UPDATE tra_documento " _
                            & " SET tra_documento.estado = 3 " _
                            & " WHERE (((tra_documento.id)=" & xRst("id") & ")); "
                
                xRst.MoveNext
            Wend
        End If
        If xRst.RecordCount = 1 Then
            MsgBox "Se retiró un documento", vbInformation, xTitulo
        Else
            MsgBox "Se retiraron " & xRst.RecordCount & " documentos", vbInformation, xTitulo
        End If
    End If
    
    Set xRst = Nothing
    
    '--Refrescando el listado
    If RstDoc.State = 0 Then Exit Sub
    RstDoc.Requery
    Dg1.Refresh
    
End Sub


Private Sub RestaurarDocumento()
'--30/10/11
    Dim xRst As New Recordset
    Dim nSQL As String

    ReDim xCampos(8, 4) As String
    
    xCampos(0, 0) = "Ruc":          xCampos(0, 1) = "rucprov":     xCampos(0, 2) = "1100":    xCampos(0, 3) = "C":    xCampos(0, 4) = "C"
    xCampos(1, 0) = "Proveedor":    xCampos(1, 1) = "proveedor":   xCampos(1, 2) = "2500":    xCampos(1, 3) = "C":    xCampos(1, 4) = "C"
    xCampos(2, 0) = "T.D.":         xCampos(2, 1) = "abrev":       xCampos(2, 2) = "500":     xCampos(2, 3) = "C":    xCampos(2, 4) = "C"
    xCampos(3, 0) = "Nª Documento": xCampos(3, 1) = "numerodoc":   xCampos(3, 2) = "1200":    xCampos(3, 3) = "C":    xCampos(3, 4) = "C"
    xCampos(4, 0) = "Fch. Doc":     xCampos(4, 1) = "fchdoc1":     xCampos(4, 2) = "1000":    xCampos(4, 3) = "C":    xCampos(4, 4) = "C"
    xCampos(5, 0) = "M":            xCampos(5, 1) = "simbolo":     xCampos(5, 2) = "500":     xCampos(5, 3) = "C":    xCampos(5, 4) = "C"
    xCampos(6, 0) = "Importe":      xCampos(6, 1) = "imptot1":     xCampos(6, 2) = "800":     xCampos(6, 3) = "N":    xCampos(6, 4) = "C"
    xCampos(7, 0) = "Glosa":        xCampos(7, 1) = "glosa":       xCampos(7, 2) = "2500":    xCampos(7, 3) = "C":    xCampos(7, 4) = "C"
    
    '--tra_documento.estado 2=tTransferido, 3=Retirado
    nSQL = "SELECT 0 as xsel, tra_documento.id, tra_documento.rucprov, tra_documento.proveedor, mae_documento.abrev, [tra_documento].[numser] & '-' & [tra_documento].[numdoc] AS numerodoc, tra_documento.fchdoc & '' as fchdoc1, mae_moneda.simbolo, tra_documento.imptot & '' as imptot1, tra_documento.glosa " _
        + vbCr + " FROM (tra_documento LEFT JOIN mae_documento ON tra_documento.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON tra_documento.idmon = mae_moneda.id " _
        + vbCr + " WHERE (((tra_documento.estado) In (3))); "
    
    CARGAR_DLL_EPSBUSCAR_SEL xCon, xRst, nSQL, xCampos(), "Restaurar Documentos"

    If xRst.State = 1 Then
        If xRst.RecordCount <> 0 Then
            While Not xRst.EOF
                '--Actualizar estado a observado
                xCon.Execute "UPDATE tra_documento " _
                            & " SET tra_documento.estado = 1 " _
                            & " WHERE (((tra_documento.id)=" & xRst("id") & ")); "
                
                xRst.MoveNext
            Wend
            
            If xRst.RecordCount = 1 Then
                MsgBox "Se restauró un documento", vbInformation, xTitulo
            Else
                MsgBox "Se restauraron " & xRst.RecordCount & " documentos", vbInformation, xTitulo
            End If
        End If
    End If
    '--Proceder a depurar los datos afin de identificar si los documentos restaurados estan listos para la transferencia
    DepurarDatos
   
    Set xRst = Nothing
    
    '--Refrescando el listado
    If RstDoc.State = 0 Then Exit Sub
    RstDoc.Requery
    Dg1.Refresh
    
    
End Sub


Sub CargaDocumentos2()
    '--111104

    Dim xNumFilas As Integer
    Dim A&
    Dim B As Integer
    Dim xFilas As Long
    Dim xFilaIni As Long
    '-------

    Dim nSQL As String

    On Error GoTo error:

    '---------------------------------------------------------------------

    Dim objExcel As Object
    Set objExcel = CreateObject("Excel.Application")
    'Dim objExcel As New Excel.Application

    objExcel.Visible = True
    objExcel.SheetsInNewWorkbook = 1

    'abre el Libro
    objExcel.WindowState = 2
    objExcel.Workbooks.Open Trim(TxtArchivo.Text)

    Frame5.Left = 3090
    Frame5.Top = 2910
    Frame5.Visible = True
    '--definiendo la estructura del rst detalle para almacenar el detalle de los documentos
    PreparaRstTmp

    '--indica el inicio de lectura de registros
    xFilaIni = 9

    xNumFilas = 1

    Fg1.Rows = 1
    Fg2.Rows = 1

    Fg1.Rows = 1

    With objExcel.ActiveSheet
        'DETERMINAMOS EL NUMERO DE FILAS CON DATOS
        LblBarra.Caption = "Calculando número de registros"
        DoEvents
        ProgressBar2.Max = 32000
        For A = xFilaIni To 32000
            ProgressBar2.Value = A
            '--verificar idsolcheque, proveedor
            If NulosC(.Cells(A, 1)) <> "" Or NulosC(.Cells(A, 8)) <> "" Then
                xNumFilas = xNumFilas + 1
            Else
                Exit For
            End If
        Next A

        xNumFilas = xNumFilas + xFilaIni
        LblBarra.Caption = "Cargando registros - Cabecera de Documentos"
        DoEvents
        ProgressBar2.Max = xNumFilas

        For A = xFilaIni To xNumFilas
            ProgressBar2.Value = A
            
            DoEvents
            '--verificar si proveedor es nulo para cancelar
            If NulosC(.Cells(A, 1)) = "" And NulosC(.Cells(A, 8)) = "" Then Exit For
            
            Fg1.Rows = Fg1.Rows + 1
            
            xFilas = Fg1.Rows - 1
            
            Fg1.TextMatrix(xFilas, 1) = NulosC(.Cells(A, 21)) '--Año registro segun cliente

            Fg1.TextMatrix(xFilas, 2) = NulosC(.Cells(A, 20)) '--Periodo registro segun cliente
            Fg1.TextMatrix(xFilas, 3) = Format(NulosC(.Cells(A, 2)), "0000") '--Correlativo
            Fg1.TextMatrix(xFilas, 4) = NulosC(.Cells(A, 8)) '--Ruc proveedor
            Fg1.TextMatrix(xFilas, 5) = NulosC(.Cells(A, 9)) '--Razon social del proveedor
            Fg1.TextMatrix(xFilas, 6) = NulosC(.Cells(A, 4)) '--Tipo de Documento
            Fg1.TextMatrix(xFilas, 7) = NulosC(.Cells(A, 5)) '--N° Serie
            Fg1.TextMatrix(xFilas, 8) = NulosC(.Cells(A, 6)) '--N° Documento

            If IsDate(CDate(.Cells(A, 22))) = True Then Fg1.TextMatrix(xFilas, 9) = Format(CDate(.Cells(A, 22)), FORMAT_DATE) '--Fecha Documento
            If IsDate(CDate(.Cells(A, 23))) = True Then Fg1.TextMatrix(xFilas, 10) = Format(CDate(.Cells(A, 23)), FORMAT_DATE)  '--Fecha Recepción
            If IsDate(CDate(.Cells(A, 24))) = True Then Fg1.TextMatrix(xFilas, 11) = Format(CDate(.Cells(A, 24)), FORMAT_DATE)  '--Fecha Vencimiento
            If IsDate(CDate(.Cells(A, 25))) = True Then Fg1.TextMatrix(xFilas, 12) = Format(CDate(.Cells(A, 25)), FORMAT_DATE)  '--Fecha Sistema

            Fg1.TextMatrix(xFilas, 13) = NulosC(.Cells(A, 17)) '--Moneda
            Fg1.TextMatrix(xFilas, 14) = NulosN(.Cells(A, 18)) '--Tipo de Cambio
            
            '--Si documento es nota de credito, mostrar los importes en positivo
            Fg1.TextMatrix(xFilas, 15) = Abs(NulosN(.Cells(A, 11))) '--Imp. Afecto
            Fg1.TextMatrix(xFilas, 16) = Abs(NulosN(.Cells(A, 12))) '--Imp Inafecto
            Fg1.TextMatrix(xFilas, 17) = Abs(NulosN(.Cells(A, 13))) '--Imp. Retencion(para honorarios)
            Fg1.TextMatrix(xFilas, 18) = Abs(NulosN(.Cells(A, 14))) '--Imp Igv
            Fg1.TextMatrix(xFilas, 19) = Abs(NulosN(.Cells(A, 10))) '--Imp. Total

            Fg1.TextMatrix(xFilas, 20) = NulosC(.Cells(A, 19)) '--Glosa
            Fg1.TextMatrix(xFilas, 21) = NulosC(.Cells(A, 35)) '--Ruc cliente
            Fg1.TextMatrix(xFilas, 22) = NulosC(.Cells(A, 36)) '--Razón Social del Cliente
            Fg1.TextMatrix(xFilas, 23) = NulosC(.Cells(A, 33)) '--Orden de Despacho


            If NulosC(.Cells(A, 8)) = "" Then
                GRID_COLOR_FONDO Fg1, xFilas, 0, xFilas, Fg1.Cols - 1, vbRed
            End If


        Next A
    End With

    DoEvents

    '*******************************************************************************************************************
    'CARGAMOS EL DETALLE DE LA VENTA
    objExcel.WindowState = 2
    objExcel.Workbooks.Open Trim(TxtArchivo2.Text)
    
    xFilaIni = 9
    xNumFilas = 0
    
    Fg2.Rows = 1
    With objExcel.ActiveSheet
        'DETERMINAMOS EL NUMERO DE FILAS CON DATOS
        LblBarra.Caption = "Calculando número de registros"
        ProgressBar2.Max = 32000
        ProgressBar2.Value = 1
        For A = xFilaIni To 32000
            ProgressBar2.Value = A
            '--verificar campo proveeodor
            If NulosC(.Cells(A, 6)) <> "" Then
                xNumFilas = xNumFilas + 1
            Else
                Exit For
            End If
        Next A

        xNumFilas = xNumFilas + xFilaIni
        
        ProgressBar2.Max = xNumFilas
        
        LblBarra.Caption = "Cargando registros - Detalle de Documentos"
        
        DoEvents
        
        For A = xFilaIni To xNumFilas
            ProgressBar2.Value = A
            DoEvents
            
            '--verificar si proveedor es nulo
            If NulosC(.Cells(A, 6)) = "" Then Exit For
            
            '--Considerar solo detalle de documentos de activos
            If LCase(NulosC(.Cells(A, 31))) <> "anulado" Then
    
                Fg2.Rows = Fg2.Rows + 1
                xFilas = Fg2.Rows - 1
                
                Fg2.TextMatrix(xFilas, 1) = NulosC(.Cells(A, 5)) '--Ruc proveedor
                Fg2.TextMatrix(xFilas, 2) = NulosC(.Cells(A, 6)) '--Razón social del proveedor
                Fg2.TextMatrix(xFilas, 3) = NulosC(.Cells(A, 9)) '--Tipo de Documento
                Fg2.TextMatrix(xFilas, 4) = NulosC(.Cells(A, 7)) '--N° Serie
                Fg2.TextMatrix(xFilas, 5) = NulosC(.Cells(A, 8)) '--N° Documento
                '--Si documento es nota de credito, mostrar los importes en positivo
                Fg2.TextMatrix(xFilas, 6) = Abs(NulosN(.Cells(A, 14))) '--Imp. Afecto
                Fg2.TextMatrix(xFilas, 7) = Abs(NulosN(.Cells(A, 15))) '--Imp Inafecto
                Fg2.TextMatrix(xFilas, 8) = Abs(NulosN(.Cells(A, 18))) '--Imp. Retencion(para honorarios)
                Fg2.TextMatrix(xFilas, 9) = Abs(NulosN(.Cells(A, 16))) '--Imp Igv
                
                Fg2.TextMatrix(xFilas, 11) = NulosC(.Cells(A, 19)) '--vinc1=Empresa
                Fg2.TextMatrix(xFilas, 12) = NulosC(.Cells(A, 23)) '--vinc2=Centro Costo
                Fg2.TextMatrix(xFilas, 13) = NulosC(.Cells(A, 24)) '--vinc3=Cuenta
                Fg2.TextMatrix(xFilas, 14) = NulosC(.Cells(A, 25)) '--vinc4=SubCuenta
    
                '--agregando registros al detalle
                RstTmpDet.AddNew
                RstTmpDet("numruc") = Fg2.TextMatrix(xFilas, 1)
                RstTmpDet("documento") = Fg2.TextMatrix(xFilas, 3)
                RstTmpDet("numser") = Fg2.TextMatrix(xFilas, 4)
                RstTmpDet("numdoc") = Fg2.TextMatrix(xFilas, 5)
                            
                '----
                RstTmpDet("impafec") = NulosN(Fg2.TextMatrix(xFilas, 6))
                RstTmpDet("impexon") = NulosN(Fg2.TextMatrix(xFilas, 7))
                RstTmpDet("impret") = NulosN(Fg2.TextMatrix(xFilas, 8))
                RstTmpDet("impigv") = NulosN(Fg2.TextMatrix(xFilas, 9))
                
                If InStr(RstTmpDet("documento"), "hon") = 0 Then
                    RstTmpDet("imptot") = RstTmpDet("impafec") + RstTmpDet("impexon") + RstTmpDet("impigv")
                Else
                    '--Si documento es recibo por honorario y se aplica retencion, este se se aplica al total del documento
                    RstTmpDet("imptot") = RstTmpDet("impafec") - RstTmpDet("impret")
                End If
                
                RstTmpDet("vinc1") = NulosC(Fg2.TextMatrix(xFilas, 11))
                RstTmpDet("vinc2") = NulosC(Fg2.TextMatrix(xFilas, 12))
                RstTmpDet("vinc3") = NulosC(Fg2.TextMatrix(xFilas, 13))
                RstTmpDet("vinc4") = NulosC(Fg2.TextMatrix(xFilas, 14))
    
                RstTmpDet.Update
                
            End If
        Next A
    End With
    '---------------------------------------------

    Frame5.Visible = False
    MsgBox "El proceso terminó con éxito", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
    objExcel.WindowState = 2
    objExcel.Workbooks.Close

    Set objExcel = Nothing
    Exit Sub
error:
'Resume
    Frame5.Visible = False
    objExcel.Workbooks.Close
    If Err.Number = 424 Then
        MsgBox Err.Description & vbCr & "El archivo fue cerrado antes de terminar de importar, vuelva a importar nuevamente.", vbCritical, xTitulo
    Else
        MsgBox Err.Description & vbCr & Err.Source, vbCritical, xTitulo
    End If
    Fg1.Rows = 1
    Fg2.Rows = 1
    Set objExcel = Nothing

End Sub




Private Sub VerDatosxDepurarDet()
    
    Dim nSQL As String
    Dim nSQLFiltro As String
    Dim xRs As New ADODB.Recordset
    '--------------------------------------------------------
    Fg7.Rows = 1
    DoEvents
    '--------------------------------------------------------
    '--Generando filtro
    If TabOne3.CurrTab = 0 Then
        If Fg3.Rows = Fg3.FixedRows Then
            MsgBox "No hay registros observados", vbInformation, xTitulo
            Exit Sub
        End If
        
        If Fg3.Row < Fg3.FixedRows Then
            MsgBox "Seleccione un registro", vbInformation, xTitulo
            Exit Sub
        End If
        
        nSQLFiltro = " and tra_documento.proveedor ='" & NulosC(Fg3.TextMatrix(Fg3.Row, 3)) & "'"
        
    ElseIf TabOne3.CurrTab = 1 Then
        nSQLFiltro = " and tra_documento.id in (SELECT tra_documento.id " _
                    & " FROM tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet " _
                    & " WHERE (((tra_documento.estado)=1) AND " _
                    & " ((tra_documentodet.vinc1)='" & NulosC(Fg5.TextMatrix(Fg5.Row, 2)) & "') AND " _
                    & " ((tra_documentodet.vinc2)='" & NulosC(Fg5.TextMatrix(Fg5.Row, 3)) & "') AND " _
                    & " ((tra_documentodet.vinc3)='" & NulosC(Fg5.TextMatrix(Fg5.Row, 4)) & "') AND " _
                    & " ((tra_documentodet.vinc4)='" & NulosC(Fg5.TextMatrix(Fg5.Row, 5)) & "')) )"

    Else
        Exit Sub
    End If
    
    
    '--------------------------------------------------------
    '--Cargando detalle documentos
    nSQL = "SELECT tra_documento.id, tra_documento.corr, tra_documento.rucprov, tra_documento.proveedor, tra_documento.documento, mae_documento.abrev, [tra_documento].[numser] & '-' & [tra_documento].[numdoc] AS numerodoc, tra_documento.fchdoc, mae_moneda.simbolo, tra_documento.tipcam, tra_documento.imptot, tra_documento.glosa " _
        + vbCr + " FROM (tra_documento LEFT JOIN mae_documento ON tra_documento.tipdoc = mae_documento.id) LEFT JOIN mae_moneda ON tra_documento.idmon = mae_moneda.id " _
        + vbCr + " Where (((tra_documento.estado) = 1)) " & nSQLFiltro _
        + vbCr + " ORDER BY tra_documento.proveedor, tra_documento.fchdoc "

    RST_Busq xRs, nSQL, xCon
    
    If xRs.RecordCount <> 0 Then
        Do While Not xRs.EOF
            Fg7.Rows = Fg7.Rows + 1
            Fg7.TextMatrix(Fg7.Rows - 1, 1) = Format(NulosC(xRs("corr")), "0000")
            Fg7.TextMatrix(Fg7.Rows - 1, 2) = NulosC(xRs("proveedor"))
            Fg7.TextMatrix(Fg7.Rows - 1, 3) = NulosC(xRs("abrev"))
            Fg7.TextMatrix(Fg7.Rows - 1, 4) = NulosC(xRs("numerodoc"))
            Fg7.TextMatrix(Fg7.Rows - 1, 5) = NulosC(xRs("fchdoc"))
            Fg7.TextMatrix(Fg7.Rows - 1, 6) = NulosC(xRs("simbolo"))
            Fg7.TextMatrix(Fg7.Rows - 1, 7) = Format(NulosN(xRs("imptot")), FORMAT_MONTO)
            Fg7.TextMatrix(Fg7.Rows - 1, 8) = NulosC(xRs("glosa"))
            xRs.MoveNext
        Loop
    End If
    Set xRs = Nothing

    '--verificar si hay datos observados para mostrar ventana
    If Fg7.Rows <> Fg7.FixedRows Then
        FraDetalle.Visible = True
        FraDetalle.Left = 1470
        FraDetalle.Top = 1380
        Command2.SetFocus
    End If
    
    If Fg7.Rows = Fg7.FixedRows Then
    
        MsgBox "No hay datos en el detalle", vbInformation, xTitulo
        
    End If
    
    
    
    
    DoEvents
End Sub





Public Function xTDB_FiltroGenerar(TDGRID As Object, Rst As ADODB.Recordset)
''    Dim tmp As String
''    Dim N As Integer
    Dim k As Integer
''    Dim C As Integer
    
    If Rst.State = 0 Then Exit Function
    
    
'--------------------
On Error GoTo errhandler
Dim Cols  As TrueOleDBGrid70.Columns
Dim Col As TrueOleDBGrid70.Column
Set Cols = Dg2.Columns
Dim C As Integer
Dim tmp As String
Dim N As Integer

C = Dg2.Col
Dg2.HoldFields

For Each Col In Cols
    If Trim(Col.FilterText) <> "" Then
        N = N + 1
        If N > 1 Then
            tmp = tmp & " AND "
        End If
        tmp = tmp & Col.DataField & " like '*" & Col.FilterText & "*'"
    End If
Next Col
Rst.Filter = tmp
Dg2.Col = C
Dg2.EditActive = True
Exit Function
errhandler:
MsgBox Err.Source & " : " & vbCrLf & Err.Description

'--------------------
End Function


Private Sub TotalRegistros(xDg As TrueOleDBGrid70.TDBGrid, xRst As ADODB.Recordset)
    '--11/11/26
    '--Permitir mostrar la cantidad de registros
On Error Resume Next

    xDg.Splits(0).Columns("fchdoc").FooterText = 0
    If xRst.State = 0 Then Exit Sub
    
    xDg.Splits(0).Columns("fchdoc").FooterText = NulosN(xRst.RecordCount)
    
Err.Clear

End Sub


Private Sub pExportar()
Exit Sub
''    '--11/11/26
''    Dim nSQL As String
''    Dim oExport As New SGI2_funciones.formularios
''    Dim RstTmp  As New ADODB.Recordset
''    Dim xCampos(15, 3) As String
''
''    TabOne1.CurrTab = 0
''
''    '0::Nombre a Mostrar;
''    '1::nombre de Campo del Rst;
''    '2::alineacion(0::derecha, 1::centro, 2::izquierda);
''    '3::ancho de columna
''    '--obs: el rst puede tener mas columnas solo se consideran los campos del array
''    xCampos(0, 0) = "Id":                   xCampos(0, 1) = "id":           xCampos(0, 2) = 2:      xCampos(0, 3) = "500"
''    xCampos(1, 0) = "Código":               xCampos(1, 1) = "codigo":       xCampos(1, 2) = 0:      xCampos(1, 3) = "900"
''    xCampos(2, 0) = "Apellidos y Nombres":  xCampos(2, 1) = "nombre":       xCampos(2, 2) = 0:      xCampos(2, 3) = "2500"
''    xCampos(3, 0) = "Sexo":                 xCampos(3, 1) = "sexo":         xCampos(3, 2) = 0:      xCampos(3, 3) = "500"
''    xCampos(4, 0) = "T.D.":                 xCampos(4, 1) = "docabrev":     xCampos(4, 2) = 0:      xCampos(4, 3) = "500"
''    xCampos(5, 0) = "Num. Doc":             xCampos(5, 1) = "numdoc":       xCampos(5, 2) = 0:      xCampos(5, 3) = "1200"
''    xCampos(6, 0) = "Fch. Nac.":            xCampos(6, 1) = "fchnac":       xCampos(6, 2) = 1:      xCampos(6, 3) = "1100"
''    xCampos(7, 0) = "Fch.Ingreso":          xCampos(7, 1) = "fching":       xCampos(7, 2) = 1:      xCampos(7, 3) = "1100"
''    xCampos(8, 0) = "Categoría":            xCampos(8, 1) = "catnomcorto":  xCampos(8, 2) = 0:      xCampos(8, 3) = "500"
''    xCampos(9, 0) = "Area":                 xCampos(9, 1) = "area":         xCampos(9, 2) = 0:      xCampos(9, 3) = "1300"
''    xCampos(10, 0) = "Cargo":               xCampos(10, 1) = "cargo":       xCampos(10, 2) = 0:     xCampos(10, 3) = "1300"
''    xCampos(11, 0) = "Pago H.N.":           xCampos(11, 1) = "paghornor":   xCampos(11, 2) = 2:     xCampos(11, 3) = "800"
''    xCampos(12, 0) = "Pago H.E.":           xCampos(12, 1) = "paghorext":   xCampos(12, 2) = 2:     xCampos(12, 3) = "800"
''    xCampos(13, 0) = "Estado":              xCampos(13, 1) = "estado":      xCampos(13, 2) = 0:     xCampos(13, 3) = "900"
''    xCampos(14, 0) = "Fch. Cese":           xCampos(14, 1) = "fchcese":     xCampos(14, 2) = 1:     xCampos(14, 3) = "1100"
''    xCampos(15, 0) = "Tip. Planilla":       xCampos(15, 1) = "destippla":   xCampos(15, 2) = 1:     xCampos(15, 3) = "1100"
''    '***************************
''
''    Set RstTmp = RstFrm
''    '***************************
''
''    oExport.VSFlexGrid_Exportar_MSExcel xCon, , "NOMINA DE PERSONAL", "", "", "NOMINA DE PERSONAL", RstTmp, xCampos
''    Set oExport = Nothing
''    Set RstTmp = Nothing
    
End Sub


Private Sub CmdActualizar_Click()
'--11/11/24
'--Temporal, desagregar de glosa tipos de gasto reembolsables
Dim nSQL As String

Me.MousePointer = vbHourglass
DoEvents

'xCon.BeginTrans

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'DESCARGA' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%descarga%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'ALMACENAJE' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%ALMACENAJE%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'FLETE' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%FLETE%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'HANDLING' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%HANDLING%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'POLIZA SEGURO COBERCONT' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%SEGURO COBERCONT%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'POLIZA SEGURO COBERCONT' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%POLIZA%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'DEVOLUCION DE CONTENEDOR' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%DEVOLUCION%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'TRAMITE DOCUMENTARIO' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%TRAMITE%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'TRAMITE DOCUMENTARIO' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like 'DOCUMENT') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'DELIVERY ORDER' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%DELIVERY%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'TRANSPORTE DE CARGA' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%TRANSPORTE%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'CUADRILLA' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%CUADRILLA%') AND ((tra_documentodet.iditem)=0)); "
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'MANIPULEO' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%MANIPULEO%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'SERVICIO DE MONTACARGA' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%MONTACARGA%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'SERVICIOS EXTRAORDINARIOS' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%SERVICIOS EXTRAORDINARIOS%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'STRECH FILM / CINTAS DE EMBALAJE' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%STRECH%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'TRASEGADO' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%TRASEGADO%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'AFORO FISICO' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%AFORO%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'B/L Transmission Fee' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%B/L%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'B/L Transmission Fee' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '% BL%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'CONTROL DE PRECINTOS' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%PRECINTOS%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'ESTIBA' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%ESTIBA%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'Gastos Administrativos' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%Administrativos%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'Gastos Administrativos' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%adm.%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'Gate IN/OUT' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%Gate IN%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'Sobreestadía de Contenedor' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%Sobreestadía%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'TRACCION' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%TRACCION%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'Transporte Cont. Vacio otro Terminal' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%VACIO%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'Equipo por Movilización de Carga' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%MOVILIZACION%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'Gastos Administrativos' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%despacho%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'DEVOLUCION DE CONTENEDOR' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%RECEPCION%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'GASTOS ADMINISTRATIVOS' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%ADM%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'Sobreestadía de Contenedor' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%SOBREESTADIA%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'DEVOLUCION DE CONTENEDOR' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%DEVOL%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'Gastos Administrativos' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%FLAT%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'Gastos Administrativos' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%INSPECCION%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'Gastos Administrativos' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%DESPACHO%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'DEVOLUCION DE CONTENEDOR' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%RECEPCION%') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'Inspección de Carga' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%INSPECCION%') AND ((tra_documentodet.iditem)=0)); "
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'Embarque Directo Naviera' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%EMBARQUE%') AND ((tra_documentodet.iditem)=0)); "
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'GASTOS ADMINISTRATIVOS' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%AGENCIAMIENTO%') AND ((tra_documentodet.iditem)=0)); "
xCon.Execute nSQL

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'STRECH FILM' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documento.glosa) Like '%EMBALAJE%') AND ((tra_documentodet.iditem)=0)); "
xCon.Execute nSQL




Me.MousePointer = vbDefault
DoEvents

'xCon.CommitTrans

MsgBox "Los datos se actualizaron correctamente", vbInformation, xTitulo

Exit Sub
xError:
'xCon.RollbackTrans
MsgBox Err.Description & vbCr & Err.Source & xTitulo
End Sub

Private Sub CmdActualizar1_Click()
'--11/11/24
'--Temporal, desagregar de glosa tipos de gasto reembolsables
Dim nSQL As String

Me.MousePointer = vbHourglass
DoEvents

'xCon.BeginTrans

nSQL = "UPDATE tra_documento INNER JOIN tra_documentodet ON tra_documento.id = tra_documentodet.iddet SET tra_documentodet.vinc4 = 'Gastos Administrativos' " _
    + vbCr + " WHERE (((tra_documentodet.vinc4)='GASTOS REEMBOLSABLES') AND ((tra_documentodet.iditem)=0));"
xCon.Execute nSQL



Me.MousePointer = vbDefault
DoEvents

'xCon.CommitTrans

MsgBox "Los datos se actualizaron correctamente", vbInformation, xTitulo

Exit Sub
xError:
'xCon.RollbackTrans
MsgBox Err.Description & vbCr & Err.Source & xTitulo
End Sub



