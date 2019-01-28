VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SIZERONE.OCX"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "VSFLEX7.OCX"
Object = "{D1333493-26F3-11D5-B046-E1A96EACB52A}#1.0#0"; "AspaTextBoxFecha.ocx"
Begin VB.Form FrmCentroCostosUni 
   Caption         =   "Contabilidad - Gastos Unificados"
   ClientHeight    =   7560
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   7560
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00C0FFC0&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6945
      Left            =   -930
      TabIndex        =   21
      Top             =   630
      Visible         =   0   'False
      Width           =   11790
      Begin VB.CommandButton Command4 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   5580
         Picture         =   "FrmCentroCostosUni.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   25
         ToolTipText     =   "Exportar a Excel"
         Top             =   6300
         Width           =   630
      End
      Begin VB.CommandButton Command3 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   4920
         Picture         =   "FrmCentroCostosUni.frx":0B0A
         Style           =   1  'Graphical
         TabIndex        =   24
         Top             =   6300
         Width           =   630
      End
      Begin VB.CommandButton Command2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   555
         Left            =   6360
         Picture         =   "FrmCentroCostosUni.frx":0E14
         Style           =   1  'Graphical
         TabIndex        =   23
         ToolTipText     =   "Exportar a Excel"
         Top             =   6300
         Width           =   630
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg2 
         Height          =   5790
         Left            =   45
         TabIndex        =   22
         Top             =   450
         Width           =   11685
         _cx             =   20611
         _cy             =   10213
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
         SelectionMode   =   0
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
         FormatString    =   $"FrmCentroCostosUni.frx":111E
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
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Gastos Unificados"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   240
         Left            =   195
         TabIndex        =   26
         Top             =   60
         Width           =   1785
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   300
         Left            =   30
         Top             =   45
         Width           =   11715
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   1
         X1              =   -15
         X2              =   11745
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   0
         X1              =   30
         X2              =   11790
         Y1              =   6930
         Y2              =   6930
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderWidth     =   2
         Index           =   1
         X1              =   11775
         X2              =   11775
         Y1              =   15
         Y2              =   6945
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Index           =   0
         X1              =   15
         X2              =   15
         Y1              =   15
         Y2              =   6945
      End
   End
   Begin VB.Frame Frame1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   960
      Left            =   0
      TabIndex        =   8
      Top             =   0
      Width           =   11880
      Begin VB.CommandButton Command1 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   8355
         Picture         =   "FrmCentroCostosUni.frx":1188
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   240
         Width           =   630
      End
      Begin VB.CommandButton CmdUnificado 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   9015
         Picture         =   "FrmCentroCostosUni.frx":15CA
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   240
         Width           =   630
      End
      Begin VB.CommandButton CmdSalir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   11115
         Picture         =   "FrmCentroCostosUni.frx":18D4
         Style           =   1  'Graphical
         TabIndex        =   7
         ToolTipText     =   "Exportar a Excel"
         Top             =   240
         Width           =   630
      End
      Begin VB.CommandButton CmdImprimir 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   9675
         Picture         =   "FrmCentroCostosUni.frx":1BDE
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   240
         Width           =   630
      End
      Begin VB.CommandButton CmdExp 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   10335
         Picture         =   "FrmCentroCostosUni.frx":1EE8
         Style           =   1  'Graphical
         TabIndex        =   6
         ToolTipText     =   "Exportar a Excel"
         Top             =   240
         Width           =   630
      End
      Begin VB.CommandButton CmdBusMon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1695
         Picture         =   "FrmCentroCostosUni.frx":29F2
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   585
         Width           =   240
      End
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchIni 
         Height          =   300
         Left            =   1170
         TabIndex        =   0
         Top             =   240
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
      Begin AspaTextBoxFecha.TextBoxFecha TxtFchFin 
         Height          =   300
         Left            =   3540
         TabIndex        =   1
         Top             =   240
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
      Begin VB.TextBox TxtIdMon 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Left            =   1170
         TabIndex        =   2
         Text            =   "TxtIdMon"
         Top             =   555
         Width           =   795
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Inicio"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   150
         TabIndex        =   13
         Top             =   285
         Width           =   735
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Fch. Final"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2685
         TabIndex        =   12
         Top             =   285
         Width           =   690
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
         Left            =   1995
         TabIndex        =   11
         Top             =   555
         Width           =   2820
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Moneda"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   150
         TabIndex        =   10
         Top             =   585
         Width           =   585
      End
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   6555
      Left            =   0
      TabIndex        =   14
      Top             =   1005
      Width           =   11880
      _cx             =   20955
      _cy             =   11562
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
      Appearance      =   2
      MousePointer    =   0
      _ConvInfo       =   1
      Version         =   700
      BackColor       =   13160660
      ForeColor       =   0
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   8388608
      Caption         =   "Tab&1|Tab&2|Tab&3"
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Index           =   2
         Left            =   12525
         TabIndex        =   19
         Top             =   375
         Width           =   11790
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   6030
            Index           =   2
            Left            =   15
            TabIndex        =   20
            Top             =   90
            Width           =   11775
            _cx             =   20770
            _cy             =   10636
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
            SelectionMode   =   0
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
            FormatString    =   $"FrmCentroCostosUni.frx":2B24
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Index           =   1
         Left            =   45
         TabIndex        =   17
         Top             =   375
         Width           =   11790
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   6030
            Index           =   1
            Left            =   15
            TabIndex        =   18
            Top             =   90
            Width           =   11775
            _cx             =   20770
            _cy             =   10636
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
            SelectionMode   =   0
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
            FormatString    =   $"FrmCentroCostosUni.frx":2BAA
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
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   6135
         Index           =   0
         Left            =   -12435
         TabIndex        =   15
         Top             =   375
         Width           =   11790
         Begin VSFlex7Ctl.VSFlexGrid Fg1 
            Height          =   6030
            Index           =   0
            Left            =   15
            TabIndex        =   16
            Top             =   90
            Width           =   11775
            _cx             =   20770
            _cy             =   10636
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
            SelectionMode   =   0
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
            FormatString    =   $"FrmCentroCostosUni.frx":2C30
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
End
Attribute VB_Name = "FrmCentroCostosUni"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim xCon1 As New ADODB.Connection
Dim xCon2 As New ADODB.Connection

Dim SeEjecuto As Boolean

Private Sub CmdBusMon_Click()
    Dim xform As New eps_librerias.FormBuscar
    Dim xRs As New ADODB.Recordset
    Dim xCampos2(2, 4) As String
    
    xCampos2(0, 0) = "Moneda":        xCampos2(0, 1) = "descripcion":      xCampos2(0, 2) = "4000":         xCampos2(0, 3) = "C"
    xCampos2(1, 0) = "Abreviatura":   xCampos2(1, 1) = "simbolo":          xCampos2(1, 2) = "1500":         xCampos2(1, 3) = "C"
    
    xform.SQLCad = "SELECT mae_moneda.* FROM mae_moneda ORDER BY descripcion"
    
    xform.Titulo = "Buscando Monedas"
    
    xform.FormaBusca = Principio
    xform.Criterio = ""
    xform.Ordenado = "descripcion"
    xform.CampoBusca = "descripcion"
    Set xform.Coneccion = xCon
    Set xRs = xform.BuscarReg(xCampos2)
    
    If xRs.State = 1 Then
        TxtIdMon.Text = xRs("id")
        LblMoneda.Caption = xRs("descripcion")
        Command1.SetFocus
    End If
    Set xform = Nothing
    Set xRs = Nothing
End Sub

Private Sub CmdExp_Click()
    If Fg1(TabOne1.CurrTab).Rows = 1 Then
        MsgBox "No se han registrado movimientos de centro de costo en la empresa :" + TabOne1.TabCaption(TabOne1.CurrTab), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim xFun As New SGI2_funciones.formularios
    xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg1(TabOne1.CurrTab), "REPORTE DE GASTOS", "Periodo Del : " + TxtFchIni.Valor + " Al : " + TxtFchFin.Valor, "Empresa :" + TabOne1.TabCaption(TabOne1.CurrTab), "gastos.xls"
    Set xFun = Nothing
End Sub

Private Sub CmdImprimir_Click()
    If Fg1(TabOne1.CurrTab).Rows = 1 Then
        MsgBox "No se han registrado movimientos de centro de costo en la empresa :" + TabOne1.TabCaption(TabOne1.CurrTab), vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        Exit Sub
    End If
    
    Dim xFun As New SGI2_funciones.formularios
    xFun.Imprimir_x_VSFlexGrid Fg2, "REPORTE DE GASTOS", "Empresa :" + TabOne1.TabCaption(TabOne1.CurrTab), "Periodo Del : " + TxtFchIni.Valor + " Al : " + TxtFchFin.Valor, True, True
    Set xFun = Nothing
End Sub

Private Sub CmdSalir_Click()
    Unload Me
End Sub

Private Sub CmdUnificado_Click()
    Frame3.Left = 60
    Frame3.Top = 285
    Frame3.Visible = True
    Consolidar
End Sub

Sub Consolidar()
    Dim A, B As Integer
    
    For A = 1 To Fg1(1).Rows - 1
        
    Next A
End Sub

Private Sub Command1_Click()
    If NulosC(TxtFchIni.Valor) = "" Then
        MsgBox "No ha especificado la fecha de inicio de la consulta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
    
    If NulosC(TxtFchFin.Valor) = "" Then
        MsgBox "No ha especificado la fecha de final de la consulta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchFin.SetFocus
        Exit Sub
    End If
    
    If NulosN(TxtIdMon.Text) = 0 Then
        MsgBox "No ha especificado la moneda de consulta", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtIdMon.SetFocus
        Exit Sub
    End If
    
    If CDate(TxtFchIni.Valor) > CDate(TxtFchFin.Valor) Then
        MsgBox "La fecha de inicio no puede ser mayor a la fecha fina", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
        TxtFchIni.SetFocus
        Exit Sub
    End If
        
    Dim A As Integer
    Dim xIndex As Integer
    Dim xRuta As String
    Dim RstEmp As New ADODB.Recordset
    
    Fg2.Cols = 3
    Fg2.Rows = 1
    Set xCon1 = AbrirConecciones(AP_RUTABD + "data.mdb")
    
    RST_Busq RstEmp, "SELECT mae_empresa.* From mae_empresa WHERE (((mae_empresa.anotra)=2007) AND ((mae_empresa.activo)=-1))", xCon1
    
    If RstEmp.RecordCount <> 0 Then
        xIndex = 0
        RstEmp.MoveFirst
        
        For A = 1 To RstEmp.RecordCount
            TabOne1.TabCaption(xIndex) = " " & Trim(RstEmp("abrevia")) & " "
            TabOne1.TabVisible(xIndex) = True
            
            xRuta = AP_RUTABD + Trim(RstEmp("ruta"))
            
            Set xCon2 = Nothing
            Set xCon2 = AbrirConecciones(xRuta)
'                If VerificarPlanesActivos > 1 Then
'                    MsgBox "No se puede consultar el plan de produccion unificado, existe mas de 1 plan activo en la empresa " + RstEmp("abrevia") + Chr(13) _
'                        & "Verifique que solo exista un plan de produccion activo.", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
'                    Set xCon1 = Nothing
'                    Set xCon2 = Nothing
'                    Unload Me
'                    Exit For
'                End If
            
            Gerencial xIndex, RstEmp("abrevia")
            RstEmp.MoveNext
            
            If RstEmp.EOF = True Then
                Exit For
            End If
            xIndex = xIndex + 1
        Next A
    End If
    
    Fg2.Cols = Fg2.Cols + 1
    Fg2.ColAlignment(Fg2.Cols - 1) = flexAlignRightTop
    Fg2.TextMatrix(0, Fg2.Cols - 1) = "TOTAL"
    Fg2.ColWidth(Fg2.Cols - 1) = 1400
    Fg2.FixedAlignment(Fg2.Cols - 1) = flexAlignCenterTop
    CalcularTotalUnificado
    TabOne1.CurrTab = 0
End Sub

Sub CalcularTotalUnificado()
    Dim A, B As Integer
    Dim Total As Double
    
    For B = 1 To Fg2.Rows - 1
        Total = 0
        For A = 3 To Fg2.Cols - 2
            Total = Total + NulosN(Fg2.TextMatrix(B, A))
        Next A
        Fg2.TextMatrix(B, Fg2.Cols - 1) = Format(Total, FORMAT_MONTO)
    Next B
    
    Fg2.Rows = Fg2.Rows + 1
    
    For A = 3 To Fg2.Cols - 1
        Total = 0
        For B = 1 To Fg2.Rows - 2
            Total = Total + NulosN(Fg2.TextMatrix(B, A))
        Next B
        Fg2.TextMatrix(Fg2.Rows - 1, A) = Format(Total, FORMAT_MONTO)
        FORMATO_CELDA Fg2, CLng(Fg2.Rows - 1), CLng(A), , True, , ""
    Next A
    GRID_COLOR_FONDO Fg2, 1, 3, Fg2.Rows - 1, Fg2.Cols - 1, &HFEFBEB, flexFillRepeat
End Sub
Private Sub Command2_Click()
    Frame3.Visible = False
End Sub

Private Sub Command3_Click()
    Dim xFun As New SGI2_funciones.formularios
    xFun.Imprimir_x_VSFlexGrid Fg2, "REPORTE DE GASTOS", " ", "Periodo Del : " + TxtFchIni.Valor + " Al : " + TxtFchFin.Valor, True, True
    Set xFun = Nothing
End Sub

Private Sub Command4_Click()
    Dim xFun As New SGI2_funciones.formularios
    xFun.VSFlexGrid_Exportar_MSExcel xCon, Fg2, "REPORTE DE GASTOS", "Periodo Del : " + TxtFchIni.Valor + " Al : " + TxtFchFin.Valor, "", "gastos.xls"
    Set xFun = Nothing
End Sub

Private Sub Form_Activate()
    If SeEjecuto = False Then
        SeEjecuto = True
    End If
End Sub

Private Sub Form_Load()
    Dim xIndex, A As Integer
    
    SeEjecuto = False
    TxtFchIni.Valor = ""
    TxtFchFin.Valor = ""
    TxtIdMon.Text = ""
    LblMoneda.Caption = ""
    Fg1(0).Rows = 1
    Fg1(1).Rows = 1
    Fg1(2).Rows = 1
    
    xIndex = 0
    
    For A = 1 To 3
        TabOne1.TabVisible(xIndex) = False
        Frame2(xIndex).BackColor = &H8000000F
        Fg1(xIndex).Rows = 1

        xIndex = xIndex + 1
    Next A
    
    Frame3.BackColor = &H8000000F
    Fg1(0).SelectionMode = flexSelectionByRow
    Fg1(1).SelectionMode = flexSelectionByRow
    Fg1(2).SelectionMode = flexSelectionByRow
    
    Fg2.SelectionMode = flexSelectionByRow
End Sub


Sub Gerencial(Index As Integer, Abrev As String)
    Dim xRst1 As New ADODB.Recordset
    Dim xRst2 As New ADODB.Recordset
    Dim xSql As String
    Dim A As Integer
    Dim xTotal As Double
    
    Fg1(Index).Cols = 4
    If TxtIdMon.Text = "1" Then
        xSql = "SELECT con_centrocosto2.codigo, con_centrocosto2.descripcion, Sum(IIf([com_compras]![idmon]=1,[com_comprascosto]![impcos],[com_comprascosto]![impcos]*[con_tc]![impven])) AS total " _
            & " FROM (com_compras RIGHT JOIN ((con_centrocosto LEFT JOIN com_comprascosto ON con_centrocosto.id = com_comprascosto.idcencos) RIGHT JOIN con_centrocosto2 " _
            & " ON con_centrocosto.idcencos2 = con_centrocosto2.id) ON com_compras.id = com_comprascosto.idcom) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
            & " WHERE (((com_compras.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (com_compras.fchdoc)<=CDate('" & TxtFchFin.Valor & "'))) GROUP BY con_centrocosto2.codigo, con_centrocosto2.descripcion"
    End If
    If TxtIdMon.Text = "2" Then
        xSql = "SELECT con_centrocosto2.codigo, con_centrocosto2.descripcion, Sum(IIf([com_compras]![idmon]=2,[com_comprascosto]![impcos],[com_comprascosto]![impcos]/[con_tc]![impven])) AS total" _
            & " FROM (com_compras RIGHT JOIN ((con_centrocosto LEFT JOIN com_comprascosto ON con_centrocosto.id = com_comprascosto.idcencos) RIGHT JOIN con_centrocosto2 " _
            & " ON con_centrocosto.idcencos2 = con_centrocosto2.id) ON com_compras.id = com_comprascosto.idcom) LEFT JOIN con_tc ON com_compras.fchdoc = con_tc.fecha " _
            & " WHERE (((com_compras.fchdoc)>=CDate('" & TxtFchIni.Valor & "') And (com_compras.fchdoc)<=CDate('" & TxtFchFin.Valor & "'))) GROUP BY con_centrocosto2.codigo, con_centrocosto2.descripcion"
    End If
    
    RST_Busq xRst1, "SELECT con_centrocosto2.* From con_centrocosto2 ORDER BY con_centrocosto2.codigo", xCon2
    RST_Busq xRst2, xSql, xCon2
    
    Fg1(Index).Rows = 1
    If xRst1.RecordCount <> 0 Then
        xRst1.MoveFirst
        
        For A = 1 To xRst1.RecordCount
            Fg1(Index).Rows = Fg1(Index).Rows + 1
            Fg1(Index).TextMatrix(A, 0) = xRst1("iduni")
            Fg1(Index).TextMatrix(A, 1) = xRst1("codigo")
            Fg1(Index).TextMatrix(A, 2) = xRst1("descripcion")
            
            xRst2.Filter = adFilterNone
            If xRst2.RecordCount <> 0 Then
                xRst2.MoveFirst
                xRst2.Filter = "codigo = '" & xRst1("codigo") & "'"
            End If
            If xRst2.EOF = False Then
                Fg1(Index).TextMatrix(A, 3) = Format(xRst2("total"), FORMAT_MONTO)
                xTotal = xTotal + xRst2("total")
            End If
            
            xRst1.MoveNext
            If xRst1.EOF = True Then Exit For
        Next A
        
        Fg1(Index).Rows = Fg1(Index).Rows + 1
        Fg1(Index).TextMatrix(Fg1(Index).Rows - 1, 2) = "TOTAL ==>"
        Fg1(Index).TextMatrix(Fg1(Index).Rows - 1, 3) = Format(xTotal, FORMAT_MONTO)
    End If
    
    'calculamos el porcentaje de cada rubro
    Fg1(Index).Cols = Fg1(Index).Cols + 1
    Fg1(Index).ColWidth(Fg1(Index).Cols - 1) = 900
    
    Fg1(Index).TextMatrix(0, Fg1(Index).Cols - 1) = "Porcentaje"
    Dim xTotPor As Double
    
    For A = 1 To Fg1(Index).Rows - 2
        Fg1(Index).TextMatrix(A, Fg1(Index).Cols - 1) = (NulosN(Fg1(Index).TextMatrix(A, Fg1(Index).Cols - 2)) / xTotal) * 100
        
        xTotPor = xTotPor + NulosN(Fg1(Index).TextMatrix(A, Fg1(Index).Cols - 1))
        If NulosN(Fg1(Index).TextMatrix(A, Fg1(Index).Cols - 1)) = 0 Then
            Fg1(Index).TextMatrix(A, Fg1(Index).Cols - 1) = ""
        Else
            Fg1(Index).TextMatrix(A, Fg1(Index).Cols - 1) = Format(Fg1(Index).TextMatrix(A, Fg1(Index).Cols - 1), FORMAT_MONTO)
        End If
        
    Next A
    Fg1(Index).TextMatrix(Fg1(Index).Rows - 1, Fg1(Index).Cols - 1) = Format(xTotPor, FORMAT_MONTO)
    
    GRID_COLOR_FONDO Fg1(Index), 1, Fg1(Index).Cols - 2, Fg1(Index).Rows - 1, Fg1(Index).Cols - 1, &HFEFBEB, flexFillRepeat
    FORMATO_CELDA Fg1(Index), Fg1(Index).Rows - 1, Fg1(Index).Cols - 3, , True, , ""
    FORMATO_CELDA Fg1(Index), Fg1(Index).Rows - 1, Fg1(Index).Cols - 2, , True, , ""
    FORMATO_CELDA Fg1(Index), Fg1(Index).Rows - 1, Fg1(Index).Cols - 1, , True, , ""
    
    'Agregamos al grid unificado
    If Index = 0 Then
        Fg2.Rows = 1
        Fg2.Cols = Fg2.Cols + 1
        Fg2.FixedAlignment(Fg2.Cols - 1) = flexAlignCenterTop
        
        Fg2.ColWidth(Fg2.Cols - 1) = 1400
        Fg2.ColAlignment(Fg2.Cols - 1) = flexAlignRightTop
        Fg2.TextMatrix(0, Fg2.Cols - 1) = Abrev
        
        xRst1.MoveFirst
        For A = 1 To xRst1.RecordCount
            Fg2.Rows = Fg2.Rows + 1
            
            Fg2.TextMatrix(A, 1) = xRst1("codigo")
            Fg2.TextMatrix(A, 2) = xRst1("descripcion")
            
            
            xRst2.Filter = adFilterNone
            If xRst2.RecordCount <> 0 Then
                xRst2.MoveFirst
                xRst2.Filter = "codigo = '" & xRst1("codigo") & "'"
            End If
            If xRst2.EOF = False Then
                Fg2.TextMatrix(A, Fg2.Cols - 1) = Format(xRst2("total"), FORMAT_MONTO)
            End If
            
            xRst1.MoveNext
            
            If xRst1.EOF = True Then Exit For
        Next A
    Else
        Fg2.Cols = Fg2.Cols + 1
        Fg2.ColWidth(Fg2.Cols - 1) = 1400
        Fg2.FixedAlignment(Fg2.Cols - 1) = flexAlignCenterTop
        Fg2.ColAlignment(Fg2.Cols - 1) = flexAlignRightTop
        Fg2.TextMatrix(0, Fg2.Cols - 1) = Abrev
        
        xRst1.MoveFirst
        For A = 1 To xRst1.RecordCount
            xRst2.Filter = adFilterNone
            If xRst2.RecordCount <> 0 Then
                xRst2.MoveFirst
                xRst2.Filter = "codigo = '" & xRst1("codigo") & "'"
            End If
            If xRst2.EOF = False Then
                Fg2.TextMatrix(A, Fg2.Cols - 1) = Format(xRst2("total"), FORMAT_MONTO)
            End If
            
            xRst1.MoveNext
            
            If xRst1.EOF = True Then Exit For
        Next A
    End If
End Sub

Private Sub TxtIdMon_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys vbTab
    End If
End Sub

Private Sub TxtIdMon_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 116 Then
        CmdBusMon_Click
    End If
End Sub

Private Sub TxtIdMon_Validate(Cancel As Boolean)
    If NulosN(TxtIdMon.Text) <> 0 Then
        LblMoneda.Caption = Busca_Codigo(TxtIdMon.Text, "id", "descripcion", "mae_moneda", "N", xCon)
        If LblMoneda.Caption = "" Then
            TxtIdMon.Text = ""
        End If
    End If
End Sub
