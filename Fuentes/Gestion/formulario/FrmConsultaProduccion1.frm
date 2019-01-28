VERSION 5.00
Object = "{9E883861-2808-4487-913D-EA332634AC0D}#1.0#0"; "SizerOne.ocx"
Object = "{D76D7128-4A96-11D3-BD95-D296DC2DD072}#1.0#0"; "Vsflex7.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form FrmConsultaProduccion1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unificado - Consultar Produccion"
   ClientHeight    =   7665
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7665
   ScaleWidth      =   11895
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame15 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Caption         =   "Frame15"
      Height          =   285
      Left            =   6180
      TabIndex        =   39
      Top             =   7275
      Width           =   5625
      Begin VB.Shape Shape4 
         BackColor       =   &H000000FF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   180
         Left            =   3195
         Top             =   45
         Width           =   540
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "= Faltante de Produccion"
         Height          =   195
         Left            =   3840
         TabIndex        =   41
         Top             =   45
         Width           =   1785
      End
      Begin VB.Shape Shape3 
         BackColor       =   &H00C00000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   180
         Left            =   900
         Top             =   45
         Width           =   540
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         BackColor       =   &H00C0C0C0&
         Caption         =   "= Sobre Produccion"
         Height          =   195
         Left            =   1545
         TabIndex        =   40
         Top             =   45
         Width           =   1410
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Caption         =   "&H80000009&"
      Height          =   6660
      Left            =   7770
      TabIndex        =   32
      Top             =   -4950
      Visible         =   0   'False
      Width           =   11775
      Begin VB.CommandButton CmdSalir 
         Height          =   555
         Left            =   10815
         Picture         =   "FrmConsultaProduccion1.frx":0000
         Style           =   1  'Graphical
         TabIndex        =   36
         ToolTipText     =   "Salir"
         Top             =   6030
         Width           =   735
      End
      Begin VB.CommandButton CmdPrin 
         Height          =   555
         Left            =   10050
         Picture         =   "FrmConsultaProduccion1.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   35
         ToolTipText     =   "Exportar MSExcel"
         Top             =   6030
         Width           =   735
      End
      Begin VB.CommandButton Command6 
         Caption         =   ">>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   990
         TabIndex        =   34
         ToolTipText     =   "Agrandar columnas"
         Top             =   6045
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.CommandButton Command5 
         Caption         =   "<<"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         Left            =   225
         TabIndex        =   33
         ToolTipText     =   "Reducir columnas"
         Top             =   6045
         Visible         =   0   'False
         Width           =   735
      End
      Begin VSFlex7Ctl.VSFlexGrid Fg7 
         Height          =   5550
         Left            =   60
         TabIndex        =   37
         Top             =   420
         Width           =   11640
         _cx             =   20532
         _cy             =   9790
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   128
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483636
         BackColorAlternate=   16777215
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
         Cols            =   9
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   0
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"FrmConsultaProduccion1.frx":0E14
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
      Begin VB.Line Line1 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   1
         X1              =   0
         X2              =   11745
         Y1              =   15
         Y2              =   15
      End
      Begin VB.Line Line2 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   0
         X1              =   11760
         X2              =   11760
         Y1              =   15
         Y2              =   6660
      End
      Begin VB.Line Line2 
         BorderColor     =   &H80000009&
         BorderWidth     =   2
         Index           =   1
         X1              =   15
         X2              =   15
         Y1              =   0
         Y2              =   6645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H8000000C&
         BorderWidth     =   2
         Index           =   0
         X1              =   0
         X2              =   11745
         Y1              =   6645
         Y2              =   6645
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00800000&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00800000&
         Height          =   330
         Left            =   30
         Top             =   45
         Width           =   11700
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Consolidado de insumos"
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
         TabIndex        =   38
         Top             =   120
         Width           =   2055
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   8235
      Top             =   -30
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483637
      ImageWidth      =   16
      ImageHeight     =   15
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   13
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion1.frx":0F3E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion1.frx":1482
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion1.frx":15DC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion1.frx":196E
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion1.frx":1AF2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion1.frx":1F46
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion1.frx":205E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion1.frx":25A2
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion1.frx":2AE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion1.frx":2BFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion1.frx":2D0E
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion1.frx":3162
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmConsultaProduccion1.frx":32CE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11895
      _ExtentX        =   20981
      _ExtentY        =   609
      ButtonWidth     =   609
      ButtonHeight    =   556
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Mostrar plan  de abastecimiento unificado"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.ToolTipText     =   "Salir"
            ImageIndex      =   12
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin SizerOneLibCtl.TabOne TabOne1 
      Height          =   7275
      Left            =   0
      TabIndex        =   1
      Top             =   375
      Width           =   11895
      _cx             =   20981
      _cy             =   12832
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
      BackColor       =   -2147483633
      ForeColor       =   -2147483630
      FrontTabColor   =   -2147483633
      BackTabColor    =   -2147483633
      TabOutlineColor =   -2147483632
      FrontTabForeColor=   8388608
      Caption         =   "  &Consulta  |   &Detalle  |New Tab|New Tab|New Tab"
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
         Height          =   6855
         Index           =   0
         Left            =   45
         TabIndex        =   18
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6660
            Index           =   0
            Left            =   30
            TabIndex        =   19
            Top             =   180
            Width           =   11745
            _cx             =   20717
            _cy             =   11747
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
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
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   13160660
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   0
            Position        =   1
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
            Begin VB.Frame Frame5 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6300
               Index           =   0
               Left            =   15
               TabIndex        =   21
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg5 
                  Height          =   6030
                  Index           =   0
                  Left            =   0
                  TabIndex        =   22
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion1.frx":3816
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
            Begin VB.Frame Frame6 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6300
               Index           =   0
               Left            =   12360
               TabIndex        =   20
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg6 
                  Height          =   6030
                  Index           =   0
                  Left            =   0
                  TabIndex        =   27
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Cols            =   12
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion1.frx":396E
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   6855
         Index           =   1
         Left            =   12540
         TabIndex        =   14
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6675
            Index           =   1
            Left            =   30
            TabIndex        =   15
            Top             =   180
            Width           =   11745
            _cx             =   20717
            _cy             =   11774
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
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
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   13160660
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   0
            Position        =   1
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
            Begin VB.Frame Frame5 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   1
               Left            =   15
               TabIndex        =   17
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg5 
                  Height          =   6030
                  Index           =   1
                  Left            =   0
                  TabIndex        =   23
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion1.frx":3AE7
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
            Begin VB.Frame Frame6 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   1
               Left            =   12360
               TabIndex        =   16
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg6 
                  Height          =   6030
                  Index           =   1
                  Left            =   0
                  TabIndex        =   28
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Cols            =   12
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion1.frx":3C3F
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Caption         =   "Frame1"
         Height          =   6855
         Index           =   2
         Left            =   12840
         TabIndex        =   10
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6675
            Index           =   2
            Left            =   30
            TabIndex        =   11
            Top             =   180
            Width           =   11745
            _cx             =   20717
            _cy             =   11774
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
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
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   13160660
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   0
            Position        =   1
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
            Begin VB.Frame Frame5 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   2
               Left            =   15
               TabIndex        =   13
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg5 
                  Height          =   6030
                  Index           =   2
                  Left            =   0
                  TabIndex        =   24
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion1.frx":3DB8
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
            Begin VB.Frame Frame6 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   2
               Left            =   12360
               TabIndex        =   12
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg6 
                  Height          =   6030
                  Index           =   2
                  Left            =   0
                  TabIndex        =   29
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Cols            =   12
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion1.frx":3F10
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   0  'None
         Height          =   6855
         Index           =   3
         Left            =   13140
         TabIndex        =   6
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6675
            Index           =   3
            Left            =   30
            TabIndex        =   7
            Top             =   180
            Width           =   11745
            _cx             =   20717
            _cy             =   11774
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
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
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   13160660
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   0
            Position        =   1
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
            Begin VB.Frame Frame5 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   3
               Left            =   15
               TabIndex        =   9
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg5 
                  Height          =   6030
                  Index           =   3
                  Left            =   0
                  TabIndex        =   25
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion1.frx":408A
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
            Begin VB.Frame Frame6 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   3
               Left            =   12360
               TabIndex        =   8
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg6 
                  Height          =   6030
                  Index           =   3
                  Left            =   0
                  TabIndex        =   30
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Cols            =   12
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion1.frx":41E2
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
      Begin VB.Frame Frame1 
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   0  'None
         Height          =   6855
         Index           =   4
         Left            =   13440
         TabIndex        =   2
         Top             =   375
         Width           =   11805
         Begin SizerOneLibCtl.TabOne TabOne2 
            Height          =   6675
            Index           =   4
            Left            =   30
            TabIndex        =   3
            Top             =   180
            Width           =   11745
            _cx             =   20717
            _cy             =   11774
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "Courier New"
               Size            =   9.75
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
            BackColor       =   12632256
            ForeColor       =   -2147483630
            FrontTabColor   =   13160660
            BackTabColor    =   12632256
            TabOutlineColor =   -2147483632
            FrontTabForeColor=   -2147483641
            Caption         =   " Terminados  | Intermedios "
            Align           =   0
            CurrTab         =   0
            FirstTab        =   0
            Style           =   0
            Position        =   1
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
            Begin VB.Frame Frame5 
               BackColor       =   &H00C0FFFF&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   4
               Left            =   15
               TabIndex        =   5
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg5 
                  Height          =   6030
                  Index           =   4
                  Left            =   0
                  TabIndex        =   26
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Cols            =   11
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion1.frx":435C
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
            Begin VB.Frame Frame6 
               BackColor       =   &H00C0FFC0&
               BorderStyle     =   0  'None
               Caption         =   "Frame5"
               Height          =   6315
               Index           =   4
               Left            =   12360
               TabIndex        =   4
               Top             =   15
               Width           =   11715
               Begin VSFlex7Ctl.VSFlexGrid Fg6 
                  Height          =   6030
                  Index           =   4
                  Left            =   0
                  TabIndex        =   31
                  Top             =   0
                  Width           =   11715
                  _cx             =   20664
                  _cy             =   10636
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
                  Cols            =   12
                  FixedRows       =   1
                  FixedCols       =   1
                  RowHeightMin    =   0
                  RowHeightMax    =   0
                  ColWidthMin     =   0
                  ColWidthMax     =   0
                  ExtendLastCol   =   0   'False
                  FormatString    =   $"FrmConsultaProduccion1.frx":44B4
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
   End
End
Attribute VB_Name = "FrmConsultaProduccion1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************************************************************
'* Nombre Archivo   : FRMCONSULTAPRODUCCION
'* Tipo             : MODULO
'* Descripcion      : FORMULARIO QUE PERMITE CONSULTAR EL PLAN DE PRODUCCION ACTIVO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* ULTIMA REVISION  : 10/11/09
'* VERSION          : 1.0
'*****************************************************************************************************
Option Explicit

Dim RstInsumos As New ADODB.Recordset      ' RECORDSET QUE ALMACENARA LOS INSUMOS
Dim SeEjecuto As Boolean                   ' VERIFICA QUE EL EVENTO ACTIVATE SE EJECUTE UNA SOLO VEZ
Dim xCon1 As New ADODB.Connection          ' CONECCION A LA BASE DE DATOS
Dim xCon2 As New ADODB.Connection          ' CONECCION A LA BASE DE DATOS

Enum Devolver
    Receta = 1
    Cantidad = 2
End Enum

Private Sub CmdPrin_Click()
    On Error GoTo error
    
    Dim oExport As New SGI2_funciones.Formularios
    
    Me.MousePointer = vbHourglass
    oExport.VSFlexGrid_Exportar_MSExcel xCon, Fg7, "Unificado de Producción", "Periodo: " & AnoTra, "", "Unificado de Producción"
    Set oExport = Nothing
    Me.MousePointer = vbDefault
    Exit Sub
    
error:
    Me.MousePointer = vbDefault
    SHOW_ERROR Me.Name, "Exportar"
End Sub

Private Sub CmdSalir_Click()
    Frame2.Visible = False
    Toolbar1.Enabled = True
    TabOne1.Enabled = True
End Sub

Private Sub Form_Activate()
    ' SEGUNDO EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    If SeEjecuto = False Then
        SeEjecuto = True
        Dim A As Integer
        Dim xIndex As Integer
        Dim xRuta As String
        Dim RstEmp As New ADODB.Recordset

        Set xCon1 = AbrirConecciones(AP_RUTABD + "data.mdb")

        RST_Busq RstEmp, "SELECT mae_empresa.* From mae_empresa WHERE (((mae_empresa.anotra)=" & AnoTra & ") AND ((mae_empresa.activo)=-1))", xCon1

        If RstEmp.RecordCount <> 0 Then
            xIndex = 0
            RstEmp.MoveFirst

            For A = 1 To RstEmp.RecordCount
                TabOne1.TabCaption(xIndex) = " " & Trim(RstEmp("abrevia")) & " "
                TabOne1.TabVisible(xIndex) = True

                xRuta = AP_RUTABD + Trim(RstEmp("ruta"))

                Set xCon2 = Nothing
                Set xCon2 = AbrirConecciones(xRuta)

                CargaTerminados xIndex
                CargarIntermedios xIndex

                RstEmp.MoveNext

                If RstEmp.EOF = True Then
                    Exit For
                End If
                xIndex = xIndex + 1
            Next A
        End If
        TabOne1.CurrTab = 0
    End If
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : CargaTerminados
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LA LISTA DE PRODUCTOS TERMINADOS DEL PLAN DE PRODUCCION ACTIVO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Indice       |  Integer   |  INDICE DEL CONTROL Fg5
'* DEVUELVE         :
'*****************************************************************************************************
Sub CargaTerminados(Indice As Integer)
    Dim A, B As Integer
    Dim xTotal As Double
    Dim RstTmp As New ADODB.Recordset
    Dim RstPro As New ADODB.Recordset
    
    RST_Busq RstPro, "SELECT alm_inventario.id, alm_inventario.idmae, alm_inventario.descripcion, Sum(ges_plaproddet.cantidad) AS SumaDecantidad, mae_unidades.abrev, " _
        & " alm_inventario.stckini, alm_inventario.idunimed FROM ges_plaprod LEFT JOIN (mae_unidades RIGHT JOIN (ges_plaproddet LEFT JOIN alm_inventario " _
        & " ON ges_plaproddet.codpro = alm_inventario.id) ON mae_unidades.id = alm_inventario.idunimed) ON ges_plaprod.id = ges_plaproddet.idpv Where (((ges_plaproddet.idmes) <> 13)) " _
        & " GROUP BY alm_inventario.id, alm_inventario.idmae, alm_inventario.descripcion, mae_unidades.abrev, alm_inventario.stckini, alm_inventario.idunimed, " _
        & " ges_plaprod.activo HAVING (((ges_plaprod.activo)=-1))", xCon2
    
    Fg5(Indice).Rows = 1
    If RstPro.RecordCount <> 0 Then
        
        RstPro.MoveFirst
        For A = 1 To RstPro.RecordCount
            
            Fg5(Indice).Rows = Fg5(Indice).Rows + 1
            Fg5(Indice).TextMatrix(A, 0) = NulosN(RstPro("idmae"))
            Fg5(Indice).TextMatrix(A, 1) = NulosC(RstPro("descripcion"))
            Fg5(Indice).TextMatrix(A, 2) = RstPro("id")
            Fg5(Indice).TextMatrix(A, 3) = NulosC(RstPro("abrev"))
            
            'Se consulta los datos del Plan de Produccion Activo
            Dim RstTmpAux As New ADODB.Recordset
            RST_Busq RstTmpAux, "SELECT ges_plaprod.id, ges_plaprod.fchini, ges_plaprod.fchfin, ges_plaprod.activo" _
                & " From ges_plaprod" _
                & " WHERE (((ges_plaprod.activo)=-1))", xCon2
            
            'Se consulta el saldo actual hasta un dia antes del nuevo plan
            Dim xTot As Double
            xTot = SaldoActual(RstPro("id"), "01/01/" & AnoTra, CDate(RstTmpAux("fchini") - 1), xCon2)
            
            Fg5(Indice).TextMatrix(A, 5) = Format(xTot, FORMAT_MONTO)
            Fg5(Indice).TextMatrix(A, 10) = NulosN(RstPro("idunimed"))
            'Total Programado
            Fg5(Indice).TextMatrix(A, 4) = Format(RstPro("SumaDecantidad"), FORMAT_MONTO)
            
            Set RstTmp = Nothing

            RST_Busq RstTmp, "SELECT pro_producciondet.iditem, Sum(pro_producciondet.cantidad) AS SumaDecantidad From pro_producciondet GROUP BY pro_producciondet.iditem " _
                & " HAVING (((pro_producciondet.iditem)=" & RstPro("id") & "))", xCon2
                
            'Se consulta lo producido desde el 1 de enero del año actual
            'hasta el primer dia de inicio de la nueva Programacion
            Dim RstTmp2 As New ADODB.Recordset
            RST_Busq RstTmp2, "SELECT pro_producciondet.iditem, Sum(pro_producciondet.cantidad) AS SumaDecantidad" _
                & " FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro" _
                & " WHERE (((pro_produccion.dia)>=CDate('" & "01/01/" & AnoTra & "') And (pro_produccion.dia)<CDate('" & RstTmpAux("fchini") & "')))" _
                & " GROUP BY pro_producciondet.iditem" _
                & " HAVING (((pro_producciondet.iditem)=" & RstPro("id") & "))", xCon2

            If RstTmp.RecordCount <> 0 Then
                If RstTmp2.RecordCount <> 0 Then
                    'Se realiza la diferencia del total producido
                    'menos lo producido hasta el primer dia de la programacion
                    Fg5(Indice).TextMatrix(A, 6) = Format(NulosN(RstTmp("SumaDecantidad") - RstTmp2("SumaDecantidad")), FORMAT_MONTO)
                Else
                    Fg5(Indice).TextMatrix(A, 6) = Format(NulosN(RstTmp("SumaDecantidad")), FORMAT_MONTO)
                End If
            Else
                Fg5(Indice).TextMatrix(A, 6) = "0.00"
            End If
            
            xTotal = NulosN(Fg5(Indice).TextMatrix(A, 5)) + NulosN(Fg5(Indice).TextMatrix(A, 6))
            Fg5(Indice).TextMatrix(A, 7) = Format(xTotal, FORMAT_MONTO)
            Fg5(Indice).TextMatrix(A, 8) = NulosN(Fg5(Indice).TextMatrix(A, 4)) - NulosN(Fg5(Indice).TextMatrix(A, 7))
            Fg5(Indice).TextMatrix(A, 8) = Format(Fg5(Indice).TextMatrix(A, 8), FORMAT_MONTO)
            
            With Fg5(Indice)
                .Select A, 8, A, 8
                .FillStyle = flexFillRepeat
                If NulosN(Fg5(Indice).TextMatrix(A, 8)) <= 0 Then
                    .CellForeColor = &HFF0000
                Else
                    .CellForeColor = &HFF&
                End If
            End With
            Fg5(Indice).TextMatrix(A, 8) = Format(Abs(Fg5(Indice).TextMatrix(A, 8)), FORMAT_MONTO)
            
            RstPro.MoveNext
            
            If RstPro.EOF = True Then
                Exit For
            End If
        Next A
    End If
    
    With Fg5(Indice)
        .Select 1, 4, Fg5(Indice).Rows - 1, 4
        .FillStyle = flexFillRepeat
        .CellBackColor = &HDDFFFF
    End With

    With Fg5(Indice)
        .Select 1, 5, Fg5(Indice).Rows - 1, 7
        .FillStyle = flexFillRepeat
        .CellBackColor = &HE0FEE7
        
        .Select 1, 1, 1, 1
    End With
End Sub


Private Sub Form_Load()
    ' PRIMER EVENTO A EJECUTARSE CUANDO SE CARGUE EL FORMULARIO
    Dim A As Integer
    Dim xIndex As Integer
    xIndex = 0
    SeEjecuto = False
    Frame15.BackColor = &HC0C0C0
    For A = 1 To 5
        TabOne1.TabVisible(xIndex) = False
        Frame1(xIndex).BackColor = &H8000000F
        Frame5(xIndex).BackColor = &H8000000F
        Frame6(xIndex).BackColor = &H8000000F
        Fg5(xIndex).ColWidth(2) = 0
        Fg5(xIndex).ColWidth(3) = 0
        Fg5(xIndex).ColWidth(9) = 0
        Fg5(xIndex).ColWidth(10) = 0
        
        Fg6(xIndex).ColWidth(2) = 0
        Fg6(xIndex).ColWidth(3) = 0
        Fg6(xIndex).ColWidth(10) = 0
        xIndex = xIndex + 1
    Next A
    
    TabOne1.CurrTab = 0
    TabOne2(0).CurrTab = 0
    
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : CargarIntermedios
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : CARGA LA LISTA DE PRODUCTOS INTERMEDIOS DEL PLAN DE PRODUCCION ACTIVO
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    Indice       |  Integer   |  INDICE DEL CONTROL Fg6
'* DEVUELVE         :
'*****************************************************************************************************
Sub CargarIntermedios(Indice As Integer)
    Dim A, B As Integer
    Dim xTotal As Double
    Dim RstTmp As New ADODB.Recordset
    Dim RstPro As New ADODB.Recordset
    Set RstPro = Nothing
    RST_Busq RstPro, "SELECT alm_inventario.id, alm_inventario.idmae, alm_inventario.descripcion, mae_unidades.abrev,  alm_inventario.stckini AS stckini1,alm_inventario.stckini, alm_inventario.idunimed, " _
        & " Sum(ges_plaproddet2.cantidad) AS SumaDecantidad FROM mae_unidades RIGHT JOIN (ges_plaprod LEFT JOIN (ges_plaproddet2 LEFT JOIN alm_inventario " _
        & " ON ges_plaproddet2.codpro = alm_inventario.id) ON ges_plaprod.id = ges_plaproddet2.idpv) ON mae_unidades.id = alm_inventario.idunimed " _
        & " Where (((ges_plaproddet2.idmes) <> 13)) GROUP BY alm_inventario.id, alm_inventario.idmae, alm_inventario.descripcion, mae_unidades.abrev, " _
        & " alm_inventario.stckini, alm_inventario.idunimed, ges_plaprod.activo HAVING (((ges_plaprod.activo)=-1))", xCon2
    
    Fg6(Indice).Rows = 1
    If RstPro.RecordCount <> 0 Then
        
        RstPro.MoveFirst
        For A = 1 To RstPro.RecordCount
            
            Fg6(Indice).Rows = Fg6(Indice).Rows + 1
            Fg6(Indice).TextMatrix(A, 0) = NulosN(RstPro("idmae"))
            Fg6(Indice).TextMatrix(A, 1) = NulosC(RstPro("descripcion"))
            Fg6(Indice).TextMatrix(A, 2) = RstPro("id")
            Fg6(Indice).TextMatrix(A, 3) = NulosC(RstPro("abrev"))
            
            'Se consulta los datos del Plan de Produccion Activo
            Dim RstTmpAux As New ADODB.Recordset
            RST_Busq RstTmpAux, "SELECT ges_plaprod.id, ges_plaprod.fchini, ges_plaprod.fchfin, ges_plaprod.activo" _
                & " From ges_plaprod" _
                & " WHERE (((ges_plaprod.activo)=-1))", xCon2
                
            'Se consulta el saldo actual hasta un dia antes del nuevo plan
            Dim xTot As Double
            xTot = SaldoActual(RstPro("id"), "01/01/" & AnoTra, CDate(RstTmpAux("fchini") - 1), xCon2)
            
            Fg6(Indice).TextMatrix(A, 5) = Format(xTot, FORMAT_MONTO)
            
            Fg6(Indice).TextMatrix(A, 10) = NulosN(RstPro("idunimed"))
            
            Fg6(Indice).TextMatrix(A, 4) = Format(RstPro("SumaDecantidad"), FORMAT_MONTO)
            
            Set RstTmp = Nothing
            
            RST_Busq RstTmp, "SELECT pro_producciondet.iditem, Sum(pro_producciondet.cantidad) AS SumaDecantidad From pro_producciondet GROUP BY pro_producciondet.iditem " _
                & " HAVING (((pro_producciondet.iditem)=" & RstPro("id") & "))", xCon2
            
            'Se consulta lo producido desde el 1 de enero del año actual
            'hasta el primer dia de inicio de la nueva Programacion
            Dim RstTmp2 As New ADODB.Recordset
            RST_Busq RstTmp2, "SELECT pro_producciondet.iditem, Sum(pro_producciondet.cantidad) AS SumaDecantidad" _
                & " FROM pro_produccion INNER JOIN pro_producciondet ON pro_produccion.id = pro_producciondet.idpro" _
                & " WHERE (((pro_produccion.dia)>=CDate('" & "01/01/" & AnoTra & "') And (pro_produccion.dia)<=CDate('" & RstTmpAux("fchini") & "')))" _
                & " GROUP BY pro_producciondet.iditem" _
                & " HAVING (((pro_producciondet.iditem)=" & RstPro("id") & "))", xCon2

            If RstTmp.RecordCount <> 0 Then
                If RstTmp2.RecordCount <> 0 Then
                    'Se realiza la diferencia del total producido
                    'menos lo producido hasta el primer dia de la programacion
                    Fg6(Indice).TextMatrix(A, 6) = Format(NulosN(RstTmp("SumaDecantidad") - RstTmp2("SumaDecantidad")), FORMAT_MONTO)
                End If
            Else
                Fg6(Indice).TextMatrix(A, 6) = "0.00"
            End If
            
            xTotal = NulosN(Fg6(Indice).TextMatrix(A, 5)) + NulosN(Fg6(Indice).TextMatrix(A, 6))
            Fg6(Indice).TextMatrix(A, 7) = Format(xTotal, FORMAT_MONTO)
            Fg6(Indice).TextMatrix(A, 8) = NulosN(Fg6(Indice).TextMatrix(A, 4)) - NulosN(Fg6(Indice).TextMatrix(A, 7))
            Fg6(Indice).TextMatrix(A, 8) = Format(Fg6(Indice).TextMatrix(A, 8), FORMAT_MONTO)
            'buscamos la receta del productos
            Fg6(Indice).TextMatrix(A, 9) = BuscaReceta(RstPro("id"), Cantidad)
            Fg6(Indice).TextMatrix(A, 11) = Format(BuscaReceta(RstPro("id"), Receta) * NulosN(Fg6(Indice).TextMatrix(A, 8)), FORMAT_MONTO)
            
            With Fg6(Indice)
                .Select A, 8, A, 8
                .FillStyle = flexFillRepeat
                If NulosN(Fg6(Indice).TextMatrix(A, 8)) <= 0 Then
                    .CellForeColor = &HFF0000
                Else
                    .CellForeColor = &HFF&
                End If
            End With
            Fg6(Indice).TextMatrix(A, 8) = Abs(NulosN(Format(Fg6(Indice).TextMatrix(A, 8))))
            
            RstPro.MoveNext
            
            If RstPro.EOF = True Then
                Exit For
            End If
        Next A
    End If
    
    With Fg6(Indice)
        .Select 1, 4, Fg6(Indice).Rows - 1, 4
        .FillStyle = flexFillRepeat
        .CellBackColor = &HDDFFFF
    End With

    With Fg6(Indice)
        .Select 1, 5, Fg6(Indice).Rows - 1, 7
        .FillStyle = flexFillRepeat
        .CellBackColor = &HE0FEE7
        .Select 1, 1, 1, 1
    End With
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : BuscaReceta
'* Tipo             : FUNCION
'* Descripcion      : BUSCA UNA RECETA
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       : PARAMENTO    |  TIPO      |  DESCRIPCION
'*                    --------------------------------------------------------------------------------
'*                    CodigoItem   |  Integer   |  ESPECIFICA EL ID DEL PRODUCTO
'*                    QueDevuelve  |  Devolver  |  ESPECIFICA EL VALOR DEL TIPO Devolver DEFINIDO
'* DEVUELVE         :
'*****************************************************************************************************
Function BuscaReceta(CodigoItem As Integer, QueDevuelve As Devolver) As Variant
    Dim xRst As New ADODB.Recordset
    
    RST_Busq xRst, "SELECT pro_receta.iditem, pro_receta.codrec, pro_recetains.iditem, pro_recetains.canpro, alm_inventario.descripcion FROM alm_inventario RIGHT JOIN " _
        & " (pro_receta INNER JOIN pro_recetains ON pro_receta.id = pro_recetains.idrec) ON alm_inventario.id = pro_recetains.iditem " _
        & " WHERE (((pro_receta.iditem)=" & CodigoItem & ") AND ((pro_receta.prirec)=1) AND ((alm_inventario.tippro)=1))", xCon2

    If xRst.RecordCount <> 0 Then
        If QueDevuelve = Cantidad Then
            BuscaReceta = NulosC(xRst("codrec"))
        Else
            BuscaReceta = NulosN(xRst("canpro"))
        End If
    Else
        BuscaReceta = 0
    End If
    Set xRst = Nothing
End Function

'*****************************************************************************************************
'* Nombre Archivo   : VerUnificado
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : MUESTRA LA INFORMACION UNIFICADA DEL PROGRAMA DE PRODUCCION, PARA ELLO JALA EL
'*                    PROGRAMA DE PRODUCCION DE LAS DEMAS BASES DE DATO ACTIVAS
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub VerUnificado()
    Dim A As Integer
    Dim Total As Double
    Dim xIndex As Integer
    
    TabOne1.Enabled = False
    Toolbar1.Enabled = False
    Frame2.Left = 75
    Frame2.Top = 705
    Frame2.Visible = True
    
    PreparaRST
    
    xIndex = 0
    
    For A = 1 To 5
        Dim B As Integer
        If TabOne1.TabVisible(xIndex) = True Then
            For B = 1 To Fg5(xIndex).Rows - 1
                RstInsumos.Filter = adFilterNone
                If RstInsumos.RecordCount <> 0 Then
                    RstInsumos.MoveFirst
                End If
                
                RstInsumos.Filter = "cod_item = '" & Fg5(xIndex).TextMatrix(B, 0) & "'"
                If RstInsumos.RecordCount = 0 Then
                    RstInsumos.AddNew
                    RstInsumos("descripcion") = Fg5(xIndex).TextMatrix(B, 1)
                    RstInsumos("unimed") = Fg5(xIndex).TextMatrix(B, 3)
                    RstInsumos("programado") = Format(NulosN(Fg5(xIndex).TextMatrix(B, 4)), FORMAT_MONTO)
                    RstInsumos("stckini") = Format(NulosN(Fg5(xIndex).TextMatrix(B, 5)), FORMAT_MONTO)
                    RstInsumos("producido") = Format(NulosN(Fg5(xIndex).TextMatrix(B, 6)), FORMAT_MONTO)
                    RstInsumos("total") = Format(NulosN(Fg5(xIndex).TextMatrix(B, 7)), FORMAT_MONTO)
                    RstInsumos("porprod") = NulosN(Fg5(xIndex).TextMatrix(B, 8))
                    RstInsumos("saldo") = Format(NulosN(Fg5(xIndex).TextMatrix(B, 9)), FORMAT_MONTO)
                    RstInsumos("cod_item") = Fg5(xIndex).TextMatrix(B, 0)
                Else
                    If RstInsumos.RecordCount = 1 Then
                        RstInsumos("programado") = RstInsumos("programado") + NulosN(Fg5(xIndex).TextMatrix(B, 4))
                        RstInsumos("porprod") = RstInsumos("porprod") + NulosN(Fg5(xIndex).TextMatrix(B, 8))
                        RstInsumos("saldo") = RstInsumos("saldo") + NulosN(Fg5(xIndex).TextMatrix(B, 9))
                    Else
                        'este error nunca debe de ocurrir
                        MsgBox "Hay mas de un items con el mismo codigo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    End If
                End If
            Next B
            
            For B = 1 To Fg6(xIndex).Rows - 1
                RstInsumos.Filter = adFilterNone
                If RstInsumos.RecordCount <> 0 Then
                    RstInsumos.MoveFirst
                End If
                RstInsumos.Filter = "cod_item = '" & Fg6(xIndex).TextMatrix(B, 0) & "'"
                If RstInsumos.RecordCount = 0 Then
                    RstInsumos.AddNew
                    RstInsumos("descripcion") = Fg6(xIndex).TextMatrix(B, 1)
                    RstInsumos("unimed") = Fg6(xIndex).TextMatrix(B, 3)
                    RstInsumos("programado") = Format(NulosN(Fg6(xIndex).TextMatrix(B, 4)), FORMAT_MONTO)
                    RstInsumos("stckini") = Format(NulosN(Fg6(xIndex).TextMatrix(B, 5)), FORMAT_MONTO)
                    RstInsumos("producido") = Format(NulosN(Fg6(xIndex).TextMatrix(B, 6)), FORMAT_MONTO)
                    RstInsumos("total") = Format(NulosN(Fg6(xIndex).TextMatrix(B, 7)), FORMAT_MONTO)
                    RstInsumos("porprod") = NulosN(Fg6(xIndex).TextMatrix(B, 8))
                    RstInsumos("saldo") = Format(NulosN(Fg6(xIndex).TextMatrix(B, 9)), FORMAT_MONTO)
                    RstInsumos("cod_item") = Fg6(xIndex).TextMatrix(B, 0)
                Else
                    If RstInsumos.RecordCount = 1 Then
                        RstInsumos("programado") = RstInsumos("programado") + NulosN(Fg6(xIndex).TextMatrix(B, 4))
                        RstInsumos("porprod") = RstInsumos("porprod") + NulosN(Fg6(xIndex).TextMatrix(B, 8))
                        RstInsumos("saldo") = RstInsumos("saldo") + NulosN(Fg6(xIndex).TextMatrix(B, 9))
                    Else
                        'este error nunca debe de ocurrir
                        MsgBox "Hay mas de un items con el mismo codigo", vbInformation + vbOKOnly + vbDefaultButton1, xTitulo
                    End If
                End If
            Next B
        End If
        xIndex = xIndex + 1
    Next A
    
    RstInsumos.Filter = adFilterNone
    RstInsumos.Sort = "descripcion"
    RstInsumos.MoveFirst
    Fg7.Rows = 1
    For A = 1 To RstInsumos.RecordCount
        Fg7.Rows = Fg7.Rows + 1
        Fg7.TextMatrix(A, 1) = RstInsumos("descripcion")
        Fg7.TextMatrix(A, 2) = RstInsumos("cod_item")
        Fg7.TextMatrix(A, 3) = RstInsumos("unimed")
        Fg7.TextMatrix(A, 4) = Format(RstInsumos("programado"), FORMAT_MONTO)
        Fg7.TextMatrix(A, 5) = Format(RstInsumos("stckini"), FORMAT_MONTO)
        Fg7.TextMatrix(A, 6) = Format(RstInsumos("producido"), FORMAT_MONTO)
        Fg7.TextMatrix(A, 7) = Format(RstInsumos("total"), FORMAT_MONTO)
        Fg7.TextMatrix(A, 8) = (NulosN(RstInsumos("programado")) - NulosN(RstInsumos("total")))
        
        With Fg7
            .Select A, 8, A, 8
            .FillStyle = flexFillRepeat
            If NulosN(Fg7.TextMatrix(A, 8)) <= 0 Then
                .CellForeColor = &HFF0000
            Else
                .CellForeColor = &HFF&
            End If
        End With
        Fg7.TextMatrix(A, 8) = Abs(NulosN(NulosN(Fg7.TextMatrix(A, 8))))
        Fg7.TextMatrix(A, 8) = Format(NulosN(Fg7.TextMatrix(A, 8)), FORMAT_MONTO)
        RstInsumos.MoveNext
        If RstInsumos.EOF = True Then
            Exit For
        End If
    Next A
    
    With Fg7
        .Select 1, 4, Fg7.Rows - 1, 4
        .FillStyle = flexFillRepeat
        .CellBackColor = &HDDFFFF
    End With

    With Fg7
        .Select 1, 5, Fg7.Rows - 1, 7
        .FillStyle = flexFillRepeat
        .CellBackColor = &HE0FEE7
        .Select 1, 1, 1, 1
    End With
    
End Sub

'*****************************************************************************************************
'* Nombre Archivo   : PreparaRST
'* Tipo             : PROCEDIMIENTO
'* Descripcion      : PREPARA UN RECORDSET TEMPORAL
'* DISEÑADO POR     : ENRIQUE POLLONGO SIERRA
'* PARAMETROS       :
'* DEVUELVE         :
'*****************************************************************************************************
Sub PreparaRST()
    Dim xFun As New eps_librerias.FuncionesData
    Dim xCampos(9, 3) As String

    xCampos(0, 0) = "cod_item":     xCampos(0, 1) = "C":      xCampos(0, 2) = "16"
    xCampos(1, 0) = "unimed":       xCampos(1, 1) = "C":      xCampos(1, 2) = "4"
    xCampos(2, 0) = "descripcion":  xCampos(2, 1) = "C":      xCampos(2, 2) = "200"
    xCampos(3, 0) = "programado":   xCampos(3, 1) = "N":      xCampos(3, 2) = "2"
    xCampos(4, 0) = "stckini":      xCampos(4, 1) = "N":      xCampos(4, 2) = "2"
    xCampos(5, 0) = "producido":    xCampos(5, 1) = "N":      xCampos(5, 2) = "2"
    xCampos(6, 0) = "total":        xCampos(6, 1) = "N":      xCampos(6, 2) = "2"
    xCampos(7, 0) = "porprod":      xCampos(7, 1) = "N":      xCampos(7, 2) = "2"
    xCampos(8, 0) = "saldo":        xCampos(8, 1) = "N":      xCampos(8, 2) = "2"
    
    Set RstInsumos = xFun.CrearRstTMP(xCampos)
    RstInsumos.Open
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    If Button.Index = 1 Then VerUnificado
    
    If Button.Index = 3 Then
        Unload Me
    End If
End Sub
